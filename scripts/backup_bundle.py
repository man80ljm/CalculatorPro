"""制作、校验和恢复加密备份；七牛上传须显式选择。密钥只读外部文件/环境。"""
from __future__ import annotations
import argparse
import hashlib
import io
import json
import os
from pathlib import Path
import shutil
import tarfile
import tempfile
from datetime import datetime, timezone
from zoneinfo import ZoneInfo
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes

MAGIC = b"CPB1"
CHUNK = 1024 * 1024
EXCLUDED = {".git", ".venv", ".local", "__pycache__", ".pytest_cache", "node_modules", "backups"}


def digest(path):
    result = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for chunk in iter(lambda: handle.read(CHUNK), b""):
            result.update(chunk)
    return result.hexdigest()


def read_key(path):
    path = Path(path).resolve()
    if os.name != "nt" and path.stat().st_mode & 0o077:
        raise ValueError("备份密钥文件权限必须为 600 或 400。")
    key = path.read_bytes()
    if len(key) != 32:
        raise ValueError("备份密钥须为 32 字节随机文件。")
    return key


def encrypt(source, destination, key):
    nonce = os.urandom(12)
    encoder = Cipher(algorithms.AES(key), modes.GCM(nonce)).encryptor()
    encoder.authenticate_additional_data(MAGIC)
    with Path(source).open("rb") as incoming, Path(destination).open("xb") as outgoing:
        os.chmod(destination, 0o600)
        outgoing.write(MAGIC + nonce)
        for chunk in iter(lambda: incoming.read(CHUNK), b""):
            outgoing.write(encoder.update(chunk))
        outgoing.write(encoder.finalize())
        outgoing.write(encoder.tag)
        outgoing.flush()
        os.fsync(outgoing.fileno())


def decrypt(source, destination, key):
    with Path(source).open("rb") as incoming:
        if incoming.read(4) != MAGIC or Path(source).stat().st_size < 32:
            raise ValueError("不是 CalculatorPro 加密备份。")
        nonce = incoming.read(12)
        incoming.seek(-16, 2)
        tag = incoming.read(16)
        remaining = incoming.tell() - 32
        incoming.seek(16)
        decoder = Cipher(algorithms.AES(key), modes.GCM(nonce, tag)).decryptor()
        decoder.authenticate_additional_data(MAGIC)
        try:
            with Path(destination).open("xb") as outgoing:
                os.chmod(destination, 0o600)
                while remaining:
                    chunk = incoming.read(min(CHUNK, remaining))
                    if not chunk:
                        raise ValueError("备份文件不完整。")
                    outgoing.write(decoder.update(chunk))
                    remaining -= len(chunk)
                outgoing.write(decoder.finalize())
        except Exception:
            Path(destination).unlink(missing_ok=True)
            raise


def _source_files(root, prefix, excluded_paths=()):
    root = Path(root).resolve()
    for parent, dirs, files in os.walk(root, followlinks=False):
        parent = Path(parent)
        dirs[:] = [name for name in dirs if name not in EXCLUDED and not (parent / name).is_symlink()]
        for name in sorted(files):
            path = parent / name
            if path.is_symlink() or path.resolve() in excluded_paths:
                continue
            if root not in path.resolve().parents:
                raise ValueError("备份路径超出指定目录。")
            yield path, f"{prefix}/{path.relative_to(root).as_posix()}"


def make_bundle(dump, uploads, code, destination, key_file, database_format="postgres"):
    dump = Path(dump).resolve()
    key_file = Path(key_file).resolve()
    key = read_key(key_file)
    if not dump.is_file() or not dump.stat().st_size:
        raise ValueError("数据库备份不存在或为空，不能生成成功备份。")
    expected = b"PGDMP" if database_format == "postgres" else b"SQLite format 3\x00"
    with dump.open("rb") as handle:
        if handle.read(len(expected)) != expected:
            raise ValueError("数据库备份格式不正确。")
    destination = Path(destination).resolve()
    for root in (Path(uploads).resolve(), Path(code).resolve()):
        if not root.is_dir() or destination == root or root in destination.parents:
            raise ValueError("备份输出必须放在代码和上传目录之外。")
    destination.mkdir(parents=True, exist_ok=True)
    os.chmod(destination, 0o700)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    name = f"calculatorpro-{stamp}-{os.urandom(4).hex()}.cpbackup"
    output = destination / name
    with tempfile.TemporaryDirectory(prefix=".building-", dir=destination) as work:
        work = Path(work)
        archive_path = work / "snapshot.tar.gz"
        members = [(dump, "database.dump")]
        members.extend(_source_files(uploads, "uploads", (key_file,)))
        members.extend(_source_files(code, "code", (key_file, dump)))
        manifest = {"version": 1, "created_at": stamp, "database_format": database_format, "files": {}}
        with tarfile.open(archive_path, "w:gz") as archive:
            for path, relative in members:
                before = (path.stat().st_size, path.stat().st_mtime_ns)
                manifest["files"][relative] = {"bytes": before[0], "sha256": digest(path)}
                archive.add(path, arcname=relative, recursive=False)
                if before != (path.stat().st_size, path.stat().st_mtime_ns):
                    raise ValueError("备份期间资料发生变化，请暂停写入后重试。")
            data = json.dumps(manifest, ensure_ascii=False).encode("utf-8")
            info = tarfile.TarInfo("manifest.json")
            info.size = len(data)
            info.mode = 0o600
            archive.addfile(info, io.BytesIO(data))
        temporary = work / "encrypted.tmp"
        encrypt(archive_path, temporary, key)
        # 发出成功前先解密校验一次；明文临时目录随后清理。
        verify_bundle(temporary, key_file)
        os.replace(temporary, output)
    return output


def verify_bundle(package, key_file, restore_dir=None):
    """只有身份认证和逐文件 SHA256 都通过，才向空的新目录恢复。"""
    with tempfile.TemporaryDirectory(prefix="calculatorpro-restore-") as work:
        work = Path(work)
        archive_path = work / "snapshot.tar.gz"
        decrypt(package, archive_path, read_key(key_file))
        staged = work / "verified"
        staged.mkdir()
        with tarfile.open(archive_path, "r:gz") as archive:
            entries = archive.getmembers()
            names = [entry.name for entry in entries]
            if len(names) != len(set(names)) or "manifest.json" not in names:
                raise ValueError("备份目录或清单重复、缺失。")
            for entry in entries:
                path = (staged / entry.name).resolve()
                if not entry.isfile() or staged not in path.parents or "\\" in entry.name:
                    raise ValueError("备份包含不安全路径。")
            manifest = json.load(archive.extractfile("manifest.json"))
            if set(names) != set(manifest["files"]) | {"manifest.json"}:
                raise ValueError("备份文件与清单不一致。")
            for entry in entries:
                path = staged / entry.name
                path.parent.mkdir(parents=True, exist_ok=True)
                with archive.extractfile(entry) as incoming, path.open("xb") as outgoing:
                    shutil.copyfileobj(incoming, outgoing, CHUNK)
                if entry.name != "manifest.json":
                    info = manifest["files"][entry.name]
                    if path.stat().st_size != info["bytes"] or digest(path) != info["sha256"]:
                        raise ValueError("备份文件校验失败。")
        if restore_dir is not None:
            target = Path(restore_dir).resolve()
            if target.exists():
                raise ValueError("恢复只能写入尚不存在的新目录。")
            shutil.copytree(staged, target)
        return manifest


def _qiniu_connection():
    import qiniu
    values = [os.environ.get(name, "").strip() for name in ("QINIU_ACCESS_KEY", "QINIU_SECRET_KEY", "QINIU_BACKUP_BUCKET")]
    if not all(values) or os.environ.get("QINIU_BACKUP_PRIVATE") != "1":
        raise ValueError("请配置七牛备份凭据，并确认使用私有空间。")
    access, secret, bucket = values
    auth = qiniu.Auth(access, secret)
    # 保持SDK自动查询区域；创建空的自定义Region会使上传端点为空。
    qiniu.config.get_default("default_zone").scheme = "https"
    manager = qiniu.BucketManager(auth, preferred_scheme="https")
    details, info = manager.bucket_info(bucket)
    if info.status_code != 200 or not details or details.get("private") != 1:
        raise ValueError("备份空间不是已确认的私有空间，已停止上传或下载。")
    return auth, manager, bucket, access, secret


def upload_qiniu(package, include_monthly=False):
    import qiniu
    auth, manager, bucket, _, _ = _qiniu_connection()
    prefix = "calculatorpro-backups/v1/"
    today = datetime.now(ZoneInfo("Asia/Shanghai"))
    daily = f"{prefix}daily/{today:%Y-%m-%d}/{Path(package).name}"
    keys = [daily]
    if today.day == 1 or include_monthly:
        keys.append(f"{prefix}monthly/{today:%Y-%m}/{Path(package).name}")
    for object_key in keys:
        days = 365 if "/monthly/" in object_key else 30
        token = auth.upload_token(bucket, object_key, 3600, {"insertOnly": 1, "deleteAfterDays": days})
        result, info = qiniu.put_file(token, object_key, str(package), version="v2", bucket_name=bucket)
        if result is None or info.status_code != 200:
            raise RuntimeError("七牛备份上传失败，本地备份已保留。")
        remote, info = manager.stat(bucket, object_key)
        if info.status_code != 200 or remote is None or remote.get("hash") != qiniu.etag(str(package)) or remote.get("fsize") != Path(package).stat().st_size:
            raise RuntimeError("七牛备份校验失败，本地备份已保留。")
    return keys


def download_qiniu(object_key, destination):
    """从七牛HTTPS S3接口下载加密包；不依赖公开域名，不覆盖已有文件。"""
    import boto3
    from botocore.config import Config
    if not object_key.startswith("calculatorpro-backups/v1/") or not object_key.endswith(".cpbackup"):
        raise ValueError("只能下载CalculatorPro加密备份。")
    _, _, bucket, access, secret = _qiniu_connection()
    region = os.environ.get("QINIU_S3_REGION", "cn-south-1")
    if region not in {"cn-east-1", "cn-east-2", "cn-north-1", "cn-south-1", "us-north-1", "ap-southeast-1", "ap-southeast-2", "ap-southeast-3"}:
        raise ValueError("七牛S3区域配置不正确。")
    target = Path(destination).resolve()
    if target.exists():
        raise ValueError("下载不能覆盖已有文件。")
    target.parent.mkdir(parents=True, exist_ok=True)
    client = boto3.client("s3", endpoint_url=f"https://s3.{region}.qiniucs.com", region_name=region,
        aws_access_key_id=access, aws_secret_access_key=secret,
        config=Config(signature_version="s3v4", s3={"addressing_style": "path"}, retries={"max_attempts": 2},
                      connect_timeout=10, read_timeout=60, response_checksum_validation="when_required"))
    response = client.get_object(Bucket=bucket, Key=object_key)
    body = response["Body"]
    size = int(response["ContentLength"])
    try:
        if shutil.disk_usage(target.parent).free < size + 200 * CHUNK:
            raise ValueError("下载所需空间不足。")
        with target.open("xb") as outgoing:
            os.chmod(target, 0o600)
            shutil.copyfileobj(body, outgoing, CHUNK)
        with target.open("rb") as handle:
            if target.stat().st_size != size or handle.read(4) != MAGIC:
                raise ValueError("下载的备份不完整或格式不正确。")
    except Exception:
        target.unlink(missing_ok=True)
        raise
    finally:
        body.close()
    return target


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    pack = commands.add_parser("pack")
    pack.add_argument("--dump", required=True)
    pack.add_argument("--uploads", required=True)
    pack.add_argument("--code", required=True)
    pack.add_argument("--destination", required=True)
    pack.add_argument("--format", choices=["postgres", "sqlite"], default="postgres")
    pack.add_argument("--qiniu", action="store_true")
    pack.add_argument("--result-file")
    verify = commands.add_parser("verify")
    verify.add_argument("--package", required=True)
    verify.add_argument("--restore-dir")
    for command in (pack, verify):
        command.add_argument("--key-file", required=True)
    upload = commands.add_parser("upload")
    upload.add_argument("--package", required=True)
    upload.add_argument("--monthly", action="store_true")
    upload.add_argument("--result-file")
    download = commands.add_parser("download")
    download.add_argument("--object-key", required=True)
    download.add_argument("--destination", required=True)
    args = parser.parse_args()
    if args.command == "pack":
        package = make_bundle(args.dump, args.uploads, args.code, args.destination, args.key_file, args.format)
        print(f"备份已生成并校验：{package}")
        if args.result_file:
            Path(args.result_file).write_text(json.dumps({"package": str(package), "sha256": digest(package)}), encoding="utf-8")
        if args.qiniu:
            keys = upload_qiniu(package)
            print(f"七牛上传并校验完成：{len(keys)} 份")
    elif args.command == "verify":
        manifest = verify_bundle(args.package, args.key_file, args.restore_dir)
        print(f"备份校验通过：{len(manifest['files'])} 个文件")
    elif args.command == "upload":
        with Path(args.package).open("rb") as handle:
            if handle.read(4) != MAGIC:
                raise ValueError("仅允许上传加密后的备份。")
        keys = upload_qiniu(args.package, args.monthly)
        if args.result_file:
            Path(args.result_file).write_text(json.dumps({"package": str(Path(args.package).resolve()),
                "sha256": digest(args.package), "bytes": Path(args.package).stat().st_size,
                "bucket": os.environ["QINIU_BACKUP_BUCKET"], "objects": keys,
                "uploaded_at": datetime.now(timezone.utc).isoformat()}, ensure_ascii=False), encoding="utf-8")
        print(f"七牛上传并校验完成：{len(keys)} 份")
    else:
        download_qiniu(args.object_key, args.destination)
        print("七牛加密备份已下载，请再执行verify进行解密及逐文件校验。")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        # 不回显 SDK 的请求、授权头或异常堆栈。
        if isinstance(exc, ValueError):
            print(f"备份未完成：{exc}")
        else:
            print("备份未完成，请检查配置、存储空间和网络。本地已有备份会保留。")
        raise SystemExit(1)
