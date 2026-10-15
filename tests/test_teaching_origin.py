"""TLS proxy origin checks for calendar, lesson-plan and teacher selection writes."""
import pytest
from starlette.requests import Request
from web_app.teaching.request_origin import same_origin


def request(origin, *, scheme="http", peer="172.20.0.1", proto="https", fetch_site="same-origin"):
    headers=[(b"host",b"calc.geekhuang.com"),(b"sec-fetch-site",fetch_site.encode())]
    if origin is not None:headers.append((b"origin",origin.encode()))
    if proto is not None:headers.append((b"x-forwarded-proto",proto.encode()))
    return Request({"type":"http","scheme":scheme,"server":("calc.geekhuang.com",80),"client":(peer,12345),
                    "root_path":"","path":"/api/lesson-plans","query_string":b"","headers":headers})


@pytest.mark.parametrize("origin,scheme,peer,proto,allowed",[
    ("https://calc.geekhuang.com","http","172.20.0.1","https",True),
    ("https://calc.geekhuang.com","https","198.51.100.2",None,True),
    ("http://calc.geekhuang.com","http","198.51.100.2",None,True),
    ("https://calc.geekhuang.com","http","198.51.100.2","https",False),
    ("https://other.example","http","172.20.0.1","https",False),
    ("http://calc.geekhuang.com","http","172.20.0.1","https",False),
    ("https://calc.geekhuang.com","http","172.20.0.1","https,http",False),
    ("null","http","172.20.0.1","https",False),
    ("https://calc.geekhuang.com/path","http","172.20.0.1","https",False),
    ("https://user@calc.geekhuang.com","http","172.20.0.1","https",False),
])
def test_origin_checks_only_accept_same_site_and_trusted_proxy(monkeypatch,origin,scheme,peer,proto,allowed):
    monkeypatch.setenv("TRUSTED_PROXIES","127.0.0.1/32,172.16.0.0/12")
    assert same_origin(request(origin,scheme=scheme,peer=peer,proto=proto)) is allowed


def test_cross_site_is_rejected_even_without_origin():
    assert not same_origin(request(None,fetch_site="cross-site"))
    assert same_origin(request(None))
