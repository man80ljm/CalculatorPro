import os
import shutil

# 定义需要修复的乱码行
replacements = {
    # === 修复 541 行及附近的下拉框乱码 ===
    'self.combo_dist.addItems(["鏍囧噯姝ｆ€?, "楂樺垎鍊惧悜", "浣庡垎鍊惧悜", "涓ゆ瀬鍒嗗寲", "妗ｄ綅鎵撳垎", "瀹屽叏闅忔満"] )': 
    '        self.combo_dist.addItems(["标准正态", "高分倾向", "低分倾向", "两极分化", "档位打分", "完全随机"])',
    
    'self.combo_spread.addItems(["澶ц法搴︼紙14-23鍒嗭級", "涓法搴︼紙7-13鍒嗭級", "灏忚法搴︼紙2-6鍒嗭級"] )':
    '        self.combo_spread.addItems(["大跨度（14-23分）", "中跨度（7-13分）", "小跨度（2-6分）"])',

    'self.combo_noise.addItems(["鏃?, "璇︽儏..."])': 
    '        self.combo_noise.addItems(["无", "详情..."])',

    'self.combo_style.addItems(["涓撲笟", "鍙ｈ", "绠€娲?, "璇︾粏", "骞介粯"] )':
    '        self.combo_style.addItems(["专业", "口语", "简洁", "详细", "幽默"])',

    # === 修复顶部按钮乱码 ===
    'self.course_open_btn = QPushButton("????")': '        self.course_open_btn = QPushButton("开课信息")',
    'self.course_basic_btn = QPushButton("??????")': '        self.course_basic_btn = QPushButton("课程基本信息")',
    'self.relation_btn = QPushButton("?????????????")': '        self.relation_btn = QPushButton("课程考核与课程目标对应关系")',
    'self.ratio_btn = QPushButton("????")': '        self.ratio_btn = QPushButton("成绩占比")',
    'self.grad_req_btn = QPushButton("??????????????")': '        self.grad_req_btn = QPushButton("课程目标与毕业要求的对应关系")',
    'self.settings_btn = QPushButton("??")': '        self.settings_btn = QPushButton("设置")',

    # === 修复其他标签乱码 ===
    'layout.addWidget(QLabel("璇疯緭鍏ュ鐢熶汉鏁?"))': '        layout.addWidget(QLabel("请输入学生人数"))',
    "title = QLabel('璇剧▼鐩爣杈炬垚鎯呭喌璇勪环鍙婃€荤粨鍒嗘瀽鎶ュ憡')": "        title = QLabel('课程目标达成情况评价及总结分析报告')",
    'self.btn_download = QPushButton("鐐瑰嚮涓嬭浇")': '        self.btn_download = QPushButton("点击下载")',
    'control_layout.addWidget(QLabel("鍒嗘暟璺ㄥ害:"))': '        control_layout.addWidget(QLabel("分数跨度:"))',
    'control_layout.addWidget(QLabel("鍒嗗竷妯″紡:"))': '        control_layout.addWidget(QLabel("分布模式:"))',
    'control_layout.addWidget(QLabel("鍣０娉ㄥ叆:"))': '        control_layout.addWidget(QLabel("噪声注入:"))',
    'control_layout.addWidget(QLabel("鎶ュ憡椋庢牸:"))': '        control_layout.addWidget(QLabel("报告风格:"))',
    'control_layout.addWidget(QLabel("瀛楁暟:"))': '        control_layout.addWidget(QLabel("字数:"))',

    # === 修复底部按钮 ===
    'self.download_btn = QPushButton("妯℃澘涓嬭浇")': '        self.download_btn = QPushButton("模板下载")',
    'self.import_btn = QPushButton("瀵煎叆鏂囦欢")': '        self.import_btn = QPushButton("导入文件")',
    'self.export_btn = QPushButton("瀵煎嚭缁撴灉")': '        self.export_btn = QPushButton("导出结果")',
    'self.ai_report_btn = QPushButton("鐢熸垚鎶ュ憡")': '        self.ai_report_btn = QPushButton("生成报告")',

    # === 修复弹窗提示 ===
    'QMessageBox.warning(self, "鎻愮ず", "璇疯緭鍏ユ湁鏁堢殑鏁板瓧")': '            QMessageBox.warning(self, "提示", "请输入有效的数字")',
    'QMessageBox.warning(self, "鎻愮ず", "骞虫椂/鏈熶腑/鏈熸湯鍗犳瘮涔嬪拰蹇呴』绛変簬1")': '            QMessageBox.warning(self, "提示", "平时/期中/期末占比之和必须等于1")',
    "QMessageBox.warning(self, '鎻愮ず', '璇峰厛鍦ㄢ€滆绋嬭€冩牳涓庤绋嬬洰鏍囧搴斿叧绯烩€濅腑璁剧疆璇剧▼鐩爣鏁伴噺')": "            QMessageBox.warning(self, '提示', '请先在“课程考核与课程目标对应关系”中设置课程目标数量')",
    "QMessageBox.warning(self, '鎻愮ず', '璇峰厛鍦ㄢ€滄垚缁╁崰姣斺€濅腑璁剧疆姣斾緥')": "            QMessageBox.warning(self, '提示', '请先在“成绩占比”中设置比例')",
}

def fix_file(file_path):
    # 转换为绝对路径，确保能找到
    abs_path = os.path.abspath(file_path)
    print(f"尝试读取文件: {abs_path}")
    
    if not os.path.exists(abs_path):
        print("❌ 错误：找不到文件！请确认 ui_app 文件夹在当前目录下。")
        return

    # 备份
    shutil.copy(abs_path, abs_path + ".bak")
    print(f"已备份为: {os.path.basename(abs_path)}.bak")

    with open(abs_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    fixed_lines = []
    count = 0
    
    for line in lines:
        stripped = line.strip()
        # 1. 移除末尾多余的 }
        if stripped == "}":
            print("已删除末尾多余的 '}'")
            continue
            
        # 2. 修复乱码行
        if stripped in replacements:
            # 保持缩进
            indent = line[:line.find(stripped)] if stripped in line else ""
            fixed_lines.append(indent + replacements[stripped].strip() + "\n")
            count += 1
        else:
            fixed_lines.append(line)

    with open(abs_path, 'w', encoding='utf-8') as f:
        f.writelines(fixed_lines)
    
    print(f"✅ 修复完成！共修复 {count} 行代码。")
    print("现在请重新运行 main.py")

if __name__ == "__main__":
    # 指定正确的相对路径
    fix_file("ui_app/main_window.py")