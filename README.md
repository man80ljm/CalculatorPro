# 逆向推算功能完整实现 - 集成指南

## 修改概览

本次修改实现了以下功能：
1. **逆向模式强制依赖关系表** - 未填写关系表时禁止下载逆向模板
2. **逆向导出包含正向全部文件** - 表2(统计表) + 表5(达成度) + Word版本
3. **额外输出"二维正向成绩表"** - 把逆向反推的方法级分数以正向模板格式展示
4. **完整计算链** - 逆向流程与正向流程输出一致

---

## 文件修改清单

### 1. excel_templates.py（完整替换）

**文件位置**: `io_app/excel_templates.py`

**修改内容**:
- `create_reverse_template()` 函数现在**强制要求** `relation_json_path` 参数
- 如果未提供关系表，会抛出 `ValueError` 异常

**直接使用**: `/home/claude/modified_files/excel_templates.py` 可以直接替换原文件

---

### 2. main_window.py（局部修改）

**文件位置**: `ui_app/main_window.py`

**修改位置**: `open_template_download()` 方法（约第805-847行）

**修改内容**:
- 逆向模式下载模板前，增加对 `self.relation_payload` 的强制检查
- 如果未填写关系表，弹出警告对话框并阻止下载

**集成方法**: 
找到原来的 `open_template_download` 方法，替换为以下代码：

```python
def open_template_download(self):
    """下载模板并保存到 outputs 目录"""
    dialog = TemplateDownloadDialog(self)
    if not dialog.exec():
        return
    count = dialog.student_count
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    
    try:
        # ===== 正向模式 =====
        if self.tabs.currentIndex() == 0:
            if not self.relation_payload:
                QMessageBox.warning(
                    self,
                    "提示",
                    "请先填写"课程考核与课程目标对应关系"。",
                )
                return
            outputs_dir = os.path.join(base_dir, "outputs")
            os.makedirs(outputs_dir, exist_ok=True)
            relation_json_path = os.path.join(outputs_dir, "relation_table.json")
            with open(relation_json_path, "w", encoding="utf-8") as f:
                json.dump(self.relation_payload, f, ensure_ascii=False, indent=2)
            output_path = create_forward_template(base_dir, count, relation_json_path)
        
        # ===== 逆向模式（修改部分） =====
        else:
            # 【新增】强制检查关系表
            if not self.relation_payload:
                QMessageBox.warning(
                    self,
                    "提示",
                    "逆向模式必须先填写"课程考核与课程目标对应关系表"。\n\n"
                    "逆向推算需要知道考核环节和考核方式的结构才能正确反推成绩。",
                )
                return
            
            # 保存关系表到文件
            outputs_dir = os.path.join(base_dir, "outputs")
            os.makedirs(outputs_dir, exist_ok=True)
            relation_json_path = os.path.join(outputs_dir, "relation_table.json")
            with open(relation_json_path, "w", encoding="utf-8") as f:
                json.dump(self.relation_payload, f, ensure_ascii=False, indent=2)
            
            # 创建逆向模板（现在会强制使用关系表）
            output_path = create_reverse_template(base_dir, count, relation_json_path)

        QMessageBox.information(
            self,
            "下载成功",
            f"模板已生成：{output_path}",
        )
    except Exception as exc:
        QMessageBox.critical(
            self,
            "错误",
            f"模板生成失败：{str(exc)}",
        )
```

---

### 3. excel_calc.py（局部修改）

**文件位置**: `core_app/excel_calc.py`

**修改内容**:
1. **新增方法** `_generate_forward_score_table()` - 生成二维正向成绩表
2. **替换方法** `process_reverse_grades()` - 完整的逆向流程实现

**集成方法**:

#### 3.1 添加新方法 `_generate_forward_score_table`

在 `ExcelCalcMixin` 类中添加以下方法（建议放在 `process_reverse_grades` 之前）：

```python
def _generate_forward_score_table(self, detail_rows: list, links: list) -> str:
    """
    根据逆向推算的明细数据，生成二维正向成绩表。
    格式：行=学生，列=考核环节下的方法，带两行表头/合并单元格
    """
    import openpyxl
    from openpyxl.styles import Alignment, Border, Side, Font
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "正向成绩表"
    
    # 构建表头结构
    ws.cell(row=1, column=1, value="姓名")
    ws.cell(row=2, column=1, value="姓名")
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)
    
    col = 2
    method_col_map = {}
    
    for link in links:
        link_name = link.get("name", "")
        methods = link.get("methods", []) or [{"name": "无"}]
        start_col = col
        
        for method in methods:
            m_name = method.get("name", "无")
            ws.cell(row=2, column=col, value=m_name)
            method_col_map[f"{link_name}-{m_name}"] = col
            col += 1
        
        end_col = col - 1
        ws.cell(row=1, column=start_col, value=link_name)
        if end_col > start_col:
            ws.merge_cells(start_row=1, start_column=start_col, end_row=1, end_column=end_col)
    
    # 填充学生数据
    row_idx = 3
    for row_data in detail_rows:
        name = row_data.get("姓名", "")
        ws.cell(row=row_idx, column=1, value=name)
        
        for key, col_num in method_col_map.items():
            score = row_data.get(key, 0)
            ws.cell(row=row_idx, column=col_num, value=round(score, 1) if isinstance(score, (int, float)) else score)
        
        row_idx += 1
    
    # 应用样式
    align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin = Side(style='thin')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_font = Font(bold=True)
    
    for r in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in r:
            cell.alignment = align
            cell.border = border
            if cell.row <= 2:
                cell.font = header_font
    
    # 调整列宽
    ws.column_dimensions['A'].width = 10
    for c in range(2, col):
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 12
    
    # 保存文件
    output_dir = os.path.join(os.path.abspath(os.path.dirname(__file__)), "..", "outputs")
    os.makedirs(output_dir, exist_ok=True)
    safe_name = self._safe_filename(self.course_name_input.text())
    output_path = os.path.join(output_dir, f"{safe_name}正向成绩表（逆向生成）.xlsx")
    wb.save(output_path)
    
    return output_path
```

#### 3.2 替换 `process_reverse_grades` 方法

用 `/home/claude/modified_files/excel_calc_patch.py` 中的 `process_reverse_grades` 方法替换原有方法。

---

## 逆向流程输出文件清单

修改后，逆向模式导出将生成以下文件：

| 文件名 | 说明 |
|--------|------|
| `{课程名}成绩明细.xlsx` | 包含3个Sheet: 成绩明细、课程成绩统计、课程目标达成情况评价结果 |
| `{课程名}正向成绩表（逆向生成）.xlsx` | **【新增】** 二维正向成绩表，格式与正向模板一致 |
| `2.课程成绩统计表.docx` | 表2的Word版本 |
| `5.基于考核结果的课程目标达成情况评价结果表.docx` | 表5的Word版本 |

---

## 测试建议

1. **测试关系表检查**
   - 不填写关系表，点击逆向模式的"模板下载"，应弹出警告
   - 填写关系表后，应能正常下载

2. **测试完整流程**
   - 填写关系表
   - 下载逆向模板
   - 填入环节总分
   - 导入并导出
   - 检查outputs目录是否生成所有预期文件

3. **测试正向成绩表格式**
   - 打开"正向成绩表（逆向生成）.xlsx"
   - 检查表头是否为二级结构（环节-方法）
   - 检查分数是否合理

---

## 后续可优化项

1. **真实性增强**（可作为高级设置）
   - 一致性系数：控制同一学生各方法分数的相关性
   - 方法难度偏移：不同考核方式有不同的平均分倾向
   - 分数锚定：限制平时分下限等

2. **UI优化**
   - 逆向模式Tab增加"必须先填写关系表"的提示文字
   - 导出完成后显示生成的文件列表