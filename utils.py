import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.cell import MergedCell

def normalize_score(score: float) -> float:
    """Normalize score to 0-100 range."""
    return max(0, min(100, score))

def get_grade_level(score: float) -> str:
    """Determine grade level based on score."""
    if score >= 90:
        return "优秀"
    elif score >= 80:
        return "良好"
    elif score >= 70:
        return "中等"
    elif score >= 60:
        return "合格"
    else:
        return "不达标"

def calculate_final_score(usual: float, midterm: float, final: float, 
                         usual_ratio: float, midterm_ratio: float, final_ratio: float) -> float:
    """Calculate final score based on ratios."""
    return usual * usual_ratio + midterm * midterm_ratio + final * final_ratio

def calculate_achievement_level(score: float) -> float:
    """Calculate achievement level as a percentage."""
    return score / 100

def adjust_column_widths(worksheet):
    """Adjust column widths based on content, handling MergedCell correctly."""
    column_widths = {}
    
    # 遍历工作表中的所有单元格
    for row in worksheet.rows:
        for cell in row:
            try:
                # 跳过 MergedCell 类型的单元格
                if cell.value and not isinstance(cell, MergedCell):
                    # 使用 cell.column 获取列号（整数），转换为列字母
                    col_letter = get_column_letter(cell.column)
                    # 计算单元格内容的字符长度（考虑中文字符）
                    cell_len = sum(2 if ord(char) > 127 else 1 for char in str(cell.value))
                    # 更新该列的最大宽度
                    column_widths[col_letter] = max(column_widths.get(col_letter, 8), cell_len + 2)
            except Exception as e:
                print(f"Error adjusting column width for cell {cell.coordinate}: {str(e)}")
                continue
    
    # 设置列宽
    for col_letter, width in column_widths.items():
        worksheet.column_dimensions[col_letter].width = min(width, 50)  # 限制最大宽度为 50