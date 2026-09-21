# process_data.py
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import re
from itertools import zip_longest
import config  # 读取配置文件
import os

MONTH_MAP = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}

def clean_phone_for_dedup(phone_str):
    if not phone_str:
        return ""
    digits = re.sub(r"\D", "", str(phone_str))
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]
    return digits

def run_excel_processing(file_path, msg_h_text, msg_k_text, global_seen_phones, global_seen_emails):
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active

    # 1. 自动寻找表头与列索引
    headers = [cell.value for cell in ws[1]]
    col_map = {}
    for idx, h in enumerate(headers, start=1):
        if h:
            h_str = str(h).strip()
            col_map[h_str] = idx

    # 兼容常见的列名变体
    phone_col = col_map.get("电话") or col_map.get("Phone") or col_map.get("手机号")
    email_col = col_map.get("邮箱") or col_map.get("Email") or col_map.get("邮 箱")
    name_col = col_map.get("姓名") or col_map.get("Name") or col_map.get("客户姓名")

    # 定义高亮颜色 (黄色/浅红)
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    pink_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    # 2. 遍历数据行进行处理与去重标记
    for row in range(2, ws.max_row + 1):
        # 电话去重检查
        if phone_col:
            p_val = ws.cell(row=row, column=phone_col).value
            if p_val:
                p_clean = clean_phone_for_dedup(p_val)
                if p_clean:
                    if p_clean in global_seen_phones:
                        ws.cell(row=row, column=phone_col).fill = yellow_fill
                    else:
                        global_seen_phones.add(p_clean)

        # 邮箱去重检查
        if email_col:
            e_val = ws.cell(row=row, column=email_col).value
            if e_val:
                e_clean = str(e_val).strip().lower()
                if e_clean:
                    if e_clean in global_seen_emails:
                        ws.cell(row=row, column=email_col).fill = pink_fill
                    else:
                        global_seen_emails.add(e_clean)

    # 3. 确保输出目录存在并保存文件
    out_dir = "处理完成"
    os.makedirs(out_dir, exist_ok=True)
    out_file = os.path.join(out_dir, f"已处理+{os.path.basename(file_path)}")
    wb.save(out_file)
    return out_file
