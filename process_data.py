# process_data.py
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import re
import os

def clean_phone_for_dedup(phone_str):
    if not phone_str:
        return ""
    digits = re.sub(r"\D", "", str(phone_str))
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]
    return digits

def check_apple_id_and_imessage(email_str, phone_str):
    """
    判定邮箱是否含 Apple ID，手机号是否支持 iMessage，并进行智能清洗。
    """
    email_clean = str(email_str).strip().lower() if email_str else ""
    
    # 1. 邮箱保留逻辑：有效格式，或包含 apple/icloud/me.com 等苹果生态后缀
    is_valid_format = bool(re.match(r"^[\w\.-]+@[\w\.-]+\.\w+$", email_clean))
    has_apple = "apple" in email_clean or "icloud" in email_clean or "me.com" in email_clean
    keep_email = is_valid_format or has_apple
    
    # 2. Apple ID 邮箱标绿判定
    email_has_appleid = has_apple or ("appleid" in email_clean)

    # 3. iMessage 手机号判定（针对海外/国内规范号码）
    phone_digits = clean_phone_for_dedup(phone_str)
    is_imessage = len(phone_digits) >= 10 
    
    return keep_email, email_has_appleid, is_imessage

def run_excel_processing(file_path, msg_h_text, msg_k_text, global_seen_phones, global_seen_emails):
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active

    # 1. 解析表头
    headers = [cell.value for cell in ws[1]]
    col_map = {}
    for idx, h in enumerate(headers, start=1):
        if h:
            col_map[str(h).strip()] = idx

    phone_col = col_map.get("电话") or col_map.get("Phone") or col_map.get("手机号")
    email_col = col_map.get("邮箱") or col_map.get("Email")
    name_col = col_map.get("姓名") or col_map.get("Name") or col_map.get("客户姓名")

    h_col_name = "Hello 姓名,"
    k_col_name = "Ambassador 话术"
    vip_col_name = "客户标签"
    
    # 自动追加新功能扩展列
    if h_col_name not in col_map:
        ws.cell(row=1, column=ws.max_column + 1, value=h_col_name)
        col_map[h_col_name] = ws.max_column
    if k_col_name not in col_map:
        ws.cell(row=1, column=ws.max_column + 1, value=k_col_name)
        col_map[k_col_name] = ws.max_column
    if vip_col_name not in col_map:
        ws.cell(row=1, column=ws.max_column + 1, value=vip_col_name)
        col_map[vip_col_name] = ws.max_column

    col_h_idx = col_map[h_col_name]
    col_k_idx = col_map[k_col_name]
    col_vip_idx = col_map[vip_col_name]

    # 颜色样式定义
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid") # 历史重复
    pink_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")   # 邮箱重复
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")  # AppleID / iMessage 亮绿

    name_col_letter = openpyxl.utils.get_column_letter(name_col) if name_col else "D"

    # 2. 逐行处理数据
    for row in range(2, ws.max_row + 1):
        # A. 写入话术公式
        safe_msg_h = msg_h_text.replace('"', '""')
        safe_msg_k = msg_k_text.replace('"', '""')
        ws.cell(row=row, column=col_h_idx).value = f'="Hello "&{name_col_letter}{row}&","&CHAR(10)&CHAR(10)&"{safe_msg_h}"'
        ws.cell(row=row, column=col_k_idx).value = f'="Hello "&{name_col_letter}{row}&","&CHAR(10)&CHAR(10)&"{safe_msg_k}"'

        p_val = ws.cell(row=row, column=phone_col).value if phone_col else ""
        e_val = ws.cell(row=row, column=email_col).value if email_col else ""

        keep_email, has_appleid, is_imessage = check_apple_id_and_imessage(e_val, p_val)

        # B. 手机号去重与 iMessage 亮绿标绿
        is_vip = False
        if phone_col and p_val:
            p_clean = clean_phone_for_dedup(p_val)
            if p_clean:
                if p_clean in global_seen_phones:
                    ws.cell(row=row, column=phone_col).fill = yellow_fill
                    is_vip = True  # 历史出现过，标记为老客户复购
                else:
                    global_seen_phones.add(p_clean)
                
                # iMessage 手机号强制亮绿
                if is_imessage:
                    ws.cell(row=row, column=phone_col).fill = green_fill

        # C. 邮箱去重、Apple ID 亮绿标绿
        if email_col and e_val:
            e_clean = str(e_val).strip().lower()
            if e_clean:
                if e_clean in global_seen_emails:
                    ws.cell(row=row, column=email_col).fill = pink_fill
                    is_vip = True
                else:
                    global_seen_emails.add(e_clean)

                # Apple ID 邮箱强制亮绿
                if has_appleid:
                    ws.cell(row=row, column=email_col).fill = green_fill

        # D. 写入智能客户分层标签
        tag_text = "⭐ VIP老客(复购)" if is_vip else "🌱 新客(首单)"
        ws.cell(row=row, column=col_vip_idx).value = tag_text

    # 3. 自动归档保存
    out_dir = "处理完成"
    os.makedirs(out_dir, exist_ok=True)
    out_file = os.path.join(out_dir, f"已处理+{os.path.basename(file_path)}")
    wb.save(out_file)
    return out_file
