# process_data.py
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import re
from itertools import zip_longest
import config  # 读取配置文件

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

def parse_and_clean_phones(phone_raw, global_seen_phones, max_limit=3):
    if not phone_raw:
        return []
    lines = [p.strip() for p in re.split(r"[\r\n]+", str(phone_raw)) if p.strip()]
    parsed_phones = []
    for idx, line in enumerate(lines):
        year_match = re.search(r"\b(20\d{2})\b", line)
        month_match = re.search(r"\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\b", line, re.IGNORECASE)
        year = int(year_match.group(1)) if year_match else 0
        month = MONTH_MAP[month_match.group(1).lower()] if month_match else 1
        date_info = ""
        if month_match and year_match:
            date_info = f"{month_match.group(1).capitalize()} {year_match.group(1)}"
        elif year_match:
            date_info = str(year_match.group(1))

        line_no_year = re.sub(r"\b20\d{2}\b", "", line)
        digits = re.sub(r"\D", "", line_no_year)
        if len(digits) == 11 and digits.startswith("1"):
            digits = digits[1:]

        formatted_phone = ""
        if len(digits) == 10:
            formatted_phone = f"+1 ({digits[:3]}) {digits[3:6]}-{digits[6:]}"
        elif len(digits) > 10:
            d10 = digits[:10]
            formatted_phone = f"+1 ({d10[:3]}) {d10[3:6]}-{d10[6:]}"
        elif len(digits) >= 7:
            formatted_phone = digits
        else:
            formatted_phone = line.split("(")[0].strip() if "(" in line else line

        if formatted_phone:
            parsed_phones.append((year, month, -idx, formatted_phone, date_info))

    parsed_phones.sort(key=lambda x: (x[0], x[1], x[2]), reverse=True)
    result = []
    for item in parsed_phones:
        phone, date_info = item[3], item[4]
        dedup_p = clean_phone_for_dedup(phone)
        if dedup_p and dedup_p not in global_seen_phones:
            global_seen_phones.add(dedup_p)
            result.append((phone, date_info))
        if len(result) == max_limit:
            break
    return result

def run_excel_processing(source_file, msg_h_text, msg_k_text, global_seen_phones, global_seen_emails):
    wb_src = openpyxl.load_workbook(source_file, data_only=True)
    ws_src = wb_src.active

    wb_tar = openpyxl.Workbook()
    ws_tar = wb_tar.active
    ws_tar.title = "Sheet1"

    headers = config.EXCEL_HEADERS
    ws_tar.append(headers)

    font_family, font_size = "宋体", 11
    color_black, color_green = "000000", "00B050"

    font_header = Font(name=font_family, size=font_size, bold=True)
    fill_header = PatternFill(start_color="F2DCDB", end_color="F2DCDB", fill_type="solid")
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin_border = Border(left=Side(style="thin", color="000000"), right=Side(style="thin", color="000000"),
                         top=Side(style="thin", color="000000"), bottom=Side(style="thin", color="000000"))

    for col_num, header in enumerate(headers, 1):
        cell = ws_tar.cell(row=1, column=col_num)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = thin_border
    ws_tar.row_dimensions[1].height = 25

    rows = list(ws_src.iter_rows(values_only=True))
    if not rows:
        return None

    header_src = [str(c).strip() if c is not None else "" for c in rows[0]]

    def get_idx(*col_names):
        for name in col_names:
            if name in header_src:
                return header_src.index(name)
            clean_target = name.lower().replace(" ", "")
            for idx, h in enumerate(header_src):
                if h.lower().replace(" ", "") == clean_target:
                    return idx
        return -1

    idx_order = get_idx("订单号", "订单编号", "Order ID", "Order No")
    idx_sku = get_idx("SKU", "尺寸", "SKU/尺寸")
    idx_time = get_idx("下单时间", "订单日期", "订单时间", "Date")
    idx_match_name = get_idx("匹配姓名")
    idx_full_name = get_idx("买家全名")
    idx_name = get_idx("买家姓名", "姓名", "Name")
    idx_email = get_idx("邮箱", "买家邮箱", "Email")
    idx_email_valid = get_idx("邮箱有效性", "邮箱状态", "Email Status", "有效性", "状态", "邮箱是否有效")
    idx_apple_id = get_idx("APPLE ID(是/否)", "APPLE ID", "AppleID", "Apple ID", "是否APPLE ID", "是否AppleID", "Apple ID(是/否)")
    idx_phone = get_idx("手机号", "电话", "手机号码", "手机", "Phone")
    idx_msg_type = get_idx("类型(iMessage/SMS)", "类型", "iMessage/SMS", "Phone Type")

    orders = []
    current_order = None

    for r_vals in rows[1:]:
        if not any(r_vals):
            continue
        raw_order_id = r_vals[idx_order] if idx_order != -1 and r_vals[idx_order] is not None else ""
        order_id_str = str(raw_order_id).strip()
        if order_id_str.endswith(".0"):
            order_id_str = order_id_str[:-2]

        if order_id_str != "":
            if current_order is None or current_order["order_id"] != order_id_str:
                if current_order:
                    orders.append(current_order)
                sku = r_vals[idx_sku] if idx_sku != -1 and r_vals[idx_sku] else ""
                time_val = r_vals[idx_time] if idx_time != -1 and r_vals[idx_time] else ""
                name_val = ""
                if idx_match_name != -1 and r_vals[idx_match_name]:
                    name_val = str(r_vals[idx_match_name]).strip()
                elif idx_full_name != -1 and r_vals[idx_full_name]:
                    name_val = str(r_vals[idx_full_name]).strip()
                elif idx_name != -1 and r_vals[idx_name]:
                    name_val = str(r_vals[idx_name]).strip()

                current_order = {
                    "order_id": order_id_str, "sku": str(sku).strip(), "time": time_val,
                    "name": name_val, "raw_phones": [], "raw_emails": [], "msg_type": ""
                }

        if current_order is None:
            continue

        email_raw = str(r_vals[idx_email]).strip() if idx_email != -1 and r_vals[idx_email] is not None else ""
        phone_raw = str(r_vals[idx_phone]).strip() if idx_phone != -1 and r_vals[idx_phone] is not None else ""
        email_valid = str(r_vals[idx_email_valid]).strip() if idx_email_valid != -1 and r_vals[idx_email_valid] is not None else ""
        apple_id = str(r_vals[idx_apple_id]).strip() if idx_apple_id != -1 and r_vals[idx_apple_id] is not None else ""

        if phone_raw:
            current_order["raw_phones"].append(phone_raw)
        if email_raw:
            for e in re.split(r"[\r\n]+", email_raw):
                clean_e = e.strip()
                if clean_e:
                    current_order["raw_emails"].append({"email": clean_e, "email_valid": email_valid, "apple_id": apple_id})

        if idx_msg_type != -1 and r_vals[idx_msg_type]:
            current_order["msg_type"] = str(r_vals[idx_msg_type]).strip()

    if current_order:
        orders.append(current_order)

    for ord_info in orders:
        all_phone_text = "\n".join(ord_info["raw_phones"])
        phones_with_dates = parse_and_clean_phones(all_phone_text, global_seen_phones, max_limit=config.DEFAULT_CONFIG["max_phones"])

        filtered_emails = []
        for e_info in ord_info["raw_emails"]:
            e_str = e_info["email"]
            e_key = e_str.lower()
            v_status = str(e_info["email_valid"]).strip().replace(" ", "").lower()
            a_status = str(e_info["apple_id"]).strip().replace(" ", "").lower()

            is_invalid = ("无效" in v_status) or ("invalid" in v_status)
            is_apple_no = ("否" in a_status) or ("no" in a_status) or (a_status == "false")

            if is_invalid and is_apple_no:
                continue
            if e_key not in global_seen_emails:
                global_seen_emails.add(e_key)
                filtered_emails.append(e_info)

        items = []
        if phones_with_dates or filtered_emails:
            for p_info, e_info in zip_longest(phones_with_dates, filtered_emails, fillvalue=None):
                p_num, p_date = ("", "") if not p_info else p_info
                e_val = e_info.get("email", "") if isinstance(e_info, dict) else ""
                a_id = e_info.get("apple_id", "") if isinstance(e_info, dict) else ""
                items.append({"phone": p_num, "date_info": p_date, "email": e_val, "apple_id": a_id, "msg_type": ord_info["msg_type"]})

        ord_info["items"] = items

    target_row_idx = 2

    def format_excel_text(text):
        parts = text.split("\n")
        quoted = ['"' + p.replace('"', '""') + '"' for p in parts]
        return '&CHAR(10)&'.join(quoted)

    formatted_h_body = format_excel_text(msg_h_text)
    formatted_k_body = format_excel_text(msg_k_text)

    for ord_info in orders:
        items = ord_info["items"]
        num_rows = max(len(items), 1)
        start_row = target_row_idx
        end_row = start_row + num_rows - 1
        date_str = ord_info["time"]

        for i in range(num_rows):
            r = start_row + i
            item = items[i] if i < len(items) else {}

            phone_val = item.get("phone", "")
            date_info_val = item.get("date_info", "")
            phone_type = item.get("msg_type", "")
            email_val = item.get("email", "")

            phone_color = color_green if phone_type.lower() == "imessage" else color_black
            a_id_status = str(item.get("apple_id", "")).strip().replace(" ", "").lower()
            is_apple_yes = ("是" in a_id_status) or ("yes" in a_id_status) or (a_id_status == "true")
            email_color = color_green if is_apple_yes else color_black

            ws_tar.cell(row=r, column=1, value=ord_info["order_id"])
            ws_tar.cell(row=r, column=2, value=ord_info["sku"])

            cell_date = ws_tar.cell(row=r, column=3, value=date_str)
            if hasattr(date_str, "strftime"):
                cell_date.number_format = "yyyy-mm-dd"

            ws_tar.cell(row=r, column=4, value=ord_info["name"])

            cell_phone = ws_tar.cell(row=r, column=5, value=phone_val)
            cell_phone.font = Font(name=font_family, size=font_size, color=phone_color)
            cell_phone.alignment = align_center
            cell_phone.border = thin_border

            cell_reply = ws_tar.cell(row=r, column=6, value=date_info_val)
            cell_reply.font = Font(name=font_family, size=font_size)
            cell_reply.alignment = align_center
            cell_reply.border = thin_border

            cell_email = ws_tar.cell(row=r, column=7, value=email_val)
            cell_email.font = Font(name=font_family, size=font_size, color=email_color)
            cell_email.alignment = align_left
            cell_email.border = thin_border

            formula_str_h = (
                f'="Hello "&D{start_row}&","&CHAR(10)&CHAR(10)'
                f'&"Thanks for purchasing our "&IF(OR(ISNUMBER(SEARCH("13X6 ST-",B{start_row})),ISNUMBER(SEARCH("13X6 SST-",B{start_row}))),SUBSTITUTE(SUBSTITUTE(B{start_row},"SST-","lace front-"),"ST-","lace front-"),B{start_row})&" wig on "&TEXT(C{start_row},"mmmm d")&". "&CHAR(10)&CHAR(10)'
                f'&{formatted_h_body}'
            )
            cell_msg = ws_tar.cell(row=r, column=8, value=formula_str_h)
            cell_msg.font = Font(name=font_family, size=font_size)
            cell_msg.alignment = align_left
            cell_msg.border = thin_border

            for col_idx in [9, 10]:
                c_empty = ws_tar.cell(row=r, column=col_idx, value="")
                c_empty.font = Font(name=font_family, size=font_size)
                c_empty.border = thin_border

            formula_str_k = (
                f'="Hello "&D{start_row}&","&CHAR(10)&CHAR(10)'
                f'&"Thank you for choosing our "&IF(OR(ISNUMBER(SEARCH("13X6 ST-",B{start_row})),ISNUMBER(SEARCH("13X6 SST-",B{start_row}))),SUBSTITUTE(SUBSTITUTE(B{start_row},"SST-","lace front-"),"ST-","lace front-"),B{start_row})&" wig on "&TEXT(C{start_row},"mmmm d")&". We hope you are loving your new hairstyle!"&CHAR(10)&CHAR(10)'
                f'&{formatted_k_body}'
            )
            cell_msg_k = ws_tar.cell(row=r, column=11, value=formula_str_k)
            cell_msg_k.font = Font(name=font_family, size=font_size)
            cell_msg_k.alignment = align_left
            cell_msg_k.border = thin_border

        if num_rows > 1:
            for col in [1, 2, 3, 4, 8, 11]:
                ws_tar.merge_cells(start_row=start_row, start_column=col, end_row=end_row, end_column=col)

        for r in range(start_row, end_row + 1):
            for col in [1, 2, 3, 4]:
                c = ws_tar.cell(row=r, column=col)
                c.font = Font(name=font_family, size=font_size)
                c.alignment = align_center
                c.border = thin_border
            for col in [8, 11]:
                c_text = ws_tar.cell(row=r, column=col)
                c_text.font = Font(name=font_family, size=font_size)
                c_text.alignment = align_left
                c_text.border = thin_border

        target_row_idx = end_row + 1

    for col_letter, width in config.COLUMN_WIDTHS.items():
        ws_tar.column_dimensions[col_letter].width = width

    return wb_tar
