from datetime import datetime
import glob
from itertools import zip_longest
import os
import re
import shutil
import sys
import traceback
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# 月份缩写映射表
MONTH_MAP = {
    "jan": 1,
    "feb": 2,
    "mar": 3,
    "apr": 4,
    "may": 5,
    "jun": 6,
    "jul": 7,
    "aug": 8,
    "sep": 9,
    "oct": 10,
    "nov": 11,
    "dec": 12,
}

# 本地记忆库文件名称
DB_FILE_NAME = "seen_database.txt"


def load_local_database(db_path):
    """从本地 txt 文件加载历史已处理过的电话和邮箱"""
    seen_phones = set()
    seen_emails = set()

    if os.path.exists(db_path):
        with open(db_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line.startswith("PHONE:"):
                    seen_phones.add(line.replace("PHONE:", ""))
                elif line.startswith("EMAIL:"):
                    seen_emails.add(line.replace("EMAIL:", ""))

    return seen_phones, seen_emails


def save_local_database(db_path, seen_phones, seen_emails):
    """将最新的电话和邮箱追加更新到本地 txt 文件"""
    with open(db_path, "w", encoding="utf-8") as f:
        for p in sorted(seen_phones):
            f.write(f"PHONE:{p}\n")
        for e in sorted(seen_emails):
            f.write(f"EMAIL:{e}\n")


def clean_phone_for_dedup(phone_str):
    """提取纯数字手机号，用于跨表精准对比"""
    if not phone_str:
        return ""
    digits = re.sub(r"\D", "", str(phone_str))
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]
    return digits


def parse_and_clean_phones(
    phone_raw, global_seen_phones, max_limit=3
):
    """精准分离手机号与日期，并进行本地记忆库对比"""
    if not phone_raw:
        return []

    lines = [
        p.strip() for p in re.split(r"[\r\n]+", str(phone_raw)) if p.strip()
    ]
    parsed_phones = []

    for idx, line in enumerate(lines):
        year_match = re.search(r"\b(20\d{2})\b", line)
        month_match = re.search(
            r"\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\b",
            line,
            re.IGNORECASE,
        )

        year = int(year_match.group(1)) if year_match else 0
        month = (
            MONTH_MAP[month_match.group(1).lower()] if month_match else 1
        )

        date_info = ""
        if month_match and year_match:
            date_info = (
                f"{month_match.group(1).capitalize()} {year_match.group(1)}"
            )
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
            formatted_phone = (
                line.split("(")[0].strip() if "(" in line else line
            )

        if formatted_phone:
            parsed_phones.append((year, month, -idx, formatted_phone, date_info))

    parsed_phones.sort(key=lambda x: (x[0], x[1], x[2]), reverse=True)

    result = []
    for item in parsed_phones:
        phone = item[3]
        date_info = item[4]

        # 本地全局记忆库手机号去重
        dedup_p = clean_phone_for_dedup(phone)
        if dedup_p and dedup_p not in global_seen_phones:
            global_seen_phones.add(dedup_p)
            result.append((phone, date_info))

        if len(result) == max_limit:
            break

    return result


def process_single_file(source_file, global_seen_phones, global_seen_emails):
    src_filename = os.path.basename(source_file)
    print(f"\n📄 正在读取源数据: {src_filename}")

    wb_src = openpyxl.load_workbook(source_file, data_only=True)
    ws_src = wb_src.active

    wb_tar = openpyxl.Workbook()
    ws_tar = wb_tar.active
    ws_tar.title = "Sheet1"

    headers = [
        "订单号",
        "尺寸",
        "订单日期",
        "姓名",
        "电话",
        "回复情况",
        "邮箱",
        "Hello 姓名,",
        "",
        "",
        "Ambassador 话术",
    ]
    ws_tar.append(headers)

    font_family = "宋体"
    font_size = 11

    color_black = "000000"
    color_green = "00B050"

    font_header = Font(name=font_family, size=font_size, bold=True)
    fill_header = PatternFill(
        start_color="F2DCDB", end_color="F2DCDB", fill_type="solid"
    )

    align_center = Alignment(
        horizontal="center", vertical="center", wrap_text=True
    )
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    thin_border = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000"),
    )

    for col_num, header in enumerate(headers, 1):
        cell = ws_tar.cell(row=1, column=col_num)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = thin_border

    ws_tar.row_dimensions[1].height = 25

    rows = list(ws_src.iter_rows(values_only=True))
    if not rows:
        print("  ⚠️ 源文件无有效数据！")
        return

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

    idx_email_valid = get_idx(
        "邮箱有效性", "邮箱状态", "Email Status", "有效性", "状态", "邮箱是否有效"
    )
    idx_apple_id = get_idx(
        "APPLE ID(是/否)",
        "APPLE ID",
        "AppleID",
        "Apple ID",
        "是否APPLE ID",
        "是否AppleID",
        "Apple ID(是/否)",
    )
    idx_phone = get_idx("手机号", "电话", "手机号码", "手机", "Phone")
    idx_msg_type = get_idx(
        "类型(iMessage/SMS)", "类型", "iMessage/SMS", "Phone Type"
    )

    orders = []
    current_order = None

    for r_vals in rows[1:]:
        if not any(r_vals):
            continue

        raw_order_id = (
            r_vals[idx_order]
            if idx_order != -1 and r_vals[idx_order] is not None
            else ""
        )
        order_id_str = str(raw_order_id).strip()
        if order_id_str.endswith(".0"):
            order_id_str = order_id_str[:-2]

        if order_id_str != "":
            if (
                current_order is None
                or current_order["order_id"] != order_id_str
            ):
                if current_order:
                    orders.append(current_order)

                sku = (
                    r_vals[idx_sku] if idx_sku != -1 and r_vals[idx_sku] else ""
                )
                time_val = (
                    r_vals[idx_time]
                    if idx_time != -1 and r_vals[idx_time]
                    else ""
                )

                name_val = ""
                if idx_match_name != -1 and r_vals[idx_match_name]:
                    name_val = str(r_vals[idx_match_name]).strip()
                elif idx_full_name != -1 and r_vals[idx_full_name]:
                    name_val = str(r_vals[idx_full_name]).strip()
                elif idx_name != -1 and r_vals[idx_name]:
                    name_val = str(r_vals[idx_name]).strip()

                current_order = {
                    "order_id": order_id_str,
                    "sku": str(sku).strip(),
                    "time": time_val,
                    "name": name_val,
                    "raw_phones": [],
                    "raw_emails": [],
                    "msg_type": "",
                }

        if current_order is None:
            continue

        email_raw = (
            str(r_vals[idx_email]).strip()
            if idx_email != -1 and r_vals[idx_email] is not None
            else ""
        )
        phone_raw = (
            str(r_vals[idx_phone]).strip()
            if idx_phone != -1 and r_vals[idx_phone] is not None
            else ""
        )

        email_valid = (
            str(r_vals[idx_email_valid]).strip()
            if idx_email_valid != -1 and r_vals[idx_email_valid] is not None
            else ""
        )
        apple_id = (
            str(r_vals[idx_apple_id]).strip()
            if idx_apple_id != -1 and r_vals[idx_apple_id] is not None
            else ""
        )

        if phone_raw:
            current_order["raw_phones"].append(phone_raw)

        if email_raw:
            for e in re.split(r"[\r\n]+", email_raw):
                clean_e = e.strip()
                if clean_e:
                    current_order["raw_emails"].append(
                        {
                            "email": clean_e,
                            "email_valid": email_valid,
                            "apple_id": apple_id,
                        }
                    )

        if idx_msg_type != -1 and r_vals[idx_msg_type]:
            current_order["msg_type"] = str(r_vals[idx_msg_type]).strip()

    if current_order:
        orders.append(current_order)

    for ord_info in orders:
        all_phone_text = "\n".join(ord_info["raw_phones"])
        phones_with_dates = parse_and_clean_phones(
            all_phone_text, global_seen_phones, max_limit=3
        )

        filtered_emails = []
        for e_info in ord_info["raw_emails"]:
            e_str = e_info["email"]
            e_key = e_str.lower()

            v_status = (
                str(e_info["email_valid"]).strip().replace(" ", "").lower()
            )
            a_status = (
                str(e_info["apple_id"]).strip().replace(" ", "").lower()
            )

            is_invalid = ("无效" in v_status) or ("invalid" in v_status)
            is_apple_no = (
                ("否" in a_status) or ("no" in a_status) or (a_status == "false")
            )

            # 剔除逻辑 1：无效 且 Apple ID 为 否
            if is_invalid and is_apple_no:
                continue

            # 剔除逻辑 2：本地全局历史数据库去重
            if e_key not in global_seen_emails:
                global_seen_emails.add(e_key)
                filtered_emails.append(e_info)

        items = []
        if phones_with_dates or filtered_emails:
            for p_info, e_info in zip_longest(
                phones_with_dates, filtered_emails, fillvalue=None
            ):
                p_num = ""
                p_date = ""
                if isinstance(p_info, tuple) and len(p_info) == 2:
                    p_num, p_date = p_info

                e_val = ""
                e_valid = ""
                a_id = ""
                if isinstance(e_info, dict):
                    e_val = e_info.get("email", "")
                    e_valid = e_info.get("email_valid", "")
                    a_id = e_info.get("apple_id", "")

                items.append(
                    {
                        "phone": p_num,
                        "date_info": p_date,
                        "email": e_val,
                        "email_valid": e_valid,
                        "apple_id": a_id,
                        "msg_type": ord_info["msg_type"],
                    }
                )

        ord_info["items"] = items

    target_row_idx = 2

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
            if not isinstance(email_val, (str, int, float)):
                email_val = str(email_val) if email_val is not None else ""

            phone_color = (
                color_green if phone_type.lower() == "imessage" else color_black
            )

            a_id_status = (
                str(item.get("apple_id", "")).strip().replace(" ", "").lower()
            )
            is_apple_yes = (
                ("是" in a_id_status)
                or ("yes" in a_id_status)
                or (a_id_status == "true")
            )
            email_color = color_green if is_apple_yes else color_black

            ws_tar.cell(row=r, column=1, value=ord_info["order_id"])
            ws_tar.cell(row=r, column=2, value=ord_info["sku"])

            cell_date = ws_tar.cell(row=r, column=3, value=date_str)
            if hasattr(date_str, "strftime"):
                cell_date.number_format = "yyyy-mm-dd"

            ws_tar.cell(row=r, column=4, value=ord_info["name"])

            cell_phone = ws_tar.cell(row=r, column=5, value=phone_val)
            cell_phone.font = Font(
                name=font_family, size=font_size, color=phone_color
            )
            cell_phone.alignment = align_center
            cell_phone.border = thin_border

            cell_reply = ws_tar.cell(row=r, column=6, value=date_info_val)
            cell_reply.font = Font(name=font_family, size=font_size)
            cell_reply.alignment = align_center
            cell_reply.border = thin_border

            cell_email = ws_tar.cell(row=r, column=7, value=email_val)
            cell_email.font = Font(
                name=font_family, size=font_size, color=email_color
            )
            cell_email.alignment = align_left
            cell_email.border = thin_border

            formula_str_h = (
                f'="Hello "&D{start_row}&","&CHAR(10)&CHAR(10)'
                f'&"Thanks for purchasing our "&IF(OR(ISNUMBER(SEARCH("13X6'
                f' ST-",B{start_row})),ISNUMBER(SEARCH("13X6'
                f' SST-",B{start_row}))),SUBSTITUTE(SUBSTITUTE(B{start_row},"SST-","lace'
                f' front-"),"ST-","lace front-"),B{start_row})&" wig on'
                f' "&TEXT(C{start_row},"mmmm d")&". "&CHAR(10)&CHAR(10)&"We\'d love to'
                " invite you to be our Brand Ambassador. Just submit a 5 star review"
                ' , you will get $30 or a human hair wig."&CHAR(10)&CHAR(10)&"Are you'
                ' interested?"'
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
                f'&"Thank you for choosing our "&IF(OR(ISNUMBER(SEARCH("13X6'
                f' ST-",B{start_row})),ISNUMBER(SEARCH("13X6'
                f' SST-",B{start_row}))),SUBSTITUTE(SUBSTITUTE(B{start_row},"SST-","lace'
                f' front-"),"ST-","lace front-"),B{start_row})&" wig on'
                f' "&TEXT(C{start_row},"mmmm d")&". We hope you are loving your new'
                ' hairstyle!"&CHAR(10)&CHAR(10)&"We\'d love to invite you to join our'
                " Straight Wig Brand Ambassador Program. Selected customers may have"
                " the chance to try our new wigs and share their"
                ' experience."&CHAR(10)&CHAR(10)&"Are you interested"'
            )
            cell_msg_k = ws_tar.cell(row=r, column=11, value=formula_str_k)
            cell_msg_k.font = Font(name=font_family, size=font_size)
            cell_msg_k.alignment = align_left
            cell_msg_k.border = thin_border

        if num_rows > 1:
            for col in [1, 2, 3, 4, 8, 11]:
                ws_tar.merge_cells(
                    start_row=start_row,
                    start_column=col,
                    end_row=end_row,
                    end_column=col,
                )

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

    column_widths = {
        "A": 22,
        "B": 16,
        "C": 22,
        "D": 22,
        "E": 20,
        "F": 14,
        "G": 32,
        "H": 45,
        "I": 12,
        "J": 12,
        "K": 45,
    }

    for col_letter, width in column_widths.items():
        ws_tar.column_dimensions[col_letter].width = width

    current_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = os.path.join(current_dir, "处理完成")
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    output_filename = f"已处理+{src_filename}"
    output_file = os.path.join(output_dir, output_filename)

    try:
        wb_tar.save(output_file)
        print(f"  ✅ 成功生成文件: 处理完成/{output_filename}")
    except PermissionError:
        print(
            f"  ❌ 保存失败！请关闭正在打开的文件: 处理完成/{output_filename}"
        )
        return

    archive_dir = os.path.join(current_dir, "历史原始数据汇总")
    if not os.path.exists(archive_dir):
        os.makedirs(archive_dir)

    target_archive_path = os.path.join(archive_dir, src_filename)
    shutil.move(source_file, target_archive_path)
    print(f"  📦 原始数据已归档: 历史原始数据汇总/{src_filename}")


def main():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    db_path = os.path.join(current_dir, DB_FILE_NAME)

    # 1. 加载本地持久化数据库
    global_seen_phones, global_seen_emails = load_local_database(db_path)
    print(
        f"💾 本地历史记忆库已加载: 已记录 {len(global_seen_phones)} 个电话，{len(global_seen_emails)} 个邮箱。"
    )

    # 2. 扫描当前目录下的表格
    excel_files = glob.glob(os.path.join(current_dir, "*.xlsx"))
    source_files = [
        f
        for f in excel_files
        if "已处理+" not in os.path.basename(f)
        and "今日输出" not in os.path.basename(f)
        and "新建" not in os.path.basename(f)
        and not os.path.basename(f).startswith("~$")
    ]

    if not source_files:
        print("❌ 当前目录下未放入需要处理的 Excel 原始文件！")
        return

    # 3. 处理表格
    for src_file in source_files:
        process_single_file(src_file, global_seen_phones, global_seen_emails)

    # 4. 将更新后的数据重新存回本地数据库 txt
    save_local_database(db_path, global_seen_phones, global_seen_emails)
    print(
        f"💾 本地数据库已同步更新！现共记录 {len(global_seen_phones)} 个电话，{len(global_seen_emails)} 个邮箱。"
    )
    print("🎉 处理完成！后续放新表格直接运行即可自动对比去重。")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("\n❌ 程序运行发生错误，详细堆栈如下：\n")
        traceback.print_exc()

    input("\n按回车键退出...")
