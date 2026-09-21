# config.py
# ------------------------------------------------------------------
# 配置文件：以后改话术、改列宽、改字段名，只需要改这个文件即可
# ------------------------------------------------------------------

# 1. 默认 H 列（奖励邀请）话术模板
DEFAULT_MSG_H = """We'd love to invite you to be our Brand Ambassador. Just submit a 5 star review , you will get $30 or a human hair wig.

Are you interested?"""

# 2. 默认 K 列（Ambassador 体验）话术模板
DEFAULT_MSG_K = """We'd love to invite you to join our Straight Wig Brand Ambassador Program. Selected customers may have the chance to try our new wigs and share their experience.

Are you interested"""

# 3. 导出的 Excel 11 列表头结构
EXCEL_HEADERS = [
    "订单号", "尺寸", "订单日期", "姓名", "电话", 
    "回复情况", "邮箱", "Hello 姓名,", "", "", "Ambassador 话术"
]

# 4. 导出的 Excel 各列列宽设置
COLUMN_WIDTHS = {
    "A": 22, "B": 16, "C": 22, "D": 22, "E": 20, 
    "F": 14, "G": 32, "H": 45, "I": 12, "J": 12, "K": 45
}

# 5. 核心处理参数规则
DEFAULT_CONFIG = {
    "max_phones": 3,               # 每个订单最多保留的去重电话数
    "enable_email_dedup": True,    # 是否开启历史邮箱去重
    "enable_phone_dedup": True,    # 是否开启历史电话去重
}
