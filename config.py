"""
考勤工资计算 - 人员分组与薪资配置
优先级：免计算 > 特殊人员 > 按部门自动匹配
"""

# ============ 1. 免计算人员（管理层/无需算工资）============
SKIP_EMPLOYEES = [
    "余耀文", "周小峰", "田宏", "王基文", "邱德芳",
    "田其斌",
]

# ============ 2. 特殊人员（不按部门，走指定规则）============
# 格式：{姓名: 规则名}
# 规则名: 'tech', 'quality', 'production', 'ouyang'
SPECIAL_EMPLOYEES = {
    "欧阳宇": "ouyang",       # 欧阳宇专属规则
    "王琴": "quality",         # 生产部但走品质部规则
    "张翱": "production",      # 模房但走生产部普工规则（区分工作日/周末）
    "耿红志": "tech",          # 品质部但走技术部规则
    "张红亮": "tech",          # 模房但走技术部规则
    "李乐平": "tech",          # 生产部但走技术部规则
    "熊其享": "tech",          # 生产部但走技术部规则
}

# ============ 3. 部门 → 规则映射 ============
DEPARTMENT_RULE_MAP = {
    "技术部": "tech",
    "品质部": "quality",
    "模房": "mold",
    "生产部": "production",
    "生产部2": "production2",
    "管理部": "skip",       # 管理部默认跳过
}

# ============ 薪资参数 ============

# --- 生产部普工 ---
# 固定底薪2100（22天*8H），工作日8H以外按17.9/H，周末23.86/H
PRODUCTION_BASE_SALARY = 2100
PRODUCTION_STANDARD_DAYS = 22
PRODUCTION_STANDARD_HOURS = 8
PRODUCTION_WEEKDAY_OT_RATE = 17.9
PRODUCTION_WEEKEND_OT_RATE = 23.86

# --- 生产部2（计时制）---
PRODUCTION2_HOURLY_RATE = 21  # 元/H

# --- 模房 ---
# 模房工价（元/H），不在此表的模房员工工价为0需补充
MOLD_HOURLY_RATES = {
    "王意": 50,
}

# --- 品质部 ---
QUALITY_MONTHLY_REST_DAYS = 2  # 月休天数
# 品质部固定底薪（加班工资按底薪计算）
QUALITY_BASE_SALARY = {
    "王琴": 7000,
    "张小平": 6000,
    "苏振鑫": 4800,
}

# --- 技术部 ---
# 应出勤=当月天数-周日天数 的人员
TECH_SUNDAY_REST = ["耿红志", "张红亮", "黎钦德", "李乐平", "熊其享"]
# 技术部其他人员 应出勤=28天
TECH_FIXED_DAYS = 28
# 固定底薪 + 岗位工资（加班工资按总和计算）
TECH_SALARY = {
    "耿红志":  {"base": 3730, "position": 4270},
    "张红亮":  {"base": 3730, "position": 8770},
    "黎钦德":  {"base": 3730, "position": 7970},
    "刘高伟":  {"base": 4020, "position": 3980},
    "陶发志":  {"base": 4020, "position": 5180},
    "王德全":  {"base": 4020, "position": 5480},
    "熊庆":    {"base": 2730, "position": 4270},
    "李乐平":  {"base": 2730, "position": 3270},
    "熊其享":  {"base": 2730, "position": 2270},
}

# --- 欧阳宇 ---
OUYANG_BASE = 2730
OUYANG_POSITION = 3970
OUYANG_SUBSIDY = 300

# ============ 通用参数 ============
OVERTIME_RATE = 17.9
NIGHT_SNACK_SUBSIDY = 15       # 夜宵补贴 元/天
FULL_ATTENDANCE_BONUS = 100    # 全勤奖 元
HIGH_TEMP_SUBSIDY = 150        # 高温补贴 元/月（7-10月，仅技术部）
HIGH_TEMP_MONTHS = [7, 8, 9, 10]

# 工作时间定义
DAY_SHIFT_START = "08:00"
DAY_SHIFT_END = "20:00"
NIGHT_SHIFT_START = "20:00"
NIGHT_SHIFT_END = "08:00"
