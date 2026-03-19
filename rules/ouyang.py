"""
欧阳宇专属规则
- 应出勤天数：当月天数 - 周日天数
- 实际出勤天数：根据考勤表，只有1个打卡也视为出勤，无需计算加班
- 基本工资 = 固定底薪 / 应出勤天数 × 实际出勤天数
- 岗位工资 = 岗位工资 / 应出勤天数 × 实际出勤天数
- 法定节假日工资 = (底薪+岗位) / 应出勤天数 × 节假日天数
- 补贴 = 300 / 应出勤天数 × 实际出勤天数
- 无需计算其他奖金和补贴
"""
from .base import (
    parse_punch_records, days_in_month, count_sundays
)


def calculate(employee_name, attendance_records, config, year, month):
    total_days = days_in_month(year, month)
    sundays = count_sundays(year, month)
    required_days = total_days - sundays

    actual_days = 0
    anomalies = []

    for record in attendance_records:
        day = record['day']
        punch_str = record.get('punch_str', '')
        times = parse_punch_records(punch_str)

        if not times:
            continue

        # 欧阳宇：只有1个打卡也视为出勤
        actual_days += 1

        if len(times) == 1:
            anomalies.append({
                'name': employee_name,
                'date': f"{year}-{month:02d}-{day:02d}",
                'type': '单次打卡(视为出勤)',
                'punch': punch_str
            })

    base_total = config.get('ouyang_base', 2730)
    position_total = config.get('ouyang_position', 3970)
    subsidy_total = config.get('ouyang_subsidy', 300)

    if required_days > 0:
        base_salary = round(base_total / required_days * actual_days, 2)
        position_salary = round(position_total / required_days * actual_days, 2)
        subsidy = round(subsidy_total / required_days * actual_days, 2)
    else:
        base_salary = 0
        position_salary = 0
        subsidy = 0

    # 法定节假日工资
    holidays = config.get('holidays', 0)
    holiday_salary = 0
    if holidays > 0 and required_days > 0:
        daily_rate = (base_total + position_total) / required_days
        holiday_salary = round(daily_rate * holidays, 2)

    total_salary = base_salary + position_salary + subsidy + holiday_salary

    return {
        'name': employee_name,
        'required_days': required_days,
        'actual_days': actual_days,
        'base_salary': base_salary,
        'position_salary': position_salary,
        'subsidy': subsidy,
        'holidays': holidays,
        'holiday_salary': holiday_salary,
        'total_salary': total_salary,
        'anomalies': anomalies,
    }
