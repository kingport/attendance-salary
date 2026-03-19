"""
技术部 + 李乐平、熊其享、张红亮、耿红志 规则
- 应出勤天数：
  - 耿红志、张红亮、黎钦德、李乐平、熊其享：当月天数 - 周日天数
  - 技术部其他人：固定28天
- 实际出勤天数：根据考勤表确认，无需计算加班
- 基本工资 = 固定底薪 / 应出勤天数 × 实际出勤天数
- 岗位工资 = 岗位工资 / 应出勤天数 × 实际出勤天数
- 法定节假日工资 = (底薪+岗位) / 应出勤天数 × 节假日天数
- 高温补贴：7-10月，150元/月，仅技术部本部
"""
from .base import (
    parse_punch_records, detect_shift_type, check_late_early,
    days_in_month, count_sundays
)


def calculate(employee_name, attendance_records, config, year, month):
    sunday_rest_list = config.get('tech_sunday_rest', [])

    if employee_name in sunday_rest_list:
        total_days = days_in_month(year, month)
        sundays = count_sundays(year, month)
        required_days = total_days - sundays
    else:
        required_days = config.get('tech_fixed_days', 28)

    actual_days = 0
    anomalies = []

    for record in attendance_records:
        day = record['day']
        punch_str = record.get('punch_str', '')
        times = parse_punch_records(punch_str)

        if not times:
            continue

        shift = detect_shift_type(times)
        status = check_late_early(times, shift)

        if status['single_punch']:
            anomalies.append({
                'name': employee_name,
                'date': f"{year}-{month:02d}-{day:02d}",
                'type': '单次打卡',
                'punch': punch_str
            })
            continue

        if status['is_late'] or status['is_early']:
            late_early = []
            if status['is_late']:
                late_early.append('迟到')
            if status['is_early']:
                late_early.append('早退')
            anomalies.append({
                'name': employee_name,
                'date': f"{year}-{month:02d}-{day:02d}",
                'type': '/'.join(late_early),
                'punch': punch_str
            })

        actual_days += 1

    # 从配置获取薪资
    salary_info = config.get('tech_salary', {}).get(employee_name, {})
    base_total = salary_info.get('base', 0)
    position_total = salary_info.get('position', 0)

    # 计算工资
    if required_days > 0:
        base_salary = round(base_total / required_days * actual_days, 2)
        position_salary = round(position_total / required_days * actual_days, 2)
    else:
        base_salary = 0
        position_salary = 0

    # 法定节假日工资
    holidays = config.get('holidays', 0)
    holiday_salary = 0
    if holidays > 0 and required_days > 0:
        daily_rate = (base_total + position_total) / required_days
        holiday_salary = round(daily_rate * holidays, 2)

    # 高温补贴：7-10月，仅技术部本部（不含李乐平、熊其享、张红亮、耿红志）
    high_temp_months = config.get('high_temp_months', [7, 8, 9, 10])
    non_tech_dept = ['李乐平', '熊其享', '张红亮', '耿红志']
    is_tech_dept = employee_name not in non_tech_dept
    high_temp = 150 if (month in high_temp_months and is_tech_dept) else 0

    total_salary = base_salary + position_salary + holiday_salary + high_temp

    return {
        'name': employee_name,
        'required_days': required_days,
        'actual_days': actual_days,
        'base_salary': base_salary,
        'position_salary': position_salary,
        'holidays': holidays,
        'holiday_salary': holiday_salary,
        'high_temp_subsidy': high_temp,
        'total_salary': total_salary,
        'anomalies': anomalies,
    }
