"""
生产部2规则（计时制）
- 按小时算工资，0.5H阶梯向下取整
- 当月工时 × 21元/H
- 不区分工作日和周末
"""
from .base import parse_punch_records, calculate_work_hours


def calculate(employee_name, attendance_records, config, year, month):
    total_hours = 0
    anomalies = []

    for record in attendance_records:
        day = record['day']
        punch_str = record.get('punch_str', '')
        times = parse_punch_records(punch_str)

        if not times:
            continue

        if len(times) == 1:
            anomalies.append({
                'name': employee_name,
                'date': f"{year}-{month:02d}-{day:02d}",
                'type': '单次打卡',
                'punch': punch_str
            })
            continue

        hours = calculate_work_hours(times)
        total_hours += hours

    hourly_rate = config.get('production2_hourly_rate', 21)
    total_salary = round(total_hours * hourly_rate, 2)

    return {
        'name': employee_name,
        'total_hours': total_hours,
        'hourly_rate': hourly_rate,
        'total_salary': total_salary,
        'anomalies': anomalies,
    }
