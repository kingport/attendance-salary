"""
模房规则
- 张翱：区分工作日/周末，同生产部普工
- 其他人：第一种打卡方式，不区分工作日/周末，整月工时 × 工价
"""
from .base import parse_punch_records, calculate_work_hours
from .production import calculate as calc_production


def calculate(employee_name, attendance_records, config, year, month):
    """
    计算模房工资
    张翱走生产部普工逻辑，其他人走计时逻辑
    """
    zhang_ao = config.get('mold_zhang_ao', '张翱')

    if employee_name == zhang_ao:
        return calc_production(employee_name, attendance_records, config, year, month)

    # 其他人：整月工时 × 工价
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

    # 工价待配置，每人可能不同
    hourly_rate = config.get('mold_hourly_rates', {}).get(employee_name, 0)

    return {
        'name': employee_name,
        'total_hours': total_hours,
        'hourly_rate': hourly_rate,
        'salary': total_hours * hourly_rate,
        'anomalies': anomalies,
    }
