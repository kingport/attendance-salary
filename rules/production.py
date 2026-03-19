"""
生产部普工规则
- 固定底薪2100（22天*8H）
- 基本工资 = 底薪/标准天数 * 工作日出勤天数
- 工作日加班：超出20:00的部分，17.9元/H
- 周末工资：周末出勤工时 * 23.86元/H（不满12H按实际）
- 夜宵补贴：夜班 15元/天（仅第三种打卡方式/夜班）
- 全勤奖：工作日出满勤 100元
"""
from .base import (
    parse_punch_records, detect_shift_type, check_late_early,
    round_hours, parse_time, is_weekend
)


def calculate(employee_name, attendance_records, config, year, month):
    base_salary_total = config.get('production_base_salary', 2100)
    standard_days = config.get('production_standard_days', 22)
    standard_hours = config.get('production_standard_hours', 8)
    weekday_ot_rate = config.get('production_weekday_ot_rate', 17.9)
    weekend_ot_rate = config.get('production_weekend_ot_rate', 23.86)

    workday_days = 0       # 工作日出勤天数
    weekend_days = 0       # 周末出勤天数
    weekday_ot_hours = 0   # 工作日加班工时（超出标准时间）
    weekend_hours = 0      # 周末总工时
    night_shift_days = 0   # 夜班天数
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

        if shift == 'night':
            night_shift_days += 1

        weekend = is_weekend(year, month, day)

        if weekend:
            weekend_days += 1
            # 周末工时 = 最后打卡 - 最早打卡
            if len(times) >= 2:
                diff = (times[-1] - times[0]).total_seconds() / 3600
                weekend_hours += round_hours(diff)
        else:
            workday_days += 1
            # 工作日加班：超出20:00的部分
            if len(times) >= 2:
                clock_out = times[-1]
                limit = parse_time("20:00")
                if clock_out > limit:
                    extra = (clock_out - limit).total_seconds() / 3600
                    weekday_ot_hours += round_hours(extra)

    # 计算工资
    base_salary = round(base_salary_total / standard_days * workday_days, 2)
    weekday_ot_salary = round(weekday_ot_hours * weekday_ot_rate, 2)
    weekend_salary = round(weekend_hours * weekend_ot_rate, 2)
    night_snack = night_shift_days * config.get('night_snack_subsidy', 15)
    full_attendance = config.get('full_attendance_bonus', 100) if workday_days >= standard_days else 0

    total_salary = base_salary + weekday_ot_salary + weekend_salary + night_snack + full_attendance

    return {
        'name': employee_name,
        'workday_days': workday_days,
        'weekend_days': weekend_days,
        'weekday_ot_hours': weekday_ot_hours,
        'weekend_hours': weekend_hours,
        'base_salary': base_salary,
        'weekday_ot_salary': weekday_ot_salary,
        'weekend_salary': weekend_salary,
        'night_shift_days': night_shift_days,
        'night_snack_subsidy': night_snack,
        'full_attendance_bonus': full_attendance,
        'total_salary': total_salary,
        'anomalies': anomalies,
    }
