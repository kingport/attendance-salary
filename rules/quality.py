"""
品质部 + 王琴 规则（排除耿红志）
- 应出勤天数 = 当月天数 - 月休天数(2天)
- 基本工资 = 固定底薪 / 应出勤天数 × 正班出勤天数
- 加班工资按底薪计算：加班工时 × (底薪/应出勤天数/8)
- 全勤奖：正班出勤天数 ≥ 应出勤天数 → 100元
- 夜宵补贴 = 夜班天数 × 15
"""
from .base import (
    parse_punch_records, detect_shift_type, check_late_early,
    round_hours, parse_time, days_in_month
)


def calculate(employee_name, attendance_records, config, year, month):
    monthly_rest = config.get('quality_monthly_rest_days', 2)
    total_days = days_in_month(year, month)
    required_days = total_days - monthly_rest  # 应出勤天数

    actual_days = 0        # 正班出勤天数
    overtime_hours = 0     # 加班工时（不区分工作日/周末）
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

        actual_days += 1

        if shift == 'night':
            night_shift_days += 1

        # 加班工时：超出正班12H的部分（白班8:00-20:00=12H, 夜班20:00-8:00=12H）
        if len(times) >= 2:
            total_work = (times[-1] - times[0]).total_seconds() / 3600
            if total_work > 12:
                overtime_hours += round_hours(total_work - 12)

    # 获取底薪
    base_salary_total = config.get('quality_base_salary', {}).get(employee_name, 0)

    # 基本工资 = 底薪 / 应出勤天数 × 出勤天数
    base_salary = round(base_salary_total / required_days * actual_days, 2) if required_days > 0 else 0

    # 加班工资按底薪算：时薪 = 底薪 / 应出勤天数 / 8H
    ot_hourly_rate = base_salary_total / required_days / 8 if required_days > 0 else 0
    overtime_salary = round(overtime_hours * ot_hourly_rate, 2)

    night_snack = night_shift_days * config.get('night_snack_subsidy', 15)
    full_attendance = config.get('full_attendance_bonus', 100) if actual_days >= required_days else 0

    total_salary = base_salary + overtime_salary + night_snack + full_attendance

    return {
        'name': employee_name,
        'required_days': required_days,
        'actual_days': actual_days,
        'base_salary': base_salary,
        'overtime_hours': overtime_hours,
        'ot_hourly_rate': round(ot_hourly_rate, 2),
        'overtime_salary': overtime_salary,
        'night_shift_days': night_shift_days,
        'night_snack_subsidy': night_snack,
        'full_attendance_bonus': full_attendance,
        'total_salary': total_salary,
        'anomalies': anomalies,
    }
