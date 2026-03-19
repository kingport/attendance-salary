"""
基础工具函数：工时取整、迟到早退判定、打卡记录解析
"""
from datetime import datetime, timedelta
from typing import List, Tuple, Optional


def round_hours(hours: float) -> float:
    """
    工时取整：0.5H阶梯向下取整
    0~0.5H → 0, 0.5~1H → 0.5, 1~1.5H → 1, ...
    """
    if hours < 0:
        return 0
    return int(hours * 2) / 2 if int(hours * 2) / 2 <= hours else int(hours * 2) / 2


def parse_time(time_str: str) -> Optional[datetime]:
    """解析时间字符串 HH:MM 为 datetime 对象（日期部分用固定值）"""
    if not time_str or not time_str.strip():
        return None
    time_str = time_str.strip()
    try:
        return datetime.strptime(time_str, "%H:%M")
    except ValueError:
        return None


def parse_punch_records(punch_str: str) -> List[datetime]:
    """
    解析打卡记录字符串，返回时间列表
    支持格式：空格分隔的多个 HH:MM 时间
    例如: "08:03  12:10  13:30  17:31  18:30  20:30"
    """
    if not punch_str or not str(punch_str).strip():
        return []
    parts = str(punch_str).strip().split()
    times = []
    for part in parts:
        t = parse_time(part)
        if t:
            times.append(t)
    return sorted(times)


def calculate_work_hours(times: List[datetime]) -> float:
    """
    根据第一种打卡方式（统一计时）计算总工时
    取第一个打卡和最后一个打卡的时间差
    """
    if len(times) < 2:
        return 0
    first = times[0]
    last = times[-1]
    diff = (last - first).total_seconds() / 3600
    return round_hours(diff)


def check_late_early(times: List[datetime], shift_type: str) -> dict:
    """
    检查迟到早退
    shift_type: 'day' 白班, 'night' 夜班

    白班：晚于8:00上班 = 迟到，早于20:00下班 = 早退
    夜班：晚于20:00上班 = 迟到，早于8:00下班 = 早退

    返回: {
        'is_late': bool,
        'is_early': bool,
        'single_punch': bool,  # 只有一个打卡
        'no_punch': bool,      # 无打卡
    }
    """
    result = {
        'is_late': False,
        'is_early': False,
        'single_punch': False,
        'no_punch': False,
    }

    if not times:
        result['no_punch'] = True
        return result

    if len(times) == 1:
        result['single_punch'] = True
        return result

    clock_in = times[0]
    clock_out = times[-1]

    if shift_type == 'day':
        start_limit = parse_time("08:00")
        end_limit = parse_time("20:00")
        if clock_in > start_limit:
            result['is_late'] = True
        if clock_out < end_limit:
            result['is_early'] = True
    elif shift_type == 'night':
        # 夜班：上班卡晚于20:00 = 迟到，下班卡早于8:00 = 早退
        start_limit = parse_time("20:00")
        end_limit = parse_time("08:00")
        if clock_in > start_limit:
            result['is_late'] = True
        if clock_out < end_limit:
            result['is_early'] = True

    return result


def detect_shift_type(times: List[datetime]) -> str:
    """
    根据打卡时间自动判断班次类型
    如果第一个打卡时间在 12:00 之后，判定为夜班；否则为白班
    """
    if not times:
        return 'unknown'
    noon = parse_time("12:00")
    if times[0] >= noon:
        return 'night'
    return 'day'


def count_sundays(year: int, month: int) -> int:
    """计算指定月份的周日天数"""
    import calendar
    count = 0
    _, days_in_month = calendar.monthrange(year, month)
    for day in range(1, days_in_month + 1):
        if calendar.weekday(year, month, day) == 6:  # 6 = Sunday
            count += 1
    return count


def days_in_month(year: int, month: int) -> int:
    """返回指定月份的总天数"""
    import calendar
    return calendar.monthrange(year, month)[1]


def is_weekend(year: int, month: int, day: int) -> bool:
    """判断是否为周末（周六或周日）"""
    import calendar
    return calendar.weekday(year, month, day) >= 5
