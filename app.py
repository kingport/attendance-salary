"""
宏耀考勤工资计算 - Web 应用
启动: python app.py
访问: http://localhost:5000
"""
import os
import re
import calendar
from datetime import datetime

from flask import Flask, render_template, request, send_file, jsonify

from config import *
from rules.production import calculate as calc_production
from rules.production2 import calculate as calc_production2
from rules.mold import calculate as calc_mold
from rules.quality import calculate as calc_quality
from rules.tech import calculate as calc_tech
from rules.ouyang import calculate as calc_ouyang

import openpyxl
from openpyxl import Workbook

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(__file__), 'uploads')
app.config['OUTPUT_FOLDER'] = os.path.join(os.path.dirname(__file__), 'output')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB

RULE_FUNCTIONS = {
    'production':  (calc_production,  '生产部普工'),
    'production2': (calc_production2, '生产部2(计时)'),
    'mold':        (calc_mold,        '模房'),
    'quality':     (calc_quality,     '品质部'),
    'tech':        (calc_tech,        '技术部'),
    'ouyang':      (calc_ouyang,      '欧阳宇专属'),
}


def clean_name(name):
    return re.sub(r'[（(].+?[）)]', '', name).strip()


def get_employee_rule(name, department):
    cname = clean_name(name)
    if '离职' in name:
        return None, '离职跳过'
    if cname in SKIP_EMPLOYEES:
        return None, '免计算'
    if cname in SPECIAL_EMPLOYEES:
        rule_key = SPECIAL_EMPLOYEES[cname]
        func, display = RULE_FUNCTIONS[rule_key]
        return func, f'{display}(特殊)'
    dept = department.split('\n')[0].strip() if department else ''
    rule_key = DEPARTMENT_RULE_MAP.get(dept)
    if rule_key == 'skip':
        return None, '部门跳过'
    if rule_key and rule_key in RULE_FUNCTIONS:
        func, display = RULE_FUNCTIONS[rule_key]
        return func, display
    return None, '未识别部门'


def build_config(holidays=0):
    return {
        'night_snack_subsidy': NIGHT_SNACK_SUBSIDY,
        'full_attendance_bonus': FULL_ATTENDANCE_BONUS,
        'overtime_rate': OVERTIME_RATE,
        'production_base_salary': PRODUCTION_BASE_SALARY,
        'production_standard_days': PRODUCTION_STANDARD_DAYS,
        'production_standard_hours': PRODUCTION_STANDARD_HOURS,
        'production_weekday_ot_rate': PRODUCTION_WEEKDAY_OT_RATE,
        'production_weekend_ot_rate': PRODUCTION_WEEKEND_OT_RATE,
        'production2_hourly_rate': PRODUCTION2_HOURLY_RATE,
        'quality_monthly_rest_days': QUALITY_MONTHLY_REST_DAYS,
        'quality_base_salary': QUALITY_BASE_SALARY,
        'mold_hourly_rates': MOLD_HOURLY_RATES,
        'tech_sunday_rest': TECH_SUNDAY_REST,
        'tech_fixed_days': TECH_FIXED_DAYS,
        'tech_salary': TECH_SALARY,
        'high_temp_months': HIGH_TEMP_MONTHS,
        'ouyang_base': OUYANG_BASE,
        'ouyang_position': OUYANG_POSITION,
        'ouyang_subsidy': OUYANG_SUBSIDY,
        'holidays': holidays,
        'required_workdays': PRODUCTION_STANDARD_DAYS,
    }


def read_attendance(filepath, year, month):
    wb = openpyxl.load_workbook(filepath, read_only=False)
    if '打卡时间' in wb.sheetnames:
        ws = wb['打卡时间']
    else:
        ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    title = str(rows[0][0]) if rows[0][0] else ''
    date_match = re.search(r'(\d{4})-(\d{2})-(\d{2})\s*至\s*(\d{4})-(\d{2})-(\d{2})', title)
    if date_match:
        start_day = int(date_match.group(3))
        end_day = int(date_match.group(6))
        total_days = end_day - start_day + 1
        first_day = start_day
    else:
        total_days = calendar.monthrange(year, month)[1]
        first_day = 1

    data_col_start = 6
    attendance = {}

    for row in rows[4:]:
        vals = list(row)
        name = vals[0]
        dept = vals[2]
        if not name or dept is None:
            continue
        name = str(name).strip()
        dept_str = str(dept).strip() if dept else ''

        records = []
        for day_idx in range(total_days):
            col_idx = data_col_start + day_idx
            day = first_day + day_idx
            punch_raw = vals[col_idx] if col_idx < len(vals) else None
            if punch_raw is not None and not isinstance(punch_raw, str):
                punch_raw = None
            punch_str = ''
            if punch_raw:
                punch_str = punch_raw.replace('\n', '  ').strip()
            weekday = calendar.weekday(year, month, day)
            records.append({
                'day': day,
                'punch_str': punch_str,
                'is_weekend': weekday >= 5,
            })

        attendance[name] = {
            'department': dept_str,
            'records': records,
        }

    return attendance, first_day, first_day + total_days - 1


def process_attendance(filepath, year, month, holidays=0):
    """处理考勤并返回结果数据"""
    attendance, start_day, end_day = read_attendance(filepath, year, month)
    config = build_config(holidays)
    config['year'] = year
    config['month'] = month

    results = []
    all_anomalies = []
    skipped = []
    rule_groups = {}

    for name, info in attendance.items():
        records = info['records']
        department = info['department']
        calc_func, rule_name = get_employee_rule(name, department)

        if calc_func is None:
            skipped.append({'name': name, 'reason': rule_name, 'department': department})
            continue

        cname = clean_name(name)
        result = calc_func(cname, records, config, year, month)
        result['rule'] = rule_name
        result['department'] = department
        results.append(result)

        if result.get('anomalies'):
            all_anomalies.extend(result['anomalies'])

        rule_groups.setdefault(rule_name, []).append(cname)

    return {
        'results': results,
        'anomalies': all_anomalies,
        'skipped': skipped,
        'rule_groups': rule_groups,
        'date_range': f"{year}年{month}月{start_day}日~{end_day}日",
        'total_employees': len(attendance),
    }


def save_results_excel(data, year, month):
    """保存结果到 Excel 并返回文件路径"""
    os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)
    results = data['results']

    # 工资表
    wb = Workbook()
    ws = wb.active
    ws.title = "工资计算结果"

    all_keys = set()
    for r in results:
        all_keys.update(k for k in r.keys() if k != 'anomalies')

    priority_keys = [
        'name', 'department', 'rule',
        'required_days', 'actual_days', 'workday_days', 'weekend_days',
        'total_hours', 'hourly_rate',
        'base_salary', 'position_salary',
        'weekday_ot_hours', 'weekday_ot_salary',
        'overtime_hours', 'ot_hourly_rate', 'overtime_salary',
        'weekend_hours', 'weekend_salary',
        'night_shift_days', 'night_snack_subsidy',
        'full_attendance_bonus', 'high_temp_subsidy',
        'subsidy', 'holidays', 'holiday_salary',
        'total_salary',
    ]
    headers = [k for k in priority_keys if k in all_keys]
    headers += sorted(all_keys - set(headers))

    cn = {
        'name': '姓名', 'department': '部门', 'rule': '规则类型',
        'required_days': '应出勤天数', 'actual_days': '实际出勤天数',
        'workday_days': '工作日出勤', 'weekend_days': '周末出勤',
        'total_hours': '总工时', 'hourly_rate': '时薪',
        'base_salary': '基本工资', 'position_salary': '岗位工资',
        'weekday_ot_hours': '工作日加班工时', 'weekday_ot_salary': '工作日加班工资',
        'overtime_hours': '加班工时', 'ot_hourly_rate': '加班时薪',
        'overtime_salary': '加班工资',
        'weekend_hours': '周末工时', 'weekend_salary': '周末工资',
        'night_shift_days': '夜班天数', 'night_snack_subsidy': '夜宵补贴',
        'full_attendance_bonus': '全勤奖', 'high_temp_subsidy': '高温补贴',
        'subsidy': '补贴', 'holidays': '法定节假日天数',
        'holiday_salary': '节假日工资', 'total_salary': '应发工资合计',
    }

    ws.append([cn.get(h, h) for h in headers])
    for r in results:
        ws.append([r.get(h, '') for h in headers])

    # 异常报告 sheet
    if data['anomalies']:
        ws2 = wb.create_sheet("异常报告")
        ws2.append(['姓名', '日期', '异常类型', '打卡内容'])
        for a in data['anomalies']:
            ws2.append([a['name'], a['date'], a['type'], a['punch']])

    output_file = os.path.join(app.config['OUTPUT_FOLDER'], f"工资计算_{year}年{month}月.xlsx")
    wb.save(output_file)
    return output_file


@app.route('/')
def index():
    now = datetime.now()
    return render_template('index.html', current_year=now.year, current_month=now.month)


@app.route('/calculate', methods=['POST'])
def calculate():
    file = request.files.get('file')
    if not file or not file.filename.endswith('.xlsx'):
        return jsonify({'error': '请上传 .xlsx 格式的考勤表'}), 400

    year = int(request.form.get('year', datetime.now().year))
    month = int(request.form.get('month', datetime.now().month))
    holidays = int(request.form.get('holidays', 0))

    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(filepath)

    try:
        data = process_attendance(filepath, year, month, holidays)
        output_file = save_results_excel(data, year, month)

        # 构建页面展示数据
        display_results = []
        for r in data['results']:
            display_results.append({
                'name': r.get('name', ''),
                'department': r.get('department', ''),
                'rule': r.get('rule', ''),
                'total_salary': round(r.get('total_salary', 0) or 0, 2),
                'details': {k: v for k, v in r.items()
                           if k not in ('name', 'department', 'rule', 'anomalies')},
            })

        return jsonify({
            'success': True,
            'date_range': data['date_range'],
            'total_employees': data['total_employees'],
            'calculated': len(data['results']),
            'skipped_count': len(data['skipped']),
            'anomaly_count': len(data['anomalies']),
            'results': display_results,
            'skipped': data['skipped'],
            'rule_groups': data['rule_groups'],
            'download_file': f"工资计算_{year}年{month}月.xlsx",
        })
    except Exception as e:
        return jsonify({'error': f'计算出错: {str(e)}'}), 500
    finally:
        if os.path.exists(filepath):
            os.remove(filepath)


@app.route('/download/<filename>')
def download(filename):
    filepath = os.path.join(app.config['OUTPUT_FOLDER'], filename)
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True)
    return '文件不存在', 404


if __name__ == '__main__':
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)
    app.run(debug=True, host='0.0.0.0', port=5000)
