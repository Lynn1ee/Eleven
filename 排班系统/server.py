"""排班系统 Web 服务器 — Flask，支持局域网多用户"""
import json
import os
import sys
import socket
import io
from flask import Flask, request, jsonify, send_file, render_template_string

# 将当前目录加入路径
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)

from schedule_engine import ScheduleEngine, DEFAULT_CONFIG
from target_calculator import TargetCalculator

app = Flask(__name__)

# 数据持久化目录
_DATA_DIR = os.path.join(SCRIPT_DIR, 'data')
_SAVED_DATA_FILE = os.path.join(_DATA_DIR, 'uploaded_raw_data.xlsx')
_CONFIG_FILE = os.path.join(_DATA_DIR, 'schedule_config.json')


def _load_config():
    """加载前端人员分组配置"""
    if os.path.exists(_CONFIG_FILE):
        try:
            with open(_CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def _save_config(config):
    """保存前端人员分组配置到磁盘"""
    os.makedirs(_DATA_DIR, exist_ok=True)
    with open(_CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=2)

# 内存中缓存最后一次生成的 Excel
_last_excel = None
_last_filename = 'schedule.xlsx'

# 目标人数计算缓存
_last_calculator = None
_last_coefficients = None
_last_daily_targets = None
_last_theoretical = None  # 未取整的理论目标值


def _save_uploaded_data(file_bytes):
    """持久化上传的原始数据到磁盘"""
    os.makedirs(_DATA_DIR, exist_ok=True)
    with open(_SAVED_DATA_FILE, 'wb') as f:
        f.write(file_bytes)


def _load_saved_data():
    """启动时加载持久化的原始数据"""
    if os.path.exists(_SAVED_DATA_FILE):
        try:
            with open(_SAVED_DATA_FILE, 'rb') as f:
                return f.read()
        except Exception:
            pass
    return None


def _restore_target_state():
    """启动时尝试恢复目标数据状态"""
    global _last_calculator, _last_coefficients, _last_daily_targets
    file_bytes = _load_saved_data()
    if file_bytes is None:
        return
    try:
        _last_calculator = TargetCalculator(file_bytes)
        _last_coefficients = _last_calculator.compute_all_coefficients()
        print(f"  [server] 已恢复上次上传的数据 ({_last_calculator.source_year}年{_last_calculator.source_month}月)")
    except Exception as e:
        print(f"  [server] 恢复数据失败: {e}")


def _get_lan_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return '127.0.0.1'


@app.route('/')
def index():
    """返回排班操作界面"""
    html_path = os.path.join(SCRIPT_DIR, 'templates', 'index.html')
    with open(html_path, 'r', encoding='utf-8') as f:
        return render_template_string(f.read())


@app.route('/api/schedule/default-config', methods=['GET'])
def default_config():
    """返回默认配置（含已保存的人员分组）"""
    config = dict(DEFAULT_CONFIG)
    saved = _load_config()
    if saved.get('groups'):
        config['groups'] = saved['groups']
    if saved.get('nightGroups'):
        config['night_groups'] = saved['nightGroups']
    if saved.get('restRequests'):
        config['rest_requests'] = saved['restRequests']
    return jsonify({'success': True, 'config': config})


@app.route('/api/schedule/config', methods=['GET'])
def get_config():
    """获取已保存的前端人员分组配置"""
    config = _load_config()
    return jsonify({'success': True, 'config': config})


@app.route('/api/schedule/config', methods=['POST'])
def save_config():
    """保存前端人员分组配置"""
    try:
        data = request.get_json(force=True)
        _save_config(data)
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@app.route('/api/calendar/info', methods=['GET'])
def calendar_info():
    """返回指定月份的日历信息（节假日、工作日、标准工时等）"""
    try:
        year = int(request.args.get('year', 2026))
        month = int(request.args.get('month', 7))
        engine = ScheduleEngine(year=year, month=month)
        holidays = engine.get_all_holidays()
        workday_count = sum(1 for d in range(1, engine.num_days + 1) if engine.is_workday(d))
        return jsonify({
            'success': True,
            'num_days': engine.num_days,
            'start_weekday': engine.start_weekday,
            'workday_count': workday_count,
            'rest_day_count': engine.num_days - workday_count,
            'standard_hours': workday_count * 8,
            'holidays': holidays,
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@app.route('/api/targets/upload', methods=['POST'])
def upload_raw_data():
    """上传原始进线数据 Excel"""
    global _last_calculator, _last_coefficients, _last_daily_targets
    try:
        file = request.files.get('file')
        if not file or file.filename == '':
            return jsonify({'success': False, 'error': '请选择文件'})

        file_bytes = file.read()
        _save_uploaded_data(file_bytes)  # 持久化到磁盘
        calculator = TargetCalculator(file_bytes)
        _last_calculator = calculator
        _last_coefficients = None
        _last_daily_targets = None

        daily_totals = calculator.get_daily_totals()
        monthly_total = sum(daily_totals.values())

        return jsonify({
            'success': True,
            'source_info': {
                'year': calculator.source_year,
                'month': calculator.source_month,
                'num_days': calculator.source_num_days,
                'start_weekday': calculator.source_start_weekday,
                'monthly_total': round(monthly_total, 2),
                'monthly_avg': round(monthly_total / calculator.source_num_days, 2),
            },
        })
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)})


@app.route('/api/targets/coefficients', methods=['GET'])
def get_coefficients():
    """获取当前已上传数据的系数计算结果"""
    global _last_calculator, _last_coefficients
    try:
        if _last_calculator is None:
            return jsonify({'success': False, 'error': '请先上传原始数据文件'})

        if _last_coefficients is None:
            _last_coefficients = _last_calculator.compute_all_coefficients()

        return jsonify({
            'success': True,
            'weekly_coeffs': _last_coefficients['weekly_coeffs'],
            'weekday_coeffs': _last_coefficients['weekday_coeffs'],
            'source_monthly_avg': _last_coefficients['source_monthly_avg'],
            'source_total': _last_coefficients['source_total'],
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@app.route('/api/targets/calculate', methods=['POST'])
def calculate_targets():
    """计算目标月每日目标人数"""
    global _last_calculator, _last_coefficients, _last_daily_targets
    try:
        if _last_calculator is None:
            return jsonify({'success': False, 'error': '请先上传原始数据文件'})

        data = request.get_json(force=True)
        year = data.get('year', 2026)
        month = data.get('month', 7)
        online_count = data.get('online_count', 26)
        max_cap = data.get('max_daily_cap')
        min_floor = data.get('min_daily_floor')

        if _last_coefficients is None:
            _last_coefficients = _last_calculator.compute_all_coefficients()

        result = _last_calculator.compute_daily_targets(
            target_year=year, target_month=month,
            online_count=online_count,
            max_daily_cap=max_cap,
            min_daily_floor=min_floor,
        )
        _last_daily_targets = result['daily_targets']

        # 构建详情
        from datetime import date
        wd_names = ['周一', '周二', '周三', '周四', '周五', '周六', '周日']
        num_days = len(result['daily_targets'])
        start_wd = date(year, month, 1).weekday()
        # 获取调试信息（实例上暂存）
        debug_info = getattr(_last_calculator, '_debug_alloc', [])
        details = []
        for day in range(1, num_days + 1):
            wd = (start_wd + day - 1) % 7
            dbg = debug_info[day - 1] if day - 1 < len(debug_info) else {}
            details.append({
                'day': day,
                'label': f'{month}/{day} {wd_names[wd]}',
                'theoretical': round(dbg.get('theoretical', 0), 2),
                'rounded': dbg.get('rounded', 0),
                'dist_to_boundary': round(dbg.get('dist', 0), 2),
                'was_adjusted': dbg.get('adjusted', False),
                'target': result['daily_targets'][day - 1],
            })
        _last_theoretical = [d['theoretical'] for d in details]

        return jsonify({
            'success': True,
            'daily_targets': result['daily_targets'],
            'total_person_trips': result['total_person_trips'],
            'workday_count': result['workday_count'],
            'details': details,
            'weekly_coeffs': _last_coefficients['weekly_coeffs'],
            'weekday_coeffs': _last_coefficients['weekday_coeffs'],
        })
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)})


@app.route('/api/targets/current', methods=['GET'])
def current_target_state():
    """返回当前缓存的目标数据状态（用于前端刷新恢复）"""
    global _last_calculator, _last_coefficients, _last_daily_targets
    if _last_calculator is None:
        return jsonify({'success': True, 'has_data': False})

    daily_totals = _last_calculator.get_daily_totals()
    monthly_total = sum(daily_totals.values())

    result = {
        'success': True,
        'has_data': True,
        'source_info': {
            'year': _last_calculator.source_year,
            'month': _last_calculator.source_month,
            'num_days': _last_calculator.source_num_days,
            'start_weekday': _last_calculator.source_start_weekday,
            'monthly_total': round(monthly_total, 2),
            'monthly_avg': round(monthly_total / _last_calculator.source_num_days, 2),
        },
        'has_coefficients': _last_coefficients is not None,
        'has_targets': _last_daily_targets is not None,
    }
    if _last_coefficients is not None:
        result['weekly_coeffs'] = _last_coefficients['weekly_coeffs']
        result['weekday_coeffs'] = _last_coefficients['weekday_coeffs']
    if _last_daily_targets is not None:
        result['daily_targets'] = _last_daily_targets
    if _last_theoretical is not None:
        result['theoretical'] = _last_theoretical
    return jsonify(result)


@app.route('/api/targets/download-template', methods=['GET'])
def download_template():
    """下载系数模板格式 Excel"""
    global _last_calculator
    try:
        if _last_calculator is None:
            return jsonify({'success': False, 'error': '请先上传原始数据文件'})

        excel_bytes = _last_calculator.export_template_excel()
        filename = f'{_last_calculator.source_year}年{_last_calculator.source_month}月_系数模板.xlsx'
        return send_file(
            excel_bytes,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename,
        )
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@app.route('/api/schedule/generate', methods=['POST'])
def generate():
    """生成排班"""
    global _last_excel, _last_filename

    try:
        import time as _st
        _t0 = _st.time()
        data = request.get_json(force=True)
        print(f"  [server] 收到请求, 开始创建引擎...", flush=True)

        year = data.get('year', 2026)
        month = data.get('month', 7)
        daily_targets = data.get('daily_targets', [])

        fenxiao = data.get('fenxiao', [])
        night_shift = data.get('night_shift', [])
        buru_ban = data.get('buru_ban', [])
        buru_support = data.get('buru_support', [])
        changbingjia = data.get('changbingjia', [])
        online_staff = data.get('online_staff', [])
        rest_requests = data.get('rest_requests', {})

        # 自动加载上月班表做跨月衔接
        prev_month_data = None
        prev_year = year if month > 1 else year - 1
        prev_month = month - 1 if month > 1 else 12
        prev_file = os.path.join(SCRIPT_DIR, f'{prev_month}月份班表.xlsx')
        if os.path.exists(prev_file):
            try:
                import calendar as _cal
                prev_num_days = _cal.monthrange(prev_year, prev_month)[1]
                prev_month_data = ScheduleEngine.load_prev_month_data(prev_file, prev_num_days)
                print(f"  [server] 已加载上月班表: {prev_file}", flush=True)
            except Exception as e:
                print(f"  [server] 加载上月班表失败: {e}", flush=True)

        engine = ScheduleEngine(
            year=year, month=month, daily_targets=daily_targets,
            fenxiao=fenxiao, night_shift=night_shift,
            buru_ban=buru_ban, buru_support=buru_support,
            changbingjia=changbingjia, online_staff=online_staff,
            theoretical_targets=_last_theoretical,
            prev_month_data=prev_month_data,
            rest_requests=rest_requests,
        )
        print(f"  [server] 引擎创建完成 ({_st.time()-_t0:.2f}s)", flush=True)

        if len(daily_targets) != engine.num_days:
            return jsonify({'success': False, 'error': f'每日目标数组长度({len(daily_targets)})与天数({engine.num_days})不匹配'})

        schedules, fenxiao_rest = engine.generate()
        print(f"  [server] generate() 完成 ({_st.time()-_t0:.2f}s)", flush=True)
        verification = engine.verify(schedules, fenxiao_rest)
        print(f"  [server] verify() 完成 ({_st.time()-_t0:.2f}s)", flush=True)
        preview = engine.preview_data(schedules)
        print(f"  [server] preview_data() 完成 ({_st.time()-_t0:.2f}s)", flush=True)

        # 生成 Excel 到内存
        excel_bytes = engine.create_excel(schedules)
        _last_excel = excel_bytes
        _last_filename = f'{month}月份班表.xlsx'
        print(f"  [server] create_excel() 完成 ({_st.time()-_t0:.2f}s)", flush=True)

        # 收集节假日信息（含日期、名称、三薪标记）
        holidays = engine.get_all_holidays()

        # 计算工作日数（不含周末和节假日，含调休补班）
        workday_count = sum(1 for d in range(1, engine.num_days + 1) if engine.is_workday(d))
        standard_hours = workday_count * 8

        return jsonify({
            'success': True,
            'verification': verification,
            'preview': preview,
            'diag': getattr(engine, '_diag', {}),
            'holidays': holidays,
            'num_days': engine.num_days,
            'start_weekday': engine.start_weekday,
            'workday_count': workday_count,
            'rest_day_count': engine.num_days - workday_count,
            'standard_hours': standard_hours,
            'download_ready': True,
            'filename': _last_filename,
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)})


@app.route('/api/schedule/download', methods=['GET'])
def download():
    """下载最后一次生成的 Excel"""
    global _last_excel, _last_filename
    if _last_excel is None:
        return jsonify({'success': False, 'error': '请先生成排班'})
    _last_excel.seek(0)
    return send_file(
        _last_excel,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=_last_filename,
    )


def main():
    _restore_target_state()  # 启动时恢复上次上传的数据
    port = 8897
    lan_ip = _get_lan_ip()
    print(f'\n  排班系统已启动')
    print(f'  本地访问: http://localhost:{port}')
    print(f'  局域网访问: http://{lan_ip}:{port}')
    print(f'  按 Ctrl+C 停止\n')
    app.run(host='0.0.0.0', port=port, threaded=True, debug=False)


if __name__ == '__main__':
    main()
