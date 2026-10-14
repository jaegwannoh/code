# -*- coding: utf-8 -*-
"""
app.py - Daily Work 웹 프로그램의 메인 파일 (파이썬 2.7 버전)

실행 방법:  run.bat 더블클릭  (또는 python app.py)
접속 주소:  내 PC  → http://localhost:5000
           다른 PC → http://내PC의IP:5000

구조
  - '@app.route(주소)' 가 붙은 함수 하나가 웹페이지 하나(또는 버튼 동작 하나)입니다.
  - 화면 모양(HTML)은 templates 폴더, 디자인(CSS)과 정렬 기능(JS)은 static 폴더에 있습니다.
  - DB 관련 코드는 db.py, 엑셀은 excel.py, 백업은 backup.py 에 나눠 두었습니다.
  - 파이썬 2.7 에서 한글을 안전하게 다루는 도우미는 compat.py 에 모아 두었습니다.

[Py2.7 차이] 표시가 있는 곳은 파이썬 3 버전(daily_work)과 다르게 쓴 부분입니다.
"""
# [Py2.7 차이] 이 줄이 있으면 '한글' 이라고만 써도 u'한글' (unicode) 이 됩니다.
#              모든 .py 파일 맨 위에 꼭 있어야 합니다.
from __future__ import unicode_literals

import os
import sys

# 백그라운드 실행(start_hidden.vbs → run.bat hidden)이면 모든 메시지를 server.log 에 씁니다.
# flask 를 불러오기 전에 해 둬야 flask 를 불러오다 난 에러도 기록됩니다.
import compat
if '--log' in sys.argv:
    compat.log_to_file(os.path.join(compat.path_text(os.path.dirname(os.path.abspath(__file__))),
                                    'server.log'))

import re
import sqlite3
from collections import Counter
from datetime import date, datetime, timedelta
from functools import wraps

from flask import (Flask, Response, abort, flash, g, redirect, render_template, request,
                   session, url_for)

try:
    from markupsafe import Markup, escape       # 보통은 여기에 있고
except ImportError:
    from jinja2 import Markup, escape           # 아주 옛 버전이면 여기에

import backup
import db
import excel
from compat import ensure_dir, error_text, log, quote, text_type, to_text

app = Flask(__name__)
app.secret_key = 'daily-work-log'           # 로그인 상태 저장용 (내부망이라 단순하게)
app.teardown_appcontext(db.close_db)        # 요청이 끝나면 DB 연결 닫기

PORT = 5000
TEAMS = ['3팀', '5팀']
GRADES = ['4급', '5급', '6급', '7급', '8급', '9급']
LEAVE_KINDS = ['연가', '출장', '교육', '기타']
WEEKDAYS = '월화수목금토일'
XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
OLD_ANALYSIS_DAYS = 90      # 분석중이 이 일수를 넘으면 빨간색으로 표시 (종료를 잊은 건 찾기)
PAGE_SIZES = [50, 100, 200] # 검색 결과 한 페이지에 보여줄 건수 (첫 번째가 기본)
PRINT_WARN = 2000           # 검색 결과가 이보다 많으면 인쇄 전에 확인창


# ===========================================================================
# 공통 도우미 함수
# ===========================================================================
def parse_date(text, default=None):
    """'2026-10-13' 같은 문자열 → date. 잘못된 값이면 default(없으면 오늘)."""
    try:
        return datetime.strptime(text, '%Y-%m-%d').date()
    except (TypeError, ValueError):
        return default or date.today()


def date_context():
    """날짜 이동(◀ ▶) 이 있는 화면에서 공통으로 쓰는 값들"""
    d = parse_date(request.args.get('date'))
    return {
        'd': d.isoformat(),
        'prev_d': (d - timedelta(days=1)).isoformat(),
        'next_d': (d + timedelta(days=1)).isoformat(),
        # [Py2.7 차이] f'{d.year}년 ...' (f-string) 대신 '{}년'.format(...)
        'd_label': '{}년 {}월 {}일 ({})'.format(d.year, d.month, d.day, WEEKDAYS[d.weekday()]),
        'today': date.today().isoformat(),
    }


def team_from_url(team):
    """주소의 '3' / '5' → '3팀' / '5팀'"""
    name = team + '팀'
    if name not in TEAMS:
        abort(404)
    return name


def go_back(default='index'):
    """버튼 처리 후 원래 보던 화면으로 돌아가기 (폼의 next 값 사용)"""
    nxt = request.form.get('next') or request.args.get('next') or ''
    if nxt.startswith('/') and not nxt.startswith('//'):
        return redirect(nxt)
    return redirect(url_for(default))


def form_text(name):
    """폼에서 글자를 꺼내 앞뒤 공백 제거"""
    return request.form.get(name, '').strip()


def query_string():
    """주소의 ? 뒤 부분 (예: 'q=abc&field=all')"""
    # [Py2.7 차이] request.query_string 은 바이트라서 글자로 바꿔 줍니다.
    return request.query_string.decode('utf-8', 'replace')


def collect_day(d):
    """하루치 업무일지 전체 (메인 화면, 엑셀, 인쇄에서 같이 사용)"""
    return {
        'leaves': db.leaves_on(d),
        'teams': dict((t, {'works': db.daily_works(d, t),
                           'missing': db.missing_users(d, t)}) for t in TEAMS),
        'signals': db.signals_on(d),
        'ongoing': db.analyses_ongoing(d),
        'finished': db.analyses_finished(d),
    }


def xlsx_response(data, filename, ascii_filename):
    """
    엑셀 파일을 브라우저로 내려보냅니다 (다운로드).
    [Py2.7 차이] 옛 Flask 의 send_file 은 한글 파일 이름을 못 보내서 에러가 납니다.
                 그래서 다운로드 머리글(Content-Disposition)을 직접 만듭니다.
                   filename="..."           : 영문 이름 (아주 옛 브라우저용)
                   filename*=UTF-8''%EC%97… : 한글 이름 (크롬/엣지/파이어폭스는 이걸 사용)
    """
    resp = Response(data, mimetype=XLSX_MIME)
    resp.headers['Content-Disposition'] = "attachment; filename=\"{}\"; filename*=UTF-8''{}".format(
        ascii_filename, quote(filename.encode('utf-8')))
    return resp


# ===========================================================================
# 로그인 처리
# ===========================================================================
@app.before_request
def load_user():
    """모든 요청 전에 로그인한 사용자를 g.user 에 넣어 둡니다."""
    g.user = None
    user_id = session.get('user_id')
    if user_id is not None:
        g.user = db.get_user(user_id)
        if g.user is None or not db.is_working(g.user):     # 계정이 삭제/전출·퇴직된 경우
            g.user = None
            session.clear()


def login_required(view):
    """로그인해야만 볼 수 있는 화면에 붙이는 장식(데코레이터)"""
    @wraps(view)
    def wrapped(*args, **kwargs):
        if g.user is None:
            return redirect(url_for('login'))
        return view(*args, **kwargs)
    return wrapped


@app.context_processor
def inject_common():
    """모든 HTML 템플릿에서 바로 쓸 수 있는 값들"""
    # here : 지금 보고 있는 화면 주소. 수정/삭제 후 이 화면으로 돌아오는 데 씁니다.
    # [Py2.7 차이] 옛 Flask 에는 request.full_path 가 없을 수 있어서 직접 만듭니다.
    qs = query_string()
    here = request.path + ('?' + qs if qs else '')
    return {'teams': TEAMS, 'grades': GRADES, 'leave_kinds': LEAVE_KINDS, 'here': here,
            'old_days': OLD_ANALYSIS_DAYS}


@app.template_global()
def url_with(**changes):
    """
    지금 주소에서 일부 값만 바꾼 주소. (검색 결과의 페이지 번호, 정렬 링크에 사용)
    예) 지금 /search?q=abc&page=2 에서 url_with(page=3) → /search?page=3&q=abc
    """
    args = request.args.to_dict(flat=False)       # target 처럼 값이 여러 개인 것도 유지
    for key, value in changes.items():
        args[key] = [value]
    # [Py2.7 차이] quote 는 바이트만 받아서 utf-8 로 바꿔서 넘깁니다.
    parts = ['{}={}'.format(quote(to_text(k).encode('utf-8'), b''),
                            quote(to_text(v).encode('utf-8'), b''))
             for k in sorted(args) for v in args[k]]
    return request.path + '?' + '&'.join(parts)


@app.template_filter('comma')
def comma(n):
    """1234 → '1,234'"""
    return '{:,}'.format(n)


@app.template_filter('highlight')
def highlight(text, q):
    """검색어를 <mark> 로 감싸 노란색으로 표시"""
    # [Py2.7 차이] str() 은 한글에서 에러 → text_type (= unicode) 사용
    safe = text_type(escape(text or ''))
    if not q:
        return Markup(safe)
    pattern = re.compile(re.escape(text_type(escape(q))), re.I | re.U)
    return Markup(pattern.sub(lambda m: '<mark>' + m.group(0) + '</mark>', safe))


@app.template_global()
def elapsed_days(start, d):
    """분석 시작일부터 d 까지 며칠째인지"""
    return (parse_date(d) - parse_date(start)).days + 1


# ===========================================================================
# 로그인 / 로그아웃 / 계정 만들기
# ===========================================================================
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        seq = form_text('seq')
        password = request.form.get('password', '')
        user = db.query('SELECT * FROM users WHERE seq=?', (seq,), one=True)
        if user is None or user['password'] != password:
            flash('순번 또는 비밀번호가 맞지 않습니다.')
        elif not db.is_working(user):
            flash('전출/퇴직 처리된 계정입니다. 계정관리에서 전출/퇴직일을 지우면 다시 쓸 수 있습니다.')
        else:
            session.clear()
            session['user_id'] = user['id']
            return redirect(url_for('index'))
    return render_template('login.html')


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


def save_account(user_id=None):
    """
    계정 추가/수정 공통 처리. 성공하면 None, 실패하면 오류 문구를 돌려줍니다.
    user_id 가 None 이면 새 계정, 있으면 수정.
    """
    seq = form_text('seq')
    team = form_text('team')
    grade = form_text('grade')
    name = form_text('name')
    password = request.form.get('password', '')
    # 전출/퇴직일 (계정 수정 화면에만 있음). 비어 있으면 재직 중
    left_text = form_text('left_date')
    left_date = parse_date(left_text).isoformat() if left_text else None

    if not seq.isdigit():
        return '순번은 숫자로 입력하세요.'
    if team not in TEAMS or not grade or not name:
        return '소속팀, 급수, 이름을 모두 입력하세요.'
    if user_id is None and not password:
        return '비밀번호를 입력하세요.'
    if (left_date and left_date <= date.today().isoformat()
            and g.user is not None and g.user['id'] == user_id):
        return '자기 계정은 전출/퇴직 처리할 수 없습니다. 다른 사람이 처리해 주세요.'

    try:
        if user_id is None:
            db.execute('INSERT INTO users (seq, team, grade, name, password) VALUES (?,?,?,?,?)',
                       (seq, team, grade, name, password))
        elif password:
            db.execute('UPDATE users SET seq=?, team=?, grade=?, name=?, password=?, left_date=? '
                       'WHERE id=?', (seq, team, grade, name, password, left_date, user_id))
        else:   # 비밀번호 칸을 비워 두면 기존 비밀번호 유지
            db.execute('UPDATE users SET seq=?, team=?, grade=?, name=?, left_date=? WHERE id=?',
                       (seq, team, grade, name, left_date, user_id))
    except sqlite3.IntegrityError:
        return '순번 {} 은(는) 이미 사용 중입니다.'.format(seq)
    return None


@app.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'POST':
        error = save_account()
        if error:
            flash(error)
        else:
            flash('계정을 만들었습니다. 로그인하세요.')
            return redirect(url_for('login'))
    return render_template('register.html')


# ===========================================================================
# 업무일지 메인 (하루 전체 요약)
# ===========================================================================
@app.route('/')
@login_required
def index():
    ctx = date_context()
    return render_template('index.html', active='index', data=collect_day(ctx['d']), **ctx)


# ===========================================================================
# 팀별 일일업무
# ===========================================================================
@app.route('/daily/<team>')
@login_required
def daily_team(team):
    t = team_from_url(team)
    ctx = date_context()
    return render_template('daily_team.html', active='daily' + team, team=team, t=t,
                           members=db.users_by_team(t),
                           works=db.daily_works(ctx['d'], t),
                           missing=db.missing_users(ctx['d'], t), **ctx)


@app.route('/daily/add', methods=['POST'])
@login_required
def daily_add():
    user = db.get_user(request.form.get('user_id'))
    title = form_text('title')
    if user is None or not title:
        flash('작성자와 업무 내용을 입력하세요.')
        return go_back()
    work_date = parse_date(request.form.get('work_date')).isoformat()
    db.execute('INSERT INTO daily_works (work_date, user_id, team, grade, name, title, detail) '
               'VALUES (?,?,?,?,?,?,?)',
               (work_date, user['id'], user['team'], user['grade'], user['name'],
                title, form_text('detail')))
    return go_back()


@app.route('/daily/<int:item_id>/edit', methods=['GET', 'POST'])
@login_required
def daily_edit(item_id):
    item = db.query('SELECT * FROM daily_works WHERE id=?', (item_id,), one=True)
    if item is None:
        abort(404)
    if request.method == 'POST':
        user = db.get_user(request.form.get('user_id'))
        title = form_text('title')
        if user is None or not title:
            flash('작성자와 업무 내용을 입력하세요.')
        else:
            db.execute('UPDATE daily_works SET work_date=?, user_id=?, team=?, grade=?, name=?, '
                       'title=?, detail=? WHERE id=?',
                       (parse_date(request.form.get('work_date')).isoformat(), user['id'],
                        user['team'], user['grade'], user['name'], title,
                        form_text('detail'), item_id))
            return go_back()
    return render_template('daily_edit.html', item=item, users=db.user_choices(item['user_id']))


@app.route('/daily/<int:item_id>/delete', methods=['POST'])
@login_required
def daily_delete(item_id):
    db.execute('DELETE FROM daily_works WHERE id=?', (item_id,))
    return go_back()


# ===========================================================================
# 3팀 신호처리 + 위성 목록
# ===========================================================================
@app.route('/signal')
@login_required
def signal():
    ctx = date_context()
    members = db.users_by_team('3팀') or db.all_users()
    return render_template('signal.html', active='signal', members=members,
                           satellites=db.satellites(), signals=db.signals_on(ctx['d']), **ctx)


@app.route('/signal/add', methods=['POST'])
@login_required
def signal_add():
    user = db.get_user(request.form.get('user_id'))
    satellite = form_text('satellite')
    case = form_text('case_notation')
    if user is None or not satellite or not case:
        flash('작성자, 위성명, CaseNotation 을 모두 입력하세요.')
        return go_back()
    db.execute('INSERT INTO signals (work_date, user_id, grade, name, satellite, case_notation) '
               'VALUES (?,?,?,?,?,?)',
               (parse_date(request.form.get('work_date')).isoformat(), user['id'],
                user['grade'], user['name'], satellite, case))
    return go_back()


@app.route('/signal/<int:item_id>/edit', methods=['GET', 'POST'])
@login_required
def signal_edit(item_id):
    item = db.query('SELECT * FROM signals WHERE id=?', (item_id,), one=True)
    if item is None:
        abort(404)
    if request.method == 'POST':
        user = db.get_user(request.form.get('user_id'))
        satellite = form_text('satellite')
        case = form_text('case_notation')
        if user is None or not satellite or not case:
            flash('작성자, 위성명, CaseNotation 을 모두 입력하세요.')
        else:
            db.execute('UPDATE signals SET work_date=?, user_id=?, grade=?, name=?, satellite=?, '
                       'case_notation=? WHERE id=?',
                       (parse_date(request.form.get('work_date')).isoformat(), user['id'],
                        user['grade'], user['name'], satellite, case, item_id))
            return go_back('signal')
    # [Py2.7 차이] 옛 Jinja2 에는 템플릿의 map(attribute=...) 이 없어서 이름 목록을 여기서 만듭니다.
    names = [s['name'] for s in db.satellites()]
    return render_template('signal_edit.html', item=item, users=db.user_choices(item['user_id']),
                           satellite_names=names)


@app.route('/signal/<int:item_id>/delete', methods=['POST'])
@login_required
def signal_delete(item_id):
    db.execute('DELETE FROM signals WHERE id=?', (item_id,))
    return go_back('signal')


@app.route('/satellite/add', methods=['POST'])
@login_required
def satellite_add():
    name = form_text('name')
    if name:
        try:
            db.execute('INSERT INTO satellites (name) VALUES (?)', (name,))
        except sqlite3.IntegrityError:
            flash('{} 은(는) 이미 목록에 있습니다.'.format(name))
    return go_back('signal')


@app.route('/satellite/<int:item_id>/delete', methods=['POST'])
@login_required
def satellite_delete(item_id):
    # 목록에서만 지워지고, 이미 등록된 신호처리 기록의 위성명은 그대로 남습니다.
    db.execute('DELETE FROM satellites WHERE id=?', (item_id,))
    return go_back('signal')


# ===========================================================================
# 5팀 분석현황
# ===========================================================================
def manager_options(keep_id=None):
    """담당자 선택 목록: 5팀을 먼저, 그 다음 다른 팀 (재직자 + 수정 중인 건의 원래 담당자)"""
    users = db.user_choices(keep_id)
    return [u for u in users if u['team'] == '5팀'] + [u for u in users if u['team'] != '5팀']


def manager_from_form():
    """폼에서 고른 담당자 계정 → (계정 번호, '급수 이름'). 없으면 (None, None)"""
    user = db.get_user(request.form.get('manager_id'))
    if user is None:
        return None, None
    return user['id'], '{} {}'.format(user['grade'], user['name'])


def analysis_who():
    """
    분석현황에서 누구 것을 볼지 (주소의 ?who=...)
      'me'  : 내가 담당자인 것   (5팀은 처음에 이것)
      'all' : 전체               (5팀이 아닌 사람은 처음에 이것)
      '12'  : 계정 번호 12 인 사람 것
    """
    who = request.args.get('who', '')
    if who in ('me', 'all') or who.isdigit():
        return who
    return 'me' if g.user['team'] == '5팀' else 'all'


@app.route('/analysis')
@login_required
def analysis():
    ctx = date_context()
    ongoing = db.analyses_ongoing(ctx['d'])
    finished = db.analyses_finished(ctx['d'])

    # 사람별 건수 (분석중 + 그날 분석종료) → 버튼과 사람 선택 목록에 표시
    counts = Counter(a['manager_id'] for a in ongoing + finished)
    people = [(u['id'], '{} {}'.format(u['grade'], u['name']), counts[u['id']])
              for u in manager_options()]

    who = analysis_who()
    if who != 'all':
        show_id = g.user['id'] if who == 'me' else int(who)
        ongoing = [a for a in ongoing if a['manager_id'] == show_id]
        finished = [a for a in finished if a['manager_id'] == show_id]
    old_count = len([a for a in ongoing
                     if elapsed_days(a['start_date'], ctx['d']) > OLD_ANALYSIS_DAYS])

    return render_template('analysis.html', active='analysis', managers=manager_options(),
                           ongoing=ongoing, finished=finished, who=who, people=people,
                           count_me=counts[g.user['id']], count_all=sum(counts.values()),
                           old_count=old_count, **ctx)


@app.route('/analysis/add', methods=['POST'])
@login_required
def analysis_add():
    manager_id, manager = manager_from_form()
    case = form_text('case_notation')
    if not manager or not case:
        flash('담당자와 CaseNotation 을 입력하세요.')
        return go_back('analysis')
    db.execute('INSERT INTO analyses (manager, manager_id, start_date, case_notation, note) '
               'VALUES (?,?,?,?,?)',
               (manager, manager_id, parse_date(request.form.get('start_date')).isoformat(), case,
                form_text('note')))
    return go_back('analysis')


@app.route('/analysis/<int:item_id>/finish', methods=['POST'])
@login_required
def analysis_finish(item_id):
    """[분석종료] 버튼: 보고 있던 날짜를 종료일로 저장"""
    item = db.query('SELECT * FROM analyses WHERE id=?', (item_id,), one=True)
    end = parse_date(request.form.get('end_date')).isoformat()
    if item is not None:
        if end < item['start_date']:
            flash('종료일이 시작일보다 빠를 수 없습니다.')
        else:
            db.execute('UPDATE analyses SET end_date=? WHERE id=?', (end, item_id))
    return go_back('analysis')


@app.route('/analysis/<int:item_id>/reopen', methods=['POST'])
@login_required
def analysis_reopen(item_id):
    """[분석중으로 되돌리기] 버튼: 종료일 지우기"""
    db.execute('UPDATE analyses SET end_date=NULL WHERE id=?', (item_id,))
    return go_back('analysis')


@app.route('/analysis/<int:item_id>/edit', methods=['GET', 'POST'])
@login_required
def analysis_edit(item_id):
    item = db.query('SELECT * FROM analyses WHERE id=?', (item_id,), one=True)
    if item is None:
        abort(404)
    if request.method == 'POST':
        if request.form.get('manager_id') == 'keep':        # 원래 담당자 그대로
            manager_id, manager = item['manager_id'], item['manager']
        else:
            manager_id, manager = manager_from_form()
        case = form_text('case_notation')
        start = parse_date(request.form.get('start_date')).isoformat()
        end_text = form_text('end_date')
        end = parse_date(end_text).isoformat() if end_text else None   # 비우면 분석중
        if not manager or not case:
            flash('담당자와 CaseNotation 을 입력하세요.')
        elif end and end < start:
            flash('종료일이 시작일보다 빠를 수 없습니다.')
        else:
            db.execute('UPDATE analyses SET manager=?, manager_id=?, start_date=?, end_date=?, '
                       'case_notation=?, note=? WHERE id=?',
                       (manager, manager_id, start, end, case, form_text('note'), item_id))
            return go_back('analysis')
    managers = manager_options(item['manager_id'])
    # 기존 담당자가 계정 목록에 있는지 (계정이 지워졌으면 화면에서 기존 값을 따로 보여줌)
    # [Py2.7 차이] 옛 Jinja2 에는 namespace() 가 없어서 이 검사를 템플릿 대신 여기서 합니다.
    manager_known = any(u['id'] == item['manager_id'] for u in managers)
    return render_template('analysis_edit.html', item=item, managers=managers,
                           manager_known=manager_known)


@app.route('/analysis/<int:item_id>/delete', methods=['POST'])
@login_required
def analysis_delete(item_id):
    db.execute('DELETE FROM analyses WHERE id=?', (item_id,))
    return go_back('analysis')


# ===========================================================================
# 휴가/부재
# ===========================================================================
def leave_form_values():
    """휴가 폼 값 검사. (값 dict, 오류문구) 를 돌려줍니다."""
    user = db.get_user(request.form.get('user_id'))
    kind = form_text('kind')
    start = parse_date(request.form.get('start_date')).isoformat()
    end = parse_date(request.form.get('end_date')).isoformat()
    if user is None or kind not in LEAVE_KINDS:
        return None, '대상자와 종류를 선택하세요.'
    if end < start:
        return None, '종료일이 시작일보다 빠를 수 없습니다.'
    return {'user': user, 'kind': kind, 'start': start, 'end': end, 'note': form_text('note')}, None


@app.route('/leave')
@login_required
def leave():
    today = date.today()
    first = today.replace(day=1)
    next_month = (first + timedelta(days=32)).replace(day=1)
    start = parse_date(request.args.get('start'), first).isoformat()
    end = parse_date(request.args.get('end'), next_month - timedelta(days=1)).isoformat()
    rows = db.query('SELECT * FROM leaves WHERE start_date<=? AND end_date>=? '
                    'ORDER BY start_date, team, name', (end, start))
    # 일수 계산해서 같이 넘기기
    items = [dict(r, days=(parse_date(r['end_date']) - parse_date(r['start_date'])).days + 1)
             for r in rows]
    return render_template('leave.html', active='leave', items=items, users=db.all_users(),
                           start=start, end=end, today=today.isoformat())


@app.route('/leave/add', methods=['POST'])
@login_required
def leave_add():
    v, error = leave_form_values()
    if error:
        flash(error)
    else:
        u = v['user']
        db.execute('INSERT INTO leaves (user_id, team, grade, name, kind, start_date, end_date, note) '
                   'VALUES (?,?,?,?,?,?,?,?)',
                   (u['id'], u['team'], u['grade'], u['name'], v['kind'], v['start'], v['end'],
                    v['note']))
    return go_back('leave')


@app.route('/leave/<int:item_id>/edit', methods=['GET', 'POST'])
@login_required
def leave_edit(item_id):
    item = db.query('SELECT * FROM leaves WHERE id=?', (item_id,), one=True)
    if item is None:
        abort(404)
    if request.method == 'POST':
        v, error = leave_form_values()
        if error:
            flash(error)
        else:
            u = v['user']
            db.execute('UPDATE leaves SET user_id=?, team=?, grade=?, name=?, kind=?, start_date=?, '
                       'end_date=?, note=? WHERE id=?',
                       (u['id'], u['team'], u['grade'], u['name'], v['kind'], v['start'], v['end'],
                        v['note'], item_id))
            return go_back('leave')
    return render_template('leave_edit.html', item=item, users=db.user_choices(item['user_id']))


@app.route('/leave/<int:item_id>/delete', methods=['POST'])
@login_required
def leave_delete(item_id):
    db.execute('DELETE FROM leaves WHERE id=?', (item_id,))
    return go_back('leave')


# ===========================================================================
# 검색
# ===========================================================================
SEARCH_TARGETS = [('daily3', '3팀 일일업무'), ('daily5', '5팀 일일업무'), ('signal', '3팀 신호처리'),
                  ('analysis', '5팀 분석현황'), ('leave', '휴가/부재')]


def search_params():
    """주소의 ?q=...&field=... 값들을 정리"""
    today = date.today()
    return {
        'q': request.args.get('q', '').strip(),
        'field': request.args.get('field', 'all'),
        'targets': request.args.getlist('target') or ['daily3', 'daily5', 'signal', 'analysis'],
        'start': parse_date(request.args.get('start'), today - timedelta(days=30)).isoformat(),
        'end': parse_date(request.args.get('end'), today).isoformat(),
    }


SEARCH_COLUMNS = [('date', '날짜'), ('kind', '구분'), ('team', '팀'), ('person', '작성자'),
                  ('content', '내용')]


def sort_state():
    """
    검색 결과 정렬 상태 → (열 이름, 내림차순인지)
    주소의 sort=열이름, dir=asc(오름)/desc(내림). 기본은 날짜 내림차순(최근 날짜가 위로).
    """
    key = request.args.get('sort')
    if key not in dict(SEARCH_COLUMNS):
        key = 'date'
    default_dir = 'desc' if key == 'date' else 'asc'
    return key, request.args.get('dir', default_dir) == 'desc'


def search_results(p):
    """
    검색 결과 전체를 정렬해서 돌려줍니다. 화면/엑셀/인쇄가 모두 같은 순서가 됩니다.
    (페이지를 나눠도 정렬은 '전체 결과' 기준이 되도록 서버에서 정렬합니다)
    """
    results = db.search_records(**p)        # 이미 최근 날짜 순
    key, desc = sort_state()
    if (key, desc) != ('date', True):
        # 파이썬 정렬은 값이 같으면 원래 순서를 지키므로, 같은 값끼리는 최근 날짜가 위로
        results.sort(key=lambda r: r[key], reverse=desc)
    return results


def paginate(items):
    """목록을 페이지로 나눕니다. 주소의 page=번호, size=한 페이지 건수"""
    size = request.args.get('size', '')
    size = int(size) if size.isdigit() and int(size) in PAGE_SIZES else PAGE_SIZES[0]
    pages = max(1, (len(items) + size - 1) // size)
    page = request.args.get('page', '')
    page = min(max(int(page), 1), pages) if page.isdigit() else 1
    first = (page - 1) * size
    return {'rows': items[first:first + size], 'page': page, 'pages': pages, 'size': size,
            'first': first + 1 if items else 0, 'last': min(first + size, len(items)),
            'numbers': page_numbers(page, pages)}


def page_numbers(page, pages):
    """
    페이지 번호 줄에 보여줄 번호들. 처음/끝 + 지금 페이지 앞뒤 2개, 건너뛴 곳은 None(…)
    예) 87페이지 중 6페이지 → [1, None, 4, 5, 6, 7, 8, None, 87]
    """
    show = sorted(n for n in set([1, pages] + list(range(page - 2, page + 3))) if 1 <= n <= pages)
    numbers = []
    for n in show:
        if numbers and n - numbers[-1] > 1:
            numbers.append(None)
        numbers.append(n)
    return numbers


def kind_counts(results):
    """검색 결과를 대상별로 센 것. 예) [('3팀 일일업무', 120), ('휴가/부재', 3)]"""
    counts = Counter(r['kind'] if r['kind'] == '휴가/부재' else '{} {}'.format(r['team'], r['kind'])
                     for r in results)
    return [(label, counts[label]) for _, label in SEARCH_TARGETS if counts[label]]


@app.route('/search')
@login_required
def search():
    p = search_params()
    searched = 'field' in request.args        # [검색] 버튼을 눌렀는지
    results = search_results(p) if searched else []
    sort_key, sort_desc = sort_state()
    return render_template('search.html', active='search', p=p, total=len(results),
                           pg=paginate(results), counts=kind_counts(results),
                           sort_key=sort_key, sort_desc=sort_desc, columns=SEARCH_COLUMNS,
                           page_sizes=PAGE_SIZES, print_warn=PRINT_WARN,
                           searched=searched, search_targets=SEARCH_TARGETS,
                           query_string=query_string())


# ===========================================================================
# 계정 관리
# ===========================================================================
@app.route('/accounts')
@login_required
def accounts():
    team = request.args.get('team', '')
    users = db.users_by_team(team, include_left=True) if team in TEAMS else db.query(
        'SELECT * FROM users ORDER BY seq')
    # 재직자를 위에, 전출/퇴직자를 아래에
    users = [u for u in users if db.is_working(u)] + [u for u in users if not db.is_working(u)]
    return render_template('accounts.html', active='accounts', users=users, team_filter=team,
                           today=date.today().isoformat())


@app.route('/accounts/add', methods=['POST'])
@login_required
def account_add():
    error = save_account()
    if error:
        flash(error)
    return go_back('accounts')


@app.route('/accounts/<int:user_id>/edit', methods=['GET', 'POST'])
@login_required
def account_edit(user_id):
    user = db.get_user(user_id)
    if user is None:
        abort(404)
    if request.method == 'POST':
        error = save_account(user_id)
        if error:
            flash(error)
        else:
            return go_back('accounts')
    return render_template('account_edit.html', user=user)


@app.route('/accounts/<int:user_id>/delete', methods=['POST'])
@login_required
def account_delete(user_id):
    # 계정을 지워도 그 사람이 쓴 업무 기록은 남아 있습니다.
    db.execute('DELETE FROM users WHERE id=?', (user_id,))
    if user_id == g.user['id']:
        session.clear()
        return redirect(url_for('login'))
    return go_back('accounts')


# ===========================================================================
# 설정 (백업)
# ===========================================================================
@app.route('/settings', methods=['GET', 'POST'])
@login_required
def settings():
    if request.method == 'POST':
        folder = form_text('backup_dir')
        days = form_text('retention_days')
        if not folder:
            flash('백업 폴더 경로를 입력하세요.')
        elif not days.isdigit() or int(days) < 1:
            flash('보관 기간은 1 이상의 숫자로 입력하세요.')
        else:
            try:
                ensure_dir(folder)          # 폴더가 없으면 만들어 봄
            except OSError as e:
                flash('폴더를 만들 수 없습니다: ' + error_text(e))
            else:
                db.set_setting('backup_dir', folder)
                db.set_setting('retention_days', int(days))
                flash('저장했습니다.')
        return redirect(url_for('settings'))
    return render_template('settings.html', active='settings',
                           backup_dir=backup.get_backup_dir(),
                           retention_days=backup.get_retention_days(),
                           files=backup.list_backups())


@app.route('/settings/backup-now', methods=['POST'])
@login_required
def backup_now():
    try:
        path = backup.run_backup(force=True)
        flash('백업했습니다: ' + path)
    except Exception as e:
        flash('백업 실패: ' + error_text(e))
    return redirect(url_for('settings'))


# ===========================================================================
# 엑셀 다운로드 / 인쇄
# ===========================================================================
@app.route('/export/day')
@login_required
def export_day():
    d = parse_date(request.args.get('date')).isoformat()
    data = excel.day_workbook(d, collect_day(d), TEAMS)
    return xlsx_response(data, '업무일지_{}.xlsx'.format(d), 'worklog_{}.xlsx'.format(d))


@app.route('/export/search')
@login_required
def export_search():
    data = excel.search_workbook(search_results(search_params()))
    today = date.today().isoformat()
    return xlsx_response(data, '검색결과_{}.xlsx'.format(today), 'search_{}.xlsx'.format(today))


@app.route('/print/day')
@login_required
def print_day():
    ctx = date_context()
    return render_template('print_day.html', data=collect_day(ctx['d']), **ctx)


@app.route('/print/search')
@login_required
def print_search():
    p = search_params()
    return render_template('print_search.html', p=p, results=search_results(p))


# ===========================================================================
# 프로그램 시작
# ===========================================================================
if __name__ == '__main__':
    db.init_db()                 # 테이블 준비
    backup.start_scheduler()     # 자동 백업 시작
    # [Py2.7 차이] 영문 윈도우 검은 창은 한글 출력 시 에러가 날 수 있어서 안내 문구는 영어로
    log('Daily Work server started: http://localhost:{}  (stop: Ctrl+C)'.format(PORT))
    # host='0.0.0.0' : 같은 망의 다른 PC에서도 접속 가능하게
    # threaded=True  : 여러 사람이 동시에 접속해도 처리
    app.run(host='0.0.0.0', port=PORT, debug=False, threaded=True)
