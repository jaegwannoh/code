# -*- coding: utf-8 -*-
"""
db.py - 데이터베이스(SQLite) 담당 파일

SQLite는 파이썬에 기본으로 들어있는 데이터베이스입니다.
별도 설치 없이 'worklog.db' 라는 파일 하나에 모든 데이터가 저장됩니다.

이 파일에는
  1) DB 연결/생성 코드
  2) 화면에서 자주 쓰는 데이터 조회 함수
가 들어 있습니다.
"""
from __future__ import unicode_literals

import os
import sqlite3
from datetime import date

from flask import g

from compat import path_text, to_text

# 이 파일(db.py)이 있는 폴더 = 프로그램 폴더
# [Py2.7 차이] 2.7 에서는 __file__ 이 바이트(str)라서 한글 폴더면 깨집니다 → path_text 로 unicode 로
BASE_DIR = path_text(os.path.dirname(os.path.abspath(__file__)))
# DB 파일 위치
DB_PATH = os.path.join(BASE_DIR, 'worklog.db')


# ---------------------------------------------------------------------------
# 테이블 구조
#   - 작성자 급수/이름/팀은 글을 쓸 때의 값을 함께 저장합니다.
#     (나중에 승진하거나 계정이 지워져도 옛날 기록은 그대로 남게 하기 위해)
# ---------------------------------------------------------------------------
SCHEMA = """
CREATE TABLE IF NOT EXISTS users (
    id       INTEGER PRIMARY KEY AUTOINCREMENT,
    seq      TEXT NOT NULL UNIQUE,      -- 순번 (로그인 ID). 앞자리 0 유지를 위해 글자로 저장
    team     TEXT NOT NULL,             -- '3팀' / '5팀'
    grade    TEXT NOT NULL,             -- 급수 (예: '7급')
    name     TEXT NOT NULL,             -- 이름
    password TEXT NOT NULL,             -- 비밀번호 (내부망이라 그대로 저장)
    left_date TEXT                      -- 전출/퇴직일. 비어 있으면(NULL) 재직 중
);

-- 일일업무
CREATE TABLE IF NOT EXISTS daily_works (
    id        INTEGER PRIMARY KEY AUTOINCREMENT,
    work_date TEXT NOT NULL,            -- 'YYYY-MM-DD'
    user_id   INTEGER NOT NULL,
    team      TEXT NOT NULL,
    grade     TEXT NOT NULL,
    name      TEXT NOT NULL,
    title     TEXT NOT NULL,            -- 업무 내용
    detail    TEXT NOT NULL DEFAULT ''  -- 세부 내용
);

-- 3팀 신호처리
CREATE TABLE IF NOT EXISTS signals (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    work_date     TEXT NOT NULL,
    user_id       INTEGER NOT NULL,
    grade         TEXT NOT NULL,
    name          TEXT NOT NULL,
    satellite     TEXT NOT NULL,
    case_notation TEXT NOT NULL
);

-- 위성 목록
CREATE TABLE IF NOT EXISTS satellites (
    id   INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE
);

-- 5팀 분석현황
--   end_date 가 비어 있으면(NULL) '분석중', 값이 있으면 그날 '분석종료'
CREATE TABLE IF NOT EXISTS analyses (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    manager       TEXT NOT NULL,        -- 담당자 (예: '6급 최지훈'). 등록 당시 급수/이름
    manager_id    INTEGER,              -- 담당자 계정 번호 (users.id). '내 분석' 을 찾을 때 사용
    start_date    TEXT NOT NULL,
    end_date      TEXT,
    case_notation TEXT NOT NULL,
    note          TEXT NOT NULL DEFAULT ''
);

-- 휴가/부재
CREATE TABLE IF NOT EXISTS leaves (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id    INTEGER NOT NULL,
    team       TEXT NOT NULL,
    grade      TEXT NOT NULL,
    name       TEXT NOT NULL,
    kind       TEXT NOT NULL,           -- 연가/출장/교육/기타
    start_date TEXT NOT NULL,
    end_date   TEXT NOT NULL,
    note       TEXT NOT NULL DEFAULT ''
);

-- 설정값 (백업 경로 등)
CREATE TABLE IF NOT EXISTS settings (
    key   TEXT PRIMARY KEY,
    value TEXT NOT NULL
);
"""


# ---------------------------------------------------------------------------
# 인덱스 (책의 '찾아보기' 같은 것)
#   데이터가 몇 년치 쌓여도 날짜/팀으로 찾을 때 표 전체를 읽지 않고 바로 찾아갑니다.
#   이미 있으면 그냥 넘어갑니다.
# ---------------------------------------------------------------------------
INDEXES = """
CREATE INDEX IF NOT EXISTS idx_daily_date     ON daily_works (work_date, team);
CREATE INDEX IF NOT EXISTS idx_daily_user     ON daily_works (user_id, work_date);
CREATE INDEX IF NOT EXISTS idx_signal_date    ON signals (work_date);
CREATE INDEX IF NOT EXISTS idx_analysis_start ON analyses (start_date);
CREATE INDEX IF NOT EXISTS idx_analysis_end   ON analyses (end_date);
CREATE INDEX IF NOT EXISTS idx_analysis_mgr   ON analyses (manager_id);
CREATE INDEX IF NOT EXISTS idx_leave_date     ON leaves (start_date, end_date);
"""


def connect():
    """DB 파일에 연결합니다. (여러 사람이 동시에 저장해도 10초까지 기다려 줌)"""
    return sqlite3.connect(DB_PATH, timeout=10)


def init_db():
    """프로그램 시작 시 한 번 실행: 테이블이 없으면 만들고, 예전 DB 는 새 구조로 고칩니다."""
    conn = connect()
    try:
        conn.executescript(SCHEMA)
        upgrade(conn)
        conn.executescript(INDEXES)
        conn.commit()
    finally:
        conn.close()


def columns(conn, table):
    """표(table)에 있는 칸 이름 목록"""
    return [to_text(r[1]) for r in conn.execute('PRAGMA table_info({})'.format(table))]


def upgrade(conn):
    """
    예전 버전에서 쓰던 worklog.db 에 새로 생긴 칸을 추가합니다. (데이터는 그대로 유지)
      - users.left_date     : 전출/퇴직일
      - analyses.manager_id : 담당자 계정 번호
    """
    if 'left_date' not in columns(conn, 'users'):
        conn.execute('ALTER TABLE users ADD COLUMN left_date TEXT')
    if 'manager_id' not in columns(conn, 'analyses'):
        conn.execute('ALTER TABLE analyses ADD COLUMN manager_id INTEGER')

    # 담당자 계정 번호가 비어 있는 분석 건은 '급수 이름' 글자가 같은 계정으로 채웁니다.
    # (같은 '급수 이름' 계정이 둘 이상이면 누구인지 알 수 없어서 비워 둡니다)
    labels = {}
    for user_id, grade, name in conn.execute('SELECT id, grade, name FROM users'):
        labels.setdefault('{} {}'.format(to_text(grade), to_text(name)), []).append(user_id)
    rows = conn.execute('SELECT id, manager FROM analyses WHERE manager_id IS NULL').fetchall()
    for item_id, manager in rows:
        ids = labels.get(to_text(manager), [])
        if len(ids) == 1:
            conn.execute('UPDATE analyses SET manager_id=? WHERE id=?', (ids[0], item_id))


# ---------------------------------------------------------------------------
# 웹 요청마다 DB 연결을 하나씩 열고, 요청이 끝나면 닫습니다.
# (g 는 Flask가 '요청 하나' 동안 값을 보관해 주는 상자입니다)
# [Py2.7 차이] 옛 Flask 의 g 에는 'in' 검사와 g.pop() 이 없을 수 있어서 getattr 을 씁니다.
# ---------------------------------------------------------------------------
def dict_row(cursor, row):
    """
    DB 결과 한 줄을 dict 로 바꿉니다. → row['name'] 처럼 이름으로 꺼낼 수 있음
    [Py2.7 차이] 파이썬 3 버전은 sqlite3.Row 를 썼지만, 2.7 의 sqlite3.Row 는
                 row[u'name'] (unicode 이름) 으로 꺼내면 IndexError 가 납니다.
                 이 프로그램은 모든 글자가 unicode 라서 dict 로 바꿔서 씁니다.
    """
    return dict((to_text(col[0]), row[i]) for i, col in enumerate(cursor.description))


def get_db():
    conn = getattr(g, '_db', None)
    if conn is None:
        conn = g._db = connect()
        conn.row_factory = dict_row
    return conn


def close_db(error=None):
    conn = getattr(g, '_db', None)
    if conn is not None:
        conn.close()
        g._db = None


def query(sql, args=(), one=False):
    """SELECT 실행. one=True 이면 첫 줄 하나만(없으면 None) 돌려줍니다."""
    rows = get_db().execute(sql, args).fetchall()
    if one:
        return rows[0] if rows else None
    return rows


def execute(sql, args=()):
    """INSERT/UPDATE/DELETE 실행 후 저장(commit)."""
    conn = get_db()
    cur = conn.execute(sql, args)
    conn.commit()
    return cur.lastrowid


# ---------------------------------------------------------------------------
# 설정값 읽기/쓰기
#   백업 스레드(웹 요청 밖)에서도 쓰기 때문에 연결을 따로 엽니다.
# ---------------------------------------------------------------------------
def get_setting(key, default=None):
    conn = connect()
    try:
        row = conn.execute('SELECT value FROM settings WHERE key=?', (key,)).fetchone()
    finally:
        conn.close()
    return row[0] if row else default


def set_setting(key, value):
    conn = connect()
    try:
        # [Py2.7 차이] str(value) 대신 to_text: 한글 경로도 에러 없이 저장
        conn.execute('INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)',
                     (key, to_text(value)))
        conn.commit()
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# 사용자
#   전출/퇴직일(left_date)이 지난 사람은 선택 목록과 미제출 인원에서 빠집니다.
#   계정과 그 사람이 쓴 기록은 그대로 남습니다.
# ---------------------------------------------------------------------------
# d 날짜에 재직 중인 사람 조건. 전출/퇴직일 당일부터 빠집니다.
WORKING_ON = '(left_date IS NULL OR left_date > ?)'


def get_user(user_id):
    return query('SELECT * FROM users WHERE id=?', (user_id,), one=True)


def is_working(user):
    """오늘 재직 중인지"""
    return not user['left_date'] or user['left_date'] > date.today().isoformat()


def all_users(include_left=False):
    if include_left:
        return query('SELECT * FROM users ORDER BY team, grade, name')
    return query('SELECT * FROM users WHERE ' + WORKING_ON + ' ORDER BY team, grade, name',
                 (date.today().isoformat(),))


def users_by_team(team, include_left=False):
    if include_left:
        return query('SELECT * FROM users WHERE team=? ORDER BY grade, name', (team,))
    return query('SELECT * FROM users WHERE team=? AND ' + WORKING_ON + ' ORDER BY grade, name',
                 (team, date.today().isoformat()))


def user_choices(keep_id=None):
    """
    작성자/대상자 선택 목록 (재직자만).
    수정 화면에서는 원래 작성자(keep_id)가 전출/퇴직했어도 목록에 남겨서
    저장할 때 작성자가 바뀌지 않게 합니다.
    """
    users = all_users()
    if keep_id is not None and not any(u['id'] == keep_id for u in users):
        kept = get_user(keep_id)
        if kept is not None:
            users.append(kept)
    return users


# ---------------------------------------------------------------------------
# 날짜별 조회 (d 는 'YYYY-MM-DD' 문자열)
# ---------------------------------------------------------------------------
def daily_works(d, team):
    return query('SELECT * FROM daily_works WHERE work_date=? AND team=? '
                 'ORDER BY grade, name, id', (d, team))


def leaves_on(d):
    """d 날짜에 휴가/부재 중인 사람"""
    return query('SELECT * FROM leaves WHERE start_date<=? AND end_date>=? '
                 'ORDER BY team, grade, name', (d, d))


def missing_users(d, team):
    """d 날짜에 일일업무를 하나도 안 쓴 사람 (휴가/부재자, 전출/퇴직자는 제외)"""
    return query(
        """
        SELECT * FROM users
        WHERE team = ?
          AND """ + WORKING_ON + """
          AND id NOT IN (SELECT user_id FROM daily_works
                         WHERE work_date = ? AND user_id IS NOT NULL)
          AND id NOT IN (SELECT user_id FROM leaves
                         WHERE start_date <= ? AND end_date >= ? AND user_id IS NOT NULL)
        ORDER BY grade, name
        """, (team, d, d, d, d))


def signals_on(d):
    return query('SELECT * FROM signals WHERE work_date=? ORDER BY id', (d,))


def satellites():
    return query('SELECT * FROM satellites ORDER BY name')


def analyses_ongoing(d):
    """d 날짜 기준 '분석중' : 시작했고, 아직 안 끝났거나 d 보다 뒤에 끝난 것"""
    return query('SELECT * FROM analyses WHERE start_date<=? '
                 'AND (end_date IS NULL OR end_date>?) ORDER BY start_date, id', (d, d))


def analyses_finished(d):
    """d 날짜에 '분석종료' 된 것"""
    return query('SELECT * FROM analyses WHERE end_date=? ORDER BY id', (d,))


# ---------------------------------------------------------------------------
# 검색
#   q      : 검색어 (비어 있으면 기간 안의 전체)
#   field  : 'all' / 'name' / 'content'
#   targets: ['daily3', 'daily5', 'signal', 'analysis', 'leave'] 중 고른 것
#   start, end : 기간 'YYYY-MM-DD'
# 결과는 화면/엑셀/인쇄에서 똑같이 쓰도록 dict 목록으로 돌려줍니다.
#
# [Py2.7 차이] 파이썬 3의 f'...' (f-string) 문법이 2.7 에는 없어서
#              '{}'.format(...) 으로 썼습니다.
# ---------------------------------------------------------------------------
def search_records(q, field, targets, start, end):
    like = '%{}%'.format(q)

    def where(name_cols, content_cols):
        """검색어 조건 SQL 조각과 값 목록을 만듭니다."""
        if not q:
            return '1=1', []
        if field == 'name':
            cols = name_cols
        elif field == 'content':
            cols = content_cols
        else:
            cols = name_cols + content_cols
        sql = '(' + ' OR '.join('{} LIKE ?'.format(c) for c in cols) + ')'
        return sql, [like] * len(cols)

    def person(r):
        return '{} {}'.format(r['grade'], r['name'])

    results = []

    # 일일업무는 3팀 / 5팀을 따로 고를 수 있습니다.
    for target, team in (('daily3', '3팀'), ('daily5', '5팀')):
        if target not in targets:
            continue
        w, a = where(['name'], ['title', 'detail'])
        rows = query('SELECT * FROM daily_works WHERE team=? AND work_date BETWEEN ? AND ? '
                     'AND ' + w, [team, start, end] + a)
        for r in rows:
            content = r['title'] + (' / ' + r['detail'] if r['detail'] else '')
            results.append({'date': r['work_date'], 'kind': '일일업무', 'team': r['team'],
                            'person': person(r), 'content': content})

    if 'signal' in targets:
        w, a = where(['name'], ['satellite', 'case_notation'])
        rows = query('SELECT * FROM signals WHERE work_date BETWEEN ? AND ? AND ' + w,
                     [start, end] + a)
        for r in rows:
            results.append({'date': r['work_date'], 'kind': '신호처리', 'team': '3팀',
                            'person': person(r),
                            'content': '{} {}'.format(r['satellite'], r['case_notation'])})

    if 'analysis' in targets:
        w, a = where(['manager'], ['case_notation', 'note'])
        # 기간이 겹치는 분석 건 (시작일 <= 검색 끝, 종료일 >= 검색 시작 또는 분석중)
        rows = query('SELECT * FROM analyses WHERE start_date<=? '
                     'AND (end_date IS NULL OR end_date>=?) AND ' + w, [end, start] + a)
        for r in rows:
            status = '분석종료' if r['end_date'] else '분석중'
            period = '~' + r['end_date'] if r['end_date'] else '~'
            content = '[{}] {} ({}{})'.format(status, r['case_notation'], r['start_date'], period)
            if r['note']:
                content += ' / ' + r['note']
            results.append({'date': r['start_date'], 'kind': '분석현황', 'team': '5팀',
                            'person': r['manager'], 'content': content})

    if 'leave' in targets:
        w, a = where(['name'], ['kind', 'note'])
        rows = query('SELECT * FROM leaves WHERE start_date<=? AND end_date>=? AND ' + w,
                     [end, start] + a)
        for r in rows:
            content = '{} {}~{}'.format(r['kind'], r['start_date'], r['end_date'])
            if r['note']:
                content += ' / ' + r['note']
            results.append({'date': r['start_date'], 'kind': '휴가/부재', 'team': r['team'],
                            'person': person(r), 'content': content})

    # 최근 날짜가 위로
    results.sort(key=lambda x: x['date'], reverse=True)
    return results
