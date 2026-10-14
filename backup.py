# -*- coding: utf-8 -*-
"""
backup.py - DB 자동 백업

- 하루에 한 번 'worklog_YYYY-MM-DD.db' 파일로 백업합니다.
- 보관 기간(기본 30일)이 지난 백업 파일은 자동으로 지웁니다.
- 백업 폴더와 보관 기간은 웹의 [설정] 화면에서 바꿀 수 있습니다.
"""
from __future__ import unicode_literals

import os
import re
import shutil
import sqlite3
import threading
import time
from datetime import date, datetime, timedelta

import db
from compat import ensure_dir, error_text, log, path_text

DEFAULT_DIR = os.path.join(db.BASE_DIR, 'backup')     # 기본 백업 폴더
DEFAULT_DAYS = 30                                      # 기본 보관 기간
FILE_RE = re.compile(r'^worklog_(\d{4}-\d{2}-\d{2})\.db$')
CHECK_INTERVAL = 60 * 60                               # 1시간마다 확인


def get_backup_dir():
    return path_text(db.get_setting('backup_dir', DEFAULT_DIR))


def get_retention_days():
    try:
        return int(db.get_setting('retention_days', DEFAULT_DAYS))
    except ValueError:
        return DEFAULT_DAYS


def copy_db(target):
    """
    DB 파일을 target 으로 안전하게 복사합니다.
    [Py2.7 차이] 파이썬 3.7 이상에 있는 sqlite3 의 backup() 기능이 2.7 에는 없습니다.
                 그래서 '잠깐 저장을 막아 두고(BEGIN IMMEDIATE) → 파일 복사' 하는 방식으로 합니다.
                 복사는 순식간에 끝나고, 그동안 저장하려는 사람은 잠시 기다렸다가 저장됩니다.
    """
    conn = sqlite3.connect(db.DB_PATH, timeout=30)
    conn.isolation_level = None             # BEGIN/ROLLBACK 을 직접 실행하기 위해
    try:
        conn.execute('BEGIN IMMEDIATE')     # 다른 사람이 저장하지 못하게 잠금
        try:
            tmp = target + '.tmp'
            shutil.copyfile(db.DB_PATH, tmp)    # 임시 이름으로 먼저 복사하고
            if os.path.exists(target):
                os.remove(target)
            os.rename(tmp, target)              # 다 복사되면 진짜 이름으로 바꿈
        finally:
            conn.execute('ROLLBACK')    # 잠금 풀기 (바꾼 내용이 없으니 ROLLBACK 해도 됨)
    finally:
        conn.close()


def run_backup(force=False):
    """
    오늘 날짜 백업 파일을 만듭니다.
    force=False 이면 오늘 백업이 이미 있을 때 건너뜁니다.
    ([지금 백업하기] 버튼은 force=True 로 오늘 파일을 덮어씁니다.)
    """
    folder = get_backup_dir()
    ensure_dir(folder)
    target = os.path.join(folder, 'worklog_{}.db'.format(date.today().isoformat()))

    if force or not os.path.exists(target):
        copy_db(target)

    delete_old_backups(folder)
    return target


def delete_old_backups(folder):
    """보관 기간이 지난 백업 파일 삭제 (파일 이름의 날짜 기준)"""
    limit = date.today() - timedelta(days=get_retention_days())
    for filename in os.listdir(folder):
        m = FILE_RE.match(filename)
        if not m:
            continue   # 우리가 만든 백업 파일이 아니면 건드리지 않음
        file_date = datetime.strptime(m.group(1), '%Y-%m-%d').date()
        if file_date <= limit:
            os.remove(os.path.join(folder, filename))


def list_backups():
    """설정 화면에 보여줄 백업 파일 목록 (최근 순)"""
    folder = get_backup_dir()
    if not os.path.isdir(folder):
        return []
    files = []
    for filename in os.listdir(folder):
        if FILE_RE.match(filename):
            path = os.path.join(folder, filename)
            files.append({
                'name': filename,
                'time': datetime.fromtimestamp(os.path.getmtime(path)).strftime('%Y-%m-%d %H:%M'),
                # [Py2.7 차이] 2.7 에서 1500 / 1024 는 1 (소수점 버림). 1024.0 으로 나눠야 1.46
                'size_kb': round(os.path.getsize(path) / 1024.0, 1),
            })
    files.sort(key=lambda f: f['name'], reverse=True)
    return files


def _loop():
    """백그라운드에서 계속 돌면서 1시간마다 오늘 백업이 있는지 확인"""
    while True:
        try:
            run_backup()
        except Exception as e:      # 백업 실패해도 프로그램은 계속 돌아가게
            log('[backup error] ' + error_text(e))
        time.sleep(CHECK_INTERVAL)


def start_scheduler():
    # daemon : 프로그램이 꺼지면 이 스레드도 같이 꺼집니다.
    # [Py2.7 차이] Thread(..., daemon=True) 처럼 한 줄로 못 쓰고 따로 지정해야 합니다.
    t = threading.Thread(target=_loop)
    t.daemon = True
    t.start()
