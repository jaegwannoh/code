# -*- coding: utf-8 -*-
"""
compat.py - 파이썬 2.7 에서 한글/파일 처리를 안전하게 하기 위한 도우미 모음

[Py2.7 차이] 왜 이 파일이 필요한가?
  파이썬 2.7 에는 문자열이 두 종류 있습니다.
    - str     : 바이트(byte) 덩어리.  'abc'
    - unicode : 진짜 글자.            u'한글'
  한글은 반드시 unicode 로 다뤄야 깨지지 않습니다.
  그래서 모든 .py 파일 맨 위에
      from __future__ import unicode_literals
  를 넣어서 '...' 라고만 써도 u'...' (unicode) 가 되게 했습니다.

  그래도 바이트(str)가 끼어드는 곳(파일 경로, 에러 메시지 등)이 있어서
  그런 곳을 이 파일의 함수로 처리합니다.

  (파이썬 3에서도 그대로 돌아가게 짜 두었습니다.)
"""
from __future__ import unicode_literals

import errno
import io
import os
import sys
import threading

# [Py2.7 차이] 2.7 에는 datetime.strptime() 을 여러 스레드(백업, 동시 접속)에서 처음 쓸 때
#              "Failed to import _strptime ..." 에러가 나는 버그가 있습니다.
#              프로그램 시작할 때 미리 불러 두면 생기지 않습니다.
import _strptime  # noqa: F401

PY2 = sys.version_info[0] == 2

if PY2:
    text_type = unicode                 # noqa: F821  (파이썬 2 에만 있는 이름)
    from urllib import quote            # 파이썬 2 위치
else:
    text_type = str
    from urllib.parse import quote      # 파이썬 3 위치

# 윈도우에서 파일 이름을 바이트로 바꿀 때 쓰는 인코딩 (보통 'mbcs')
FS_ENCODING = sys.getfilesystemencoding() or 'utf-8'


def to_text(value):
    """
    무엇이든 unicode 글자로 바꿉니다.
    [Py2.7 차이] 파이썬 2 에서 str(u'한글') 을 하면 UnicodeEncodeError 가 납니다.
                 그래서 str() 대신 이 함수를 씁니다.
    """
    if value is None:
        return ''
    if isinstance(value, text_type):
        return value
    if isinstance(value, bytes):
        for enc in ('utf-8', FS_ENCODING, 'cp949'):
            try:
                return value.decode(enc)
            except (UnicodeDecodeError, LookupError):
                pass
        return value.decode('utf-8', 'replace')
    try:
        return text_type(value)
    except UnicodeError:
        return to_text(repr(value))


def error_text(e):
    """
    에러(예외)를 화면에 보여줄 글자로 바꿉니다.
    [Py2.7 차이] 윈도우 에러 메시지는 바이트(str)로 오는 경우가 있어서
                 그냥 '{}'.format(e) 를 하면 에러 처리 중에 또 에러가 날 수 있습니다.
    """
    try:
        if PY2:
            # 에러에 담긴 값(args)을 하나씩 글자로 바꿔서 이어 붙입니다.
            parts = [to_text(a) for a in getattr(e, 'args', ())]
            if getattr(e, 'filename', None):
                parts.append(to_text(e.filename))
            return ' '.join(parts) or to_text(e.__class__.__name__)
        return text_type(e)
    except Exception:
        return to_text(repr(e))


def path_text(path):
    """파일 경로를 unicode 로 (한글 폴더 이름도 깨지지 않게)"""
    return to_text(path)


def ensure_dir(folder):
    """
    폴더가 없으면 만듭니다. 이미 있으면 아무것도 안 합니다.
    [Py2.7 차이] 파이썬 3의 os.makedirs(folder, exist_ok=True) 가 2.7 에는 없습니다.
    """
    try:
        os.makedirs(folder)
    except OSError as e:
        if e.errno != errno.EEXIST or not os.path.isdir(folder):
            raise


def log(message):
    """
    검은 창(콘솔)에 메시지 출력.
    [Py2.7 차이] 영문 윈도우 콘솔은 한글을 못 찍어서 print 하다가 프로그램이 죽을 수 있습니다.
                 출력할 수 없는 글자는 '?' 로 바꿔서 절대 죽지 않게 합니다.
    """
    text = to_text(message) + '\n'
    stream = sys.stdout
    try:
        if PY2:
            enc = getattr(stream, 'encoding', None) or 'ascii'
            try:
                data = text.encode(enc, 'replace')
            except LookupError:            # 'cp65001' 처럼 2.7이 모르는 인코딩
                data = text.encode('ascii', 'replace')
            stream.write(data)
        else:
            stream.write(text)
        stream.flush()
    except Exception:
        pass


class SizeLimitedLog(object):
    """
    server.log 파일에 메시지를 쓰는 '가짜 화면(stdout)'.
    파일이 max_bytes 를 넘으면 server_old.log 로 이름을 바꾸고 새로 씁니다.
    (서버를 몇 달씩 켜 둬도 로그 파일이 끝없이 커지지 않게)
    """
    encoding = 'utf-8'      # log() 와 logging 이 이 값을 보고 글자를 바이트로 바꿉니다.

    def __init__(self, path, max_bytes):
        self.path = path
        self.old_path = os.path.splitext(path)[0] + '_old.log'
        self.max_bytes = max_bytes
        self.lock = threading.Lock()
        self.f = None
        self._rotate()      # 켤 때마다 직전 기록은 server_old.log 로

    def _rotate(self):
        if self.f is not None:
            self.f.close()
        try:
            if os.path.exists(self.path):
                if os.path.exists(self.old_path):
                    os.remove(self.old_path)
                os.rename(self.path, self.old_path)
        except OSError:
            pass            # 이름을 못 바꾸면 기존 파일 뒤에 이어서 씀
        self.f = io.open(self.path, 'ab')

    def write(self, data):
        if isinstance(data, text_type):
            data = data.encode('utf-8', 'replace')
        with self.lock:
            try:
                self.f.write(data)
                self.f.flush()
                if self.f.tell() > self.max_bytes:
                    self._rotate()
            except Exception:
                pass        # 기록하다 실패해도 서버는 죽지 않게

    def flush(self):
        pass

    def isatty(self):
        return False


def log_to_file(path, max_bytes=5 * 1024 * 1024):
    """
    이제부터 화면(stdout)과 에러(stderr)에 나오는 모든 메시지를 path 파일에 씁니다.
    백그라운드 실행(start_hidden.vbs)일 때 app.py 맨 앞에서 부릅니다.
    """
    sys.stdout = sys.stderr = SizeLimitedLog(path, max_bytes)
