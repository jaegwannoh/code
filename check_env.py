# -*- coding: utf-8 -*-
"""
check_env.py - 회사 PC에 필요한 모듈이 있는지 확인하는 파일

실행:  python check_env.py
       (run.bat 과 같은 python 을 써야 합니다. Anaconda Prompt 에서 실행하면 편합니다.)
문제가 있으면 이 창의 내용을 그대로 복사해서 알려주세요.
"""
from __future__ import print_function

import sys

print('python   :', sys.version.split()[0], '(' + sys.executable + ')')

ok = True
for name in ('sqlite3', 'flask', 'jinja2', 'werkzeug', 'markupsafe', 'openpyxl'):
    try:
        module = __import__(name)
        version = getattr(module, '__version__', '')
        if name == 'sqlite3':
            version = 'SQLite ' + module.sqlite_version
        print('%-9s: OK  %s' % (name, version))
    except Exception as e:
        ok = False
        print('%-9s: MISSING  (%s)' % (name, e))

print()
print('All good. Run run.bat' if ok else 'Some modules are missing. Send this screen.')
