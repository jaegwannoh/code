# -*- coding: utf-8 -*-
"""
excel.py - 엑셀 파일 만들기 (openpyxl 라이브러리 사용)

파일을 디스크에 저장하지 않고 메모리(BytesIO)에서 만들어
바로 브라우저로 내려보냅니다.

[Py2.7 차이] 회사 PC의 openpyxl 은 2.x 옛 버전일 수 있습니다.
             옛 버전에도 있는 기능만 쓰도록 짰습니다.
             (ws['A'], ws.iter_rows(min_row=..), wb.remove() 같은 새 기능은 안 씀)
"""
from __future__ import unicode_literals

from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

try:
    from openpyxl.utils import get_column_letter        # openpyxl 2.4 이상
except ImportError:
    from openpyxl.cell import get_column_letter         # 그보다 옛 버전

from compat import to_text

HEADER_FILL = PatternFill(fill_type='solid', start_color='FFEEF0EC', end_color='FFEEF0EC')


def _new_workbook():
    """빈 엑셀 파일. 기본으로 생기는 빈 시트(Sheet)는 지웁니다."""
    wb = Workbook()
    if hasattr(wb, 'remove'):
        wb.remove(wb.active)            # openpyxl 2.4 이상
    else:
        wb.remove_sheet(wb.active)      # 옛 버전
    return wb


def _add_sheet(wb, title, headers, rows):
    """시트 하나를 만들어 머리글 + 데이터를 채웁니다."""
    ws = wb.create_sheet(title=title)
    # 칸마다 가장 긴 글자 수를 기억해 두었다가 칸 너비를 맞춥니다.
    longest = [len(h) for h in headers]

    # 1줄: 머리글
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col)
        cell.value = h
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal='center')

    # 2줄부터: 데이터. 긴 글은 줄바꿈해서 보이게
    wrap = Alignment(wrap_text=True, vertical='top')
    for r, row in enumerate(rows, 2):
        for col, value in enumerate(row, 1):
            text = to_text(value)
            cell = ws.cell(row=r, column=col)
            cell.value = text
            cell.alignment = wrap
            longest[col - 1] = max(longest[col - 1], len(text))

    # 칸 너비를 글자 수에 맞춰 대충 조절 (한글은 한 글자가 2칸 정도)
    for col, n in enumerate(longest, 1):
        ws.column_dimensions[get_column_letter(col)].width = min(n * 2 + 2, 60)


def _to_bytes(wb):
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def day_workbook(d, data, teams):
    """하루치 업무일지 → 시트: 팀별 일일업무 / 3팀 신호처리 / 5팀 분석현황 / 휴가·부재"""
    wb = _new_workbook()

    for t in teams:
        works = data['teams'][t]['works']
        missing = data['teams'][t]['missing']
        rows = [(d, w['grade'], w['name'], w['title'], w['detail']) for w in works]
        if missing:
            names = ', '.join('{} {}'.format(u['grade'], u['name']) for u in missing)
            rows.append(('', '', '', '미제출 {}명'.format(len(missing)), names))
        _add_sheet(wb, t + ' 일일업무', ['날짜', '급수', '이름', '업무내용', '세부내용'], rows)

    _add_sheet(wb, '3팀 신호처리', ['날짜', '급수', '이름', '위성명', 'CaseNotation'],
               [(s['work_date'], s['grade'], s['name'], s['satellite'], s['case_notation'])
                for s in data['signals']])

    rows = [('분석중', a['manager'], a['start_date'], '', a['case_notation'], a['note'])
            for a in data['ongoing']]
    rows += [('분석종료', a['manager'], a['start_date'], a['end_date'], a['case_notation'], a['note'])
             for a in data['finished']]
    _add_sheet(wb, '5팀 분석현황', ['상태', '담당자', '시작일', '종료일', 'CaseNotation', '비고'], rows)

    _add_sheet(wb, '휴가·부재', ['팀', '급수', '이름', '종류', '시작일', '종료일', '비고'],
               [(l['team'], l['grade'], l['name'], l['kind'], l['start_date'], l['end_date'], l['note'])
                for l in data['leaves']])
    return _to_bytes(wb)


def search_workbook(results):
    """검색 결과 → 시트 하나"""
    wb = _new_workbook()
    _add_sheet(wb, '검색결과', ['날짜', '구분', '팀', '작성자', '내용'],
               [(r['date'], r['kind'], r['team'], r['person'], r['content']) for r in results])
    return _to_bytes(wb)
