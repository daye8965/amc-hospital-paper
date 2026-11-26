# -*- coding: utf-8 -*-
"""
기구조직도 엑셀의 (변경) 탭에서 조직 단위 목록을 뽑아 org-units.json 으로 저장합니다.

    python tools/extract-org.py "C:\\path\\기구 조직도 변경(안).xlsx"

조직도는 표가 아니라 그림처럼 그려져 있어서, 셀이 놓인 '행 번호'가 곧 계층입니다.
조직 개편으로 행 위치가 바뀌면 아래 ROW_CAT 을 손봐야 합니다.

뽑은 결과는 index.html 의 ORG_UNITS 에 붙여 넣습니다(런타임에 이 파일을 읽지는 않습니다).
"""
import collections
import json
import re
import sys

import openpyxl

# 행 번호 → 조직 계층
ROW_CAT = {14: '직속 조직', 18: '행정·지원', 19: '행정·지원', 39: '전문병원',
           41: '센터', 43: '센터', 57: '진료지원', 69: '진료과', 92: '진료지원'}

# 81행은 진료 분과·간호·행정이 한 줄에 섞여 있어 이름으로 가릅니다
NURSE = set('암병원간호1팀 암병원간호2팀 내과간호1팀 내과간호2팀 외과간호1팀 외과간호2팀 심장병원간호팀 '
            '어린이병원간호팀 수술간호팀 중환자간호팀 응급간호팀 외래간호팀 중앙공급팀 간호교육행정팀 '
            '통합돌봄지원팀'.split())
SUPPORT = set('재활의학팀 방사선종양팀 영상의학팀 핵의학팀 방사성의약품제조소 진단검사의학팀 병리팀 '
              '약제팀 영양팀 특수검사팀'.split())
EDU = set('교육수련팀 아카데미운영팀 시뮬레이션센터 임상해부교육센터'.split())
RESEARCH = set('연구기획팀 연구지원팀 R&D사업팀'.split())
ADMIN = set('인사팀 노사협력팀 재무회계팀 원무팀 법무팀 적정진료팀 사회복지팀 총무팀 자재팀 시설팀 '
            '보안관리팀 후생사업팀 의공팀'.split())

# 기사에서 실제로 쓰이는 약칭
ALIASES = [('IDC', '행정·지원'), ('디콘', '행정·지원')]


def main(path):
    ws = openpyxl.load_workbook(path, data_only=True)['변경']
    by_row = collections.defaultdict(list)
    for row in ws.iter_rows():
        for c in row:
            if c.value is None:
                continue
            text = ' '.join(str(c.value).split())
            if text:
                by_row[c.row].append(text)

    units, seen = [], set()
    for r, names in by_row.items():
        if r in (1, 3):            # 제목과 기준일
            continue
        for raw in names:
            flat = re.sub(r'\s+', '', raw)      # 세로쓰기라 '아 카 데 미' 처럼 들어옵니다
            if not flat or flat in seen:
                continue
            if r == 81:
                cat = ('간호' if flat in NURSE else '진료지원' if flat in SUPPORT
                       else '교육' if flat in EDU else '연구' if flat in RESEARCH
                       else '행정·지원' if flat in ADMIN else '진료 분과')
            else:
                cat = ROW_CAT.get(r)
                if cat is None:
                    continue
            seen.add(flat)
            units.append({'n': flat, 'c': cat})

    for name, cat in ALIASES:
        if name not in seen:
            units.append({'n': name, 'c': cat})

    with open('org-units.json', 'w', encoding='utf-8') as f:
        json.dump(units, f, ensure_ascii=False, separators=(',', ':'))

    counts = collections.Counter(u['c'] for u in units)
    print('조직 %d곳' % len(units))
    for cat, n in counts.most_common():
        print('  %-8s %d' % (cat, n))


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    main(sys.argv[1])
