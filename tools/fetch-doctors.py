# -*- coding: utf-8 -*-
"""
amc.seoul.kr 에서 의료진 이름 → 소속 진료과 표를 받아 doctors.json 으로 저장합니다.

병원보 기사의 비고에는 "박준규, 안중호 (11.9)" 처럼 부서 없이 의료진 이름만 적힌 경우가 많습니다.
그 이름으로 부서를 찾으려고 만든 표입니다.

    python tools/fetch-doctors.py

진료과가 개편되면 다시 돌리면 됩니다. 164개 진료과·센터 페이지를 훑어 1~2분 걸립니다.
"""
import json
import re
import sys
import time
import urllib.request

BASE = 'https://www.amc.seoul.kr'
LIST = BASE + '/asan/staff/base/staffBaseInfoList.do'
UA = 'Mozilla/5.0 (compatible; amc-hospital-paper/1.0)'


def get(url):
    req = urllib.request.Request(url, headers={'User-Agent': UA})
    with urllib.request.urlopen(req, timeout=30) as r:
        return r.read().decode('utf-8', 'replace')


def dept_codes():
    """진료과 선택 상자에서 D001 같은 코드와 이름을 꺼냅니다."""
    html = get(LIST)
    out, seen = [], set()
    for m in re.finditer(r'<option[^>]*value="([^"]*depts/(D\d+)/[^"]*)"[^>]*>([^<]+)</option>', html):
        code, name = m.group(2), m.group(3).strip()
        if code not in seen:
            seen.add(code)
            out.append((code, name))
    return out


def doctors_of(code):
    """한 진료과의 의료진 목록. [(이름, [소속...]), ...]"""
    html = get(LIST + '?searchHpCd=' + code)
    out = []
    # .doctor_info 하나가 의료진 한 명, 그 안의 첫 표 행이 소속 진료과입니다
    for block in re.findall(r'<div class="doctor_info">(.*?)</table>', html, re.S):
        nm = re.search(r'<p class="doctor_name">.*?<a[^>]*>([^<]+)</a>', block, re.S)
        if not nm:
            continue
        name = nm.group(1).strip()
        row = re.search(r'<th[^>]*>\s*진료과\s*</th>\s*<td>(.*?)</td>', block, re.S)
        depts = re.findall(r'<a[^>]*>([^<]+)</a>', row.group(1)) if row else []
        depts = [d.strip() for d in depts if d.strip()]
        if name:
            out.append((name, depts))
    return out


def main():
    codes = dept_codes()
    print('진료과·센터 %d개' % len(codes))
    primary, every = {}, {}
    for i, (code, name) in enumerate(codes, 1):
        try:
            found = doctors_of(code)
        except Exception as e:
            print('  ! %s %s 실패: %s' % (code, name, e))
            continue
        for who, depts in found:
            dl = depts or [name]
            primary.setdefault(who, dl[0])      # 첫 소속을 대표 부서로 봅니다
            every.setdefault(who, [])
            for d in dl:
                if d not in every[who]:
                    every[who].append(d)
        if i % 20 == 0 or i == len(codes):
            print('  %d/%d · 의료진 %d명' % (i, len(codes), len(primary)))
        time.sleep(0.1)                          # 사이트에 부담 주지 않도록

    with open('doctors.json', 'w', encoding='utf-8') as f:
        json.dump({'t': time.strftime('%Y-%m-%d'), 'primary': primary, 'all': every},
                  f, ensure_ascii=False, separators=(',', ':'))
    print('doctors.json 저장 · 의료진 %d명' % len(primary))


if __name__ == '__main__':
    sys.exit(main())
