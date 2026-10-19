# -*- coding: utf-8 -*-
"""
对未找到公告的基金，查询天天基金网确认渠道
"""
import sys
sys.stdout.reconfigure(encoding='utf-8')
import json
import ssl
import urllib.request
import re
import os

XL = r'D:\opencode\project\oracle_xrisk\20261013_QDII额度\QDII额度.xlsx'
RESULT_FILE = r'D:\opencode\project\oracle_xrisk\20261013_QDII额度\_results.json'
TTJJ_DIR = r'D:\opencode\project\oracle_xrisk\20261013_QDII额度\_ttjj_cache'

os.makedirs(TTJJ_DIR, exist_ok=True)

CTX = ssl.create_default_context()
CTX.check_hostname = False
CTX.verify_mode = ssl.CERT_NONE
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
}


def fetch_page(code):
    """fetch and cache the page"""
    code_short = code.replace('.OF', '')
    fn = os.path.join(TTJJ_DIR, f'{code_short}.html')
    if os.path.exists(fn):
        with open(fn, 'r', encoding='utf-8', errors='ignore') as f:
            return f.read()
    url = f'https://fundf10.eastmoney.com/jjjz_{code_short}.html'
    try:
        req = urllib.request.Request(url, headers=HEADERS)
        data = urllib.request.urlopen(req, timeout=15, context=CTX).read()
        with open(fn, 'wb') as f:
            f.write(data)
        return data.decode('utf-8', errors='ignore')
    except Exception as e:
        return None


def parse_state(html):
    """parse 交易状态 and limit from html"""
    if not html:
        return None, None
    # find 交易状态 followed by state
    m = re.search(r'交易状态[：:]\s*</?\w*>\s*([^<\n\r]+)', html)
    state = m.group(1).strip() if m else None
    # find 购买上限
    m2 = re.search(r'单日累计购买上限[^<\n\r]*?(\d[\d,\.]*\d)[^<\n\r]*?元', html)
    limit = m2.group(1).replace(',', '') if m2 else None
    # find 限购 类型
    state_m = re.search(r'(限大额|限购|暂停申购|开放申购|限制大额申购)', html)
    state_simple = state_m.group(1) if state_m else None
    return state_simple, limit


if __name__ == '__main__':
    with open(RESULT_FILE, 'r', encoding='utf-8') as f:
        results = json.load(f)
    missing = [r for r in results if r['channel'] == '未找到公告']
    print(f'missing count: {len(missing)}')

    for r in missing:
        code = r['code']
        print(f'\n--- {code} {r["name"]} (limit={r["limit"]}) ---')
        html = fetch_page(code)
        if html:
            state, limit = parse_state(html)
            print(f'  ttjj state: {state}, ttjj limit: {limit}')
        else:
            print('  fetch failed')
