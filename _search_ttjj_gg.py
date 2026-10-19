# -*- coding: utf-8 -*-
"""
对未找到公告的基金，从天天基金网公告页查找具体公告
"""
import sys
sys.stdout.reconfigure(encoding='utf-8')
import json
import ssl
import urllib.request
import re
import os

RESULT_FILE = r'D:\opencode\project\oracle_xrisk\20261013_QDII额度\_results.json'
TTJJ_DIR = r'D:\opencode\project\oracle_xrisk\20261013_QDII额度\_ttjj_cache'
GG_DIR = r'D:\opencode\project\oracle_xrisk\20261013_QDII额度\_ttjj_gg'
os.makedirs(GG_DIR, exist_ok=True)

CTX = ssl.create_default_context()
CTX.check_hostname = False
CTX.verify_mode = ssl.CERT_NONE
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
}


def fetch_gg(code):
    """fetch 天天基金网公告列表"""
    code_short = code.replace('.OF', '')
    fn = os.path.join(GG_DIR, f'{code_short}.html')
    if os.path.exists(fn):
        with open(fn, 'r', encoding='utf-8', errors='ignore') as f:
            return f.read()
    # Try multiple pages (first page only for now)
    url = f'https://fundf10.eastmoney.com/jjgg_{code_short}.html'
    try:
        req = urllib.request.Request(url, headers=HEADERS)
        data = urllib.request.urlopen(req, timeout=15, context=CTX).read()
        with open(fn, 'wb') as f:
            f.write(data)
        return data.decode('utf-8', errors='ignore')
    except Exception as e:
        return None


def search_gg_for_limit(code):
    """查找公告中包含'限制大额申购'或'调整大额申购'的链接"""
    html = fetch_gg(code)
    if not html:
        return None
    # Look for announcements with relevant keywords
    # Common pattern: <a href="...gonggao/CODENAME,ANXXXXXXXXXXXXXX.html" ...>公告标题</a>
    links = re.findall(r'<a[^>]*href="([^"]*gonggao/[^"]+)"[^>]*>([^<]+)</a>', html)
    relevant = []
    for href, title in links:
        if any(kw in title for kw in ['调整大额申购', '限制大额申购', '暂停大额申购', '调整申购',
                                       '限制申购', '暂停申购', '调整直销', '调整代销']):
            relevant.append((href, title.strip()))
    return relevant


if __name__ == '__main__':
    with open(RESULT_FILE, 'r', encoding='utf-8') as f:
        results = json.load(f)
    missing = [r for r in results if r['channel'] == '未找到公告']
    print(f'missing count: {len(missing)}')

    for r in missing[:5]:
        code = r['code']
        print(f'\n--- {code} {r["name"]} ---')
        links = search_gg_for_limit(code)
        if links:
            for href, title in links[:3]:
                print(f'  [{title[:50]}] {href}')
        else:
            print('  no relevant announcement found on ttjj')
