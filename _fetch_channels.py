# -*- coding: utf-8 -*-
"""
批量查询基金限制大额申购公告，识别适用渠道
"""
import sys
sys.stdout.reconfigure(encoding='utf-8')
import json
import subprocess
import openpyxl
import re
import os
import time
import urllib.request

XL = r'D:\opencode\project\oracle_xrisk\20261013_QDII额度\QDII额度.xlsx'
OUT_DIR = r'D:\opencode\project\oracle_xrisk\20261013_QDII额度\_tmp_pdfs'
os.makedirs(OUT_DIR, exist_ok=True)
QUERY_SCRIPT = r'C:\Users\yuanyuan\.workbuddy\skills\skill_2053082432761950208\scripts\query.py'
RESULT_FILE = r'D:\opencode\project\oracle_xrisk\20261013_QDII额度\_results.json'


def query_nd(q):
    r = subprocess.run(['python', QUERY_SCRIPT, '--query', q],
                       capture_output=True, text=True, encoding='utf-8', timeout=60)
    try:
        return json.loads(r.stdout)
    except Exception:
        return None


def parse_anns(j):
    out = []
    try:
        for blk in j['data']['apiData']['apiRecall']:
            tag = blk.get('tag', '')
            desc = blk.get('desc', '')
            type_ = blk.get('type', '')
            if not (('公告' in desc) or ('公告' in type_) or ('公告' in tag)):
                continue
            content = blk.get('content', '')
            for line in content.split('\n'):
                m = re.search(r'\|\s*(\d{4}-\d{2}-\d{2})\s*\|\s*([^|]+?)\s*\|\s*([^|]+?)\s*\|\s*(https?://[^\s|]+)', line)
                if m:
                    out.append({
                        'date': m.group(1),
                        'title': m.group(2).strip(),
                        'type': m.group(3).strip(),
                        'url': m.group(4).strip()
                    })
    except Exception:
        pass
    return out


def extract_pdf_text(pdf_path):
    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            return '\n'.join((p.extract_text() or '') for p in pdf.pages)
    except ImportError:
        try:
            import pypdf
            reader = pypdf.PdfReader(pdf_path)
            return '\n'.join((p.extract_text() or '') for p in reader.pages)
        except Exception:
            return ''


def download_pdf(url, code):
    safe_code = code.replace('/', '_').replace('\\', '_')
    fn = os.path.join(OUT_DIR, f'{safe_code}.pdf')
    if os.path.exists(fn):
        return fn
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        data = urllib.request.urlopen(req, timeout=20).read()
        with open(fn, 'wb') as f:
            f.write(data)
        return fn
    except Exception:
        return None


def normalize_name(name):
    n = name
    for s in ['人民币钞', '人民币汇', '美元钞', '美元汇', '人民币']:
        if n.endswith(s):
            n = n[:-len(s)]
    if n and n[-1] in 'ABCDEF':
        n = n[:-1]
    return n


def expand_name(name):
    """尝试生成多种查询名，覆盖短简称与全称"""
    n = normalize_name(name)
    candidates = [name, n]
    # 短简称补全
    if n.startswith('摩根新兴市场') and '全球' not in n:
        candidates.append(n.replace('摩根新兴市场', '摩根全球新兴市场混合'))
    if n == '富国蓝筹精选':
        candidates.append('富国蓝筹精选股票')
    if n == '富国亚洲收益':
        candidates.append('富国亚洲收益债券')
    if n == '华夏大中华企业精选':
        candidates.append('华夏大中华企业精选灵活配置混合')
    if n == '华宝海外新能源汽车':
        candidates.append('华宝海外新能源汽车股票')
    if n == '华宝致远':
        candidates.append('华宝标普美国品质')
    if n == '建信新兴市场优选':
        candidates.append('建信新兴市场优选混合')
    if n == '南方全球精选':
        candidates.append('南方全球精选配置')
    if n == '摩根亚太优势':
        candidates.append('摩根亚太优势混合')
    if n == '摩根全球多元':
        candidates.append('摩根全球多元配置')
    if n == '摩根海外稳健配置':
        candidates.append('摩根海外稳健配置混合')
    if n == '摩根富时REITs':
        candidates.append('摩根富时发达市场REITs')
    if n == '博时亚洲票息':
        candidates.append('博时亚洲票息收益债券')
    if n == '博时大中华':
        candidates.append('博时大中华亚太精选股票')
    if n == '博时抗通胀':
        candidates.append('博时抗通胀增强回报')
    return candidates


def channel_from_text(text):
    if '上述调整仅适用于投资者通过直销渠道' in text or '仅适用于投资者通过直销渠道' in text:
        return '直销'
    if '上述调整仅适用于投资者通过代销' in text or '仅适用于投资者通过代销渠道' in text:
        return '代销'

    has_dx = bool(re.search(r'各代销|代销机构|代销渠道', text))
    has_ds = bool(re.search(r'直销柜台|直销渠道|直销平台|网上直销|直销中心|直销 APP|本公司直销', text))

    special = []
    if '博时财富' in text:
        special.append('博时财富')
    if '京东肯特' in text or '京东平台' in text:
        special.append('京东肯特')
    if '天天基金' in text:
        special.append('天天基金')

    if special and has_ds:
        return '直销+代销（含' + '、'.join(special) + '）'
    if special and has_dx and not has_ds:
        return '代销（含' + '、'.join(special) + '）'
    if has_ds and not has_dx:
        return '直销'
    if has_dx and not has_ds:
        return '代销'
    if has_ds and has_dx:
        return '直销+代销'
    return '全部渠道'


def title_channel_hint(title):
    if '直销渠道' in title or ('直销' in title and '代销' not in title):
        return '直销'
    if '代销渠道' in title or ('代销' in title and '直销' not in title):
        return '代销'
    return '全部渠道'


def get_anns_for(name):
    """尝试多种命名变体查询"""
    candidates = expand_name(name)
    for cand in candidates:
        for q in [f'{cand} 公告', f'{cand}']:
            j = query_nd(q)
            if not j:
                continue
            anns = parse_anns(j)
            if anns:
                return anns, q, cand
    return [], '', ''


def process_one(code, name, limit):
    anns, q_used, matched_name = get_anns_for(name)
    kws = ['调整大额申购', '暂停大额申购', '限制申购', '暂停申购', '调整申购',
           '调整直销', '调整代销', '大额申购']
    relevant = [a for a in anns if any(k in a['title'] for k in kws)]
    relevant.sort(key=lambda x: x['date'], reverse=True)
    if not relevant:
        return {'code': code, 'name': name, 'limit': limit,
                'channel': '未找到公告', 'ann_date': '', 'ann_title': '',
                'note': f'q={q_used} matched={matched_name}'}
    latest = relevant[0]
    pdf = download_pdf(latest['url'], code)
    if pdf:
        text = extract_pdf_text(pdf)
        ch = channel_from_text(text)
    else:
        ch = title_channel_hint(latest['title'])
    return {'code': code, 'name': name, 'limit': limit,
            'channel': ch, 'ann_date': latest['date'], 'ann_title': latest['title'][:80],
            'note': f'q={q_used} matched={matched_name}'}


if __name__ == '__main__':
    wb = openpyxl.load_workbook(XL, data_only=True)
    ws = wb.active
    funds = []
    for i, row in enumerate(ws.iter_rows(values_only=True), 1):
        if i == 1 or row[4] is None:
            continue
        funds.append((row[0], row[1], row[4]))

    print(f'total funds with limit: {len(funds)}')

    # 加载已有结果（断点续跑）
    results = []
    if os.path.exists(RESULT_FILE):
        with open(RESULT_FILE, 'r', encoding='utf-8') as f:
            results = json.load(f)
    done_codes = {r['code'] for r in results}

    for idx, (code, name, limit) in enumerate(funds):
        if code in done_codes:
            continue
        try:
            r = process_one(code, name, limit)
            results.append(r)
            with open(RESULT_FILE, 'w', encoding='utf-8') as f:
                json.dump(results, f, ensure_ascii=False, indent=2)
            print(f'[{idx+1}/{len(funds)}] {code} {name[:20]} -> {r["channel"]}')
        except Exception as e:
            print(f'[{idx+1}/{len(funds)}] ERR {code}: {e}')
        time.sleep(0.5)

    print(f'\ndone. total: {len(results)}')
