# -*- coding: utf-8 -*-
"""
天天基金「海外额度宝」额度排行数据抓取

数据源：天天基金移动端公开接口 fundmobapi.eastmoney.com
  1) FundMNRank?FundType=6   -> QDII 基金全集（代码/名称/近1年/规模/净值）
  2) FundBaseTypeInformation -> 申购状态(含单日投资上限)、基金类型、公司、规模

输出：_quota_raw.json
"""
import sys, os, json, time, ssl, re, random
import urllib.request, urllib.error

sys.stdout.reconfigure(encoding='utf-8')

BASE = "https://fundmobapi.eastmoney.com"
HDR = {
    "User-Agent": "Mozilla/5.0 (iPhone; CPU iPhone OS 16_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Mobile/15E148",
    "Referer": "https://fund.eastmoney.com/",
    "Accept": "application/json, text/plain, */*",
}
CTX = ssl.create_default_context()
CTX.check_hostname = False
CTX.verify_mode = ssl.CERT_NONE

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_quota_raw.json")
COMMON = "deviceid=Wap&plat=Wap&product=EFund&version=2.0.0"


def fetch(url, retry=3):
    for i in range(retry):
        try:
            req = urllib.request.Request(url, headers=HDR)
            with urllib.request.urlopen(req, timeout=15, context=CTX) as r:
                return r.read().decode("utf-8", "ignore")
        except Exception as e:
            if i == retry - 1:
                return None
            time.sleep(1.0 + i)
    return None


def get_qdii_universe():
    """分页拉取 QDII 基金全集"""
    out, page = [], 1
    while True:
        u = (f"{BASE}/FundMNewApi/FundMNRank?FundType=6&SortColumn=SYL_1N&Sort=desc"
             f"&pageIndex={page}&pageSize=30&plat=Android&appType=ttjj&product=EFund"
             f"&Version=1&deviceid=Wap&_={int(time.time()*1000)}")
        d = fetch(u)
        if not d:
            print(f"  page {page} fetch failed, stop")
            break
        try:
            j = json.loads(d)
        except Exception:
            print(f"  page {page} parse failed")
            break
        ds = j.get("Datas") or []
        if not ds:
            break
        out.extend(ds)
        total = j.get("TotalCount") or 0
        print(f"  page {page}: +{len(ds)} (cum {len(out)}/{total})")
        if len(out) >= total or len(ds) < 30:
            break
        page += 1
        time.sleep(0.4)
    return out


def parse_sgzt(sgzt):
    """解析申购状态文本 -> (状态, 单日上限文本)"""
    s = (sgzt or "").strip()
    return s


def get_detail(code):
    u = f"{BASE}/FundMApi/FundBaseTypeInformation.ashx?FCODE={code}&{COMMON}&_={int(time.time()*1000)}"
    for _ in range(3):
        d = fetch(u)
        if d:
            try:
                j = json.loads(d)
            except Exception:
                j = None
            if j and str(j.get("ErrCode")) == "0" and j.get("Datas"):
                return j.get("Datas") or {}
        time.sleep(1.0)
    return {}


def main():
    print("== step1: QDII universe ==")
    univ = get_qdii_universe()
    codes = []
    seen = set()
    for r in univ:
        c = r.get("FCODE")
        if c and c not in seen:
            seen.add(c)
            codes.append(r.get("FCODE"))
    print("universe:", len(codes))

    prev = {}
    if os.path.exists(OUT):
        try:
            prev = {x["FCODE"]: x for x in json.load(open(OUT, encoding="utf-8"))}
            print("resume from", len(prev))
        except Exception:
            prev = {}

    results = [prev[c] for c in codes if c in prev]
    todo = [c for c in codes if c not in prev]
    print("todo:", len(todo))

    for i, code in enumerate(todo, 1):
        d = get_detail(code)
        if d:
            d["FCODE"] = code
            results.append(d)
        else:
            results.append({"FCODE": code, "_failed": True})
        if i % 25 == 0 or i == len(todo):
            json.dump(results, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
            print(f"  {i}/{len(todo)} saved")
        time.sleep(0.30 + random.random() * 0.20)

    json.dump(results, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("done, total", len(results))


if __name__ == "__main__":
    main()
