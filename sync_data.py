# -*- coding: utf-8 -*-
"""
QDII「海外额度宝」额度数据同步

数据源：天天基金移动端公开接口 fundmobapi.eastmoney.com
  1) FundMNRank?FundType=6          -> QDII 基金全集（代码/名称/近1年/规模/净值）
  2) FundBaseTypeInformation.ashx   -> 申购状态(含单日投资上限)、类型、公司、规模

产物：
  _quota_raw.json              原始接口数据（兼容旧脚本）
  data/quota_data.json         加工后的看板数据（Web 服务读取）
  天天基金_海外额度宝_额度排行.xlsx   Excel 排行榜
  QDII额度看板.html             离线静态快照（数据已内嵌）

用法：
  python sync_data.py                 # 全量重新抓取并生成全部产物
  python sync_data.py --reuse-raw     # 不抓数，直接用已有的 _quota_raw.json 重建产物
  python sync_data.py --workers 6     # 并发抓取线程数（默认 6）
  python sync_data.py --no-excel --no-static
"""
from __future__ import annotations

import argparse
import json
import os
import random
import re
import ssl
import sys
import time
import urllib.request
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime

sys.stdout.reconfigure(encoding="utf-8")

HERE = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(HERE)  # oracle_xrisk 目录（含 xrisk_db 包）
sys.path.insert(0, HERE)
sys.path.insert(0, PROJECT_ROOT)

# 复用 Excel 脚本里的业务口径（区域推断 / 额度解析），保持单一数据源
from _build_excel import infer_region, parse_limit, fmt_size, classify_qdii  # noqa: E402

DATA_DIR = os.path.join(HERE, "data")
RAW_JSON = os.path.join(HERE, "_quota_raw.json")
UNIVERSE_JSON = os.path.join(HERE, "_universe.json")
WEB_JSON = os.path.join(DATA_DIR, "quota_data.json")
TEMPLATE = os.path.join(HERE, "templates", "index.html")
STATIC_HTML = os.path.join(HERE, "QDII额度看板.html")

BASE = "https://fundmobapi.eastmoney.com"
COMMON = "deviceid=Wap&plat=Wap&product=EFund&version=2.0.0"
HDR = {
    "User-Agent": ("Mozilla/5.0 (iPhone; CPU iPhone OS 16_0 like Mac OS X) "
                   "AppleWebKit/605.1.15 (KHTML, like Gecko) Mobile/15E148"),
    "Referer": "https://fund.eastmoney.com/",
    "Accept": "application/json, text/plain, */*",
}
CTX = ssl.create_default_context()
CTX.check_hostname = False
CTX.verify_mode = ssl.CERT_NONE

STATUS_ORDER = {"开放申购": 0, "限大额": 1, "封闭期": 2, "暂停申购": 3, "场内交易": 4, "认购期": 5}


def log(msg: str) -> None:
    print(f"[{datetime.now():%H:%M:%S}] {msg}", flush=True)


BAD_RESPONSE_LOG = os.path.join(HERE, "_bad_response.log")
MIN_UNIVERSE = int(os.environ.get("QDII_MIN_UNIVERSE", "600"))   # 全集健康下限
MIN_ROWS = int(os.environ.get("QDII_MIN_ROWS", "500"))           # 入库健康下限


def save_bad(tag: str, text: str | None) -> None:
    """把异常响应落盘，便于事后定位（只保留最近若干条）"""
    try:
        with open(BAD_RESPONSE_LOG, "a", encoding="utf-8") as f:
            f.write(f"\n===== {datetime.now():%Y-%m-%d %H:%M:%S} | {tag} =====\n")
            f.write(((text or "")[:800]).replace("\r", " ").replace("\n", " ") + "\n")
    except Exception:
        pass


def fetch(url: str, retry: int = 3, timeout: int = 15):
    for i in range(retry):
        try:
            req = urllib.request.Request(url, headers=HDR)
            with urllib.request.urlopen(req, timeout=timeout, context=CTX) as r:
                return r.read().decode("utf-8", "ignore")
        except Exception:
            if i == retry - 1:
                return None
            time.sleep(0.8 + i * 0.7)
    return None


def parse_json(text: str | None):
    """宽松解析：剥离 BOM/空白，区分各类异常。返回 (obj, err)"""
    if text is None:
        return None, "请求失败"
    t = (text or "").lstrip("\ufeff").strip()
    if not t:
        return None, "响应为空"
    if t[0] not in "{[":
        return None, f"非JSON响应(head={t[:40]!r})"
    try:
        j = json.loads(t)
    except Exception as e:
        return None, f"JSON解析错误({e})"
    if not isinstance(j, dict):
        return None, f"顶层非对象(type={type(j).__name__}, val={str(j)[:40]!r})"
    return j, None


# --------------------------------------------------------------------------- #
# 取数
# --------------------------------------------------------------------------- #
def _rank_page(page: int):
    """抓取单页，返回 (datas, total, err)"""
    u = (f"{BASE}/FundMNewApi/FundMNRank?FundType=6&SortColumn=SYL_1N&Sort=desc"
         f"&pageIndex={page}&pageSize=30&plat=Android&appType=ttjj&product=EFund"
         f"&Version=1&deviceid=Wap&_={int(time.time() * 1000)}")
    d = fetch(u)
    j, err = parse_json(d)
    if err:
        save_bad(f"rank page {page}", d)
        return [], 0, err
    return (j.get("Datas") or []), int(j.get("TotalCount") or 0), None


def _paged_once() -> tuple:
    """单轮完整分页抓取，返回 (list, err)"""
    out, page, total = [], 1, 0
    while page <= 40:
        ds, tc, err = _rank_page(page)
        if err:
            # 单页异常再试 2 次（接口偶发返回 null/HTML/风控页）
            for _ in range(2):
                time.sleep(1.5)
                ds, tc, err = _rank_page(page)
                if not err:
                    break
            if err:
                return out, f"第 {page} 页 {err}"
        total = tc or total
        if not ds:
            break
        out.extend(ds)
        log(f"  第 {page} 页 +{len(ds)}（累计 {len(out)}/{total}）")
        if (total and len(out) >= total) or len(ds) < 30:
            break
        page += 1
        time.sleep(0.3)
    return out, None


def _prev_universe() -> list:
    """兜底：上一次成功抓取的全集"""
    try:
        if os.path.exists(UNIVERSE_JSON):
            u = json.load(open(UNIVERSE_JSON, encoding="utf-8"))
            if isinstance(u, list) and len(u) >= MIN_UNIVERSE:
                return u
    except Exception:
        pass
    return []


def get_qdii_universe(rounds: int = 3) -> tuple:
    """分页拉取 QDII 基金全集。返回 (list, source)
    source: live=实时抓取 / fallback=复用上次全集 / fail=彻底失败
    """
    last_err = ""
    for r in range(rounds):
        out, err = _paged_once()
        if err or len(out) < MIN_UNIVERSE:
            last_err = err or f"仅 {len(out)} 只（<{MIN_UNIVERSE}）"
            log(f"  全集第 {r + 1}/{rounds} 轮异常：{last_err}")
            if r < rounds - 1:
                wait = 20 * (r + 1)
                log(f"  {wait}s 后重试…")
                time.sleep(wait)
            continue
        log(f"  全集抓取成功：{len(out)} 只")
        return out, "live"

    fb = _prev_universe()
    if fb:
        log(f"  实时全集不可用（{last_err}），复用上次全集 {len(fb)} 只继续抓明细")
        return fb, "fallback"
    log(f"  全集抓取失败且无可用兜底：{last_err}")
    return [], "fail"


def get_detail(code: str) -> dict:
    u = (f"{BASE}/FundMApi/FundBaseTypeInformation.ashx?FCODE={code}"
         f"&{COMMON}&_={int(time.time() * 1000)}")
    for _ in range(3):
        d = fetch(u)
        if d:
            try:
                j = json.loads(d)
            except Exception:
                j = None
            if j and str(j.get("ErrCode")) == "0" and j.get("Datas"):
                return j["Datas"] or {}
        time.sleep(0.7)
    return {}


def fetch_all_details(codes: list, workers: int = 6) -> dict:
    """并发抓取明细，返回 {code: detail}"""
    res, done, failed = {}, 0, []
    total = len(codes)
    with ThreadPoolExecutor(max_workers=workers) as ex:
        futs = {ex.submit(get_detail, c): c for c in codes}
        for f in as_completed(futs):
            code = futs[f]
            try:
                d = f.result()
            except Exception:
                d = {}
            if d:
                d["FCODE"] = code
                res[code] = d
            else:
                failed.append(code)
            done += 1
            if done % 50 == 0 or done == total:
                log(f"  明细 {done}/{total}（失败 {len(failed)}）")
    if failed:
        log(f"  重试失败明细 {len(failed)} 只…")
        for code in failed:
            d = get_detail(code)
            if d:
                d["FCODE"] = code
                res[code] = d
    return res


# --------------------------------------------------------------------------- #
# 加工
# --------------------------------------------------------------------------- #
def fetch_region_internal() -> dict:
    """从 XRISK Oracle 查询 QDII 基金主投区域（内部口径）
    返回 {fund_code: region_label}
    """
    try:
        from xrisk_db import XriskDB
        with XriskDB() as db:
            df = db.extractSQL("""
                SELECT f.i_code1 AS code, t.factor_name AS region_internal
                FROM xrisk.vcrp_efd_fund_type f, xrisk.TFND_STYLE_FACTOR t
                WHERE t.f_code = f.i_code1
                  AND f.class_name1_yh LIKE '%QDII%'
                  AND t.report_date = '2026-06-30'
                  AND t.report_flag = '年报'
                  AND t.factor = 'NLABEL_STYLE'
            """)
        mapping = {}
        for _, row in df.iterrows():
            code = str(row["CODE"]).strip()
            mapping[code] = str(row["REGION_INTERNAL"]).strip()
        log(f"Oracle 主投区域（内部）：{len(mapping)} 只")
        return mapping
    except Exception as e:
        log(f"Oracle 主投区域查询失败（忽略）：{e}")
        return {}


REGION_INTERNAL_JSON = os.path.join(HERE, "_region_internal.json")


def map_region_tab(region_internal):
    """把内部区域标签映射到5个 tab 类别"""
    r = (region_internal or "").strip()
    if r == "美国":
        return "美国"
    if r == "欧洲":
        return "欧洲"
    if r == "新兴市场":
        return "新兴市场"
    # 大中华：港股XX、全市场、中概
    if r in ("港股宽基", "港股科技", "港股消费", "港股医药", "港股红利", "港股金融",
             "全市场", "中概", "红利"):
        return "大中华"
    # 其余归为其他
    return "其他"


def num(v):
    try:
        f = float(v)
        return round(f, 2)
    except Exception:
        return None


def build_records(details: dict, universe: list, region_map: dict = None) -> list:
    """加工为看板记录"""
    if region_map is None:
        region_map = {}
    meta = {}
    for r in universe:
        c = r.get("FCODE")
        if c and c not in meta:
            meta[c] = r

    rows = []
    codes = list(dict.fromkeys([r.get("FCODE") for r in universe if r.get("FCODE")]))
    for code in codes:
        d = details.get(code) or {}
        m = meta.get(code) or {}
        name = (d.get("SHORTNAME") or m.get("SHORTNAME") or "").strip()
        if not name:
            continue
        ftype = d.get("FTYPE") or m.get("FTYPE") or ""
        lv, ltxt, status = parse_limit(d.get("SGZT"))
        region_internal = region_map.get(code, "")
        rows.append({
            "code": code,
            "name": name,
            "ftype": ftype,
            "qdii_cat": classify_qdii(ftype),
            "region": infer_region(name, ftype),
            "region_internal": region_internal,
            "region_tab": map_region_tab(region_internal),
            "limit_val": lv,
            "limit_txt": ltxt,
            "status": status,
            "sgzt_raw": d.get("SGZT") or "",
            "ret_1d": num(d.get("RZDF") or m.get("RZDF")),
            "ret_1w": num(d.get("SYL_Z") or m.get("SYL_Z")),
            "ret_1m": num(d.get("SYL_Y") or m.get("SYL_Y")),
            "ret_3m": num(d.get("SYL_3Y") or m.get("SYL_3Y")),
            "ret_1n": num(d.get("SYL_1N") or m.get("SYL_1N")),
            "ret_ytd": num(m.get("SYL_JN")),
            "nav": d.get("DWJZ") or m.get("DWJZ") or "",
            "nav_date": d.get("FSRQ") or m.get("FSRQ") or "",
            "company": d.get("JJGS") or "",
            "size": fmt_size(d.get("ENDNAV")),
            "has_amount": lv is not None,
        })

    def sort_key(r):
        g = STATUS_ORDER.get(r["status"], 9)
        if g == 1 and r["limit_val"] is not None:
            return (g, -r["limit_val"], r["name"])
        return (g, 0, r["name"])

    rows.sort(key=sort_key)
    return rows


def build_stats(rows: list) -> dict:
    by_status = Counter(r["status"] for r in rows)
    by_region = Counter(r["region"] for r in rows)
    by_company = Counter(r["company"] for r in rows if r["company"])
    amounts = [r["limit_val"] for r in rows if r["limit_val"] is not None]
    return {
        "total": len(rows),
        "by_status": [{"k": k, "v": v} for k, v in
                      sorted(by_status.items(), key=lambda x: (STATUS_ORDER.get(x[0], 9), -x[1]))],
        "by_region": [{"k": k, "v": v} for k, v in by_region.most_common()],
        "by_company_top": [{"k": k, "v": v} for k, v in by_company.most_common(12)],
        "open_count": by_status.get("开放申购", 0),
        "limit_count": by_status.get("限大额", 0),
        "pause_count": by_status.get("暂停申购", 0),
        "other_count": sum(v for k, v in by_status.items()
                           if k not in ("开放申购", "限大额", "暂停申购")),
        "undisclosed": sum(1 for r in rows if r["status"] == "限大额" and r["limit_val"] is None),
        "max_amount": max(amounts) if amounts else None,
    }


def render_static(data: dict) -> str | None:
    """把数据内嵌进模板，生成离线可双击打开的静态快照"""
    if not os.path.exists(TEMPLATE):
        return None
    html = open(TEMPLATE, encoding="utf-8").read()
    payload = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")
    marker = "window.__EMBEDDED__ = null; /*__EMBED_DATA__*/"
    if marker not in html:
        return None
    html = html.replace(marker, f"window.__EMBEDDED__ = {payload}; /*__EMBED_DATA__*/")
    with open(STATIC_HTML, "w", encoding="utf-8") as f:
        f.write(html)
    return STATIC_HTML


# --------------------------------------------------------------------------- #
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--reuse-raw", action="store_true", help="不抓数，直接用 _quota_raw.json 重建产物")
    ap.add_argument("--workers", type=int, default=6, help="明细并发线程数（默认 6）")
    ap.add_argument("--no-excel", action="store_true", help="不重新生成 Excel")
    ap.add_argument("--no-static", action="store_true", help="不生成静态 HTML 快照")
    args = ap.parse_args()

    t0 = time.time()
    os.makedirs(DATA_DIR, exist_ok=True)

    uni_src = "reuse"
    if args.reuse_raw:
        if not os.path.exists(RAW_JSON):
            log("缺少 _quota_raw.json，无法复用")
            sys.exit(1)
        raw = json.load(open(RAW_JSON, encoding="utf-8"))
        details = {x["FCODE"]: x for x in raw if x.get("FCODE")}
        # 尝试加载 universe（rank API 数据，含 SYL_JN 等字段）
        if os.path.exists(UNIVERSE_JSON):
            universe = json.load(open(UNIVERSE_JSON, encoding="utf-8"))
            log(f"复用已有原始数据：{len(universe)} 只（含 universe）")
        else:
            universe = list(details.values())
            log(f"复用已有原始数据：{len(universe)} 只（无 universe，YTD 可能缺失）")
    else:
        log("步骤 1/4：拉取 QDII 基金全集…")
        universe, uni_src = get_qdii_universe()
        codes = list(dict.fromkeys([r.get("FCODE") for r in universe if r.get("FCODE")]))
        log(f"QDII 基金全集：{len(codes)} 只（来源：{uni_src}）")
        if not codes:
            log("！全集为空，放弃本次同步（保留上一次数据）")
            sys.exit(2)

        log(f"步骤 2/4：抓取申购状态与单日额度（{args.workers} 并发）…")
        details = fetch_all_details(codes, workers=args.workers)

        raw = [details[c] for c in codes if c in details]
        for c in codes:
            if c not in details:
                raw.append({"FCODE": c, "_failed": True})
        json.dump(raw, open(RAW_JSON, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        json.dump(universe, open(UNIVERSE_JSON, "w", encoding="utf-8"), ensure_ascii=False)
        log(f"原始数据已保存：{RAW_JSON}")

    log("步骤 3/4：加工数据…")
    # 从 Oracle 查询内部主投区域
    if os.path.exists(REGION_INTERNAL_JSON) and args.reuse_raw:
        region_map = json.load(open(REGION_INTERNAL_JSON, encoding="utf-8"))
        log(f"复用内部区域数据：{len(region_map)} 只")
    else:
        region_map = fetch_region_internal()
        if region_map:
            json.dump(region_map, open(REGION_INTERNAL_JSON, "w", encoding="utf-8"), ensure_ascii=False)
    rows = build_records(details, universe, region_map)

    # ---- 数据健康校验：异常结果绝不覆盖上一次的好数据 ----
    prev_total = 0
    try:
        if os.path.exists(WEB_JSON):
            prev_total = len(json.load(open(WEB_JSON, encoding="utf-8")).get("items") or [])
    except Exception:
        prev_total = 0
    if len(rows) < MIN_ROWS or (prev_total and len(rows) < prev_total * 0.8):
        log(f"！数据异常：本次 {len(rows)} 条（上次 {prev_total} 条，下限 {MIN_ROWS}），"
            f"放弃写入，保留上一次数据")
        save_bad(f"rows too few: {len(rows)} vs {prev_total}", "")
        sys.exit(3)

    stats = build_stats(rows)
    updated = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    data = {
        "updated_at": updated,
        "updated_date": updated[:10],
        "source": "天天基金移动端接口 fundmobapi.eastmoney.com",
        "universe": ("QDII 基金全集（FundType=6）" if uni_src == "live"
                     else "QDII 基金全集（复用上次代码表，实时抓取失败）"),
        "stats": stats,
        "items": rows,
    }

    json.dump(data, open(WEB_JSON, "w", encoding="utf-8"), ensure_ascii=False)
    log(f"看板数据已保存：{WEB_JSON}（{len(rows)} 条）")

    if not args.no_excel:
        try:
            import _build_excel
            _build_excel.main()
            log("Excel 排行榜已刷新")
        except Exception as e:
            log(f"Excel 生成失败（忽略）：{e}")

    if not args.no_static:
        p = render_static(data)
        log(f"静态快照已生成：{p}" if p else "静态快照跳过（模板缺少内嵌标记）")

    log(f"完成，用时 {time.time() - t0:.1f}s")
    log(f"  状态分布 {[(x['k'], x['v']) for x in stats['by_status']]}")
    log(f"  限大额未披露额度 {stats['undisclosed']} 只")


if __name__ == "__main__":
    main()
