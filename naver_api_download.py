"""
네이버 SA API 자동 다운로더
전략:
  쇼핑검색 → 기존 '광고 캠페인 리스트.csv' 에서 소재ID 로드 → /stats 직접 조회
  파워링크 → /ncc/adgroups 로 그룹 열거 → 그룹ID를 소재ID로 사용 → /stats 조회
  (/ncc/ads 엔드포인트 미사용)

생성 파일:
  raw/노출클릭_{n}월.csv
  raw/구매완료_{n}월.csv
실행: python naver_api_download.py
"""

import hashlib
import hmac
import base64
import time
import csv
import json
import requests
import pandas as pd
from pathlib import Path
from datetime import date
import calendar
from datetime import timedelta

BASE_URL  = "https://api.naver.com"
ENV_FILE  = Path(__file__).parent / ".env"
RAW_DIR   = Path(__file__).parent / "raw"
REF_DIR   = Path(__file__).parent / "기준데이터"

STATS_FIELDS = '["impCnt","clkCnt","salesAmt","purchaseCcnt","purchaseConvAmt"]'
BATCH_SIZE   = 100


# ── 인증 ──────────────────────────────────────────────────────────

def load_env():
    env = {}
    for line in ENV_FILE.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if "=" in line and not line.startswith("#"):
            k, v = line.split("=", 1)
            env[k.strip()] = v.strip()
    return env


def _sign(ts, method, uri, secret_key):
    msg = "{}.{}.{}".format(ts, method, uri)
    h = hmac.new(bytes(secret_key, "utf-8"), bytes(msg, "utf-8"), hashlib.sha256)
    return base64.b64encode(h.digest()).decode("utf-8")


def _headers(method, uri, api_key, secret_key, customer_id):
    ts = str(round(time.time() * 1000))
    return {
        "Content-Type": "application/json; charset=UTF-8",
        "X-Timestamp" : ts,
        "X-API-KEY"   : api_key,
        "X-Customer"  : str(customer_id),
        "X-Signature" : _sign(ts, method, uri, secret_key),
    }


def api_get(path, creds, params=None):
    api_key, secret_key, customer_id = creds
    try:
        r = requests.get(
            BASE_URL + path,
            headers=_headers("GET", path, api_key, secret_key, customer_id),
            params=params,
            timeout=30,
        )
        return r
    except Exception as e:
        print(f"  [요청 오류] {path}: {e}")
        return None


# ── API 헬퍼 ─────────────────────────────────────────────────────

def fetch_list(path, creds, params):
    r = api_get(path, creds, {**params, "size": 1000})
    if r is None or not r.ok or not r.text.strip():
        return []
    try:
        data = r.json()
    except Exception:
        return []
    if isinstance(data, list):
        return data
    if isinstance(data, dict):
        return data.get("items", [])
    return []


def fetch_stats(ids, time_range, creds):
    if not ids:
        return []
    r = api_get("/stats", creds, {
        "ids"       : ",".join(ids),
        "fields"    : STATS_FIELDS,
        "timeUnit"  : "MONTHLY",
        "timeRange" : time_range,
    })
    if r is None or not r.ok or not r.text.strip():
        return []
    try:
        data = r.json()
    except Exception:
        return []
    if isinstance(data, list):
        return data
    if isinstance(data, dict):
        return data.get("data", data.get("items", []))
    return []


# ── CSV 저장 ──────────────────────────────────────────────────────

def write_csv(path, date_header, columns, rows):
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        f.write(date_header + "\n")
        writer = csv.DictWriter(f, fieldnames=columns)
        writer.writeheader()
        writer.writerows(rows)


# ── 소재 메타 구성 ────────────────────────────────────────────────

def load_ss_meta():
    """쇼핑검색: 광고 캠페인 리스트.csv → {소재ID: {campaign, adgroup}}"""
    paths = list(REF_DIR.glob("광고 캠페인 리스트*.csv"))
    # 파워링크 파일 제외, 쇼핑검색 파일만
    path = next((p for p in paths if "파워링크" not in p.stem), None)
    if not path:
        print("[ERROR] 기준데이터/광고 캠페인 리스트.csv 가 없습니다.")
        return {}

    df = pd.read_csv(path, encoding="utf-8-sig")
    df.columns = df.columns.str.strip()

    meta = {}
    for _, row in df.iterrows():
        sojae_id = str(row.get("소재 ID", "")).strip()
        if sojae_id and sojae_id != "nan":
            meta[sojae_id] = {
                "campaign" : str(row.get("캠페인 이름", "")).strip(),
                "adgroup"  : str(row.get("광고그룹 이름", "")).strip(),
            }
    return meta


def load_pl_meta(creds):
    """파워링크: API로 WEB_SITE 캠페인의 광고그룹 열거 → {adgroup_id: {campaign, adgroup}}
    /ncc/adgroups 는 campaignIds 파라미터를 무시하고 계정 전체 adgroup을 반환하므로
    반환된 adgroup의 nccCampaignId 를 직접 확인하여 WEB_SITE 소속만 필터링한다.
    """
    all_camps = fetch_list("/ncc/campaigns", creds, {})
    web_camps = [c for c in all_camps if c.get("campaignTp") == "WEB_SITE"]
    web_camp_ids = {c.get("nccCampaignId"): c.get("name", "") for c in web_camps}
    print(f"  파워링크 캠페인: {len(web_camps)}개")

    # 한 번만 조회 후 nccCampaignId 로 실제 파워링크 adgroup만 필터
    all_ags = fetch_list("/ncc/adgroups", creds, {})
    meta = {}
    for ag in all_ags:
        ag_camp_id = ag.get("nccCampaignId", "")
        if ag_camp_id not in web_camp_ids:
            continue
        ag_id   = ag.get("nccAdgroupId", "")
        ag_name = ag.get("name", "")
        if ag_id:
            meta[ag_id] = {
                "campaign" : web_camp_ids[ag_camp_id],
                "adgroup"  : ag_name,
            }

    return meta


# ── 통계 배치 조회 ────────────────────────────────────────────────

def collect_stats(meta, time_range, creds, label):
    ids       = list(meta.keys())
    all_stats = []
    total     = len(ids)

    if total == 0:
        return []

    print(f"  {label} 통계 조회: {total}개 ID → {-(-total // BATCH_SIZE)}회 배치")

    # 첫 번째 배치 응답 확인용 출력
    first_batch = ids[:min(BATCH_SIZE, total)]
    r = api_get("/stats", creds, {
        "ids"      : ",".join(first_batch),
        "fields"   : STATS_FIELDS,
        "timeUnit" : "MONTHLY",
        "timeRange": time_range,
    })
    if r is None:
        print(f"  [{label}] 첫 배치 요청 실패 (None)")
    else:
        print(f"  [{label}] 첫 배치 HTTP: {r.status_code} | body: {r.text[:200]}")
        try:
            data  = r.json()
            if isinstance(data, list):
                items = data
            elif isinstance(data, dict):
                items = data.get("data", data.get("items", []))
            else:
                items = []
            all_stats.extend(items)
        except Exception as e:
            print(f"  [{label}] JSON 파싱 오류: {e}")

    time.sleep(0.1)

    for i in range(BATCH_SIZE, total, BATCH_SIZE):
        batch = ids[i : i + BATCH_SIZE]
        stats = fetch_stats(batch, time_range, creds)
        all_stats.extend(stats)
        time.sleep(0.1)

    print(f"  → 수신 레코드: {len(all_stats)}건")
    return all_stats


# ── 메인 ──────────────────────────────────────────────────────────

def run():
    env   = load_env()
    creds = (env["SA_ACCESS_LICENSE"], env["SA_SECRET_KEY"], env["SA_CUSTOMER_ID"])

    today = date.today()

    yesterday = today - timedelta(days=1)

    # 매월 1일에 실행하면 이전 달(완전한 데이터)을 조회
    # 그 외에는 이번 달 1일~어제까지 조회 (당일 실시간 데이터 제외)
    if today.day == 1:
        if today.month == 1:
            target = date(today.year - 1, 12, 1)
        else:
            target = date(today.year, today.month - 1, 1)
        until_date = date(target.year, target.month, calendar.monthrange(target.year, target.month)[1])
        print(f"  (오늘 {today.month}월 1일 → 완전한 이전 달 {target.month}월 데이터 조회)")
    else:
        target     = today
        until_date = yesterday  # 당일 실시간 제외

    last_day = calendar.monthrange(target.year, target.month)[1]
    month_ko = f"{target.month}월"
    date_hdr = (
        f"조회기간: {target.year}.{target.month:02d}.01."
        f"~{until_date.year}.{until_date.month:02d}.{until_date.day:02d}."
    )
    # /stats API 는 timeRange (JSON) 필요
    time_range = json.dumps({
        "since": f"{target.year}-{target.month:02d}-01",
        "until": f"{until_date.year}-{until_date.month:02d}-{until_date.day:02d}",
    })

    RAW_DIR.mkdir(exist_ok=True)

    # ── 1. 쇼핑검색 소재 메타 (캠페인 리스트 파일) ──────────────
    print("[1] 쇼핑검색 소재 목록 로드 중...")
    ss_meta = load_ss_meta()
    print(f"  소재 {len(ss_meta)}개 로드 완료")

    if not ss_meta:
        print("  [주의] 쇼핑검색 소재 없음 — 기준데이터 폴더를 확인하세요.")

    # ── 2. 파워링크 광고그룹 메타 (API 조회) ────────────────────
    print("\n[2] 파워링크 광고그룹 조회 중...")
    pl_meta = load_pl_meta(creds)
    print(f"  광고그룹 {len(pl_meta)}개 수집 완료")

    # ── 3. 통계 조회 ─────────────────────────────────────────────
    print(f"\n[3] 통계 조회 중 ({today.year}-{today.month:02d})...")
    ss_stats = collect_stats(ss_meta, time_range, creds, "쇼핑검색")
    pl_stats = collect_stats(pl_meta, time_range, creds, "파워링크")

    # ── 4. CSV 행 구성 ────────────────────────────────────────────
    click_rows = []
    conv_rows  = []

    def add_rows(stats, meta, type_ko):
        for stat in stats:
            entity_id = stat.get("id")
            m = meta.get(entity_id, {})
            imp  = int(stat.get("impCnt",   0) or 0)
            clk  = int(stat.get("clkCnt",   0) or 0)
            cost = int(stat.get("salesAmt", 0) or 0)
            ccnt = int(stat.get("purchaseCcnt",    0) or 0)
            camt = int(stat.get("purchaseConvAmt", 0) or 0)

            click_rows.append({
                "소재"              : entity_id,
                "캠페인유형"         : type_ko,
                "캠페인"            : m.get("campaign", ""),
                "광고그룹"          : m.get("adgroup",  ""),
                "노출수"            : imp,
                "클릭수"            : clk,
                "총비용(VAT포함,원)" : cost,
            })
            conv_rows.append({
                "소재"                  : entity_id,
                "캠페인유형"             : type_ko,
                "구매완료 전환수"         : ccnt,
                "구매완료 전환매출액(원)" : camt,
            })

    add_rows(ss_stats, ss_meta, "쇼핑검색")
    add_rows(pl_stats, pl_meta, "파워링크")

    # ── 5. 저장 ───────────────────────────────────────────────────
    click_path = RAW_DIR / f"노출클릭_{month_ko}.csv"
    conv_path  = RAW_DIR / f"구매완료_{month_ko}.csv"

    click_cols = ["소재", "캠페인유형", "캠페인", "광고그룹",
                  "노출수", "클릭수", "총비용(VAT포함,원)"]
    conv_cols  = ["소재", "캠페인유형", "구매완료 전환수", "구매완료 전환매출액(원)"]

    write_csv(click_path, date_hdr, click_cols, click_rows)
    write_csv(conv_path,  date_hdr, conv_cols,  conv_rows)

    print(f"\n[완료] 저장 완료!")
    print(f"  raw/노출클릭_{month_ko}.csv  ({len(click_rows)}행)")
    print(f"  raw/구매완료_{month_ko}.csv  ({len(conv_rows)}행)")
    print()
    print("다음: python naver_ad_pipeline.py")


if __name__ == "__main__":
    run()
