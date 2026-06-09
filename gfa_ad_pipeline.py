"""
네이버 GFA 광고 데이터 파이프라인 (애드부스트 + DA/카탈로그 소재보고서)
==========================================================================
폴더 구조:
  naver_ad_tool/
  ├── raw/
  │   └── GFA/
  │       ├── 애드부스트_N월.csv    ← 매달 추가
  │       └── 소재보고서_N월.csv    ← 매달 추가
  ├── 기준데이터/
  │   └── 스토어 등록 상품 리스트.csv
  ├── output/
  │   └── gfa_data.csv
  └── gfa_ad_pipeline.py
"""

import pandas as pd
import re
import sys
from pathlib import Path


# ============================================================
# 경로 설정
# ============================================================
BASE_DIR = Path(__file__).parent
GFA_DIR  = BASE_DIR / "raw" / "GFA"
REF_DIR  = BASE_DIR / "기준데이터"
OUT_DIR  = BASE_DIR / "output"


# ============================================================
# 애드부스트 캠페인명 → 카테고리
# 새 캠페인 추가 시 여기에 넣으세요 (순서대로 먼저 매칭됨)
# ============================================================
ADV_CAMP_KEYWORDS = [
    ("시스템에어컨", "시스템에어컨"),  # 에어컨보다 먼저 체크
    ("에어컨",      "에어컨"),
    ("TV",          "TV"),
    ("공기청정기",   "공기청정기"),
    ("청소기",       "청소기"),
    ("제습기",       "제습기"),
    ("SD",           "SD"),
]

# 전상품 캠페인 fallback: 상품명 키워드 → 카테고리
# 순서 중요 — 시스템에어컨이 에어컨보다 먼저 매칭되어야 함
NAME_KEYWORDS = [
    (["시스템에어컨"],                                           "시스템에어컨"),
    (["무풍에어컨", "멀티형에어컨"],                              "에어컨"),
    (["QLED", "OLED", "Neo QLED", "더프레임", "Crystal UHD", "UHD TV", " TV "], "TV"),
    (["큐브에어", "블루스카이", "공기청정기"],                     "공기청정기"),
    (["비스포크 제트", "파워건", "청소기"],                        "청소기"),
    (["무풍제습기", "제습기"],                                    "제습기"),
    (["사운드바", "블루투스스피커"],                               "SD"),
]


# ============================================================
# 유틸
# ============================================================

def find_gfa_files(keyword):
    if not GFA_DIR.exists():
        return []
    return sorted(GFA_DIR.glob(f"*{keyword}*.csv"))


def read_csv_auto(path):
    for enc in ["utf-8-sig", "cp949", "euc-kr"]:
        try:
            return pd.read_csv(path, encoding=enc)
        except (UnicodeDecodeError, UnicodeError):
            continue
    raise ValueError(f"인코딩 감지 실패: {path.name}")


def load_store_list():
    candidates = list(REF_DIR.glob("*상품*리스트*")) + list(REF_DIR.glob("*스토어*등록*"))
    candidates = [f for f in candidates if not f.name.startswith("~")]
    if not candidates:
        print("  [WARNING] 스토어 등록 상품 리스트 파일을 찾지 못했습니다. 모델명은 (미매칭) 처리됩니다.")
        return pd.DataFrame(columns=["상품ID", "model"])

    path = candidates[0]
    df = read_csv_auto(path) if path.suffix == ".csv" else pd.read_excel(path)
    df.columns = df.columns.str.strip()
    df = df.rename(columns={
        "상품번호(스마트스토어)": "상품ID",
        "판매자상품코드"        : "model",
    })
    df["상품ID"] = pd.to_numeric(
        df["상품ID"].astype(str).str.replace(",", ""), errors="coerce"
    ).astype("Int64")
    return df[["상품ID", "model"]].dropna(subset=["상품ID"]).drop_duplicates(subset=["상품ID"])


def parse_date(val):
    """기간 컬럼 값 → YYYY-MM. 예: '2026.05.' → '2026-05'"""
    m = re.search(r"(\d{4})\.(\d{2})\.", str(val))
    if m:
        return f"{m.group(1)}-{m.group(2)}"
    return str(val).strip()


def numeric(s):
    return pd.to_numeric(s.astype(str).str.replace(",", ""), errors="coerce").fillna(0)


def camp_to_category(campaign):
    """캠페인명 → 카테고리. 매칭 실패 시 None"""
    if pd.isna(campaign):
        return None
    for keyword, cat in ADV_CAMP_KEYWORDS:
        if keyword in str(campaign):
            return cat
    return None


def name_to_category(product_name):
    """상품명 키워드 → 카테고리 (전상품 캠페인 fallback)"""
    if pd.isna(product_name):
        return "그 외 상품"
    name = str(product_name)
    for keywords, cat in NAME_KEYWORDS:
        if any(kw in name for kw in keywords):
            return cat
    return "그 외 상품"


# ============================================================
# 애드부스트 처리
# ============================================================

def process_adv(files, store_df):
    frames = []
    for f in files:
        df = read_csv_auto(f)
        df.columns = df.columns.str.strip()

        df = df.rename(columns={
            "상품명"             : "name",
            "쇼핑몰 상품 ID"    : "상품ID",
            "캠페인 이름"        : "campaign",
            "기간"               : "기간",
            "총비용"             : "spend",
            "노출수"             : "impressions",
            "클릭수"             : "clicks",
            "구매완료 수"        : "conv",
            "장바구니 담기 수"   : "cart",
            "구매완료 전환매출액" : "revenue",
            "장바구니 전환매출액" : "cart_revenue",
        })

        # 캠페인 이름 컬럼이 없는 구버전 파일 대응 (상품명 키워드 fallback으로 처리)
        if "campaign" not in df.columns:
            df["campaign"] = None

        for col in ["spend", "impressions", "clicks", "conv", "cart", "revenue", "cart_revenue"]:
            if col in df.columns:
                df[col] = numeric(df[col])

        df["date"] = df["기간"].apply(parse_date)

        # 상품ID → 모델명 join
        df["상품ID"] = pd.to_numeric(
            df["상품ID"].astype(str).str.replace(",", ""), errors="coerce"
        ).astype("Int64")
        df = df.merge(store_df, on="상품ID", how="left")
        df["model"] = df["model"].fillna("(미매칭)")

        # 카테고리: 캠페인명 우선, 실패 시 상품명 키워드 fallback
        df["category"] = df["campaign"].apply(camp_to_category)
        mask = df["category"].isna()
        df.loc[mask, "category"] = df.loc[mask, "name"].apply(name_to_category)
        df["category"] = df["category"].fillna("그 외 상품")

        frames.append(df)

    result = pd.concat(frames, ignore_index=True)

    # 월+모델+상품명+카테고리 단위로 집계
    result = (
        result.groupby(["date", "model", "name", "category"], dropna=False)
        .agg(
            spend        =("spend",        "sum"),
            impressions  =("impressions",  "sum"),
            clicks       =("clicks",       "sum"),
            conv         =("conv",         "sum"),
            revenue      =("revenue",      "sum"),
            cart         =("cart",         "sum"),
            cart_revenue =("cart_revenue", "sum"),
        )
        .reset_index()
    )
    result["subtype"] = "애드부스트"
    return result


# ============================================================
# 소재보고서 처리
# ============================================================

def process_soja(files):
    frames = []
    for f in files:
        df = read_csv_auto(f)
        df.columns = df.columns.str.strip()

        df = df.rename(columns={
            "캠페인 이름"        : "campaign",
            "기간"               : "기간",
            "총비용"             : "spend",
            "노출수"             : "impressions",
            "클릭수"             : "clicks",
            "구매완료 수"        : "conv",
            "장바구니 담기 수"   : "cart",
            "구매완료 전환매출액" : "revenue",
            "장바구니 전환매출액" : "cart_revenue",
        })

        for col in ["spend", "impressions", "clicks", "conv", "cart", "revenue", "cart_revenue"]:
            if col in df.columns:
                df[col] = numeric(df[col])

        df["date"] = df["기간"].apply(parse_date)

        # 캠페인명에 '카탈로그' 포함 여부로 DA/카탈로그 구분
        df["ad_type"] = df["campaign"].apply(
            lambda c: "카탈로그" if "카탈로그" in str(c) else "DA"
        )

        frames.append(df)

    result = pd.concat(frames, ignore_index=True)

    # 월+광고유형+캠페인 단위로 집계
    result = (
        result.groupby(["date", "ad_type", "campaign"], dropna=False)
        .agg(
            spend        =("spend",        "sum"),
            impressions  =("impressions",  "sum"),
            clicks       =("clicks",       "sum"),
            conv         =("conv",         "sum"),
            revenue      =("revenue",      "sum"),
            cart         =("cart",         "sum"),
            cart_revenue =("cart_revenue", "sum"),
        )
        .reset_index()
    )
    result["subtype"] = "소재보고서"
    return result


# ============================================================
# 파이프라인 실행
# ============================================================

def run_pipeline():
    OUT_DIR.mkdir(exist_ok=True)

    adv_files  = find_gfa_files("애드부스트")
    soja_files = find_gfa_files("소재보고서")

    if not adv_files and not soja_files:
        print("[INFO] raw/GFA/ 폴더에 GFA 파일이 없습니다. 건너뜁니다.")
        return None

    print("[GFA] 파일 감지")
    print(f"  애드부스트 : {', '.join(f.name for f in adv_files) or '없음'}")
    print(f"  소재보고서 : {', '.join(f.name for f in soja_files) or '없음'}")
    print()

    store_df = load_store_list()
    print(f"  상품리스트 : {len(store_df):,}건 로드됨")
    print()

    frames = []

    if adv_files:
        print("[애드부스트] 처리 중...")
        adv = process_adv(adv_files, store_df)
        matched = (adv["model"] != "(미매칭)").sum()
        print(f"  결과: {len(adv):,}행  |  모델 매칭: {matched}/{len(adv)}")
        print(f"  카테고리별: {adv['category'].value_counts().to_dict()}")
        frames.append(adv)

    if soja_files:
        print()
        print("[소재보고서] 처리 중...")
        soja = process_soja(soja_files)
        da_c  = (soja["ad_type"] == "DA").sum()
        cat_c = (soja["ad_type"] == "카탈로그").sum()
        print(f"  결과: {len(soja):,}행  |  DA: {da_c}개 캠페인, 카탈로그: {cat_c}개 캠페인")
        frames.append(soja)

    result = pd.concat(frames, ignore_index=True)
    out = OUT_DIR / "gfa_data.csv"
    result.to_csv(out, index=False, encoding="utf-8-sig")

    print()
    print(f"[완료] {out.name} 저장됨")
    print(f"   총 {len(result):,}행  /  {result['date'].nunique()}개월")
    return result


if __name__ == "__main__":
    run_pipeline()
