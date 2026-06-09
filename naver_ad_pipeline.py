"""
네이버 SA 광고 데이터 파이프라인 (쇼핑검색 + 파워링크)
========================================================
폴더 구조:
  naver_ad_tool/
  ├── raw/              ← 네이버에서 받은 raw 파일 여기에 넣기 (매달 교체)
  │   ├── 노출클릭_*.csv
  │   └── 구매완료_*.csv
  ├── 기준데이터/        ← 거의 바뀌지 않는 기준 파일 (변경 시만 교체)
  │   ├── 광고_캠페인_리스트.csv          ← 쇼핑검색용
  │   ├── 모델명 리스트.csv               ← 파이프라인이 자동 생성/누적
  │   └── 스토어_등록_상품_리스트.csv (.xlsx도 가능)
  ├── output/           ← 결과 파일 자동 저장
  │   └── sa_data.csv
  ├── naver_ad_pipeline.py
  └── 실행.bat
"""

import pandas as pd
import re
import sys
from pathlib import Path


# ============================================================
# ★ 카테고리 매핑 (세분류/소분류 → 8개 카테고리)
# 새 제품군이 생기면 여기에 추가하세요
# ============================================================
CATEGORY_MAP = {
    # TV
    "LEDTV"          : "TV",
    "QLEDTV"         : "TV",
    "OLEDTV"         : "TV",
    "TV"             : "TV",
    # 에어컨 (시스템에어컨은 별도)
    "창문형에어컨"    : "에어컨",
    "멀티형에어컨"    : "에어컨",
    "스탠드형에어컨"  : "에어컨",
    "벽걸이형에어컨"  : "에어컨",
    # 시스템에어컨
    "시스템에어컨"    : "시스템에어컨",
    # 공기청정기
    "공기청정기"      : "공기청정기",
    "공기정화기"      : "공기청정기",
    "공기정화기필터"  : "그 외 상품",
    # 청소기
    "핸디스틱청소기"  : "청소기",
    "진공청소기"      : "청소기",
    "로봇청소기"      : "청소기",
    "청소기"         : "청소기",
    # 제습기
    "일반용제습기"    : "제습기",
    # SD (사운드바/스피커)
    "사운드바시스템"  : "SD",
    "블루투스스피커"  : "SD",
}

# 캠페인명 → 카테고리 추출 규칙 (미매칭 소재 fallback용)
CAMP_EXTRACT_CATEGORY_RULES = [
    (["시스템에어컨"],                                            "시스템에어컨"),
    (["멀티형에어컨", "스탠드에어컨", "벽걸이에어컨", "창문형에어컨"], "에어컨"),
    (["TV"],                                                     "TV"),
    (["SD사운드바", "사운드바"],                                  "SD"),
    (["공기청정기"],                                              "공기청정기"),
    (["청소기"],                                                  "청소기"),
    (["제습기"],                                                  "제습기"),
]

# ★ 파워링크 캠페인명 → 카테고리 (직접 매핑)
# 새 캠페인이 생기면 여기에 추가하세요
PL_CAMPAIGN_CATEGORY = {
    "파워_TV"              : "TV",
    "파워_에어컨"           : "에어컨",
    "파워_시스템에어컨"      : "시스템에어컨",
    "파워_시스템에어컨_26년" : "시스템에어컨",
    "파워_공기청정기"        : "공기청정기",
    "파워_청소기"           : "청소기",
    "파워_제습기"           : "제습기",
    "파워_SD"              : "SD",
    "파워_공식몰"           : "그 외 상품",
}


# ============================================================
# 파일 자동 감지
# ============================================================

BASE_DIR        = Path(__file__).parent
RAW_DIR         = BASE_DIR / "raw"
REF_DIR         = BASE_DIR / "기준데이터"
OUT_DIR         = BASE_DIR / "output"
MODEL_LIST_PATH = REF_DIR / "모델명 리스트.csv"
MODEL_LIST_COLS = ["자동추출_모델명", "상품ID", "캠페인명", "자동추출_카테고리",
                   "수동 적용_모델명", "수동 적용_카테고리"]

def find_files(directory: Path, keywords: list, extensions=("*.csv", "*.xlsx")) -> list:
    """키워드를 모두 포함하는 파일을 디렉토리에서 모두 탐색"""
    candidates = []
    for ext in extensions:
        candidates.extend(directory.glob(ext))
    result = []
    for path in candidates:
        name = path.stem
        if all(kw in name for kw in keywords):
            result.append(path)
    return sorted(result)


# ============================================================
# 로더
# ============================================================

def read_file(path: Path) -> pd.DataFrame:
    if str(path).endswith(".csv"):
        return pd.read_csv(path, encoding="utf-8-sig")
    else:
        return pd.read_excel(path)


def extract_date_from_header(path: Path) -> str:
    """헤더 첫 줄에서 날짜 추출. 예: '2026.03.01.~2026.03.18.' → '2026-03'"""
    with open(path, encoding="utf-8-sig") as f:
        header = f.readline()
    m = re.search(r"(\d{4})\.(\d{2})\.\d{2}\.", header)
    if m:
        return f"{m.group(1)}-{m.group(2)}"
    return None


def load_click_data(path: Path) -> pd.DataFrame:
    """노출클릭 CSV 로드 — 쇼핑검색+파워링크 모두 반환, 캠페인명/광고그룹명 포함"""
    raw = pd.read_csv(path, skiprows=1, encoding="utf-8-sig")
    raw.columns = raw.columns.str.strip()

    if "총비용(VAT포함,원)" in raw.columns:
        raw = raw.rename(columns={"총비용(VAT포함,원)": "총비용"})

    rename_map = {"소재": "소재ID", "캠페인": "캠페인명", "광고그룹": "광고그룹명",
                  "노출수": "impressions", "클릭수": "clicks", "총비용": "spend"}

    if "월별" in raw.columns:
        rename_map["월별"] = "date"
        df = raw.rename(columns=rename_map)[
            ["date", "소재ID", "캠페인유형", "캠페인명", "광고그룹명", "impressions", "clicks", "spend"]
        ].copy()
        df["date"] = df["date"].str.replace(".", "-", regex=False).str.rstrip("-")
    else:
        date_str = extract_date_from_header(path)
        df = raw.rename(columns=rename_map)[
            ["소재ID", "캠페인유형", "캠페인명", "광고그룹명", "impressions", "clicks", "spend"]
        ].copy()
        df.insert(0, "date", date_str)

    df = df[df["캠페인유형"].isin(["쇼핑검색", "파워링크"])].copy()

    # 삭제된 소재/그룹은 ID 뒤에 '(삭제)' 텍스트가 붙으므로 제거
    df["소재ID"] = df["소재ID"].astype(str).str.replace(r"\(삭제\)", "", regex=True).str.strip()

    for col in ["impressions", "clicks", "spend"]:
        df[col] = pd.to_numeric(df[col].astype(str).str.replace(",", ""), errors="coerce")
    return df


def load_conv_data(path: Path) -> pd.DataFrame:
    """전환유형 CSV 로드 — 쇼핑검색+파워링크 모두 반환, 캠페인유형 컬럼 포함"""
    raw = pd.read_csv(path, skiprows=1, encoding="utf-8-sig")
    raw.columns = raw.columns.str.strip()

    date_str = extract_date_from_header(path)

    df = raw.rename(columns={
        "소재"                  : "소재ID",
        "구매완료 전환수"         : "conv",
        "구매완료 전환매출액(원)"  : "revenue",
    })[["소재ID", "캠페인유형", "conv", "revenue"]].copy()
    df.insert(0, "date", date_str)

    # 쇼핑검색 + 파워링크만 유지
    df = df[df["캠페인유형"].isin(["쇼핑검색", "파워링크"])].copy()

    # 삭제된 소재/그룹은 ID 뒤에 '(삭제)' 텍스트가 붙으므로 제거
    df["소재ID"] = df["소재ID"].astype(str).str.replace(r"\(삭제\)", "", regex=True).str.strip()

    for col in ["conv", "revenue"]:
        df[col] = pd.to_numeric(df[col].astype(str).str.replace(",", ""), errors="coerce")
    return df


def load_campaign_list(path: Path) -> pd.DataFrame:
    """쇼핑검색 캠페인 리스트: 소재ID → 상품ID 매핑"""
    df = read_file(path)
    df.columns = df.columns.str.strip()
    df = df.rename(columns={
        "소재 ID"     : "소재ID",
        "쇼핑몰 상품ID": "상품ID",
        "캠페인 이름"  : "캠페인명",
        "카테고리"     : "캠페인_카테고리",
    })
    df["상품ID"] = pd.to_numeric(df["상품ID"], errors="coerce").astype("Int64")
    return df[["소재ID", "상품ID", "캠페인명", "캠페인_카테고리"]].drop_duplicates(subset=["소재ID"])


def load_store_list(path: Path) -> pd.DataFrame:
    df = read_file(path)
    df.columns = df.columns.str.strip()
    df = df.rename(columns={
        "상품번호(스마트스토어)": "상품ID",
        "판매자상품코드"        : "model_raw",
    })
    df["상품ID"] = pd.to_numeric(df["상품ID"], errors="coerce").astype("Int64")
    df["분류키"] = df["세분류"].fillna(df["소분류"])
    return df[["상품ID", "model_raw", "분류키"]].drop_duplicates(subset=["상품ID"])


def build_model_category_lookup(store_df: pd.DataFrame) -> dict:
    return (
        store_df[store_df["분류키"].notna()]
        .drop_duplicates(subset=["model_raw"])
        .set_index("model_raw")["분류키"]
        .to_dict()
    )


def apply_category_map(df: pd.DataFrame) -> pd.DataFrame:
    df["category"] = df["분류키"].map(CATEGORY_MAP).fillna("그 외 상품")
    return df


def extract_model_from_campaign(campaign_name):
    """캠페인명에서 모델명·카테고리 추출.
    예) '🔴TV65인치_KQ65SF8EAEXKR' → ('KQ65SF8EAEXKR', 'TV')
    추출 불가 시 (None, 카테고리) 반환.
    """
    name = re.sub(r"\(삭제\)", "", str(campaign_name)).strip()

    category = "그 외 상품"
    for keywords, cat in CAMP_EXTRACT_CATEGORY_RULES:
        if any(kw in name for kw in keywords):
            category = cat
            break

    parts = name.split("_")
    if len(parts) >= 2:
        candidate = parts[-1].strip()
        if re.match(r"^[A-Za-z0-9]", candidate):
            return candidate, category
    return None, category


def load_model_list() -> pd.DataFrame:
    """모델명 리스트 로드. 없으면 빈 DataFrame 반환."""
    if MODEL_LIST_PATH.exists():
        df = pd.read_csv(MODEL_LIST_PATH, encoding="utf-8-sig")
        df.columns = df.columns.str.strip()
        for col in MODEL_LIST_COLS:
            if col not in df.columns:
                df[col] = ""
        return df[MODEL_LIST_COLS].fillna("")
    return pd.DataFrame(columns=MODEL_LIST_COLS)


def update_model_list(current_models: pd.DataFrame, existing: pd.DataFrame) -> pd.DataFrame:
    """기존 리스트에 새 모델 행 추가. 수동 적용 값은 보존. 저장 후 반환."""
    new_rows = current_models[
        ~current_models["자동추출_모델명"].isin(existing["자동추출_모델명"])
    ].copy()
    new_rows["수동 적용_모델명"]   = ""
    new_rows["수동 적용_카테고리"] = ""

    updated = pd.concat([existing, new_rows[MODEL_LIST_COLS]], ignore_index=True)
    updated = updated.drop_duplicates(subset=["자동추출_모델명"], keep="first")

    # 기존 행의 상품ID 공란 → 현재 데이터에서 채우기
    id_map = (
        current_models.dropna(subset=["상품ID"])
        .set_index("자동추출_모델명")["상품ID"]
    )
    blank = updated["상품ID"].astype(str).str.strip().isin(["", "nan", "<NA>"])
    updated.loc[blank, "상품ID"] = updated.loc[blank, "자동추출_모델명"].map(id_map)

    updated.to_csv(MODEL_LIST_PATH, index=False, encoding="utf-8-sig")
    return updated


# ============================================================
# 파이프라인 실행
# ============================================================

def run_pipeline(
    file_click=None, file_conv=None,
    file_campaign=None, file_store=None,
    output_file=None,
):
    OUT_DIR.mkdir(exist_ok=True)

    # ── 파일 자동 감지 ──
    fc_list = [Path(file_click)]   if file_click    else find_files(RAW_DIR, ["노출클릭"])
    fv_list = [Path(file_conv)]    if file_conv     else find_files(RAW_DIR, ["구매완료"])

    # 쇼핑검색 캠페인 리스트 (파워링크 파일 있어도 무시)
    all_camp = find_files(REF_DIR, ["캠페인", "리스트"])
    fa  = Path(file_campaign) if file_campaign else next((f for f in all_camp if "파워링크" not in f.stem), None)
    fs  = Path(file_store)    if file_store    else (find_files(REF_DIR, ["상품", "리스트"]) or [None])[0]
    out = Path(output_file)   if output_file   else OUT_DIR / "sa_data.csv"

    # ── 필수 파일 존재 확인 ──
    missing = []
    if not fc_list:                    missing.append(("노출클릭 파일",          "raw/"))
    if not fv_list:                    missing.append(("구매완료 파일",          "raw/"))
    if fa is None or not fa.exists():  missing.append(("캠페인 리스트 (쇼핑검색)", "기준데이터/"))
    if fs is None or not fs.exists():  missing.append(("상품 리스트",            "기준데이터/"))

    if missing:
        print("❌ 아래 파일을 찾지 못했습니다:")
        for name, loc in missing:
            print(f"  [{name}] → {loc} 폴더 확인 필요")
        input("\nEnter 키를 눌러 종료...")
        sys.exit(1)

    print("▶ 파일 감지 완료")
    print(f"  노출클릭  : {', '.join(f.name for f in fc_list)}")
    print(f"  구매완료  : {', '.join(f.name for f in fv_list)}")
    print(f"  캠페인    : {fa.name}")
    print(f"  상품리스트: {fs.name}")
    print(f"  ※ 파워링크는 raw 보고서의 캠페인/광고그룹 컬럼을 직접 사용")
    print()

    # ── 로드 ──
    print("▶ 데이터 로드 중...")
    all_click = pd.concat([load_click_data(f) for f in fc_list], ignore_index=True)
    all_conv  = pd.concat([load_conv_data(f)  for f in fv_list], ignore_index=True)

    # 광고 유형별 분리
    ss_clicks = all_click[all_click["캠페인유형"] == "쇼핑검색"].drop(columns=["캠페인유형"]).copy()
    ss_convs  = all_conv[all_conv["캠페인유형"]  == "쇼핑검색"].drop(columns=["캠페인유형"]).copy()
    pl_clicks = all_click[all_click["캠페인유형"] == "파워링크"].drop(columns=["캠페인유형"]).copy()
    pl_convs  = all_conv[all_conv["캠페인유형"]  == "파워링크"].drop(columns=["캠페인유형"]).copy()

    print(f"  쇼핑검색 노출클릭: {len(ss_clicks):,}행")
    print(f"  파워링크 노출클릭: {len(pl_clicks):,}행")

    campaign_df      = load_campaign_list(fa)
    store_df         = load_store_list(fs)
    model_cat_lookup = build_model_category_lookup(store_df)

    # ══════════════════════════════════════════
    # [쇼핑검색] 기존 처리 흐름
    # ══════════════════════════════════════════
    # ── 모델명 리스트 로드 ──
    model_list   = load_model_list()
    _mask_m        = model_list["수동 적용_모델명"].astype(str).str.strip() != ""
    override_model = model_list[_mask_m].set_index("자동추출_모델명")["수동 적용_모델명"].to_dict()
    _mask_c        = model_list["수동 적용_카테고리"].astype(str).str.strip() != ""
    override_cat   = model_list[_mask_c].set_index("자동추출_모델명")["수동 적용_카테고리"].to_dict()
    is_new_list = not MODEL_LIST_PATH.exists()
    print(f"▶ 모델명 리스트: {'신규 생성 예정' if is_new_list else f'{len(model_list)}개 로드 (수동 적용 {len(override_model)}건)'}")

    print("\n▶ [쇼핑검색] STEP 1: 노출클릭 + 전환유형 병합...")
    ss = ss_clicks.merge(ss_convs, on=["date", "소재ID"], how="left")
    ss["conv"]    = ss["conv"].fillna(0).astype(int)
    ss["revenue"] = ss["revenue"].fillna(0).astype(int)

    print("▶ [쇼핑검색] STEP 2: 소재ID → 상품번호 매칭...")
    ss["원본캠페인명"] = ss["캠페인명"]   # 미매칭 소재의 캠페인명 추출용으로 보존
    ss = ss.drop(columns=["캠페인명", "광고그룹명"], errors="ignore")
    ss = ss.merge(campaign_df, on="소재ID", how="left")
    print(f"  매칭 성공: {ss['상품ID'].notna().sum():,}/{len(ss):,}행")

    print("▶ [쇼핑검색] STEP 3: 상품번호 → 모델명/카테고리 매칭...")
    ss = ss.merge(store_df, on="상품ID", how="left")
    print(f"  매칭 성공: {ss['model_raw'].notna().sum():,}/{len(ss):,}행")

    no_cat_mask = ss["분류키"].isna() & ss["model_raw"].notna()
    if no_cat_mask.any():
        ss.loc[no_cat_mask, "분류키"] = ss.loc[no_cat_mask, "model_raw"].map(model_cat_lookup)

    ss = apply_category_map(ss)

    fallback_mask = (ss["model_raw"].isna() | ss["분류키"].isna()) & ss["캠페인_카테고리"].notna()
    if fallback_mask.any():
        fallback_cat = ss.loc[fallback_mask, "캠페인_카테고리"].str.split("/").str[-1]
        ss.loc[fallback_mask, "category"] = fallback_cat.map(CATEGORY_MAP).fillna("그 외 상품")

    # campaign_df에도 없는 소재 → 원본 캠페인명에서 카테고리 추출
    raw_cat_mask = ss["model_raw"].isna() & ss["캠페인_카테고리"].isna() & ss["원본캠페인명"].notna()
    if raw_cat_mask.any():
        ss.loc[raw_cat_mask, "category"] = ss.loc[raw_cat_mask, "원본캠페인명"].apply(
            lambda c: extract_model_from_campaign(c)[1]
        )

    def resolve_model(row):
        # 1순위: 스토어 상품 리스트 매칭
        if pd.notna(row["model_raw"]):
            return row["model_raw"]
        # 2순위: campaign_df 캠페인명에서 모델 추출
        camp = row.get("캠페인명")
        if pd.notna(camp) and str(camp).strip():
            extracted, _ = extract_model_from_campaign(str(camp))
            if extracted:
                return extracted
            return f"(미매칭)_{camp}"
        # 3순위: 원본(raw) 캠페인명에서 모델 추출
        orig = row.get("원본캠페인명")
        if pd.notna(orig) and str(orig).strip():
            extracted, _ = extract_model_from_campaign(str(orig))
            if extracted:
                return extracted
            return f"(미매칭)_{orig}"
        return "(미매칭)"

    ss["model_auto"]    = ss.apply(resolve_model, axis=1)
    ss["category_auto"] = ss["category"].copy()

    # 수동 적용 override
    ss["model"]    = ss["model_auto"].map(lambda m: override_model.get(m, m))
    ss["category"] = ss["model_auto"].map(override_cat).fillna(ss["category_auto"])

    print("▶ [쇼핑검색] STEP 4: 집계...")
    ss_result = (
        ss.groupby(["date", "category", "model"], dropna=False)
        .agg(spend=("spend","sum"), clicks=("clicks","sum"),
             impressions=("impressions","sum"), conv=("conv","sum"),
             revenue=("revenue","sum"))
        .reset_index()
        .sort_values(["date", "category", "model"])
    )
    ss_result["ad_type"] = "쇼핑검색"
    print(f"  쇼핑검색 결과: {len(ss_result):,}행")

    # 모델명 리스트용 정보 수집 (쇼핑검색)
    ss_model_info = (
        ss[["model_auto", "상품ID", "원본캠페인명", "category_auto"]]
        .rename(columns={
            "model_auto"    : "자동추출_모델명",
            "원본캠페인명"   : "캠페인명",
            "category_auto" : "자동추출_카테고리",
        })
        .drop_duplicates(subset=["자동추출_모델명"])
    )
    ss_model_info["캠페인명"] = (
        ss_model_info["캠페인명"].fillna("확인 불가").replace("", "확인 불가")
    )

    # ══════════════════════════════════════════
    # [파워링크] 새로운 처리 흐름
    # ══════════════════════════════════════════
    pl_result     = pd.DataFrame()
    pl_model_info = pd.DataFrame(columns=["자동추출_모델명", "캠페인명", "자동추출_카테고리"])

    if not pl_clicks.empty:
        print("\n▶ [파워링크] STEP 1: 노출클릭 + 전환유형 병합...")
        pl = pl_clicks.merge(pl_convs, on=["date", "소재ID"], how="left")
        pl["conv"]    = pl["conv"].fillna(0).astype(int)
        pl["revenue"] = pl["revenue"].fillna(0).astype(int)

        print("▶ [파워링크] STEP 2: 캠페인명 → 카테고리, 광고그룹명 → 모델명...")
        cat_mapped = pl["캠페인명"].map(PL_CAMPAIGN_CATEGORY)
        pl["category"] = cat_mapped.fillna("그 외 상품")

        mapped = cat_mapped.notna()
        pl.loc[mapped,  "model"] = "<파워>" + pl.loc[mapped,  "광고그룹명"].fillna("(미매칭)")
        pl.loc[~mapped, "model"] = "<파워>" + pl.loc[~mapped, "캠페인명"].fillna("(미매칭)")
        print(f"  카테고리 매핑: {mapped.sum():,}/{len(pl):,}행 (나머지 → '그 외 상품')")

        print("▶ [파워링크] STEP 3: 집계...")
        pl_result = (
            pl.groupby(["date", "category", "model"], dropna=False)
            .agg(spend=("spend","sum"), clicks=("clicks","sum"),
                 impressions=("impressions","sum"), conv=("conv","sum"),
                 revenue=("revenue","sum"))
            .reset_index()
            .sort_values(["date", "category", "model"])
        )
        pl_result["ad_type"] = "파워링크"
        print(f"  파워링크 결과: {len(pl_result):,}행")

        # 모델명 리스트용 정보 수집 (파워링크)
        pl_model_info = (
            pl[["model", "캠페인명", "category"]]
            .rename(columns={
                "model"    : "자동추출_모델명",
                "category" : "자동추출_카테고리",
            })
            .drop_duplicates(subset=["자동추출_모델명"])
        )
        pl_model_info["캠페인명"] = pl_model_info["캠페인명"].fillna("확인 불가")

    # ══════════════════════════════════════════
    # 모델명 리스트 업데이트 & 저장
    # ══════════════════════════════════════════
    all_model_info = pd.concat(
        [ss_model_info] + ([pl_model_info] if not pl_result.empty else []),
        ignore_index=True
    ).drop_duplicates(subset=["자동추출_모델명"])

    updated_list = update_model_list(all_model_info, model_list)
    new_count    = len(updated_list) - len(model_list)
    print(f"\n▶ 모델명 리스트 저장: 기준데이터/모델명 리스트.csv")
    print(f"   전체 {len(updated_list)}개 모델 (신규 추가 {new_count}개)")

    # ══════════════════════════════════════════
    # 최종 합치기 & 저장
    # ══════════════════════════════════════════
    frames = [f for f in [ss_result, pl_result] if not f.empty]
    result = pd.concat(frames, ignore_index=True)

    col_order = ["date", "ad_type", "category", "model",
                 "spend", "clicks", "impressions", "conv", "revenue"]
    result = result[col_order]

    for col in ["spend", "clicks", "impressions", "conv", "revenue"]:
        result[col] = result[col].astype(int)

    result.to_csv(out, index=False, encoding="utf-8-sig")

    print(f"\n[완료] {out.name} 저장됨")
    print(f"   총 {len(result):,}행 / {result['date'].nunique()}개월")
    print(f"   광고유형별: {result['ad_type'].value_counts().to_dict()}\n")

    summary = result.groupby(["ad_type", "category"])[["spend", "conv", "revenue"]].sum()
    summary["ROAS"] = (summary["revenue"] / summary["spend"]).round(1)
    print("--- 광고유형/카테고리별 요약 ---")
    print(summary.to_string())
    return result


if __name__ == "__main__":
    args = sys.argv[1:]
    run_pipeline(
        file_click    = args[0] if len(args) > 0 else None,
        file_conv     = args[1] if len(args) > 1 else None,
        file_campaign = args[2] if len(args) > 2 else None,
        file_store    = args[3] if len(args) > 3 else None,
        output_file   = args[4] if len(args) > 4 else None,
    )
