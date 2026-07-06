# -*- coding: utf-8 -*-
"""
ClaimLens 부속 도구 - 효능효과/용법용량 ITEM_SEQ 일괄순회 채우기
====================================================================
배경
  기존 drug_api_extractor.py의 "③ 식약처 의약품 상세정보
  (getDrugPrdtPrmsnDtlInq06)"는 CONFIG(필드매핑)는 정확하지만,
  이 오퍼레이션은 원래 품목 1건(ITEM_SEQ 또는 ITEM_NAME) 단위로
  조회하는 API라서 대량추출 탭의 "검색어 1개" 방식으로는
  21,878건 전체를 채울 수 없었음 (96%가 비어있던 근본 원인).

이 페이지가 하는 일
  1) 기존 Master 엑셀(품목기준코드 컬럼 필수)을 업로드
  2) 효능효과/용법용량이 비어있는 행만 추출
  3) 지정한 범위(start_idx~end_idx)만큼 ITEM_SEQ를 하나씩 순회하며
     ③ 상세조회 API 호출 → EE_DOC_DATA/UD_DOC_DATA/NB_DOC_DATA 수집
  4) HTML 태그 제거 후 session_state에 누적 (품목기준코드 기준 중복제거)
  5) 누적분을 원본 Master와 병합해서 엑셀 다운로드
  6) 그래도 못 채운 행은 "동일 성분+함량+제형" 그룹 내 값 상속으로 2차 보충
     (신뢰도 컬럼에 "성분상속"이라고 표시해 원본 API 값과 구분)

Streamlit Cloud 타임아웃을 피하려면 한 번에 500~1000건씩 여러 세션에
나눠 돌리는 것을 권장 (사이드바에서 범위 지정).
"""

import re
import time

import pandas as pd
import requests
import streamlit as st

st.set_page_config(page_title="효능효과/용법용량 일괄채우기", page_icon="🧩", layout="wide")

DETAIL_BASE_URL = "http://apis.data.go.kr/1471000/DrugPrdtPrmsnInfoService07"
DETAIL_OPERATION = "getDrugPrdtPrmsnDtlInq06"  # drug_api_extractor.py와 동일 (확정된 오퍼레이션)

TAG_RE = re.compile(r"<[^>]+>")


def strip_html(text: str) -> str:
    if not text:
        return ""
    text = TAG_RE.sub(" ", str(text))
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalize_service_key(raw_key: str, key_mode: str) -> str:
    if key_mode == "인코딩(Encoding) 키를 붙여넣었어요":
        import urllib.parse as up
        try:
            return up.unquote(raw_key)
        except Exception:
            return raw_key
    return raw_key


def call_detail_api(service_key: str, item_seq: str, timeout: int = 15):
    """ITEM_SEQ 1건에 대해 상세조회 API 호출. drug_api_extractor.py의 call_api와
    동일한 파싱 규칙(response 래핑 유무 모두 지원)을 사용."""
    url = f"{DETAIL_BASE_URL}/{DETAIL_OPERATION}"
    params = {
        "serviceKey": service_key,
        "item_seq": item_seq,
        "numOfRows": 1,
        "pageNo": 1,
        "type": "json",
    }
    try:
        resp = requests.get(url, params=params, timeout=timeout)
    except requests.exceptions.RequestException as e:
        return None, f"네트워크 오류: {e}"

    text = resp.text.strip()
    if text.startswith("<"):
        return None, f"XML 응답(에러 가능성): {text[:200]}"

    try:
        data = resp.json()
    except ValueError:
        return None, f"알 수 없는 응답: {text[:200]}"

    try:
        root_obj = data["response"] if "response" in data and isinstance(data["response"], dict) else data
        body = root_obj["body"]
        header = root_obj["header"]
        result_code = str(header.get("resultCode", ""))
        if result_code not in ("00", "0"):
            return None, f"[{result_code}] {header.get('resultMsg', '')}"

        items_raw = body.get("items", "")
        if items_raw in ("", None):
            return None, "결과 없음 (해당 품목 상세문서 미등록)"
        if isinstance(items_raw, dict) and "item" in items_raw:
            item_val = items_raw["item"]
            item = item_val[0] if isinstance(item_val, list) else item_val
        elif isinstance(items_raw, list):
            item = items_raw[0]
        else:
            return None, "예상치 못한 items 구조"

        return {
            "품목기준코드": str(item.get("ITEM_SEQ", item_seq)),
            "효능효과": strip_html(item.get("EE_DOC_DATA", "")),
            "용법용량": strip_html(item.get("UD_DOC_DATA", "")),
            "사용상주의사항_API": strip_html(item.get("NB_DOC_DATA", "")),
        }, None
    except (KeyError, TypeError):
        return None, f"예상치 못한 JSON 구조: {str(data)[:300]}"


# ============================================================
# 사이드바
# ============================================================
with st.sidebar:
    st.header("🔑 인증 설정")
    service_key_raw = st.text_input("공공데이터포털 서비스키", type="password")
    key_mode = st.radio(
        "붙여넣은 키의 종류",
        ["디코딩(Decoding) 키를 붙여넣었어요", "인코딩(Encoding) 키를 붙여넣었어요"],
        index=0,
    )
    service_key = normalize_service_key(service_key_raw, key_mode)

    st.divider()
    st.header("⚙️ 수집 범위 (청크)")
    st.caption("Streamlit Cloud 타임아웃 방지를 위해 한 번에 500~1000건 권장")
    start_idx = st.number_input("시작 인덱스 (0부터)", min_value=0, value=0, step=100)
    end_idx = st.number_input("종료 인덱스 (미포함)", min_value=1, value=500, step=100)
    sleep_sec = st.slider("요청 간 대기시간(초)", 0.0, 1.0, 0.15, 0.05)

    st.divider()
    if st.button("🗑️ 누적 세션 초기화", use_container_width=True):
        st.session_state.pop("filled_results", None)
        st.success("초기화 완료")

st.title("🧩 효능효과/용법용량 ITEM_SEQ 일괄채우기")
st.caption("Master 엑셀 업로드 → 빈 칸인 품목만 골라 ITEM_SEQ로 하나씩 조회 → 누적 병합")

# ============================================================
# 1) Master 엑셀 업로드
# ============================================================
master_file = st.file_uploader("Master 엑셀 업로드 (품목기준코드/약품코드, 성분명, 함량, 제형, 효능효과, 용법용량 컬럼 필요)", type=["xlsx"])

if master_file:
    df = pd.read_excel(master_file, sheet_name=0)
    st.write(f"전체 {len(df):,}행 로드 완료")

    code_col = "약품코드" if "약품코드" in df.columns else ("품목기준코드" if "품목기준코드" in df.columns else None)
    if code_col is None:
        st.error("품목기준코드(또는 약품코드) 컬럼을 찾을 수 없습니다.")
        st.stop()

    missing_mask = df["효능효과"].isna() & df["용법용량"].isna() if "효능효과" in df.columns else df.index >= 0
    missing_df = df[missing_mask].reset_index(drop=True)
    st.info(f"효능효과/용법용량 둘 다 비어있는 행: {len(missing_df):,}건")

    chunk_df = missing_df.iloc[int(start_idx):int(end_idx)]
    st.write(f"이번 청크에서 조회할 품목: {len(chunk_df):,}건 (인덱스 {start_idx}~{end_idx})")

    if "filled_results" not in st.session_state:
        st.session_state["filled_results"] = {}  # 품목기준코드 -> dict

    run = st.button("🚀 이 범위 수집 시작", type="primary", disabled=not service_key)
    if not service_key:
        st.warning("사이드바에 서비스키를 입력해주세요.")

    if run:
        progress = st.progress(0.0)
        status = st.empty()
        errors = []
        n = len(chunk_df)
        for i, (_, row) in enumerate(chunk_df.iterrows()):
            item_seq = str(row[code_col]).strip()
            result, err = call_detail_api(service_key, item_seq)
            if result:
                st.session_state["filled_results"][item_seq] = result
            else:
                errors.append(f"{item_seq}: {err}")
            progress.progress((i + 1) / max(n, 1))
            status.text(f"{i+1}/{n} 처리 중... (누적 성공 {len(st.session_state['filled_results']):,}건)")
            time.sleep(sleep_sec)

        st.success(f"이번 청크 완료. 누적 수집: {len(st.session_state['filled_results']):,}건")
        if errors:
            with st.expander(f"⚠️ 실패/미등록 {len(errors)}건"):
                for e in errors[:200]:
                    st.text(e)

    # ============================================================
    # 2) 누적 결과를 원본 Master와 병합
    # ============================================================
    if st.session_state["filled_results"]:
        st.divider()
        st.subheader("📥 누적 결과 병합 다운로드")

        result_df = pd.DataFrame(st.session_state["filled_results"].values())
        merged = df.copy()
        merged[code_col] = merged[code_col].astype(str).str.strip()
        result_df["품목기준코드"] = result_df["품목기준코드"].astype(str).str.strip()

        merged = merged.merge(
            result_df, left_on=code_col, right_on="품목기준코드", how="left", suffixes=("", "_신규")
        )

        if "효능효과_신규" in merged.columns:
            merged["효능효과"] = merged["효능효과"].where(merged["효능효과"].notna() & (merged["효능효과"] != ""), merged["효능효과_신규"])
            merged["용법용량"] = merged["용법용량"].where(merged["용법용량"].notna() & (merged["용법용량"] != ""), merged["용법용량_신규"])
            merged = merged.drop(columns=["효능효과_신규", "용법용량_신규", "품목기준코드"], errors="ignore")

        still_missing = merged["효능효과"].isna().sum() if "효능효과" in merged.columns else 0
        st.write(f"병합 후 여전히 비어있는 행: {still_missing:,}건")

        import io
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            merged.to_excel(writer, index=False, sheet_name="Drug_Master_병합")
        st.download_button(
            "⬇️ 병합된 Master 엑셀 다운로드",
            data=buf.getvalue(),
            file_name="약품_Master_효능효과_채움_병합.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # ============================================================
    # 3) 2차 보충: 동일 성분+함량+제형 그룹 상속
    # ============================================================
    st.divider()
    st.subheader("🔁 2차 보충: 동일 성분·함량·제형 그룹 상속")
    st.caption(
        "API로도 채워지지 않는 품목(제네릭이 자체 문서 없이 원개발사 문서를 참조하는 경우 등)에 대해, "
        "같은 성분명+함량+제형 그룹 안에 채워진 값이 있으면 그대로 상속합니다. "
        "'신뢰도' 컬럼에 '성분상속'이라고 표시되므로 원본 API 값과 구분해서 검수할 수 있습니다."
    )
    do_inherit = st.button("성분 그룹 상속 실행 (병합 다운로드 이후 사용 권장)")
    if do_inherit:
        base = merged if "merged" in dir() and isinstance(merged, pd.DataFrame) else df
        base = base.copy()
        if "신뢰도" not in base.columns:
            base["신뢰도"] = ""

        group_cols = [c for c in ["성분명", "함량", "제형"] if c in base.columns]
        if not group_cols:
            st.error("성분명/함량/제형 컬럼이 없어 그룹 상속을 수행할 수 없습니다.")
        else:
            filled_before = base["효능효과"].notna().sum()

            def pick_reference(group: pd.DataFrame):
                ref = group[group["효능효과"].notna() & (group["효능효과"] != "")]
                if ref.empty:
                    return None
                return ref.iloc[0]

            for _, group in base.groupby(group_cols):
                ref_row = pick_reference(group)
                if ref_row is None:
                    continue
                empty_idx = group[group["효능효과"].isna() | (group["효능효과"] == "")].index
                base.loc[empty_idx, "효능효과"] = ref_row["효능효과"]
                base.loc[empty_idx, "용법용량"] = ref_row["용법용량"]
                base.loc[empty_idx, "신뢰도"] = base.loc[empty_idx, "신뢰도"].where(
                    base.loc[empty_idx, "신뢰도"] != "", "성분상속(동일 성분·함량·제형 참조)"
                )

            filled_after = base["효능효과"].notna().sum()
            st.success(f"성분 그룹 상속으로 {filled_after - filled_before:,}건 추가 채움")

            import io
            buf2 = io.BytesIO()
            with pd.ExcelWriter(buf2, engine="openpyxl") as writer:
                base.to_excel(writer, index=False, sheet_name="Drug_Master_최종")
            st.download_button(
                "⬇️ 성분상속 반영 최종 엑셀 다운로드",
                data=buf2.getvalue(),
                file_name="약품_Master_효능효과_최종.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
else:
    st.info("먼저 Master 엑셀을 업로드해주세요.")
