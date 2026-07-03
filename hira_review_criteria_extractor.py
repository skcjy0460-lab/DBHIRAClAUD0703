# -*- coding: utf-8 -*-
"""
심사기준(수가기준) API 통합 추출기
==================================================
건강보험심사평가원 공공데이터포털 API를 이용하여
  1) 개별 항목 검색/상세조회
  2) 전체 데이터 자동 페이지네이션 → 엑셀(.xlsx) 일괄 추출
두 가지 기능을 제공하는 Streamlit 앱.

대상 API (data.go.kr, 공공데이터포털)
  - 건강보험심사평가원_수가기준정보조회서비스   (B551182/mdfeeCrtrInfoService)
  - 건강보험심사평가원_신포괄기준정보조회서비스 (B551182/NdrgStdInfoService)

⚠ 중요 안내
  data.go.kr의 Open API 상세페이지(Swagger)는 자바스크립트로 렌더링되어
  검색 엔진/크롤러로는 "정확한 응답 필드명"까지는 확인이 불가능합니다.
  따라서 이 앱은 필드명을 하드코딩하지 않고, 최초 호출 시 실제 API가
  반환하는 필드를 자동으로 감지한 뒤 "필드 매핑" 화면에서 조정윤님이
  직접 확인/매핑하도록 설계했습니다. (실무 API는 문서와 실제 응답이
  다른 경우가 잦아 이 방식이 훨씬 안전합니다.)

  또한 Operation ID 중 아래 2개만 검색으로 100% 확인되었습니다.
    - getPharmacyMdfeeList  (수가기준정보조회서비스 · 약국수가목록)
    - getNdrgPayList        (신포괄기준정보조회서비스 · 신포괄지급목록)
  진료수가목록/한방수가목록/신포괄분류코드 등 나머지 Operation은
  사이드바에서 "직접 입력"으로 추가할 수 있게 해두었습니다.
  (조정윤님이 활용신청 후 Swagger 화면에서 정확한 이름을 확인하여
   입력하시면 즉시 동작합니다.)
"""

import io
import time
import json
import datetime
import xml.etree.ElementTree as ET

import requests
import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ------------------------------------------------------------------
# 기본 설정
# ------------------------------------------------------------------
st.set_page_config(
    page_title="심사기준 · 수가기준 API 추출기",
    page_icon="🏥",
    layout="wide",
)

REQUEST_TIMEOUT = 15
MAX_RETRY = 3
RETRY_BACKOFF_SEC = 1.5
PAGE_SLEEP_SEC = 0.15  # 공공데이터포털 트래픽 제한 보호용 딜레이

# data.go.kr 표준 에러코드 매핑 (참고: 공공데이터포털 개발가이드)
DATA_GO_KR_ERROR_MAP = {
    "1": "APPLICATION ERROR",
    "4": "HTTP ERROR",
    "5": "서비스 연결실패 오류",
    "10": "잘못된 요청 파라미터 오류 (INVALID_REQUEST_PARAMETER_ERROR)",
    "11": "필수요청 파라미터 누락 오류",
    "12": "해당 오픈API서비스가 없거나 폐기됨",
    "20": "서비스 접근거부 오류",
    "21": "일시적으로 사용할 수 없는 서비스 키",
    "22": "서비스 요청제한횟수 초과 오류 (하루 트래픽 초과)",
    "30": "등록되지 않은 서비스키 (SERVICE_KEY_IS_NOT_REGISTERED_ERROR) → 서비스키 오타/미승인 확인",
    "31": "기한만료된 서비스키",
    "32": "등록되지 않은 IP",
    "33": "서명되지 않은 호출",
    "99": "기타에러",
}

# ------------------------------------------------------------------
# 서비스(엔드포인트) 정의
# ------------------------------------------------------------------
SERVICES = {
    "수가기준정보조회서비스": {
        "base_url": "https://apis.data.go.kr/B551182/mdfeeCrtrInfoService",
        "description": (
            "의료수가코드/의료수가코드명/의료수가분류번호 등을 기준으로 "
            "진료수가목록·약국수가목록·한방수가목록의 심사기준(수가) 정보를 조회합니다."
        ),
        "operations": {
            "약국수가목록 (getPharmacyMdfeeList) ✅확인됨": "getPharmacyMdfeeList",
            "진료수가목록 (getMedFeeList) ⚠추정": "getMedFeeList",
            "한방수가목록 (getOriMdfeeList) ⚠추정": "getOriMdfeeList",
            "직접 입력": "__custom__",
        },
        # 검색용 파라미터 후보 (실제 응답 필드 감지 후 조정 가능)
        "search_params": ["mdfeeCd", "mdfeeNm", "mdfeeDivNo"],
        "search_labels": {
            "mdfeeCd": "수가코드",
            "mdfeeNm": "수가명(수가코드명)",
            "mdfeeDivNo": "수가분류번호",
        },
    },
    "신포괄기준정보조회서비스": {
        "base_url": "https://apis.data.go.kr/B551182/NdrgStdInfoService",
        "description": (
            "신포괄수가의 분류코드(행위/약제/치료재료), 포괄구분코드, 분류유형코드 등을 "
            "기준으로 적용개시일·적용종료일 등 신포괄 급여기준 정보를 조회합니다."
        ),
        "operations": {
            "신포괄지급목록 (getNdrgPayList) ✅확인됨": "getNdrgPayList",
            "신포괄분류목록 (getNdrgClsfList) ⚠추정": "getNdrgClsfList",
            "직접 입력": "__custom__",
        },
        "search_params": ["clsfCd", "clsfNm"],
        "search_labels": {
            "clsfCd": "분류코드",
            "clsfNm": "분류명",
        },
    },
}

# ------------------------------------------------------------------
# 세션 상태 초기화
# ------------------------------------------------------------------
def init_state():
    defaults = {
        "raw_records": [],          # 최근 수집한 원본 레코드 (list[dict])
        "detected_fields": [],      # 감지된 실제 응답 필드명 목록
        "field_mapping": {},        # {표준필드: 실제API필드}
        "last_service": None,
        "last_operation": None,
        "fetch_log": [],            # 수집 로그(디버깅용)
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


init_state()

STANDARD_FIELDS = [
    ("계획일자/시행일자", "plan_date"),
    ("일련번호", "seq_no"),
    ("구분", "category"),
    ("심사지침/개최일자", "guideline_date"),
    ("관련근거", "reference"),
    ("제목", "title"),
    ("내용(심사기준 관련 내용)", "content"),
]


# ------------------------------------------------------------------
# API 호출 유틸
# ------------------------------------------------------------------
def build_key_param(service_key: str, key_type: str):
    """
    공공데이터포털 서비스키는 '인코딩된 키(Encoding)'와
    '디코딩된 키(Decoding)' 두 가지 형태로 발급됩니다.
      - Decoding 키: requests의 params에 그대로 넣으면 requests가 1회 인코딩 → 정상
      - Encoding 키: 이미 %인코딩 되어 있어 params에 넣으면 이중 인코딩되어 오류 발생
    따라서 Encoding 키는 URL 뒤에 직접 이어붙이고, params에는 넣지 않는다.
    """
    if key_type == "디코딩 키 (일반 텍스트, 예: 대부분 +,/,= 문자가 보임)":
        return {"serviceKey": service_key}, ""
    else:
        return {}, f"?serviceKey={service_key}&"


def parse_xml_items(xml_text: str):
    """XML 응답을 list[dict]로 변환 (JSON 미지원 오퍼레이션 대비 fallback)."""
    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError:
        return [], None, "XML 파싱 실패 (응답이 XML/JSON 어느 쪽도 아닐 수 있습니다)"

    # 표준 공공데이터포털 XML 구조: response > header > resultCode/resultMsg
    #                              response > body > items > item ... > (필드)
    result_code = root.findtext(".//resultCode")
    result_msg = root.findtext(".//resultMsg")
    total_count_txt = root.findtext(".//totalCount")
    total_count = int(total_count_txt) if total_count_txt and total_count_txt.isdigit() else None

    items = []
    for item in root.findall(".//items/item"):
        record = {child.tag: (child.text or "").strip() for child in item}
        if record:
            items.append(record)

    error_msg = None
    if result_code not in (None, "00", "0"):
        mapped = DATA_GO_KR_ERROR_MAP.get(result_code, "알 수 없는 오류")
        error_msg = f"[{result_code}] {result_msg or ''} → {mapped}"

    return items, total_count, error_msg


def parse_json_items(payload: dict):
    """JSON 응답을 list[dict]로 변환."""
    try:
        body = payload["response"]["body"]
        header = payload["response"].get("header", {})
    except (KeyError, TypeError):
        return [], None, "예상치 못한 JSON 구조입니다 (response/body 없음)"

    result_code = str(header.get("resultCode", "00"))
    result_msg = header.get("resultMsg", "")
    total_count = body.get("totalCount")

    items_raw = body.get("items")
    items = []
    if isinstance(items_raw, dict):
        inner = items_raw.get("item", [])
        items = inner if isinstance(inner, list) else [inner]
    elif isinstance(items_raw, list):
        items = items_raw

    error_msg = None
    if result_code not in ("00", "0"):
        mapped = DATA_GO_KR_ERROR_MAP.get(result_code, "알 수 없는 오류")
        error_msg = f"[{result_code}] {result_msg} → {mapped}"

    return items, total_count, error_msg


def call_api(base_url, operation, service_key, key_type, extra_params, page_no=1, num_rows=100):
    """API 단건 호출. (items, total_count, error_msg) 반환."""
    params, url_prefix_query = build_key_param(service_key, key_type)
    params.update({
        "pageNo": page_no,
        "numOfRows": num_rows,
        "_type": "json",
    })
    params.update({k: v for k, v in extra_params.items() if v not in (None, "")})

    url = f"{base_url}/{operation}"
    if url_prefix_query:
        # 인코딩 키 방식: 서비스키를 URL에 직접 결합(재인코딩 방지), 나머지는 params로 전송
        full_url = url + url_prefix_query.rstrip("&")
    else:
        full_url = url

    last_err = None
    for attempt in range(1, MAX_RETRY + 1):
        try:
            resp = requests.get(full_url, params=params, timeout=REQUEST_TIMEOUT)
            resp.raise_for_status()
            text = resp.text.strip()

            if text.startswith("<"):
                # XML로 응답 (JSON 미지원 오퍼레이션 또는 오류 XML)
                return parse_xml_items(text)
            else:
                payload = json.loads(text)
                return parse_json_items(payload)

        except requests.exceptions.RequestException as e:
            last_err = f"네트워크 오류: {e}"
        except json.JSONDecodeError:
            last_err = f"JSON 파싱 실패. 응답 원문 일부: {text[:200]}"
        except Exception as e:
            last_err = f"알 수 없는 오류: {e}"

        if attempt < MAX_RETRY:
            time.sleep(RETRY_BACKOFF_SEC * attempt)

    return [], None, last_err


def fetch_all_pages(base_url, operation, service_key, key_type, extra_params,
                     num_rows=500, max_rows_limit=None, progress_callback=None):
    """전체 페이지 자동 순회하여 모든 레코드 수집."""
    all_items = []
    page_no = 1
    total_count = None
    errors = []

    while True:
        items, total_count, err = call_api(
            base_url, operation, service_key, key_type, extra_params,
            page_no=page_no, num_rows=num_rows,
        )
        if err:
            errors.append(f"{page_no}페이지: {err}")
            break

        if not items:
            break

        all_items.extend(items)

        if progress_callback:
            progress_callback(len(all_items), total_count)

        if max_rows_limit and len(all_items) >= max_rows_limit:
            all_items = all_items[:max_rows_limit]
            break

        if total_count is not None and len(all_items) >= total_count:
            break

        page_no += 1
        time.sleep(PAGE_SLEEP_SEC)

        # 안전장치: 무한루프 방지 (혹시 totalCount 미제공 API 대비)
        if page_no > 3000:
            errors.append("페이지 3000회 초과 → 중단 (totalCount 미제공 API 가능성)")
            break

    return all_items, total_count, errors


# ------------------------------------------------------------------
# 엑셀 변환 (스타일 적용)
# ------------------------------------------------------------------
def dataframe_to_styled_excel(df: pd.DataFrame, sheet_name="심사기준DB") -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.sheets[sheet_name]

        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True, size=11)
        thin = Side(style="thin", color="D9D9D9")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        wrap_align = Alignment(vertical="top", wrap_text=True)

        for col_idx, col_name in enumerate(df.columns, start=1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border

            # 열 너비 자동 조정 (내용 길이 기반, 최대 60)
            max_len = max(
                [len(str(col_name))] +
                [len(str(v)) for v in df[col_name].astype(str).tolist()[:200]]
            )
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 4, 12), 60)

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
            for cell in row:
                cell.alignment = wrap_align
                cell.border = border

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    return output.getvalue()


# ------------------------------------------------------------------
# 사이드바: 접속 설정
# ------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ API 접속 설정")

    service_name = st.selectbox("서비스 선택", list(SERVICES.keys()))
    service_cfg = SERVICES[service_name]
    st.caption(service_cfg["description"])

    op_label = st.selectbox("Operation(기능) 선택", list(service_cfg["operations"].keys()))
    operation = service_cfg["operations"][op_label]
    if operation == "__custom__":
        operation = st.text_input(
            "Operation ID 직접 입력",
            placeholder="예: getMedFeeList",
            help="data.go.kr 활용신청 후 Swagger 화면(요청주소)에서 정확한 명칭을 확인해 입력하세요.",
        )

    st.divider()
    service_key = st.text_input(
        "서비스키(인증키)",
        type="password",
        help="공공데이터포털 마이페이지 > 개발계정 상세보기에서 발급받은 키를 입력하세요. 이 앱은 키를 저장/전송하지 않고 세션 내에서만 사용합니다.",
    )
    key_type = st.radio(
        "서비스키 형태",
        [
            "디코딩 키 (일반 텍스트, 예: 대부분 +,/,= 문자가 보임)",
            "인코딩 키 (%2B, %2F 등 %가 보이는 키)",
        ],
        index=0,
    )

    base_url_override = st.text_input("Base URL (필요시만 수정)", value=service_cfg["base_url"])

    st.divider()
    with st.expander("🔎 검색 조건 (선택)"):
        st.caption("비워두면 전체 조회됩니다. '개별 조회' 탭 검색에도 함께 사용됩니다.")
        extra_params = {}
        for p in service_cfg["search_params"]:
            label = service_cfg["search_labels"].get(p, p)
            extra_params[p] = st.text_input(f"{label} ({p})", key=f"param_{service_name}_{p}")

        custom_param_raw = st.text_area(
            "추가 파라미터 (key=value, 줄바꿈으로 구분)",
            placeholder="예:\nbaseDate=20260101",
            help="검색 조건에 없는 파라미터를 추가로 넣고 싶을 때 사용하세요.",
        )
        for line in custom_param_raw.splitlines():
            if "=" in line:
                k, v = line.split("=", 1)
                extra_params[k.strip()] = v.strip()

    st.divider()
    st.caption(
        "💡 이 앱은 서비스키를 서버에 저장하지 않습니다. 브라우저 세션이 끝나면 사라지니, "
        "실제 운영 시에는 st.secrets 또는 환경변수 사용을 권장합니다."
    )


# ------------------------------------------------------------------
# 메인 화면
# ------------------------------------------------------------------
st.title("🏥 심사기준(수가기준) API 통합 추출기")
st.caption(
    "공공데이터포털 API를 이용해 심사기준/수가기준 데이터를 개별 조회하거나, "
    "전체 데이터를 한 번에 엑셀로 추출하여 자체 DB 구축에 활용할 수 있습니다."
)

if not operation:
    st.warning("⬅️ 왼쪽 사이드바에서 Operation ID를 선택하거나 직접 입력해주세요.")
    st.stop()

tab_bulk, tab_single, tab_mapping, tab_log = st.tabs(
    ["📦 전체 리스트 추출 (엑셀)", "🔍 개별 조회", "🧩 필드 매핑 설정", "🪵 요청 로그"]
)

# ------------------------------------------------------------------
# TAB 1: 전체 리스트 추출
# ------------------------------------------------------------------
with tab_bulk:
    st.subheader("전체 데이터 일괄 추출 → 엑셀 변환")
    st.markdown(
        "API 페이지네이션을 자동으로 끝까지 순회하며 **전체 데이터를 한 번에 수집**합니다. "
        "하나씩 검색/입력할 필요 없이, 수집 완료 후 엑셀로 다운로드하여 바로 자체 DB 자료로 사용하실 수 있습니다."
    )

    col1, col2, col3 = st.columns(3)
    with col1:
        num_rows = st.number_input("페이지당 요청 건수 (numOfRows)", min_value=10, max_value=1000, value=500, step=10)
    with col2:
        limit_rows = st.number_input("최대 수집 건수 (0=제한없음, 테스트용)", min_value=0, value=0, step=100)
    with col3:
        st.metric("현재 캐시된 레코드 수", len(st.session_state["raw_records"]))

    if st.button("🚀 전체 데이터 수집 시작", type="primary", use_container_width=True):
        if not service_key:
            st.error("서비스키를 입력해주세요.")
        else:
            progress_bar = st.progress(0.0, text="수집 시작...")
            status_box = st.empty()

            def _cb(collected, total):
                if total:
                    pct = min(collected / total, 1.0)
                    progress_bar.progress(pct, text=f"{collected:,} / {total:,} 건 수집 중...")
                else:
                    status_box.info(f"{collected:,}건 수집 중... (전체 건수 미확인)")

            t0 = time.time()
            items, total_count, errors = fetch_all_pages(
                base_url_override, operation, service_key, key_type, extra_params,
                num_rows=int(num_rows),
                max_rows_limit=int(limit_rows) if limit_rows > 0 else None,
                progress_callback=_cb,
            )
            elapsed = time.time() - t0
            progress_bar.progress(1.0, text="완료")

            st.session_state["raw_records"] = items
            st.session_state["detected_fields"] = sorted({k for row in items for k in row.keys()})
            st.session_state["last_service"] = service_name
            st.session_state["last_operation"] = operation
            st.session_state["fetch_log"].append({
                "시각": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "서비스": service_name,
                "operation": operation,
                "수집건수": len(items),
                "전체건수(API보고)": total_count,
                "소요시간(초)": round(elapsed, 1),
                "오류": "; ".join(errors) if errors else "",
            })

            if errors:
                st.error("⚠️ 수집 중 오류가 발생했습니다:\n\n" + "\n".join(f"- {e}" for e in errors))
            if items:
                st.success(f"✅ 총 {len(items):,}건 수집 완료 (소요 {elapsed:.1f}초)")
            elif not errors:
                st.warning("수집된 데이터가 없습니다. 서비스키/Operation ID/검색조건을 확인해주세요.")

    st.divider()

    if st.session_state["raw_records"]:
        df_raw = pd.DataFrame(st.session_state["raw_records"])
        st.markdown(f"**미리보기** (실제 API 원본 필드 그대로 · 총 {len(df_raw):,}행)")
        st.dataframe(df_raw, use_container_width=True, height=350)

        st.markdown("#### 다운로드 옵션")
        use_mapping = st.checkbox(
            "표준 컬럼(일련번호/구분/심사지침·개최일자/관련근거/제목/내용 등)으로 매핑하여 다운로드",
            value=bool(st.session_state["field_mapping"]),
            help="'🧩 필드 매핑 설정' 탭에서 먼저 매핑을 지정해야 적용됩니다. 지정하지 않으면 원본 필드 그대로 다운로드됩니다.",
        )

        if use_mapping and st.session_state["field_mapping"]:
            mapped_df = pd.DataFrame()
            for std_label, std_key in STANDARD_FIELDS:
                src_field = st.session_state["field_mapping"].get(std_key)
                if src_field and src_field in df_raw.columns:
                    mapped_df[std_label] = df_raw[src_field]
                else:
                    mapped_df[std_label] = ""
            export_df = mapped_df
        else:
            export_df = df_raw

        excel_bytes = dataframe_to_styled_excel(export_df)
        fname = f"{service_name}_{operation}_{datetime.date.today().isoformat()}.xlsx"
        st.download_button(
            "⬇️ 엑셀(.xlsx) 다운로드",
            data=excel_bytes,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
    else:
        st.info("아직 수집된 데이터가 없습니다. 위의 '전체 데이터 수집 시작' 버튼을 눌러주세요.")


# ------------------------------------------------------------------
# TAB 2: 개별 조회
# ------------------------------------------------------------------
with tab_single:
    st.subheader("개별 항목 검색 / 상세조회")

    sub_tab_local, sub_tab_live = st.tabs(["📂 캐시된 전체 데이터에서 검색 (빠름)", "🌐 API 단건 실시간 조회"])

    with sub_tab_local:
        if not st.session_state["raw_records"]:
            st.info("먼저 '📦 전체 리스트 추출' 탭에서 전체 데이터를 한 번 수집하면, 이후에는 API 호출 없이 즉시 검색할 수 있습니다.")
        else:
            df_raw = pd.DataFrame(st.session_state["raw_records"])
            keyword = st.text_input("검색어 (제목/내용/전체 필드 통합 검색)", key="local_search_kw")
            target_cols = st.multiselect(
                "검색 대상 컬럼 (선택 안 하면 전체 컬럼 대상)",
                options=list(df_raw.columns),
            )

            if keyword:
                cols_to_search = target_cols if target_cols else list(df_raw.columns)
                mask = df_raw[cols_to_search].astype(str).apply(
                    lambda col: col.str.contains(keyword, case=False, na=False)
                ).any(axis=1)
                result_df = df_raw[mask]
            else:
                result_df = df_raw

            st.caption(f"검색 결과: {len(result_df):,}건")
            st.dataframe(result_df, use_container_width=True, height=300)

            if len(result_df) > 0:
                idx = st.selectbox(
                    "상세보기 할 행 선택",
                    options=list(result_df.index),
                    format_func=lambda i: " | ".join(str(result_df.loc[i, c])[:20] for c in result_df.columns[:3]),
                )
                st.markdown("##### 상세 정보")
                detail = result_df.loc[idx].to_dict()
                for k, v in detail.items():
                    st.markdown(f"**{k}**: {v}")

    with sub_tab_live:
        st.caption("서버에 직접 단건(또는 소량) 조회 요청을 보냅니다. 왼쪽 사이드바의 검색 조건이 함께 적용됩니다.")
        if st.button("🔍 실시간 조회 실행"):
            if not service_key:
                st.error("서비스키를 입력해주세요.")
            else:
                items, total_count, err = call_api(
                    base_url_override, operation, service_key, key_type,
                    extra_params, page_no=1, num_rows=50,
                )
                if err:
                    st.error(f"오류: {err}")
                if items:
                    st.success(f"{len(items)}건 조회됨 (전체 {total_count if total_count is not None else '알수없음'}건 중)")
                    st.dataframe(pd.DataFrame(items), use_container_width=True)
                    st.session_state["detected_fields"] = sorted(
                        set(st.session_state["detected_fields"]) | {k for row in items for k in row.keys()}
                    )
                elif not err:
                    st.warning("조회된 데이터가 없습니다.")


# ------------------------------------------------------------------
# TAB 3: 필드 매핑 설정
# ------------------------------------------------------------------
with tab_mapping:
    st.subheader("표준 컬럼 ↔ 실제 API 필드 매핑")
    st.markdown(
        "말씀하신 6개 표준 컬럼(계획일자·일련번호·구분·심사지침·관련근거·제목·내용)과 "
        "**실제 API가 반환하는 필드명**은 서비스에 따라 다를 수 있습니다. "
        "아래에서 한 번 매핑해두면, '전체 리스트 추출' 탭 다운로드 시 자동으로 표준 컬럼명으로 변환됩니다."
    )

    if not st.session_state["detected_fields"]:
        st.info(
            "아직 감지된 API 필드가 없습니다. '전체 리스트 추출' 또는 '개별 조회 > 실시간 조회'를 "
            "한 번 실행하면 실제 응답 필드 목록이 자동으로 채워집니다."
        )
    else:
        st.success(f"감지된 실제 API 필드 {len(st.session_state['detected_fields'])}개: " +
                    ", ".join(st.session_state["detected_fields"]))

        options = ["(매핑 안함)"] + st.session_state["detected_fields"]
        new_mapping = {}
        for std_label, std_key in STANDARD_FIELDS:
            current = st.session_state["field_mapping"].get(std_key, "(매핑 안함)")
            idx = options.index(current) if current in options else 0
            picked = st.selectbox(f"{std_label} ←", options, index=idx, key=f"map_{std_key}")
            if picked != "(매핑 안함)":
                new_mapping[std_key] = picked

        if st.button("💾 매핑 저장", type="primary"):
            st.session_state["field_mapping"] = new_mapping
            st.success("매핑이 저장되었습니다. '전체 리스트 추출' 탭에서 매핑 적용 다운로드를 이용하세요.")


# ------------------------------------------------------------------
# TAB 4: 요청 로그
# ------------------------------------------------------------------
with tab_log:
    st.subheader("수집 이력")
    if st.session_state["fetch_log"]:
        st.dataframe(pd.DataFrame(st.session_state["fetch_log"]), use_container_width=True)
    else:
        st.caption("아직 수집 이력이 없습니다.")

    st.divider()
    st.markdown("##### 자주 발생하는 오류 코드 안내")
    st.dataframe(
        pd.DataFrame(
            [{"코드": k, "의미": v} for k, v in DATA_GO_KR_ERROR_MAP.items()]
        ),
        use_container_width=True,
        height=250,
    )
