"""
VC투자팀 Track Record 백엔드
Google Sheets API를 사용하여 데이터를 가져오고 계산 로직을 처리합니다.
"""

from flask import Flask, jsonify, request
from flask_cors import CORS
import requests
from datetime import date, timedelta
import calendar
import math
import os
import json
import threading
import gspread
from google.oauth2.service_account import Credentials

app = Flask(__name__)
CORS(app)

# Google Sheets API 설정
SHEET_ID = "1xYZqiMRclX6OkQn0V-RVwteC4D8OMUjDd22-3fb3hbg"
API_KEY = "AIzaSyDTbY36CW4NqgqiZIv-_FuoRWMzAykNZ3U"

# DIRECT_HOLDINGS 쓰기용 서비스 계정 (읽기는 API_KEY로, 쓰기는 이 계정으로 수행)
# 로컬 개발: secrets/service-account.json 파일 사용
# 배포(Vercel 등): GOOGLE_SERVICE_ACCOUNT_JSON 환경변수에 서비스 계정 JSON 전체를 문자열로 저장
SERVICE_ACCOUNT_FILE = "secrets/service-account.json"
SHEETS_WRITE_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# 캐시된 데이터 - 시트 원본(_data_cache), dict 변환본(_dicts_cache), 조회 인덱스(_index_cache)는
# 같은 원본에서 파생되므로 반드시 함께 무효화해야 합니다 (clear_data_caches / invalidate_sheet 사용)
_data_cache = {}
_dicts_cache = {}
_index_cache = {}

# 앱이 사용하는 전체 시트 목록. 캐시가 빈 상태에서 첫 조회가 발생하면 이 시트들을
# values:batchGet 한 번으로 모두 가져옵니다 (시트별 개별 요청 대비 HTTP 왕복 ~13회 -> 1회).
ALL_SHEETS = [
    "FUND_BASICS", "FUND_BASE_FX", "FUND_FX", "FUND_SUBSCRIPTION",
    "FUND_EXPENSE", "FUND_DISTRIBUTION",
    "DIRECT_INVESTMENT", "DIRECT_COMPANY", "DIRECT_REALIZATION",
    "DIRECT_CONVERSION", "DIRECT_MARKUP_A", "DIRECT_MARKUP_E", "DIRECT_HOLDINGS",
]


def clear_data_caches():
    """확인 버튼 클릭 시마다 Google Sheets에서 최신 데이터를 다시 읽도록 모든 캐시를 비웁니다."""
    _data_cache.clear()
    _dicts_cache.clear()
    _index_cache.clear()


def invalidate_sheet(sheet_name):
    """특정 시트의 캐시만 무효화합니다 (해당 시트에 쓰기 후 최신값 반영용)."""
    _data_cache.pop(sheet_name, None)
    _dicts_cache.pop(sheet_name, None)
    _index_cache.pop(sheet_name, None)


# 쓰기용 gspread 핸들 캐시 - 서비스 계정 인증(토큰 교환)과 시트 메타데이터 조회를 요청마다 반복하지 않습니다
_gspread_ws_cache = {}
_gspread_lock = threading.Lock()


def get_gspread_worksheet(sheet_name):
    """쓰기 권한이 있는 서비스 계정으로 워크시트 핸들을 가져옵니다 (최초 1회만 인증 후 재사용)."""
    with _gspread_lock:
        if sheet_name not in _gspread_ws_cache:
            service_account_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
            if service_account_json:
                creds = Credentials.from_service_account_info(json.loads(service_account_json), scopes=SHEETS_WRITE_SCOPES)
            else:
                creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SHEETS_WRITE_SCOPES)
            gc = gspread.authorize(creds)
            sh = gc.open_by_key(SHEET_ID)
            _gspread_ws_cache[sheet_name] = sh.worksheet(sheet_name)
        return _gspread_ws_cache[sheet_name]


def _clean_sheet_values(sheet_name, data):
    """"DATA" 시트를 제외한 모든 시트는 Row 2가 한글 설명이므로 스킵합니다."""
    if sheet_name != "DATA" and len(data) > 2:
        # 헤더 + Row 3부터의 데이터 (Row 2 한글 설명 스킵)
        return [data[0]] + data[2:]
    return data


def get_sheet_data(sheet_name):
    """
    Google Sheets에서 데이터를 가져옵니다.
    캐시가 완전히 빈 상태(확인 클릭 직후)면 ALL_SHEETS 전체를 values:batchGet 한 번으로 가져와
    캐시를 채웁니다. 그 외(쓰기 후 단일 시트 재조회 등)에는 해당 시트만 개별 요청합니다.
    valueRenderOption=UNFORMATTED_VALUE를 사용합니다 - 기본값(FORMATTED_VALUE)은 셀 표시 형식(예: 소수점
    자릿수 반올림)이 적용된 문자열을 반환해 실제 저장값과 다를 수 있기 때문입니다 (예: 74.98이 "75"로 반환됨).
    이 경우 날짜 컬럼은 문자열이 아닌 Google Sheets 일련번호(정수)로 반환되므로 parse_date()에서 함께 처리합니다.
    """
    if sheet_name in _data_cache:
        return _data_cache[sheet_name]

    if not _data_cache and sheet_name in ALL_SHEETS:
        try:
            ranges = "&".join(f"ranges={s}" for s in ALL_SHEETS)
            url = (
                f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}/values:batchGet"
                f"?key={API_KEY}&valueRenderOption=UNFORMATTED_VALUE&{ranges}"
            )
            response = requests.get(url)
            if response.status_code == 200:
                # valueRanges는 요청한 ranges 순서대로 반환됩니다
                for name, value_range in zip(ALL_SHEETS, response.json().get("valueRanges", [])):
                    _data_cache[name] = _clean_sheet_values(name, value_range.get("values", []))
                if sheet_name in _data_cache:
                    return _data_cache[sheet_name]
        except Exception as e:
            print(f"Error batch-fetching sheets: {e}")
        # batchGet 실패 시 아래 개별 요청으로 폴백

    try:
        url = (
            f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}/values/{sheet_name}"
            f"?key={API_KEY}&valueRenderOption=UNFORMATTED_VALUE"
        )
        response = requests.get(url)

        if response.status_code == 200:
            cleaned_data = _clean_sheet_values(sheet_name, response.json().get("values", []))
            _data_cache[sheet_name] = cleaned_data
            return cleaned_data
        else:
            return []
    except Exception as e:
        print(f"Error fetching {sheet_name}: {e}")
        return []


def get_sheet_dicts(sheet_name):
    """rows_to_dicts(get_sheet_data(...)) 결과를 시트별로 캐시합니다 (환율 조회 등 반복 호출 시 재변환 방지)."""
    if sheet_name not in _dicts_cache:
        _dicts_cache[sheet_name] = rows_to_dicts(get_sheet_data(sheet_name))
    return _dicts_cache[sheet_name]


def get_column_index(headers, column_name):
    """헤더에서 컬럼 인덱스를 찾습니다. (공백 무시)"""
    for i, header in enumerate(headers):
        if header.strip() == column_name.strip():
            return i
    return -1


def rows_to_dicts(data):
    """시트 데이터를 [{header: value}, ...] 형태로 변환합니다. (헤더 공백 제거)"""
    if not data or len(data) < 1:
        return []

    headers = [h.strip() for h in data[0]]
    result = []
    for row in data[1:]:
        d = {}
        for i, h in enumerate(headers):
            d[h] = row[i] if i < len(row) else ""
        result.append(d)
    return result


def parse_number(value):
    """'  1,234,567 ' 같은 문자열을 float으로 변환합니다."""
    if value is None:
        return 0.0
    s = str(value).strip().replace(",", "").replace(" ", "").replace("%", "")
    if s == "" or s == "-":
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def has_value(value):
    """빈 칸('', None, 공백)인지 여부. parse_number()는 빈 칸도 0.0으로 반환하므로,
    '입력 안 함'과 '실제로 0'을 구분해야 하는 곳(전환/마크업 등)에서 사용합니다."""
    return value is not None and str(value).strip() != ""


def parse_date(value):
    """
    'YYYY-MM-DD' 문자열 또는 Google Sheets 일련번호(숫자, UNFORMATTED_VALUE로 조회 시 날짜 셀의 형태)를
    date 객체로 변환합니다. 실패 시 None.
    """
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        # Google Sheets/Excel 날짜 일련번호: 1899-12-30을 0일차로 계산
        return date(1899, 12, 30) + timedelta(days=int(value))
    try:
        s = str(value).strip()
        parts = s.split("-")
        if len(parts) != 3:
            return None
        return date(int(parts[0]), int(parts[1]), int(parts[2]))
    except (ValueError, TypeError):
        return None


def format_date_iso(value):
    """parse_date()로 파싱한 뒤 'YYYY-MM-DD' 문자열로 반환합니다 (프론트엔드 응답용, 일련번호도 문자열로 변환). 실패 시 None."""
    d = parse_date(value)
    return d.isoformat() if d else None


def get_fund_code_from_nickname(nickname):
    """펀드 닉네임에서 펀드 코드를 가져옵니다."""
    data = get_sheet_data("FUND_BASICS")
    if not data or len(data) < 1:
        return None

    headers = data[0]
    nickname_idx = get_column_index(headers, "FUND_NICKNAME")
    code_idx = get_column_index(headers, "FUND_CODE")

    if nickname_idx == -1 or code_idx == -1:
        return None

    for row in data[1:]:
        if nickname_idx < len(row) and row[nickname_idx] == nickname:
            if code_idx < len(row):
                return row[code_idx]

    return None


def get_fund_info(fund_code):
    """FUND_BASICS에서 펀드 정보를 dict로 반환합니다."""
    funds = get_sheet_dicts("FUND_BASICS")
    for f in funds:
        if f.get("FUND_CODE") == fund_code:
            return f
    return None


@app.route('/api/funds', methods=['GET'])
def get_funds():
    """펀드 목록을 반환합니다."""
    data = get_sheet_data("FUND_BASICS")

    if not data or len(data) < 1:
        return jsonify([])

    headers = data[0]
    nickname_idx = get_column_index(headers, "FUND_NICKNAME")
    inception_idx = get_column_index(headers, "FUND_INCEPTION")

    if nickname_idx == -1:
        return jsonify([])

    funds = []
    for row_idx in range(1, len(data)):
        row = data[row_idx]
        nickname = row[nickname_idx] if nickname_idx < len(row) else None
        if nickname:
            funds.append({
                "nickname": nickname,
                "row_index": row_idx,
                "inception": format_date_iso(row[inception_idx]) if inception_idx != -1 and inception_idx < len(row) else None,
            })

    return jsonify(funds)


@app.route('/api/currencies', methods=['GET'])
def get_currencies():
    """
    선택된 펀드의 집행 통화, 펀드 통화(FUND_CURRENCY), 대체환율 필요 여부를 반환합니다.

    needs_fx_other: 집행환율(EXECUTED) 선택 시 CALL_ID가 없어 적용할 수 없는 항목
    (FUND_SUBSCRIPTION 등)의 통화가 펀드 통화와 다른 경우 True. Femtech처럼 약정통화(USD)와
    펀드 통화(KRW)가 다른 펀드가 대표적인 예시입니다.
    """
    fund_nickname = request.args.get('fund')

    empty_response = {"currencies": [], "fund_currency": "KRW", "needs_fx_other": False}

    if not fund_nickname:
        return jsonify(empty_response)

    fund_code = get_fund_code_from_nickname(fund_nickname)
    if not fund_code:
        return jsonify(empty_response)

    fund_info = get_fund_info(fund_code)
    fund_currency = (fund_info.get("FUND_CURRENCY") if fund_info else None) or "KRW"

    data = get_sheet_data("DIRECT_INVESTMENT")

    currencies = set()
    if data and len(data) >= 1:
        headers = data[0]
        fund_code_idx = get_column_index(headers, "FUND_CODE")
        currency_idx = get_column_index(headers, "CURRENCY_INV")

        if fund_code_idx != -1 and currency_idx != -1:
            for row in data[1:]:
                if fund_code_idx < len(row) and row[fund_code_idx] == fund_code:
                    if currency_idx < len(row):
                        currency = row[currency_idx].strip() if row[currency_idx] else ""
                        if currency and currency != "KRW":
                            currencies.add(currency)

    # 약정통화(CURRENCY_SUB)가 펀드 통화와 다르면, FUND_SUBSCRIPTION은 CALL_ID가 없어
    # 집행환율로는 환산할 수 없으므로 대체환율 선택이 필요합니다.
    needs_fx_other = fund_currency != "KRW"
    subscription_currency = None
    for s in get_sheet_dicts("FUND_SUBSCRIPTION"):
        if s.get("FUND_CODE") != fund_code:
            continue
        sub_currency = (s.get("CURRENCY_SUB") or "").strip() or fund_currency
        if subscription_currency is None:
            subscription_currency = sub_currency
        if sub_currency != fund_currency:
            needs_fx_other = True
            if sub_currency != "KRW":
                currencies.add(sub_currency)

    return jsonify({
        "currencies": sorted(list(currencies)),
        "fund_currency": fund_currency,
        "needs_fx_other": needs_fx_other,
        "subscription_currency": subscription_currency or fund_currency,
    })


@app.route('/api/fund-base-fx', methods=['GET'])
def get_fund_base_fx():
    """FUND_BASE_FX 데이터를 반환합니다."""
    data = get_sheet_data("FUND_BASE_FX")

    if not data or len(data) < 1:
        return jsonify({})

    result = {}
    for row in data[1:]:
        if len(row) > 1:
            key = f"{row[0]}_{row[1]}"
            result[key] = row

    return jsonify(result)


@app.route('/api/direct-markup', methods=['GET'])
def get_direct_markup():
    """DIRECT_MARKUP 데이터를 반환합니다 (ACTUAL 또는 EXPECTED)."""
    markup_type = request.args.get('type', 'ACTUAL')
    sheet_name = "DIRECT_MARKUP_A" if markup_type == "ACTUAL" else "DIRECT_MARKUP_E"

    data = get_sheet_data(sheet_name)
    return jsonify(data if data else [])


# ============================================================
# Investment Overview
# ============================================================

def _get_base_fx_index():
    """FUND_BASE_FX를 (FUND_CODE, CURRENCY) -> BASE_RATE 인덱스로 캐시합니다 (첫 매칭 행 우선)."""
    if "FUND_BASE_FX" not in _index_cache:
        index = {}
        for row in get_sheet_dicts("FUND_BASE_FX"):
            key = (row.get("FUND_CODE"), row.get("CURRENCY", "").strip())
            if key not in index:
                index[key] = parse_number(row.get("BASE_RATE"))
        _index_cache["FUND_BASE_FX"] = index
    return _index_cache["FUND_BASE_FX"]


def _get_fund_fx_rate_index():
    """FUND_FX를 CALL_ID -> BUY_RATE_FX 인덱스로 캐시합니다 (INITIAL 타입 우선, 없으면 첫 번째 행)."""
    if "FUND_FX" not in _index_cache:
        first_rows = {}
        initial_rows = {}
        for row in get_sheet_dicts("FUND_FX"):
            call_id = row.get("CALL_ID")
            if call_id not in first_rows:
                first_rows[call_id] = row
            if call_id not in initial_rows and row.get("FX_TYPE", "").strip() == "INITIAL":
                initial_rows[call_id] = row
        _index_cache["FUND_FX"] = {
            call_id: parse_number(initial_rows.get(call_id, row).get("BUY_RATE_FX"))
            for call_id, row in first_rows.items()
        }
    return _index_cache["FUND_FX"]


def get_rate_to_krw(currency, fx_option, fx_rates, fund_code, call_id, as_of):
    """
    특정 통화를 KRW로 환산하는 환율을 반환합니다.
    - EXECUTED: FUND_FX 시트에서 CALL_ID로 매칭된 INITIAL 타입의 BUY_RATE_FX 사용
    - BASE: FUND_BASE_FX 시트에서 (FUND_CODE, CURRENCY)로 매칭된 BASE_RATE 사용
    - SPOT: 사용자가 입력한 fx_rates 사용
    KRW는 항상 1을 반환합니다. 환율을 찾을 수 없으면 None을 반환합니다.
    """
    if currency == "KRW":
        return 1.0

    if fx_option == "SPOT":
        rate = fx_rates.get(currency)
        return parse_number(rate) if rate is not None else None

    if fx_option == "BASE":
        return _get_base_fx_index().get((fund_code, currency))

    if fx_option == "EXECUTED":
        return _get_fund_fx_rate_index().get(call_id)

    return None


def convert_amount(amount, native_currency, target_currency, fx_option, fx_rates, fund_code, call_id, as_of, label=None):
    """
    금액을 native_currency에서 target_currency로 환산합니다.
    변환 불가 시 (None, {"currency": 못찾은통화, "label": 어떤 항목인지}) 형태로 반환합니다.
    """
    if native_currency == target_currency:
        return amount, None

    rate_native = get_rate_to_krw(native_currency, fx_option, fx_rates, fund_code, call_id, as_of)
    if rate_native is None:
        return None, {"currency": native_currency, "label": label}

    rate_target = get_rate_to_krw(target_currency, fx_option, fx_rates, fund_code, call_id, as_of)
    if rate_target is None:
        return None, {"currency": target_currency, "label": label}

    # amount(native) -> KRW -> target
    amount_krw = amount * rate_native
    return amount_krw / rate_target, None


def add_fx_warning(warnings, warn):
    """convert_amount가 반환한 {"currency", "label"}을 통화별로 모읍니다."""
    if warn:
        warnings.setdefault(warn["currency"], []).append(warn.get("label"))


def merge_fx_warnings(target, source):
    """compute_investment_rows 등 다른 곳에서 모은 통화별 warnings를 합칩니다."""
    for currency, labels in source.items():
        target.setdefault(currency, []).extend(labels)


def format_fx_warnings(warnings):
    """{"USD": ["A", "B", None]} -> ["USD 환율 부재: A, B"] 형태의 문구로 변환합니다."""
    messages = []
    for currency in sorted(warnings.keys()):
        labels = list(dict.fromkeys(l for l in warnings[currency] if l))
        if labels:
            messages.append(f"{currency} 환율 부재: {', '.join(labels)}")
        else:
            messages.append(f"{currency} 환율을 찾을 수 없습니다")
    return messages


def calculate_xirr(cashflows, guess=0.1):
    """
    cashflows: [(date, amount), ...] (amount는 유출(-)/유입(+)). 뉴턴법으로 XIRR을 계산합니다.
    수렴하지 않으면 None을 반환합니다.
    """
    cashflows = [(d, a) for d, a in cashflows if d is not None]
    if len(cashflows) < 2:
        return None

    t0 = min(d for d, _ in cashflows)
    years = [(d - t0).days / 365.0 for d, _ in cashflows]
    amounts = [a for _, a in cashflows]

    def npv(rate):
        return sum(a / (1 + rate) ** t for a, t in zip(amounts, years))

    def npv_derivative(rate):
        return sum(-t * a / (1 + rate) ** (t + 1) for a, t in zip(amounts, years))

    rate = guess
    for _ in range(100):
        f = npv(rate)
        df = npv_derivative(rate)
        if df == 0:
            return None
        new_rate = rate - f / df
        if new_rate <= -0.999:
            new_rate = -0.999
        if abs(new_rate - rate) < 1e-7:
            return new_rate
        rate = new_rate

    return None


def compute_and_write_holdings(fund_code, as_of):
    """
    선택된 펀드의 as-of 시점(연/월) 기준 현재 보유 주식 현황을 계산하여 DIRECT_HOLDINGS 시트에 씁니다.

    계산 로직 (as_of 이하 날짜의 이벤트만 반영):
    - DIRECT_INVESTMENT: 최초 투자 시점의 종목/주수/통화/단가를 시작점으로 사용
    - DIRECT_CONVERSION: as_of 이하 날짜의 전환 이력을 날짜순으로 순차 적용 (종목/주수/통화/단가를 전환 결과로 갱신)
    - DIRECT_REALIZATION: as_of 이하 날짜의 실현(매각) 주식수 합계를 차감
    - DIRECT_MARKUP_A: as_of 이하 날짜 중 최신 마크업의 단가/통화(CURRENCY_MKA)로 PRICE_PER_SHARE/통화를 갱신
      (주수는 변경하지 않음 - SHARES_MKA는 회사 전체 발행주식수이지 우리 보유주수가 아님). 마크업 통화가 투자/전환
      통화와 다를 수 있으므로 (예: USD 투자건에 KRW로 마크업) 통화도 함께 갱신해야 가격 단위 오류를 막을 수 있음

    쓰기 실패(권한 등)는 경고만 남기고 조용히 무시합니다 (Investment Overview 조회 자체를 막지 않기 위함).
    """
    try:
        investments = get_sheet_dicts("DIRECT_INVESTMENT")

        # CALL_ID별로 한 번에 그룹핑하고 날짜도 미리 파싱해둡니다 (투자 건마다 전체 리스트 재탐색/재파싱 방지)
        conversions_by_call = {}
        for c in get_sheet_dicts("DIRECT_CONVERSION"):
            conv_date = parse_date(c.get("DATE_CONV"))
            if conv_date:
                conversions_by_call.setdefault(c.get("CALL_ID"), []).append((conv_date, c))

        realizations_by_call = {}
        for r in get_sheet_dicts("DIRECT_REALIZATION"):
            realizations_by_call.setdefault(r.get("CALL_ID"), []).append((parse_date(r.get("DATE_REAL")), r))

        markups_by_call = {}
        for m in get_sheet_dicts("DIRECT_MARKUP_A"):
            markup_date = parse_date(m.get("DATE_MKA"))
            if markup_date:
                markups_by_call.setdefault(m.get("CALL_ID"), []).append((markup_date, m))

        fund_investments = [inv for inv in investments if inv.get("FUND_CODE") == fund_code]

        updates = {}
        for inv in fund_investments:
            call_id = inv.get("CALL_ID")
            inv_date = parse_date(inv.get("DATE_INV"))
            if inv_date and inv_date > as_of:
                continue  # 아직 투자가 발생하지 않은 시점

            sec_type = (inv.get("SECURITY_TYPE_INV") or "").strip()
            currency = (inv.get("CURRENCY_INV") or "").strip()
            shares = parse_number(inv.get("SHARES_INV"))
            price = parse_number(inv.get("PRICE_PER_SHARE_INV"))

            call_conversions = sorted(
                (dc for dc in conversions_by_call.get(call_id, []) if dc[0] <= as_of),
                key=lambda dc: dc[0],
            )
            # 전환/마크업 행이 있어도 특정 칸(주수/단가 등)이 비어 있으면 "아직 확정 안 됨"으로 보고
            # 직전 값(투자 단가 등)을 그대로 유지합니다 - 빈 칸을 parse_number()가 0으로 반환해
            # 유효했던 원가를 0으로 덮어써버리는 것을 방지하기 위함입니다.
            for _, conv in call_conversions:
                sec_type = (conv.get("SECURITY_TYPE_CONV") or "").strip() or sec_type
                currency = (conv.get("CURRENCY_CONV") or "").strip() or currency
                if has_value(conv.get("SHARES_CONV")):
                    shares = parse_number(conv.get("SHARES_CONV"))
                if has_value(conv.get("PRICE_PER_SHARE_CONV")):
                    price = parse_number(conv.get("PRICE_PER_SHARE_CONV"))

            realized_shares = sum(
                parse_number(r.get("SHARES_REAL"))
                for real_date, r in realizations_by_call.get(call_id, [])
                if not (real_date and real_date > as_of)
            )
            shares -= realized_shares

            call_markups = [dm for dm in markups_by_call.get(call_id, []) if dm[0] <= as_of]
            if call_markups:
                latest_markup = max(call_markups, key=lambda dm: dm[0])[1]
                currency = (latest_markup.get("CURRENCY_MKA") or "").strip() or currency
                if has_value(latest_markup.get("PRICE_PER_SHARE_MKA")):
                    price = parse_number(latest_markup.get("PRICE_PER_SHARE_MKA"))

            updates[call_id] = (sec_type, shares, currency, price)

        if not updates:
            return

        # 행 번호 매핑은 방금 읽어온 캐시 데이터에서 계산합니다 (쓰기 전 별도 시트 재조회 불필요).
        # cleaned 데이터: index 0 = 시트 1행(헤더), index i>=1 = 시트 (i+2)행 (2행 한글 설명은 스킵됨)
        holdings_data = get_sheet_data("DIRECT_HOLDINGS")
        row_map = {}
        current_values = {}
        for i, row in enumerate(holdings_data[1:], start=1):
            if row and row[0]:
                row_map[row[0]] = i + 2
                current_values[row[0]] = row[1:5]

        def holding_unchanged(call_id, sec_type, shares, currency, price):
            """시트의 현재 B~E 값이 계산 결과와 이미 같으면 True (동일 값 재기록 생략용)."""
            current = current_values.get(call_id)
            if current is None or len(current) < 4:
                return False
            return (
                str(current[0]).strip() == sec_type
                and str(current[2]).strip() == currency
                and math.isclose(parse_number(current[1]), shares, rel_tol=1e-9, abs_tol=1e-9)
                and math.isclose(parse_number(current[3]), price, rel_tol=1e-9, abs_tol=1e-9)
            )

        batch = []
        new_rows = []
        for call_id, (sec_type, shares, currency, price) in updates.items():
            if call_id in row_map:
                if holding_unchanged(call_id, sec_type, shares, currency, price):
                    continue
                row_num = row_map[call_id]
                batch.append({"range": f"B{row_num}:E{row_num}", "values": [[sec_type, shares, currency, price]]})
            else:
                # 이 CALL_ID의 행이 DIRECT_HOLDINGS에 아직 없으면 (신규 투자건) 새 행을 추가합니다.
                # 기존에는 이런 CALL_ID를 그냥 건너뛰어서 보유 현황이 영영 계산되지 않았습니다.
                new_rows.append([call_id, sec_type, shares, currency, price])

        if batch or new_rows:
            ws = get_gspread_worksheet("DIRECT_HOLDINGS")
            if batch:
                ws.batch_update(batch, value_input_option="USER_ENTERED")
            if new_rows:
                ws.append_rows(new_rows, value_input_option="USER_ENTERED")
            invalidate_sheet("DIRECT_HOLDINGS")  # 캐시 무효화 -> 이후 조회 시 최신값 반영
    except Exception as e:
        print(f"[경고] DIRECT_HOLDINGS 갱신 실패 (무시하고 계속 진행): {e}")


@app.route('/api/investment-overview', methods=['POST'])
def investment_overview():
    """
    Investment Overview 테이블 데이터를 계산합니다.

    Request body:
    {
        "year": 2026,
        "month": 5,
        "fund": "군공",
        "fx_option": "SPOT" | "BASE" | "EXECUTED",  // 투자금액 계산 방식 (AMOUNT_INV 환산에만 사용)
        "fx_rates": {"USD": 1350},  // SPOT 환율 (펀드에 집행된 모든 통화에 대해 항상 필요 - REALIZED/UNREALIZED은 항상 SPOT 사용)
        "markup_option": "ACTUAL" | "EXPECTED",
        "display_currency": "NATIVE" | "FUND"   // 투자통화 vs 펀드통화 토글
    }

    참고:
    - 투자금액(AMOUNT_INV)과 투자잔액은 fx_option에 따라 EXECUTED/BASE/SPOT 중 선택한 방식으로 환산됩니다.
    - REALIZED/UNREALIZED은 fx_option과 무관하게 항상 SPOT 환율(fx_rates)로 환산됩니다.
    - UNREALIZED: markup_option=ACTUAL이면 DIRECT_HOLDINGS(SHARES_HELD × PRICE_PER_SHARE_HELD) 기준.
      markup_option=EXPECTED이면 DIRECT_MARKUP_E가 DIRECT_MARKUP_A보다 최신인 CALL_ID에 한해
      POSTVAL_MKE/POSTVAL_INV 배수를 AMOUNT_INV에 적용해 TOTAL_VALUE를 추정(가격 정보가 없는 미확정 라운드이므로).
      그 외에는 ACTUAL과 동일하게 DIRECT_HOLDINGS 기준.
    - 투자잔액: SHARES_HELD × 취득가(전환 이력이 있으면 DIRECT_CONVERSION의 최신 단가, 없으면 DIRECT_INVESTMENT 투자 단가).
    """
    clear_data_caches()  # 확인 버튼 클릭 시마다 Google Sheets에서 최신 데이터를 다시 읽어옴
    params = request.json or {}

    year = int(params.get('year'))
    month = int(params.get('month'))
    fund_nickname = params.get('fund')
    fx_option = params.get('fx_option')
    fx_rates = params.get('fx_rates', {}) or {}
    markup_option = params.get('markup_option', 'ACTUAL')  # ACTUAL | EXPECTED
    display_currency = params.get('display_currency', 'NATIVE')  # NATIVE | FUND

    fund_code = get_fund_code_from_nickname(fund_nickname)
    if not fund_code:
        return jsonify({"error": "펀드를 찾을 수 없습니다"}), 400

    fund_info = get_fund_info(fund_code)
    fund_currency = fund_info.get("FUND_CURRENCY") if fund_info else "KRW"

    # 선택된 연월의 마지막 날짜 (as-of 기준일)
    last_day = calendar.monthrange(year, month)[1]
    as_of = date(year, month, last_day)

    rows, warnings = compute_investment_rows(
        fund_code, fund_currency, as_of, fx_option, fx_rates, markup_option, display_currency
    )

    return jsonify({
        "fund_currency": fund_currency,
        "display_currency": display_currency,
        "as_of": as_of.isoformat(),
        "rows": rows,
        "warnings": format_fx_warnings(warnings),
    })


def compute_investment_rows(fund_code, fund_currency, as_of, fx_option, fx_rates, markup_option, display_currency):
    """
    /api/investment-overview와 /api/fund-cashflow(Fund NAV 합산)가 공유하는 투자 건별 계산 로직입니다.
    (call_id, company_name, amount_inv, realized, unrealized, total_value, moic 등을 담은 rows 리스트와
    warnings 집합을 반환합니다.)
    """
    # as-of 시점 기준 보유 현황을 재계산하여 DIRECT_HOLDINGS 시트에 반영 (확인 버튼 클릭 시마다 갱신)
    compute_and_write_holdings(fund_code, as_of)

    # 데이터 로드
    investments = get_sheet_dicts("DIRECT_INVESTMENT")
    companies = get_sheet_dicts("DIRECT_COMPANY")
    holdings = get_sheet_dicts("DIRECT_HOLDINGS")

    company_name_map = {c.get("COMPANY_ID"): c.get("COMPANY_NAME") for c in companies}
    company_country_map = {c.get("COMPANY_ID"): c.get("COMPANY_COUNTRY") for c in companies}
    # CALL_ID 기준 보유 주식수 (compute_and_write_holdings가 as_of 시점 기준으로 갱신한 값)
    holdings_map = {h.get("CALL_ID"): h for h in holdings if h.get("CALL_ID")}

    # CALL_ID별로 한 번에 그룹핑하고 날짜도 미리 파싱해둡니다 (투자 건마다 전체 리스트 재탐색/재파싱 방지)
    realizations_by_call = {}
    for r in get_sheet_dicts("DIRECT_REALIZATION"):
        realizations_by_call.setdefault(r.get("CALL_ID"), []).append((parse_date(r.get("DATE_REAL")), r))

    conversions_by_call = {}
    for c in get_sheet_dicts("DIRECT_CONVERSION"):
        conv_date = parse_date(c.get("DATE_CONV"))
        if conv_date:
            conversions_by_call.setdefault(c.get("CALL_ID"), []).append((conv_date, c))

    markups_a_by_call = {}
    for m in get_sheet_dicts("DIRECT_MARKUP_A"):
        markup_date = parse_date(m.get("DATE_MKA"))
        if markup_date:
            markups_a_by_call.setdefault(m.get("CALL_ID"), []).append((markup_date, m))

    markups_e_by_call = {}
    for m in get_sheet_dicts("DIRECT_MARKUP_E"):
        markup_date = parse_date(m.get("DATE_MKE"))
        if markup_date:
            markups_e_by_call.setdefault(m.get("CALL_ID"), []).append((markup_date, m))

    # 펀드 + as-of 날짜로 투자 건 필터링
    fund_investments = []
    for inv in investments:
        if inv.get("FUND_CODE") != fund_code:
            continue
        inv_date = parse_date(inv.get("DATE_INV"))
        if inv_date and inv_date > as_of:
            continue
        fund_investments.append(inv)

    rows = []
    warnings = {}

    for inv in fund_investments:
        call_id = inv.get("CALL_ID")
        company_id = inv.get("COMPANY_ID")
        company_label = company_name_map.get(company_id, company_id)
        native_currency = (inv.get("CURRENCY_INV") or "").strip()
        amount_inv = parse_number(inv.get("AMOUNT_INV"))
        shares_inv = parse_number(inv.get("SHARES_INV"))
        post_shares_inv = parse_number(inv.get("POST_SHARES_INV"))
        postval_inv = parse_number(inv.get("POSTVAL_INV"))

        target_currency = native_currency if display_currency == "NATIVE" else fund_currency

        # 투자금액 환산
        converted_amount_inv, warn = convert_amount(
            amount_inv, native_currency, target_currency, fx_option, fx_rates, fund_code, call_id, as_of, label=company_label
        )
        if warn:
            add_fx_warning(warnings, warn)

        # REALIZED: 이 CALL_ID의 실현 내역 합산 (as-of 날짜 이하만)
        realized_total = 0.0
        realized_conversion_failed = False
        for real_date, real in realizations_by_call.get(call_id, []):
            if real_date and real_date > as_of:
                continue
            real_currency = (real.get("CURRENCY_REAL") or "").strip() or native_currency
            real_amount = parse_number(real.get("AMOUNT_REAL"))

            # REALIZED/UNREALIZED은 투자금액 계산 방식(fx_option)과 무관하게 항상 SPOT 환율을 사용합니다
            real_target = real_currency if display_currency == "NATIVE" else fund_currency
            converted_real, real_warn = convert_amount(
                real_amount, real_currency, real_target, "SPOT", fx_rates, fund_code, call_id, as_of, label=company_label
            )
            if real_warn:
                add_fx_warning(warnings, real_warn)
                realized_conversion_failed = True
                continue
            realized_total += converted_real

        # 지분율(투자당시): 투자 시점의 인수주식수(SHARES_INV) / 투자 후 총주식수(POST_SHARES_INV) 기준
        # 지분율(기준일): DIRECT_HOLDINGS의 보유 주식수(SHARES_HELD) 기준 (필터에서 선택한 as-of 시점 기준 최신 보유 현황)
        # (SHARES_INV == 1인 전환형 투자는 아직 주식으로 전환되지 않았으므로 두 경우 모두 투자 당시 추정치로 대체)
        holding = holdings_map.get(call_id)
        shares_held = parse_number(holding.get("SHARES_HELD")) if holding else None
        shares_held = shares_held if shares_held else None

        # 투자잔액: SHARES_HELD × 취득가 (투자 단가에서 시작, DIRECT_CONVERSION의 최신 단가가 있으면 그것으로 갱신.
        # REALIZATION/MARKUP은 반영하지 않음)
        remaining_balance = None
        if shares_held:
            cost_basis_currency = native_currency
            cost_basis_price = parse_number(inv.get("PRICE_PER_SHARE_INV"))

            call_conversions = [dc for dc in conversions_by_call.get(call_id, []) if dc[0] <= as_of]
            if call_conversions:
                latest_conv = max(call_conversions, key=lambda dc: dc[0])[1]
                cost_basis_currency = (latest_conv.get("CURRENCY_CONV") or "").strip() or cost_basis_currency
                # PRICE_PER_SHARE_CONV가 비어 있으면 아직 확정 전이므로 투자 단가를 그대로 유지합니다.
                if has_value(latest_conv.get("PRICE_PER_SHARE_CONV")):
                    cost_basis_price = parse_number(latest_conv.get("PRICE_PER_SHARE_CONV"))

            remaining_target = cost_basis_currency if display_currency == "NATIVE" else fund_currency
            converted_remaining, remaining_warn = convert_amount(
                shares_held * cost_basis_price, cost_basis_currency, remaining_target, fx_option, fx_rates, fund_code, call_id, as_of, label=company_label
            )
            if remaining_warn:
                add_fx_warning(warnings, remaining_warn)
            else:
                remaining_balance = converted_remaining

        # UNREALIZED:
        # - ACTUAL: DIRECT_HOLDINGS(SHARES_HELD × PRICE_PER_SHARE_HELD) 기준 (compute_and_write_holdings가 이미 as-of 반영)
        # - EXPECTED: DIRECT_MARKUP_E가 DIRECT_MARKUP_A보다 최신이면 (미확정 라운드라 주당가격이 없으므로)
        #   POSTVAL_MKE / POSTVAL_INV 배수를 AMOUNT_INV에 적용해 TOTAL_VALUE 추정. 그 외에는 ACTUAL과 동일하게 처리.
        unrealized = None
        use_expected_postval = False
        latest_e_row = None
        if markup_option == "EXPECTED":
            call_markups_a = [dm for dm in markups_a_by_call.get(call_id, []) if dm[0] <= as_of]
            call_markups_e = [dm for dm in markups_e_by_call.get(call_id, []) if dm[0] <= as_of]
            latest_a_date = max((dm[0] for dm in call_markups_a), default=None)
            if call_markups_e:
                latest_e_date, latest_e_row = max(call_markups_e, key=lambda dm: dm[0])
                use_expected_postval = latest_a_date is None or latest_e_date > latest_a_date

        if use_expected_postval:
            # EXPECTED(미확정 라운드) 추정은 POSTVAL_MKE/POSTVAL_INV 배수만 필요하고 실제 보유
            # 주식수(SHARES_HELD)는 쓰지 않으므로, shares_held가 없어도(전환 전 SAFE 등) 계산합니다.
            if postval_inv:
                moic_expected = parse_number(latest_e_row.get("POSTVAL_MKE")) / postval_inv
                total_target = native_currency if display_currency == "NATIVE" else fund_currency
                converted_total, total_warn = convert_amount(
                    moic_expected * amount_inv, native_currency, total_target, "SPOT", fx_rates, fund_code, call_id, as_of, label=company_label
                )
                if total_warn:
                    add_fx_warning(warnings, total_warn)
                elif not realized_conversion_failed:
                    unrealized = converted_total - realized_total
        elif shares_held:
            currency_held = (holding.get("CURRENCY_HELD") or "").strip() or native_currency
            price_held = parse_number(holding.get("PRICE_PER_SHARE_HELD"))
            if price_held:
                unrealized_target = currency_held if display_currency == "NATIVE" else fund_currency
                converted_unrealized, unrealized_warn = convert_amount(
                    shares_held * price_held, currency_held, unrealized_target, "SPOT", fx_rates, fund_code, call_id, as_of, label=company_label
                )
                if unrealized_warn:
                    add_fx_warning(warnings, unrealized_warn)
                else:
                    unrealized = converted_unrealized

        realized_value = None if realized_conversion_failed else realized_total
        total_value = None
        moic = None
        if realized_value is not None and unrealized is not None:
            total_value = realized_value + unrealized
            if converted_amount_inv:
                moic = total_value / converted_amount_inv

        ownership_investment_pct = None
        ownership_investment_estimated = False
        if shares_inv == 1 and postval_inv:
            ownership_investment_pct = (amount_inv / postval_inv) * 100
            ownership_investment_estimated = True
        elif shares_inv and post_shares_inv:
            ownership_investment_pct = (shares_inv / post_shares_inv) * 100

        ownership_asof_pct = None
        ownership_asof_estimated = False
        if shares_inv == 1 and postval_inv:
            ownership_asof_pct = (amount_inv / postval_inv) * 100
            ownership_asof_estimated = True
        elif shares_held and post_shares_inv:
            ownership_asof_pct = (shares_held / post_shares_inv) * 100

        rows.append({
            "call_id": call_id,
            "company_name": company_name_map.get(company_id, company_id),
            "company_country": company_country_map.get(company_id),
            "currency_inv": native_currency,
            "display_currency": target_currency,
            "date_inv": format_date_iso(inv.get("DATE_INV")),
            "amount_inv": converted_amount_inv,
            "realized": realized_value,
            "unrealized": unrealized,
            "remaining_balance": remaining_balance,
            "total_value": total_value,
            "moic": moic,
            "ownership_investment_pct": ownership_investment_pct,
            "ownership_investment_estimated": ownership_investment_estimated,
            "ownership_asof_pct": ownership_asof_pct,
            "ownership_asof_estimated": ownership_asof_estimated,
        })

    return rows, warnings


@app.route('/api/fund-cashflow', methods=['POST'])
def fund_cashflow():
    """
    Fund Cash Flow (LP 관점의 펀드 현금흐름)를 계산합니다.

    Request body: /api/investment-overview와 동일 + fx_option_other (year, month, fund, fx_option, fx_option_other, fx_rates, markup_option)

    - Cash Out(-): DIRECT_INVESTMENT(투자 집행), FUND_EXPENSE(펀드 비용)
    - Cash In(+): FUND_DISTRIBUTION(분배), Fund NAV
      (기준일 기준 모든 투자건의 TOTAL VALUE(REALIZED+UNREALIZED) 합계를 수익자에게 반환한다고 가정하는 마지막 항목.
      REALIZED된 금액도 아직 실제로 수익자에게 분배되지 않았다는 전제 하에 포함합니다.)

    모든 금액은 펀드 통화(FUND_CURRENCY) 기준이며, 선택한 fx_option(EXECUTED/BASE/SPOT)으로 환산합니다.
    단, FUND_EXPENSE/FUND_DISTRIBUTION은 CALL_ID가 없어 EXECUTED를 적용할 수 없으므로,
    fx_option이 EXECUTED이면 fx_option_other(BASE/SPOT)로 대신 환산합니다.
    """
    clear_data_caches()  # 확인 버튼 클릭 시마다 Google Sheets에서 최신 데이터를 다시 읽어옴
    params = request.json or {}

    year = int(params.get('year'))
    month = int(params.get('month'))
    fund_nickname = params.get('fund')
    fx_option = params.get('fx_option')
    fx_rates = params.get('fx_rates', {}) or {}
    markup_option = params.get('markup_option', 'ACTUAL')

    # FUND_EXPENSE/FUND_DISTRIBUTION은 CALL_ID가 없어 집행환율(EXECUTED)을 적용할 수 없으므로,
    # 집행환율 선택 시에는 별도로 지정한 대체 방식(BASE/SPOT)을 사용합니다.
    other_fx_option = params.get('fx_option_other') if fx_option == "EXECUTED" else fx_option

    fund_code = get_fund_code_from_nickname(fund_nickname)
    if not fund_code:
        return jsonify({"error": "펀드를 찾을 수 없습니다"}), 400

    fund_info = get_fund_info(fund_code)
    fund_currency = fund_info.get("FUND_CURRENCY") if fund_info else "KRW"

    last_day = calendar.monthrange(year, month)[1]
    as_of = date(year, month, last_day)

    investments = get_sheet_dicts("DIRECT_INVESTMENT")
    companies = get_sheet_dicts("DIRECT_COMPANY")
    expenses = get_sheet_dicts("FUND_EXPENSE")
    distributions = get_sheet_dicts("FUND_DISTRIBUTION")

    company_name_map = {c.get("COMPANY_ID"): c.get("COMPANY_NAME") for c in companies}

    entries = []
    warnings = {}

    # Cash Out: DIRECT_INVESTMENT (투자 집행)
    for inv in investments:
        if inv.get("FUND_CODE") != fund_code:
            continue
        inv_date = parse_date(inv.get("DATE_INV"))
        if inv_date and inv_date > as_of:
            continue

        native_currency = (inv.get("CURRENCY_INV") or "").strip()
        amount_inv = parse_number(inv.get("AMOUNT_INV"))
        company_label = company_name_map.get(inv.get("COMPANY_ID"), inv.get("COMPANY_ID"))
        converted, warn = convert_amount(
            amount_inv, native_currency, fund_currency, fx_option, fx_rates, fund_code, inv.get("CALL_ID"), as_of, label=company_label
        )
        if warn:
            add_fx_warning(warnings, warn)
            continue

        entries.append({
            "date": format_date_iso(inv.get("DATE_INV")),
            "amount": -converted,
            "note": company_label,
            "type": "investment",
        })

    # Cash Out: FUND_EXPENSE (펀드 비용)
    for exp in expenses:
        if exp.get("FUND_CODE") != fund_code:
            continue
        exp_date = parse_date(exp.get("DATE_EXP"))
        if exp_date and exp_date > as_of:
            continue

        exp_currency = (exp.get("CURRENCY_EXP") or "").strip() or fund_currency
        amount_exp = parse_number(exp.get("AMOUNT_EXP"))
        exp_label = (exp.get("NOTE_EXP") or "").strip() or format_date_iso(exp.get("DATE_EXP")) or "비용"
        converted, warn = convert_amount(
            amount_exp, exp_currency, fund_currency, other_fx_option, fx_rates, fund_code, None, as_of, label=exp_label
        )
        if warn:
            add_fx_warning(warnings, warn)
            continue

        entries.append({
            "date": format_date_iso(exp.get("DATE_EXP")),
            "amount": -converted,
            "note": (exp.get("NOTE_EXP") or "").strip(),
            "type": "expense",
        })

    # Cash In: FUND_DISTRIBUTION (분배)
    for dist in distributions:
        if dist.get("FUND_CODE") != fund_code:
            continue
        dist_date = parse_date(dist.get("DATE_DIST"))
        if dist_date and dist_date > as_of:
            continue

        dist_currency = (dist.get("CURRENCY_DIST") or "").strip() or fund_currency
        amount_dist = parse_number(dist.get("AMOUNT_DIST"))
        dist_label = (dist.get("NOTE_DIST") or "").strip() or format_date_iso(dist.get("DATE_DIST")) or "분배"
        converted, warn = convert_amount(
            amount_dist, dist_currency, fund_currency, other_fx_option, fx_rates, fund_code, None, as_of, label=dist_label
        )
        if warn:
            add_fx_warning(warnings, warn)
            continue

        entries.append({
            "date": format_date_iso(dist.get("DATE_DIST")),
            "amount": converted,
            "note": (dist.get("NOTE_DIST") or "").strip(),
            "type": "distribution",
        })

    # Cash In: Fund NAV (기준일 기준 모든 투자건의 TOTAL VALUE 합계 - 항상 마지막 항목)
    nav_rows, nav_warnings = compute_investment_rows(
        fund_code, fund_currency, as_of, fx_option, fx_rates, markup_option, "FUND"
    )
    merge_fx_warnings(warnings, nav_warnings)
    nav_total = sum(r["total_value"] for r in nav_rows if r.get("total_value") is not None)

    entries.append({
        "date": as_of.isoformat(),
        "amount": nav_total,
        "note": "NAV",
        "type": "nav",
    })

    entries.sort(key=lambda e: e["date"])

    return jsonify({
        "fund_currency": fund_currency,
        "as_of": as_of.isoformat(),
        "rows": entries,
        "warnings": format_fx_warnings(warnings),
    })


@app.route('/api/fund-metrics', methods=['POST'])
def fund_metrics():
    """
    Fund Metrics: BASICS / PERFORMANCE / DEPLOYMENT / INVESTMENT / REALIZATION / HEDGE

    Request body: /api/investment-overview와 동일 (year, month, fund, fx_option, fx_rates, markup_option)

    - 약정액: FUND_SUBSCRIPTION 시트에서 이 펀드의 LP별 AMOUNT_SUB 합계
    - NAV/MOIC/IRR: Investment Overview와 동일한 as-of 계산을 재사용 (compute_investment_rows)
    - IRR: Gross(투자/실현/분배/NAV, 비용 제외)와 Net(비용 포함) 현금흐름을 각각 XIRR로 계산
    - Investment: 총 투자 건수/포트폴리오 기업 수/Minimum·Average·Maximum Ticket/Average Ownership (nav_rows 재사용)
    - Hedge: 통화별 보유 헤지율 = SUM(FUND_FX의 INITIAL 타입 SELL_AMOUNT_FX, 통화별) / SUM(DIRECT_INVESTMENT의 AMOUNT_INV, 통화별).
      정산익/정산손 = FUND_EXPENSE에서 EXPENSE_TYPE이 각각 "SETTLEMENT GAIN"/"SETTLEMENT LOSS"인 항목의 합
    """
    clear_data_caches()
    params = request.json or {}

    year = int(params.get('year'))
    month = int(params.get('month'))
    fund_nickname = params.get('fund')
    fx_option = params.get('fx_option')
    fx_rates = params.get('fx_rates', {}) or {}
    markup_option = params.get('markup_option', 'ACTUAL')
    # 약정(FUND_SUBSCRIPTION)은 CALL_ID가 없어 집행환율을 적용할 수 없으므로, 집행환율 선택 시
    # 대체환율(BASE/SPOT)로 대신 환산합니다 (Fund Cash Flow의 비용/분배 환산과 동일한 방식).
    fx_option_other = params.get('fx_option_other')
    other_fx_option = fx_option_other if fx_option == "EXECUTED" else fx_option

    fund_code = get_fund_code_from_nickname(fund_nickname)
    if not fund_code:
        return jsonify({"error": "펀드를 찾을 수 없습니다"}), 400

    fund_info = get_fund_info(fund_code) or {}
    fund_currency = fund_info.get("FUND_CURRENCY") or "KRW"

    last_day = calendar.monthrange(year, month)[1]
    as_of = date(year, month, last_day)

    warnings = {}

    def to_fund_currency(amount, native_currency, call_id, method="SPOT", label=None):
        converted, warn = convert_amount(amount, native_currency, fund_currency, method, fx_rates, fund_code, call_id, as_of, label=label)
        if warn:
            add_fx_warning(warnings, warn)
            return None
        return converted

    # ---------- BASICS ----------
    target_irr_raw = fund_info.get("FUND_TARGET_IRR")
    target_hedge_raw = fund_info.get("FUND_TARGET_HEDGE_RATE")

    subscriptions = get_sheet_dicts("FUND_SUBSCRIPTION")
    commitment_total = 0.0
    subscription_currency = None
    for s in subscriptions:
        if s.get("FUND_CODE") != fund_code:
            continue
        sub_currency = (s.get("CURRENCY_SUB") or "").strip() or fund_currency
        if subscription_currency is None:
            subscription_currency = sub_currency
        amount_sub = parse_number(s.get("AMOUNT_SUB"))
        converted, warn = convert_amount(
            amount_sub, sub_currency, fund_currency, other_fx_option, fx_rates, fund_code, None, as_of, label="약정액"
        )
        if warn:
            add_fx_warning(warnings, warn)
            continue
        commitment_total += converted

    basics = {
        "currency": fund_currency,
        "subscription_currency": subscription_currency or fund_currency,
        "commitment_amount": commitment_total,
        "target_irr_type": (fund_info.get("FUND_TARGET_IRR_TYPE") or "").strip() or None,
        "target_irr": parse_number(target_irr_raw) * 100 if target_irr_raw not in (None, "") else None,
        "date_subscription": format_date_iso(fund_info.get("FUND_SUBSCRIPTION")),
        "date_inception": format_date_iso(fund_info.get("FUND_INCEPTION")),
        "date_investment_termination": format_date_iso(fund_info.get("FUND_INVESTMENT_TERMINATION")),
        "date_maturity": format_date_iso(fund_info.get("FUND_MATURITY")),
    }

    # ---------- 데이터 로드 ----------
    investments = get_sheet_dicts("DIRECT_INVESTMENT")
    companies = get_sheet_dicts("DIRECT_COMPANY")
    realizations = get_sheet_dicts("DIRECT_REALIZATION")
    expenses = get_sheet_dicts("FUND_EXPENSE")
    distributions = get_sheet_dicts("FUND_DISTRIBUTION")
    fund_fx_rows = get_sheet_dicts("FUND_FX")

    company_name_map = {c.get("COMPANY_ID"): c.get("COMPANY_NAME") for c in companies}

    fund_investments = []
    for inv in investments:
        if inv.get("FUND_CODE") != fund_code:
            continue
        inv_date = parse_date(inv.get("DATE_INV"))
        if inv_date and inv_date > as_of:
            continue
        fund_investments.append(inv)

    # ---------- 투자금 / 펀드비용 / Capital Called ----------
    total_invested = 0.0
    investment_cashflow = []  # (date, -amount) for IRR
    for inv in fund_investments:
        native_currency = (inv.get("CURRENCY_INV") or "").strip()
        amount_inv = parse_number(inv.get("AMOUNT_INV"))
        company_label = company_name_map.get(inv.get("COMPANY_ID"), inv.get("COMPANY_ID"))
        converted, warn = convert_amount(amount_inv, native_currency, fund_currency, fx_option, fx_rates, fund_code, inv.get("CALL_ID"), as_of, label=company_label)
        if warn:
            add_fx_warning(warnings, warn)
            continue
        total_invested += converted
        investment_cashflow.append((parse_date(inv.get("DATE_INV")), -converted))

    total_fund_expense = 0.0
    expense_cashflow = []  # (date, -amount) for Net IRR
    settlement_gain_only = 0.0  # Hedge 카드용: 정산익 = SETTLEMENT GAIN 항목만
    settlement_loss_only = 0.0  # Hedge 카드용: 정산손 = SETTLEMENT LOSS 항목만
    for exp in expenses:
        if exp.get("FUND_CODE") != fund_code:
            continue
        exp_date = parse_date(exp.get("DATE_EXP"))
        if exp_date and exp_date > as_of:
            continue
        exp_currency = (exp.get("CURRENCY_EXP") or "").strip() or fund_currency
        amount_exp = parse_number(exp.get("AMOUNT_EXP"))
        exp_label = (exp.get("NOTE_EXP") or "").strip() or format_date_iso(exp.get("DATE_EXP")) or "비용"
        converted = to_fund_currency(amount_exp, exp_currency, None, label=exp_label)
        if converted is None:
            continue
        total_fund_expense += converted
        expense_cashflow.append((exp_date, -converted))
        expense_type = (exp.get("EXPENSE_TYPE") or "").strip().upper()
        if expense_type == "SETTLEMENT GAIN":
            settlement_gain_only += converted
        elif expense_type == "SETTLEMENT LOSS":
            settlement_loss_only += -converted

    capital_called = total_invested + total_fund_expense

    # ---------- REALIZATION: TXN_TYPE_REAL별 ----------
    realized_by_type = {"장내매각": 0.0, "장외매각": 0.0, "이자수령": 0.0}
    realization_cashflow = []  # (date, +amount) for IRR
    fund_call_ids = {inv.get("CALL_ID") for inv in fund_investments}
    # CALL_ID -> 투자 건 매핑 (첫 번째 매칭 행 우선 - 기존 next() 선형 탐색과 동일)
    inv_by_call_id = {}
    for i in fund_investments:
        inv_by_call_id.setdefault(i.get("CALL_ID"), i)
    for real in realizations:
        if real.get("CALL_ID") not in fund_call_ids:
            continue
        real_date = parse_date(real.get("DATE_REAL"))
        if real_date and real_date > as_of:
            continue
        inv = inv_by_call_id.get(real.get("CALL_ID"))
        native_currency = (inv.get("CURRENCY_INV") or "").strip() if inv else fund_currency
        real_currency = (real.get("CURRENCY_REAL") or "").strip() or native_currency
        real_amount = parse_number(real.get("AMOUNT_REAL"))
        real_label = company_name_map.get(inv.get("COMPANY_ID"), inv.get("COMPANY_ID")) if inv else None
        converted = to_fund_currency(real_amount, real_currency, real.get("CALL_ID"), label=real_label)
        if converted is None:
            continue
        txn_type = (real.get("TXN_TYPE_REAL") or "").strip()
        if txn_type in realized_by_type:
            realized_by_type[txn_type] += converted
        realization_cashflow.append((real_date, converted))

    total_realized = sum(realized_by_type.values())

    # ---------- FUND_DISTRIBUTION ----------
    total_distributions = 0.0
    distribution_cashflow = []
    for dist in distributions:
        if dist.get("FUND_CODE") != fund_code:
            continue
        dist_date = parse_date(dist.get("DATE_DIST"))
        if dist_date and dist_date > as_of:
            continue
        dist_currency = (dist.get("CURRENCY_DIST") or "").strip() or fund_currency
        amount_dist = parse_number(dist.get("AMOUNT_DIST"))
        dist_label = (dist.get("NOTE_DIST") or "").strip() or format_date_iso(dist.get("DATE_DIST")) or "분배"
        converted = to_fund_currency(amount_dist, dist_currency, None, label=dist_label)
        if converted is None:
            continue
        total_distributions += converted
        distribution_cashflow.append((dist_date, converted))

    # ---------- NAV (Investment Overview와 동일 로직 재사용) ----------
    nav_rows, nav_warnings = compute_investment_rows(fund_code, fund_currency, as_of, fx_option, fx_rates, markup_option, "FUND")
    merge_fx_warnings(warnings, nav_warnings)
    nav_total = sum(r["total_value"] for r in nav_rows if r.get("total_value") is not None)

    # ---------- INVESTMENT (건수/티켓 사이즈/지분율 등, nav_rows 재사용 - 이미 펀드통화로 환산되어 있음) ----------
    investment_amounts = [r["amount_inv"] for r in nav_rows if r.get("amount_inv") is not None]
    ownership_values = [r["ownership_asof_pct"] for r in nav_rows if r.get("ownership_asof_pct") is not None]
    portfolio_companies = len({inv.get("COMPANY_ID") for inv in fund_investments})

    investment_metrics = {
        "deal_count": len(fund_investments),
        "portfolio_companies": portfolio_companies,
        "min_ticket_size": min(investment_amounts) if investment_amounts else None,
        "avg_ticket_size": (sum(investment_amounts) / len(investment_amounts)) if investment_amounts else None,
        "max_ticket_size": max(investment_amounts) if investment_amounts else None,
        "avg_ownership": (sum(ownership_values) / len(ownership_values)) if ownership_values else None,
    }

    # ---------- MOIC ----------
    moic_gross = nav_total / total_invested if total_invested else None
    moic_net = (nav_total + total_distributions) / capital_called if capital_called else None

    # ---------- IRR (XIRR) ----------
    gross_cf = investment_cashflow + realization_cashflow + distribution_cashflow + [(as_of, nav_total)]
    net_cf = gross_cf + expense_cashflow
    irr_gross = calculate_xirr(gross_cf)
    irr_net = calculate_xirr(net_cf)

    # ---------- DEPLOYMENT ----------
    deployment = {
        "capital_called": capital_called,
        "invested": total_invested,
        "fund_expense": total_fund_expense,
        "remaining_commitment_excl_fee": (commitment_total - total_invested) if commitment_total else None,
        "remaining_commitment_incl_fee": (commitment_total - capital_called) if commitment_total else None,
        "deployment_rate_excl_fee": (total_invested / commitment_total * 100) if commitment_total else None,
        "deployment_rate_incl_fee": (capital_called / commitment_total * 100) if commitment_total else None,
    }

    # ---------- HEDGE ----------
    foreign_currencies = sorted({
        (inv.get("CURRENCY_INV") or "").strip()
        for inv in fund_investments
        if (inv.get("CURRENCY_INV") or "").strip() and (inv.get("CURRENCY_INV") or "").strip() != fund_currency
    })
    hedge_by_currency = {}
    for currency in foreign_currencies:
        invested_in_currency = sum(
            parse_number(inv.get("AMOUNT_INV"))
            for inv in fund_investments
            if (inv.get("CURRENCY_INV") or "").strip() == currency
        )
        hedged_amount = sum(
            parse_number(fx.get("SELL_AMOUNT_FX"))
            for fx in fund_fx_rows
            if (fx.get("FX_CURRENCY") or "").strip() == currency
            and (fx.get("FX_TYPE") or "").strip() == "INITIAL"
            and fx.get("CALL_ID") in fund_call_ids
        )
        hedge_by_currency[currency] = (hedged_amount / invested_in_currency * 100) if invested_in_currency else None

    return jsonify({
        "fund_currency": fund_currency,
        "as_of": as_of.isoformat(),
        "basics": basics,
        "performance": {
            "nav": nav_total,
            "moic_gross": moic_gross,
            "moic_net": moic_net,
            "irr_gross": irr_gross * 100 if irr_gross is not None else None,
            "irr_net": irr_net * 100 if irr_net is not None else None,
        },
        "deployment": deployment,
        "investment": investment_metrics,
        "realization": {
            "total": total_realized,
            "onmarket": realized_by_type["장내매각"],
            "offmarket": realized_by_type["장외매각"],
            "interest": realized_by_type["이자수령"],
        },
        "hedge": {
            "target_rate": parse_number(target_hedge_raw) * 100 if target_hedge_raw not in (None, "") else None,
            "by_currency": hedge_by_currency,
            "settlement_gain": settlement_gain_only,
            "settlement_loss": settlement_loss_only,
        },
        "warnings": format_fx_warnings(warnings),
    })


@app.route('/api/verify-cashflow', methods=['POST'])
def verify_cashflow():
    """
    "검증" 탭의 "통합자산운용시스템 Cash Flow" 업로드 검증용 기준값을 계산합니다.

    Request body: { "year": 2026, "month": 6, "fund": "군공" }

    기준값 = SUM(DIRECT_INVESTMENT.AMOUNT_INV) + SUM(FUND_EXPENSE.AMOUNT_EXP), 이 펀드(FUND_CODE)의
    as-of(연/월) 이하 날짜 항목만. 투자통화가 펀드 약정통화와 다르면 FUND_FX에서 CALL_ID로 매칭되는
    FX_TYPE="INITIAL" 행의 BUY_RATE_FX로 환산합니다 (= EXECUTED 방식과 동일).
    """
    clear_data_caches()
    params = request.json or {}

    year = int(params.get('year'))
    month = int(params.get('month'))
    fund_nickname = params.get('fund')

    fund_code = get_fund_code_from_nickname(fund_nickname)
    if not fund_code:
        return jsonify({"error": "펀드를 찾을 수 없습니다"}), 400

    fund_info = get_fund_info(fund_code) or {}
    fund_currency = fund_info.get("FUND_CURRENCY") or "KRW"

    last_day = calendar.monthrange(year, month)[1]
    as_of = date(year, month, last_day)

    investments = get_sheet_dicts("DIRECT_INVESTMENT")
    companies = get_sheet_dicts("DIRECT_COMPANY")
    expenses = get_sheet_dicts("FUND_EXPENSE")

    company_name_map = {c.get("COMPANY_ID"): c.get("COMPANY_NAME") for c in companies}

    warnings = {}

    total_invested = 0.0
    for inv in investments:
        if inv.get("FUND_CODE") != fund_code:
            continue
        inv_date = parse_date(inv.get("DATE_INV"))
        if inv_date and inv_date > as_of:
            continue
        native_currency = (inv.get("CURRENCY_INV") or "").strip()
        amount_inv = parse_number(inv.get("AMOUNT_INV"))
        company_label = company_name_map.get(inv.get("COMPANY_ID"), inv.get("COMPANY_ID"))
        converted, warn = convert_amount(
            amount_inv, native_currency, fund_currency, "EXECUTED", {}, fund_code, inv.get("CALL_ID"), as_of, label=company_label
        )
        if warn:
            add_fx_warning(warnings, warn)
            continue
        total_invested += converted

    total_expense = 0.0
    for exp in expenses:
        if exp.get("FUND_CODE") != fund_code:
            continue
        exp_date = parse_date(exp.get("DATE_EXP"))
        if exp_date and exp_date > as_of:
            continue
        exp_currency = (exp.get("CURRENCY_EXP") or "").strip() or fund_currency
        amount_exp = parse_number(exp.get("AMOUNT_EXP"))
        exp_label = (exp.get("NOTE_EXP") or "").strip() or format_date_iso(exp.get("DATE_EXP")) or "비용"
        converted, warn = convert_amount(
            amount_exp, exp_currency, fund_currency, "EXECUTED", {}, fund_code, None, as_of, label=exp_label
        )
        if warn:
            add_fx_warning(warnings, warn)
            continue
        total_expense += converted

    return jsonify({
        "fund_currency": fund_currency,
        "as_of": as_of.isoformat(),
        "total_invested": total_invested,
        "total_expense": total_expense,
        "total_cost": total_invested + total_expense,
        "warnings": format_fx_warnings(warnings),
    })


@app.route('/api/verify-shares', methods=['POST'])
def verify_shares():
    """
    "검증" 탭의 "HINTS 신탁재산명세부" 업로드 검증용 기준값을 계산합니다.

    Request body: { "year": 2026, "month": 5, "fund": "군공" }

    DIRECT_INVESTMENT.HINTS_ID가 "N/A"가 아닌 CALL_ID에 한해, as-of(연/월) 시점 기준으로
    DIRECT_HOLDINGS.SHARES_HELD를 재계산해서 반환합니다 (프론트에서 HINTS_ID로 업로드 파일과 매칭 비교).
    """
    clear_data_caches()
    params = request.json or {}

    year = int(params.get('year'))
    month = int(params.get('month'))
    fund_nickname = params.get('fund')

    fund_code = get_fund_code_from_nickname(fund_nickname)
    if not fund_code:
        return jsonify({"error": "펀드를 찾을 수 없습니다"}), 400

    last_day = calendar.monthrange(year, month)[1]
    as_of = date(year, month, last_day)

    # as-of 시점 기준 보유 현황 재계산 (다른 탭과 동일한 방식)
    compute_and_write_holdings(fund_code, as_of)

    investments = get_sheet_dicts("DIRECT_INVESTMENT")
    companies = get_sheet_dicts("DIRECT_COMPANY")
    holdings = get_sheet_dicts("DIRECT_HOLDINGS")

    company_name_map = {c.get("COMPANY_ID"): c.get("COMPANY_NAME") for c in companies}
    holdings_map = {h.get("CALL_ID"): h for h in holdings if h.get("CALL_ID")}

    rows = []
    for inv in investments:
        if inv.get("FUND_CODE") != fund_code:
            continue
        hints_id = (inv.get("HINTS_ID") or "").strip()
        if not hints_id or hints_id.upper() == "N/A":
            continue
        inv_date = parse_date(inv.get("DATE_INV"))
        if inv_date and inv_date > as_of:
            continue

        call_id = inv.get("CALL_ID")
        holding = holdings_map.get(call_id)
        shares_held = parse_number(holding.get("SHARES_HELD")) if holding and holding.get("SHARES_HELD") not in (None, "") else None

        rows.append({
            "call_id": call_id,
            "company_name": company_name_map.get(inv.get("COMPANY_ID"), inv.get("COMPANY_ID")),
            "hints_id": hints_id,
            "shares_held": shares_held,
        })

    return jsonify({
        "as_of": as_of.isoformat(),
        "rows": rows,
    })


@app.route('/api/calculate', methods=['POST'])
def calculate_returns():
    """
    펀드 수익률을 계산합니다. (준비 중)
    """
    params = request.json

    result = {
        "status": "ready",
        "message": "계산 로직 준비 중",
        "params": params,
    }

    return jsonify(result)


@app.route('/api/health', methods=['GET'])
def health_check():
    """서버 상태 확인"""
    return jsonify({"status": "ok"})


if __name__ == '__main__':
    app.run(debug=True, port=5001, threaded=True)
