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
    "FUND_BASICS", "FUND_BASE_FX", "FUND_FX_BUY", "FUND_SUB",
    "FUND_EXPENSE", "FUND_DIST",
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


def get_investment_shares_and_price(inv):
    """
    DIRECT_INVESTMENT의 두 트랜치(INV1/INV2 - 한 CALL_ID가 두 종류의 증권을 동시에
    인수한 경우)를 합산해 (보유주식수, 가중평균 단가, 종목유형)을 반환합니다.
    현재는 모든 행이 INV1만 채워져 있지만(INV2=0), 두 트랜치가 함께 채워져도
    주식수 가중평균 단가로 자연스럽게 확장됩니다.
    """
    shares1 = parse_number(inv.get("SHARES_INV1"))
    price1 = parse_number(inv.get("PRICE_PER_SHARE_INV1"))
    shares2 = parse_number(inv.get("SHARES_INV2"))
    price2 = parse_number(inv.get("PRICE_PER_SHARE_INV2"))
    shares = shares1 + shares2
    price = (shares1 * price1 + shares2 * price2) / shares if shares else 0.0
    sec_type = (inv.get("SECURITY_TYPE_INV1") or "").strip() or (inv.get("SECURITY_TYPE_INV2") or "").strip()
    return shares, price, sec_type


def resolve_investment_tranches(inv, call_conversions):
    """
    DIRECT_INVESTMENT의 두 트랜치(INV1/INV2)에 DIRECT_CONVERSION 이력을 날짜순으로 적용합니다.
    한 CALL_ID가 서로 다른 두 증권을 동시에 보유하다 그 중 하나만 전환되는 경우
    (예: RCPS 115,384주 + CS 115,384주로 투자했다가 RCPS만 CS로 전환)를 위한 것으로,
    각 전환 행은 SECURITY_TYPE_PRE_CONV와 종목유형이 일치하는 트랜치에만 적용되고
    나머지 트랜치는 그대로 유지됩니다. call_conversions는 이미 as_of 이하로
    필터링/정렬된 (date, row) 튜플 리스트여야 합니다.
    반환: [[sec_type, shares, price, currency], [sec_type, shares, price, currency]] (트랜치 2개)
    """
    base_currency = (inv.get("CURRENCY_INV") or "").strip()
    tranches = [
        [
            (inv.get("SECURITY_TYPE_INV1") or "").strip(),
            parse_number(inv.get("SHARES_INV1")),
            parse_number(inv.get("PRICE_PER_SHARE_INV1")),
            base_currency,
        ],
        [
            (inv.get("SECURITY_TYPE_INV2") or "").strip(),
            parse_number(inv.get("SHARES_INV2")),
            parse_number(inv.get("PRICE_PER_SHARE_INV2")),
            base_currency,
        ],
    ]

    for _, conv in call_conversions:
        pre_type = (conv.get("SECURITY_TYPE_PRE_CONV") or "").strip()
        target = next((t for t in tranches if t[0] == pre_type), None)
        if target is None:
            continue  # 일치하는 트랜치가 없으면(데이터 불일치) 적용하지 않음
        target[0] = (conv.get("SECURITY_TYPE_POST_CONV") or "").strip() or target[0]
        target[3] = (conv.get("CURRENCY_CONV") or "").strip() or target[3]
        # 전환은 현금 재투입이 아니므로 취득원가(총액)를 그대로 보존합니다 - PRICE_PER_SHARE_CONV를
        # 새 취득단가로 쓰지 않고, 주수가 바뀐 만큼만 단가를 재계산합니다 (총원가 = 주수 x 단가 불변).
        # 빈 칸(SHARES_CONV 미기재)이면 아직 확정 전이므로 이전 값을 그대로 유지합니다.
        if has_value(conv.get("SHARES_CONV")):
            new_shares = parse_number(conv.get("SHARES_CONV"))
            if new_shares:
                target[2] = (target[1] * target[2]) / new_shares
            target[1] = new_shares

    return tranches


def combine_tranches(tranches):
    """resolve_investment_tranches의 두 트랜치를 (총 주수, 가중평균 단가, 종목유형, 통화)로 합칩니다."""
    total_shares = sum(t[1] for t in tranches)
    price = (sum(t[1] * t[2] for t in tranches) / total_shares) if total_shares else 0.0
    sec_type = next((t[0] for t in tranches if t[1] and t[0]), "") or next((t[0] for t in tranches if t[0]), "")
    currency = next((t[3] for t in tranches if t[1]), tranches[0][3] if tranches else "")
    return total_shares, price, sec_type, currency


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
    # 집행환율로는 환산할 수 없으므로 대체환율 선택이 필요합니다. Investment Overview는
    # "펀드통화" 모드에서 약정통화를 타겟으로 쓰므로, 약정통화 자체가 KRW가 아니면 그 통화로
    # 실거래 환율이 없는 투자건(KRW로 집행된 건 등)도 대체환율이 필요할 수 있습니다.
    needs_fx_other = fund_currency != "KRW"
    subscription_currency = None
    for s in get_sheet_dicts("FUND_SUB"):
        if s.get("FUND_CODE") != fund_code:
            continue
        sub_currency = (s.get("CURRENCY_SUB") or "").strip() or fund_currency
        if subscription_currency is None:
            subscription_currency = sub_currency
        if sub_currency != fund_currency:
            needs_fx_other = True
        if sub_currency != "KRW":
            needs_fx_other = True
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
    """FUND_FX_BUY를 CALL_ID -> BUY_RATE_FX 인덱스로 캐시합니다 (INITIAL 타입 우선, 없으면 첫 번째 행)."""
    if "FUND_FX_BUY" not in _index_cache:
        first_rows = {}
        initial_rows = {}
        for row in get_sheet_dicts("FUND_FX_BUY"):
            call_id = row.get("CALL_ID")
            if call_id not in first_rows:
                first_rows[call_id] = row
            if call_id not in initial_rows and row.get("FX_TYPE", "").strip() == "INITIAL":
                initial_rows[call_id] = row
        _index_cache["FUND_FX_BUY"] = {
            call_id: parse_number(initial_rows.get(call_id, row).get("BUY_RATE_FX"))
            for call_id, row in first_rows.items()
        }
    return _index_cache["FUND_FX_BUY"]


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


# 통화별 정상 환율 범위 (통화 1단위 = ? KRW). 벗어나면 오탈자(예: 환율 칸에 금액을 잘못 입력)로 의심합니다.
# 프론트엔드의 SPOT 환율 입력창 검증(FX_SANITY_RANGE, index.html)과 동일한 값을 사용합니다.
FX_SANITY_RANGE = {
    "USD": (1000, 1600),
    "EUR": (1100, 1800),
    "GBP": (1300, 2100),
    "JPY": (6, 12),
    "CNY": (150, 250),
}


def is_fx_rate_suspicious(currency, rate):
    """rate가 이 통화의 정상 범위를 벗어나면 True (범위가 정의되지 않은 통화는 항상 False)."""
    range_ = FX_SANITY_RANGE.get(currency)
    if not range_ or rate is None:
        return False
    return rate < range_[0] or rate > range_[1]


def format_rate_sanity_warnings(warnings):
    """{"USD": ["A", "B"]} -> ["USD 환율 확인 필요(비정상 범위): A, B"] 형태의 문구로 변환합니다."""
    messages = []
    for currency in sorted(warnings.keys()):
        labels = list(dict.fromkeys(l for l in warnings[currency] if l))
        if labels:
            messages.append(f"{currency} 환율 확인 필요(비정상 범위): {', '.join(labels)}")
    return messages


def check_fund_fx_rate_sanity(fund_code, call_ids):
    """
    이 펀드에서 실제 쓰이는 집행환율(FUND_FX_BUY.BUY_RATE_FX)과 약정일환율(FUND_BASE_FX.BASE_RATE)이
    통화별 정상 범위를 벗어나면 경고를 생성합니다 - 환율 칸에 금액을 잘못 입력하는 등의 오타를 잡기 위함
    (예: 환율이 7,500,000처럼 들어가 있으면 실제 계산에 그대로 쓰여 투자금액이 터무니없이 커집니다).
    call_ids: 이 펀드에서 실제 사용 중인 CALL_ID 집합 (관련 없는 다른 펀드의 오타까지 보고하지 않기 위함)
    """
    warnings = {}
    for row in get_sheet_dicts("FUND_FX_BUY"):
        call_id = row.get("CALL_ID")
        if call_id not in call_ids:
            continue
        currency = (row.get("FX_CURRENCY_BUY") or "").strip()
        rate = parse_number(row.get("BUY_RATE_FX"))
        if is_fx_rate_suspicious(currency, rate):
            add_fx_warning(warnings, {"currency": currency, "label": f"{call_id}(집행환율 {rate:,.0f})"})

    for row in get_sheet_dicts("FUND_BASE_FX"):
        if row.get("FUND_CODE") != fund_code:
            continue
        currency = (row.get("CURRENCY") or "").strip()
        rate = parse_number(row.get("BASE_RATE"))
        if is_fx_rate_suspicious(currency, rate):
            add_fx_warning(warnings, {"currency": currency, "label": f"약정일환율 {rate:,.0f}"})

    return warnings


def get_subscription_currency(fund_code, fund_currency):
    """이 펀드의 약정통화(FUND_SUBSCRIPTION.CURRENCY_SUB)를 반환합니다. 값이 없으면 펀드통화로 폴백합니다."""
    for s in get_sheet_dicts("FUND_SUB"):
        if s.get("FUND_CODE") != fund_code:
            continue
        sub_currency = (s.get("CURRENCY_SUB") or "").strip()
        if sub_currency:
            return sub_currency
    return fund_currency


def convert_amount_with_fallback(amount, native_currency, target_currency, fx_option, fx_option_other, fx_rates, fund_code, call_id, as_of, label=None):
    """
    fx_option(주로 집행환율)으로 먼저 환산을 시도하고, 실패하면 대체환율(fx_option_other: BASE/SPOT)로
    재시도합니다. CALL_ID는 있지만 그 통화 조합에 대한 실거래(집행) 환율이 없는 경우를 위한 것입니다.
    """
    converted, warn = convert_amount(amount, native_currency, target_currency, fx_option, fx_rates, fund_code, call_id, as_of, label=label)
    if warn and fx_option == "EXECUTED" and fx_option_other and fx_option_other != fx_option:
        converted, warn = convert_amount(amount, native_currency, target_currency, fx_option_other, fx_rates, fund_code, call_id, as_of, label=label)
    return converted, warn


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
    try:
        for _ in range(100):
            f = npv(rate)
            df = npv_derivative(rate)
            if df == 0:
                return None
            new_rate = rate - f / df
            # 뉴턴법이 극단적인 값으로 발산하면(현금흐름이 병적인 경우) 안전한 범위로 clamp합니다
            # (clamp 없이는 (1+rate)**t가 오버플로되어 전체 요청이 500 에러로 죽을 수 있음).
            new_rate = max(-0.999, min(new_rate, 1e6))
            if abs(new_rate - rate) < 1e-7:
                return new_rate
            rate = new_rate
    except (OverflowError, ZeroDivisionError, ValueError):
        return None

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

            call_conversions = sorted(
                (dc for dc in conversions_by_call.get(call_id, []) if dc[0] <= as_of),
                key=lambda dc: dc[0],
            )
            # 두 트랜치(INV1/INV2 - 한 CALL_ID가 서로 다른 두 증권을 동시에 보유한 경우)에
            # 전환 이력을 각각 적용합니다 (SECURITY_TYPE_PRE_CONV가 일치하는 트랜치만 전환됨).
            tranches = resolve_investment_tranches(inv, call_conversions)
            shares, price, sec_type, currency = combine_tranches(tranches)

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
        "fx_option_other": "BASE" | "SPOT",  // fx_option이 EXECUTED인데 실거래 환율이 없는 투자건에 쓸 대체환율
        "fx_rates": {"USD": 1350},  // SPOT 환율 (펀드에 집행된 모든 통화에 대해 항상 필요 - REALIZED/UNREALIZED은 항상 SPOT 사용)
        "markup_option": "ACTUAL" | "EXPECTED",
        "display_currency": "NATIVE" | "FUND"   // 투자통화 vs 펀드통화 토글
    }

    참고:
    - 투자금액(AMOUNT_INV)과 투자잔액은 fx_option에 따라 EXECUTED/BASE/SPOT 중 선택한 방식으로 환산됩니다.
      fx_option이 EXECUTED인데 해당 CALL_ID에 실거래 환율이 없으면 fx_option_other로 재시도합니다.
    - REALIZED/UNREALIZED은 fx_option과 무관하게 항상 SPOT 환율(fx_rates)로 환산됩니다.
    - UNREALIZED: markup_option=ACTUAL이면 DIRECT_HOLDINGS(SHARES_HELD × PRICE_PER_SHARE_HELD) 기준.
      markup_option=EXPECTED이면 DIRECT_MARKUP_E가 DIRECT_MARKUP_A보다 최신인 CALL_ID에 한해
      POSTVAL_MKE/POSTVAL_INV 배수를 AMOUNT_INV에 적용해 TOTAL_VALUE를 추정(가격 정보가 없는 미확정 라운드이므로).
      그 외에는 ACTUAL과 동일하게 DIRECT_HOLDINGS 기준.
    - 투자잔액: SHARES_HELD × 취득가(전환 이력이 있으면 DIRECT_CONVERSION의 최신 단가, 없으면 DIRECT_INVESTMENT 투자 단가).
    - display_currency="FUND"는 내부 회계 통화(FUND_BASICS.FUND_CURRENCY)가 아니라 약정통화
      (FUND_SUBSCRIPTION.CURRENCY_SUB, LP가 실제로 약정한 통화) 기준으로 환산합니다.
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
    # 약정(FUND_SUBSCRIPTION)은 CALL_ID가 없어 집행환율을 적용할 수 없으므로, 집행환율 선택 시
    # 실거래 환율이 없는 투자건은 대체환율(BASE/SPOT)로 대신 환산합니다.
    fx_option_other = params.get('fx_option_other')

    fund_code = get_fund_code_from_nickname(fund_nickname)
    if not fund_code:
        return jsonify({"error": "펀드를 찾을 수 없습니다"}), 400

    fund_info = get_fund_info(fund_code)
    fund_currency = fund_info.get("FUND_CURRENCY") if fund_info else "KRW"
    # Investment Overview는 LP(수익자) 관점이므로 "펀드통화" 모드는 내부 회계 통화가 아니라
    # 약정통화(FUND_SUBSCRIPTION.CURRENCY_SUB) 기준으로 보여줍니다.
    display_fund_currency = get_subscription_currency(fund_code, fund_currency)

    # 선택된 연월의 마지막 날짜 (as-of 기준일)
    last_day = calendar.monthrange(year, month)[1]
    as_of = date(year, month, last_day)

    rows, warnings, rate_warnings = compute_investment_rows(
        fund_code, display_fund_currency, as_of, fx_option, fx_rates, markup_option, display_currency, fx_option_other=fx_option_other
    )
    fof_rows, fof_warnings = compute_fof_investment_rows(
        fund_code, display_fund_currency, as_of, fx_option, fx_rates, display_currency, fx_option_other=fx_option_other
    )
    merge_fx_warnings(warnings, fof_warnings)

    return jsonify({
        "fund_currency": display_fund_currency,
        "display_currency": display_currency,
        "as_of": as_of.isoformat(),
        "rows": rows,
        "fof_rows": fof_rows,
        "warnings": format_fx_warnings(warnings) + format_rate_sanity_warnings(rate_warnings),
    })


def compute_investment_rows(fund_code, fund_currency, as_of, fx_option, fx_rates, markup_option, display_currency, fx_option_other=None):
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
    rate_warnings = check_fund_fx_rate_sanity(fund_code, {inv.get("CALL_ID") for inv in fund_investments})

    for inv in fund_investments:
        call_id = inv.get("CALL_ID")
        company_id = inv.get("COMPANY_ID")
        company_label = company_name_map.get(company_id, company_id)
        native_currency = (inv.get("CURRENCY_INV") or "").strip()
        amount_inv = parse_number(inv.get("AMOUNT_INV"))
        shares_inv, _, _ = get_investment_shares_and_price(inv)
        post_shares_inv = parse_number(inv.get("POST_SHARES_INV"))
        postval_inv = parse_number(inv.get("POSTVAL_INV"))

        target_currency = native_currency if display_currency == "NATIVE" else fund_currency

        # 투자금액 환산 (집행환율에 실거래 환율이 없으면 대체환율로 재시도)
        converted_amount_inv, warn = convert_amount_with_fallback(
            amount_inv, native_currency, target_currency, fx_option, fx_option_other, fx_rates, fund_code, call_id, as_of, label=company_label
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
            call_conversions = sorted(
                (dc for dc in conversions_by_call.get(call_id, []) if dc[0] <= as_of),
                key=lambda dc: dc[0],
            )
            tranches = resolve_investment_tranches(inv, call_conversions)
            _, cost_basis_price, _, cost_basis_currency = combine_tranches(tranches)

            remaining_target = cost_basis_currency if display_currency == "NATIVE" else fund_currency
            converted_remaining, remaining_warn = convert_amount_with_fallback(
                shares_held * cost_basis_price, cost_basis_currency, remaining_target, fx_option, fx_option_other, fx_rates, fund_code, call_id, as_of, label=company_label
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

    return rows, warnings, rate_warnings


def compute_fof_investment_rows(fund_code, fund_currency, as_of, fx_option, fx_rates, display_currency, fx_option_other=None):
    """
    재간접(FoF) 투자 건별 계산 로직입니다 (예: 벤2가 DCVC V, L.P.에 투자한 경우).
    FOF_FUND=투자 대상 재간접 펀드 정보(DIRECT_COMPANY 역할), FOF_INVESTMENT=이 펀드의 투자건
    (DIRECT_INVESTMENT 역할), FOF_CALL=재간접 펀드에서 발생한 Capital Call, FOF_DIST=Distribution,
    FOF_CAS=해외 펀드가 발행한 Capital Account Statement(평가금액, 우리가 콜한 금액에 대한 평가).

    - 투자금액(Capital Called) = as_of 이하 FOF_CALL 누적
    - 투자잔액(원가 기준) = 투자금액 - as_of 이하 RETURN_OF_CAPITAL 누적
    - REALIZED = as_of 이하 TOTAL_DISTRIBUTION 누적 (원본상환+이익분배금+재투자 전체)
    - UNREALIZED = as_of 이하 최신 FOF_CAS 평가금액 + (그 CAS 기준일 이후 ~ as_of에 발생한 콜은 아직
      CAS에 반영되지 않았으므로 원가로 가산). CAS가 아예 없으면 콜한 원가 전체를 평가금액으로 봄.
    - TOTAL VALUE = REALIZED + UNREALIZED, MOIC = TOTAL VALUE / 투자금액
    """
    fof_fund_map = {f.get("FOF_CODE"): f for f in get_sheet_dicts("FOF_FUND")}

    calls_by_inv = {}
    for c in get_sheet_dicts("FOF_CALL"):
        d = parse_date(c.get("DATE_CALL"))
        if d:
            calls_by_inv.setdefault(c.get("FOF_INV_ID"), []).append((d, c))

    dists_by_inv = {}
    for d_row in get_sheet_dicts("FOF_DIST"):
        d = parse_date(d_row.get("DATE_DIST"))
        if d:
            dists_by_inv.setdefault(d_row.get("FOF_INV_ID"), []).append((d, d_row))

    cas_by_inv = {}
    for c in get_sheet_dicts("FOF_CAS"):
        d = parse_date(c.get("DATE_CAS"))
        if d:
            cas_by_inv.setdefault(c.get("FOF_INV_ID"), []).append((d, c))

    fund_fof_investments = [inv for inv in get_sheet_dicts("FOF_INVESTMENT") if inv.get("FUND_CODE") == fund_code]

    rows = []
    warnings = {}

    for inv in fund_fof_investments:
        fof_inv_id = inv.get("FOF_INV_ID")
        fof_fund = fof_fund_map.get(inv.get("FOF_ID"), {})
        fund_label = fof_fund.get("FOF_NAME") or inv.get("FOF_ID")
        native_currency = (inv.get("FOF_INV_CURRENCY") or "").strip()
        target_currency = native_currency if display_currency == "NATIVE" else fund_currency

        call_list = sorted((dc for dc in calls_by_inv.get(fof_inv_id, []) if dc[0] <= as_of), key=lambda dc: dc[0])
        if not call_list:
            continue  # as_of 시점에 아직 콜이 발생하지 않음 (투자 시작 전)

        date_inv = call_list[0][0]

        # 투자금액(Capital Called): FOF_CALL은 CALL_ID가 없어 집행환율을 적용할 수 없으므로 대체환율 사용
        amount_called = 0.0
        for call_date, call_row in call_list:
            call_currency = (call_row.get("CURRENCY_CALL") or "").strip() or native_currency
            call_amount = parse_number(call_row.get("AMOUNT_CALL"))
            converted, warn = convert_amount_with_fallback(
                call_amount, call_currency, target_currency, fx_option, fx_option_other, fx_rates, fund_code, None, as_of, label=fund_label
            )
            if warn:
                add_fx_warning(warnings, warn)
                continue
            amount_called += converted

        # 약정금액
        commitment_amount, commit_warn = convert_amount_with_fallback(
            parse_number(inv.get("FOF_INV_AMT")), native_currency, target_currency, fx_option, fx_option_other, fx_rates, fund_code, None, as_of, label=fund_label
        )
        if commit_warn:
            add_fx_warning(warnings, commit_warn)
            commitment_amount = None

        # 분배 내역 (REALIZED = TOTAL_DISTRIBUTION 누적, 투자잔액용 RETURN_OF_CAPITAL 누적)
        # REALIZED/UNREALIZED은 투자금액 계산 방식과 무관하게 항상 SPOT 환율을 사용합니다 (직투 로직과 동일)
        dist_list = [(d, r) for d, r in dists_by_inv.get(fof_inv_id, []) if d <= as_of]
        realized_total = 0.0
        return_of_capital_total = 0.0
        dist_conversion_failed = False
        for dist_date, dist_row in dist_list:
            dist_currency = (dist_row.get("CURRENCY_DIST") or "").strip() or native_currency
            real_target = dist_currency if display_currency == "NATIVE" else fund_currency

            converted_total, warn = convert_amount(
                parse_number(dist_row.get("TOTAL_DISTRIBUTION")), dist_currency, real_target, "SPOT", fx_rates, fund_code, None, as_of, label=fund_label
            )
            if warn:
                add_fx_warning(warnings, warn)
                dist_conversion_failed = True
                continue
            realized_total += converted_total

            converted_roc, roc_warn = convert_amount(
                parse_number(dist_row.get("RETURN_OF_CAPITAL")), dist_currency, real_target, "SPOT", fx_rates, fund_code, None, as_of, label=fund_label
            )
            if not roc_warn:
                return_of_capital_total += converted_roc

        # 투자잔액 = 투자금액(원가) - 회수된 원금(RETURN_OF_CAPITAL)
        remaining_balance = amount_called - return_of_capital_total

        # UNREALIZED: 최신 CAS 평가금액 + (그 이후 ~ as_of까지 발생한 콜은 아직 미반영이므로 원가로 가산)
        cas_list = sorted((dc for dc in cas_by_inv.get(fof_inv_id, []) if dc[0] <= as_of), key=lambda dc: dc[0])
        unrealized = None
        if cas_list:
            latest_cas_date, latest_cas_row = cas_list[-1]
            cas_currency = (latest_cas_row.get("CURRENCY_CAS") or "").strip() or native_currency
            cas_amount = parse_number(latest_cas_row.get("AMOUNT_CAS"))
            uncalled_since_cas = sum(
                parse_number(c.get("AMOUNT_CALL")) for call_date, c in call_list if call_date > latest_cas_date
            )
            unrealized_target = cas_currency if display_currency == "NATIVE" else fund_currency
            converted_unrealized, unrealized_warn = convert_amount(
                cas_amount + uncalled_since_cas, cas_currency, unrealized_target, "SPOT", fx_rates, fund_code, None, as_of, label=fund_label
            )
            if unrealized_warn:
                add_fx_warning(warnings, unrealized_warn)
            else:
                unrealized = converted_unrealized
        else:
            # CAS가 아직 한 번도 없으면 콜한 원가 전체를 평가금액으로 봅니다.
            unrealized_target = native_currency if display_currency == "NATIVE" else fund_currency
            converted_unrealized, unrealized_warn = convert_amount_with_fallback(
                amount_called, native_currency, unrealized_target, fx_option, fx_option_other, fx_rates, fund_code, None, as_of, label=fund_label
            )
            if unrealized_warn:
                add_fx_warning(warnings, unrealized_warn)
            else:
                unrealized = converted_unrealized

        total_value = None
        moic = None
        if not dist_conversion_failed and unrealized is not None:
            total_value = realized_total + unrealized
            if amount_called:
                moic = total_value / amount_called

        rows.append({
            "fof_inv_id": fof_inv_id,
            "fund_name": fund_label,
            "country": fof_fund.get("FOF_COUNTRY"),
            "currency_inv": native_currency,
            "display_currency": target_currency,
            "date_inv": date_inv.isoformat(),
            "commitment_amount": commitment_amount,
            "amount_called": amount_called,
            "remaining_balance": remaining_balance,
            "realized": None if dist_conversion_failed else realized_total,
            "unrealized": unrealized,
            "total_value": total_value,
            "moic": moic,
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
    fx_option_other = params.get('fx_option_other')
    other_fx_option = fx_option_other if fx_option == "EXECUTED" else fx_option

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
    distributions = get_sheet_dicts("FUND_DIST")

    company_name_map = {c.get("COMPANY_ID"): c.get("COMPANY_NAME") for c in companies}

    entries = []
    warnings = {}
    rate_warnings = {}

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

    # Cash Out: FOF_CALL 대응 FUND_FX_BUY(FX_TYPE=INITIAL) - 재간접 펀드 Capital Call 시 실제로
    # 집행된 KRW 금액(BUY_KRW)을 그대로 씁니다. FOF_CALL.AMOUNT_CALL(외화 표시 콜 금액)을 별도
    # 환율로 재계산하면 실제 집행 당시 환율과 달라지므로, 실거래 기록인 FUND_FX_BUY를 사용합니다.
    fof_fund_map = {f.get("FOF_CODE"): f for f in get_sheet_dicts("FOF_FUND")}
    fof_inv_map = {i.get("FOF_INV_ID"): i for i in get_sheet_dicts("FOF_INVESTMENT") if i.get("FUND_CODE") == fund_code}
    fof_call_to_inv = {
        c.get("CALL_ID"): c.get("FOF_INV_ID")
        for c in get_sheet_dicts("FOF_CALL")
        if c.get("FOF_INV_ID") in fof_inv_map
    }
    for fx_row in get_sheet_dicts("FUND_FX_BUY"):
        if (fx_row.get("FX_TYPE") or "").strip() != "INITIAL":
            continue
        fof_inv_id = fof_call_to_inv.get(fx_row.get("CALL_ID"))
        if not fof_inv_id:
            continue
        buy_date = parse_date(fx_row.get("BUY_DATE_FX"))
        if buy_date and buy_date > as_of:
            continue

        fof_inv = fof_inv_map[fof_inv_id]
        fof_fund = fof_fund_map.get(fof_inv.get("FOF_ID"), {})
        fund_label = fof_fund.get("FOF_NAME") or fof_inv.get("FOF_ID")
        buy_krw = parse_number(fx_row.get("BUY_KRW"))
        converted, warn = convert_amount(buy_krw, "KRW", fund_currency, "SPOT", fx_rates, fund_code, None, as_of, label=fund_label)
        if warn:
            add_fx_warning(warnings, warn)
            continue

        entries.append({
            "date": format_date_iso(fx_row.get("BUY_DATE_FX")),
            "amount": -converted,
            "note": fund_label,
            "type": "fof_call",
        })

    # Cash In: FOF_DIST (재간접 펀드 분배 - CALL_ID가 없어 항상 대체환율/SPOT 사용)
    for dist_row in get_sheet_dicts("FOF_DIST"):
        fof_inv = fof_inv_map.get(dist_row.get("FOF_INV_ID"))
        if not fof_inv:
            continue
        dist_date = parse_date(dist_row.get("DATE_DIST"))
        if dist_date and dist_date > as_of:
            continue

        fof_fund = fof_fund_map.get(fof_inv.get("FOF_ID"), {})
        fund_label = fof_fund.get("FOF_NAME") or fof_inv.get("FOF_ID")
        dist_currency = (dist_row.get("CURRENCY_DIST") or "").strip() or (fof_inv.get("FOF_INV_CURRENCY") or "").strip()
        amount_dist = parse_number(dist_row.get("TOTAL_DISTRIBUTION"))
        converted, warn = convert_amount_with_fallback(
            amount_dist, dist_currency, fund_currency, fx_option, fx_option_other, fx_rates, fund_code, None, as_of, label=fund_label
        )
        if warn:
            add_fx_warning(warnings, warn)
            continue

        entries.append({
            "date": format_date_iso(dist_row.get("DATE_DIST")),
            "amount": converted,
            "note": fund_label,
            "type": "fof_distribution",
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

    # Cash In: Fund NAV (기준일 기준 모든 투자건의 TOTAL VALUE 합계 - 항상 마지막 항목, 직투+재간접 합산)
    nav_rows, nav_warnings, nav_rate_warnings = compute_investment_rows(
        fund_code, fund_currency, as_of, fx_option, fx_rates, markup_option, "FUND", fx_option_other=fx_option_other
    )
    merge_fx_warnings(warnings, nav_warnings)
    merge_fx_warnings(rate_warnings, nav_rate_warnings)
    nav_total = sum(r["total_value"] for r in nav_rows if r.get("total_value") is not None)

    fof_nav_rows, fof_nav_warnings = compute_fof_investment_rows(
        fund_code, fund_currency, as_of, fx_option, fx_rates, "FUND", fx_option_other=fx_option_other
    )
    merge_fx_warnings(warnings, fof_nav_warnings)
    nav_total += sum(r["total_value"] for r in fof_nav_rows if r.get("total_value") is not None)

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
        "warnings": format_fx_warnings(warnings) + format_rate_sanity_warnings(rate_warnings),
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
    rate_warnings = {}

    def to_fund_currency(amount, native_currency, call_id, method="SPOT", label=None):
        converted, warn = convert_amount(amount, native_currency, fund_currency, method, fx_rates, fund_code, call_id, as_of, label=label)
        if warn:
            add_fx_warning(warnings, warn)
            return None
        return converted

    # ---------- BASICS ----------
    target_irr_raw = fund_info.get("FUND_TARGET_IRR")
    target_hedge_raw = fund_info.get("FUND_TARGET_HEDGE_RATE")

    subscriptions = get_sheet_dicts("FUND_SUB")
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
    distributions = get_sheet_dicts("FUND_DIST")
    fund_fx_rows = get_sheet_dicts("FUND_FX_BUY")

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

    # ---------- FOF_CALL 대응 FUND_FX_BUY(FX_TYPE=INITIAL) - 재간접 펀드 Capital Call 시 실제로
    # 집행된 KRW 금액(BUY_KRW)을 그대로 합산합니다 (Fund Cash Flow와 동일한 방식 - 실거래 기록 우선). ----------
    fof_fund_map = {f.get("FOF_CODE"): f for f in get_sheet_dicts("FOF_FUND")}
    fof_inv_map = {i.get("FOF_INV_ID"): i for i in get_sheet_dicts("FOF_INVESTMENT") if i.get("FUND_CODE") == fund_code}
    fof_call_to_inv = {
        c.get("CALL_ID"): c.get("FOF_INV_ID")
        for c in get_sheet_dicts("FOF_CALL")
        if c.get("FOF_INV_ID") in fof_inv_map
    }
    total_fof_called = 0.0
    for fx_row in get_sheet_dicts("FUND_FX_BUY"):
        if (fx_row.get("FX_TYPE") or "").strip() != "INITIAL":
            continue
        fof_inv_id = fof_call_to_inv.get(fx_row.get("CALL_ID"))
        if not fof_inv_id:
            continue
        buy_date = parse_date(fx_row.get("BUY_DATE_FX"))
        if buy_date and buy_date > as_of:
            continue
        fof_inv = fof_inv_map[fof_inv_id]
        fof_fund = fof_fund_map.get(fof_inv.get("FOF_ID"), {})
        fund_label = fof_fund.get("FOF_NAME") or fof_inv.get("FOF_ID")
        converted = to_fund_currency(parse_number(fx_row.get("BUY_KRW")), "KRW", None, label=fund_label)
        if converted is None:
            continue
        total_fof_called += converted
        investment_cashflow.append((buy_date, -converted))

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

    capital_called = total_invested + total_fof_called + total_fund_expense

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

    # ---------- FOF_DIST (재간접 펀드 분배) ----------
    for dist_row in get_sheet_dicts("FOF_DIST"):
        fof_inv = fof_inv_map.get(dist_row.get("FOF_INV_ID"))
        if not fof_inv:
            continue
        dist_date = parse_date(dist_row.get("DATE_DIST"))
        if dist_date and dist_date > as_of:
            continue
        fof_fund = fof_fund_map.get(fof_inv.get("FOF_ID"), {})
        fund_label = fof_fund.get("FOF_NAME") or fof_inv.get("FOF_ID")
        dist_currency = (dist_row.get("CURRENCY_DIST") or "").strip() or (fof_inv.get("FOF_INV_CURRENCY") or "").strip()
        converted = to_fund_currency(parse_number(dist_row.get("TOTAL_DISTRIBUTION")), dist_currency, None, label=fund_label)
        if converted is None:
            continue
        total_distributions += converted
        distribution_cashflow.append((dist_date, converted))

    # ---------- NAV (Investment Overview와 동일 로직 재사용, 직투+재간접 합산) ----------
    nav_rows, nav_warnings, nav_rate_warnings = compute_investment_rows(fund_code, fund_currency, as_of, fx_option, fx_rates, markup_option, "FUND", fx_option_other=fx_option_other)
    merge_fx_warnings(warnings, nav_warnings)
    merge_fx_warnings(rate_warnings, nav_rate_warnings)
    nav_total = sum(r["total_value"] for r in nav_rows if r.get("total_value") is not None)

    fof_nav_rows, fof_nav_warnings = compute_fof_investment_rows(fund_code, fund_currency, as_of, fx_option, fx_rates, "FUND", fx_option_other=fx_option_other)
    merge_fx_warnings(warnings, fof_nav_warnings)
    nav_total += sum(r["total_value"] for r in fof_nav_rows if r.get("total_value") is not None)

    # ---------- INVESTMENT (건수/티켓 사이즈/지분율 등, nav_rows 재사용 - 이미 펀드통화로 환산되어 있음) ----------
    # 지분율/포트폴리오 기업 수는 직투에만 있는 개념이라 재간접 투자는 제외하고, 건수/티켓사이즈는 합산합니다.
    investment_amounts = [r["amount_inv"] for r in nav_rows if r.get("amount_inv") is not None]
    investment_amounts += [r["amount_called"] for r in fof_nav_rows if r.get("amount_called") is not None]
    ownership_values = [r["ownership_asof_pct"] for r in nav_rows if r.get("ownership_asof_pct") is not None]
    portfolio_companies = len({inv.get("COMPANY_ID") for inv in fund_investments})

    investment_metrics = {
        "deal_count": len(fund_investments) + len(fof_nav_rows),
        "portfolio_companies": portfolio_companies,
        "min_ticket_size": min(investment_amounts) if investment_amounts else None,
        "avg_ticket_size": (sum(investment_amounts) / len(investment_amounts)) if investment_amounts else None,
        "max_ticket_size": max(investment_amounts) if investment_amounts else None,
        "avg_ownership": (sum(ownership_values) / len(ownership_values)) if ownership_values else None,
    }

    # ---------- MOIC ----------
    moic_gross = nav_total / (total_invested + total_fof_called) if (total_invested + total_fof_called) else None
    moic_net = (nav_total + total_distributions) / capital_called if capital_called else None

    # ---------- IRR (XIRR) ----------
    gross_cf = investment_cashflow + realization_cashflow + distribution_cashflow + [(as_of, nav_total)]
    net_cf = gross_cf + expense_cashflow
    irr_gross = calculate_xirr(gross_cf)
    irr_net = calculate_xirr(net_cf)

    # ---------- DEPLOYMENT (직투+재간접 Capital Call 합산) ----------
    invested_total = total_invested + total_fof_called
    deployment = {
        "capital_called": capital_called,
        "invested": invested_total,
        "fund_expense": total_fund_expense,
        "remaining_commitment_excl_fee": (commitment_total - invested_total) if commitment_total else None,
        "remaining_commitment_incl_fee": (commitment_total - capital_called) if commitment_total else None,
        "deployment_rate_excl_fee": (invested_total / commitment_total * 100) if commitment_total else None,
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
            if (fx.get("FX_CURRENCY_BUY") or "").strip() == currency
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
        "warnings": format_fx_warnings(warnings) + format_rate_sanity_warnings(rate_warnings),
    })


@app.route('/api/verify-cashflow', methods=['POST'])
def verify_cashflow():
    """
    "검증" 탭의 "통합자산운용시스템 Cash Flow" 업로드 검증용 기준값을 계산합니다.

    Request body: { "year": 2026, "month": 6, "fund": "군공" }

    수익자에게 요청한 돈(캐피탈콜) = SUM(DIRECT_INVESTMENT.AMOUNT_INV) + SUM(FUND_EXPENSE.AMOUNT_EXP)
    + SUM(재간접 펀드 Capital Call 실제 집행 KRW), 이 펀드(FUND_CODE)의 as-of(연/월) 이하 날짜 항목만.
    투자통화가 펀드 약정통화와 다르면 FUND_FX_BUY에서 CALL_ID로 매칭되는 FX_TYPE="INITIAL" 행의
    BUY_RATE_FX로 환산합니다(= EXECUTED 방식과 동일). 재간접(FOF_CALL) 항목은 FUND_FX_BUY의 INITIAL
    행에 이미 기록된 실제 집행 KRW(BUY_KRW)를 그대로 씁니다 (Fund Cash Flow/Fund Metrics와 동일한 방식).

    수익자에게 나간 돈(분배) = FUND_DIST를 DISTRIBUTION_TYPE별로 집계 - "RETURN OF CAPITAL"은
    return_of_capital(업로드 파일의 "원본(해지)"과 대조), "REALIZED GAIN"은 realized_gain(업로드
    파일의 "이익분배금"과 대조)으로 각각 반환합니다.
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

    # 재간접(FOF_CALL) - FUND_FX_BUY의 INITIAL 행에 기록된 실제 집행 KRW(BUY_KRW)를 그대로 합산
    fof_inv_map = {i.get("FOF_INV_ID"): i for i in get_sheet_dicts("FOF_INVESTMENT") if i.get("FUND_CODE") == fund_code}
    fof_call_to_inv = {
        c.get("CALL_ID"): c.get("FOF_INV_ID")
        for c in get_sheet_dicts("FOF_CALL")
        if c.get("FOF_INV_ID") in fof_inv_map
    }
    total_fof_invested = 0.0
    for fx_row in get_sheet_dicts("FUND_FX_BUY"):
        if (fx_row.get("FX_TYPE") or "").strip() != "INITIAL":
            continue
        if fx_row.get("CALL_ID") not in fof_call_to_inv:
            continue
        buy_date = parse_date(fx_row.get("BUY_DATE_FX"))
        if buy_date and buy_date > as_of:
            continue
        converted, warn = convert_amount(
            parse_number(fx_row.get("BUY_KRW")), "KRW", fund_currency, "SPOT", {}, fund_code, None, as_of, label=fx_row.get("CALL_ID")
        )
        if warn:
            add_fx_warning(warnings, warn)
            continue
        total_fof_invested += converted

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

    # 수익자에게 나간 돈(FUND_DIST) - DISTRIBUTION_TYPE별로 원본상환/이익분배금을 구분 집계합니다.
    # 업로드 파일의 "원본(해지)"/"이익분배금" 컬럼과 각각 대조합니다.
    return_of_capital_total = 0.0
    realized_gain_total = 0.0
    for dist in get_sheet_dicts("FUND_DIST"):
        if dist.get("FUND_CODE") != fund_code:
            continue
        dist_date = parse_date(dist.get("DATE_DIST"))
        if dist_date and dist_date > as_of:
            continue
        dist_type = (dist.get("DISTRIBUTION_TYPE") or "").strip().upper()
        if dist_type not in ("RETURN OF CAPITAL", "REALIZED GAIN"):
            continue
        dist_currency = (dist.get("CURRENCY_DIST") or "").strip() or fund_currency
        dist_label = (dist.get("NOTE_DIST") or "").strip() or format_date_iso(dist.get("DATE_DIST")) or "분배"
        converted, warn = convert_amount(
            parse_number(dist.get("AMOUNT_DIST")), dist_currency, fund_currency, "EXECUTED", {}, fund_code, None, as_of, label=dist_label
        )
        if warn:
            add_fx_warning(warnings, warn)
            continue
        if dist_type == "RETURN OF CAPITAL":
            return_of_capital_total += converted
        else:
            realized_gain_total += converted

    return jsonify({
        "fund_currency": fund_currency,
        "as_of": as_of.isoformat(),
        "total_invested": total_invested + total_fof_invested,
        "total_expense": total_expense,
        "total_cost": total_invested + total_fof_invested + total_expense,
        "return_of_capital": return_of_capital_total,
        "realized_gain": realized_gain_total,
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
