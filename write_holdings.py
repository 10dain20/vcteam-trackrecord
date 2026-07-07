"""
DIRECT_HOLDINGS 시트에 현재 보유 주식 현황을 계산하여 직접 기록하는 스크립트.

계산 로직:
- DIRECT_INVESTMENT의 최초 투자 내역(종목/주수/통화/단가)을 시작점으로 삼음
- DIRECT_CONVERSION이 있으면 SHARES_TARGET -> SHARES_CONV로 주식수/종목/단가 갱신 (RATIO_CONV 순차 적용, 최신 전환 기준)
- DIRECT_REALIZATION의 SHARES_REAL 합계를 차감 (일부 매각분 반영)
- DIRECT_MARKUP_A에 최신 PRICE_PER_SHARE_MKA가 있으면 그 값으로 단가 갱신 (주수는 그대로 유지, MARKUP 시트의 SHARES_MKA는
  회사 전체 발행주식수이므로 우리 보유주수와 무관 - 사용하지 않음)

CALL_ID: 105600NEARDC1의 SECURITY_TYPE_CONV가 "RCPS"로 되어 있으나 NOTE_CONV("RCPS→CS 전환")와
뒤이은 105600NEARDC2(액면분할, CS 기준)의 정합성을 근거로 CS로 처리함 (원본 시트 오타로 추정, 사용자에게 확인 요청함).
"""

import gspread
from google.oauth2.service_account import Credentials

SHEET_ID = "1xYZqiMRclX6OkQn0V-RVwteC4D8OMUjDd22-3fb3hbg"
SERVICE_ACCOUNT_FILE = "secrets/service-account.json"

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# 계산된 최종 보유 현황 (CALL_ID 순서는 DIRECT_HOLDINGS 시트의 기존 행 순서와 동일)
HOLDINGS = [
    # CALL_ID,       SECURITY_TYPE_HELD, SHARES_HELD, CURRENCY_HELD, PRICE_PER_SHARE_HELD
    ("105600NAVIB", "RCPS", 595947, "KRW", 8390),
    ("105600ARADC", "CB", 1, "KRW", 2000000000),
    ("105600WAVIP", "CS", 180235, "KRW", 8550),
    ("105600AUTOC", "CS", 137034, "KRW", 21891.67),
    ("105600NEARD", "CS", 181550, "KRW", 22031.80),
    ("105600ICEYE", "CPS", 580703, "EUR", 74.98),
    ("105600ZENZA", "RCPS", 324470, "KRW", 18492),
    ("105600COCHB", "RCPS", 13713, "USD", 255),
    ("105600RLWRS", "SAFE", 1, "USD", 2000000),
    ("105600MADDA", "RCPS", 21579, "KRW", 231710),
]


def main():
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SHEET_ID)
    ws = sh.worksheet("DIRECT_HOLDINGS")

    existing = ws.get_all_values()
    header_row = existing[0]
    call_id_col = 0  # column A

    # CALL_ID -> 실제 행 번호 (1-indexed) 매핑
    row_map = {}
    for i, row in enumerate(existing[2:], start=3):  # row 1=header, row 2=한글설명
        if row and row[call_id_col]:
            row_map[row[call_id_col]] = i

    updates = []
    for call_id, sec_type, shares, currency, price in HOLDINGS:
        if call_id not in row_map:
            print(f"[경고] {call_id}가 DIRECT_HOLDINGS 시트에 없습니다. 건너뜁니다.")
            continue
        row_num = row_map[call_id]
        updates.append({
            "range": f"B{row_num}:E{row_num}",
            "values": [[sec_type, shares, currency, price]],
        })

    if updates:
        ws.batch_update(updates, value_input_option="USER_ENTERED")
        print(f"{len(updates)}개 행 업데이트 완료.")
    else:
        print("업데이트할 내용이 없습니다.")


if __name__ == "__main__":
    main()
