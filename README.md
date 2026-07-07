# VC투자팀 Track Record 웹사이트

펀드 수익률 분석 및 투자 데이터 관리 플랫폼

## 구조

- **Frontend**: `index.html` (정적 웹사이트, JavaScript UI)
- **Backend**: `app.py` (Python Flask 서버)
- **Data Source**: Google Sheets API

## 실행 방법

### 1. 필수 라이브러리 설치

```bash
pip3 install -r requirements.txt
```

### 2. 백엔드 서버 실행 (포트 5001)

```bash
python3 app.py
```

서버가 `http://localhost:5001`에서 실행됩니다.

### 3. 프론트엔드 서버 실행 (포트 8743)

다른 터미널에서:
```bash
python3 -m http.server 8743
```

### 4. 웹사이트 접속

브라우저에서 `http://localhost:8743` 접속

## 기능

### 탭 1: VC투자팀
- 기본 팀 정보 페이지

### 탭 2: 펀드별 수익률
4가지 입력 옵션으로 수익률 계산:

1. **DATE**: 연도/월 선택
2. **FUND**: 펀드 선택 (FUND_BASICS의 FUND_NICKNAME)
3. **FX**: 환율 적용 방식
   - **EXECUTED**: 집행 시점 환율
   - **BASE**: FUND_BASE_FX 시트의 환율
   - **SPOT**: 사용자 입력 환율 (선택한 펀드의 집행 통화만 표시)
4. **MARKUP**: 마크업 반영 방식
   - **ACTUAL**: DIRECT_MARKUP_A만 사용
   - **EXPECTED**: DIRECT_MARKUP_A + DIRECT_MARKUP_E 사용

## 데이터 구조

Google Sheets의 각 시트:
- **Row 1**: 데이터베이스 컬럼명
- **Row 2**: 한글 설명 (자동으로 스킵됨)
- **Row 3+**: 실제 데이터

사용되는 시트:
- `FUND_BASICS`: 펀드 기본 정보
- `DIRECT_INVESTMENT`: 직접투자 거래내역
- `FUND_BASE_FX`: 펀드별 기준 환율
- `DIRECT_MARKUP_A`: 실제 마크업
- `DIRECT_MARKUP_E`: 예상 마크업

## 변수 구조

계산을 위해 준비된 변수 (`fundReturnState`):

```javascript
{
  selectedDate: { year, month },           // 선택된 연도, 월
  selectedFund: string,                    // 펀드명
  selectedFundNickname: string,            // 펀드 닉네임
  fundCurrencies: string[],                // 해당 펀드의 집행 통화들
  fxOption: "EXECUTED"|"BASE"|"SPOT",     // 선택된 FX 옵션
  fxRates: {[currency]: number},           // { "USD": 1200, "EUR": 1300 }
  markupOption: "ACTUAL"|"EXPECTED"        // 선택된 마크업 옵션
}
```

## API 엔드포인트

### GET `/api/funds`
펀드 목록 반환
```json
[
  { "nickname": "군공", "row_index": 1 },
  { "nickname": "방수", "row_index": 2 }
]
```

### GET `/api/currencies?fund=<fund_name>`
펀드의 집행 통화 반환
```json
["EUR", "USD"]
```

### GET `/api/fund-base-fx`
기준 환율 데이터 반환

### GET `/api/direct-markup?type=ACTUAL|EXPECTED`
마크업 데이터 반환

### POST `/api/calculate`
수익률 계산 (준비 중)

## 비밀번호

기본 비밀번호: `Maxibest11^^`

`index.html`의 `PASSWORD` 변수에서 변경 가능

## 주의사항

- Google Sheet는 공개 설정되어야 함
- 로컬 개발 시 두 개의 서버가 동시에 실행되어야 함 (Frontend 8743, Backend 5001)
- Chrome/Safari 등 최신 브라우저 필요

## Vercel 배포

이 저장소는 `vercel.json`으로 Flask 백엔드(`app.py`)와 정적 프론트엔드(`index.html`)를
하나의 프로젝트로 함께 배포하도록 구성되어 있습니다.

1. [vercel.com](https://vercel.com)에서 New Project → 이 GitHub 저장소 Import
2. Environment Variables에 `GOOGLE_SERVICE_ACCOUNT_JSON` 추가
   - `secrets/service-account.json` 파일 내용 전체(JSON)를 값으로 붙여넣기
   - 이 변수가 없으면 DIRECT_HOLDINGS 갱신(쓰기) 기능만 실패하고, 나머지 조회 기능은 정상 동작
3. Deploy 클릭 → 완료 후 `https://<project>.vercel.app`으로 접속

배포 환경에서는 프론트엔드가 같은 도메인의 `/api`로 상대 경로 호출을 하므로
별도 CORS/URL 설정이 필요 없습니다.
