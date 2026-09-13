# InterestPrint 제품 · 배송비 대시보드

[InterestPrint](https://www.interestprint.com) 카테고리 페이지에서 **제품 정보(이름·판매가·정가·이미지·SKU)** 와
**국가별 배송비(배송사별 전체)** 를 한 번에 긁어 **엑셀 장표**와 **HTML 대시보드**로 정리하는 도구입니다.

기본 조회 국가: **미국(US) · 대만(TW) · 일본(JP)** — 사이트의 "Shipping Calculate"를 일일이 클릭할 필요 없이,
내부 배송비 계산 API(`shipping_pirce_estimate`)를 직접 호출해 수량 1개 기준 실시간 값을 가져옵니다.

## 왜 효율적인가

| 항목 | 방식 | 요청 수 |
|---|---|---|
| 제품 목록·가격·이미지·SKU | 카테고리 목록 페이지 1번 파싱 (제품 상세 페이지 불필요) | 카테고리당 1 |
| 배송비 (배송사 전체) | 배송비 API 직접 호출 | 제품당 = 국가 수(기본 3) |

로그인·브라우저 자동화 없이 순수 HTTP 요청으로 동작합니다.

## 설치

```bash
pip install requests beautifulsoup4 lxml openpyxl pillow
```

## 사용법

### 1) 수집 (`scrape.py`) → `output/data.json`

```bash
# 카테고리 목록 파일로 (권장)
python scrape.py --categories-file categories.txt

# 카테고리 직접 지정 (여러 번 가능)
python scrape.py --category lunch_boxes-248 --category tote_bags-210

# 단일 제품 URL
python scrape.py --url https://www.interestprint.com/custom/1723-2256.html

# 전체 카테고리(242개) — 매우 크고 느림, 차단 위험. --delay 를 넉넉히.
python scrape.py --all --delay 1.0
```

옵션
- `--delay`  요청 사이 지연(초, 기본 0.8). 대량 수집 시 1.0 이상 권장.
- `--out`    결과 JSON 경로 (기본 `output/data.json`).

### 2) 대시보드 생성 (`build_dashboard.py`)

```bash
python build_dashboard.py
# → output/interestprint_dashboard.xlsx  (제품 이미지 임베드)
# → output/dashboard.html                (브라우저용 카드형 대시보드)
```

## 카테고리 슬러그 찾는 법

`https://www.interestprint.com/custom` 페이지의 카테고리 링크가 `/custom/<슬러그>-<ID>` 형태입니다.
예: `Lunch Bags → lunch_boxes-248`, `Tote Bags → tote_bags-210`.
`categories.txt` 에 한 줄에 하나씩 (뒤에 `#` 주석 가능) 넣으면 됩니다.

## 산출물 구성

**엑셀 (`interestprint_dashboard.xlsx`)**
- `제품 대시보드` 시트 — 제품 1행. 이미지 · 판매가/정가 · 국가별 **최저 배송비 / 배송사 수 / 배송사별 상세**
- `배송비 상세` 시트 — (제품 × 국가 × 배송사) 롱포맷. 자동필터/피벗테이블로 분석
- `README` 시트 — 수집 조건 안내

**HTML (`dashboard.html`)** — 제품 카드 그리드, 국가별 배송사·가격 표. 브라우저로 바로 열람.

## 조회 국가 바꾸기

`scrape.py` 상단 `DEFAULT_COUNTRIES` 에서 `(표시명, short_code)` 를 추가/수정하세요.
`short_code` 는 배송비 계산 팝업의 국가 `<option short_code="...">` 값과 동일합니다 (예: 한국 `KR`).

## 참고 / 주의

- 가격은 USD, 배송비는 **수량 1개** 기준입니다. 수량별 배송비가 필요하면 `scrape.py` 의 `qty` 를 조정하세요.
- 판매가/정가/배송비는 **수집 시점** 값입니다. 최신화하려면 다시 실행하세요.
- 대량 수집 시 사이트에 부담을 주지 않도록 `--delay` 를 충분히 두세요.
