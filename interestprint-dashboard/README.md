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
- `--no-shipping`  배송비 없이 **카탈로그(가격 포함)만** 빠르게 수집.
- `--out`    결과 JSON 경로 (기본 `output/data.json`).

### 1-대) 대규모 권장 워크플로 — 전체 카탈로그 → 저단가 후보만 배송비

전체 사이트(242 카테고리, 약 3,300개 제품)에 배송비를 모두 붙이면 1만 요청에 육박해
느리고 차단 위험이 큽니다. **직접 팔 저단가 제품 발굴**이 목적이면 2단계로:

```bash
# ① 전체 카탈로그 + 단가만 빠르게 (배송비 X) — 약 242 요청
python scrape.py --all --no-shipping --delay 0.3 --out output/catalog.json

# ② 저단가 후보에만 배송비 조회 (원가 $5 이하) — 카탈로그 재사용, 재수집 없음
python scrape.py --from-catalog output/catalog.json --max-price 5 --out output/data.json
#   또는 상위 N개만:  --top-cheapest 300
```

`--from-catalog` 옵션
- `--max-price X`     원가(판매가) X 이하 제품만 배송비 조회
- `--top-cheapest N`  원가 낮은 순 상위 N개만 배송비 조회
- 결과 JSON 에는 **전체 제품이 유지**되며(전체 단가 랭킹용), 선택된 후보만 배송비가 채워집니다.

### 2) 대시보드 · 판매후보 분석 생성 (`build_dashboard.py`)

```bash
python build_dashboard.py
# → output/interestprint_dashboard.xlsx  (분석 엑셀, 후보 이미지 임베드)
# → output/catalog_by_price.csv          (전체 제품 원가순 CSV)
# → output/dashboard.html                (저단가 후보 카드형 대시보드)
```

## 카테고리 슬러그 찾는 법

`https://www.interestprint.com/custom` 페이지의 카테고리 링크가 `/custom/<슬러그>-<ID>` 형태입니다.
예: `Lunch Bags → lunch_boxes-248`, `Tote Bags → tote_bags-210`.
`categories.txt` 에 한 줄에 하나씩 (뒤에 `#` 주석 가능) 넣으면 됩니다.

## 산출물 구성

**엑셀 (`interestprint_dashboard.xlsx`)**
- `저단가 판매후보` 시트 — 배송비까지 조회된 후보를 **원가 낮은 순**으로. 이미지 +
  원가/권장정가/**마진/마진율** + 국가별 최저배송 + **예상착지원가**(원가+US최저배송).
  마진율 60% 이상은 초록 강조.
- `전체 카탈로그(단가순)` 시트 — 수집된 **모든 제품**의 원가/마진 마스터 목록(자동필터).
- `배송비 상세` 시트 — (후보 × 국가 × 배송사) 롱포맷. 피벗/필터 분석용.
- `README` 시트 — 수집 조건 안내.

**CSV (`catalog_by_price.csv`)** — 전체 제품 원가순. 구글시트/엑셀에서 자유 필터·정렬.

**HTML (`dashboard.html`)** — 저단가 후보 카드 그리드 + 국가별 배송사·가격. 브라우저로 열람.

### 용어 (판매후보 관점)
- **원가(base)** = 사이트 판매가(POD 공급가) ≒ 내가 물건에 지불하는 값.
- **권장정가(retail)** = 사이트 Retail Price(제안 소비자가).
- **마진** = 권장정가 − 원가, **마진율** = 마진 / 권장정가.
- 원가가 낮을수록 내 디자인 마진을 붙일 여유가 큼 → 판매 후보로 매력적.
  (배송비까지 감안하려면 `예상착지원가` 참고.)

> 참고: `output/` 의 `interestprint_dashboard.xlsx`·`dashboard.html` 은 용량이 커서 저장소에
> 커밋하지 않습니다(`.gitignore`). `output/data.json` 으로 `build_dashboard.py` 를 다시 돌리면
> 언제든 재생성됩니다. 원가순 CSV(`catalog_by_price.csv`)와 원본 데이터(json)는 커밋됩니다.

## 조회 국가 바꾸기

`scrape.py` 상단 `DEFAULT_COUNTRIES` 에서 `(표시명, short_code)` 를 추가/수정하세요.
`short_code` 는 배송비 계산 팝업의 국가 `<option short_code="...">` 값과 동일합니다 (예: 한국 `KR`).

## 참고 / 주의

- 가격은 USD, 배송비는 **수량 1개** 기준입니다. 수량별 배송비가 필요하면 `scrape.py` 의 `qty` 를 조정하세요.
- 판매가/정가/배송비는 **수집 시점** 값입니다. 최신화하려면 다시 실행하세요.
- 대량 수집 시 사이트에 부담을 주지 않도록 `--delay` 를 충분히 두세요.
