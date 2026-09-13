#!/usr/bin/env python3
"""
InterestPrint 제품/배송비 수집기
================================

InterestPrint(https://www.interestprint.com) 카테고리 페이지에서 제품 정보
(이름·판매가·정가·이미지·SKU)를 긁고, 각 제품의 국가별 배송비(배송사별 전체)를
내부 배송비 API로 조회해 하나의 JSON 으로 저장한다.

수집 방식(효율)
    - 카테고리당 요청 1번으로 제품 목록/가격/이미지 전부 확보 (제품 상세 페이지 불필요)
    - 제품당 배송비 요청 = 대상 국가 수 (기본 3: 미국/대만/일본)
    - 예의상 요청 사이 지연 + 재시도 + 로컬 캐시(json) 지원

사용법
    python scrape.py --category lunch_boxes-248
    python scrape.py --category lunch_boxes-248 --category tote_bags-210
    python scrape.py --categories-file categories.txt
    python scrape.py --all                 # /custom 인덱스의 전체 카테고리(242개, 매우 큼)
    python scrape.py --url https://www.interestprint.com/custom/1723-2256.html

결과: output/data.json  (이후 build_dashboard.py 로 엑셀/HTML 생성)
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import time
from dataclasses import dataclass, field, asdict
from pathlib import Path

import requests
from bs4 import BeautifulSoup

BASE = "https://www.interestprint.com"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/120.0 Safari/537.36")

# 조회할 국가: (표시명, short_code). 배송 API 의 country 파라미터는 short_code 를 사용.
DEFAULT_COUNTRIES = [
    ("미국 (US)", "US"),
    ("대만 (TW)", "TW"),
    ("일본 (JP)", "JP"),
]

HERE = Path(__file__).resolve().parent
OUTPUT_DIR = HERE / "output"


# --------------------------------------------------------------------------- #
# 데이터 모델
# --------------------------------------------------------------------------- #
@dataclass
class ShipOption:
    country: str          # 표시명 (예: "일본 (JP)")
    country_code: str     # short_code (예: "JP")
    carrier: str          # 배송사/서비스명 (예: "DHL Express")
    price: float | None   # 배송비 (USD). 파싱 실패 시 None
    price_text: str       # 원문 가격 문자열 (예: "$44.98")


@dataclass
class Product:
    pid: str
    model: str
    sku: str
    name: str
    sale_price: float | None
    sale_price_text: str
    retail_price: float | None
    retail_price_text: str
    image: str
    url: str
    category_slug: str
    category_name: str
    shipping: list[ShipOption] = field(default_factory=list)


# --------------------------------------------------------------------------- #
# HTTP
# --------------------------------------------------------------------------- #
def make_session() -> requests.Session:
    s = requests.Session()
    s.headers.update({
        "User-Agent": UA,
        "Accept-Language": "en-US,en;q=0.9",
    })
    return s


def get(session: requests.Session, url: str, *, retries: int = 4,
        backoff: float = 2.0, **kwargs) -> requests.Response | None:
    """GET/POST 공통 재시도 래퍼. method 는 kwargs 로 넘어온 data 유무로 결정."""
    method = "POST" if "data" in kwargs else "GET"
    last_exc = None
    for attempt in range(retries):
        try:
            resp = session.request(method, url, timeout=30, **kwargs)
            if resp.status_code == 200:
                return resp
            last_exc = f"HTTP {resp.status_code}"
        except requests.RequestException as exc:  # noqa: BLE001
            last_exc = exc
        wait = backoff * (2 ** attempt)
        print(f"    ! {method} {url} 실패({last_exc}), {wait:.0f}s 후 재시도 "
              f"({attempt + 1}/{retries})", file=sys.stderr)
        time.sleep(wait)
    print(f"    ✗ 최종 실패: {method} {url} ({last_exc})", file=sys.stderr)
    return None


# --------------------------------------------------------------------------- #
# 파싱
# --------------------------------------------------------------------------- #
def _price_to_float(text: str) -> float | None:
    m = re.search(r"[\d,]+\.\d{2}|\d+", text or "")
    if not m:
        return None
    try:
        return float(m.group(0).replace(",", ""))
    except ValueError:
        return None


def parse_category_products(html: str, slug: str, cat_name: str) -> list[Product]:
    """카테고리 목록 페이지에서 제품 카드를 파싱한다."""
    soup = BeautifulSoup(html, "lxml")
    products: dict[str, Product] = {}
    for card in soup.select(".ca-product"):
        a = card.select_one('a[href*="/custom/"]')
        if not a:
            continue
        # 제품 URL: /custom/<slug 또는 모델번호>-<pid>.html
        # 모델 부분이 숫자(예: 1723)일 수도, 텍스트 슬러그일 수도 있음
        m = re.search(r"/custom/([A-Za-z0-9_]+)-(\d+)\.html", a.get("href", ""))
        if not m:
            continue
        model, pid = m.group(1), m.group(2)
        if pid in products:
            continue
        name_el = card.select_one(".products-list-text")
        name = name_el.get_text(strip=True) if name_el else ""
        price_el = card.select_one(".products-list-price")
        sale_text = price_el.get_text(strip=True) if price_el else ""
        retail_text = ""
        for p in card.select("p"):
            if "Retail Price" in p.get_text():
                rm = re.search(r"\$[\d.,]+", p.get_text())
                retail_text = rm.group(0) if rm else ""
                break
        img_el = card.select_one("img")
        img = ""
        if img_el:
            img = img_el.get("src") or img_el.get("data-original") or ""
        code_el = card.select_one("[data-product-code]")
        sku = code_el.get("data-product-code") if code_el else ""
        products[pid] = Product(
            pid=pid, model=model, sku=sku, name=name,
            sale_price=_price_to_float(sale_text), sale_price_text=sale_text,
            retail_price=_price_to_float(retail_text), retail_price_text=retail_text,
            image=img,
            url=f"{BASE}/custom/{model}-{pid}.html",
            category_slug=slug, category_name=cat_name,
        )
    return list(products.values())


def parse_shipping(html: str, country_name: str, country_code: str) -> list[ShipOption]:
    """배송비 API(HTML) 응답에서 배송사/가격을 파싱한다."""
    soup = BeautifulSoup(html, "lxml")
    options: list[ShipOption] = []
    for row in soup.select(".p-select1"):
        carrier_el = row.select_one(".col-xs-9 p.g-fs16") or row.select_one("p.g-fs16")
        carrier = carrier_el.get_text(strip=True) if carrier_el else ""
        price_el = row.select_one(".col-xs-3")
        price_text = price_el.get_text(strip=True) if price_el else ""
        if not price_text:
            # fallback: 행 안에서 $금액 검색
            pm = re.search(r"\$[\d.,]+", row.get_text())
            price_text = pm.group(0) if pm else ""
        if not carrier and not price_text:
            continue
        options.append(ShipOption(
            country=country_name, country_code=country_code,
            carrier=carrier, price=_price_to_float(price_text),
            price_text=price_text,
        ))
    return options


def parse_single_product(html: str, url: str) -> Product | None:
    """제품 상세 페이지(--url 사용 시)에서 최소 정보를 파싱한다."""
    soup = BeautifulSoup(html, "lxml")
    m = re.search(r"/custom/([A-Za-z0-9_]+)-(\d+)\.html", url)
    if not m:
        pm = re.search(r"pid'\s*:\s*'(\d+)'", html)
        pid = pm.group(1) if pm else ""
        model = ""
    else:
        model, pid = m.group(1), m.group(2)
    title = soup.select_one("title")
    name = ""
    if title:
        name = re.sub(r"[-|].*$", "", title.get_text()).replace("Custom", "").strip()
    og = soup.select_one('meta[property="og:image"]')
    img = og.get("content") if og else ""
    # 가격: 상세 페이지의 판매가 텍스트 추출 시도
    sale_text = ""
    price_el = soup.find(string=re.compile(r"\$\d"))
    if price_el:
        pm = re.search(r"\$[\d.,]+", price_el)
        sale_text = pm.group(0) if pm else ""
    return Product(
        pid=pid, model=model, sku="", name=name,
        sale_price=_price_to_float(sale_text), sale_price_text=sale_text,
        retail_price=None, retail_price_text="",
        image=img, url=url, category_slug="(single)", category_name="(single)",
    )


# --------------------------------------------------------------------------- #
# 수집 로직
# --------------------------------------------------------------------------- #
def fetch_shipping(session: requests.Session, product: Product,
                   countries, delay: float) -> None:
    for cname, ccode in countries:
        resp = get(
            session, f"{BASE}/shop/shipping_pirce_estimate?shipping_estimate",
            data={
                "country": ccode,
                "address_state": "",
                "product_id": product.pid,
                "container": "shipping_estimate",
                "qty": "1",
                "cp": product.url,
            },
            headers={"X-Requested-With": "XMLHttpRequest",
                     "Referer": f"{BASE}/shop/shippingprice_estimation?pid={product.pid}"},
        )
        if resp is None:
            continue
        opts = parse_shipping(resp.text, cname, ccode)
        product.shipping.extend(opts)
        prices = ", ".join(f"{o.carrier} {o.price_text}" for o in opts) or "(없음)"
        print(f"      · {cname}: {prices}")
        time.sleep(delay)


def enumerate_all_categories(session: requests.Session) -> list[tuple[str, str]]:
    resp = get(session, f"{BASE}/custom")
    if resp is None:
        return []
    cats = re.findall(r'/custom/([a-z0-9_]+)-(\d+)"[^>]*>\s*([^<]+)', resp.text)
    seen: dict[str, tuple[str, str]] = {}
    for slug, cid, name in cats:
        key = f"{slug}-{cid}"
        name = BeautifulSoup(name, "lxml").get_text(strip=True)
        seen.setdefault(key, (key, name))
    return list(seen.values())


def scrape(categories, urls, countries, delay: float) -> list[Product]:
    session = make_session()
    all_products: list[Product] = []

    for slug in categories:
        cat_url = f"{BASE}/custom/{slug}"
        print(f"\n[카테고리] {slug}  ({cat_url})")
        resp = get(session, cat_url)
        if resp is None:
            continue
        # 카테고리 표시명 추출
        soup = BeautifulSoup(resp.text, "lxml")
        h1 = soup.select_one("h1")
        cat_name = h1.get_text(strip=True) if h1 else slug
        products = parse_category_products(resp.text, slug, cat_name)
        print(f"  제품 {len(products)}개 발견")
        for p in products:
            print(f"  - [{p.pid}] {p.name}  판매가 {p.sale_price_text or 'N/A'}")
            fetch_shipping(session, p, countries, delay)
            all_products.append(p)
        time.sleep(delay)

    for url in urls:
        print(f"\n[단일 제품] {url}")
        resp = get(session, url)
        if resp is None:
            continue
        p = parse_single_product(resp.text, url)
        if p is None:
            continue
        print(f"  [{p.pid}] {p.name}")
        fetch_shipping(session, p, countries, delay)
        all_products.append(p)

    return all_products


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def main() -> None:
    ap = argparse.ArgumentParser(description="InterestPrint 제품/배송비 수집기")
    ap.add_argument("--category", action="append", default=[],
                    help="카테고리 슬러그 (예: lunch_boxes-248). 여러 번 지정 가능")
    ap.add_argument("--categories-file",
                    help="카테고리 슬러그를 줄바꿈으로 나열한 파일")
    ap.add_argument("--url", action="append", default=[],
                    help="단일 제품 상세 URL. 여러 번 지정 가능")
    ap.add_argument("--all", action="store_true",
                    help="/custom 인덱스의 전체 카테고리 수집 (242개, 매우 큼/느림)")
    ap.add_argument("--delay", type=float, default=0.8,
                    help="요청 사이 지연(초). 기본 0.8")
    ap.add_argument("--out", default=str(OUTPUT_DIR / "data.json"),
                    help="결과 JSON 경로")
    args = ap.parse_args()

    session_for_all = make_session()
    categories: list[str] = list(args.category)
    if args.categories_file:
        for line in Path(args.categories_file).read_text(encoding="utf-8").splitlines():
            line = line.split("#", 1)[0].strip()  # 인라인 주석 제거
            if line:
                categories.append(line)
    if args.all:
        print("전체 카테고리 열거 중...")
        categories = [slug for slug, _ in enumerate_all_categories(session_for_all)]
        print(f"  {len(categories)}개 카테고리")

    if not categories and not args.url:
        ap.error("최소한 --category, --categories-file, --all, --url 중 하나가 필요합니다.")

    products = scrape(categories, args.url, DEFAULT_COUNTRIES, args.delay)

    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "source": "interestprint.com",
        "countries": [{"name": n, "code": c} for n, c in DEFAULT_COUNTRIES],
        "scraped_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "product_count": len(products),
        "products": [asdict(p) for p in products],
    }
    out_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2),
                        encoding="utf-8")
    print(f"\n✓ 제품 {len(products)}개 저장 → {out_path}")
    print(f"  다음: python build_dashboard.py --data {out_path}")


if __name__ == "__main__":
    main()
