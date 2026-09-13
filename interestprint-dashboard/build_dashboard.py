#!/usr/bin/env python3
"""
대시보드 빌더
=============

scrape.py 가 만든 output/data.json 을 읽어 두 가지 산출물을 생성한다.

    1) 엑셀 (.xlsx)  — 제품 이미지 임베드 + 국가별 배송사별 배송비
        · "제품 대시보드" 시트: 제품 1행, 국가별 최저가/배송사수 요약
        · "배송비 상세" 시트: (제품 × 국가 × 배송사) 롱포맷 (필터/피벗용)
    2) HTML 대시보드 — 브라우저로 바로 보는 카드형 요약

사용법
    python build_dashboard.py                       # output/data.json 사용
    python build_dashboard.py --data output/data.json
"""
from __future__ import annotations

import argparse
import html as html_lib
import json
from collections import defaultdict
from io import BytesIO
from pathlib import Path

import requests
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

HERE = Path(__file__).resolve().parent
OUTPUT_DIR = HERE / "output"
IMG_CACHE = OUTPUT_DIR / "images"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/120.0 Safari/537.36")

THUMB = 90  # 엑셀 셀 안 썸네일 픽셀

HEADER_FILL = PatternFill("solid", fgColor="1F2937")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
SUBFILL = PatternFill("solid", fgColor="EEF2FF")
THIN = Side(style="thin", color="D1D5DB")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


def load(data_path: Path) -> dict:
    return json.loads(data_path.read_text(encoding="utf-8"))


def download_image(url: str, pid: str) -> Path | None:
    if not url:
        return None
    if url.startswith("//"):
        url = "https:" + url
    IMG_CACHE.mkdir(parents=True, exist_ok=True)
    ext = ".jpg"
    dest = IMG_CACHE / f"{pid}{ext}"
    if dest.exists() and dest.stat().st_size > 0:
        return dest
    try:
        r = requests.get(url, headers={"User-Agent": UA}, timeout=30)
        if r.status_code == 200 and r.content:
            dest.write_bytes(r.content)
            return dest
    except requests.RequestException:
        pass
    return None


def shipping_by_country(product: dict) -> dict[str, list[dict]]:
    grouped: dict[str, list[dict]] = defaultdict(list)
    for opt in product.get("shipping", []):
        grouped[opt["country"]].append(opt)
    return grouped


def min_price(opts: list[dict]) -> float | None:
    vals = [o["price"] for o in opts if o.get("price") is not None]
    return min(vals) if vals else None


# --------------------------------------------------------------------------- #
# 엑셀
# --------------------------------------------------------------------------- #
def build_excel(data: dict, xlsx_path: Path) -> None:
    products = data["products"]
    countries = [c["name"] for c in data["countries"]]

    wb = Workbook()
    ws = wb.active
    ws.title = "제품 대시보드"

    # 헤더
    base_cols = ["이미지", "제품ID", "SKU", "제품명", "판매가(USD)", "정가(USD)", "카테고리"]
    ship_cols: list[str] = []
    for c in countries:
        ship_cols += [f"{c} 최저 배송비", f"{c} 배송사 수", f"{c} 배송사별"]
    headers = base_cols + ship_cols + ["제품 링크"]

    for col, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER
        cell.border = BORDER
    ws.freeze_panes = "D2"

    # 열 너비
    widths = [14, 8, 10, 40, 12, 12, 18]
    for c in countries:
        widths += [14, 12, 34]
    widths += [24]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    row = 2
    for p in products:
        ws.row_dimensions[row].height = THUMB * 0.78
        # 이미지
        img_path = download_image(p.get("image", ""), p["pid"])
        if img_path:
            try:
                xi = XLImage(str(img_path))
                xi.width = xi.height = THUMB
                ws.add_image(xi, f"A{row}")
            except Exception:  # noqa: BLE001
                ws.cell(row=row, column=1, value="(이미지)")
        ws.cell(row=row, column=2, value=p.get("pid", ""))
        ws.cell(row=row, column=3, value=p.get("sku", ""))
        ws.cell(row=row, column=4, value=p.get("name", ""))
        ws.cell(row=row, column=5, value=p.get("sale_price"))
        ws.cell(row=row, column=6, value=p.get("retail_price"))
        ws.cell(row=row, column=7, value=p.get("category_name", ""))

        grouped = shipping_by_country(p)
        col = 8
        for cname in countries:
            opts = grouped.get(cname, [])
            mp = min_price(opts)
            ws.cell(row=row, column=col, value=mp)
            ws.cell(row=row, column=col + 1, value=len(opts))
            detail = "\n".join(
                f"{o['carrier']}: {o['price_text']}" for o in opts) or "-"
            ws.cell(row=row, column=col + 2, value=detail)
            col += 3

        link_cell = ws.cell(row=row, column=col, value=p.get("url", ""))
        link_cell.hyperlink = p.get("url", "")
        link_cell.font = Font(color="2563EB", underline="single")

        # 스타일
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=row, column=c)
            cell.border = BORDER
            cell.alignment = LEFT if c in (4,) or c >= 8 else CENTER
        row += 1

    # 통화 서식
    for r in range(2, row):
        for c in [5, 6] + [8 + 3 * i for i in range(len(countries))]:
            ws.cell(row=r, column=c).number_format = '"$"#,##0.00'

    # -------- 시트 2: 배송비 상세 (롱포맷) --------
    ws2 = wb.create_sheet("배송비 상세")
    det_headers = ["제품ID", "제품명", "카테고리", "국가", "국가코드",
                   "배송사/서비스", "배송비(USD)", "제품 링크"]
    for col, h in enumerate(det_headers, start=1):
        cell = ws2.cell(row=1, column=col, value=h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER
        cell.border = BORDER
    ws2.freeze_panes = "A2"
    for i, w in enumerate([8, 38, 18, 12, 10, 28, 14, 24], start=1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    r = 2
    for p in products:
        for o in p.get("shipping", []):
            ws2.cell(row=r, column=1, value=p.get("pid", ""))
            ws2.cell(row=r, column=2, value=p.get("name", ""))
            ws2.cell(row=r, column=3, value=p.get("category_name", ""))
            ws2.cell(row=r, column=4, value=o.get("country", ""))
            ws2.cell(row=r, column=5, value=o.get("country_code", ""))
            ws2.cell(row=r, column=6, value=o.get("carrier", ""))
            pc = ws2.cell(row=r, column=7, value=o.get("price"))
            pc.number_format = '"$"#,##0.00'
            ws2.cell(row=r, column=8, value=p.get("url", ""))
            for c in range(1, 9):
                ws2.cell(row=r, column=c).border = BORDER
                ws2.cell(row=r, column=c).alignment = LEFT if c in (2, 3, 6) else CENTER
            r += 1
    ws2.auto_filter.ref = f"A1:H{max(1, r - 1)}"

    # -------- 시트 3: 안내 --------
    ws3 = wb.create_sheet("README")
    notes = [
        ["InterestPrint 제품 · 배송비 대시보드"],
        [f"수집 시각: {data.get('scraped_at', '')}"],
        [f"제품 수: {data.get('product_count', 0)}"],
        [f"조회 국가: {', '.join(countries)}  (수량 1개 기준)"],
        [""],
        ["· '제품 대시보드' 시트: 제품별 판매가/정가 + 국가별 최저 배송비, 배송사별 상세"],
        ["· '배송비 상세' 시트: (제품×국가×배송사) 롱포맷 — 자동필터/피벗테이블로 분석"],
        ["· 배송비는 InterestPrint 배송비 계산 API(shipping_pirce_estimate) 실시간 조회값"],
        ["· 가격은 USD, 배송비는 수량 1개 기준. 정가/판매가는 수집 시점 값"],
    ]
    for i, line in enumerate(notes, start=1):
        cell = ws3.cell(row=i, column=1, value=line[0])
        if i == 1:
            cell.font = Font(bold=True, size=14)
    ws3.column_dimensions["A"].width = 90

    xlsx_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(xlsx_path)
    print(f"✓ 엑셀 저장 → {xlsx_path}")


# --------------------------------------------------------------------------- #
# HTML
# --------------------------------------------------------------------------- #
def build_html(data: dict, html_path: Path) -> None:
    products = data["products"]
    countries = [c["name"] for c in data["countries"]]
    esc = html_lib.escape

    cards = []
    for p in products:
        grouped = shipping_by_country(p)
        ship_html = []
        for cname in countries:
            opts = grouped.get(cname, [])
            mp = min_price(opts)
            rows = "".join(
                f'<tr><td>{esc(o["carrier"])}</td>'
                f'<td class="pr">{esc(o["price_text"] or "-")}</td></tr>'
                for o in opts) or '<tr><td colspan="2">배송 옵션 없음</td></tr>'
            mp_txt = f"${mp:,.2f}" if mp is not None else "-"
            ship_html.append(f"""
            <div class="country">
              <div class="chead"><span>{esc(cname)}</span>
                <span class="min">최저 {mp_txt}</span></div>
              <table class="stbl"><tbody>{rows}</tbody></table>
            </div>""")
        sale = f'${p["sale_price"]:,.2f}' if p.get("sale_price") is not None else esc(p.get("sale_price_text", "-"))
        retail = f'${p["retail_price"]:,.2f}' if p.get("retail_price") is not None else ""
        img = p.get("image", "")
        if img.startswith("//"):
            img = "https:" + img
        cards.append(f"""
        <div class="card">
          <a href="{esc(p.get('url',''))}" target="_blank" rel="noopener">
            <img loading="lazy" src="{esc(img)}" alt="{esc(p.get('name',''))}"></a>
          <div class="body">
            <div class="cat">{esc(p.get('category_name',''))}</div>
            <a class="name" href="{esc(p.get('url',''))}" target="_blank" rel="noopener">{esc(p.get('name',''))}</a>
            <div class="price"><span class="sale">{sale}</span>
              {'<span class="retail">'+retail+'</span>' if retail else ''}</div>
            <div class="ship">{''.join(ship_html)}</div>
          </div>
        </div>""")

    doc = f"""<!doctype html>
<html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>InterestPrint 제품 · 배송비 대시보드</title>
<style>
  :root {{ color-scheme: light dark; }}
  * {{ box-sizing: border-box; }}
  body {{ margin:0; font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"Malgun Gothic",sans-serif;
         background:#f6f7f9; color:#1f2937; }}
  header {{ background:#111827; color:#fff; padding:20px 24px; }}
  header h1 {{ margin:0 0 4px; font-size:20px; }}
  header p {{ margin:0; color:#9ca3af; font-size:13px; }}
  .grid {{ display:grid; grid-template-columns:repeat(auto-fill,minmax(300px,1fr));
           gap:16px; padding:20px; max-width:1400px; margin:0 auto; }}
  .card {{ background:#fff; border:1px solid #e5e7eb; border-radius:12px; overflow:hidden;
           display:flex; flex-direction:column; }}
  .card > a {{ display:block; background:#fff; text-align:center; }}
  .card img {{ width:100%; max-width:220px; height:auto; margin:12px auto 0; }}
  .body {{ padding:12px 14px 16px; display:flex; flex-direction:column; gap:8px; }}
  .cat {{ font-size:11px; color:#6b7280; text-transform:uppercase; letter-spacing:.03em; }}
  .name {{ font-weight:600; font-size:14px; line-height:1.35; color:#111827; text-decoration:none; }}
  .name:hover {{ color:#2563eb; }}
  .price {{ display:flex; align-items:baseline; gap:8px; }}
  .sale {{ font-size:18px; font-weight:700; color:#dc2626; }}
  .retail {{ font-size:13px; color:#9ca3af; text-decoration:line-through; }}
  .ship {{ display:flex; flex-direction:column; gap:8px; margin-top:4px; }}
  .country {{ border:1px solid #eef0f3; border-radius:8px; overflow:hidden; }}
  .chead {{ display:flex; justify-content:space-between; align-items:center;
            background:#eef2ff; padding:6px 10px; font-size:12px; font-weight:600; }}
  .chead .min {{ color:#4338ca; }}
  .stbl {{ width:100%; border-collapse:collapse; font-size:12px; }}
  .stbl td {{ padding:4px 10px; border-top:1px solid #f1f2f4; }}
  .stbl td.pr {{ text-align:right; font-variant-numeric:tabular-nums; white-space:nowrap; }}
  @media (prefers-color-scheme: dark) {{
    body {{ background:#0b0f17; color:#e5e7eb; }}
    .card {{ background:#111827; border-color:#1f2937; }}
    .card > a {{ background:#fff; }}
    .name {{ color:#f3f4f6; }}
    .country {{ border-color:#1f2937; }}
    .chead {{ background:#1e293b; }}
    .chead .min {{ color:#a5b4fc; }}
    .stbl td {{ border-color:#1f2937; }}
  }}
</style></head><body>
<header>
  <h1>InterestPrint 제품 · 배송비 대시보드</h1>
  <p>제품 {data.get('product_count',0)}개 · 조회 국가: {esc(', '.join(countries))} (수량 1개 기준) · 수집 {esc(data.get('scraped_at',''))}</p>
</header>
<div class="grid">{''.join(cards)}</div>
</body></html>"""
    html_path.parent.mkdir(parents=True, exist_ok=True)
    html_path.write_text(doc, encoding="utf-8")
    print(f"✓ HTML 저장 → {html_path}")


def main() -> None:
    ap = argparse.ArgumentParser(description="InterestPrint 대시보드 빌더")
    ap.add_argument("--data", default=str(OUTPUT_DIR / "data.json"))
    ap.add_argument("--xlsx", default=str(OUTPUT_DIR / "interestprint_dashboard.xlsx"))
    ap.add_argument("--html", default=str(OUTPUT_DIR / "dashboard.html"))
    args = ap.parse_args()

    data = load(Path(args.data))
    build_excel(data, Path(args.xlsx))
    build_html(data, Path(args.html))


if __name__ == "__main__":
    main()
