#!/usr/bin/env python3
"""
인터랙티브 HTML 대시보드 빌더
=============================

data.json(전체 카탈로그 + 일부 배송비)을 받아, 브라우저에서 **카테고리·가격·정렬·검색**으로
자유롭게 훑어볼 수 있는 단일 HTML 파일을 만든다. 모든 제품 데이터를 인라인 JSON 으로
포함하고, 필터링은 순수 자바스크립트로 클라이언트에서 처리한다. (외부 CDN 이미지는
브라우저에서 지연 로드 — 인터넷 연결 필요.)
"""
from __future__ import annotations

import json
from collections import defaultdict
from pathlib import Path


def _min(opts):
    vals = [o["price"] for o in opts if o.get("price") is not None]
    return round(min(vals), 2) if vals else None


def _records(data: dict):
    country_codes = [c["code"] for c in data["countries"]]
    recs = []
    for p in data["products"]:
        base = p.get("sale_price")
        retail = p.get("retail_price")
        mpct = None
        if base is not None and retail:
            mpct = round((retail - base) / retail * 100, 1)
        by_code = defaultdict(list)
        for o in p.get("shipping", []):
            by_code[o["country_code"]].append(o)
        ship = {}
        for code in country_codes:
            opts = by_code.get(code)
            if opts:
                ship[code] = {
                    "min": _min(opts),
                    "opts": [[o.get("carrier", ""), o.get("price_text", "")] for o in opts],
                }
        img = p.get("image", "")
        if img.startswith("//"):
            img = "https:" + img
        recs.append({
            "id": p.get("pid", ""),
            "n": p.get("name", ""),
            "c": p.get("category_name", ""),
            "s": base,
            "r": retail,
            "m": mpct,
            "img": img,
            "u": p.get("url", ""),
            "sh": ship,
        })
    return recs, country_codes


def build_html(data: dict, html_path: Path) -> None:
    recs, country_codes = _records(data)
    countries = [(c["code"], c["name"]) for c in data["countries"]]
    prices = [r["s"] for r in recs if r["s"] is not None]
    pmin = min(prices) if prices else 0
    pmax = max(prices) if prices else 0
    n_ship = sum(1 for r in recs if r["sh"])

    payload = json.dumps(recs, ensure_ascii=False, separators=(",", ":"))
    payload = payload.replace("</", "<\\/")  # </script> 방어
    countries_js = json.dumps(countries, ensure_ascii=False)

    doc = """<!doctype html>
<html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>InterestPrint 제품 탐색기</title>
<style>
  :root{ color-scheme:light dark; --bg:#f6f7f9; --fg:#1f2937; --card:#fff; --line:#e5e7eb;
         --muted:#6b7280; --accent:#4338ca; --chip:#eef2ff; --sale:#dc2626; }
  @media (prefers-color-scheme:dark){ :root{ --bg:#0b0f17; --fg:#e5e7eb; --card:#111827;
         --line:#1f2937; --muted:#9ca3af; --accent:#a5b4fc; --chip:#1e293b; --sale:#f87171; } }
  *{box-sizing:border-box}
  body{margin:0;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"Malgun Gothic",sans-serif;
       background:var(--bg);color:var(--fg)}
  header{position:sticky;top:0;z-index:10;background:var(--card);border-bottom:1px solid var(--line);
         padding:12px 18px;box-shadow:0 1px 6px rgba(0,0,0,.04)}
  h1{margin:0 0 8px;font-size:17px}
  .sub{color:var(--muted);font-size:12px;margin-bottom:10px}
  .controls{display:flex;flex-wrap:wrap;gap:10px;align-items:flex-end}
  .fld{display:flex;flex-direction:column;gap:3px}
  .fld label{font-size:11px;color:var(--muted);font-weight:600}
  input,select{font:inherit;padding:6px 8px;border:1px solid var(--line);border-radius:8px;
       background:var(--bg);color:var(--fg)}
  input[type=number]{width:92px}
  #q{width:200px}
  .btn{padding:6px 12px;border:1px solid var(--line);border-radius:8px;background:var(--chip);
       color:var(--fg);cursor:pointer;font-size:13px}
  .count{margin-left:auto;font-size:13px;color:var(--muted);align-self:center}
  .count b{color:var(--fg)}
  .grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(240px,1fr));gap:14px;
        padding:16px;max-width:1500px;margin:0 auto}
  .sec-title{grid-column:1/-1;font-size:15px;font-weight:700;margin:8px 2px 0;
        border-bottom:2px solid var(--line);padding-bottom:6px}
  .card{background:var(--card);border:1px solid var(--line);border-radius:12px;overflow:hidden;
        display:flex;flex-direction:column}
  .card>a{display:block;background:#fff;text-align:center;min-height:150px}
  .card img{width:100%;max-width:180px;height:auto;margin:10px auto 0}
  .body{padding:10px 12px 14px;display:flex;flex-direction:column;gap:6px}
  .cat{font-size:10.5px;color:var(--muted);text-transform:uppercase;letter-spacing:.03em}
  .name{font-weight:600;font-size:13px;line-height:1.35;color:var(--fg);text-decoration:none}
  .name:hover{color:var(--accent)}
  .price{display:flex;align-items:baseline;gap:7px;flex-wrap:wrap}
  .sale{font-size:17px;font-weight:700;color:var(--sale)}
  .retail{font-size:12px;color:var(--muted);text-decoration:line-through}
  .mg{font-size:11px;color:#059669;font-weight:600}
  @media (prefers-color-scheme:dark){.mg{color:#34d399}}
  .badges{display:flex;flex-wrap:wrap;gap:5px;margin-top:2px}
  .bdg{font-size:11px;background:var(--chip);border-radius:6px;padding:2px 7px;color:var(--fg)}
  .bdg.none{opacity:.5}
  .more{font-size:11px;color:var(--accent);cursor:pointer;user-select:none}
  .ship-det{display:none;margin-top:4px}
  .ship-det.open{display:block}
  .ship-det table{width:100%;border-collapse:collapse;font-size:11px}
  .ship-det td{padding:2px 6px;border-top:1px solid var(--line)}
  .ship-det td.pr{text-align:right;white-space:nowrap;font-variant-numeric:tabular-nums}
  .ship-det .ch{font-weight:600;color:var(--accent);padding-top:5px}
  .empty{grid-column:1/-1;text-align:center;color:var(--muted);padding:40px}
</style></head><body>
<header>
  <h1>InterestPrint 제품 탐색기 — 디자인 상품 브레인스토밍</h1>
  <div class="sub" id="sub"></div>
  <div class="controls">
    <div class="fld"><label>검색(제품명)</label><input id="q" type="search" placeholder="예: mug, bag, pillow"></div>
    <div class="fld"><label>카테고리</label><select id="cat"></select></div>
    <div class="fld"><label>원가 최소($)</label><input id="pmin" type="number" step="0.5" min="0"></div>
    <div class="fld"><label>원가 최대($)</label><input id="pmax" type="number" step="0.5" min="0"></div>
    <div class="fld"><label>정렬</label><select id="sort">
      <option value="s">원가 낮은 순</option>
      <option value="-s">원가 높은 순</option>
      <option value="-m">마진율 높은 순</option>
      <option value="n">이름순</option>
    </select></div>
    <div class="fld"><label>&nbsp;</label><label style="font-weight:400;font-size:12px"><input id="shipOnly" type="checkbox"> 배송비 있는 것만</label></div>
    <div class="fld"><label>&nbsp;</label><label style="font-weight:400;font-size:12px"><input id="group" type="checkbox" checked> 카테고리별 그룹</label></div>
    <button class="btn" id="reset">초기화</button>
    <div class="count" id="count"></div>
  </div>
</header>
<div class="grid" id="grid"></div>
<script>
const DATA = __PAYLOAD__;
const COUNTRIES = __COUNTRIES__;
const PMIN = __PMIN__, PMAX = __PMAX__;

const $ = s => document.querySelector(s);
const grid = $("#grid");

// 카테고리 목록 + 개수
const catCount = {};
DATA.forEach(r => { catCount[r.c] = (catCount[r.c]||0)+1; });
const cats = Object.keys(catCount).sort((a,b)=>a.localeCompare(b));
const catSel = $("#cat");
catSel.innerHTML = '<option value="">전체 카테고리 ('+DATA.length+')</option>' +
  cats.map(c=>'<option value="'+c.replace(/"/g,'&quot;')+'">'+c+' ('+catCount[c]+')</option>').join('');

$("#pmin").value = Math.floor(PMIN);
$("#pmax").value = Math.ceil(PMAX);
$("#sub").textContent = '전체 '+DATA.length+'개 · 배송비 조회됨 '+DATA.filter(r=>Object.keys(r.sh).length).length+
  '개 · 원가 $'+PMIN.toFixed(2)+'~$'+PMAX.toFixed(2)+' · 조회국가 '+COUNTRIES.map(c=>c[1]).join(', ')+' (수량 1개)';

function esc(s){ return (s||"").replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])); }

function cardHTML(r){
  const retail = r.r!=null ? '<span class="retail">$'+r.r.toFixed(2)+'</span>' : '';
  const mg = r.m!=null ? '<span class="mg">마진 '+r.m.toFixed(0)+'%</span>' : '';
  const sale = r.s!=null ? '$'+r.s.toFixed(2) : '-';
  let badges='', det='';
  const codes = Object.keys(r.sh);
  if(codes.length){
    badges = COUNTRIES.map(([code,name])=>{
      const s=r.sh[code];
      return s ? '<span class="bdg">'+name.split(' ')[0]+' 최저 $'+s.min.toFixed(2)+'</span>'
               : '<span class="bdg none">'+name.split(' ')[0]+' -</span>';
    }).join('');
    det = '<div class="more" data-more>배송사별 보기 ▾</div><div class="ship-det"><table>' +
      COUNTRIES.map(([code,name])=>{
        const s=r.sh[code]; if(!s) return '';
        return '<tr><td class="ch" colspan="2">'+esc(name)+'</td></tr>' +
          s.opts.map(o=>'<tr><td>'+esc(o[0])+'</td><td class="pr">'+esc(o[1])+'</td></tr>').join('');
      }).join('') + '</table></div>';
  } else {
    badges = '<span class="bdg none">배송비 미조회</span>';
  }
  return '<div class="card">'+
    '<a href="'+esc(r.u)+'" target="_blank" rel="noopener"><img loading="lazy" src="'+esc(r.img)+'" alt=""></a>'+
    '<div class="body"><div class="cat">'+esc(r.c)+'</div>'+
    '<a class="name" href="'+esc(r.u)+'" target="_blank" rel="noopener">'+esc(r.n)+'</a>'+
    '<div class="price"><span class="sale">'+sale+'</span>'+retail+mg+'</div>'+
    '<div class="badges">'+badges+'</div>'+det+'</div></div>';
}

function render(){
  const q = $("#q").value.trim().toLowerCase();
  const cat = catSel.value;
  const lo = parseFloat($("#pmin").value); const hi = parseFloat($("#pmax").value);
  const sort = $("#sort").value;
  const shipOnly = $("#shipOnly").checked;
  const group = $("#group").checked;

  let rows = DATA.filter(r=>{
    if(cat && r.c!==cat) return false;
    if(q && !r.n.toLowerCase().includes(q)) return false;
    if(shipOnly && !Object.keys(r.sh).length) return false;
    if(!isNaN(lo) && (r.s==null || r.s<lo)) return false;
    if(!isNaN(hi) && (r.s==null || r.s>hi)) return false;
    return true;
  });
  const cmp = {
    's':(a,b)=>(a.s??1e9)-(b.s??1e9),
    '-s':(a,b)=>(b.s??-1)-(a.s??-1),
    '-m':(a,b)=>(b.m??-1)-(a.m??-1),
    'n':(a,b)=>a.n.localeCompare(b.n),
  }[sort];
  rows.sort(cmp);

  $("#count").innerHTML = '<b>'+rows.length+'</b>개 표시';
  if(!rows.length){ grid.innerHTML='<div class="empty">조건에 맞는 제품이 없습니다.</div>'; return; }

  if(group && !cat){
    const byCat={};
    rows.forEach(r=>{ (byCat[r.c]=byCat[r.c]||[]).push(r); });
    const order=Object.keys(byCat).sort((a,b)=>byCat[b].length-byCat[a].length || a.localeCompare(b));
    grid.innerHTML = order.map(c=>
      '<div class="sec-title">'+esc(c)+' — '+byCat[c].length+'개</div>'+
      byCat[c].map(cardHTML).join('')
    ).join('');
  } else {
    grid.innerHTML = rows.map(cardHTML).join('');
  }
}

grid.addEventListener('click', e=>{
  const m = e.target.closest('[data-more]');
  if(m){ const d=m.nextElementSibling; d.classList.toggle('open');
    m.textContent = d.classList.contains('open') ? '배송사별 접기 ▴' : '배송사별 보기 ▾'; }
});
["q","cat","pmin","pmax","sort","shipOnly","group"].forEach(id=>{
  const el=$("#"+id); el.addEventListener(el.tagName==='INPUT'&&el.type!=='checkbox'?'input':'change', render);
});
$("#reset").addEventListener('click', ()=>{
  $("#q").value=""; catSel.value=""; $("#pmin").value=Math.floor(PMIN); $("#pmax").value=Math.ceil(PMAX);
  $("#sort").value="s"; $("#shipOnly").checked=false; $("#group").checked=true; render();
});
render();
</script>
</body></html>"""
    doc = (doc.replace("__PAYLOAD__", payload)
              .replace("__COUNTRIES__", countries_js)
              .replace("__PMIN__", f"{pmin}")
              .replace("__PMAX__", f"{pmax}"))
    html_path.parent.mkdir(parents=True, exist_ok=True)
    html_path.write_text(doc, encoding="utf-8")
    print(f"✓ HTML(탐색기) 저장 → {html_path}  (전체 {len(recs)}개 / 배송비 {n_ship}개)")
