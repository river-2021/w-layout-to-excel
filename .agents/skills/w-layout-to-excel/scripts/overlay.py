"""w-layout-to-excel: overlay check page - render layout.json onto the drawing.

Usage:
  python overlay.py <layout.json> --image <png> -o <check.html> \
      [--rules rules.json] [--drop R36,...] [--no-open]

Without --rules: record boxes colored by confidence (blue=ok, red dashed=low).
With --rules: full numbering runs first (reuses generate.py logic), records are
colored per zone, every bay shows its rack_group number, hover shows full code.
The check page is opened in the default browser unless --no-open is given.
"""
import argparse
import json
import os
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import generate as gen

ZONE_COLORS = ["#2196F3", "#FF9800", "#4CAF50", "#9C27B0",
               "#00BCD4", "#795548", "#E91E63", "#3F51B5"]

HTML = """<!DOCTYPE html>
<html lang="zh">
<head>
<meta charset="utf-8">
<title>布局识别核对</title>
<style>
  body { font-family: -apple-system, sans-serif; margin: 20px; background: #f5f5f5; }
  h2 { margin: 0 0 8px; font-size: 18px; }
  .legend { margin-bottom: 12px; font-size: 13px; }
  .legend label { margin-right: 16px; cursor: pointer; user-select: none; }
  .sw { display: inline-block; width: 14px; height: 14px; vertical-align: -2px;
        margin-right: 4px; border-radius: 3px; }
  .wrap { position: relative; display: inline-block; max-width: 100%;
          box-shadow: 0 2px 8px rgba(0,0,0,.15); }
  .wrap img { display: block; width: __W__px; max-width: 100%; }
  .wrap svg { position: absolute; inset: 0; width: 100%; height: 100%; }
  .rec { fill: rgba(33,150,243,.08); stroke: #2196F3; stroke-width: 2; }
  .lab { font: 700 15px sans-serif; fill: #1565C0; paint-order: stroke;
         stroke: #fff; stroke-width: 3px; }
  .low .rec { fill: rgba(244,67,54,.12); stroke: #F44336; stroke-width: 3;
              stroke-dasharray: 6 4; }
  .low .lab { fill: #C62828; }
  .bay { fill: rgba(33,150,243,.15); stroke: #1565C0; stroke-width: 1; }
  .code { font: 700 7px sans-serif; fill: #222; text-anchor: middle;
          paint-order: stroke; stroke: #fff; stroke-width: 2px; }
  .exc { fill: rgba(255,193,7,.15); stroke: #FFC107; stroke-width: 1.5;
         stroke-dasharray: 3 3; }
  .off { display: none; }
  .tip { font-size: 12px; color: #666; margin-top: 8px; }
</style>
</head>
<body>
<h2>__TITLE__</h2>
<div class="legend">__LEGEND__</div>
<div class="wrap"><img src="__IMG__"><svg id="svg" viewBox="0 0 __W__ __H__"
  preserveAspectRatio="none"></svg></div>
<div class="tip">__TIP__</div>
<script>
const DATA = __DATA__;
const svg = document.getElementById('svg');
const NS = 'http://www.w3.org/2000/svg';
function el(n, at, parent) {
  const e = document.createElementNS(NS, n);
  for (const k in at) e.setAttribute(k, at[k]);
  (parent || svg).appendChild(e); return e;
}
DATA.records.forEach(r => {
  const g = el('g', {class: r.kind + (r.zone ? ' z' + r.zone : '')});
  const [x0, y0, x1, y1] = r.box;
  const head = el('g', {}, g);
  el('rect', {x: x0, y: y0, width: x1 - x0, height: y1 - y0, class: 'rec'}, head);
  const lab = el('text', {x: x0 + 2, y: y0 - 4 < 8 ? y0 + 14 : y0 - 4, class: 'lab'}, head);
  lab.textContent = r.id + (r.zone ? '·' + r.zone : '');
  el('title', {}, head).textContent =
    `${r.id}${r.zone ? '（区 ' + r.zone + '）' : ''}：${r.racks} 排 / ${r.cells} 格` +
    (r.kind === 'low' ? ' [低置信]' : '');
  if (r.color) {
    head.querySelector('rect').setAttribute('style',
      `stroke:${r.color};fill:${r.color}22`);
    lab.setAttribute('style', `fill:${r.color}`);
  }
  const gb = el('g', {class: 'bays' + (DATA.with_codes ? '' : ' off')}, g);
  const gc = el('g', {class: 'codes' + (DATA.with_codes ? '' : ' off')}, g);
  r.bays.forEach(b => {
    const br = el('rect', {x: b.x0, y: b.y0, width: b.x1 - b.x0,
      height: b.y1 - b.y0, class: 'bay'}, gb);
    if (b.code) el('title', {}, br).textContent = b.code;
    if (b.lab) {
      const t = el('text', {x: (b.x0 + b.x1) / 2, y: (b.y0 + b.y1) / 2 + 2.5,
        class: 'code'}, gc);
      t.textContent = b.lab;
      if (b.code) el('title', {}, t).textContent = b.code;
    }
  });
});
const gExc = el('g', {class: 'exc-g off'});
DATA.excluded.forEach(x => {
  if (!Array.isArray(x.bbox)) return;
  const [x0, y0, x1, y1] = x.bbox;
  const g = el('rect', {x: x0, y: y0, width: x1 - x0, height: y1 - y0, class: 'exc'}, gExc);
  el('title', {}, g).textContent = '排除: ' + (x.reason || '');
});
document.querySelectorAll('.legend input').forEach(cb => {
  cb.addEventListener('change', () => {
    const t = cb.dataset.t;
    if (t === 'exc') gExc.classList.toggle('off', !cb.checked);
    else if (t === 'img')
      document.querySelector('.wrap img').style.visibility = cb.checked ? '' : 'hidden';
    else document.querySelectorAll('svg .' + t).forEach(
      g => g.classList.toggle('off', !cb.checked));
  });
});
</script>
</body>
</html>"""


def open_check_page(path):
    try:
        if sys.platform == "darwin":
            cmd = ["open", path]
        elif os.name == "nt":
            os.startfile(path)
            return True
        else:
            cmd = ["xdg-open", path]
        subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return True
    except Exception as exc:
        print(f"[overlay] 警告: 无法自动打开核对页（可手动打开）: {exc}",
              file=sys.stderr)
        return False


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("layout_json")
    ap.add_argument("--image", required=True, help="底图 PNG/JPG 路径")
    ap.add_argument("-o", "--out", default="check.html")
    ap.add_argument("--rules", help="rules.json：开启编号预览（分区配色+每格货架组号）")
    ap.add_argument("--drop", default="", help="剔除的记录 id，逗号分隔（前缀匹配）")
    ap.add_argument("--no-open", action="store_true", help="生成后不自动打开核对页")
    args = ap.parse_args()

    for path, what in ((args.layout_json, "layout"), (args.image, "底图")):
        if not os.path.exists(path):
            raise SystemExit(f"[overlay] 错误: {what} 文件不存在: {path}")
    with open(args.layout_json, encoding="utf-8") as f:
        layout = json.load(f)
    racks = layout.get("racks", [])
    if not racks:
        raise SystemExit("[overlay] 错误: layout.json 中没有货架")
    if args.drop:
        patterns = [p.strip() for p in args.drop.split(",") if p.strip()]
        racks = [r for r in racks if not gen.drop_matches(r, patterns)]
    if not racks:
        raise SystemExit("[overlay] 错误: --drop 后没有剩余货架")

    rules = None
    if args.rules:
        if not os.path.exists(args.rules):
            raise SystemExit(f"[overlay] 错误: rules 文件不存在: {args.rules}")
        with open(args.rules, encoding="utf-8") as f:
            rules = json.load(f)

    zone_color = {}
    zone_of = None
    rg_idx = -1
    sep = "-"
    if rules:
        racks_ordered = gen.order_racks(racks)
        gen.generate_codes(racks_ordered, rules)
        zone_of = gen.build_zone_lookup(rules)
        prefixes = sorted({(zone_of(r["id"]) or {}).get("prefix", "") for r in racks_ordered})
        for i, p in enumerate(prefixes):
            zone_color[p] = ZONE_COLORS[i % len(ZONE_COLORS)]
        code_cfg = rules.get("code", {})
        sep = code_cfg.get("sep", "-")
        rg_idx = code_cfg.get("segments", []).index("rack_group") \
            if "rack_group" in code_cfg.get("segments", []) else -1

    records = {}
    order = []
    for r in racks:
        rec = r.get("record", r["id"])
        if rec not in records:
            records[rec] = {"bays": [], "conf": 1.0, "racks": 0}
            order.append(rec)
        e = records[rec]
        e["bays"].extend(r["bays"])
        e["racks"] += 1
        e["conf"] = min(e["conf"], r.get("confidence", 1.0))

    rec_data = []
    for rec in order:
        e = records[rec]
        zone = ""
        if zone_of is not None:
            zone = (zone_of(rec) or {}).get("prefix", "")
        bays = []
        for b in e["bays"]:
            codes = b.get("_codes") or []
            full = "\n".join(codes)
            lab = ""
            if codes and rg_idx >= 0:
                parts = codes[0].split(sep)
                if rg_idx < len(parts):
                    lab = parts[rg_idx]
            bays.append({"x0": b["x0"], "y0": b["y0"], "x1": b["x1"], "y1": b["y1"],
                         "code": full, "lab": lab})
        xs0 = [b["x0"] for b in e["bays"]]; xs1 = [b["x1"] for b in e["bays"]]
        ys0 = [b["y0"] for b in e["bays"]]; ys1 = [b["y1"] for b in e["bays"]]
        rec_data.append({
            "id": rec,
            "box": [min(xs0), min(ys0), max(xs1), max(ys1)],
            "racks": e["racks"],
            "cells": len(e["bays"]),
            "kind": "low" if e["conf"] < 0.5 else "ok",
            "zone": zone,
            "color": zone_color.get(zone),
            "bays": bays,
        })

    size = layout.get("image_size") or {}
    w = size.get("width") or max(b["x1"] for r in racks for b in r["bays"])
    h = size.get("height") or max(b["y1"] for r in racks for b in r["bays"])

    with_codes = rules is not None
    zone_html = "".join(
        f'<span style="margin-right:14px"><span class="sw" style="background:{zone_color[z] or "#bbb"}"></span>{z or "(未分区)"}区</span>'
        for z in sorted(zone_color)) if with_codes else ""
    legend = '<label><input type="checkbox" data-t="img" checked><span class="sw" style="background:repeating-linear-gradient(45deg,#ddd,#ddd 2px,#fff 2px,#fff 4px)"></span>底图</label>' + zone_html + (
        '<label><input type="checkbox" data-t="codes" checked><span class="sw" style="background:#666"></span>编号预览</label>'
        if with_codes else
        '<label><input type="checkbox" data-t="ok" checked><span class="sw" style="background:#2196F3"></span>正常货架</label>'
        '<label><input type="checkbox" data-t="low" checked><span class="sw" style="background:#F44336"></span>低置信</label>'
    ) + f'<label><input type="checkbox" data-t="bays"{" checked" if with_codes else ""}><span class="sw" style="background:#90CAF9"></span>格子填充</label>' \
        '<label><input type="checkbox" data-t="exc"><span class="sw" style="background:#FFC107"></span>排除区域</label>'

    data = {
        "with_codes": with_codes,
        "records": rec_data,
        "excluded": layout.get("excluded_regions") or [],
    }
    html = (HTML.replace("__TITLE__", "布局识别核对" + (" · 编号预览" if with_codes else ""))
            .replace("__LEGEND__", legend)
            .replace("__IMG__", os.path.abspath(args.image))
            .replace("__W__", str(int(w))).replace("__H__", str(int(h)))
            .replace("__TIP__", ("每格小字=货架组号，hover 看完整货位码；外框颜色=分区。"
                                 if with_codes else
                                 "蓝=正常货架；红虚线=低置信记录；黄=GPT 排除区域。"))
            .replace("__DATA__", json.dumps(data, ensure_ascii=False)))

    out_dir = os.path.dirname(os.path.abspath(args.out))
    os.makedirs(out_dir, exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as f:
        f.write(html)
    out_path = os.path.abspath(args.out)
    opened = False if args.no_open else open_check_page(out_path)
    print(json.dumps({
        "ok": True,
        "output": out_path,
        "opened": opened,
        "records": len(rec_data),
        "cells": sum(r["cells"] for r in rec_data),
        "zones": sorted(zone_color) if with_codes else [],
    }, ensure_ascii=False))


if __name__ == "__main__":
    main()
