"""w-layout-to-excel: stage 4 - build the Excel layout from layout.json + rules.json.

Usage:
  python generate.py <layout.json> -r <rules.json> -o <output.xlsx>

Excel mapping:
  - one physical rack = one row (or column when orientation.transpose)
  - cells (货架组) = bordered cells with light-gray fill, containing all level codes
  - racks of the same block (back-to-back) are adjacent with NO blank row between
  - a blank thin row/column marks an aisle between different blocks
"""
import argparse
import json
import os
import statistics
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, PatternFill, Side
from openpyxl.utils import get_column_letter

DEFAULT_FILL = "D9D9D9"
DEFAULT_STROKE = "808080"


def drop_matches(rack, patterns):
    rid = rack["id"]
    rec = rack.get("record", rid)
    return any(rid == p or rec == p or rid.startswith(p + "-") for p in patterns)


def rack_columns(racks):
    cols = {}
    for r in racks:
        vertical = r.get("dir", "H") == "V"
        key = round(min(b["x0"] for b in r["bays"]) if vertical
                    else min(b["y0"] for b in r["bays"]))
        cols.setdefault(key, []).append(r)
    return [cols[k] for k in sorted(cols)]


def pnum(v, width):
    if not width or width <= 0:
        return str(v)
    return str(v).rjust(width, "0")


def hex_rgb(rgb, default):
    if not rgb:
        return default
    parts = [max(0, min(255, int(round(float(c) * 255)))) for c in rgb[:3]]
    return "".join(f"{p:02X}" for p in parts)


def order_racks(racks):
    for i, r in enumerate(racks):
        r.setdefault("block", r.get("id") or f"B{i + 1}")
    v_major = sum(1 for r in racks if r.get("dir", "H") == "V") > len(racks) / 2
    if v_major:
        key = lambda r: (min(b["x0"] for b in r["bays"]), min(b["y0"] for b in r["bays"]))
    else:
        key = lambda r: (min(b["y0"] for b in r["bays"]), min(b["x0"] for b in r["bays"]))
    blocks = {}
    for r in racks:
        blocks.setdefault(r["block"], []).append(r)
    ordered = []
    for _, members in sorted(blocks.items(), key=lambda kv: min(key(r) for r in kv[1])):
        ordered.extend(sorted(members, key=key))
    return ordered


def build_zone_lookup(rules):
    zones = rules.get("zones") or []
    exact = {}
    for z in zones:
        for pat in z.get("rack_ids", []):
            if pat != "*":
                exact[pat] = z
    catchall = next((z for z in zones if "*" in z.get("rack_ids", [])), None)

    def zone_of(rid):
        if rid in exact:
            return exact[rid]
        for pat, z in exact.items():
            if rid.startswith(pat + "-"):
                return z
        return catchall

    return zone_of


def assign_aisle_sides(racks_ordered, zone_of, continuous=True):
    """Aisle = real gap between neighbouring rack records (zones renumber from 01).

    Walk records of each zone in reading order, tracking the aisle counter `a`:
      - pair record  : left column  -> aisle a, EVEN cells (02,04,...)
                       right column -> aisle a+1, ODD cells (01,03,...); a += 1
      - single record: first in zone -> aisle a, ODD; otherwise -> aisle a, EVEN
    Cells number along the rack line in reading order (top->bottom for V).
    `continuous=False` restarts the sequence in each cross-aisle segment.
    Sets r["_aisle"] and bay["_seq"] (1-based position within the column).
    """
    records = {}
    for r in racks_ordered:
        rec = r.get("record", r["id"])
        records.setdefault(rec, []).append(r)
    for r in racks_ordered:
        zkey = (zone_of(r["id"]) or {}).get("prefix", "")
        r["_zone"] = zkey

    def rec_sort_key(rec):
        racks = records[rec]
        b = racks[0]["bays"][0]
        if racks[0].get("dir", "H") == "V":
            return (min(x["bays"][0]["x0"] for x in racks), b["y0"])
        return (b["y0"], min(x["bays"][0]["x0"] for x in racks))

    for zkey in {r["_zone"] for r in racks_ordered}:
        zone_recs = [rec for rec in records if {r["_zone"] for r in records[rec]} == {zkey}]
        zone_recs.sort(key=rec_sort_key)
        a = 1
        bad_cols = []
        for pos, rec in enumerate(zone_recs, 1):
            col_list = rack_columns(records[rec])
            if len(col_list) > 2:
                bad_cols.append(f"{rec}({len(col_list)}列)")
            if len(col_list) >= 2:
                plan = [(col_list[0], a, "even"), (col_list[-1], a + 1, "odd")]
                a += 1
            elif pos == 1:
                plan = [(col_list[0], a, "odd")]
            else:
                plan = [(col_list[0], a, "even")]
            for col_racks, aisle, side in plan:
                for rack in col_racks:
                    rack["_aisle"] = aisle
                if continuous:
                    col_bays = sorted((b for r in col_racks for b in r["bays"]),
                                      key=(lambda b: b["y0"]) if col_racks[0].get("dir", "H") == "V"
                                      else (lambda b: b["x0"]))
                    for i, bay in enumerate(col_bays, 1):
                        bay["_seq"] = 2 * i - 1 if side == "odd" else 2 * i
                else:
                    for rack in col_racks:
                        bays = sorted(rack["bays"],
                                      key=(lambda b: b["y0"]) if rack.get("dir", "H") == "V"
                                      else (lambda b: b["x0"]))
                        for i, bay in enumerate(bays, 1):
                            bay["_seq"] = 2 * i - 1 if side == "odd" else 2 * i
        if bad_cols:
            raise SystemExit("[generate] 错误: numbering_mode=aisle_sides 要求每条记录最多 "
                             f"2 条并列列（背靠背一对）；记录 {'、'.join(bad_cols)} 超出。"
                             "请在 layout.json 中把多余列拆成独立记录，或改用 rack_group_scope 计数")


def generate_codes(racks_ordered, rules):
    code = rules.get("code", {})
    segments = code.get("segments", ["zone", "aisle", "rack_group", "level"])
    sep = code.get("sep", "-")
    pads = code.get("pads", {})
    scope = code.get("rack_group_scope", "aisle")
    pos_count = max(1, int(code.get("position_count", 1) or 1))
    serp = code.get("serpentine", {}) or {}
    serp_on = bool(serp.get("enabled", False))
    corner = serp.get("start_corner", "top_left")
    first_dir = serp.get("first_dir", "right")

    levels_cfg = rules.get("levels", {"mode": "global", "count": 1})
    global_levels = int(levels_cfg.get("count", 1) or 1)
    per_rack_levels = levels_cfg.get("per_rack", {}) or {}

    zone_of = build_zone_lookup(rules)
    if rules.get("zones"):
        unmatched = sorted({r.get("record", r["id"]) for r in racks_ordered
                            if zone_of(r["id"]) is None})
        if unmatched:
            print(f"[generate] 警告: 以下货架未匹配任何分区（zone 段将为空）: "
                  f"{', '.join(unmatched)}", file=sys.stderr)
    aisle_sides = code.get("numbering_mode") == "aisle_sides"
    if aisle_sides:
        assign_aisle_sides(racks_ordered, zone_of,
                           continuous=code.get("continuous_segments", True))
        serp_on = False
    v_major = sum(1 for r in racks_ordered if r.get("dir", "H") == "V") > len(racks_ordered) / 2
    base_flip = (v_major and corner.startswith("bottom")) or \
                (not v_major and corner.endswith("right"))
    odd_rev = first_dir in ("right", "down")

    zone_block_no = {}
    cell_zone = {}
    cell_block = {}
    cell_rack = {}
    cell_global = 0

    for idx, r in enumerate(racks_ordered):
        rid = r["id"]
        z = zone_of(rid) or {}
        zkey = z.get("prefix", "")
        blk = r["block"]
        if not aisle_sides:
            zone_key = (zkey, blk)
            if zone_key not in zone_block_no:
                zone_block_no[zone_key] = sum(1 for k in zone_block_no if k[0] == zkey) + 1
            r["_zone"] = zkey
            r["_aisle"] = zone_block_no[zone_key]
        r["_rack"] = len([x for x in racks_ordered[:idx + 1]
                          if (zone_of(x["id"]) or {}).get("prefix", "") == zkey])

        bays = sorted(r["bays"],
                      key=(lambda b: b["y0"]) if r.get("dir", "H") == "V" else (lambda b: b["x0"]))
        if serp_on:
            parity_rev = (idx % 2 == 1) if odd_rev else (idx % 2 == 0)
            if base_flip != parity_rev:
                bays = bays[::-1]

        levels = per_rack_levels.get(rid, global_levels)
        if not isinstance(levels, int) or levels < 1:
            levels = 1

        for bay in bays:
            if aisle_sides:
                seq = bay.get("_seq", 1)
            else:
                cell_global += 1
                cell_zone[zkey] = cell_zone.get(zkey, 0) + 1
                cell_block[blk] = cell_block.get(blk, 0) + 1
                cell_rack[rid] = cell_rack.get(rid, 0) + 1
                seq = {"global": cell_global, "zone": cell_zone[zkey],
                       "aisle": cell_block[blk], "rack": cell_rack[rid]}.get(scope, cell_block[blk])
            codes = []
            for lv in range(1, levels + 1):
                for p in range(1, pos_count + 1):
                    parts = []
                    for s in segments:
                        if s == "zone":
                            parts.append(pnum(zkey, pads.get("zone", 0)) if zkey else "")
                        elif s == "aisle":
                            parts.append(pnum(r["_aisle"], pads.get("aisle", 2)))
                        elif s == "rack":
                            parts.append(pnum(r["_rack"], pads.get("rack", 2)))
                        elif s == "rack_group":
                            parts.append(pnum(seq, pads.get("rack_group", 2)))
                        elif s == "level":
                            parts.append(pnum(lv, pads.get("level", 0)))
                        elif s == "position":
                            parts.append(pnum(p, pads.get("position", 0)))
                    joined = sep.join(x for x in parts if x)
                    if joined:
                        codes.append(joined)
            bay["_codes"] = codes


def cell_specs(racks_ordered):
    specs = []
    prev_block = None
    for r in racks_ordered:
        if prev_block is not None and r["block"] != prev_block:
            specs.append(("aisle", None))
        prev_block = r["block"]
        specs.append(("rack", [{"codes": bay.get("_codes", []),
                                "fill": r.get("fill"),
                                "stroke": r.get("stroke")} for bay in r["bays"]]))
    return specs


def transpose_specs(specs):
    blocks = []
    cur = []
    for kind, row in specs:
        if kind == "rack":
            cur.append(row)
        elif cur:
            blocks.append(cur)
            cur = []
    if cur:
        blocks.append(cur)
    max_pos = max(len(row) for blk in blocks for row in blk)
    out = []
    for pos in range(max_pos):
        row = []
        for bi, blk in enumerate(blocks):
            if bi > 0:
                row.append({"aisle": True})
            for rack_row in blk:
                row.append(rack_row[pos] if pos < len(rack_row) else {"empty": True})
        out.append(("rack", row))
    return out


def mirror_specs(racks_ordered):
    """Mirror the drawing: one rack line = one sheet column (lines in x order,
    one blank aisle column between records); rows come from a global band grid
    over all bay positions, so cross-aisle gaps show as blank cells/rows."""
    v_major = sum(1 for r in racks_ordered if r.get("dir", "H") == "V") > len(racks_ordered) / 2
    if not v_major:
        raise SystemExit("[generate] 错误: orientation.mode=mirror 目前只支持竖排(V)货架"
                         "为主的仓库；横排仓库请用 classic 或 transpose")
    records = {}
    for r in racks_ordered:
        records.setdefault(r.get("record", r["id"]), []).append(r)

    def rec_key(rec):
        racks = records[rec]
        if v_major:
            return (min(b["x0"] for r in racks for b in r["bays"]),
                    min(b["y0"] for r in racks for b in r["bays"]))
        return (min(b["y0"] for r in racks for b in r["bays"]),
                min(b["x0"] for r in racks for b in r["bays"]))

    rec_order = sorted(records, key=rec_key)
    pos = (lambda b: b["y0"]) if v_major else (lambda b: b["x0"])
    diffs = []
    for r in racks_ordered:
        starts = sorted(pos(b) for b in r["bays"])
        diffs.extend(y - x for x, y in zip(starts, starts[1:]))
    pitch = statistics.median(diffs) if diffs else 24.0
    spacer_thr = 1.5 * pitch

    counts = {}
    for r in racks_ordered:
        for b in r["bays"]:
            counts[pos(b)] = counts.get(pos(b), 0) + 1
    radius = 0.55 * pitch
    seeds = []
    for y in sorted(counts.items(), key=lambda kv: (-kv[1], kv[0])):
        if all(abs(y[0] - s) > radius for s in seeds):
            seeds.append(y[0])
    seeds.sort()
    band_idx = {y: min(range(len(seeds)), key=lambda i: abs(seeds[i] - y))
                for y in counts}
    spacer_after = {i for i in range(len(seeds) - 1)
                    if seeds[i + 1] - seeds[i] > spacer_thr}
    row_of = {}
    extra = 0
    for i in range(len(seeds)):
        row_of[i] = i + extra
        if i in spacer_after:
            extra += 1
    n_rows = len(seeds) + extra

    col_info = []
    ncols = 0
    prev_rec = None
    for rec in rec_order:
        if prev_rec is not None:
            ncols += 1
        for col in rack_columns(records[rec]):
            col_info.append((ncols, col))
            ncols += 1
        prev_rec = rec

    grid = [[{"empty": True} for _ in range(ncols)] for _ in range(n_rows)]

    def ensure_rows(n):
        while len(grid) < n:
            grid.append([{"empty": True} for _ in range(ncols)])

    for cidx, col in col_info:
        col_bays = sorted(((pos(b), rack, b) for rack in col for b in rack["bays"]),
                          key=lambda t: t[0])
        prev_row = prev_y = None
        for _y, rack, bay in col_bays:
            natural = row_of[band_idx[_y]]
            if prev_row is None:
                row = natural
            elif _y - prev_y > spacer_thr:
                row = max(natural, prev_row + 2)
            else:
                row = prev_row + 1
            ensure_rows(row + 1)
            grid[row][cidx] = {
                "codes": bay.get("_codes", []),
                "fill": rack.get("fill"),
                "stroke": rack.get("stroke"),
            }
            prev_row, prev_y = row, _y

    specs = []
    for row_cells in grid:
        if any(not c.get("empty") for c in row_cells):
            specs.append(("rack", row_cells))
        else:
            specs.append(("aisle", None))
    return specs, ncols


def write_xlsx(specs, rules, out_path, max_cells):
    wb = Workbook()
    ws = wb.active
    ws.title = rules.get("sheet_name", "layout")
    style = rules.get("style", {}) or {}
    col_w = rules.get("cell_size", {}).get("col_width", 18)

    for rc, (kind, row) in enumerate(specs, 1):
        if kind == "aisle" or row is None:
            ws.row_dimensions[rc].height = 8
            continue
        max_codes = max((len(c.get("codes") or []) for c in row
                         if not c.get("aisle") and not c.get("empty")), default=1)
        ws.row_dimensions[rc].height = max(20, 15 * max_codes)
        for cc, spec in enumerate(row, 1):
            if spec.get("aisle") or spec.get("empty"):
                continue
            cell = ws.cell(row=rc, column=cc)
            cell.value = "\n".join(spec.get("codes") or []) or None
            cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
            stroke = style.get("stroke") or hex_rgb(spec.get("stroke"), DEFAULT_STROKE)
            side = Side(style="thin", color=stroke)
            cell.border = Border(left=side, right=side, top=side, bottom=side)
            fill = style.get("fill") or \
                (hex_rgb(spec["fill"], DEFAULT_FILL) if spec.get("fill") else DEFAULT_FILL)
            cell.fill = PatternFill("solid", fgColor=fill)

    for cc in range(1, max_cells + 1):
        ws.column_dimensions[get_column_letter(cc)].width = col_w
    ws.freeze_panes = "B2"
    wb.save(out_path)
    return out_path


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("layout_json")
    ap.add_argument("-r", "--rules", required=True)
    ap.add_argument("-o", "--out", default="layout.xlsx")
    ap.add_argument("--drop", default="",
                    help="剔除的记录/货架 id，逗号分隔（前缀匹配，如 R36,R05）")
    args = ap.parse_args()

    for path, what in ((args.layout_json, "layout"), (args.rules, "rules")):
        if not os.path.exists(path):
            raise SystemExit(f"[generate] 错误: {what} 文件不存在: {path}")
    with open(args.layout_json, encoding="utf-8") as f:
        layout = json.load(f)
    with open(args.rules, encoding="utf-8") as f:
        rules = json.load(f)

    racks = layout.get("racks", [])
    if args.drop:
        patterns = [p.strip() for p in args.drop.split(",") if p.strip()]
        before = len(racks)
        racks = [r for r in racks if not drop_matches(r, patterns)]
        print(f"[generate] --drop 已剔除 {before - len(racks)} 排"
              f"（模式: {', '.join(patterns)}）", file=sys.stderr)
    if not racks:
        raise SystemExit("layout.json 中没有货架排（可能被 --drop 全部剔除），无法生成 Excel。"
                         "请先运行 normalize.py 或人工补充 layout.json")
    empty = [r.get("id") for r in racks if not r.get("bays")]
    if empty:
        raise SystemExit(f"layout.json 中以下货架没有格子: {', '.join(empty)}")

    racks_ordered = order_racks(racks)
    generate_codes(racks_ordered, rules)

    specs = cell_specs(racks_ordered)
    max_cells = max((len(row) for kind, row in specs if kind == "rack"), default=0)
    orient = rules.get("orientation", {}) or {}
    if orient.get("mode") == "mirror":
        specs, max_cells = mirror_specs(racks_ordered)
    elif orient.get("transpose", False):
        specs = transpose_specs(specs)
        max_cells = max((len(row) for kind, row in specs), default=0)

    out = write_xlsx(specs, rules, args.out, max_cells)
    print(json.dumps({
        "ok": True,
        "output": os.path.abspath(out),
        "n_blocks": len({r["block"] for r in racks_ordered}),
        "n_racks": len(racks_ordered),
        "n_cells": sum(len(r["bays"]) for r in racks_ordered),
    }, ensure_ascii=False))


if __name__ == "__main__":
    main()
