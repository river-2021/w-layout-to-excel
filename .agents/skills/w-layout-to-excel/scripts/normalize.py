"""w-layout-to-excel: stage 1 - normalize GPT layout JSON into the standard layout.json.

Usage:
  python normalize.py <gpt6.json> -o <layout.json> [--strict]

Input variants (auto-detected):
  grouped: {"rack_positions": [{"id", "x_ranges": [[a,b],...], "y_ranges": [...]}]}
           every x_range x y_range combination is one cell (cartesian product)
  flat:    {"racks": [{"id", "cells": [{"x0","y0","x1","y1"}, ...]}]}

Output layout.json (consumed by generate.py):
  racks[]: {id, record, dir, block, confidence, notes, bays[]}
  - back-to-back columns (tiny depth gap) are split into physical racks sharing one block
  - large gaps along the rack line (cross aisles) split into separate blocks
  - low-confidence records / excluded regions are surfaced in notes + needs_review
"""
import argparse
import json
import os
import statistics

LOW_CONF_MARKERS = ("较低", "低", "不确定")
LOW_CONF_VALUE = 0.45


def fail(msg):
    raise SystemExit(f"[normalize] 错误: {msg}")


def load_json(path):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        fail(f"文件不存在: {path}")
    except json.JSONDecodeError as e:
        fail(f"{path} 不是合法 JSON（行 {e.lineno} 列 {e.colno}: {e.msg}）")


def norm_conf(v):
    if v is None or isinstance(v, bool):
        return 1.0
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        if any(m in v for m in LOW_CONF_MARKERS):
            return LOW_CONF_VALUE
        try:
            return float(v)
        except ValueError:
            return 1.0
    return 1.0


def parse_ranges(pairs, what, rid):
    if not isinstance(pairs, list) or not pairs:
        fail(f"记录 {rid}: {what} 缺失或为空（应为 [[start,end], ...]）")
    out = []
    for p in pairs:
        ok = (isinstance(p, (list, tuple)) and len(p) == 2
              and all(isinstance(c, (int, float)) and not isinstance(c, bool) for c in p))
        if not ok:
            fail(f"记录 {rid}: {what} 含非法区间 {p!r}，应为 [start, end] 两个数字")
        a, b = float(p[0]), float(p[1])
        if a >= b:
            fail(f"记录 {rid}: {what} 区间起点必须小于终点: {p!r}")
        out.append((a, b))
    out.sort()
    for i in range(1, len(out)):
        if out[i][0] < out[i - 1][1] - 1e-6:
            return out, True
    return out, False


def median_or(xs, default):
    return statistics.median(xs) if xs else default


def split_by_gap(ranges, thr):
    groups = [[ranges[0]]]
    for r in ranges[1:]:
        if r[0] - groups[-1][-1][1] > thr:
            groups.append([r])
        else:
            groups[-1].append(r)
    return groups


def interval_clusters(intervals, tol):
    clusters = []
    for iv in sorted(intervals):
        for c in clusters:
            if abs(c[0][0] - iv[0]) <= tol and abs(c[0][1] - iv[1]) <= tol:
                c.append(iv)
                break
        else:
            clusters.append([iv])
    return clusters


def flat_record_to_ranges(rec, rid):
    cells = rec.get("cells")
    if not isinstance(cells, list) or not cells:
        fail(f"记录 {rid}: cells 缺失或为空（应为 [{{x0,y0,x1,y1}}, ...]）")
    xs, ys = [], []
    for c in cells:
        if not isinstance(c, dict):
            fail(f"记录 {rid}: cell {c!r} 不是对象")
        try:
            x0, y0, x1, y1 = (float(c[k]) for k in ("x0", "y0", "x1", "y1"))
        except (KeyError, TypeError, ValueError):
            fail(f"记录 {rid}: cell 缺少 x0/y0/x1/y1 数字字段: {c!r}")
        if x0 >= x1 or y0 >= y1:
            fail(f"记录 {rid}: cell 坐标非法（需 x0<x1, y0<y1）: {c!r}")
        xs.append((x0, x1))
        ys.append((y0, y1))
    x_tol = max(1.0, 0.2 * median_or([b - a for a, b in xs], 1.0))
    y_tol = max(1.0, 0.2 * median_or([b - a for a, b in ys], 1.0))
    xr = sorted((min(a for a, _ in c), max(b for _, b in c)) for c in interval_clusters(xs, x_tol))
    yr = sorted((min(a for a, _ in c), max(b for _, b in c)) for c in interval_clusters(ys, y_tol))
    centers = [((a + b) / 2, (c + d) / 2) for (a, b), (c, d) in zip(xs, ys)]
    return xr, yr, centers


def process_record(rec, index, notes):
    rid = rec.get("id")
    if not rid or not isinstance(rid, str):
        fail(f"第 {index} 条货架记录缺少字符串 id")
    centers = None
    if "cells" in rec:
        xr, yr, centers = flat_record_to_ranges(rec, rid)
    else:
        xr, overlap_x = parse_ranges(rec.get("x_ranges"), "x_ranges", rid)
        yr, overlap_y = parse_ranges(rec.get("y_ranges"), "y_ranges", rid)
        if overlap_x:
            notes.append(f"{rid}: x_ranges 存在重叠，已按起点排序处理")
        if overlap_y:
            notes.append(f"{rid}: y_ranges 存在重叠，已按起点排序处理")

    span_x = xr[-1][1] - xr[0][0]
    span_y = yr[-1][1] - yr[0][0]
    if len(yr) > len(xr):
        direction = "V"
    elif len(xr) > len(yr):
        direction = "H"
    else:
        direction = "H" if span_x >= span_y else "V"

    length, depth = (yr, xr) if direction == "V" else (xr, yr)
    starts = [a for a, _ in length]
    pitch = median_or([b - a for a, b in zip(starts, starts[1:])],
                      median_or([b - a for a, b in length], 10.0))
    length_thr = max(2.0, 1.15 * pitch)
    segs = split_by_gap(length, length_thr)
    depth_width = median_or([b - a for a, b in depth], 10.0)
    depth_thr = max(2.0, depth_width)

    conf = norm_conf(rec.get("confidence"))
    review = rec.get("review_reason")
    racks = []
    k = 0
    for si, seg in enumerate(segs):
        dgroups = split_by_gap(depth, depth_thr)
        for bi, dg in enumerate(dgroups):
            block = f"{rid}-s{si + 1}" if len(dgroups) == 1 else f"{rid}-s{si + 1}b{bi + 1}"
            for dr in dg:
                k += 1
                bays = []
                for a, b in seg:
                    if direction == "V":
                        bays.append({"x0": dr[0], "y0": a, "x1": dr[1], "y1": b})
                    else:
                        bays.append({"x0": a, "y0": dr[0], "x1": b, "y1": dr[1]})
                dims = [(bay["y1"] - bay["y0"]) if direction == "V"
                        else (bay["x1"] - bay["x0"]) for bay in bays]
                med = median_or(dims, 0.0) or 1.0
                n_bad = sum(1 for d in dims if d < 0.5 * med or d > 1.6 * med)
                racks.append({
                    "id": f"{rid}-{k}",
                    "record": rid,
                    "dir": direction,
                    "block": block,
                    "confidence": conf,
                    "notes": [f"{n_bad} 个格子尺寸异常"] if n_bad else [],
                    "bays": bays,
                })
    if centers is not None:
        dropped = 0
        for rack in racks:
            kept = []
            for bay in rack["bays"]:
                cx = (bay["x0"] + bay["x1"]) / 2
                cy = (bay["y0"] + bay["y1"]) / 2
                tolx = max(3.0, 0.35 * (bay["x1"] - bay["x0"]))
                toly = max(3.0, 0.35 * (bay["y1"] - bay["y0"]))
                if any(abs(cx - rx) <= tolx and abs(cy - ry) <= toly for rx, ry in centers):
                    kept.append(bay)
                else:
                    dropped += 1
            rack["bays"] = kept
        racks = [r for r in racks if r["bays"]]
        if dropped:
            notes.append(f"{rid}: 分组展开出 {dropped} 个图中不存在的格子，已剔除")
    if k == 1 and racks:
        racks[0]["id"] = rid
    if conf < 0.5:
        notes.append(f"{rid}: 置信度低（{review}）" if review else f"{rid}: 置信度低")
    elif review:
        notes.append(f"{rid}: {review}")
    if rec.get("last_y_range_truncated"):
        notes.append(f"{rid}: 末端格子被图像边缘截断")
    return racks


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input", help="GPT 输出的货架 JSON")
    ap.add_argument("-o", "--out", default="layout.json", help="输出 layout.json 路径")
    ap.add_argument("--strict", action="store_true", help="把警告当错误处理")
    args = ap.parse_args()

    data = load_json(args.input)
    if not isinstance(data, dict):
        fail("顶层必须是 JSON 对象")

    if isinstance(data.get("rack_positions"), list) and data["rack_positions"]:
        records, fmt = data["rack_positions"], "grouped"
    elif (isinstance(data.get("racks"), list) and data["racks"]
          and isinstance(data["racks"][0], dict) and "cells" in data["racks"][0]):
        records, fmt = data["racks"], "flat"
    else:
        fail("无法识别输入格式：需要 rack_positions[]（x_ranges/y_ranges 分组）"
             "或 racks[].cells（扁平 bbox），参见 reference/gpt6_prompt.md")

    notes = []
    image = data.get("image_size") if isinstance(data.get("image_size"), dict) else {}
    img_w, img_h = image.get("width"), image.get("height")
    all_racks = []
    ids = set()
    for i, rec in enumerate(records, 1):
        if not isinstance(rec, dict):
            fail(f"第 {i} 条货架记录不是对象: {rec!r}")
        racks = process_record(rec, i, notes)
        for r in racks:
            if r["id"] in ids:
                notes.append(f"货架 id 重复: {r['id']}")
            ids.add(r["id"])
            if img_w or img_h:
                xs = [b["x0"] for b in r["bays"]] + [b["x1"] for b in r["bays"]]
                ys = [b["y0"] for b in r["bays"]] + [b["y1"] for b in r["bays"]]
                if (img_w and (min(xs) < -0.5 or max(xs) > img_w + 0.5)
                        or img_h and (min(ys) < -0.5 or max(ys) > img_h + 0.5)):
                    notes.append(f"{r['id']}: 坐标超出图像范围")
        all_racks.extend(racks)
    if not all_racks:
        fail("没有解析出任何货架，请检查输入 JSON")

    excluded = data.get("excluded_regions") or []
    for ex in excluded:
        if isinstance(ex, dict) and ex.get("reason"):
            notes.append(f"排除区域: {ex['reason']}")

    if args.strict:
        problems = [n for n in notes]
        if problems:
            for n in problems:
                print(f"[normalize] 警告: {n}", file=os.sys.stderr)
            fail(f"--strict 模式：存在 {len(problems)} 条警告，请修正后重试")

    needs_review = []
    seen = set()
    for r in all_racks:
        rec_id = r.get("record", r["id"])
        if r["confidence"] < 0.5 and rec_id not in seen:
            seen.add(rec_id)
            needs_review.append(rec_id)

    layout = {
        "source": "gpt6",
        "source_file": os.path.abspath(args.input),
        "format": fmt,
        "image_size": image or None,
        "confidence": round(min((r["confidence"] for r in all_racks), default=1.0), 3),
        "needs_review": needs_review,
        "notes": notes,
        "n_records": len(records),
        "n_racks": len(all_racks),
        "n_blocks": len({r["block"] for r in all_racks}),
        "n_cells": sum(len(r["bays"]) for r in all_racks),
        "excluded_regions": excluded,
        "racks": all_racks,
    }
    out_dir = os.path.dirname(os.path.abspath(args.out))
    os.makedirs(out_dir, exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(layout, f, ensure_ascii=False, indent=2)

    print(json.dumps({
        "ok": True,
        "output": os.path.abspath(args.out),
        "format": fmt,
        "n_records": layout["n_records"],
        "n_racks": layout["n_racks"],
        "n_blocks": layout["n_blocks"],
        "n_cells": layout["n_cells"],
        "confidence": layout["confidence"],
        "needs_review": needs_review,
        "n_notes": len(notes),
    }, ensure_ascii=False))


if __name__ == "__main__":
    main()
