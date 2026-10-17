# -*- coding: utf-8 -*-
"""
VIN 查询工具生成器
读取车型库 Excel，把数据嵌入单文件 HTML 查询工具。
车型库更新后，重新运行本脚本即可刷新网页里的数据。

用法：
    python build_vin_tool.py [车型库.xlsx路径]
默认路径：D:\\中国车\\Oscar中国车主表.xlsx
输出：同目录下 VIN车型查询工具.html
"""
import json
import re
import sys
from pathlib import Path

import openpyxl
import dbstore  # SQLite 数据层：build 时把期初数据落入 DB（web 修改经 API 写回）

DEFAULT_XLSX = r"D:\中国车\Oscar中国车主表.xlsx"
HERE = Path(__file__).parent
TEMPLATE = HERE / "template.html"
OUTPUT = HERE / "VIN车型查询工具.html"
INDEX = HERE / "index.html"  # 部署用（端口 3001 服务默认页）
PRODUCT_XLSX = r"D:\中国车\中国车产品数据表.xlsx"  # 产品表：按车型分表，sheet 名 = 车型库 vehicle_id（即 cid / VID）

HEADERS = ["vehicle_id", "vehicle_description", "make", "model", "submodel", "year",
           "capacity", "kw", "engine", "chassis", "series", "vin", "body_type", "fuel", "drive"]


def parse_year_range(s):
    """把 '01/2017 ~ 01/2020' / '2021 ~ on' / '1951 ~ 1962' 解析为 (startYear, endYear|None)"""
    if not s:
        return (None, None)
    s = str(s)
    years = [int(y) for y in re.findall(r"(?<!\d)(19\d{2}|20\d{2})(?!\d)", s)]
    open_ended = bool(re.search(r"~\s*on\b", s, re.I)) or s.strip().lower().endswith("on")
    if not years:
        return (None, None)
    start = years[0]
    end = None
    if open_ended:
        end = None
    elif len(years) > 1:
        end = years[-1]
    else:
        end = start
    return (start, end)


def load_records(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    records, skipped = [], []
    for ws in wb.worksheets:
        for row in ws.iter_rows(min_row=2, values_only=True):
            row = list(row) + [None] * (len(HEADERS) - len(row))
            vid, desc, make, model, submodel, year, capacity, kw, engine, chassis, series, vin, body, fuel, drive = row[:15]
            vin_raw = str(vin or "").strip()
            # 去掉 'Orig: CHINA' 之类后缀，取前缀
            prefix = vin_raw.split("Orig")[0].strip().upper()
            origin = ""
            m = re.search(r"Orig[:\s]*([A-Za-z ]+)", vin_raw)
            if m:
                origin = m.group(1).strip()
            prefix = "".join(ch for ch in prefix if ch.isalnum())
            if not prefix or prefix == ".":
                skipped.append({"brand": make, "model": model, "sub": submodel, "year": year})
                continue
            ys, ye = parse_year_range(year)
            records.append({
                "prefix": prefix,
                "brand": str(make or "").strip(),
                "model": str(model or "").strip(),
                "sub": str(submodel or "").strip(),
                "year": str(year or "").strip(),
                "ys": ys,
                "ye": ye,
                "cap": str(capacity or "").strip(),
                "kw": str(kw or "").strip(),
                "engine": str(engine or "").strip(),
                "chassis": str(chassis or "").strip() if str(chassis or "").strip() != "." else "",
                "series": str(series or "").strip() if str(series or "").strip() != "." else "",
                "body": str(body or "").strip(),
                "fuel": str(fuel or "").strip(),
                "drive": str(drive or "").strip(),
                "vid": vid,
                "origin": origin,
            })
    wb.close()
    return records, skipped


# 库内确有错填的 VIN 前缀（已用真实 17 位 VIN 样本核实，来源 17vin / 实车/公告）
# key = vehicle_id, value = (错误前缀, 正确前缀)。Excel 中 vehicle_id 通常以字符串存储，故此处用 str 键；
# apply_corrections 同时兼容 int 键。仅当当前前缀等于错误值时才修正，可重复运行。
PREFIX_CORRECTIONS = {
    "52583": ("LSKG4AL1", "LSKG4GL1"),   # LDV G10 真实前缀为 LSKG4GL1（LSKG4AL1 是 MIFA 的）
    "55514": ("LSKG4AL1", "LSKG4GL1"),
    "56574": ("LSKG4AL1", "LSKG4GL1"),
    "61285": ("LSKG5GL1", "LSH14J7C"),   # LDV DELIVER 9 真实前缀为 LSH14J7C（LSKG5GL1 是 V80 的）
}

# 是否把"共享同一 VDS 前缀（厂家同一车辆特征码）的库内记录合并为一条"。
# True  = 合并，查询时同前缀只返回一款（适合"查到即唯一"场景）
# False = 不合并，查询时同前缀的多款车型全部作为候选列出（适合需区分电池/配置的场景）
MERGE_SAME_VDS = False


def apply_corrections(records):
    for r in records:
        vid = r.get("vid")
        c = PREFIX_CORRECTIONS.get(str(vid)) if vid is not None else None
        if c is None and isinstance(vid, int):
            c = PREFIX_CORRECTIONS.get(vid)
        if c and r["prefix"] == c[0]:
            r["prefix"] = c[1]
            r["origin"] = (r.get("origin") or "") + " [前缀已校正]"
    return records


def merge_by_prefix(records):
    """把共享同一 VDS 前缀（即厂家同一车辆特征码）的库内记录合并为一条。
    这样每个 VIN 前缀只对应唯一一款车，查询返回单一结果。"""
    from collections import defaultdict
    groups = defaultdict(list)
    for r in records:
        groups[r["prefix"]].append(r)
    out = []
    for prefix, recs in groups.items():
        if len(recs) == 1:
            out.append(recs[0])
            continue
        ys = [r["ys"] for r in recs if r["ys"]]
        yes = [r["ye"] for r in recs if r["ye"] is not None]
        ys_v = min(ys) if ys else None
        ye_v = max(yes) if yes else None
        if ys_v and ye_v is not None:
            year = f"{ys_v} ~ {ye_v}"
        elif ys_v and ye_v is None:
            year = f"{ys_v} ~ on"
        elif ys_v:
            year = str(ys_v)
        else:
            year = ""

        def join(field):
            vals = []
            for r in recs:
                v = (r.get(field) or "").strip()
                if v and v != "." and v not in vals:
                    vals.append(v)
            return " / ".join(vals)

        m = recs[0]
        out.append({
            "prefix": prefix,
            "brand": m["brand"],
            "model": join("model"),
            "sub": join("sub"),
            "year": year,
            "ys": ys_v, "ye": ye_v,
            "cap": join("cap"),
            "kw": join("kw"),
            "engine": join("engine"),
            "chassis": join("chassis"),
            "series": join("series"),
            "body": join("body"),
            "fuel": join("fuel"),
            "drive": join("drive"),
            "vid": m["vid"],
            "origin": f"合并自 {len(recs)} 条同VDS记录",
        })
    return out


def load_products(xlsx_path):
    """读取产品表（D:\\中国车\\中国车产品数据表.xlsx）。
    该文件按车型分表：每个车型一个 sheet，sheet 名 = 车型库 vehicle_id（即 cid / VID）。
    每个分表内有一段零件表，列含 品名(中文) / Part(英文) / Brand / CDAP No / OE Part No / 供应商 等。
    返回 { vid(str): [ {cn,en,brand,cdap,oe,notes,supplier,cost,moq,grp}, ... ] }，仅含以 VID 命名的数字分表。"""
    import os
    def cell(r, j):
        v = r[j] if r and len(r) > j else None
        return "" if v is None else str(v).strip()
    if not os.path.exists(xlsx_path):
        print(f"[产品] 找不到产品表: {xlsx_path}，跳过产品注入")
        return {}
    try:
        wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    except Exception as e:
        print(f"[产品] 打开产品表失败: {e}，跳过产品注入")
        return {}
    products = {}
    for ws in wb.worksheets:
        title = str(ws.title).strip()
        if not title.isdigit():
            continue  # 只处理以 VID 命名的分表（LIST / Tom提供 / Nick提供 等跳过）
        rows = list(ws.iter_rows(values_only=True))
        hdr = -1
        for i, r in enumerate(rows):
            if r and any(str(c).strip() == "OE Part No" for c in r if c is not None):
                hdr = i
                break
        if hdr < 0:
            continue
        items = []
        for r in rows[hdr + 1:]:
            cn = cell(r, 2)
            if not cn:
                continue
            oe = cell(r, 6)
            if oe.endswith(".0"):
                oe = oe[:-2]
            items.append({
                "cn": cn, "en": cell(r, 3), "brand": cell(r, 4), "cdap": cell(r, 5),
                "oe": oe, "notes": cell(r, 7), "supplier": cell(r, 8),
                "cost": cell(r, 9), "moq": cell(r, 10), "grp": cell(r, 1),
            })
        if items:
            products[title] = items
    wb.close()
    print(f"[产品] 解析到 {len(products)} 个车型的零件表")
    return products


def main():
    xlsx = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(DEFAULT_XLSX)
    if not xlsx.exists():
        print(f"找不到车型库文件: {xlsx}")
        sys.exit(1)
    records, skipped = load_records(xlsx)
    # 校正错填前缀 → 同 VDS 合并（每个 VIN 前缀对应唯一一款车）
    records = apply_corrections(records)
    if MERGE_SAME_VDS:
        records = merge_by_prefix(records)
    brands = sorted({r["brand"] for r in records})
    prefixes = sorted({r["prefix"] for r in records})
    wmis = sorted({p[:3] for p in prefixes if len(p) >= 3})

    meta = {
        "source": xlsx.name,
        "brands": brands,
        "nRecords": len(records),
        "nPrefixes": len(prefixes),
        "nWmi": len(wmis),
        "nSkipped": len(skipped),
    }
    products = load_products(PRODUCT_XLSX)
    meta["nProducts"] = sum(len(v) for v in products.values())
    meta["nVehiclesWithProducts"] = len(products)
    # 把期初数据落入 SQLite（仅当库空时执行；web 端修改受保护，不会被覆盖）。
    # 失败不阻断 HTML 生成（页面仍可启动，只是读不到数据）。
    try:
        dbstore.seed_from_excel(str(xlsx), PRODUCT_XLSX)
    except Exception as e:
        print(f"[警告] 期初数据写入 SQLite 失败：{e}")

    tpl = TEMPLATE.read_text(encoding="utf-8")
    # 数据不再内嵌：页面启动时从 /api/data 拉取（保证 web 修改落库后全员可见）
    html = tpl.replace("__DATA__", "null") \
              .replace("__META__", "null") \
              .replace("__PRODUCTS__", "{}")
    OUTPUT.write_text(html, encoding="utf-8")
    INDEX.write_text(html, encoding="utf-8")  # 同步部署文件
    print(f"已生成: {OUTPUT}")
    print(f"已同步部署页: {INDEX}")
    print(f"记录 {len(records)} 条 | 品牌 {len(brands)} 个 | 前缀 {len(prefixes)} 个 | WMI {len(wmis)} 个 | 无前缀跳过 {len(skipped)} 条")
    if products:
        print(f"产品：{len(products)} 个车型含产品，共 {meta['nProducts']} 条产品记录（来源：{PRODUCT_XLSX}）")
    else:
        print(f"产品：未注入（找不到 {PRODUCT_XLSX} 或无可解析分表）")
    if skipped:
        print("无 VIN 前缀、无法被 VIN 匹配的记录：")
        for s in skipped[:20]:
            print(f"  - {s['brand']} {s['model']} {s['sub'] or ''} ({s['year']})")


if __name__ == "__main__":
    main()
