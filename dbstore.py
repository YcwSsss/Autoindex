# -*- coding: utf-8 -*-
"""
SQLite 数据层：车型库(vehicles) + 产品(products) + VDS 真值表(vds_map)。
期初数据来自 Excel（车型/产品）与 vds_seed.json（VDS 真值表），统一落入 SQLite；
web 端的增删改直接写回 SQLite，所有同事共享同一份真源（不再依赖本机 localStorage）。

使用：
  - seed_from_excel()        把 Excel + VDS 种子落入 DB（仅当库空或 force 时，避免覆盖 web 修改）
  - get_all_data()           一次性返回前端启动所需的全部数据
  - upsert_product / delete_product / add_vehicle / delete_vehicles_web   web 写回
"""
import json
import sqlite3
from pathlib import Path

HERE = Path(__file__).parent
DB_PATH = HERE / "vin_data.db"

DEFAULT_VEHICLE_XLSX = r"D:\中国车\Oscar中国车主表.xlsx"
DEFAULT_PRODUCT_XLSX = r"D:\中国车\中国车产品数据表.xlsx"
VDS_SEED = HERE / "vds_seed.json"

VEHICLE_COLS = ["vid", "prefix", "brand", "model", "sub", "year", "ys", "ye",
                "cap", "kw", "engine", "chassis", "series", "body", "fuel", "drive", "origin",
                "vin17", "region", "transmission"]  # vin17: 核实过的完整 17 位实际 VIN 码（用于查询时精确命中）

PRODUCT_COLS = ["vid", "grp", "cn", "en", "brand", "cdap", "oe", "notes", "supplier", "cost", "moq", "loc", "interchange"]


def get_db():
    conn = sqlite3.connect(str(DB_PATH), timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA busy_timeout=30000")
    return conn


def init_db():
    conn = get_db()
    c = conn.cursor()
    c.executescript("""
    CREATE TABLE IF NOT EXISTS vehicles(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      vid TEXT, prefix TEXT, brand TEXT, model TEXT, sub TEXT,
      year TEXT, ys INTEGER, ye INTEGER, cap TEXT, kw TEXT,
      engine TEXT, chassis TEXT, series TEXT, body TEXT, fuel TEXT,
      drive TEXT, origin TEXT, vin17 TEXT,
      region TEXT DEFAULT '澳新地区'
    );
    CREATE TABLE IF NOT EXISTS products(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      vid TEXT, grp TEXT, cn TEXT, en TEXT, brand TEXT,
      cdap TEXT, oe TEXT, notes TEXT, supplier TEXT, cost TEXT, moq TEXT, loc TEXT,
      interchange TEXT DEFAULT ''
    );
    CREATE TABLE IF NOT EXISTS product_vehicles(
      product_id INTEGER,
      vid TEXT,
      PRIMARY KEY(product_id, vid)
    );
    CREATE TABLE IF NOT EXISTS vds_map(
      vds TEXT PRIMARY KEY,
      models TEXT,   -- JSON：字符串(唯一车型) 或 数组(碰撞多车型)
      src TEXT       -- 'lib' / 'web'
    );
    CREATE TABLE IF NOT EXISTS categories(
      code TEXT PRIMARY KEY,       -- 品类代码（如 A, A1, A301）
      cn TEXT DEFAULT '',          -- 品类中文名
      en TEXT DEFAULT '',          -- 品类英文名
      level INTEGER NOT NULL DEFAULT 1,  -- 1/2/3 级
      parent_code TEXT DEFAULT ''  -- 上级品类代码（一级为空）
    );
    CREATE TABLE IF NOT EXISTS product_catalog(
      grp TEXT NOT NULL,       -- 品类代码（= 产品表 grp）
      cn  TEXT NOT NULL,       -- 产品中文名
      en  TEXT DEFAULT '',     -- 产品英文名
      PRIMARY KEY(grp, cn)
    );
    CREATE TABLE IF NOT EXISTS users(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      username TEXT UNIQUE NOT NULL,   -- 登录用户名（也可作为登录账号）
      email TEXT UNIQUE,               -- 邮箱（与用户名都可作为登录账号）
      password TEXT NOT NULL,          -- 存 salt:sha256
      role TEXT NOT NULL DEFAULT 'user',  -- 'admin' | 'pm' | 'user'
      active INTEGER NOT NULL DEFAULT 1,  -- 1=启用 0=停用
      cats TEXT DEFAULT '',            -- 产品经理管理品类（逗号分隔，如 A,B,D,E,F）
      created_at TEXT
    );
    CREATE TABLE IF NOT EXISTS feedbacks(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      user_id INTEGER NOT NULL,
      username TEXT NOT NULL,
      content TEXT NOT NULL,
      created_at TEXT NOT NULL,
      reply TEXT DEFAULT '',
      replied_by TEXT DEFAULT '',
      replied_at TEXT DEFAULT ''
    );
    CREATE INDEX IF NOT EXISTS idx_prod_vid ON products(vid);
    CREATE INDEX IF NOT EXISTS idx_veh_vid ON vehicles(vid);
    CREATE INDEX IF NOT EXISTS idx_prod_grp ON products(grp);
    """)

    # 操作日志表
    c.execute("""CREATE TABLE IF NOT EXISTS logs(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      username TEXT,
      action TEXT,
      detail TEXT,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )""")
    # 旧库迁移：补 vin17 列（新库 CREATE TABLE 已含，重复 ALTER 会报 duplicate column，忽略即可）
    try:
        c.execute("ALTER TABLE vehicles ADD COLUMN vin17 TEXT")
    except sqlite3.OperationalError:
        pass
    try:
        c.execute("CREATE INDEX IF NOT EXISTS idx_veh_vin17 ON vehicles(vin17)")
    except sqlite3.OperationalError:
        pass
    # 旧库迁移：product_vehicles 补 remark 列（产品针对车型的备注）
    try:
        cols = [r[1] for r in c.execute("PRAGMA table_info(product_vehicles)").fetchall()]
        if 'remark' not in cols:
            c.execute("ALTER TABLE product_vehicles ADD COLUMN remark TEXT DEFAULT ''")
    except sqlite3.OperationalError:
        pass
    # 旧库迁移：products 补 loc 列（产品位置/安装位，如 左前/右后）
    try:
        cols = [r[1] for r in c.execute("PRAGMA table_info(products)").fetchall()]
        if 'loc' not in cols:
            c.execute("ALTER TABLE products ADD COLUMN loc TEXT DEFAULT ''")
    except sqlite3.OperationalError:
        pass
    _sync_catalog(conn)   # 产品目录从 products 表初始化（仅新增，不删）
    conn.commit()
    conn.close()


def _records_from_excel(xlsx):
    from build_vin_tool import load_records, apply_corrections, merge_by_prefix, MERGE_SAME_VDS
    records, skipped = load_records(xlsx)
    records = apply_corrections(records)
    if MERGE_SAME_VDS:
        records = merge_by_prefix(records)
    return records, skipped


def _products_from_excel(xlsx):
    from build_vin_tool import load_products
    return load_products(xlsx)


def seed_from_excel(vehicle_xlsx=DEFAULT_VEHICLE_XLSX, product_xlsx=DEFAULT_PRODUCT_XLSX,
                    vds_json=VDS_SEED, force=False):
    """把期初数据落入 SQLite。仅当 vehicles 为空或 force=True 时执行（保护 web 端修改）。"""
    init_db()
    conn = get_db()
    c = conn.cursor()
    n = c.execute("SELECT COUNT(*) AS n FROM vehicles").fetchone()["n"]
    if not force and n > 0:
        print(f"[seed] 车型库已有 {n} 条，跳过种子（如需重置为期初，用 force=True）")
        conn.close()
        return {"vehicles": n, "products": 0, "vds": 0, "skipped": 0, "seeded": False}

    if force:
        c.execute("DELETE FROM vehicles")
        c.execute("DELETE FROM products")
        c.execute("DELETE FROM vds_map")

    recs, skipped = _records_from_excel(vehicle_xlsx)
    for r in recs:
        c.execute("INSERT INTO vehicles(" + ",".join(VEHICLE_COLS) + ") VALUES (" +
                  ",".join("?" * len(VEHICLE_COLS)) + ")",
                  [r.get(col) for col in VEHICLE_COLS])
    veh_n = len(recs)

    p = _products_from_excel(product_xlsx)
    pn = 0
    for vid, items in p.items():
        for it in items:
            c.execute("INSERT INTO products(" + ",".join(PRODUCT_COLS) + ") VALUES (" +
                      ",".join("?" * len(PRODUCT_COLS)) + ")",
                      [vid, it.get("grp", ""), it.get("cn", ""), it.get("en", ""),
                       it.get("brand", ""), it.get("cdap", ""), it.get("oe", ""),
                       it.get("notes", ""), it.get("supplier", ""), it.get("cost", ""),
                       it.get("moq", "")])
            pn += 1

    vds_n = 0
    if vds_json and Path(vds_json).exists():
        seed = json.loads(Path(vds_json).read_text(encoding="utf-8"))
        vmap = seed.get("VDS_MAP", {})
        vsrc = seed.get("VDS_SRC", {})
        for k, val in vmap.items():
            c.execute("INSERT OR REPLACE INTO vds_map(vds, models, src) VALUES(?,?,?)",
                      [k, json.dumps(val, ensure_ascii=False), vsrc.get(k, "web")])
        vds_n = len(vmap)

    conn.commit()
    conn.close()
    print(f"[seed] 完成：车型 {veh_n} 条（skipped {len(skipped)}），产品 {pn} 条，VDS {vds_n} 条")
    return {"vehicles": veh_n, "products": pn, "vds": vds_n, "skipped": len(skipped), "seeded": True}


def get_all_products():
    """返回产品表：{ vid(str): [ {id,grp,cn,en,brand,cdap,oe,notes,supplier,cost,moq,loc}, ... ] }
    供 /api/products 懒加载使用。包含 products.vid (owner) 和 product_vehicles (额外关联)。"""
    conn = get_db()
    c = conn.cursor()
    prows = c.execute("SELECT id," + ",".join(PRODUCT_COLS) + " FROM products").fetchall()
    products = {}
    # 先按 owner vid 分组
    for r in prows:
        d = dict(r)
        vid = str(d["vid"]).strip()
        pid = d["id"]
        if vid:
            products.setdefault(vid, []).append(
                {k: d[k] for k in ["id", "grp", "cn", "en", "brand", "cdap", "oe", "notes", "supplier", "cost", "moq", "loc", "interchange"]})
    # 再把 product_vehicles 中的额外关联也加进去
    pvrows = c.execute("SELECT product_id, vid FROM product_vehicles").fetchall()
    pid_map = {r["id"]: dict(r) for r in prows}
    for r in pvrows:
        pid = r["product_id"]
        vid = str(r["vid"]).strip()
        if pid in pid_map and vid:
            d = pid_map[pid]
            owner_vid = str(d["vid"]).strip()
            # 只在不是 owner 的情况下才添加（避免重复）
            if vid != owner_vid:
                products.setdefault(vid, []).append(
                    {k: d[k] for k in ["id", "grp", "cn", "en", "brand", "cdap", "oe", "notes", "supplier", "cost", "moq", "loc", "interchange"]})
    conn.close()
    return products


def get_product_vehicles(pid):
    """返回某产品关联的所有车型：owner(products.vid) + 额外关联(product_vehicles)。"""
    conn = get_db()
    c = conn.cursor()
    p = c.execute("SELECT vid FROM products WHERE id=?", [pid]).fetchone()
    owner = (p["vid"] or "").strip() if p else ""
    rows = c.execute("SELECT vid FROM product_vehicles WHERE product_id=?", [pid]).fetchall()
    conn.close()
    res = []
    if owner:
        res.append({"vid": owner, "owner": True})
    seen = {owner}
    for r in rows:
        v = (r["vid"] or "").strip()
        if v and v not in seen:
            seen.add(v)
            res.append({"vid": v, "owner": False})
    return res


def get_product_vehicles_by_code(oe, cdap):
    """按 OE+CDAP 聚合所有关联车型（products 表中的 owner vid + product_vehicles 额外关联），去重返回。含 remark。"""
    oe = (oe or '').strip()
    cdap = (cdap or '').strip()
    if not oe and not cdap:
        return []
    conn = get_db()
    c = conn.cursor()
    where = []
    params = []
    if oe:
        where.append("oe=?")
        params.append(oe)
    if cdap:
        where.append("cdap=?")
        params.append(cdap)
    prows = c.execute("SELECT id, vid FROM products WHERE " + " AND ".join(where), params).fetchall()
    res = []
    seen = set()
    pids = []
    for r in prows:
        vid = (r["vid"] or "").strip()
        pid = r["id"]
        if vid and vid not in seen:
            seen.add(vid)
            res.append({"vid": vid, "owner": True, "pid": pid, "remark": ""})
        pids.append(pid)
    # 查 product_vehicles 中的额外关联及备注
    if pids:
        placeholders = ",".join(["?"] * len(pids))
        rows = c.execute(
            f"SELECT pv.vid, pv.product_id, COALESCE(pv.remark,'') AS remark "
            f"FROM product_vehicles pv WHERE pv.product_id IN ({placeholders})", pids).fetchall()
        for r in rows:
            vid = (r["vid"] or "").strip()
            if vid and vid not in seen:
                seen.add(vid)
                res.append({"vid": vid, "owner": False, "pid": r["product_id"], "remark": r["remark"]})
            else:
                # 已存在（可能是 owner 行），补上 remark
                for item in res:
                    if item["vid"] == vid and not item.get("remark"):
                        item["remark"] = r["remark"]
                        break
    conn.close()
    return res


def add_product_vehicle(pid, vid):
    vid = (vid or "").strip()
    if not vid:
        return False
    conn = get_db()
    c = conn.cursor()
    c.execute("INSERT OR IGNORE INTO product_vehicles(product_id, vid) VALUES(?,?)", [pid, vid])
    # 若产品尚无 owner 车型，则把新增的车型设为 owner（保持 VIN 查询可用）
    p = c.execute("SELECT vid FROM products WHERE id=?", [pid]).fetchone()
    if p and not (p["vid"] or "").strip():
        c.execute("UPDATE products SET vid=? WHERE id=?", [vid, pid])
    conn.commit()
    conn.close()
    return True


def delete_product_vehicle(pid, vid):
    conn = get_db()
    c = conn.cursor()
    # 1. 删除 product_vehicles 表中的关联记录（非主关联）
    c.execute("DELETE FROM product_vehicles WHERE product_id=? AND vid=?", [pid, vid])
    # 2. 若该 vid 是产品主关联（products.vid 字段直接存储），则清空主关联
    c.execute("UPDATE products SET vid='' WHERE id=? AND vid=?", [pid, vid])
    conn.commit()
    conn.close()
    return True


def get_vehicle_product_remarks(vid):
    """返回某车型下所有产品-车型备注：[{product_id, remark}]（仅 product_vehicles 表中有记录的）。"""
    vid = (vid or "").strip()
    if not vid:
        return []
    conn = get_db()
    c = conn.cursor()
    rows = c.execute("SELECT product_id, remark FROM product_vehicles WHERE vid=?", [vid]).fetchall()
    conn.close()
    return [{"product_id": r["product_id"], "remark": r["remark"] or ""} for r in rows]


def set_product_vehicle_remark(pid, vid, remark):
    """设置某产品针对某车型的备注（按 product_id+vid 唯一）。会自动确保 product_vehicles 行存在。"""
    vid = (vid or "").strip()
    remark = (remark if remark is not None else "").strip()
    if not pid or not vid:
        raise ValueError("缺少 product_id 或 vid")
    conn = get_db()
    c = conn.cursor()
    c.execute("INSERT OR IGNORE INTO product_vehicles(product_id, vid) VALUES(?,?)", [pid, vid])
    c.execute("UPDATE product_vehicles SET remark=? WHERE product_id=? AND vid=?", [remark, pid, vid])
    # 若产品尚无 owner 车型，则把备注车型设为 owner（保持 VIN 查询可用）
    p = c.execute("SELECT vid FROM products WHERE id=?", [pid]).fetchone()
    if p and not (p["vid"] or "").strip():
        c.execute("UPDATE products SET vid=? WHERE id=?", [vid, pid])
    conn.commit()
    conn.close()
    return True


def get_all_data(include_products=True):
    conn = get_db()
    c = conn.cursor()
    vrows = c.execute("SELECT id," + ",".join(VEHICLE_COLS) + " FROM vehicles").fetchall()
    rows = [dict(r) for r in vrows]

    products = {}
    if include_products:
        prows = c.execute("SELECT id," + ",".join(PRODUCT_COLS) + " FROM products").fetchall()
        for r in prows:
            d = dict(r)
            vid = str(d["vid"])
            products.setdefault(vid, []).append(
                {k: d[k] for k in ["id", "grp", "cn", "en", "brand", "cdap", "oe", "notes", "supplier", "cost", "moq", "interchange"]})
    else:
        n_products = c.execute("SELECT COUNT(*) FROM products").fetchone()[0]
        n_vids = c.execute("SELECT COUNT(DISTINCT vid) FROM products").fetchone()[0]

    vrows2 = c.execute("SELECT vds, models, src FROM vds_map").fetchall()
    VDS_MAP, VDS_SRC = {}, {}
    for r in vrows2:
        try:
            val = json.loads(r["models"])
        except Exception:
            val = r["models"]
        VDS_MAP[r["vds"]] = val
        VDS_SRC[r["vds"]] = r["src"] or "web"

    brands = sorted({r["brand"] for r in rows if r["brand"]})
    _region_set = set()
    for r in rows:
        for part in (r.get("region") or '').split(','):
            part = part.strip()
            if part:
                _region_set.add(part)
    regions = sorted(_region_set)
    prefixes = sorted({r["prefix"] for r in rows if r["prefix"]})
    wmis = sorted({p[:3] for p in prefixes if p and len(p) >= 3})
    meta = {
        "source": "SQLite (vin_data.db)",
        "brands": brands,
        "regions": regions,
        "nRecords": len(rows),
        "nPrefixes": len(prefixes),
        "nWmi": len(wmis),
        "nSkipped": 0,
        "nProducts": (sum(len(v) for v in products.values()) if include_products else n_products),
        "nVehiclesWithProducts": (len(products) if include_products else n_vids),
    }
    conn.close()
    # 首屏不返回 DATA 列表（避免大数据量；车型走 /api/vehicles 分页拉取）
    return {"DATA": rows, "META": meta, "PRODUCTS": products, "VDS_MAP": VDS_MAP, "VDS_SRC": VDS_SRC}


def get_vehicles_page(limit=200, offset=0, q="", brand="", region=""):
    """分页拉取车型；支持品牌、地区过滤与模糊搜索（按 brand/model/sub/prefix/engine/chassis/series/vid/vin17/body/fuel）。
    limit=-1 表示不限制，返回全部匹配数据。
    返回 (rows, total)。rows 已按 id 升序，便于追加分页时保持稳定。"""
    if limit == -1:
        limit = None  # 不限制
    else:
        limit = max(1, min(int(limit or 200), 20000))
    offset = max(0, int(offset or 0))
    conn = get_db()
    c = conn.cursor()
    where = []
    params = []
    if q:
        like = f"%{q}%"
        where.append("(" +
                     "brand LIKE ? OR model LIKE ? OR sub LIKE ? OR prefix LIKE ? OR " +
                     "engine LIKE ? OR chassis LIKE ? OR series LIKE ? OR vid LIKE ? OR " +
                     "vin17 LIKE ? OR body LIKE ? OR fuel LIKE ? OR drive LIKE ? OR kw LIKE ? OR " +
                     "transmission LIKE ?" +
                     ")")
        params.extend([like] * 14)
    if brand:
        where.append("brand = ?")
        params.append(brand)
    if region:
        where.append("region LIKE ?")
        params.append(f"%{region}%")
    where_sql = (" WHERE " + " AND ".join(where)) if where else ""
    total = c.execute(f"SELECT COUNT(*) AS n FROM vehicles{where_sql}", params).fetchone()["n"]
    if limit is None:
        # 不限制，返回全部
        rows = c.execute(
            f"SELECT id," + ",".join(VEHICLE_COLS) +
            f" FROM vehicles{where_sql} ORDER BY id",
            params
        ).fetchall()
    else:
        rows = c.execute(
            f"SELECT id," + ",".join(VEHICLE_COLS) +
            f" FROM vehicles{where_sql} ORDER BY id LIMIT ? OFFSET ?",
            params + [limit, offset]
        ).fetchall()
    conn.close()
    return [dict(r) for r in rows], int(total)


def list_regions():
    """返回所有不重复的地区列表（拆分逗号分隔的多值，按名称排序）。"""
    conn = get_db()
    c = conn.cursor()
    rows = c.execute("SELECT region FROM vehicles WHERE region IS NOT NULL AND region != ''").fetchall()
    conn.close()
    seen = set()
    for r in rows:
        for part in (r[0] or '').split(','):
            part = part.strip()
            if part:
                seen.add(part)
    return sorted(seen)


# ---------- 级联筛选（字段互约束） ----------

# 所有可筛选字段（数据库列名）
_FILTER_FIELDS = [
    "region", "brand", "model", "sub", "year", "cap", "engine",
    "chassis", "series", "kw", "body", "fuel", "drive", "transmission", "origin"
]


def get_filter_options(filters: dict) -> dict:
    """根据已选条件，返回每个筛选字段的可选值列表。
    filters: {field_name: selected_value, ...}，如 {"region": "澳新地区", "brand": "MG"}
    返回: {"region": [...], "brand": [...], "model": [...], ...}，每个字段排除自身条件后查询。
    """
    conn = get_db()
    c = conn.cursor()
    result = {}
    for field in _FILTER_FIELDS:
        # 构建 WHERE：用 filters 中除当前字段外的所有条件
        where_parts = []
        params = []
        for f, v in filters.items():
            if f == field or not v:
                continue
            if f in _FILTER_FIELDS:
                where_parts.append(f"[{f}] LIKE ?")
                params.append(f"%{v}%")
        # 加 NOT NULL / NOT EMPTY 条件
        where_parts.append(f"[{field}] IS NOT NULL AND [{field}] != ''")
        where_sql = " WHERE " + " AND ".join(where_parts)
        rows = c.execute(
            f"SELECT DISTINCT [{field}] FROM vehicles{where_sql} ORDER BY [{field}]",
            params
        ).fetchall()
        result[field] = [r[0] for r in rows]
    conn.close()
    return result


# ---------- 品类管理（产品分组） ----------

def _sync_categories(conn):
    """保证产品表里出现过的每个 grp 代码在 categories 表都有对应行（名称留空，默认一级）。"""
    c = conn.cursor()
    for r in c.execute("SELECT DISTINCT TRIM(grp) AS g FROM products WHERE TRIM(grp)<>''").fetchall():
        c.execute("INSERT OR IGNORE INTO categories(code, cn, en, level, parent_code) VALUES(?,'','',1,'')", [r["g"]])
    conn.commit()


def _sync_catalog(conn):
    """保证 products 表里出现过的每个 (grp,cn) 组合在 product_catalog 都有对应行。"""
    c = conn.cursor()
    rows = c.execute(
        "SELECT TRIM(grp) AS g, TRIM(cn) AS n, TRIM(COALESCE(en,'')) AS e "
        "FROM products WHERE TRIM(grp)<>'' AND TRIM(cn)<>''").fetchall()
    for r in rows:
        c.execute("INSERT OR IGNORE INTO product_catalog(grp, cn, en) VALUES(?,?,?)",
                  [r["g"], r["n"], r["e"]])
        if r["e"]:
            c.execute("UPDATE product_catalog SET en=? WHERE grp=? AND cn=? AND (en IS NULL OR en='')",
                      [r["e"], r["g"], r["n"]])
    conn.commit()


def add_catalog_product(grp, cn, en=""):
    """在产品目录中新增/更新一个产品（某品类下）。"""
    grp = str(grp or "").strip().upper()
    cn = str(cn or "").strip()
    en = str(en or "").strip().upper()
    if not grp:
        raise ValueError("品类代码不能为空")
    if not cn:
        raise ValueError("产品中文名不能为空")
    conn = get_db()
    c = conn.cursor()
    c.execute("INSERT OR IGNORE INTO categories(code) VALUES(?)", [grp])
    c.execute("INSERT OR REPLACE INTO product_catalog(grp, cn, en) VALUES(?,?,?)", [grp, cn, en])
    conn.commit()
    conn.close()
    return True


def update_catalog_product(grp, old_cn, cn="", en=""):
    """更新产品目录中的产品名（中英文），同步更新 products 表中所有匹配记录。"""
    grp = str(grp or "").strip().upper()
    old_cn = str(old_cn or "").strip()
    cn = str(cn or "").strip()
    en = str(en or "").strip().upper()
    if not grp or not old_cn:
        raise ValueError("品类代码和原产品名不能为空")
    if not cn:
        raise ValueError("新产品中文名不能为空")
    conn = get_db()
    c = conn.cursor()
    # 更新 product_catalog
    c.execute("UPDATE product_catalog SET cn=?, en=? WHERE grp=? AND cn=?", [cn, en, grp, old_cn])
    cat_updated = c.rowcount
    # 同步更新 products 表中所有匹配的记录
    c.execute("UPDATE products SET cn=?, en=? WHERE TRIM(grp)=? AND TRIM(cn)=?", [cn, en, grp, old_cn])
    prod_updated = c.rowcount
    conn.commit()
    conn.close()
    return {"catalog": cat_updated, "products": prod_updated}


def delete_catalog_product(grp, cn):
    """删除产品目录中的产品。若该产品名下已有关联的产品号码（products 表），则拒绝删除。"""
    grp = str(grp or "").strip().upper()
    cn = str(cn or "").strip()
    if not grp or not cn:
        raise ValueError("品类代码和产品名不能为空")
    conn = get_db()
    c = conn.cursor()
    # 检查 products 表中是否有该 (grp, cn) 下的产品号码
    count = c.execute("SELECT COUNT(*) FROM products WHERE TRIM(grp)=? AND TRIM(cn)=?", [grp, cn]).fetchone()[0]
    if count > 0:
        conn.close()
        raise ValueError(f"该产品下已有 {count} 个产品号码，无法删除")
    c.execute("DELETE FROM product_catalog WHERE grp=? AND cn=?", [grp, cn])
    deleted = c.rowcount
    conn.commit()
    conn.close()
    if deleted == 0:
        raise ValueError("产品不存在或已被删除")
    return {"deleted": deleted}


def export_catalog_excel():
    """导出品类+产品目录为 Excel（bytes），含品类代码/中英文名 + 每个产品中英文名。"""
    from io import BytesIO
    try:
        import openpyxl
    except ImportError:
        raise RuntimeError("需要安装 openpyxl: pip install openpyxl")
    cats, _ = list_categories()  # cats 含 products 列表
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "品类产品目录"
    # 表头
    headers = ["品类代码", "品类中文名", "品类英文名", "产品中文名", "产品英文名"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = openpyxl.styles.Font(bold=True)
        cell.fill = openpyxl.styles.PatternFill(start_color="EEEDFE", end_color="EEEDFE", fill_type="solid")
    row = 2
    for cat in cats:
        prods = cat.get("products") or []
        if not prods:
            ws.cell(row=row, column=1, value=cat["code"])
            ws.cell(row=row, column=2, value=cat.get("cn", ""))
            ws.cell(row=row, column=3, value=cat.get("en", ""))
            row += 1
        else:
            first = True
            for p in prods:
                if first:
                    ws.cell(row=row, column=1, value=cat["code"])
                    ws.cell(row=row, column=2, value=cat.get("cn", ""))
                    ws.cell(row=row, column=3, value=cat.get("en", ""))
                    first = False
                ws.cell(row=row, column=4, value=p["cn"])
                ws.cell(row=row, column=5, value=p.get("en", ""))
                row += 1
    # 调整列宽
    ws.column_dimensions['A'].width = 10
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 22
    ws.column_dimensions['D'].width = 22
    ws.column_dimensions['E'].width = 30
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def list_categories():
    """返回品类树（含每个品类的产品记录数、去重产品清单）。

    返回扁平列表，每项含 code/cn/en/level/parent_code，前端自行构建树。
    产品清单来自产品目录 product_catalog；count 为该品类产品记录数，
    products[].n 为使用该产品名（grp+cn）的车型数。
    """
    conn = get_db()
    _sync_categories(conn)
    _sync_catalog(conn)
    c = conn.cursor()
    cats = [dict(r) for r in c.execute(
        "SELECT code, cn, en, level, COALESCE(parent_code,'') AS parent_code "
        "FROM categories ORDER BY code"
    ).fetchall()]
    for cat in cats:
        g = cat["code"]
        cat["count"] = c.execute("SELECT COUNT(*) AS n FROM products WHERE TRIM(grp)=?", [g]).fetchone()["n"]
        cat["products"] = [
            {"cn": r["cn"], "en": r["en"],
             "n": c.execute("SELECT COUNT(*) AS n FROM products WHERE TRIM(grp)=? AND TRIM(cn)=?",
                            [g, r["cn"]]).fetchone()["n"],
             "valid": c.execute("SELECT COUNT(*) AS n FROM products WHERE TRIM(grp)=? AND TRIM(cn)=? "
                                "AND (TRIM(COALESCE(cdap,''))<>'' OR TRIM(COALESCE(oe,''))<>'')",
                                [g, r["cn"]]).fetchone()["n"]}
            for r in c.execute("SELECT cn, en FROM product_catalog WHERE grp=? ORDER BY en", [g]).fetchall()
        ]
    # 未归类的产品记录（grp 为空）单独提示
    cat_un = c.execute("SELECT COUNT(*) AS n FROM products WHERE TRIM(COALESCE(grp,''))=''").fetchone()["n"]
    conn.close()
    return cats, int(cat_un)


def add_category(code, cn="", en="", level=1, parent_code=""):
    import re
    code = str(code or "").strip().upper()
    level = int(level or 1)
    parent_code = str(parent_code or "").strip().upper()
    if not code:
        raise ValueError("品类代码不能为空")
    if not re.fullmatch(r"[A-Z0-9-]{1,10}", code):
        raise ValueError("品类代码仅限 1-10 位字母 / 数字 / 连字符")
    if level not in (1, 2, 3):
        raise ValueError("品类级别必须为 1、2 或 3")
    if level > 1 and not parent_code:
        raise ValueError("二/三级品类必须指定上级品类")
    conn = get_db()
    c = conn.cursor()
    if c.execute("SELECT 1 FROM categories WHERE code=?", [code]).fetchone():
        conn.close()
        raise ValueError(f"品类代码 {code} 已存在")
    if level > 1 and not c.execute("SELECT 1 FROM categories WHERE code=?", [parent_code]).fetchone():
        conn.close()
        raise ValueError(f"上级品类 {parent_code} 不存在")
    c.execute("INSERT INTO categories(code, cn, en, level, parent_code) VALUES(?,?,?,?,?)",
              [code, cn, en, level, parent_code if level > 1 else ""])
    conn.commit()
    conn.close()
    return code


def update_category(code, cn="", en="", parent_code=None):
    code = str(code or "").strip().upper()
    conn = get_db()
    c = conn.cursor()
    if parent_code is not None:
        parent_code = str(parent_code or "").strip().upper()
        n = c.execute("UPDATE categories SET cn=?, en=?, parent_code=? WHERE code=?",
                       [cn, en, parent_code, code]).rowcount
    else:
        n = c.execute("UPDATE categories SET cn=?, en=? WHERE code=?", [cn, en, code]).rowcount
    conn.commit()
    conn.close()
    if not n:
        raise ValueError("品类不存在")
    return True


def delete_category(code):
    """删除品类（仅当该品类下没有子品类和产品记录时允许）。"""
    code = str(code or "").strip().upper()
    conn = get_db()
    c = conn.cursor()
    # 检查子品类
    children = c.execute("SELECT COUNT(*) AS n FROM categories WHERE parent_code=?", [code]).fetchone()["n"]
    if children:
        conn.close()
        raise ValueError(f"品类 {code} 下仍有 {children} 个子品类，请先删除子品类")
    # 检查产品
    n = c.execute("SELECT COUNT(*) AS n FROM products WHERE TRIM(grp)=?", [code]).fetchone()["n"]
    if n:
        conn.close()
        raise ValueError(f"品类 {code} 下仍有 {n} 条产品记录，请先清空或转移这些产品")
    c.execute("DELETE FROM categories WHERE code=?", [code])
    deleted = c.rowcount
    conn.commit()
    conn.close()
    if not deleted:
        raise ValueError("品类不存在")
    return deleted


def _default_product_template():
    """标准产品模板 = 产品目录（product_catalog）中的全部 (grp, cn, en)。

    新增到目录的产品会自动成为所有车型的默认产品。
    """
    conn = get_db()
    c = conn.cursor()
    rows = c.execute(
        "SELECT grp, cn, en FROM product_catalog WHERE TRIM(grp)<>'' AND TRIM(cn)<>'' ORDER BY grp, en").fetchall()
    conn.close()
    return [{"grp": r["grp"], "cn": r["cn"], "en": r["en"] or ""} for r in rows]


def get_all_catalog():
    """返回产品目录全量 (grp,cn,en)，供前端「添加产品」下拉（按品类产品名称搜索填充）。"""
    conn = get_db()
    c = conn.cursor()
    rows = c.execute(
        "SELECT grp, cn, en FROM product_catalog WHERE TRIM(grp)<>'' AND TRIM(cn)<>'' ORDER BY grp, en").fetchall()
    conn.close()
    return [{"grp": r["grp"], "cn": r["cn"], "en": (r["en"] or "")} for r in rows]


def seed_default_products_for_vid(vid):
    """给单个车型补齐默认产品模板（已有产品则跳过）。返回插入条数。"""
    vid = str(vid or "").strip()
    if not vid:
        return 0
    template = _default_product_template()
    if not template:
        return 0
    conn = get_db()
    c = conn.cursor()
    has = c.execute("SELECT COUNT(*) AS n FROM products WHERE vid=?", [vid]).fetchone()["n"]
    if has:
        conn.close()
        return 0
    for t in template:
        c.execute("INSERT INTO products(vid,grp,cn,en,brand,cdap,oe,notes,supplier,cost,moq) "
                  "VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                  [vid, t["grp"], t["cn"], t["en"], "", "", "", "", "", "", ""])
    conn.commit()
    conn.close()
    return len(template)


def seed_default_products():
    """给所有还没有产品记录的车型批量插入"标准产品模板"（OE/品牌等字段留空）。

    幂等：已有产品的 vid 跳过，重复调用不会产生重复记录。
    """
    template = _default_product_template()
    if not template:
        return {"ok": False, "error": "库中没有任何带产品的车型，无法推导模板", "inserted": 0}

    conn = get_db()
    c = conn.cursor()
    vids_with = {r["vid"] for r in
                 c.execute("SELECT DISTINCT vid FROM products WHERE TRIM(COALESCE(vid,''))<>''").fetchall()}
    all_vids = [r["vid"] for r in
                c.execute("SELECT DISTINCT vid FROM vehicles WHERE TRIM(COALESCE(vid,''))<>''").fetchall()]
    targets = [v for v in all_vids if v not in vids_with]

    inserted = 0
    for v in targets:
        for t in template:
            c.execute("INSERT INTO products(vid,grp,cn,en,brand,cdap,oe,notes,supplier,cost,moq) "
                      "VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                      [v, t["grp"], t["cn"], t["en"], "", "", "", "", "", "", ""])
            inserted += 1
    conn.commit()
    conn.close()
    return {"ok": True, "template": len(template), "vehicles": len(targets),
            "inserted": inserted, "skipped": len(vids_with)}


def reseed_default_products():
    """把**所有已有产品的车型**补齐到产品目录（product_catalog）全集。

    与 seed_default_products 不同：这里不跳过已有产品的车型，而是逐个对照
    catalog，把车型缺失的 (grp,cn) 补进来（按 vid+grp+cn 去重，不会重复插入）。
    车型已有的、catalog 之外的额外产品会保留，不会被删除。幂等、可重复调用。
    """
    template = _default_product_template()
    if not template:
        return {"ok": False, "error": "catalog 为空，无法补齐", "inserted": 0}

    conn = get_db()
    c = conn.cursor()
    vids = [r["vid"] for r in c.execute(
        "SELECT DISTINCT vid FROM products WHERE TRIM(COALESCE(vid,''))<>''").fetchall()]

    inserted = 0
    for vid in vids:
        have = {(r["grp"], r["cn"]) for r in
                c.execute("SELECT grp, cn FROM products WHERE vid=?", [vid]).fetchall()}
        for t in template:
            if (t["grp"], t["cn"]) in have:
                continue
            c.execute("INSERT INTO products(vid,grp,cn,en,brand,cdap,oe,notes,supplier,cost,moq) "
                      "VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                      [vid, t["grp"], t["cn"], t["en"], "", "", "", "", "", "", ""])
            inserted += 1
    conn.commit()
    conn.close()
    return {"ok": True, "template": len(template), "vehicles": len(vids), "inserted": inserted}


# ---------- web 写回 ----------

def upsert_product(vid, pid, grp="", cn="", en="", brand="", cdap="", oe="",
                   notes="", supplier="", cost="", moq="", loc="", interchange=""):
    conn = get_db()
    c = conn.cursor()
    pid = int(pid) if pid else 0
    if pid > 0:
        c.execute("UPDATE products SET vid=?,grp=?,cn=?,en=?,brand=?,cdap=?,oe=?,notes=?,supplier=?,cost=?,moq=?,loc=?,interchange=? WHERE id=?",
                  [vid, grp, cn, en, brand, cdap, oe, notes, supplier, cost, moq, loc, interchange, pid])
        new_id = pid
    else:
        cur = c.execute("INSERT INTO products(" + ",".join(PRODUCT_COLS) + ") VALUES (" +
                        ",".join("?" * len(PRODUCT_COLS)) + ")",
                        [vid, grp, cn, en, brand, cdap, oe, notes, supplier, cost, moq, loc, interchange])
        new_id = cur.lastrowid
    conn.commit()
    conn.close()
    return new_id


def delete_product(pid):
    conn = get_db()
    c = conn.cursor()
    c.execute("DELETE FROM products WHERE id=?", [int(pid)])
    conn.commit()
    conn.close()


def add_vehicle(rec):
    conn = get_db()
    c = conn.cursor()
    rec = dict(rec)
    # vin17 支持多条：规范化为逗号分隔的合法 17 位 VIN 列表
    if rec.get("vin17"):
        rec["vin17"] = _norm_vin17(rec["vin17"])
    # 地区默认"澳新地区"
    if not rec.get("region"):
        rec["region"] = "澳新地区"
    vals = [rec.get(col) for col in VEHICLE_COLS]
    cur = c.execute("INSERT INTO vehicles(" + ",".join(VEHICLE_COLS) + ") VALUES (" +
                    ",".join("?" * len(VEHICLE_COLS)) + ")", vals)
    new_id = cur.lastrowid
    conn.commit()
    conn.close()
    return new_id


def delete_vehicles_web():
    conn = get_db()
    c = conn.cursor()
    c.execute("DELETE FROM vehicles WHERE origin=?", ["web"])
    n = c.rowcount
    conn.commit()
    conn.close()
    return n


def delete_vehicle(vid):
    """按行 id 删除单条车型（仅删除 web 新增较为安全，但仍由调用方控制）。"""
    conn = get_db()
    c = conn.cursor()
    c.execute("DELETE FROM vehicles WHERE id=?", [int(vid)])
    n = c.rowcount
    conn.commit()
    conn.close()
    return n


def _split_vins(v):
    """把逗号/空白/分号分隔的字符串拆成 VIN 列表（去空、去重、保序、大写）。"""
    import re
    toks = [t for t in re.split(r"[,;\s、，；]+", str(v or "").strip().upper()) if t]
    seen, out = set(), []
    for t in toks:
        if t not in seen:
            seen.add(t)
            out.append(t)
    return out


def _norm_vin17(v):
    """规范化 vin17 字段：仅保留合法的 17 位 VIN，逗号连接；无合法项返回 None。"""
    import re
    vins = [t for t in _split_vins(v) if re.fullmatch(r"[A-HJ-NPR-Z0-9]{17}", t)]
    return ",".join(vins) if vins else None


def update_vehicle_vins(row_id, vins_raw):
    """登记/更新某条车型的实际 17 位 VIN 码（支持多条，逗号/空白/分号分隔）。
    传入整体列表（全量覆盖）；传空清除。每条须 17 位且不含 I/O/Q；
    同一 VIN 不允许登记到多条车型。返回 (ok, vins_list 或 error_message)。"""
    import re
    row_id = int(row_id)
    vins = _split_vins(vins_raw)
    bad = [t for t in vins if not re.fullmatch(r"[A-HJ-NPR-Z0-9]{17}", t)]
    if bad:
        return False, "以下 VIN 不合法（须 17 位且不含 I/O/Q）：" + "、".join(bad[:5]) + ("…" if len(bad) > 5 else "")
    if len(vins) > 50:
        return False, "单条车型最多登记 50 个实际 VIN"
    conn = get_db()
    c = conn.cursor()
    # 同一 VIN 只允许登记到一条车型（防止重复登记导致查询歧义）
    if vins:
        vset = set(vins)
        for r in c.execute("SELECT id, model, vin17 FROM vehicles "
                           "WHERE vin17 IS NOT NULL AND vin17<>'' AND id<>?", [row_id]).fetchall():
            other = _split_vins(r["vin17"])
            clash = vset.intersection(other)
            if clash:
                conn.close()
                return False, ("VIN " + "、".join(sorted(clash)) +
                               f" 已登记到其他车型（id={r['id']} {r['model']}）")
    c.execute("UPDATE vehicles SET vin17=? WHERE id=?", [",".join(vins) if vins else None, row_id])
    n = c.rowcount
    conn.commit()
    conn.close()
    if not n:
        return False, "车型记录不存在"
    return True, vins


# 兼容旧调用名
def update_vehicle_vin17(row_id, vin17):
    return update_vehicle_vins(row_id, vin17)


# ---------- 用户与权限（登录模块） ----------

import hashlib as _hashlib
import os as _os
from datetime import datetime as _dt

# 产品经理默认管理品类（与产品表 grp 字段一致：A/B/D/E/F）
PM_DEFAULT_CATS = "A,B,D,E,F"


def _now():
    return _dt.now().strftime("%Y-%m-%d %H:%M:%S")


def _norm_cats(cats):
    """把 逗号分隔字符串 / 列表 / 元组 规范化为大写、去重、保序的逗号串（如 'A,B,D,E,F'）。"""
    if isinstance(cats, str):
        items = [x.strip() for x in cats.split(",") if x.strip()]
    else:
        items = [str(x).strip() for x in (cats or []) if str(x).strip()]
    seen, out = set(), []
    for it in items:
        k = it.upper()
        if k not in seen:
            seen.add(k)
            out.append(k)
    return ",".join(out)


def _hash_pw(pw, salt=None):
    """返回 'salt:sha256'，salt 随机。pw 为空也允许（调用方应校验）。"""
    pw = pw or ""
    if salt is None:
        salt = _os.urandom(8).hex()
    h = _hashlib.sha256((salt + ":" + pw).encode("utf-8")).hexdigest()
    return salt + ":" + h


def _check_pw(pw, stored):
    if not stored or ":" not in stored:
        return False
    salt, h = stored.split(":", 1)
    return h == _hashlib.sha256((salt + ":" + pw).encode("utf-8")).hexdigest()


def init_users():
    """建表并种子 admin 账号（admin / admin123，管理员）。幂等。"""
    conn = get_db()
    c = conn.cursor()
    c.execute("""
    CREATE TABLE IF NOT EXISTS users(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      username TEXT UNIQUE NOT NULL,
      email TEXT UNIQUE,
      password TEXT NOT NULL,
      role TEXT NOT NULL DEFAULT 'user',
      active INTEGER NOT NULL DEFAULT 1,
      cats TEXT DEFAULT '',
      region_group TEXT DEFAULT '',
      created_at TEXT
    )""")
    if not c.execute("SELECT 1 FROM users WHERE username=?", ["admin"]).fetchone():
        c.execute(
            "INSERT INTO users(username,email,password,role,active,cats,created_at) "
            "VALUES(?,?,?,?,?,?,?)",
            ["admin", "admin@autoindex.local", _hash_pw("admin123"), "admin", 1, "", _now()])
    conn.commit()
    conn.close()


def get_user_by_login(login):
    """按用户名或邮箱匹配（用户名与邮箱均大小写不敏感）。返回 row dict 或 None。"""
    login = (login or "").strip()
    if not login:
        return None
    conn = get_db()
    c = conn.cursor()
    u = c.execute("SELECT * FROM users WHERE LOWER(username)=LOWER(?)", [login]).fetchone()
    if not u:
        u = c.execute("SELECT * FROM users WHERE LOWER(email)=LOWER(?)", [login]).fetchone()
    conn.close()
    return dict(u) if u else None


def verify_user(login, pw):
    """校验凭证且账号启用。返回用户 dict（不含 password）或 None。"""
    u = get_user_by_login(login)
    if not u:
        return None
    if not u.get("active"):
        return None
    if not _check_pw(pw, u.get("password", "")):
        return None
    u.pop("password", None)
    return u


def update_me(uid, email=None, password=None, current_password=None):
    """用户自行修改邮箱/密码。必须先验证 current_password。返回更新后的用户 dict（不含 password）。"""
    uid = int(uid)
    if current_password is None:
        raise ValueError("请输入当前密码")
    conn = get_db()
    c = conn.cursor()
    u = c.execute("SELECT * FROM users WHERE id=?", [uid]).fetchone()
    if not u:
        conn.close()
        raise ValueError("用户不存在")
    if not _check_pw(current_password, u["password"] or ""):
        conn.close()
        raise ValueError("当前密码错误")
    sets, params = [], []
    if email is not None:
        email = (email or "").strip().lower()
        if email and not _is_email(email):
            conn.close()
            raise ValueError("邮箱格式不合法")
        if email and c.execute("SELECT 1 FROM users WHERE email=? AND id<>?", [email, uid]).fetchone():
            conn.close()
            raise ValueError(f"邮箱 {email} 已被其他用户使用")
        sets.append("email=?")
        params.append(email or None)
    if password is not None:
        if not password:
            conn.close()
            raise ValueError("新密码不能为空")
        sets.append("password=?")
        params.append(_hash_pw(password))
    if sets:
        c.execute("UPDATE users SET " + ",".join(sets) + " WHERE id=?", params + [uid])
    conn.commit()
    u = c.execute(
        "SELECT id,username,email,role,active,cats,created_at FROM users WHERE id=?", [uid]).fetchone()
    conn.close()
    return dict(u)


def list_users():
    """返回全部用户（不含 password）。"""
    conn = get_db()
    c = conn.cursor()
    rows = c.execute(
        "SELECT id,username,email,role,active,cats,region_group,created_at FROM users ORDER BY id").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def add_user(username, email, password, role, cats=None, region_group=None):
    username = (username or "").strip()
    email = (email or "").strip().lower()
    role = (role or "user").strip()
    if not username:
        raise ValueError("用户名不能为空")
    if role not in ("admin", "pm", "user"):
        raise ValueError("权限必须是 admin / pm / user 之一")
    if not password:
        raise ValueError("初始密码不能为空")
    if email and not _is_email(email):
        raise ValueError("邮箱格式不合法")
    conn = get_db()
    c = conn.cursor()
    if c.execute("SELECT 1 FROM users WHERE LOWER(username)=LOWER(?)", [username]).fetchone():
        conn.close()
        raise ValueError(f"用户名 {username} 已存在")
    if email and c.execute("SELECT 1 FROM users WHERE email=?", [email]).fetchone():
        conn.close()
        raise ValueError(f"邮箱 {email} 已存在")
    if role == "pm":
        cats = _norm_cats(cats if cats is not None else PM_DEFAULT_CATS)
        if not cats:
            raise ValueError("产品经理必须至少选择一个管理品类")
        region_group = (region_group or "").strip()
        if region_group not in ("CDC", "CDI"):
            region_group = ""
    else:
        cats = ""
        region_group = ""
    c.execute(
        "INSERT INTO users(username,email,password,role,active,cats,region_group,created_at) "
        "VALUES(?,?,?,?,?,?,?,?)",
        [username, email or None, _hash_pw(password), role, 1, cats, region_group, _now()])
    new_id = c.lastrowid
    conn.commit()
    conn.close()
    return new_id


def update_user(uid, role=None, password=None, active=None, email=None, cats=None, region_group=None):
    """更新用户。role 变为 pm 时默认写入全部管理品类（可用 cats 指定）；
    显式传 cats 可调整 pm 用户的管理品类；离开 pm 时清空。
    region_group 仅对 pm 用户生效（CDC/CDI）。返回更新后的用户 dict。"""
    uid = int(uid)
    conn = get_db()
    c = conn.cursor()
    u = c.execute("SELECT * FROM users WHERE id=?", [uid]).fetchone()
    if not u:
        conn.close()
        raise ValueError("用户不存在")
    old_role = u["role"]
    sets, params = [], []
    if email is not None:
        email = (email or "").strip().lower()
        if email and not _is_email(email):
            conn.close()
            raise ValueError("邮箱格式不合法")
        if email and c.execute("SELECT 1 FROM users WHERE email=? AND id<>?", [email, uid]).fetchone():
            conn.close()
            raise ValueError(f"邮箱 {email} 已被其他用户使用")
        sets.append("email=?"); params.append(email or None)
    if role is not None:
        role = role.strip()
        if role not in ("admin", "pm", "user"):
            conn.close()
            raise ValueError("权限必须是 admin / pm / user 之一")
        sets.append("role=?"); params.append(role)
        if role == "pm":
            new_cats = _norm_cats(cats if cats is not None else PM_DEFAULT_CATS)
            if not new_cats:
                raise ValueError("产品经理必须至少选择一个管理品类")
            sets.append("cats=?"); params.append(new_cats)
            rg = (region_group or "").strip()
            sets.append("region_group=?"); params.append(rg if rg in ("CDC", "CDI") else "")
        elif old_role == "pm":
            sets.append("cats=?"); params.append("")
            sets.append("region_group=?"); params.append("")
    else:
        if cats is not None and old_role == "pm":
            # 未改角色但显式调整管理品类（仅 pm 生效）
            new_cats = _norm_cats(cats)
            if not new_cats:
                raise ValueError("产品经理必须至少选择一个管理品类")
            sets.append("cats=?"); params.append(new_cats)
        if region_group is not None and old_role == "pm":
            rg = (region_group or "").strip()
            sets.append("region_group=?"); params.append(rg if rg in ("CDC", "CDI") else "")
    if password is not None:
        if not password:
            conn.close()
            raise ValueError("密码不能为空")
        sets.append("password=?"); params.append(_hash_pw(password))
    if active is not None:
        sets.append("active=?"); params.append(1 if active else 0)
    if sets:
        c.execute("UPDATE users SET " + ",".join(sets) + " WHERE id=?", params + [uid])
    conn.commit()
    u = c.execute(
        "SELECT id,username,email,role,active,cats,region_group,created_at FROM users WHERE id=?", [uid]).fetchone()
    conn.close()
    return dict(u)


def set_active(uid, active):
    return update_user(uid, active=active)


def delete_user(uid):
    uid = int(uid)
    conn = get_db()
    c = conn.cursor()
    n = c.execute("DELETE FROM users WHERE id=?", [uid]).rowcount
    conn.commit()
    conn.close()
    if not n:
        raise ValueError("用户不存在")
    return n


def _is_email(s):
    import re
    return bool(re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", s or ""))


def export_vehicles_excel(vids=None, filters=None):
    """导出车型+产品为 Excel（LIST 汇总表 + 每个 VID 一个 sheet）。

    车型来源（二选一，优先 vids）：
    - vids: 前端传来的「当前查询结果显示的 VID 列表」（车型筛选 / VIN / VID 查询的结果）
    - filters: 兼容旧逻辑，按属性筛选条件匹配车型
    布局参考 D:\\linshi\\新模板.xlsx：
    - LIST: Brand | Vehicle(宽) | VID | VIN(超链接)
    - VID sheet: R2 Vehicle:/D2车型(合并D2:J2)  R3 VIN:/D3VIN  R4 VID:/D4VID  R6表头(从B列)  R7+数据
      列: B=#  C=Category(英文品类)  D=Part Description(英文品名)  E=CDAP No  F=OE Part No  G=Notes  H=Supplier  I=COST CNY  J=MOQ
    """
    import openpyxl
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

    conn = get_db()
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    # 1. 确定要导出的车型：优先用前端传来的查询结果 VID 列表；否则按筛选条件匹配
    if vids:
        vids = [str(x).strip() for x in vids if str(x).strip()]
        if not vids:
            conn.close()
            return None
        ph = ','.join(['?'] * len(vids))
        vehicles = c.execute(f"SELECT * FROM vehicles WHERE vid IN ({ph}) ORDER BY vid", vids).fetchall()
    else:
        vehicles = c.execute("SELECT * FROM vehicles ORDER BY vid").fetchall()
        for fk, fv in (filters or []):
            key = fk.upper()
            val = fv.upper() if key not in ('YEAR', 'MAKE', 'ORIGIN') else fv
            filtered = []
            for v in vehicles:
                field_map = {
                    'MAKE': ('brand',), 'MODEL': ('model',), 'SUBMODEL': ('sub',),
                    'YEAR': ('year',), 'CAPACITY': ('cap',), 'ENGINE': ('engine',),
                    'CHASSIS': ('chassis',), 'SERIES': ('series',), 'POWER': ('kw',),
                    'BODY': ('body',), 'FUEL': ('fuel',), 'DRIVE': ('drive',),
                    'ORIGIN': ('origin',), 'REGION': ('region',),
                }
                cols = field_map.get(key, ())
                if not cols:
                    continue
                match = False
                for col in cols:
                    cell_val = str(v[col] or '')
                    if key == 'YEAR':
                        if val and val in cell_val:
                            match = True
                            break
                    else:
                        if val and val in cell_val.upper():
                            match = True
                            break
                if match:
                    filtered.append(v)
            vehicles = filtered

    if not vehicles:
        conn.close()
        return None

    # 2. 获取每个车型的产品和备注
    vid_list = [str(v['vid']) for v in vehicles]
    placeholders = ','.join(['?'] * len(vid_list))
    all_products = c.execute(
        f"SELECT * FROM products WHERE vid IN ({placeholders}) ORDER BY vid, grp, cn",
        vid_list
    ).fetchall()

    prods_by_vid = {}
    for p in all_products:
        vid = str(p['vid'])
        prods_by_vid.setdefault(vid, []).append(dict(p))

    # 备注
    remarks_raw = c.execute(
        f"SELECT pv.*, p.vid as owner_vid FROM product_vehicles pv "
        f"LEFT JOIN products p ON p.id=pv.product_id "
        f"WHERE pv.vid IN ({placeholders})",
        vid_list
    ).fetchall()

    remark_map = {}
    for r in remarks_raw:
        rid = str(r['vid'])
        remark_map[(r['product_id'], rid)] = r['remark'] or ''

    # 品类英文映射（categories.code → en，如 F → BODY PARTS）
    cat_rows = c.execute("SELECT code, en FROM categories").fetchall()
    cat_en_map = {r['code']: (r['en'] or '') for r in cat_rows}

    conn.close()

    # 3. 构建 Excel
    wb = openpyxl.Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    bold_font = Font(bold=True)
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    wrap_align = Alignment(wrap_text=True, vertical='center', horizontal='left')
    center_align = Alignment(horizontal='center', vertical='center')

    def _build_vehicle_name(v):
        parts = [v['brand'] or '', v['model'] or '']
        if v['sub']: parts.append(v['sub'])
        if v['year']: parts.append(v['year'])
        extra = []
        if v['body']: extra.append(v['body'])
        if v['cap']: extra.append(f"{v['cap']} cc")
        if v['fuel']: extra.append(v['fuel'])
        if v['engine']: extra.append(v['engine'])
        if extra: parts.append(', '.join(extra))
        return ' '.join(parts).strip()

    # ---- LIST 汇总表 ----
    ws_list = wb.create_sheet("LIST", 0)
    list_headers = ["Brand", "Vehicle", "VID", "VIN"]
    for ci, h in enumerate(list_headers, 1):
        cell = ws_list.cell(row=1, column=ci, value=h)
        cell.font = bold_font
        cell.fill = header_fill
        cell.border = thin_border

    for ri, v in enumerate(vehicles, 2):
        vid_str = str(v['vid'])
        vehicle_name = _build_vehicle_name(v)
        vin_str = (v['vin17'] or '').replace(',', ' / ')
        ws_list.cell(row=ri, column=1, value=v['brand'] or '').border = thin_border
        ws_list.cell(row=ri, column=2, value=vehicle_name).border = thin_border
        ws_list.cell(row=ri, column=3, value=int(vid_str)).border = thin_border
        vin_cell = ws_list.cell(row=ri, column=4, value=vin_str)
        vin_cell.border = thin_border

    # LIST 列宽（匹配模板）
    ws_list.column_dimensions['A'].width = 8.89
    ws_list.column_dimensions['B'].width = 117.78
    ws_list.column_dimensions['C'].width = 8.89
    ws_list.column_dimensions['D'].width = 26.33

    # ---- 每个 VID 一个 sheet（匹配新模板布局）----
    prod_headers = ["#", "Category", "Part Description", "CDAP No", "OE Part No",
                     "Notes", "Supplier", "COST CNY", "MOQ"]
    # 新模板列宽
    vid_col_widths = {'A': 0.45, 'B': 3.78, 'C': 21.11, 'D': 41.22,
                       'E': 26.56, 'F': 27.67, 'G': 23.56, 'H': 24.44,
                       'I': 16.67, 'J': 10.89}

    for v in vehicles:
        vid_str = str(v['vid'])
        safe_name = vid_str[:31]
        try:
            ws = wb.create_sheet(safe_name)
        except Exception:
            safe_name = f"VID_{vid_str}"[:31]
            ws = wb.create_sheet(safe_name)

        vehicle_name = _build_vehicle_name(v)
        vin_str = (v['vin17'] or '').replace(',', ' / ')

        # R2: Vehicle 标签 + 车型名（D2:J2 合并）
        c2 = ws.cell(row=2, column=3, value="Vehicle:")
        c2.font = bold_font
        c2.alignment = Alignment(horizontal='left')
        d2 = ws.cell(row=2, column=4, value=vehicle_name)
        d2.alignment = wrap_align
        ws.merge_cells('D2:J2')

        # R3: VIN
        c3 = ws.cell(row=3, column=3, value="VIN:")
        c3.font = bold_font
        c3.alignment = Alignment(horizontal='left')
        ws.cell(row=3, column=4, value=vin_str).alignment = wrap_align

        # R4: VID
        c4 = ws.cell(row=4, column=3, value="VID:")
        c4.font = bold_font
        c4.alignment = Alignment(horizontal='left')
        ws.cell(row=4, column=4, value=int(vid_str)).alignment = wrap_align

        # R6: 表头（从 B 列开始）
        for ci, h in enumerate(prod_headers, 2):
            cell = ws.cell(row=6, column=ci, value=h)
            cell.font = bold_font
            cell.fill = header_fill
            cell.border = thin_border
            cell.alignment = center_align
        # Notes 列(G=col 7)特殊：wrap_text
        ws.cell(row=6, column=7).alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')

        # R7+: 数据行（从 B 列开始）
        prods = prods_by_vid.get(vid_str, [])
        for pi, p in enumerate(prods, 7):
            pid = p['id']
            rem = remark_map.get((pid, vid_str), '') or ((p['notes'] or ''))
            seq_no = pi - 6  # 序号从1开始
            cat_en = cat_en_map.get(p['grp'] or '', '')

            row_data = [
                seq_no,              # B: # 序号
                cat_en,              # C: Category（英文品类名）
                (p['en'] or ''),     # D: Part Description（英文品名）
                (p['cdap'] or ''),   # E: CDAP No
                (p['oe'] or ''),     # F: OE Part No
                rem,                 # G: Notes
                (p['supplier'] or ''),# H: Supplier
                (p['cost'] or ''),   # I: COST CNY
                (p['moq'] or ''),    # J: MOQ
            ]
            for ci, val in enumerate(row_data, 2):
                cell = ws.cell(row=pi, column=ci, value=val)
                cell.border = thin_border
                cell.alignment = wrap_align

        # 设置列宽（匹配模板）
        for col_letter, w in vid_col_widths.items():
            ws.column_dimensions[col_letter].width = w

    from io import BytesIO
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()



# ================= 反馈建议 =================

def add_feedback(user_id, username, content):
    """用户提交反馈"""
    from datetime import datetime
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    conn = get_db()
    c = conn.cursor()
    c.execute("INSERT INTO feedbacks(user_id, username, content, created_at) VALUES(?,?,?,?)",
              (user_id, username, content.strip(), now))
    conn.commit()
    fid = c.lastrowid
    conn.close()
    return fid


def list_feedbacks(user_id=None, is_admin=False):
    """获取反馈列表：管理员看全部，普通用户只看自己的"""
    conn = get_db()
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    if is_admin:
        rows = c.execute("SELECT * FROM feedbacks ORDER BY created_at DESC").fetchall()
    else:
        rows = c.execute("SELECT * FROM feedbacks WHERE user_id=? ORDER BY created_at DESC",
                         (user_id,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def reply_feedback(feedback_id, reply, replied_by):
    """管理员回复反馈"""
    from datetime import datetime
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    conn = get_db()
    c = conn.cursor()
    c.execute("UPDATE feedbacks SET reply=?, replied_by=?, replied_at=? WHERE id=?",
              (reply.strip(), replied_by, now, feedback_id))
    conn.commit()
    affected = c.rowcount
    conn.close()
    return affected > 0


def delete_feedback(feedback_id, user_id=None, is_admin=False):
    """删除反馈：管理员可删任何，用户只能删自己的"""
    conn = get_db()
    c = conn.cursor()
    if is_admin:
        c.execute("DELETE FROM feedbacks WHERE id=?", (feedback_id,))
    else:
        c.execute("DELETE FROM feedbacks WHERE id=? AND user_id=?", (feedback_id, user_id))
    conn.commit()
    affected = c.rowcount
    conn.close()
    return affected > 0


if __name__ == "__main__":
    import sys
    force = "--force" in sys.argv
    seed_from_excel(force=force)


def add_log(username, action, detail=''):
    """添加操作日志"""
    conn = get_db()
    c = conn.cursor()
    c.execute("INSERT INTO logs(username, action, detail, created_at) VALUES(?,?,?,?)", [username, action, detail, _now()])
    conn.commit()
    conn.close()

def get_logs(limit=100, offset=0):
    """获取操作日志（按时间倒序）"""
    conn = get_db()
    c = conn.cursor()
    total = c.execute("SELECT COUNT(*) FROM logs").fetchone()[0]
    rows = c.execute("SELECT id, username, action, detail, created_at FROM logs ORDER BY id DESC LIMIT ? OFFSET ?", [limit, offset]).fetchall()
    conn.close()
    return {'total': total, 'logs': [dict(r) for r in rows]}
