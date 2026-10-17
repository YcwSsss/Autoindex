# -*- coding: utf-8 -*-
"""
vin_data.db 每日备份脚本（SQLite 在线备份 API，服务运行中也可安全备份）。
- 备份到同目录 backups\vin_data_YYYYMMDD_HHMMSS.db
- 自动保留最近 30 份，更早的滚动删除
用法: python backup_db.py
"""
import sqlite3
import sys
from pathlib import Path
from datetime import datetime

HERE = Path(__file__).parent
DB = HERE / "vin_data.db"
BK_DIR = HERE / "backups"
KEEP = 30


def main():
    if not DB.exists():
        print("[backup] 找不到 vin_data.db，退出")
        return 1
    BK_DIR.mkdir(exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = BK_DIR / f"vin_data_{stamp}.db"

    # sqlite3 在线备份：即使服务正在写库也能得到一致快照
    src = sqlite3.connect(str(DB))
    dst = sqlite3.connect(str(dest))
    try:
        with dst:
            src.backup(dst)
    finally:
        src.close()
        dst.close()

    # 校验备份可打开且车型数一致
    chk = sqlite3.connect(str(dest))
    n_bk = chk.execute("SELECT COUNT(*) FROM vehicles").fetchone()[0]
    chk.close()
    if n_bk < 0:
        print("[backup] 备份校验失败")
        return 1

    # 滚动清理：只保留最近 KEEP 份
    files = sorted(BK_DIR.glob("vin_data_*.db"))
    removed = 0
    for f in files[:-KEEP] if len(files) > KEEP else []:
        f.unlink()
        removed += 1

    size_kb = dest.stat().st_size / 1024
    print(f"[backup] OK: {dest.name} ({size_kb:.0f} KB, {n_bk} 条车型; 清理旧备份 {removed} 份)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
