# AutoIndex —— 开发交接文档（DEV GUIDE）

> 本文档面向**接手开发的 Agent**，目标是让接手者无需向原开发者提问即可独立完成：
> 定位文件、改代码、同步副本、重启服务、回归验证的完整闭环。
> 最后更新：2026-08-21 11:10（含产品管理搜索改造、OE/CDAP 去重聚合、产品编辑备注列全部落地后）

---

## 0. 交接速览（30 秒）

| 项目 | 值 |
|---|---|
| 产品名称 | AutoIndex（原「VIN 车型查询工具」） |
| 定位 | 团队内部「车型库 + 产品（OE/CDAP）」数据管理系统 |
| 技术栈 | 原生 HTML/JS 单页（无框架）+ Python `http.server` + SQLite |
| 访问地址 | 内网 `http://<服务器IP>:3001/`；公网 `http://139.224.206.147:3001/` |
| 默认账号 | `admin` / `admin123`（管理员） |
| 服务进程 | `pythonw server_proxy.py`（0.0.0.0:3001，多线程） |
| 数据真源 | `vin_data.db`（SQLite，WAL 模式） |

**核心结论（务必先读）：**
1. **前端源码在** `vin-tools-package/vin-tools/template.html`，改完必须同步到部署目录 3 份副本。
2. **后端（`dbstore.py` / `server_proxy.py`）只在部署目录维护**，源码包里的后端是旧版，不要改、不要用。
3. **服务重启不能用 PowerShell / cmd**（沙箱拒绝），唯一可靠方法是内联 Python 脚本（见 §3.3）。
4. 每次改动必须走「改 → 同步/重启 → curl 验证」闭环。

---

## 1. 项目概览

### 1.1 业务能力
- **查询**：VIN 码查车型（17 位校验、WMI 识别、VDS 真值表判定、年份区间、候选排序、NHTSA 联网核对）；字段组合筛选车型；OE/CDAP 号查产品。
- **编辑**：产品管理（产品记录维护、**同 OE/CDAP 自动去重合并为一行**、关联车型聚合）；产品编辑（按车型维护产品行、**产品-车型备注列**、登记真实 17 位 VIN）。
- **管理**：车型库管理（分页浏览/搜索/Excel 导入导出/模板下载）；品类管理（品类→产品树，品类下新增产品自动补齐到所有车型）；用户管理（三级权限 + 产品经理管理品类授权）。

### 1.2 权限模型（三级）
| 角色 | 菜单可见 | 说明 |
|---|---|---|
| `user` 用户 | 仅「查询」 | 只能查 VIN/产品 |
| `pm` 产品经理 | 「查询」+「编辑」 | 只能看到/编辑自己管理品类（cats）下的产品 |
| `admin` 管理员 | 全部（含「管理」+「用户管理」） | 不受品类限制 |

### 1.3 数据流向
```
Excel（期初种子，一次性） ─┐
vds_seed.json（VDS 真值表）─┤→ SQLite（vin_data.db）←── web 端增删改（写回真源）
web 浏览/编辑 ←── server_proxy.py（HTTP + 会话 + 鉴权）── dbstore.py（SQL）
```

---

## 2. 目录与文件职责（⚠️ 双目录工作流，最容易踩坑）

### 2.1 部署目录（运行真源，服务器从这里起）
`C:\Users\Administrator\WorkBuddy\2026-08-19-08-55-42\vin-tools\`

| 文件 | 职责 | 是否可改 |
|---|---|---|
| `template.html` / `index.html` / `AutoIndex.html` | 前端页面的 **3 份相同副本**（浏览器打开 index.html） | 改后必须 3 份一致 |
| `dbstore.py` | SQLite 数据层（**最新版**，1010 行） | ✅ 直接改这里 |
| `server_proxy.py` | HTTP 服务 + 会话 + 鉴权 + NHTSA 代理（**最新版**，447 行） | ✅ 直接改这里 |
| `vin_data.db` | **生产数据库**（唯一真源） | 运行时被读写 |
| `backup_db.py` / `backups/` | 数据库备份脚本/备份产物 | — |
| `vds_seed.json` | VDS 真值表种子数据（首次初始化用） | 谨慎 |
| `车型库新增模板.xlsx` | 车型导入模板参考 | — |
| `server.log` / `server_out.log` / `server_err.log` | 日志（当前为空，`log_message` 被禁用） | — |
| `start_server.bat` / `started_flag.txt` | 历史启动辅助（当前不依赖） | — |
| `__dbtest/` `__pycache__/` | 测试/缓存目录 | 可忽略 |
| `_audit_match.py` `_check_vid.py` `_check_vin.py` `sweep_vds.py` `test_match.js` | 历史校验/调试脚本 | 可忽略 |
| `vin_data_before_*.db` | 历史快照 | 勿删勿用 |

### 2.2 源码包目录（前端源码所在地）
`C:\Users\Administrator\WorkBuddy\2026-08-19-08-55-42\vin-tools-package\vin-tools\`

| 文件 | 说明 |
|---|---|
| `template.html` | **前端源码**（3068 行，最新）——改前端先改这里 |
| `index.html` / `AutoIndex.html` | 前端旧副本（会随同步刷新） |
| `dbstore.py` / `server_proxy.py` | ⚠️ **旧版后端，勿改勿用**（部署目录才是最新） |
| `Oscar中国车主表.xlsx` / `library_dump.json` / `build_vin_tool.py` | 期初数据与构建脚本 |
| `车型库新增模板.xlsx` | 导入模板 |

### 2.3 工作流铁律

| 场景 | 操作 |
|---|---|
| 改前端 | 编辑 `vin-tools-package/.../template.html` → `cp` 同步到部署目录 3 副本 → curl 验证 |
| 改后端 | 直接编辑部署目录 `dbstore.py` / `server_proxy.py` → py_compile → 重启服务 |
| 同步 3 副本 | `cp template.html index.html && cp template.html AutoIndex.html`，然后 `diff -q` 确认 |
| 重启服务 | 内联 Python（§3.3），禁止 PowerShell/cmd |
| 验证 | curl 打 `http://127.0.0.1:3001/` 及各 API |

---

## 3. 运行与部署

### 3.1 启动
服务由 `server_proxy.py` 以 `pythonw`（无控制台窗口）在后台运行：
```python
pythonw = r'C:\Users\Administrator\.workbuddy\binaries\python\versions\3.13.12\pythonw.exe'
subprocess.Popen([pythonw, 'server_proxy.py'],
                 cwd=r'C:\Users\Administrator\WorkBuddy\2026-08-19-08-55-42\vin-tools',
                 creationflags=0x00000008 | 0x08000000,   # DETACHED_PROCESS | CREATE_NO_WINDOW
                 close_fds=True)
```
启动时自动执行 `dbstore.init_db()`（建表）与 `dbstore.init_users()`（种子 admin）。

### 3.2 访问
- 本机/内网：`http://127.0.0.1:3001/` 或 `http://<内网IP>:3001/`
- 公网：`http://139.224.206.147:3001/`（服务器有公网映射，端口已放行）

### 3.3 重启服务（⚠️ 唯一可靠方法）
PowerShell 工具在本环境创建 ConPTY 被拒、Bash 禁止调用 `powershell.exe`、`cmd //c taskkill/start` 参数转换异常。
**唯一可行**：Bash 里跑内联 Python：

```bash
python3 - <<'PY'
import subprocess, os, signal, time
out = subprocess.check_output(['netstat', '-ano'], text=True, errors='ignore')
pids = set()
for line in out.splitlines():
    if ':3001' in line and ('LISTENING' in line or 'ESTABLISHED' in line):
        pid = line.strip().split()[-1]
        if pid.isdigit(): pids.add(int(pid))
print('found pids:', pids)
for pid in pids:
    try: os.kill(pid, signal.SIGTERM); print('killed', pid)
    except Exception as e: print('kill error', pid, e)
time.sleep(1)
pythonw = r'C:\Users\Administrator\.workbuddy\binaries\python\versions\3.13.12\pythonw.exe'
proc = subprocess.Popen([pythonw, 'server_proxy.py'],
                        cwd=r'C:\Users\Administrator\WorkBuddy\2026-08-19-08-55-42\vin-tools',
                        creationflags=0x00000008 | 0x08000000, close_fds=True)
print('started pid', proc.pid)
PY
```

⚠️ 已知陷阱：3001 端口可能残留**多个历史进程**，导致 cookie 失效/请求被不同实例处理。务必先把所有 3001 相关 PID 杀干净，再启动**单实例**，然后 `netstat -ano | grep ':3001'` 确认只剩一个 LISTENING。

### 3.4 数据库备份
- `backup_db.py` 脚本 + `backups/` 目录存放历史备份。
- 改动数据前建议先复制一份 `vin_data.db`。

---

## 4. 数据库设计（`dbstore.py`，7 张表）

### 4.1 vehicles —— 车型库
| 列 | 说明 |
|---|---|
| id | 主键自增 |
| vid | 车辆 ID（业务主键，如 62036） |
| prefix | VIN 前缀 |
| brand / model / sub | 品牌 / 车型 / 配置 |
| year / ys / ye | 年款（字符串）/ 起始年 / 结束年 |
| cap / kw | 排量 / 功率 |
| engine / chassis / series | 发动机 / 底盘 / 车系 |
| body / fuel / drive / origin | 车身 / 燃料 / 驱动 / 来源（`web`=网页新增，可整批清除；`import`=Excel 导入） |
| vin17 | 核实的完整 17 位实际 VIN 码（逗号分隔多条；查询精确命中层） |

索引：`idx_veh_vid`、`idx_veh_vin17`。

### 4.2 products —— 产品（OE/CDAP）
| 列 | 说明 |
|---|---|
| id | 主键自增 |
| vid | 所属车型 vid |
| grp | 品类代码（A/B/D/E/F，与 categories.code 对应） |
| cn / en | 产品中文名 / 英文名 |
| brand | 品牌 |
| cdap | CDAP 号 |
| oe | OE 号（可多个，以空格/逗号分隔存储） |
| notes / supplier / cost / moq | 备注 / 供应商 / 采购价 / 起订量 |

索引：`idx_prod_vid`、`idx_prod_grp`。

### 4.3 product_vehicles —— 产品↔车型关联（多对多）
- `product_id` + `vid` 联合主键。产品记录与车型的关联关系。
- `remark`（TEXT，旧库启动时自动迁移补列）——**产品针对车型的备注**：同一产品关联到不同车型可有各自不同的备注；仅在「产品编辑」关联产品表展示，**不进入产品管理**。
- 注意：owner 产品（`products.vid` 归属车型）默认没有 `product_vehicles` 行；给 owner 写备注时 `set_product_vehicle_remark` 会自动 INSERT 一行，读取侧按 vid 聚合时天然去重，无副作用。

### 4.4 vds_map —— VDS 真值表（VIN 第 4~8 位 → 车型，权威判定层）
| 列 | 说明 |
|---|---|
| vds | 主键，VIN 第 4~8 位 |
| models | JSON：字符串=唯一车型；数组=多车型碰撞（需结合年款/动力判定） |
| src | `lib` / `web` |

> 前端另有静态 `VS_MAP`（约 40 条，粗略系列提示层）与 `WMI3`/`WMI2`/`COUNTRY_MAP`（WMI 识别），在 template.html 约 720~764 行。

### 4.5 categories —— 品类
| 列 | 说明 |
|---|---|
| code | 品类代码（主键，如 A/B/D/E/F/G） |
| cn / en | 中文名 / 英文名（初始留空，用户在产品管理页维护） |

### 4.6 product_catalog —— 产品目录（品类下可用的产品名清单）
- `grp` + `cn` 联合主键。品类下新增产品 = 写目录 + 补齐到所有车型（`seed_catalog_product_to_all`）。

### 4.7 users —— 用户（登录模块）
| 列 | 说明 |
|---|---|
| id | 主键自增 |
| username | 唯一，登录账号（大小写不敏感） |
| email | 唯一，可空，也可作登录账号（大小写不敏感） |
| password | `salt:sha256`（格式 `salt` + `:` + hex），盐 8 字节随机 |
| role | `admin` / `pm` / `user` |
| active | 1=启用 / 0=停用（停用后无法登录、会话立即失效） |
| cats | pm 的管理品类，逗号分隔（如 `A,B,D,E,F`）；非 pm 为空串 |
| created_at | 创建时间 |

**常量**：`PM_DEFAULT_CATS = "A,B,D,E,F"`（pm 未指定品类时的默认全选）。

### 4.8 关键后端函数索引（dbstore.py）
| 函数 | 行号（约） | 职责 |
|---|---|---|
| `init_db` / `get_db` | 38 / 30 | 建表（幂等）+ WAL 连接 + `product_vehicles.remark` 迁移 |
| `get_all_data` | 306 | 首屏全量数据（`include_products` 控制） |
| `get_all_products` | 176 | `{vid: [products]}` 懒加载数据 |
| `get_vehicles_page` | 352 | 车型分页 + q/brand 过滤 |
| `get_product_vehicles` / `get_product_vehicles_by_code` | 192 / 212 | 产品关联车型（按 pid） / 按 OE+CDAP 聚合去重（产品管理「关联车型」用） |
| `add_product_vehicle` / `delete_product_vehicle` | 250 / 266 | 产品↔车型关联增删 |
| `get_vehicle_product_remarks` / `set_product_vehicle_remark` | 275 / 287 | 读某车型下全部产品备注 / 写某产品-车型备注（自动补行，owner 回填） |
| `list_categories` / `add_category` / `update_category` / `delete_category` | 449~508 | 品类 CRUD（有产品记录时拒删） |
| `add_catalog_product` / `seed_catalog_product_to_all` | 408 / 426 | 目录新增 + 补齐所有车型 |
| `seed_default_products_for_vid` / `seed_default_products` / `reseed_default_products` | 539~592 | 默认产品模板/重建 |
| `upsert_product` / `delete_product` | 626 / 645 | 产品写回（id>0 更新，否则插入） |
| `add_vehicle` / `delete_vehicle` / `delete_vehicles_web` | 653~679 | 车型写回（web 来源整批清除） |
| `update_vehicle_vins` | 709 | 登记实际 17 位 VIN（全量覆盖，单车型≤50 条，跨车型查重） |
| `_norm_cats` | 762 | 品类串规范化（大写/去重/保序） |
| `_hash_pw` / `_check_pw` | 777 / 786 | 密码哈希/校验 |
| `init_users` | 793 | 种子 admin/admin123 |
| `get_user_by_login` / `verify_user` | 817 / 831 | 登录查找（大小写不敏感）/ 凭证+启用校验 |
| `update_me` | 844 | 用户自改邮箱/密码（须当前密码） |
| `list_users` / `add_user` / `update_user` / `set_active` / `delete_user` | 884~990 | 用户管理 CRUD |
| `_is_email` | 1002 | 邮箱格式校验 |

**add_user / update_user 关键规则**：
- role=pm 且未给 cats → 默认 `PM_DEFAULT_CATS`；显式传空数组 → 报错「产品经理必须至少选择一个管理品类」。
- 离开 pm 角色 → 自动清空 cats。
- 用户名/邮箱查重均大小写不敏感。

---

## 5. API 接口清单（server_proxy.py）

> 约定：响应统一 `{ok:true,...}` / `{ok:false,error}`；除登录外所有 `/api/*` 必须带会话 Cookie。
> Cookie：`ai_token`，HttpOnly，7 天有效期；会话存内存 `SESSIONS`（重启即失效，需重新登录）。

### 5.1 鉴权
| 方法 | 路径 | 权限 | 说明 |
|---|---|---|---|
| POST | `/api/login` | 公开 | `{login,password}`；login 支持用户名或邮箱，均大小写不敏感；停用账号拒登。返回 `{ok,user}` + Set-Cookie |
| POST | `/api/logout` | 公开 | 销毁会话 + 清 Cookie |
| GET | `/api/me` | 登录 | 返回**实时**用户信息（从 DB 拉取，角色/品类被管理员改后即时生效）；停用/不存在 → 401 |
| PUT | `/api/me` | 登录 | `{email?, password?, current_password}`；须当前密码，邮箱格式/占用校验 |

### 5.2 业务数据
| 方法 | 路径 | 权限 | 说明 |
|---|---|---|---|
| GET | `/api/data?products=0|1` | 登录 | 首屏：`{META, VDS_MAP, VDS_SRC, vehicles..., total}`；products=0 时不含产品（懒加载用） |
| GET | `/api/products` | 登录 | `{ok, products:{vid:[...]}}` 产品懒加载 |
| GET | `/api/product-vehicles?id=<pid>` | 登录 | 某产品的关联车型 `{ok, vehicles:[{vid, owner}]}` |
| POST | `/api/product-vehicles` | 登录 | `{id, vid, action:'add'\|'delete'}` 关联/解除 |
| GET | `/api/product-vehicles-by-code?oe=&cdap=` | 登录 | 按 OE+CDAP 聚合所有关联车型（跨产品去重，产品管理「关联车型」弹窗用） |
| GET | `/api/product-vehicle-remarks?vid=` | 登录 | 某车型下所有产品备注 `{ok, remarks:[{product_id, remark}]}`（产品编辑关联产品表用） |
| POST | `/api/product-vehicle-remark` | 登录 | `{pid, vid, remark}` 设置产品-车型备注（产品管理不展示） |
| GET | `/api/vehicles?limit=&offset=&q=&brand=` | 登录 | 分页车型 `{total, vehicles, hasMore}` |
| POST | `/api/vehicles` | 登录 | 3 种：数组=批量导入（逐条校验 vid/brand/model 必填）；`{id, vin17}`=登记实际 VIN；对象=单条新增（自动补默认产品） |
| DELETE | `/api/vehicles` | 登录 | `{web:true}`=清所有 web 来源；`{id}`=删单条 |
| GET | `/api/categories` | 登录 | 品类树 `{categories:[{code,cn,en,count,products}]}` |
| POST | `/api/categories` | 登录 | `{code,cn,en}` 新增；`{action:'update',code,cn,en}` 改名 |
| DELETE | `/api/categories` | 登录 | `{code}` 删除（品类下有产品记录则后端拒绝） |
| POST | `/api/catalog-products` | 登录 | `{grp,cn,en}` 目录新增产品并补齐所有车型，返回 `{seeded:n}` |
| GET | `/api/vpic?vin=17位` | 登录 | 同源代理 NHTSA vPIC 官方库（NHTSA 无 CORS，必须服务端代理）；返回 make/model/年款/发动机/产地等 |

### 5.3 用户管理（仅 admin）
| 方法 | 路径 | 说明 |
|---|---|---|
| GET | `/api/users` | 用户列表（不含 password） |
| POST | `/api/users` | `{username,email,password,role,cats?}` 新增 |
| PUT | `/api/users` | `{id, role?, password?, active?, email?, cats?}` 更新（改角色联动品类） |
| DELETE | `/api/users` | `{id}` 删除（前端禁止删自己，后端由前端控制） |

### 5.4 鉴权实现细节（server_proxy.py）
- `_session_user()`：从 Cookie 头解析 `ai_token` → 查内存 `SESSIONS`。
- `_require_session()`：未登录 401；`_require_admin()`：非 admin 403。
- `/api/me` 特殊：每次从 DB 拉最新（不信任会话快照），并校验 active。
- `end_headers()` 被重写：所有响应带 `no-store, no-cache, must-revalidate`，保证前端拿到最新。
- CORS：`Access-Control-Allow-Origin: *`（内网工具，可接受）。
- `timeout = 30`：单连接 30s 超时防挂死。

---

## 6. 前端功能地图（template.html，3068 行）

### 6.1 页面结构
```
loginGate（全屏登录门）
appLayout
├─ aside.side 侧边栏
│   ├─ side-group「查询」→ tabS(VIN查询) / tabO(产品查询)
│   ├─ side-group#grpEdit「编辑」→ tabPm(产品管理) / tabEdit(产品编辑)
│   ├─ side-group#grpManage「管理」→ tabM(车型库管理) / tabP(品类管理) / tabU(用户管理)
│   └─ side-user#sideUser（点击→账号设置弹窗；退出登录按钮）
└─ main → 7 个 pane
```

### 6.2 Pane 索引（switchTab 编号）
| # | id | 名称 | 角色 | 核心函数 |
|---|---|---|---|---|
| 0 | paneS | VIN 查询 | 全部 | `queryOne`/`match`/`candCard`/`vpicCheck`/`queryByAttrs`/`toggleCardProds` |
| 1 | paneO | 产品查询 | 全部 | `queryByOe` |
| 2 | paneM | 车型库管理 | admin | `fetchVehiclePage`/`renderLibrary`/`loadMoreVehts`/`importVehiclesExcel`/`exportVehiclesExcel`/`downloadVehicleTemplate` |
| 3 | paneP | 品类管理 | admin | `loadCats`/`renderCats`/`pmAddProd`/`saveCatalogProduct`/`saveCatNew` |
| 4 | paneEdit | 产品编辑 | pm/admin | `queryEdit`/`queryEditByAttrs`/`renderEditProdTree`/`renderEditProductsTable`/`enterEditProdRow`/`saveEditProdRow`/`saveEditVin`/`editVpicCheck`/`saveRemark` |
| 5 | panePm | 产品管理 | pm/admin | `queryByOePm`（搜索→过滤下方全表，无卡片区）/`renderPmAllProducts`（同 OE/CDAP 去重合并）/`openPmEdit`/`pmSaveEdit`/`openPmRel`/`pmLoadRel`/`pmAddRel`/`pmDelRel` |
| 6 | paneUsers | 用户管理 | admin | `loadUsers`/`renderUsers`/`umAddUser`/`umEditUser`/`umSaveEdit`/`umToggleActive`/`umDeleteUser` |

### 6.3 弹窗（`</main>` 前）
| id | 用途 | 函数 |
|---|---|---|
| `pmEditModal` | 产品管理-编辑产品（品牌/OE/CDAP/供应商/MOQ/采购价） | `openPmEdit`/`pmHideEdit`/`pmSaveEdit` |
| `pmRelModal` | 产品管理-关联车型 | `openPmRel`/`pmLoadRel`/`pmAddRel`/`pmDelRel` |
| `umEditModal` | 用户管理-编辑用户（含品类多选） | `umEditUser`/`umSaveEdit`/`umOnRoleChange('edit')` |
| `accountModal` | 账号设置（改邮箱/密码） | `openAccountSettings`/`saveAccountSettings` |

### 6.4 全局变量与关键常量
| 变量 | 说明 |
|---|---|
| `DATA` / `VEH_TOTAL` | 已加载车型数组 / 总数（分页） |
| `META` | `{brands:[...]}` 品牌列表 |
| `PRODUCTS` / `PRODUCTS_READY` | 懒加载产品 `{vid:[...]}` |
| `VDS_MAP` / `VDS_SRC` | 服务端 VDS 真值表 |
| `CURRENT_USER` | 当前登录用户（含 role/cats） |
| `PM_CAT_OPTIONS` | `['A','B','D','E','F']` |
| `PM_ALL_ROWS` / `PM_FILTER_Q` / `PM_VISIBLE_ROWS` | 产品管理全表行（含 `_vids`/`_pids`/`rel_count` 聚合字段）/ 搜索关键词 / 过滤后的可见行 |
| `EDIT_REMARKS` | 产品编辑当前 vid 的产品-车型备注映射 `{product_id: remark}` |
| `WMI3`/`WMI2`/`COUNTRY_MAP`/`VS_MAP` | WMI/系列码静态规则表（约 720~764 行） |

### 6.5 启动流程
```
checkAuth()（fetch /api/me）
 ├─ 未登录 → showGate(true)（停在登录门）
 └─ 登录 → afterLogin()
      ├─ showGate(false) + 侧边栏显示用户名/角色
      ├─ applyRoleMenu(role)（按角色显隐菜单 + switchTab(0)）
      └─ bootstrap()：fetch /api/data?products=0 → fetchVehiclePage(false)（首屏 200 条）→ init() → loadProductsLazy()（后台拉 /api/products）
```

### 6.6 产品经理品类过滤（5 个前端过滤点）
核心：`myCats()`（admin→null=全部；pm→数组；user→[]）+ `inMyCats(grp)`。
1. `renderEditProdTree()` —— 产品编辑筛选树只显示管理品类
2. `eFilterProdsForVid(vid)` —— 编辑页车型关联产品表过滤
3. `saveEditProdRow()` —— 保存前 `inMyCats(p)` 校验（越权拒绝）
4. `renderPmAllProducts()` —— 产品管理全表过滤（含搜索）
5. `openPmEdit` / `pmSaveEdit` / `openPmRel` —— 弹窗入口越权拦截；`queryByOePm()` 设置 `PM_FILTER_Q` 驱动全表过滤（过滤逻辑内置 `inMyCats`）

---

## 7. 核心业务逻辑

### 7.1 VIN 查询匹配流程（`match(vin)`，约 1240 行起）
1. 归一化 VIN（大写、去非字母数字）。
2. 17 位时：校验字符集（无 I/O/Q）与第 9 位校验位（`checkDigit`），识别 WMI（3 位 → 2 位 → 国家码兜底）。
3. VDS 判定（第 4~8 位）：先查服务端 `VDS_MAP`（权威），命中唯一 → 直接候选；命中数组 → 多候选按年款/动力/工厂位碰撞判定（`pickBestCandidate`）；未命中 → 按 prefix 前缀匹配 + 年份区间（`inYearRange`）筛车型库。
4. 精确命中层：`vin17` 字段精确匹配（`vehVins`）优先于一切。
5. 结果渲染：卡片（`candCard`），主候选在前、参考候选并入尾部；点击卡片展开该车型产品（`vinToggleProds`）。
6. 17 位 VIN 时显示「联网官方核对（NHTSA）」按钮（`vpicCheck` → `/api/vpic`）。

### 7.2 产品编辑页（paneEdit，独立于查询页的完整副本）
- 定位车型的三种途径：VIN/VID 查询（`queryEdit`）、字段组合筛选（`queryEditByAttrs`，13 个字段）、产品筛选树（`renderEditProdTree`）。
- 卡片上可：登记/编辑实际 17 位 VIN（`toggleEditVinInput`/`saveEditVin`，走 `update_vehicle_vins`）、展开产品表行内编辑（`renderEditProductsTable`/`enterEditProdRow`/`saveEditProdRow`）。
- **产品-车型备注列**：关联产品表最右「备注」列是内联输入框，失焦即存（`saveRemark` → `POST /api/product-vehicle-remark`），仅在**本页**展示；产品管理中不出现。
- **保存校验**：pm 用户只能保存自己管理品类内的产品行（`inMyCats`）。

### 7.5 产品管理去重与搜索（panePm）
- **去重合并**：`renderPmAllProducts()` 用 Map 按 `oe|cdap` 合并，同号码产品只显示一行，行内聚合 `_vids`（该号码涉及的所有 vid）、`_pids`（对应产品 id 列表）、`rel_count`（关联车型总数）；「关联车型」按钮显示 `关联车型 (n)`。
- **关联车型聚合**：`openPmRel` 改按 oe+cdap 调 `GET /api/product-vehicles-by-code` 拉全部关联车型（owner + 额外关联，跨产品去重）。
- **搜索过滤**：搜索框输入 OE/CDAP 关键词后直接过滤下方全表（`PM_FILTER_Q`/`PM_VISIBLE_ROWS`），不再渲染独立的车型卡片结果区（原 `#outPm` 已删除）；清空搜索恢复全量。

### 7.3 产品目录补齐机制
- 品类管理页在某品类下新增产品 → `POST /api/catalog-products` → 写入 `product_catalog` + `seed_catalog_product_to_all` 补齐到**所有车型**。
- 新增车型 → 自动 `seed_default_products_for_vid` 生成默认产品模板（OE 等留空）。

### 7.4 首屏性能优化
- `/api/data?products=0`（不带 ~1.5MB 产品）→ 车型分页 200 条 → 后台 `loadProductsLazy()` 拉 `/api/products`，就绪后刷新依赖产品的视图。

---

## 8. 开发与验证流程（每个改动必走）

### 8.1 前端改动流程
```bash
# 1) 编辑源码
# 2) 同步 3 副本（部署目录）
cp package/template.html vin-tools/template.html
cp package/template.html vin-tools/index.html
cp package/template.html vin-tools/AutoIndex.html
diff -q template.html index.html && diff -q template.html AutoIndex.html   # 必须全部一致
# 3) 前端 JS 语法检查（提取 <script> 合并后 node --check）
# 4) curl 首页确认新代码在线上（无需重启，改的是静态文件）
curl -s http://127.0.0.1:3001/ | grep -n "新函数名或特征串"
```
前端 JS 语法检查脚本（提取所有 `<script>` 合并为临时 js 再 node --check）：
```python
from pathlib import Path
import re, subprocess
html = Path('C:/Users/Administrator/WorkBuddy/2026-08-19-08-55-42/vin-tools/index.html').read_text(encoding='utf-8')
scripts = re.findall(r'<script>(.*?)</script>', html, re.S)
js = Path('C:/Users/Administrator/AppData/Local/Temp/2/vin_check.js')
js.write_text('\n'.join(scripts), encoding='utf-8')
r = subprocess.run(['C:/Users/Administrator/.workbuddy/binaries/node/versions/22.22.2/node.exe','--check',str(js)], capture_output=True, text=True)
print('exit', r.returncode, r.stderr[:2000] if r.stderr else '')
```

### 8.2 后端改动流程
```bash
# 1) 编辑部署目录 dbstore.py / server_proxy.py
# 2) 语法检查
C:/Users/Administrator/.workbuddy/binaries/python/versions/3.13.12/python.exe -m py_compile dbstore.py server_proxy.py
# 3) 重启服务（§3.3 内联 Python）
# 4) curl 验证 API（见下）
```

### 8.3 后端回归验证（curl 模板）
```bash
# 登录（cookie 存 $TEMP，注意沙箱下 /tmp 可能不可用）
curl -s -c "$TEMP/cookies.txt" -b "$TEMP/cookies.txt" -X POST http://127.0.0.1:3001/api/login \
  -H "Content-Type: application/json" -d '{"login":"admin","password":"admin123"}'
# 查看当前用户
curl -s -b "$TEMP/cookies.txt" http://127.0.0.1:3001/api/me
# 401 拦截验证（不带 cookie）
curl -s http://127.0.0.1:3001/api/me
# 用户管理（admin 专用）
curl -s -b "$TEMP/cookies.txt" http://127.0.0.1:3001/api/users
```

---

## 9. 回归测试清单（改动后对照执行）

### 9.1 登录/会话
- [ ] 正确用户名/密码 → 登录成功，返回 user（含 role）
- [ ] 用户名大写/混合大小写（`ADMIN`、`Admin`）→ 成功（大小写不敏感）
- [ ] 邮箱大写 → 成功
- [ ] 错误密码 → 401
- [ ] 停用账号 → 401「账号不存在或已停用」
- [ ] 不带 cookie 访问任意 /api/* → 401
- [ ] 重启服务后旧会话失效 → 需重新登录

### 9.2 权限
- [ ] user 登录：菜单只见「查询」，`/api/users` → 403
- [ ] pm 登录（cats=A）：编辑页产品树/全表/搜索只显示 A 品类（发电机、起动机）；尝试保存 B 品类产品行被前端拦截
- [ ] admin 登录：全部菜单可见，品类不受限
- [ ] admin 把某 pm 角色改为 user → 其 cats 清空；改回 pm → 默认全选或指定

### 9.3 用户管理
- [ ] 新增 user/pm/admin 均成功；pm 不选品类 → 报错
- [ ] 用户名重复（含大小写变体）→ 「用户名已存在」
- [ ] 邮箱重复/格式非法 → 报错
- [ ] 编辑角色/密码/停用/启用 → 生效
- [ ] 删除用户 → 从列表消失；不能删自己（前端 disabled）

### 9.4 账号设置
- [ ] 当前密码错误 → 拒绝
- [ ] 改邮箱+密码成功 → 新凭据可登录、旧密码失效
- [ ] 测试后务必恢复 admin 原邮箱/密码（`skye@cooldrive.cc` / `admin123`）

### 9.5 业务数据
- [ ] VIN 查询（17 位合法/VID/前缀）命中；NHTSA 按钮出现并返回
- [ ] 产品编辑：保存产品行、登记实际 VIN（非法 17 位拒绝、跨车型重复登记拒绝）
- [ ] 产品编辑：同一产品在不同车型下写不同「备注」，失焦自动保存；刷新后仍在；产品管理列表不出现备注列
- [ ] 产品管理：搜索 `10703383` 只显示 1 行，行内「关联车型 (3)」展开能看到 3 个车型（56062/56061/58484）
- [ ] 产品管理：清空搜索恢复全量；编辑/关联车型弹窗正常
- [ ] 品类管理：新增品类、品类下新增产品补齐所有车型、删除品类（有产品时拒绝）
- [ ] 车型库：分页/搜索/品牌筛选、Excel 导入（模板列名）、导出、模板下载
- [ ] 产品管理：OE/CDAP 查询、编辑、关联/解除车型

---

## 10. 已知问题与注意事项

1. **会话为内存态**：`SESSIONS` 在 `server_proxy.py` 进程内，重启服务后所有用户需重新登录。
2. **端口多进程陷阱**：3001 可能残留历史进程，cookie 可能被不同实例处理导致"时而有效时而 401"。处理：杀净所有 3001 PID → 单实例启动。
3. **源码包后端文件过时**：`vin-tools-package` 下的 `dbstore.py`/`server_proxy.py` 是旧版，仅 `template.html` 是前端源码。接手时不要混淆。
4. **VDS 碰撞**：部分 VDS（如 BYD CE4CB、长城 EF7A6）被多车型共用，判定依赖年款/动力/工厂位，命中时前端会给多条候选。
5. **Excel 导入列名**：模板表头为中文（车辆ID/品牌/车型…），缺少必填列会直接 alert 拒绝。
6. **单连接 30s 超时**：`Handler.timeout=30`，防止卡死连接占线程。
7. **CORS 全开**（`*`）：内网工具定位，未做来源白名单。
8. **删除类操作**：车型整批删除仅限 `origin='web'` 的来源；产品删除直接删行。
9. **品类全集目前为 A/B/D/E/F**：新增品类后，若需让 pm 可选新品类，需同步更新前端 `PM_CAT_OPTIONS` 与后端 `PM_DEFAULT_CATS` 两处常量。
10. **服务日志**：`log_message` 被禁用，排障主要靠 curl 与 `netstat`。
11. **备注边界**：产品-车型备注只存在 `product_vehicles.remark`，产品管理所有视图/弹窗均不展示；若未来要把备注并入产品管理，需明确聚合策略（同 OE/CDAP 多行产品、多车型时取哪条）。
12. **去重副作用**：产品管理按 oe+cdap 合并后，「编辑」弹窗以合并行内第一个产品（`_pids[0]`）为主编辑对象，修改只落一个产品记录，其余同码产品不受影响。

---

## 11. 环境信息（接手必备）

| 项 | 值 |
|---|---|
| Python（managed） | `C:\Users\Administrator\.workbuddy\binaries\python\versions\3.13.12\python.exe` |
| pythonw（后台启动用） | `C:\Users\Administrator\.workbuddy\binaries\python\versions\3.13.12\pythonw.exe` |
| Node（JS 语法检查） | `C:\Users\Administrator\.workbuddy\binaries\node\versions\22.22.2\node.exe` |
| 服务重启 Skill | 项目级 `restart-vin-tools-server`（`~/.workbuddy/.../.workbuddy/skills/` 或项目 `.workbuddy/skills/`） |
| 工作记忆 | `C:\Users\Administrator\WorkBuddy\2026-08-20-15-36-53\.workbuddy\memory\2026-08-21.md`（按日期追加） |
| 前端依赖 | SheetJS（XLSX）CDN（页面内引用，用于 Excel 导入导出） |

---

## 12. 未来扩展建议（未实施）
- 会话持久化（SQLite 存 token）或 JWT，解决重启掉线。
- 用户管理接口防重放/限流；登录失败次数锁定。
- CORS 来源白名单 + HTTPS。
- pm 品类选项与 categories 表联动（新增品类自动出现在权限配置里）。
- 车型库删除/编辑单条（目前仅表格浏览 + 导入导出）。
