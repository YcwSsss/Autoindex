#!/usr/bin/env python3
# AutoIndex —— 本地服务（端口 3001）
# 作用：① 托管 vin-tools 目录下的静态页面（index.html 等）
#      ② /api/vpic?vin=XXXX 同源代理 NHTSA 官方 VIN 库（NHTSA 不开放 CORS，浏览器无法直接调用）
#      ③ /api/data 返回车型库+产品+VDS 真值表（来自 SQLite）
#      ④ /api/products、/api/vehicles 接收 web 端增删改并写回 SQLite
import http.server, socketserver, urllib.request, json, os, sys, secrets, re
from urllib.parse import urlparse, parse_qs, quote

PORT = 3001
HERE = os.path.dirname(os.path.abspath(__file__))
VPIC = 'https://vpic.nhtsa.dot.gov/api/vehicles/DecodeVinValues/'

# 数据层（SQLite）：与 server_proxy.py 同目录的 dbstore.py
sys.path.insert(0, HERE)
import dbstore  # noqa: E402

# 登录会话：token -> {'username','role'}（内存态，重启即失效；密码存 SQLite）
SESSIONS = {}
COOKIE_NAME = "ai_token"
SESSION_MAX_AGE = 7 * 24 * 3600  # 7 天


class Handler(http.server.SimpleHTTPRequestHandler):
    # 单个连接超时为 30s，防止半开/卡住的连接长期占用线程
    timeout = 30

    def _send_json(self, obj, code=200):
        body = json.dumps(obj, ensure_ascii=False).encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, DELETE, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.send_header('Cache-Control', 'no-store')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _read_json(self):
        try:
            length = int(self.headers.get('Content-Length') or 0)
        except ValueError:
            length = 0
        raw = self.rfile.read(length) if length else b''
        if not raw:
            return {}
        try:
            return json.loads(raw.decode('utf-8'))
        except Exception:
            return {}

    def _session_user(self):
        cookie = self.headers.get('Cookie') or ''
        m = re.search(r'%s=([^;]+)' % re.escape(COOKIE_NAME), cookie)
        if not m:
            return None
        return SESSIONS.get(m.group(1))

    def _require_session(self):
        u = self._session_user()
        if not u:
            self._send_json({'ok': False, 'error': '未登录或登录已过期'}, 401)
            return None
        return u

    def _log(self, action, detail=''):
        s = self._session_user()
        username = s['username'] if s else ''
        dbstore.add_log(username, action, detail)

    def _require_admin(self):
        u = self._require_session()
        if u is None:
            return None
        if u.get('role') != 'admin':
            self._send_json({'ok': False, 'error': '需要管理员权限'}, 403)
            return None
        return u

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, DELETE, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def do_GET(self):
        # 级联筛选选项：尝试获取登录用户，按地区组权限过滤 region 选项
        if self.path.startswith('/api/filter-options'):
            try:
                q = parse_qs(urlparse(self.path).query)
                filters = {k: v[0] for k, v in q.items()}
                options = dbstore.get_filter_options(filters)
                # 地区组权限：CDC 只看澳新，CDI 排除澳新
                user = self._session_user()
                if user:
                    rg = (user.get('region_group') or '').strip()
                    if rg == 'CDC':
                        options['region'] = [r for r in options.get('region', []) if r == '澳新地区']
                    elif rg == 'CDI':
                        options['region'] = [r for r in options.get('region', []) if r != '澳新地区']
                return self._send_json({'ok': True, 'options': options})
            except Exception as e:
                return self._send_json({'ok': False, 'error': str(e)}, 500)
        
        # 除静态页面外，所有 /api/* 需登录
        if self.path.startswith('/api/'):
            if not self._require_session():
                return
        if self.path.startswith('/api/data'):
            # 首屏优化：/api/data?products=0 不携带 PRODUCTS（~1.5MB），产品型号走 /api/products 懒加载
            try:
                q = parse_qs(urlparse(self.path).query)
                include_products = (q.get('products', ['1'])[0] or '1') != '0'
                return self._send_json(dbstore.get_all_data(include_products=include_products))
            except Exception as e:
                return self._send_json({'ok': False, 'error': str(e)}, 500)
        if self.path.startswith('/api/products'):
            # 仅产品型号数据（懒加载）：/api/products → {ok, products:{vid:[...]}}
            try:
                return self._send_json({'ok': True, 'products': dbstore.get_all_products()})
            except Exception as e:
                return self._send_json({'ok': False, 'error': str(e)}, 500)
        if self.path.startswith('/api/product-vehicles-by-code'):
            # 按 OE+CDAP 聚合关联车型：/api/product-vehicles-by-code?oe=...&cdap=...
            try:
                q = parse_qs(urlparse(self.path).query)
                oe = (q.get('oe', [''])[0] or '').strip()
                cdap = (q.get('cdap', [''])[0] or '').strip()
                return self._send_json({'ok': True, 'vehicles': dbstore.get_product_vehicles_by_code(oe, cdap)})
            except Exception as e:
                return self._send_json({'ok': False, 'error': str(e)}, 500)
        if self.path.startswith('/api/product-vehicles'):
            # 取某产品关联的所有车型：/api/product-vehicles?id=<product_id> → {ok, vehicles:[{vid,owner}]}
            try:
                q = parse_qs(urlparse(self.path).query)
                pid = int(q.get('id', ['0'])[0] or 0)
                return self._send_json({'ok': True, 'vehicles': dbstore.get_product_vehicles(pid)})
            except Exception as e:
                return self._send_json({'ok': False, 'error': str(e)}, 500)
        if self.path.startswith('/api/product-vehicle-remarks'):
            # 取某车型下所有产品备注：/api/product-vehicle-remarks?vid=<vid> → {ok, remarks:[{product_id, remark}]}
            try:
                q = parse_qs(urlparse(self.path).query)
                vid = (q.get('vid', [''])[0] or '').strip()
                return self._send_json({'ok': True, 'remarks': dbstore.get_vehicle_product_remarks(vid)})
            except Exception as e:
                return self._send_json({'ok': False, 'error': str(e)}, 500)
        if self.path.startswith('/api/regions'):
            # 地区列表：/api/regions → {ok, regions:['澳新地区', ...]}
            try:
                regions = dbstore.list_regions()
                return self._send_json({'ok': True, 'regions': regions})
            except Exception as e:
                return self._send_json({'ok': False, 'error': str(e)}, 500)
        if self.path.startswith('/api/vehicles'):
            # 分页拉取车型：/api/vehicles?limit=200&offset=0&q=&brand=&region=
            try:
                q = parse_qs(urlparse(self.path).query)
                limit = int(q.get('limit', ['200'])[0] or 200)
                offset = int(q.get('offset', ['0'])[0] or 0)
                kw = (q.get('q', [''])[0] or '').strip()
                br = (q.get('brand', [''])[0] or '').strip()
                rg = (q.get('region', [''])[0] or '').strip()
                rows, total = dbstore.get_vehicles_page(limit=limit, offset=offset, q=kw, brand=br, region=rg)
                return self._send_json({
                    'ok': True,
                    'total': total,
                    'vehicles': rows,
                    'hasMore': offset + len(rows) < total,
                    'limit': limit,
                    'offset': offset,
                })
            except Exception as e:
                return self._send_json({'ok': False, 'error': str(e)}, 500)
        if self.path.startswith('/api/filter-options'):
            # 级联筛选：根据已选条件返回各字段的可选值
            try:
                q = parse_qs(urlparse(self.path).query)
                filters = {k: v[0] for k, v in q.items()}
                options = dbstore.get_filter_options(filters)
                return self._send_json({'ok': True, 'options': options})
            except Exception as e:
                return self._send_json({'ok': False, 'error': str(e)}, 500)
        if self.path.startswith('/api/categories'):
            # 品类列表：/api/categories → {ok, categories:[{code,cn,en,count,products:[{cn,en,n}]}], unclassified}
            try:
                cats, un = dbstore.list_categories()
                return self._send_json({'ok': True, 'categories': cats, 'unclassified': un})
            except Exception as e:
                return self._send_json({'ok': False, 'error': str(e)}, 500)
        if self.path.split('?')[0] == '/api/catalog':
            # 产品目录全量（供「添加产品」下拉）：/api/catalog → {ok, items:[{grp,cn,en}]}
            try:
                return self._send_json({'ok': True, 'items': dbstore.get_all_catalog()})
            except Exception as e:
                return self._send_json({'ok': False, 'error': str(e)}, 500)
        if self.path.startswith('/api/vpic'):
            q = parse_qs(urlparse(self.path).query)
            vin = (q.get('vin', [''])[0] or '').strip().upper()
            if len(vin) != 17:
                return self._send_json({'ok': False, 'error': 'VIN 必须为 17 位'}, 400)
            try:
                req = urllib.request.Request(VPIC + vin + '?format=json',
                                             headers={'User-Agent': 'Mozilla/5.0'})
                with urllib.request.urlopen(req, timeout=12) as r:
                    data = json.loads(r.read().decode('utf-8'))
                res = (data.get('Results') or [{}])[0]
                g = lambda *ks: next((str(res[k]) for k in ks if str(res.get(k) or '').strip()), '')
                note = '' if any(str(v).strip() for v in res.values()) else \
                    'NHTSA 无此 VIN 数据（多为中国/澳洲等未在美注册车型）'
                out = {'ok': True,
                       'make': g('Make'), 'model': g('Model'),
                       'modelYear': g('Model Year'), 'trim': g('Trim'),
                       'series': g('Series'), 'engineModel': g('Engine Model'),
                       'engineCylinders': g('Engine Cylinders'),
                       'displacementL': g('Displacement (L)'),
                       'fuelType': g('Fuel Type - Primary'),
                       'bodyClass': g('Body Class'), 'driveType': g('Drive Type'),
                       'plantCountry': g('Plant Country'),
                       'plantCompany': g('Plant Company Name'),
                       'vehicleType': g('Vehicle Type'), 'note': note}
                return self._send_json(out)
            except Exception as e:
                return self._send_json({'ok': False, 'error': str(e)}, 502)
        # ---- 鉴权相关（需登录） ----
        if self.path.startswith('/api/me'):
            s = self._session_user()
            if not s:
                return self._send_json({'ok': False, 'error': '未登录或登录已过期'}, 401)
            # 从 DB 拉最新用户信息（角色/管理品类可能已被管理员修改，会话快照不实时）
            u = dbstore.get_user_by_login(s.get('username') or '')
            if not u or not u.get('active'):
                return self._send_json({'ok': False, 'error': '账号不存在或已停用'}, 401)
            u.pop('password', None)
            return self._send_json({'ok': True, 'user': u})
        if self.path.startswith('/api/users'):
            if not self._require_admin():
                return
            return self._send_json({'ok': True, 'users': dbstore.list_users()})
        if self.path.startswith('/api/feedbacks'):
            # 反馈列表：管理员看全部，普通用户只看自己的
            u = self._session_user()
            is_admin = u.get('role') == 'admin'
            try:
                items = dbstore.list_feedbacks(user_id=u['id'], is_admin=is_admin)
                return self._send_json({'ok': True, 'feedbacks': items})
            except Exception as e:
                return self._send_json({'ok': False, 'error': str(e)}, 500)
        if self.path.startswith('/api/logs'):
            # 操作日志查询（仅管理员）
            s = self._require_session()
            if not s: return
            if s['role'] != 'admin':
                return self._send_json({'ok': False, 'error': '仅管理员可查看'}, 403)
            try:
                q = parse_qs(urlparse(self.path).query)
                limit = int(q.get('limit', ['100'])[0] or 100)
                offset = int(q.get('offset', ['0'])[0] or 0)
                return self._send_json({'ok': True, **dbstore.get_logs(limit, offset)})
            except Exception as e:
                return self._send_json({'ok': False, 'error': str(e)}, 500)
        if self.path.startswith('/api/user-guide'):
            # 使用说明：读取 USER_GUIDE.md 返回 markdown 文本
            try:
                guide_path = os.path.join(HERE, 'USER_GUIDE.md')
                with open(guide_path, 'r', encoding='utf-8') as f:
                    text = f.read()
                return self._send_json({'ok': True, 'content': text})
            except FileNotFoundError:
                return self._send_json({'ok': False, 'error': '文档不存在'}, 404)
            except Exception as e:
                return self._send_json({'ok': False, 'error': str(e)}, 500)
        return super().do_GET()

    def do_POST(self):
        d = self._read_json()
        # 除登录外，所有 /api/* 需登录
        if self.path.startswith('/api/') and not self.path.startswith('/api/login'):
            if not self._require_session():
                return
        if self.path.startswith('/api/categories'):
            # 新增品类 {code, cn, en, level, parent_code} / 更新品类 {action:'update', code, cn, en, parent_code}
            try:
                if d.get('action') == 'update':
                    self._log('品类管理', '更新品类 ' + str(d.get('code','')))
                    dbstore.update_category(
                        d.get('code'), d.get('cn', ''), d.get('en', ''),
                        parent_code=d.get('parent_code'))
                    return self._send_json({'ok': True})
                self._log('品类管理', '新增品类: 代码=' + str(d.get('code','')) + ' 名称=' + str(d.get('cn',''))); code = dbstore.add_category(
                    d.get('code'), d.get('cn', ''), d.get('en', ''),
                    level=d.get('level', 1),
                    parent_code=d.get('parent_code', ''))
                return self._send_json({'ok': True, 'code': code})
            except ValueError as e:
                return self._send_json({'ok': False, 'error': str(e)}, 400)
            except Exception as e:
                return self._send_json({'ok': False, 'error': str(e)}, 500)
        if self.path.startswith('/api/products'):
            self._log('产品管理', '新增/编辑产品: VID=' + str(d.get('vid','')) + ' 产品=' + str(d.get('cn','')) + ' OE=' + str(d.get('oe','')) + ' CDAP=' + str(d.get('cdap','')))
            pid = dbstore.upsert_product(
                vid=str(d.get('vid', '')), pid=d.get('id'),
                grp=d.get('grp', ''), cn=d.get('cn', ''), en=d.get('en', ''),
                brand=d.get('brand', ''), cdap=d.get('cdap', ''), oe=d.get('oe', ''),
                notes=d.get('notes', ''), supplier=d.get('supplier', ''),
                cost=d.get('cost', ''), moq=d.get('moq', ''), loc=d.get('loc', ''),
                interchange=d.get('interchange', ''))
            return self._send_json({'ok': True, 'id': pid})
        if self.path.startswith('/api/product-vehicles'):
            # 关联/解除车型：{id, vid, action:'add'|'delete'}
            try:
                pid = int(d.get('id'))
                vid = str(d.get('vid', '')).strip()
                action = d.get('action','')
                self._log('产品管理', '关联车型' + ('添加' if action=='add' else '解除') + ': VID=' + str(d.get('vid','')) + ' 产品ID=' + str(d.get('id','')))
                action = (d.get('action') or 'add')
                if action == 'delete':
                    dbstore.delete_product_vehicle(pid, vid)
                else:
                    dbstore.add_product_vehicle(pid, vid)
                return self._send_json({'ok': True, 'vehicles': dbstore.get_product_vehicles(pid)})
            except Exception as e:
                return self._send_json({'ok': False, 'error': str(e)}, 400)
        if self.path.startswith('/api/product-vehicle-remark'):
            # 设置某产品针对某车型的备注：{pid, vid, remark}
            try:
                pid = int(d.get('pid'))
                vid = str(d.get('vid', '')).strip()
                remark = d.get('remark', '')
                dbstore.set_product_vehicle_remark(pid, vid, remark)
                return self._send_json({'ok': True})
            except Exception as e:
                return self._send_json({'ok': False, 'error': str(e)}, 400)
        if self.path.startswith('/api/feedbacks'):
            # 提交反馈 {content} 或管理员回复 {id, reply}
            u = self._session_user()
            if not u:
                return self._send_json({'ok': False, 'error': '未登录'}, 401)
            try:
                if 'reply' in d and 'id' in d:
                    # 管理员回复
                    if u.get('role') != 'admin':
                        return self._send_json({'ok': False, 'error': '需要管理员权限'}, 403)
                    ok = dbstore.reply_feedback(int(d['id']), d['reply'], u['username'])
                    if ok:
                        return self._send_json({'ok': True})
                    return self._send_json({'ok': False, 'error': '反馈不存在'}, 404)
                elif 'content' in d:
                    # 用户提交反馈
                    if not d['content'].strip():
                        return self._send_json({'ok': False, 'error': '内容不能为空'}, 400)
                    fid = dbstore.add_feedback(u['id'], u['username'], d['content'])
                    return self._send_json({'ok': True, 'id': fid})
                else:
                    return self._send_json({'ok': False, 'error': '参数缺失'}, 400)
            except Exception as e:
                return self._send_json({'ok': False, 'error': str(e)}, 500)
        if self.path.startswith('/api/catalog-products'):
            # 在产品目录（某品类下）新增产品 {grp, cn, en}：仅写入目录，不自动补齐到车型
            try:
                grp = d.get('grp', '')
                cn = d.get('cn', '')
                en = d.get('en', '')
                dbstore.add_catalog_product(grp, cn, en)
                return self._send_json({'ok': True})
            except ValueError as e:
                return self._send_json({'ok': False, 'error': str(e)}, 400)
            except Exception as e:
                return self._send_json({'ok': False, 'error': str(e)}, 500)
        if self.path.startswith('/api/export-catalog'):
            # 导出品类+产品目录为 Excel
            try:
                data = dbstore.export_catalog_excel()
                date_str = __import__('datetime').date.today().isoformat()
                cn_name = f"品类产品目录_{date_str}.xlsx"
                ascii_name = f"Catalog_Export_{date_str}.xlsx"
                self.send_response(200)
                self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                self.send_header('Content-Disposition', f'attachment; filename="{ascii_name}"; filename*=UTF-8\'\'{quote(cn_name)}')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.send_header('Cache-Control', 'no-store')
                self.send_header('Content-Length', str(len(data)))
                self.end_headers()
                self.wfile.write(data)
                return
            except Exception as e:
                return self._send_json({'ok': False, 'error': f'导出失败: {e}'}, 500)
        if self.path.startswith('/api/vehicles'):
            # 批量导入：请求体为数组 [{...}, {...}]，逐条 add_vehicle
            if isinstance(d, list):
                self._log('车型库管理', '批量导入车型: 数量=' + str(len(d)))
                ids = []
                errors = []
                for idx, item in enumerate(d):
                    try:
                        r = dict(item)
                        if not r.get('vid') or not str(r.get('vid')).strip():
                            errors.append({'idx': idx, 'vid': r.get('vid'), 'error': '车辆ID(vid) 必填'})
                            continue
                        if not r.get('brand') or not str(r.get('brand')).strip():
                            errors.append({'idx': idx, 'vid': r.get('vid'), 'error': '品牌(brand) 必填'})
                            continue
                        if not r.get('model') or not str(r.get('model')).strip():
                            errors.append({'idx': idx, 'vid': r.get('vid'), 'error': '车型(model) 必填'})
                            continue
                        r['origin'] = r.get('origin') or 'import'  # 导入标记
                        nid = dbstore.add_vehicle(r)
                        ids.append(nid)
                        try:
                            dbstore.seed_default_products_for_vid(r.get('vid'))
                        except Exception:
                            pass
                    except Exception as e:
                        errors.append({'idx': idx, 'vid': (item or {}).get('vid'), 'error': str(e)})
                return self._send_json({'ok': True, 'count': len(ids), 'ids': ids, 'errors': errors,
                                       'skipped': len(errors)})
            # 带 id + vin17 → 登记/更新该车型的实际 17 位 VIN 码（支持多条，全量覆盖）
            if d.get('id') is not None and 'vin17' in d:
                ok, res = dbstore.update_vehicle_vins(d.get('id'), d.get('vin17'))
                if ok:
                    vins_str = ', '.join(res) if res else '(清空)'
                    self._log('VIN码登记', f'VID={d.get("id")} 登记VIN: {vins_str}')
                return self._send_json({'ok': True, 'vins': res} if ok
                                       else {'ok': False, 'error': res}, 200 if ok else 400)
            rec = dict(d)
            rec['origin'] = rec.get('origin') or 'web'  # web 新增标记，便于整体清除
            nid = dbstore.add_vehicle(rec)
            # 新增车型自动带上默认产品模板（OE 等留空，便于后续补录）——每个车型都"必须有产品"
            try:
                n_prods = dbstore.seed_default_products_for_vid(rec.get('vid'))
            except Exception:
                n_prods = 0
            return self._send_json({'ok': True, 'id': nid, 'seeded_products': n_prods})
        # ---- 登录 / 登出 / 用户管理（管理员） ----
        if self.path.startswith('/api/login'):
            login = (d.get('login') or '').strip()
            pw = d.get('password') or ''
            u = dbstore.verify_user(login, pw)
            if not u:
                return self._send_json({'ok': False, 'error': '用户名/邮箱或密码错误，或账号已停用'}, 401)
            token = secrets.token_hex(16)
            SESSIONS[token] = {'id': u['id'], 'username': u['username'], 'role': u['role'], 'region_group': u.get('region_group', '')}
            self._log('登录', '用户 ' + u['username'] + ' 登录成功')
            body = json.dumps({'ok': True, 'user': u}, ensure_ascii=False).encode('utf-8')
            self.send_response(200)
            self.send_header('Set-Cookie', '%s=%s; HttpOnly; Path=/; SameSite=Lax; Max-Age=%d'
                             % (COOKIE_NAME, token, SESSION_MAX_AGE))
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.send_header('Cache-Control', 'no-store')
            self.send_header('Content-Length', str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        if self.path.startswith('/api/logout'):
            auth = self.headers.get('Cookie') or ''
            m = re.search(r'%s=([^;]+)' % re.escape(COOKIE_NAME), auth)
            if m:
                SESSIONS.pop(m.group(1), None)
            body = json.dumps({'ok': True}, ensure_ascii=False).encode('utf-8')
            self.send_response(200)
            self.send_header('Set-Cookie', '%s=; HttpOnly; Path=/; Max-Age=0' % COOKIE_NAME)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.send_header('Cache-Control', 'no-store')
            self.send_header('Content-Length', str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        if self.path.startswith('/api/users'):
            if not self._require_admin():
                return
            try:
                self._log('用户管理', '新增用户: 用户名=' + str(d.get('username','')) + ' 角色=' + str(d.get('role','')))
                new_id = dbstore.add_user(d.get('username'), d.get('email'),
                                          d.get('password'), d.get('role'), cats=d.get('cats'),
                                          region_group=d.get('region_group'))
                return self._send_json({'ok': True, 'id': new_id})
            except ValueError as e:
                return self._send_json({'ok': False, 'error': str(e)}, 400)
            except Exception as e:
                return self._send_json({'ok': False, 'error': str(e)}, 500)
        if self.path.startswith('/api/export-vehicles'):
            # 导出当前查询结果为 Excel：POST body = {"vids":[...]}（优先）或兼容旧 {"filters":[[field,value],...]}
            try:
                vids = d.get('vids') if isinstance(d, dict) else None
                filters = d.get('filters') if isinstance(d, dict) else (d if isinstance(d, list) else None)
                data = dbstore.export_vehicles_excel(vids=vids, filters=filters)
                if not data:
                    return self._send_json({'ok': False, 'error': '无匹配数据，请调整筛选条件'}, 400)
                filename = f"AutoIndex_Export_{__import__('datetime').date.today().isoformat()}.xlsx"
                self.send_response(200)
                self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.send_header('Cache-Control', 'no-store')
                self.send_header('Content-Length', str(len(data)))
                self.end_headers()
                self.wfile.write(data)
                return
            except Exception as e:
                return self._send_json({'ok': False, 'error': f'导出失败: {e}'}, 500)
        return self._send_json({'ok': False, 'error': 'unknown endpoint'}, 404)

    def do_PUT(self):
        d = self._read_json()
        if self.path.startswith('/api/'):
            if not self._require_session():
                return
        if self.path.startswith('/api/user-guide'):
            # 使用说明：保存 markdown 内容（仅管理员）
            s = self._require_session()
            if not s:
                return
            if s['role'] != 'admin':
                return self._send_json({'ok': False, 'error': '仅管理员可编辑'}, 403)
            try:
                guide_path = os.path.join(HERE, 'USER_GUIDE.md')
                with open(guide_path, 'w', encoding='utf-8') as f:
                    self._log('使用说明', '编辑使用说明')
                    f.write(d.get('content', ''))
                return self._send_json({'ok': True})
            except Exception as e:
                return self._send_json({'ok': False, 'error': str(e)}, 500)
        if self.path.startswith('/api/me'):
            s = self._require_session()
            if not s:
                return
            try:
                uid = s.get('id')
                if not uid:
                    u = dbstore.get_user_by_login(s.get('username') or '')
                    if not u:
                        return self._send_json({'ok': False, 'error': '未找到当前用户'}, 401)
                    uid = u['id']
                updated = dbstore.update_me(int(uid), email=d.get('email'),
                                            password=d.get('password') if d.get('password') else None,
                                            current_password=d.get('current_password'))
                return self._send_json({'ok': True, 'user': updated})
            except ValueError as e:
                return self._send_json({'ok': False, 'error': str(e)}, 400)
            except Exception as e:
                return self._send_json({'ok': False, 'error': str(e)}, 500)
        if self.path.startswith('/api/users'):
            if not self._require_admin():
                return
            try:
                updated = dbstore.update_user(int(d.get('id')), role=d.get('role'),
                                             password=d.get('password'), active=d.get('active'),
                                             email=d.get('email'), cats=d.get('cats'),
                                             region_group=d.get('region_group'))
                return self._send_json({'ok': True, 'user': updated})
            except ValueError as e:
                return self._send_json({'ok': False, 'error': str(e)}, 400)
            except Exception as e:
                return self._send_json({'ok': False, 'error': str(e)}, 500)
        if self.path.startswith('/api/catalog-products'):
            # 更新产品目录中的产品名 {grp, old_cn, cn, en}
            try:
                grp = d.get('grp', '')
                old_cn = d.get('old_cn', '')
                cn = d.get('cn', '')
                en = d.get('en', '')
                result = dbstore.update_catalog_product(grp, old_cn, cn, en)
                return self._send_json({'ok': True, **result})
            except ValueError as e:
                return self._send_json({'ok': False, 'error': str(e)}, 400)
            except Exception as e:
                return self._send_json({'ok': False, 'error': str(e)}, 500)
        return self._send_json({'ok': False, 'error': 'unknown endpoint'}, 404)

    def do_DELETE(self):
        d = self._read_json()
        if self.path.startswith('/api/') and not self.path.startswith('/api/login'):
            if not self._require_session():
                return
        if self.path.startswith('/api/categories'):
            # 删除品类 {code}（品类下有产品记录时后端会拒绝）
            try:
                self._log('品类管理', '删除品类: 代码=' + str(d.get('code')))
                dbstore.delete_category(d.get('code'))
                return self._send_json({'ok': True})
            except ValueError as e:
                return self._send_json({'ok': False, 'error': str(e)}, 400)
            except Exception as e:
                return self._send_json({'ok': False, 'error': str(e)}, 500)
        if self.path.startswith('/api/catalog-products'):
            # 删除产品目录中的产品 {grp, cn}
            try:
                grp = d.get('grp', '')
                cn = d.get('cn', '')
                result = dbstore.delete_catalog_product(grp, cn)
                return self._send_json({'ok': True, **result})
            except ValueError as e:
                return self._send_json({'ok': False, 'error': str(e)}, 400)
            except Exception as e:
                return self._send_json({'ok': False, 'error': str(e)}, 500)
        if self.path.startswith('/api/products'):
            pid = d.get('id')
            self._log('产品管理', '删除产品: ID=' + str(pid))
            dbstore.delete_product(pid)
            return self._send_json({'ok': True})
        if self.path.startswith('/api/vehicles'):
            if d.get('web'):
                self._log('车型库管理', '删除车型: VID=' + str(d.get('id', d.get('web', ''))))
                n = dbstore.delete_vehicles_web()
                return self._send_json({'ok': True, 'deleted': n})
            if 'id' in d and d.get('id') is not None:
                n = dbstore.delete_vehicle(d.get('id'))
                return self._send_json({'ok': True, 'deleted': n})
            return self._send_json({'ok': False, 'error': 'unsupported'}, 400)
        if self.path.startswith('/api/users'):
            if not self._require_admin():
                return
            try:
                dbstore.delete_user(int(d.get('id')))
                return self._send_json({'ok': True})
            except ValueError as e:
                return self._send_json({'ok': False, 'error': str(e)}, 400)
            except Exception as e:
                return self._send_json({'ok': False, 'error': str(e)}, 500)
        if self.path.startswith('/api/feedbacks'):
            # 删除反馈 {id}：管理员可删任何，用户只能删自己的
            u = self._session_user()
            is_admin = u.get('role') == 'admin'
            try:
                ok = dbstore.delete_feedback(int(d.get('id')), user_id=u['id'], is_admin=is_admin)
                if ok:
                    return self._send_json({'ok': True})
                return self._send_json({'ok': False, 'error': '反馈不存在或无权删除'}, 404)
            except Exception as e:
                return self._send_json({'ok': False, 'error': str(e)}, 500)
        return self._send_json({'ok': False, 'error': 'unknown endpoint'}, 404)

    def end_headers(self):
        # 多重 no-cache 头确保浏览器每次拉最新（避免车型库更新后页面停在旧版）
        self.send_header('Cache-Control', 'no-store, no-cache, must-revalidate, max-age=0')
        self.send_header('Pragma', 'no-cache')
        self.send_header('Expires', '0')
        super().end_headers()

    def log_message(self, *a):
        pass


if __name__ == '__main__':
    os.chdir(HERE)
    # 启动时确保表结构存在（不强制种子；种子请用 build_vin_tool.py 或 python dbstore.py）
    dbstore.init_db()
    dbstore.init_users()  # 建 users 表并种子 admin / admin123
    # 多线程：每个连接独立线程，单连接卡住不影响其他请求；守护线程避免残留
    class ThreadingHTTPServer(socketserver.ThreadingTCPServer):
        allow_reuse_address = True
        daemon_threads = True
    with ThreadingHTTPServer(('0.0.0.0', PORT), Handler) as httpd:
        print('VIN 工具服务(多线程)已启动：http://0.0.0.0:%d  (Ctrl+C 停止)' % PORT)
        httpd.serve_forever()
