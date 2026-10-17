# -*- coding: utf-8 -*-
"""
扫描 Excel 库，生成 lib-derived VDS_MAP 新条目（按 WMI 分组）。
输出供直接复制到 template.html。
"""
import openpyxl, re, json
from collections import defaultdict, OrderedDict

wb = openpyxl.load_workbook(r'D:\中国车\Oscar中国车主表.xlsx', data_only=True)
groups = defaultdict(list)
for ws in wb.worksheets:
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or len(row) < 12: continue
        vid, desc, make, model, submodel, year, capacity, kw, engine, chassis, series, vin = list(row)[:12]
        if not vin: continue
        prefix = str(vin).split('Orig')[0].strip().upper()
        prefix = ''.join(c for c in prefix if c.isalnum())
        if len(prefix) < 8: continue
        wmi = prefix[:3]
        vds = prefix[3:8]
        groups[(wmi, vds)].append({
            'brand': make, 'model': model, 'sub': submodel,
            'year': year, 'kw': kw, 'engine': engine
        })
wb.close()

# 当前 VDS_MAP 已收录的（避免重复）
existing = {
    'WH409','EB409','Z14U3','WP439','A36U9','E2409','W54U9','WJ4U6','A24W3','W9439','W74C3','W74U3','W7409','WS439','W94U9','WT409','WR409',
    'CE4CB','CE4CC','CF6CD','C74C4','C76C4','CD6C4','D74C4','CH4CD',
    'EF4A5','EF6C5','FG7A6','EE4A5','EEUA5','EF7A6','EF5A5','EF5A6','EFUA6','EFUA5','FFFA5','FF7A5','FFSA5','FF9A6','FGUA7','CBF19','CBD18','CA319','CA419','DB417','CBE17','CBE19','DBE19','EF6A5','ECMA4','EDMA4','FGVA7'
}

brand_section = {
    'LSJ':'MG (WMI LSJ)', 'LC0':'BYD商用车 (WMI LC0)', 'LGX':'BYD (WMI LGX)', 'LGW':'长城 GWM (WMI LGW)',
    'LNN':'奇瑞新能源 (WMI LNN)', 'LVT':'奇瑞 (WMI LVT)', 'LVV':'奇瑞 (WMI LVV)',
    'LSF':'上汽大通 MAXUS (WMI LSF)', 'LSH':'上汽大通 MAXUS (WMI LSH)', 'LSK':'上汽大通 MAXUS (WMI LSK)',
    'LB3':'吉利 GEELY (WMI LB3)', 'L6T':'吉利新 (WMI L6T)', 'LPE':'BYD (WMI LPE)',
    'LJ1':'江淮 JAC (WMI LJ1)', 'SAR':'MG UK (WMI SAR)', 'SAX':'MG UK (WMI SAX)',
}

by_wmi = OrderedDict()
for (wmi, vds), recs in sorted(groups.items(), key=lambda x: (x[0][0], x[0][1])):
    if vds in existing: continue
    by_wmi.setdefault(wmi, []).append((vds, recs))

for wmi, items in by_wmi.items():
    label = brand_section.get(wmi, wmi)
    print(f'// ===== {label} =====')
    for vds, recs in items:
        names = sorted({(str(r['brand'] or '').strip() + ' ' + str(r['model'] or '').strip()).strip() for r in recs})
        subs  = sorted({str(r['sub'] or '').strip() for r in recs if str(r['sub'] or '').strip() not in ('','.')})
        if len(names) > 1:
            name = ' / '.join(names)
        elif subs and len(subs) <= 2:
            name = names[0] + ' (' + '/'.join(subs) + ')'
        elif len(subs) > 2:
            name = names[0] + ' (' + str(len(subs)) + '种配置)'
        else:
            name = names[0]
        print('  "' + vds + '":"' + name + '",')
    print()