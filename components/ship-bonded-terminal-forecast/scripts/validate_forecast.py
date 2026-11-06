#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""综保区码头船舶作业预报单输出验证"""
import re
import sys
import zipfile

import openpyxl

# 必须有值（来自概况A表）
REQUIRED = {
    'J6': '申报日期',
    'B7': '标题',
    'C9': '船舶英文名',
    'D9': '国籍',
    'E9': '船代',
    'F9': '货代',
    'H9': '船舶长度',
    'I9': '载重吨',
    'C15': 'IMO号码',
    'F15': '呼号',
    'J15': '建造年月',
    'C16': '船长',
    'F16': '船宽',
    'J16': '型深',
    'C17': '总吨',
    'F17': '净吨',
    'J17': '载重吨',
}

# 必须为空（用户指定留空）
MUST_BLANK = {
    'B9': '船舶名称', 'G9': '货主', 'J9': '抵港吃水', 'K9': '货种',
    'B11': '总载货量', 'C11': '上/本/下港', 'D11': '计划减载量',
    'E11': '预抵时间', 'F11': '作业舱口', 'G11': '航次号',
    'C18': '船舶TPC', 'C24': '备注',
    'C22': '舱口计划', 'D22': '舱口计划', 'E22': '舱口计划', 'F22': '舱口计划',
    'G22': '舱口计划', 'H22': '舱口计划', 'I22': '舱口计划', 'J22': '舱口计划',
    'K22': '吃水',
    'C23': '舱口计划', 'D23': '舱口计划', 'E23': '舱口计划', 'F23': '舱口计划',
    'G23': '舱口计划', 'H23': '舱口计划', 'I23': '舱口计划', 'J23': '舱口计划',
    'K23': '吃水',
}

# 单位/千位分隔符不应出现
UNIT_RE = re.compile(r'(?i)\b(MT|M\.?T\.?|TEU|M)\b')
COMMA_RE = re.compile(r'[0-9],[0-9]')

# 允许保留默认的单元格
DEFAULT_OK = {'H11': '否', 'I11': '无限制'}


def main(path):
    ok = True
    print(f'== 验证 {path} ==')

    try:
        with zipfile.ZipFile(path) as z:
            bad = z.testzip()
            if bad:
                print(f'  ❌ zip 损坏: {bad}')
                ok = False
            else:
                print('  ✅ zip 完整')
    except Exception as e:
        print(f'  ❌ 无法打开 zip: {e}')
        return 1

    try:
        wb = openpyxl.load_workbook(path)
        ws = wb['Sheet1']
    except Exception as e:
        print(f'  ❌ openpyxl 加载失败: {e}')
        return 1
    print('  ✅ openpyxl 可打开, sheet:', wb.sheetnames)

    if '综保区码头船舶作业预报单' not in str(ws['B1'].value):
        print(f'  ❌ B1 标题被破坏: {ws["B1"].value!r}')
        ok = False
    else:
        print('  ✅ 模板标题完整')

    # 申报日期格式 YYYYMMDD
    dv = str(ws['J6'].value or '')
    if not re.fullmatch(r'\d{8}', dv):
        print(f'  ❌ 申报日期非 YYYYMMDD: {dv!r}')
        ok = False
    else:
        print(f'  ✅ 申报日期 {dv}')

    missing = [f'{c}({n})' for c, n in REQUIRED.items()
               if not str(ws[c].value or '').strip()]
    if missing:
        print('  ❌ 未填: ' + ', '.join(missing))
        ok = False
    else:
        print('  ✅ 必填字段齐全')

    notblank = []
    for c, n in MUST_BLANK.items():
        v = ws[c].value
        if v is not None and str(v).strip() != '':
            notblank.append(f'{c}({n})={v!r}')
    if notblank:
        print('  ❌ 应为空但有值: ' + ', '.join(notblank))
        ok = False
    else:
        print('  ✅ 指定留空字段均为空')

    # 单位 / 千位分隔符
    issues = []
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if not isinstance(v, str) or not v.strip():
                continue
            if UNIT_RE.search(v):
                issues.append(f'{cell.coordinate} 含单位: {v!r}')
            if COMMA_RE.search(v):
                issues.append(f'{cell.coordinate} 含千位分隔符: {v!r}')
    if issues:
        print('  ❌ ' + '; '.join(issues))
        ok = False
    else:
        print('  ✅ 无单位残留、无千位分隔符')

    print('  --- 关键值 ---')
    for coord in ['B7', 'B9', 'C9', 'D9', 'E9', 'F9', 'G9', 'H9', 'I9', 'J9', 'K9',
                  'B11', 'C11', 'D11', 'E11', 'F11', 'G11', 'H11', 'I11',
                  'C15', 'F15', 'J15', 'C16', 'F16', 'J16', 'C17', 'F17', 'J17', 'C18',
                  'K22', 'K23', 'C24']:
        v = ws[coord].value
        if v is not None and str(v).strip():
            print(f'    {coord} = {v}')

    print('  ✅ 全部通过' if ok else '  ❌ 存在问题')
    return 0 if ok else 1


if __name__ == '__main__':
    sys.exit(main(sys.argv[1]))
