#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
概况A表 (Ship's Particulars Form A) 解析器
支持: .doc (antiword+textutil), .docx (python-docx), .xls (xlrd), .xlsx (openpyxl), .pdf (pymupdf)

A 表两种版式:
  1) 表格版 (doc/xls/xlsx): label 和 value 在相邻/上下 cell
  2) 自由文本版 (PDF): "Ship's Name\\nWHALE SHARK" 竖排

输出: dict (ship_name/imo/callsign/flag/built/loa/breadth/depth/gt/nt/dwt/...)
"""
import re
import subprocess
import sys
import os
import json

# ---------------- 文本提取 ----------------

def _run(cmd):
    try:
        p = subprocess.run(cmd, capture_output=True, timeout=60)
        return p.stdout.decode('utf-8', 'ignore')
    except Exception:
        return ''


def _resolve(path):
    """路径可能含 U+00A0 等不可见字符 → glob 匹配真实文件"""
    if os.path.exists(path):
        return path
    d = os.path.dirname(path) or '.'
    b = os.path.basename(path)
    if not os.path.isdir(d):
        return path
    import glob as _g
    pat = ''.join('*' if (ord(c) > 127 or c in "'\"() ") else c for c in b)
    for c in _g.glob(os.path.join(d, pat)):
        if os.path.isfile(c):
            return c
    return path


def extract(path):
    """-> (rows: list[list[str]], flat: str)"""
    path = _resolve(path)
    ext = os.path.splitext(path)[1].lower()
    rows = []
    if ext == '.doc':
        # antiword 给表格结构但常截断；textutil 给完整字段（单元格用 \x07 分隔）→ 两者合并
        tmp = None
        target = path
        base = os.path.basename(path)
        if any(ch in base for ch in "'\"() ") or any(ord(c) > 127 for c in base):
            import tempfile
            tmp = tempfile.NamedTemporaryFile(suffix='.doc', delete=False)
            tmp.write(open(path, 'rb').read())
            tmp.close()
            target = tmp.name
        t_aw = _run(['antiword', target])
        t_tu = _run(['textutil', '-convert', 'txt', '-stdout', target])
        if tmp:
            try:
                os.unlink(tmp.name)
            except OSError:
                pass
        for line in t_aw.split('\n'):
            rows.append([c.strip() for c in line.split('|')])
        flat = t_aw
        if t_tu.strip():
            t_tu_n = t_tu.replace('\x07', '\n')
            flat = t_aw + '\n' + t_tu_n
            for line in t_tu_n.split('\n'):
                rows.append([line.strip()])
        return rows, flat
    if ext == '.xls':
        import xlrd
        wb = xlrd.open_workbook(path)
        for s in wb.sheet_names():
            ws = wb.sheet_by_name(s)
            for r in range(ws.nrows):
                rows.append([str(ws.cell_value(r, c)).strip() for c in range(ws.ncols)])
        return rows, '\n'.join(' | '.join(r) for r in rows)
    if ext == '.docx':
        from docx import Document
        doc = Document(path)
        for table in doc.tables:
            for row in table.rows:
                rows.append([cell.text.strip() for cell in row.cells])
        flat = '\n'.join(paragraph.text for paragraph in doc.paragraphs)
        flat += '\n' + '\n'.join(' | '.join(row) for row in rows)
        return rows, flat
    if ext in ('.xlsx', '.xlsm'):
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True)
        for s in wb.sheetnames:
            ws = wb[s]
            for r in range(1, ws.max_row + 1):
                rows.append([str(ws.cell(r, c).value or '').strip()
                             for c in range(1, ws.max_column + 1)])
        return rows, '\n'.join(' | '.join(r) for r in rows)
    if ext == '.pdf':
        import fitz
        doc = fitz.open(path)
        t = '\n'.join(p.get_text() for p in doc)
        for line in t.split('\n'):
            rows.append([line.strip()])
        return rows, t
    t = open(path, errors='ignore').read()
    for line in t.split('\n'):
        rows.append([line.strip()])
    return rows, t


# ---------------- label 定义 ----------------

LABELS = {
    'ship_name': (['ship s name', "ship's name", 'ship name', 'ship s', '船名'], 'upper'),
    'imo':       (['imo no', 'imo number', 'imo', 'imo no.', '船舶编号'], 'digits'),
    'callsign':  (['call sign', 'callsign', '呼号'], 'raw'),
    'flag':      (['flag', 'nationality', '船籍港', '船旗'], 'upper'),
    'built':     (['built year', 'built', '建造年月', '建造年份', 'date of build'], 'raw'),
    'shiptype':  (['ship-type', 'ship type', 'ship s-type', '船舶类型', 'type'], 'upper'),
    'loa':       (['loa', 'length overall', '船长'], 'num'),
    'breadth':   (['breadth', 'breadth molded', '船宽'], 'num'),
    'depth':     (['depth', 'depth molded', '型深'], 'num'),
    'gt':        (['gross tonnage', 'gross', '总吨', 'intern. grt', 'intern grt', 'grt'], 'num'),
    'nt':        (['net tonnage', 'net', '净吨', 'intern. nrt', 'intern nrt', 'nrt'], 'num'),
    'dwt':       (['dwt / teu', 'dwt/teu', 'dwt', '载重吨', 'deadweight', 'dw '], 'num'),
    'power':     (['m / power', 'm/power', '主机功率', 'power me'], 'upper'),
    'maxheight': (['max height', '最高高度', 'airdraft total', 'air draft'], 'num'),
    'regno':     (['initial reg', 'registry no', '登记号', 'register no', 'official no'], 'upper'),
    'owner':     (['owner', 'operator', 'owners', '船东', '公司'], 'upper'),
    'security':  (['security level', '保安等级'], 'raw'),
    'captain':   (['signature of captain', 'signature of master', '船长签名'], 'upper'),
}

BAD_VAL = re.compile(
    r'^(—|-|/|n/?a|none|null|if not applicable|no fill|待填|无|空|year\s*&?\s*month|'
    r'istry|istry no|istry port|level|security level|port|owner|operator|keel|allowance|'
    r'moulded|survey|issued|valid|exempted|content)$', re.I)

# label 被换行拆开时的碎片（非真值）
CONTINUATION = re.compile(
    r'^(name|port|istry|istry port|istry no|no\.?$|level|operator|tonnage|moulded|'
    r'from keel|keel|allowance|survey|issued|valid|last annual|exempted|content|'
    r'mean draft|summer|fresh water|year|month|&?\s*month|m / power|ship-type|'
    r'no\. of coc|valid until|not applicable|if not applicable)$', re.I)

# 纯 label 词被当值的兜底
LABEL_WORDS = ('FLAG', 'LOA', 'NAME', 'BUILT', 'YEAR', 'DWT', 'TEU', 'NET', 'GROSS',
                'BREADTH', 'DEPTH', 'CALL', 'SIGN', 'TYPE', 'SHIP', 'OPERATOR',
                'CONTENT', 'CERTIFICATE', 'SECURITY', 'POWER', 'HEIGHT', 'DRAFT')


def _clean(kind, v):
    if v is None:
        return None
    v = str(v).strip()
    if not v or BAD_VAL.match(v):
        return None
    if kind == 'digits':
        d = re.sub(r'\D', '', v)
        return d[:7] if len(d) >= 5 else None
    if kind == 'upper':
        v = re.sub(r'\s+', ' ', v).upper().strip(' .,-')
        v = re.sub(r'\s+(PORT|ISTRY|ISTRY PORT)$', '', v).strip()
        if v in ('', 'PORT', 'ISTRY') or CONTINUATION.match(v):
            return None
        if v in LABEL_WORDS:
            return None
        return v
    if kind == 'num':
        # 去单位 (M / MT / TEU) 与千位分隔符 → 纯数字
        s = re.sub(r'(?i)\b(MT|M\.?T\.?|TEU|M)\b', '', v)
        s = re.sub(r'(?i)\b(KG|KGS|TON|TONS)\b', '', s)
        s = s.replace(',', '').replace('，', '').replace(' ', '')
        m = re.search(r'-?[0-9]+(?:\.[0-9]+)?', s)
        if not m:
            return None
        f = float(m.group(0))
        # 整数化：避免 225.0
        if abs(f - round(f)) < 1e-9:
            return int(round(f))
        return f
    if kind == 'raw':
        v = re.sub(r'\s+', ' ', v).strip()
        if BAD_VAL.match(v) or CONTINUATION.match(v) or v.upper() in LABEL_WORDS:
            return None
        # 呼号等：字母数字混合，保留但转大写
        if re.fullmatch(r'[A-Za-z0-9\-/]{3,12}', v):
            return v.upper()
        return v
    return v


def _nl(s):
    s = str(s).lower().replace('’', "'").replace('‘', "'")
    s = re.sub(r'\s+', ' ', s).strip().rstrip(':：').strip()
    return s


def _match_label(s):
    lbl = _nl(s)
    if not lbl or len(lbl) > 40:
        return None
    if re.search(r'certificat|seafarer|competency|inspection|declar|seaworth|hereby|'
                 r'owner\s*/\s*operator', lbl):
        return None
    for key, (aliases, _k) in LABELS.items():
        for a in aliases:
            if lbl == _nl(a):
                return key
    # 短前缀（label 被截断: "Ship's", "Flag/Reg-"）
    for key, (aliases, _k) in LABELS.items():
        for a in aliases:
            al = _nl(a)
            if len(al) >= 4 and lbl.startswith(al[:4]):
                return key
    return None


# ---------------- 自由文本 fallback ----------------

FLAT_PATTERNS = [
    ('imo',       r'IMO\s*(?:NO\.?|NUMBER)?\s*[:：]?\s*([0-9]{7})'),
    ('callsign',  r'Call\s*Sign\s*[:：]?\s*([A-Z0-9]{3,8})'),
    ('built',     r'Built\s*(?:Year)?\s*[:：]?\s*(\d{4}(?:\.\d{1,2}){0,2})'),
    ('shiptype',  r'Ship[\-\s]*type\s*[:：]?\s*([A-Z][A-Z /\-]{2,40})'),
    ('loa',       r'\bLOA\b\s*[:：]?\s*([0-9]+(?:\.[0-9]+)?)'),
    ('breadth',   r'Breadth\s*(?:Moulded)?\s*[:：]?\s*([0-9]+(?:\.[0-9]+)?)'),
    ('depth',     r'Depth\s*(?:Moulded)?\s*[:：]?\s*([0-9]+(?:\.[0-9]+)?)'),
    ('gt',        r'Gross\s*(?:Tonnage)?\s*[:：]?\s*([0-9]+(?:\.[0-9]+)?)'),
    ('nt',        r'Net\s*Tonnage\s*[:：]?\s*([0-9]+(?:\.[0-9]+)?)'),
    ('dwt',       r'DWT\s*/?\s*TEU\s*[:：]?\s*([0-9]+(?:\.[0-9]+)?)'),
    ('maxheight', r'Max\s*Height\s*(?:From\s*Keel)?\s*[:：]?\s*([0-9]+(?:\.[0-9]+)?)'),
    ('power',     r'M\s*/\s*Power\s*[:：]?\s*([0-9]+\s*KW)'),
    ('regno',     r'Initial\s*Reg(?:istry)?\s*No\.?\s*[:：]?\s*([A-Za-z0-9\-]{4,30})'),
    ('owner',     r'Owner\s*/\s*Operator\s*[:：]?\s*([A-Z][A-Z0-9 ,.\-&]{3,60})'),
    ('security',  r'Security\s*Level\s*[:：]?\s*([0-9])'),
    ('captain',   r'Signature\s*of\s*(?:Captain|master)\s*[:：]?\s*([A-Z][A-Z ,.\-]{2,40})'),
    ('ship_name', r"Ship[\s'’\-]*s?\s*Name\s*[:：]?\s*([A-Z][A-Z0-9 ]{2,40})"),
]


def parse_particulars(path):
    rows, flat = extract(path)
    res = {}

    # 表格版：label/value 同排 或 label 在下方
    for idx, cells in enumerate(rows):
        for i, c in enumerate(cells):
            key = _match_label(c)
            if not key or key in res:
                continue
            found = False
            # 右侧（跳过 ':' 之类的分隔符）
            for j in range(i + 1, min(i + 5, len(cells))):
                cs = cells[j].strip()
                if not cs or re.fullmatch(r'[:：/\-–]+', cs):
                    continue
                if _match_label(cells[j]):
                    break
                v = _clean(LABELS[key][1], cs)
                if v:
                    res[key] = v
                    found = True
                    break
            if found:
                continue
            # 下方（label 碎片行需跳过）
            for nrow in range(idx + 1, min(idx + 5, len(rows))):
                nxt = rows[nrow]
                if not nxt or not any(nxt):
                    continue
                for j in (list(range(max(0, i - 1), min(i + 3, len(nxt)))) or [0]):
                    if j >= len(nxt):
                        continue
                    if _match_label(nxt[j]) or CONTINUATION.match(nxt[j].strip()):
                        continue
                    v = _clean(LABELS[key][1], nxt[j])
                    if v:
                        res[key] = v
                        found = True
                        break
                if found:
                    break

    # 自由文本 fallback
    for key, pat in FLAT_PATTERNS:
        if key in res:
            continue
        m = re.search(pat, flat, re.I | re.S)
        if m:
            v = _clean(LABELS[key][1], m.group(1))
            if v:
                res[key] = v

    # 竖排 "label\nvalue"
    lines = [l.strip() for l in flat.split('\n')]
    for i, ln in enumerate(lines):
        key = _match_label(ln)
        if not key or key in res:
            continue
        for j in range(i + 1, min(i + 4, len(lines))):
            nx = lines[j]
            if not nx or _match_label(nx) or CONTINUATION.match(nx):
                continue
            v = _clean(LABELS[key][1], nx)
            if v:
                res[key] = v
                break
        if key in res:
            continue
        m = re.match(r'^(.*?)\s{2,}([A-Z0-9][^\n]{1,40})$', ln)
        if m and _match_label(m.group(1)):
            v = _clean(LABELS[key][1], m.group(2))
            if v:
                res[key] = v

    # 船名跨格 (GREAT | VANGUARD / GREEN | PORTOCEL)
    if 'ship_name' in res:
        for idx, cells in enumerate(rows):
            p = None
            for j, c in enumerate(cells):
                cs = c.strip()
                if not cs or len(cs) > 20:
                    continue
                if 'particular' in cs.lower() or 'certificate' in cs.lower():
                    continue
                if (re.search(r"Ship[\s'’\-]*s?\b", cs, re.I)
                        and not re.search(r"ship[\s'’\-]*s?\s*type", cs, re.I)):
                    p = j
                    break
            if p is None:
                continue
            cands = []
            for j in (p + 1, p + 2):
                if j < len(cells):
                    cs = cells[j].strip()
                    if re.fullmatch(r"[A-Z][A-Z0-9]{1,}", cs) and cs != 'NAME':
                        cands.append(cs)
                        break
            if idx + 1 < len(rows):
                nxt = rows[idx + 1]
                for j in (p + 1, p + 2):
                    if j < len(nxt):
                        ns = nxt[j].strip()
                        if re.fullmatch(r"[A-Z][A-Z0-9]{1,}", ns) and ns != 'NAME':
                            cands.append(ns)
                            break
            if cands:
                res['ship_name'] = (' '.join(cands[:2])).upper()
                break

    # 旗国 / 船籍港
    if 'flag' in res and '/' in str(res['flag']):
        pr = [x.strip() for x in str(res['flag']).split('/')]
        res['flag'] = pr[0]
        if len(pr) > 1 and pr[1]:
            res['registry_port'] = pr[1]

    # --- 特殊版式补充 ---
    # 建造年: "Date Launched / Delivery : 20.07.2012 / 28.09.2012" → 2012
    if 'built' not in res:
        m = re.search(r'Date\s*(?:Launched|Delivery|Keel\s*Laid)[^\n]*?(\d{2})[./-](\d{2})[./-](\d{4})',
                      flat, re.I)
        if m:
            res['built'] = m.group(3)
    # 建造年不能是船厂名
    if res.get('built') and not re.search(r'\d{4}', str(res['built'])):
        res.pop('built', None)
        m = re.search(r'(?:Date\s*(?:Launched|Delivery)|Built)[^\n]{0,40}?(\d{4})', flat, re.I)
        if m:
            res['built'] = m.group(1)

    # 载重吨: "MARK | DRAFT | Freeboard | Displacement | DW [mt]" 表头 + Summer 行
    # 夏季载重吨在 DW 列（行内最后一个数字），前面是吃水/干舷/排水量
    if 'dwt' not in res or (isinstance(res.get('dwt'), (int, float)) and res['dwt'] < 1000):
        # 先确认表头含 DW 列
        has_dw = bool(re.search(r'DW\s*\[mt\]', flat, re.I))
        for cells in rows:
            if not cells:
                continue
            lbl = ' '.join(cells[:2]).strip().lower()
            if lbl.startswith('summer'):
                nums = []
                for c in cells:
                    for m in re.finditer(r'[0-9][0-9,]*(?:\.[0-9]+)?', c or ''):
                        v = _clean('num', m.group(0))
                        if v is not None:
                            nums.append(v)
                if nums:
                    res['dwt'] = nums[-1] if has_dw else max(nums)
                break

    res['_source'] = os.path.basename(path)
    res['_text'] = flat
    return res


if __name__ == '__main__':
    data = parse_particulars(sys.argv[1])
    data.pop('_text', None)
    print(json.dumps(data, ensure_ascii=False, indent=2))
