#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
腾讯会议考勤统计（零第三方依赖）
================================
读取腾讯会议导出的「成员参会概况」.xlsx，按花名册判定每位成员的
准时 / 迟到 / 缺勤 / 请假，输出 Markdown 日报；一次给多个文件时再出周报。

只用 Python 标准库（zipfile + xml）解析 xlsx，不需要 pandas / openpyxl。

用法：
    python3 attendance_stat.py 导出.xlsx --roster roster.md                 # 单日 → 日报
    python3 attendance_stat.py 周一.xlsx 周二.xlsx … --roster roster.md     # 多日 → 日报 + 周报
    python3 attendance_stat.py 导出.xlsx --leave "小鹿,阿布"                 # 当日请假名单
    python3 attendance_stat.py *.xlsx --leave-file leaves.csv               # 按日期的请假表（日期,名字）
    python3 attendance_stat.py 导出.xlsx --on-time 20:00 --absent-after 20:30  # 改判定时间
    python3 attendance_stat.py 导出.xlsx --csv --out ./out                   # 额外导出 CSV 明细

判定规则优先级：命令行参数 > 花名册「规则」表 > 默认值（准时线 06:00:00，缺勤线 07:00:00，
统计日周一至周五，每周 1 次请假不计）。
"""

import argparse
import csv
import datetime
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

M_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

WEEKDAY_CN = '一二三四五六日'

DEFAULT_RULES = {
    'title': '会议考勤',
    'on_time': datetime.time(6, 0, 0),       # 首次入会 ≤ 准时线 → 准时
    'absent_after': datetime.time(7, 0, 0),  # 首次入会 > 缺勤线 → 缺勤；None = 不设缺勤线
    'workdays': {0, 1, 2, 3, 4},             # 统计日，Python weekday()：周一=0 … 周日=6
    'free_leaves': 1,                        # 每周免罚请假次数，超出的按迟到计
    'member_identity': '成员',               # 花名册里身份为它的人才计入应到
}


# ---------------------------------------------------------------- 小工具
def parse_time(value):
    """'06:00' / '06:00:00' → time；'无' / 'none' / '' → None。"""
    s = str(value).strip()
    if s.lower() in ('', 'none', 'null', '无', '不设', '-'):
        return None
    m = re.fullmatch(r'(\d{1,2}):(\d{2})(?::(\d{2}))?', s)
    if not m:
        raise ValueError('时间格式应为 HH:MM 或 HH:MM:SS：%s' % value)
    h, mi, se = int(m.group(1)), int(m.group(2)), int(m.group(3) or 0)
    return datetime.time(h, mi, se)


def parse_workdays(value):
    """'1-5' / '1,3,5' / '周一至周五' / '周一、周三' → {weekday 下标}（周一=0）。"""
    cn = {'一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '日': 7, '天': 7}
    s = re.sub(r'周|星期|礼拜', '', str(value).strip())
    s = s.replace('至', '-').replace('到', '-').replace('~', '-').replace('～', '-')
    s = s.replace('，', ',').replace('、', ',').replace(' ', '')

    def one(tok):
        if tok in cn:
            return cn[tok]
        if tok.isdigit() and 1 <= int(tok) <= 7:
            return int(tok)
        raise ValueError('无法识别的统计日：%s' % value)

    days = set()
    for part in s.split(','):
        if not part:
            continue
        if '-' in part:
            a, b = part.split('-', 1)
            a, b = one(a), one(b)
            if a > b:
                raise ValueError('统计日区间写反了：%s' % value)
            days.update(range(a, b + 1))
        else:
            days.add(one(part))
    if not days:
        raise ValueError('统计日为空：%s' % value)
    return {d - 1 for d in days}


def workdays_text(workdays):
    return '、'.join('周' + WEEKDAY_CN[d] for d in sorted(workdays))


def next_second(t):
    dt = datetime.datetime.combine(datetime.date(2000, 1, 1), t) + datetime.timedelta(seconds=1)
    return dt.time()


def fmt_hms(t):
    return t.strftime('%H:%M:%S')


def fmt_hm(dt):
    return dt.strftime('%H:%M:%S') if dt else '--:--:--'


# ---------------------------------------------------------------- xlsx 读取
def read_xlsx_sheets(path):
    """读取 xlsx，返回 [(sheet 名, rows)]，rows = [{列字母: 值}]。纯标准库。"""
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        wb = ET.fromstring(z.read('xl/workbook.xml'))
        rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        relmap = {r.get('Id'): r.get('Target') for r in rels}
        sst = []
        if 'xl/sharedStrings.xml' in names:
            ss = ET.fromstring(z.read('xl/sharedStrings.xml'))
            for si in ss.findall('{%s}si' % M_NS):
                sst.append(''.join(t.text or '' for t in si.iter('{%s}t' % M_NS)))
        sheets = []
        for s in wb.find('{%s}sheets' % M_NS):
            name = s.get('name')
            rid = s.get('{%s}id' % R_NS)
            target = relmap.get(rid, '').lstrip('/')
            if not target.startswith('xl/'):
                target = 'xl/' + target
            if target not in names:
                continue
            sh = ET.fromstring(z.read(target))
            rows = []
            for row in sh.iter('{%s}row' % M_NS):
                cells = {}
                for c in row.findall('{%s}c' % M_NS):
                    ref = c.get('r') or ''
                    m = re.match(r'[A-Z]+', ref)
                    col = m.group() if m else ''
                    t = c.get('t')
                    v = c.find('{%s}v' % M_NS)
                    isv = c.find('{%s}is' % M_NS)
                    if t == 's' and v is not None:
                        val = sst[int(v.text)] if v.text and int(v.text) < len(sst) else ''
                    elif t == 'inlineStr' and isv is not None:
                        val = ''.join(x.text or '' for x in isv.iter('{%s}t' % M_NS))
                    elif v is not None:
                        val = v.text or ''
                    else:
                        val = ''
                    if col:
                        cells[col] = val
                rows.append(cells)
            sheets.append((name, rows))
    return sheets


def excel_serial_to_dt(serial):
    """Excel 日期序列号 → datetime（1900 日期系统）。"""
    try:
        f = float(serial)
    except ValueError:
        return None
    if f < 10000:   # 纯时间或过小的数，不是完整日期
        return None
    try:
        dt = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=f)
    except OverflowError:
        return None
    # 浮点误差会让 06:00:00 变成 06:00:00.00001，取整到秒，避免边界误判
    return (dt + datetime.timedelta(microseconds=500000)).replace(microsecond=0)


def parse_dt(s):
    """解析首次入会时间（字符串或 Excel 序列号）→ datetime，失败返回 None。"""
    if s is None:
        return None
    s = str(s).strip()
    if not s:
        return None
    if re.fullmatch(r'-?\d+(\.\d+)?', s):
        return excel_serial_to_dt(s)
    m = re.search(r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})[ T](\d{1,2}):(\d{2}):(\d{2})', s)
    if m:
        try:
            return datetime.datetime(*[int(x) for x in m.groups()])
        except ValueError:
            return None
    return None


# ---------------------------------------------------------------- 花名册
def normalize(s):
    """归一化显示名：去空白，全角括号转半角，英文不分大小写。"""
    s = str(s).replace('（', '(').replace('）', ')')
    return re.sub(r'\s+', '', s).casefold()


RULE_KEYS = {
    '名称': 'title', '标题': 'title', '活动名称': 'title',
    '准时线': 'on_time',
    '缺勤线': 'absent_after',
    '统计日': 'workdays',
    '每周免罚请假': 'free_leaves', '每周免罚请假次数': 'free_leaves',
    '计入应到的身份': 'member_identity',
}


def parse_roster(md_path):
    """
    解析花名册 Markdown（唯一数据源）。返回 dict：
      roster:            {规范名归一化: {'name', 'identity', 'note'}}
      aliases:           {显示名归一化: 规范名归一化}
      excluded:          {排除人名归一化: 原文}
      excluded_aliases:  {已退出人员显示名归一化: 'EXCL:' + 规范名}
      rules:             {规则键: 原文值}（来自「规则」表，可为空）
    """
    with open(md_path, encoding='utf-8') as f:
        text = f.read()
    roster, aliases, excluded, excluded_aliases, rules = {}, {}, {}, {}, {}
    section = ''
    for raw in text.splitlines():
        line = raw.strip()
        if line.startswith('## '):
            section = line[3:].strip()
            continue
        if not section:
            continue
        if line.startswith('|'):
            cells = [c.strip() for c in line.strip('|').split('|')]
            if not cells or all(c == '' for c in cells):
                continue
            if all(re.fullmatch(r':?-{2,}:?', c) for c in cells if c):
                continue
            if any(('规范名' in c) or ('显示名' in c) for c in cells) or cells[0] in ('项目', '规则'):
                continue
            if '在册' in section or '成员名单' in section:
                if cells[0]:
                    roster[normalize(cells[0])] = {
                        'name': cells[0],
                        'identity': cells[1] if len(cells) > 1 and cells[1] else '成员',
                        'note': cells[2] if len(cells) > 2 else '',
                    }
            elif '映射' in section:
                if len(cells) >= 2 and cells[0] and cells[1]:
                    canon = normalize(cells[1])
                    note = cells[2] if len(cells) > 2 else ''
                    if ('退出' in note) or ('请假' in note) or ('排除' in note):
                        target = 'EXCL:' + cells[1]
                        for alias in cells[0].split('/'):
                            a = normalize(alias)
                            if a:
                                excluded_aliases[a] = target
                        excluded[canon] = cells[1]
                        excluded_aliases.setdefault(canon, target)
                    else:
                        for alias in cells[0].split('/'):
                            a = normalize(alias)
                            if a:
                                aliases[a] = canon
            elif '规则' in section:
                if len(cells) >= 2 and cells[0] in RULE_KEYS:
                    rules[RULE_KEYS[cells[0]]] = cells[1]
        elif '排除' in section:
            m = re.match(r'^[-*]\s+([^（(]+)', line)
            if m:
                nm = m.group(1).strip()
                excluded[normalize(nm)] = nm
    # 规范名本身也是一个可用的显示名：会议昵称和规范名一致、或只多了括号后缀时，无需另行登记
    for key in roster:
        aliases.setdefault(key, key)
    return {'roster': roster, 'aliases': aliases, 'excluded': excluded,
            'excluded_aliases': excluded_aliases, 'rules': rules}


def coerce_md_rule(key, text):
    """
    花名册「规则」表的值允许写成说明文字（如「首次入会 ≤ 06:00:00」「周一至周五（周六不计）」），
    这里从文字里提取出能用的值；提取不到返回 None，调用方会保留默认值。
    """
    s = str(text).strip()
    if key == 'title':
        return s or None
    if key in ('on_time', 'absent_after'):
        m = re.search(r'\d{1,2}:\d{2}(?::\d{2})?', s)
        if m:
            return parse_time(m.group())
        if key == 'absent_after' and s.lower() in ('无', '不设', 'none', '-'):
            return 'NONE'
        return None
    if key == 'workdays':
        try:
            return parse_workdays(s)
        except ValueError:
            m = re.search(r'(?:周|星期)[一二三四五六日天](?:\s*[至到~～-]\s*(?:周|星期)?[一二三四五六日天])?'
                          r'(?:\s*[、,，]\s*(?:周|星期)?[一二三四五六日天])*', s)
            if m:
                try:
                    return parse_workdays(m.group())
                except ValueError:
                    return None
            return None
    if key == 'free_leaves':
        m = re.search(r'\d+', s)
        return int(m.group()) if m else None
    if key == 'member_identity':
        return s or None
    return None


def build_rules(md_rules, args=None):
    """合并规则：命令行 > 花名册「规则」表 > 默认值。"""
    rules = dict(DEFAULT_RULES)
    for key, text in (md_rules or {}).items():
        val = coerce_md_rule(key, text)
        if val is None:
            label = next((k for k, v in RULE_KEYS.items() if v == key), key)
            print('⚠️ 花名册规则「%s」读不出有效值（%s），使用默认值' % (label, text))
            continue
        rules[key] = None if val == 'NONE' else val
    if args is not None:
        for key in ('title', 'on_time', 'absent_after', 'workdays', 'free_leaves'):
            val = getattr(args, key, None)
            if val is not None:
                rules[key] = val
    # 统一类型
    if isinstance(rules['on_time'], str):
        rules['on_time'] = parse_time(rules['on_time'])
        if rules['on_time'] is None:
            raise ValueError('准时线不能为空')
    if isinstance(rules['absent_after'], str):
        rules['absent_after'] = parse_time(rules['absent_after'])
    if isinstance(rules['workdays'], str):
        rules['workdays'] = parse_workdays(rules['workdays'])
    rules['free_leaves'] = int(rules['free_leaves'])
    if rules['absent_after'] is not None and rules['absent_after'] < rules['on_time']:
        raise ValueError('缺勤线（%s）不能早于准时线（%s）' % (
            fmt_hms(rules['absent_after']), fmt_hms(rules['on_time'])))
    return rules


def match_name(display, aliases):
    """
    多层昵称匹配：① 整串精确 → ② 括号前的基础名 → ③ 最后一对括号里的名字（含一层嵌套）。
    返回规范名归一化 key；匹配不到返回 None。
    """
    norm = normalize(display)
    if norm in aliases:
        return aliases[norm]
    if '(' in norm:
        base = norm.split('(')[0]
        if base in aliases:
            return aliases[base]
        i, j = norm.rfind('('), norm.rfind(')')
        if 0 <= i < j:
            inner = norm[i + 1:j]
            if inner in aliases:
                return aliases[inner]
            if '(' in inner:
                i2, j2 = inner.rfind('('), inner.rfind(')')
                if 0 <= i2 < j2 and inner[i2 + 1:j2] in aliases:
                    return aliases[inner[i2 + 1:j2]]
    return None


# ---------------------------------------------------------------- 会议解析
def parse_meeting(path):
    """
    解析一份腾讯会议导出，返回：
      {'date', 'meeting_id', 'creator', 'total', 'participants': [(显示名, 首次入会 datetime)]}
    """
    sheets = read_xlsx_sheets(path)
    target_rows = None
    for name, rows in sheets:
        if '参会概况' in name:
            target_rows = rows
            break
    if target_rows is None:
        for name, rows in sheets:
            if any('昵称' in ''.join(map(str, r.values())) and '入会' in ''.join(map(str, r.values()))
                   for r in rows):
                target_rows = rows
                break
    if target_rows is None:
        raise ValueError('未找到「成员参会概况」表：%s' % os.path.basename(path))

    header_idx = None
    for i, row in enumerate(target_rows[:40]):
        joined = ''.join(str(v) for v in row.values())
        if '昵称' in joined and '入会' in joined and '时间' in joined:
            header_idx = i
            break
    if header_idx is None:
        raise ValueError('未找到表头行（需含「昵称」「首次入会时间」）：%s' % os.path.basename(path))

    meta = {}
    for row in target_rows[:header_idx]:
        items = list(row.items())
        if len(items) >= 2 and str(items[0][1]).strip().endswith((':', '：')):
            meta[str(items[0][1]).strip().rstrip('：:')] = str(items[1][1]).strip()

    hdr = {str(v).strip(): col for col, v in target_rows[header_idx].items()}
    nick_col = next((c for k, c in hdr.items() if '昵称' in k), None)
    time_col = next((c for k, c in hdr.items() if '入会时间' in k), None)
    if nick_col is None or time_col is None:
        raise ValueError('缺少「昵称 / 首次入会时间」列：%s' % os.path.basename(path))

    participants = []
    for row in target_rows[header_idx + 1:]:
        nick = str(row.get(nick_col, '') or '').strip()
        if not nick or '昵称' in nick:
            continue
        participants.append((nick, parse_dt(row.get(time_col, ''))))

    date = parse_dt(meta.get('预定开始时间', ''))
    if date is None:
        date = next((dt for _, dt in participants if dt), None)
    return {
        'date': date.date() if date else None,
        'meeting_id': meta.get('会议号', ''),
        'creator': meta.get('会议创建者', ''),
        'total': meta.get('参会用户总数', str(len(participants))),
        'participants': participants,
    }


# ---------------------------------------------------------------- 判定与统计
def classify(dt, rules):
    """首次入会 ≤ 准时线 → 准时；≤ 缺勤线（或不设缺勤线）→ 迟到；更晚或未入会 → 缺勤。"""
    if dt is None:
        return '缺勤'
    t = dt.time()
    if t <= rules['on_time']:
        return '准时'
    if rules['absent_after'] is None or t <= rules['absent_after']:
        return '迟到'
    return '缺勤'


def resolve_names(names, data):
    """把名字列表（支持别名）解析成规范名归一化 key 集合；返回 (keys, 无法识别的名字)。"""
    keys, unknown = set(), []
    for token in names:
        token = token.strip()
        if not token:
            continue
        key = data['aliases'].get(normalize(token))
        if key is None and normalize(token) in data['roster']:
            key = normalize(token)
        if key:
            keys.add(key)
        else:
            unknown.append(token)
    return keys, unknown


def split_names(text):
    return [t for t in re.split(r'[,，、;；\s]+', text or '') if t.strip()]


def load_leave_file(path, data):
    """读取请假表 CSV（列：日期,名字；日期写 2026-03-02）→ {date: {key}}。"""
    leaves = {}
    with open(path, encoding='utf-8-sig', newline='') as f:
        for row in csv.reader(f):
            if len(row) < 2 or not row[0].strip() or row[0].strip() in ('日期', 'date'):
                continue
            try:
                d = datetime.date.fromisoformat(row[0].strip())
            except ValueError:
                print('⚠️ 请假表日期格式不对（应为 2026-03-02）：%s' % row[0])
                continue
            keys, unknown = resolve_names([row[1]], data)
            for u in unknown:
                print('⚠️ 请假表里的名字不在花名册：%s' % u)
            leaves.setdefault(d, set()).update(keys)
    return leaves


def build_daily(meeting, data, leave_keys, rules):
    """
    单日统计，返回：
      {'date', 'meeting_id', 'creator', 'total', 'present': {key: dt}, 'status': {key: 状态},
       'unmatched': [(显示名, dt)], 'nonstat': [(name, identity, dt)], 'excluded_seen': [(name, dt)]}
    """
    roster, excluded = data['roster'], data['excluded']
    all_aliases = dict(data['aliases'])
    all_aliases.update(data['excluded_aliases'])
    present, nonstat, unmatched, excluded_seen = {}, [], [], []
    for display, dt in meeting['participants']:
        norm = normalize(display)
        if norm in excluded:
            excluded_seen.append((excluded[norm], dt))
            continue
        key = match_name(display, all_aliases)
        if key is None:
            # 兜底：显示名里含排除人员的名字（≥3 个字时才比对，避免误伤）
            hit = next((k for k in excluded if len(k) >= 3 and k in norm), None)
            if hit:
                excluded_seen.append((excluded[hit], dt))
            else:
                unmatched.append((display, dt))
            continue
        if key.startswith('EXCL:'):
            excluded_seen.append((key[5:], dt))
            continue
        info = roster.get(key)
        if info is None:
            unmatched.append((display, dt))
            continue
        if info['identity'] != rules['member_identity']:
            nonstat.append((info['name'], info['identity'], dt))
            continue
        # 同一人多个设备 / 多次入会：取最早的首次入会时间
        if key not in present or (dt is not None and (present[key] is None or dt < present[key])):
            present[key] = dt

    status = {}
    for key, info in roster.items():
        if info['identity'] != rules['member_identity']:
            continue
        if key in present:
            status[key] = classify(present[key], rules)
        elif key in leave_keys:
            status[key] = '请假'
        else:
            status[key] = '缺勤'
    return {
        'date': meeting['date'], 'meeting_id': meeting['meeting_id'],
        'creator': meeting['creator'], 'total': meeting['total'],
        'present': present, 'status': status, 'unmatched': unmatched,
        'nonstat': nonstat, 'excluded_seen': excluded_seen,
    }


# ---------------------------------------------------------------- 报告输出
def rules_text(rules):
    on = fmt_hms(rules['on_time'])
    if rules['absent_after'] is None:
        return '首次入会 ≤ %s 为准时；晚于 %s 为迟到；未入会为缺勤' % (on, on)
    return '首次入会 ≤ %s 为准时；%s ~ %s 为迟到；%s 之后入会或未入会为缺勤' % (
        on, fmt_hms(next_second(rules['on_time'])), fmt_hms(rules['absent_after']),
        fmt_hms(rules['absent_after']))


def render_daily(day, data, rules):
    """渲染单日日报（Markdown 文本）。"""
    roster = data['roster']
    d = day['date']
    date_str = d.strftime('%Y年%m月%d日') if d else '未知日期'
    wd = WEEKDAY_CN[d.weekday()] if d else '?'
    present, status = day['present'], day['status']

    on_time = sorted([k for k, v in status.items() if v == '准时'], key=lambda k: present[k])
    late = sorted([k for k, v in status.items() if v == '迟到'], key=lambda k: present[k])
    absent = sorted([k for k, v in status.items() if v == '缺勤'], key=lambda k: roster[k]['name'])
    overtime = [k for k in absent if present.get(k)]
    leaves = sorted([k for k, v in status.items() if v == '请假'], key=lambda k: roster[k]['name'])
    members = len(status)

    out = []
    out.append('# 📅 【%s】%s（周%s）考勤日报' % (rules['title'], date_str, wd))
    out.append('')
    out.append('> 判定规则：%s。' % rules_text(rules))
    out.append('')
    out.append('## 基本信息')
    out.append('- 会议号：%s ｜ 创建者：%s ｜ 导出参会记录：%s 条' % (
        day['meeting_id'] or '—', day['creator'] or '—', day['total']))
    if d and d.weekday() not in rules['workdays']:
        out.append('- ⚠️ 本日为周%s，不在统计日（%s）内，结果仅供参考。' % (wd, workdays_text(rules['workdays'])))
    out.append('- 应到：%s 人（花名册中身份为「%s」的人）' % (members, rules['member_identity']))
    uniq = len(present) + len(day['nonstat']) + len(day['unmatched']) + len(day['excluded_seen'])
    out.append('- 实到：成员 %s 人；另有非统计身份 %s 人、无法匹配 %s 人、排除人员 %s 人' % (
        len(present), len(day['nonstat']), len(day['unmatched']), len(day['excluded_seen'])))
    try:
        dup = int(day['total']) - uniq
    except (TypeError, ValueError):
        dup = 0
    if dup > 0:
        out.append('- 导出里有 %s 条同一人的重复入会记录，已按最早入会时间合并' % dup)
    out.append('')

    out.append('## ✅ 准时：%s 人' % len(on_time))
    out.append('，'.join('%s %s' % (roster[k]['name'], fmt_hm(present[k])) for k in on_time) or '无')
    out.append('')
    out.append('## ⏰ 迟到：%s 人' % len(late))
    out.append('，'.join('%s %s' % (roster[k]['name'], fmt_hm(present[k])) for k in late) or '无')
    out.append('')
    out.append('## ❌ 缺勤：%s 人' % len(absent))
    out.append('，'.join(roster[k]['name'] for k in absent) or '无')
    if overtime:
        out.append('')
        out.append('其中超过缺勤线才入会：%s' % '；'.join(
            '%s（%s）' % (roster[k]['name'], fmt_hm(present[k])) for k in overtime))
    out.append('')
    out.append('## 📝 请假：%s 人' % len(leaves))
    out.append('，'.join(roster[k]['name'] for k in leaves) or '无')
    out.append('')
    out.append('## 🚫 不计入统计的身份')
    out.append('，'.join('%s（%s）%s' % (n, idt, fmt_hm(dt)) for n, idt, dt in day['nonstat']) or '无')
    out.append('')
    out.append('## 👻 排除人员参会（只记录，不统计）')
    out.append('，'.join('%s %s' % (n, fmt_hm(dt)) for n, dt in day['excluded_seen']) or '无')
    out.append('')
    out.append('## ❓ 无法匹配（会议昵称不在花名册）')
    if day['unmatched']:
        for n, dt in day['unmatched']:
            out.append('- %s（%s）→ 请确认身份，补进花名册「名字映射」' % (n, fmt_hm(dt)))
    else:
        out.append('无')
    out.append('')
    pct = len(on_time) / members * 100 if members else 0
    out.append('## 📊 汇总')
    out.append('准时 %s/%s（%.1f%%）｜ 迟到 %s ｜ 缺勤 %s ｜ 请假 %s' % (
        len(on_time), members, pct, len(late), len(absent), len(leaves)))
    if on_time:
        out.append('')
        out.append('最早入会：%s（%s）' % (roster[on_time[0]]['name'], fmt_hm(present[on_time[0]])))
    out.append('')
    out.append('> 摄像头、请假理由、特殊豁免等导出表里没有的信息，需要管理员人工确认；本报告不做任何金额或罚则计算。')
    return '\n'.join(out) + '\n'


def render_week(days, data, rules):
    """多日 → 周报：个人状态矩阵、需关注名单、未全勤名单。days 为 build_daily 的结果列表。"""
    roster = data['roster']
    days = sorted(days, key=lambda x: x['date'] or datetime.date(1970, 1, 1))
    d0, d1 = days[0]['date'], days[-1]['date']
    span = '%s ~ %s' % (d0.isoformat(), d1.isoformat()) if d0 and d1 else '多日'
    counted = [i for i, x in enumerate(days) if x['date'] and x['date'].weekday() in rules['workdays']]

    out = []
    out.append('# 📅 【%s】%s 考勤周报' % (rules['title'], span))
    out.append('')
    out.append('> 判定规则：%s。统计日：%s；每周 %s 次请假不计，超出部分按迟到计。' % (
        rules_text(rules), workdays_text(rules['workdays']), rules['free_leaves']))
    for x in days:
        if x['date'] and x['date'].weekday() not in rules['workdays']:
            out.append('')
            out.append('⚠️ %s（周%s）不在统计日内，表中列出但不计入统计。' % (
                x['date'].isoformat(), WEEKDAY_CN[x['date'].weekday()]))
    out.append('')

    head = ['姓名'] + ['周%s %s' % (WEEKDAY_CN[x['date'].weekday()], x['date'].strftime('%m-%d'))
                      if x['date'] else '?' for x in days] + ['迟到', '请假', '缺勤', '需关注次数']
    out.append('| ' + ' | '.join(head) + ' |')
    out.append('|' + '---|' * len(head))
    short = {'准时': '准', '迟到': '迟', '缺勤': '缺', '请假': '假'}
    rows = []
    for key, info in sorted(roster.items(), key=lambda kv: kv[1]['name']):
        if info['identity'] != rules['member_identity']:
            continue
        cells, late_n, leave_n, absent_n, all_on_time = [], 0, 0, 0, True
        for i, x in enumerate(days):
            v = x['status'].get(key)
            cells.append(short.get(v, '·'))
            if i not in counted:
                continue
            if v != '准时':
                all_on_time = False
            late_n += v == '迟到'
            leave_n += v == '请假'
            absent_n += v == '缺勤'
        attention = late_n + max(0, leave_n - rules['free_leaves'])
        rows.append((info['name'], cells, late_n, leave_n, absent_n, attention, all_on_time))
        out.append('| %s | %s | %s | %s | %s | %s |' % (
            info['name'], ' | '.join(cells), late_n, leave_n, absent_n, attention))
    out.append('')
    out.append('图例：准 = 准时，迟 = 迟到，缺 = 缺勤，假 = 请假')
    out.append('')
    out.append('## ⏰ 需关注（迟到 + 超出免罚次数的请假）')
    need = [r for r in rows if r[5] > 0]
    out.append('，'.join('%s ×%s' % (r[0], r[5]) for r in need) or '无')
    out.append('')
    out.append('## ❌ 有缺勤')
    miss = [r for r in rows if r[4] > 0]
    out.append('，'.join('%s ×%s' % (r[0], r[4]) for r in miss) or '无')
    out.append('')
    out.append('## 🏅 统计日全部准时')
    if not counted:
        out.append('本期没有统计日的数据')
    else:
        full = [r[0] for r in rows if r[6]]
        out.append('，'.join(full) if full else '无')
    out.append('')
    out.append('> 请假以群内确认为准；只给了 `--leave` 时，周报会把这些人在每一天都记为请假，'
               '按日期区分请用 `--leave-file`。')
    return '\n'.join(out) + '\n'


# ---------------------------------------------------------------- 主流程
def find_roster(cli_path):
    if cli_path:
        return cli_path
    here = os.path.dirname(os.path.abspath(__file__))
    for p in ('roster.md', os.path.join(here, '..', 'references', 'roster.md')):
        if os.path.exists(p):
            return p
    sys.exit('❌ 没找到花名册。请用 --roster 指定，或把 examples/roster.example.md 复制成 roster.md 后修改。')


def main(argv=None):
    ap = argparse.ArgumentParser(description='腾讯会议考勤统计：一个文件出日报，多个文件再出周报')
    ap.add_argument('files', nargs='+', help='腾讯会议导出的「成员参会概况」.xlsx')
    ap.add_argument('--roster', help='花名册 Markdown（默认依次找 ./roster.md、../references/roster.md）')
    ap.add_argument('--title', help='报告标题，如「晨读打卡」（默认取花名册规则表，否则「会议考勤」）')
    ap.add_argument('--on-time', dest='on_time', help='准时线 HH:MM[:SS]，默认 06:00:00')
    ap.add_argument('--absent-after', dest='absent_after', help='缺勤线 HH:MM[:SS]，写 none 表示不设，默认 07:00:00')
    ap.add_argument('--workdays', help='统计日，如 1-5、1,3,5、周一至周五，默认周一至周五')
    ap.add_argument('--free-leaves', dest='free_leaves', type=int, help='每周免罚请假次数，默认 1')
    ap.add_argument('--leave', default='', help='请假名单（所有日期通用），如 "小鹿,阿布"')
    ap.add_argument('--leave-file', dest='leave_file', help='按日期的请假表 CSV：日期,名字')
    ap.add_argument('--out', default='.', help='输出目录，默认当前目录')
    ap.add_argument('--csv', action='store_true', help='额外导出每日 CSV 明细')
    ap.add_argument('--quiet', action='store_true', help='不在终端打印报告正文')
    args = ap.parse_args(argv)

    roster_path = find_roster(args.roster)
    if not os.path.exists(roster_path):
        sys.exit('❌ 花名册不存在：%s' % roster_path)
    data = parse_roster(roster_path)
    if not data['roster']:
        sys.exit('❌ 花名册解析为空，请检查「在册人员名单」表格格式')
    try:
        rules = build_rules(data['rules'], args)
    except ValueError as e:
        sys.exit('❌ 规则有误：%s' % e)
    members = [k for k, v in data['roster'].items() if v['identity'] == rules['member_identity']]
    print('✅ 花名册：在册 %s 人（计入应到 %s 人）、别名 %s 条、排除 %s 人' % (
        len(data['roster']), len(members), len(data['aliases']), len(data['excluded'])))

    common_leaves, unknown = resolve_names(split_names(args.leave), data)
    for u in unknown:
        print('⚠️ 请假名单里的名字不在花名册：%s' % u)
    dated_leaves = load_leave_file(args.leave_file, data) if args.leave_file else {}

    os.makedirs(args.out, exist_ok=True)
    days = []
    for fpath in args.files:
        if not os.path.exists(fpath):
            print('⚠️ 跳过不存在的文件：%s' % fpath)
            continue
        try:
            meeting = parse_meeting(fpath)
        except (ValueError, KeyError, zipfile.BadZipFile) as e:
            print('❌ %s：%s' % (os.path.basename(fpath), e))
            continue
        leave_keys = set(common_leaves) | dated_leaves.get(meeting['date'], set())
        day = build_daily(meeting, data, leave_keys, rules)
        days.append(day)
        report = render_daily(day, data, rules)
        if not args.quiet:
            print(report)
        name = '%s_考勤日报.md' % day['date'].strftime('%m%d') if day['date'] else '考勤日报.md'
        with open(os.path.join(args.out, name), 'w', encoding='utf-8') as f:
            f.write(report)
        print('💾 日报：%s' % os.path.join(args.out, name))
        if args.csv and day['date']:
            cname = '%s_考勤明细.csv' % day['date'].strftime('%m%d')
            with open(os.path.join(args.out, cname), 'w', encoding='utf-8-sig', newline='') as f:
                w = csv.writer(f)
                w.writerow(['姓名', '状态', '首次入会时间'])
                for key, info in sorted(data['roster'].items(), key=lambda kv: kv[1]['name']):
                    if info['identity'] != rules['member_identity']:
                        continue
                    dt = day['present'].get(key)
                    w.writerow([info['name'], day['status'].get(key, '缺勤'),
                                dt.strftime('%Y-%m-%d %H:%M:%S') if dt else ''])
            print('💾 CSV：%s' % os.path.join(args.out, cname))

    if len(days) > 1:
        week = render_week(days, data, rules)
        if not args.quiet:
            print(week)
        dates = [x['date'] for x in days if x['date']]
        name = '%s-%s_考勤周报.md' % (min(dates).strftime('%m%d'), max(dates).strftime('%m%d')) if dates else '考勤周报.md'
        with open(os.path.join(args.out, name), 'w', encoding='utf-8') as f:
            f.write(week)
        print('💾 周报：%s' % os.path.join(args.out, name))
    return 0 if days else 1


if __name__ == '__main__':
    sys.exit(main())
