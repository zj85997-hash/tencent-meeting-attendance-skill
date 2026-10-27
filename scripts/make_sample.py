#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生成虚构的腾讯会议「成员参会概况」导出表（.xlsx），用于演示和测试。

    python3 scripts/make_sample.py              # 写到 examples/exports/
    python3 scripts/make_sample.py --out DIR    # 写到指定目录

人物与 examples/roster.example.md 对应，全部是虚构的。
一周 5 天的数据覆盖了脚本的所有判定情况：边界时间、迟到、超时缺勤、未入会、
请假、一人多设备、显示名变体、非统计身份、已退出人员、无法匹配的访客。
只用 Python 标准库写 xlsx。
"""

import argparse
import datetime
import os
import zipfile
from xml.sax.saxutils import escape

WEEK = [datetime.date(2026, 3, 2) + datetime.timedelta(days=i) for i in range(5)]  # 周一至周五

# 每人默认的入会时间（时:分:秒）。None = 当天没来。
BASE = {
    '晨光': '05:40:00', '小满': '05:45:10',
    '海盐（上海）': '05:52:31', '小鹿(三组)': '05:55:02', '麦穗': '05:48:45',
    '山茶花': '05:57:20', 'Abu': '05:58:03', '青柠-手机': '05:59:30', '木木': '05:50:12',
    '南风（李）': '05:53:40', '知夏': '05:46:58', '布丁妈妈': '05:56:15', '星河': '05:58:59',
    '安安': '05:51:07', 'LILY-广州': '05:54:44', 'Kevin(Zhang)': '05:49:30', '七七': '05:57:57',
    '一一': '05:58:30', '一一的iPad': '06:02:00',
}

# 按日期覆盖默认时间（下标 0 = 周一）。
OVERRIDES = {
    0: {'木木': '07:05:00',          # 超过缺勤线才入会 → 缺勤
        'Kevin(Zhang)': '06:00:01',  # 边界：比准时线晚 1 秒 → 迟到
        '星河': None,                # 没来、也没请假 → 缺勤
        '旧书签': '05:59:00'},        # 已退出人员 → 只记录不统计
    1: {'青柠-手机': '06:03:12',      # 迟到
        '布丁妈妈': None},            # 请假（见 leaves.example.csv）
    2: {'海盐（上海）': '06:00:00',    # 边界：正好准时线 → 准时
        'Abu': None,                 # 缺勤
        '布丁妈妈': None},            # 本周第 2 次请假 → 周报按迟到计
    3: {'青柠-手机': '06:15:40',      # 迟到
        '访客123': '06:05:00'},       # 不在花名册 → 无法匹配
    4: {'Abu': '06:30:05',           # 迟到
        'LILY-广州': None},           # 缺勤
}

IDENTITY = {'晨光': '主持人'}


def col_letter(i):
    s = ''
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def write_xlsx(path, sheet_name, rows):
    """把二维字符串列表写成单 sheet 的 xlsx（共享字符串表，和腾讯会议导出一致）。"""
    strings, index = [], {}

    def sid(s):
        if s not in index:
            index[s] = len(strings)
            strings.append(s)
        return index[s]

    xml_rows = []
    for r, row in enumerate(rows, 1):
        cells = []
        for c, val in enumerate(row):
            if val is None or val == '':
                continue
            cells.append('<c r="%s%d" t="s"><v>%d</v></c>' % (col_letter(c), r, sid(str(val))))
        xml_rows.append('<row r="%d">%s</row>' % (r, ''.join(cells)))

    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    rns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    pns = 'http://schemas.openxmlformats.org/package/2006/relationships'
    files = {
        '[Content_Types].xml':
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            '</Types>',
        '_rels/.rels':
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="%s">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>' % pns,
        'xl/workbook.xml':
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="%s" xmlns:r="%s"><sheets>'
            '<sheet name="%s" sheetId="1" r:id="rId1"/></sheets></workbook>' % (ns, rns, escape(sheet_name)),
        'xl/_rels/workbook.xml.rels':
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="%s">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '</Relationships>' % pns,
        'xl/styles.xml':
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<styleSheet xmlns="%s">'
            '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
            '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
            '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
            '</styleSheet>' % ns,
        'xl/worksheets/sheet1.xml':
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="%s"><sheetData>%s</sheetData></worksheet>' % (ns, ''.join(xml_rows)),
    }
    files['xl/sharedStrings.xml'] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<sst xmlns="%s" count="%d" uniqueCount="%d">%s</sst>' % (
            ns, len(strings), len(strings),
            ''.join('<si><t xml:space="preserve">%s</t></si>' % escape(s) for s in strings)))
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, content in files.items():
            z.writestr(name, content)


def day_rows(i, day):
    times = dict(BASE)
    times.update(OVERRIDES.get(i, {}))
    records = []
    for name, hms in times.items():
        if hms is None:
            continue
        first = datetime.datetime.combine(day, datetime.time.fromisoformat(hms))
        last = datetime.datetime.combine(day, datetime.time(7, 0, 0)) + datetime.timedelta(seconds=len(name) * 7)
        stay = max(last - first, datetime.timedelta(minutes=1))
        minutes, seconds = divmod(int(stay.total_seconds()), 60)
        records.append([name, first.strftime('%Y-%m-%d %H:%M:%S'), last.strftime('%Y-%m-%d %H:%M:%S'),
                        '1', '%d分钟%d秒' % (minutes, seconds), IDENTITY.get(name, '参会者')])
    records.sort(key=lambda r: r[1])
    start = datetime.datetime.combine(day, datetime.time(5, 30))
    meta = [
        ['会议主题:', '每日晨读会（示例）'],
        ['会议号:', '900000001'],
        ['会议创建者:', '晨光'],
        ['预定开始时间:', start.strftime('%Y-%m-%d %H:%M:%S')],
        ['预定结束时间:', start.replace(hour=7, minute=0).strftime('%Y-%m-%d %H:%M:%S')],
        ['累计会议时长:', '90分钟'],
        ['参会用户总数:', str(len(records))],
        [],
        ['用户昵称（入会昵称）', '首次入会时间', '最后退会时间', '入会次数', '累计参会时长', '身份'],
    ]
    return meta + records


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    ap = argparse.ArgumentParser(description='生成虚构的腾讯会议导出表')
    ap.add_argument('--out', default=os.path.join(here, '..', 'examples', 'exports'))
    args = ap.parse_args()
    os.makedirs(args.out, exist_ok=True)
    for i, day in enumerate(WEEK):
        path = os.path.join(args.out, '每日晨读会-示例-%s.xlsx' % day.strftime('%Y%m%d'))
        write_xlsx(path, '成员参会概况', day_rows(i, day))
        print('✅ %s' % path)


if __name__ == '__main__':
    main()
