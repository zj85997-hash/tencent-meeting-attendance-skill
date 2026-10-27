# -*- coding: utf-8 -*-
"""运行：python3 -m unittest discover -s tests -v"""

import datetime
import os
import subprocess
import sys
import tempfile
import types
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, 'scripts'))

import attendance_stat as A  # noqa: E402
import make_sample as S      # noqa: E402

ROSTER = os.path.join(ROOT, 'examples', 'roster.example.md')
LEAVES = os.path.join(ROOT, 'examples', 'leaves.example.csv')
SCRIPT = os.path.join(ROOT, 'scripts', 'attendance_stat.py')


def at(hms, day=datetime.date(2026, 3, 2)):
    return datetime.datetime.combine(day, datetime.time.fromisoformat(hms))


def cli(**kw):
    base = dict(title=None, on_time=None, absent_after=None, workdays=None, free_leaves=None)
    base.update(kw)
    return types.SimpleNamespace(**base)


class RulesTest(unittest.TestCase):
    def test_default_boundaries(self):
        rules = A.build_rules({})
        self.assertEqual(A.classify(at('06:00:00'), rules), '准时')
        self.assertEqual(A.classify(at('06:00:01'), rules), '迟到')
        self.assertEqual(A.classify(at('07:00:00'), rules), '迟到')
        self.assertEqual(A.classify(at('07:00:01'), rules), '缺勤')
        self.assertEqual(A.classify(None, rules), '缺勤')

    def test_no_absent_line(self):
        rules = A.build_rules({'absent_after': '无'})
        self.assertIsNone(rules['absent_after'])
        self.assertEqual(A.classify(at('09:30:00'), rules), '迟到')

    def test_cli_overrides_roster(self):
        rules = A.build_rules({'on_time': '06:00:00'}, cli(on_time='20:00', absent_after='20:30'))
        self.assertEqual(A.classify(at('19:59:59'), rules), '准时')
        self.assertEqual(A.classify(at('20:15:00'), rules), '迟到')
        self.assertEqual(A.classify(at('20:30:01'), rules), '缺勤')

    def test_descriptive_roster_values(self):
        rules = A.build_rules({'on_time': '首次入会 ≤ 05:30', 'workdays': '周一至周五（周六不计）',
                               'free_leaves': '每周 2 次'})
        self.assertEqual(rules['on_time'], datetime.time(5, 30))
        self.assertEqual(rules['workdays'], {0, 1, 2, 3, 4})
        self.assertEqual(rules['free_leaves'], 2)

    def test_unreadable_roster_value_keeps_default(self):
        rules = A.build_rules({'on_time': '看群公告'})
        self.assertEqual(rules['on_time'], datetime.time(6, 0))

    def test_absent_line_before_on_time_is_rejected(self):
        with self.assertRaises(ValueError):
            A.build_rules({}, cli(on_time='07:00', absent_after='06:30'))

    def test_parse_workdays(self):
        self.assertEqual(A.parse_workdays('1-5'), {0, 1, 2, 3, 4})
        self.assertEqual(A.parse_workdays('周一、周三、周五'), {0, 2, 4})
        self.assertEqual(A.parse_workdays('6,7'), {5, 6})
        with self.assertRaises(ValueError):
            A.parse_workdays('5-1')

    def test_excel_serial_rounds_to_second(self):
        serial = (at('06:00:00') - datetime.datetime(1899, 12, 30)) / datetime.timedelta(days=1)
        self.assertEqual(A.parse_dt(repr(serial)), at('06:00:00'))
        self.assertEqual(A.parse_dt('2026/3/2 06:00:00'), at('06:00:00'))


class RosterTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.data = A.parse_roster(ROSTER)

    def test_counts(self):
        roster = self.data['roster']
        self.assertEqual(len(roster), 18)
        self.assertEqual(len([v for v in roster.values() if v['identity'] == '成员']), 16)
        self.assertEqual(sorted(self.data['excluded'].values()), ['旧书签', '远方'])

    def test_rules_table(self):
        rules = A.build_rules(self.data['rules'])
        self.assertEqual(rules['title'], '晨读打卡（示例）')
        self.assertEqual(rules['absent_after'], datetime.time(7, 0))

    def test_name_variants(self):
        cases = [('小鹿(三组)', '小鹿'), ('南风（李）', '南风'), ('LILY-广州', 'Lily'),
                 ('Kevin(Zhang)', 'Kevin'), ('kevin', 'Kevin'), ('一一的iPad', '一一'),
                 ('海盐（上海）', '海盐'), ('Abu', '阿布'), ('麦穗', '麦穗')]
        for display, name in cases:
            key = A.match_name(display, self.data['aliases'])
            self.assertIsNotNone(key, display)
            self.assertEqual(self.data['roster'][key]['name'], name, display)
        self.assertIsNone(A.match_name('访客123', self.data['aliases']))


class EndToEndTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.files = []
        for i, day in enumerate(S.WEEK):
            path = os.path.join(cls.tmp.name, '每日晨读会-%s.xlsx' % day.strftime('%Y%m%d'))
            S.write_xlsx(path, '成员参会概况', S.day_rows(i, day))
            cls.files.append(path)
        cls.data = A.parse_roster(ROSTER)
        cls.rules = A.build_rules(cls.data['rules'])

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def daily(self, i, leaves=()):
        keys, unknown = A.resolve_names(list(leaves), self.data)
        self.assertEqual(unknown, [])
        return A.build_daily(A.parse_meeting(self.files[i]), self.data, keys, self.rules)

    def names(self, day):
        return {self.data['roster'][k]['name']: v for k, v in day['status'].items()}

    def test_monday(self):
        day = self.daily(0)
        s = self.names(day)
        self.assertEqual(day['date'], datetime.date(2026, 3, 2))
        self.assertEqual(day['meeting_id'], '900000001')
        self.assertEqual(s['Kevin'], '迟到')      # 06:00:01
        self.assertEqual(s['木木'], '缺勤')       # 07:05 入会
        self.assertEqual(s['星河'], '缺勤')       # 没来
        self.assertEqual(s['一一'], '准时')       # 两台设备取最早
        self.assertEqual(list(s.values()).count('准时'), 13)
        self.assertEqual([n for n, _ in day['excluded_seen']], ['旧书签'])
        self.assertEqual(sorted(n for n, _, _ in day['nonstat']), ['小满', '晨光'])
        self.assertEqual(day['unmatched'], [])

    def test_wednesday_boundary_and_leave(self):
        s = self.names(self.daily(2, leaves=['布丁妈妈']))
        self.assertEqual(s['海盐'], '准时')       # 正好 06:00:00
        self.assertEqual(s['阿布'], '缺勤')
        self.assertEqual(s['布丁'], '请假')

    def test_thursday_unmatched(self):
        day = self.daily(3)
        self.assertEqual([n for n, _ in day['unmatched']], ['访客123'])
        self.assertEqual(self.names(day)['青柠'], '迟到')

    def test_cli_week_report(self):
        out = os.path.join(self.tmp.name, 'out')
        proc = subprocess.run([sys.executable, SCRIPT, *self.files, '--roster', ROSTER,
                               '--leave-file', LEAVES, '--out', out, '--quiet'],
                              capture_output=True, text=True)
        self.assertEqual(proc.returncode, 0, proc.stderr)
        with open(os.path.join(out, '0302-0306_考勤周报.md'), encoding='utf-8') as f:
            week = f.read()
        self.assertIn('| 布丁 | 准 | 假 | 假 | 准 | 准 | 0 | 2 | 0 | 1 |', week)   # 第 2 次请假按迟到计
        self.assertIn('| 青柠 | 准 | 迟 | 准 | 迟 | 准 | 2 | 0 | 0 | 2 |', week)
        self.assertIn('| 阿布 | 准 | 准 | 缺 | 准 | 迟 | 1 | 0 | 1 | 1 |', week)
        for d in ('0302', '0303', '0304', '0305', '0306'):
            self.assertTrue(os.path.exists(os.path.join(out, '%s_考勤日报.md' % d)))

    def test_cli_override_on_time(self):
        out = os.path.join(self.tmp.name, 'out2')
        proc = subprocess.run([sys.executable, SCRIPT, self.files[0], '--roster', ROSTER,
                               '--on-time', '05:50', '--out', out, '--quiet'],
                              capture_output=True, text=True)
        self.assertEqual(proc.returncode, 0, proc.stderr)
        with open(os.path.join(out, '0302_考勤日报.md'), encoding='utf-8') as f:
            report = f.read()
        self.assertIn('## ✅ 准时：2 人', report)   # 周一只有知夏 05:46:58、麦穗 05:48:45 早于 05:50
