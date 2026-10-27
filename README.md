# 腾讯会议考勤统计 Skill

> 把腾讯会议导出的「成员参会概况」表丢给 AI，或者自己跑一行命令，就能得到准时 / 迟到 / 缺勤 / 请假名单，以及一份周报。
> 零依赖 Python，可以单独运行，也可以作为 Skill 装进 Claude Code、WorkBuddy 等支持 Skill 的 AI 工具。

**Tencent Meeting attendance counter** — reads the "成员参会概况" (participant summary) .xlsx exported from Tencent Meeting, matches nicknames against a Markdown roster, and writes daily and weekly attendance reports. Pure Python standard library; works as a standalone CLI or as an AI agent skill. Docs are in Chinese.

---

## 为什么做这个

我在一个读书会负责晨读打卡统计：每天早上 6 点开腾讯会议，几十个人入会。人工对名单很费时间，最麻烦的是大家的会议昵称五花八门，有「小鹿(三组)」「xx的iPad」、英文名，还有家长代孩子入会。

后来我把规则写成 Skill 交给 AI 来跑，群里的朋友把文件夹装进自己的 AI 工具也能用。这个仓库是它的通用版：读书会的专属规则都去掉了，样例数据全部是虚构的。

## 能做什么

- **日报**：准时（按入会时间排序）、迟到、缺勤（单独标出超时入会）、请假、非统计身份、排除人员、无法匹配的昵称，最后是汇总。
- **周报**：每人每天的状态矩阵，以及需关注名单、缺勤名单、全勤名单。
- **昵称匹配**：整串匹配、括号拆分、英文不分大小写、一人多设备自动合并。
- **规则可配置**：准时线、缺勤线、统计日、每周免罚请假次数，写在花名册里或用命令行参数改。
- **CSV 明细**：可以额外导出，方便放进表格。

## 30 秒试跑

需要 Python 3.8 及以上，不用装任何包：

```bash
git clone https://github.com/zj85997-hash/tencent-meeting-attendance-skill.git
cd tencent-meeting-attendance-skill

# 用自带的虚构数据跑一周
python3 scripts/attendance_stat.py examples/exports/*.xlsx \
  --roster examples/roster.example.md \
  --leave-file examples/leaves.example.csv \
  --out out
```

跑完 `out/` 里有 5 份日报和 1 份周报，和 [`examples/output/`](examples/output/) 里的一样。周报片段：

| 姓名 | 周一 03-02 | 周二 03-03 | 周三 03-04 | 周四 03-05 | 周五 03-06 | 迟到 | 请假 | 缺勤 | 需关注次数 |
|---|---|---|---|---|---|---|---|---|---|
| 布丁 | 准 | 假 | 假 | 准 | 准 | 0 | 2 | 0 | 1 |
| 青柠 | 准 | 迟 | 准 | 迟 | 准 | 2 | 0 | 0 | 2 |
| 阿布 | 准 | 准 | 缺 | 准 | 迟 | 1 | 0 | 1 | 1 |

## 用在自己的群

1. **导出数据**：腾讯会议 → 历史会议 → 会议详情 → 导出「成员参会概况」，得到一个 .xlsx。
2. **写花名册**：把 [`examples/roster.example.md`](examples/roster.example.md) 复制成 `roster.md`，填上：
   - **规则**：比如晚上 8 点的课，准时线写 `20:00:00`。
   - **在册人员名单**：身份写「成员」的人计入应到；主持人、助教写别的身份。
   - **名字映射**：会议昵称和名单不一致的人，在这里登记。昵称完全一样，或者只多了括号后缀的，不用登记。
3. **运行**：

```bash
python3 scripts/attendance_stat.py 你的导出.xlsx --roster roster.md
```

第一次跑，多半会有几个「❓ 无法匹配」的昵称。把它们补进名字映射，再跑一次就齐了。

### 常用参数

| 参数 | 作用 | 例子 |
|---|---|---|
| `--roster` | 花名册路径（默认 `./roster.md`） | `--roster roster.md` |
| `--on-time` | 准时线 | `--on-time 20:00` |
| `--absent-after` | 缺勤线；写 `none` 表示不设 | `--absent-after 20:30` |
| `--workdays` | 统计日 | `--workdays 1-5`、`--workdays 周一,周三,周五` |
| `--free-leaves` | 每周免罚请假次数 | `--free-leaves 2` |
| `--leave` | 请假名单（所有日期通用） | `--leave "小鹿,阿布"` |
| `--leave-file` | 按日期的请假表（CSV：日期,名字） | `--leave-file leaves.csv` |
| `--out` | 输出目录 | `--out reports` |
| `--csv` | 额外导出每日 CSV 明细 | |
| `--quiet` | 不在终端打印报告正文 | |

优先级：命令行参数 > 花名册「规则」表 > 默认值（06:00:00 / 07:00:00 / 周一至周五 / 每周 1 次）。

## 作为 AI Skill 使用

[`SKILL.md`](SKILL.md) 写好了触发条件、工作流和注意事项，可以直接给 AI 读。

- **Claude Code**：把整个文件夹放进 `~/.claude/skills/tencent-meeting-attendance/`（或项目里的 `.claude/skills/`），之后在对话里发导出表，说「统计今天的考勤」就行。
- **WorkBuddy、豆包等其他 AI 工具**：把文件夹作为技能或附件交给它，并提醒它「按 SKILL.md 的流程，用 scripts 里的脚本统计」。

AI 只负责调用脚本、解读报告、追问无法匹配的昵称，**判定由脚本完成**，避免 AI 自己数错或编造。

## 目录结构

```
├── SKILL.md                  # 给 AI 读的技能说明
├── scripts/
│   ├── attendance_stat.py    # 统计脚本（零依赖）
│   └── make_sample.py        # 生成虚构导出表（演示 / 测试用）
├── examples/
│   ├── roster.example.md     # 花名册模板（虚构人物）
│   ├── leaves.example.csv    # 请假表模板
│   ├── exports/              # 虚构的一周导出表
│   └── output/               # 用上面数据跑出的日报和周报
└── tests/
    └── test_attendance.py    # 单元测试
```

## 测试

```bash
python3 -m unittest discover -s tests -v
```

测试覆盖以下情况：边界时间（06:00:00 / 06:00:01 / 07:00:01）、昵称变体、多设备去重、请假与免罚次数、非统计日、命令行覆盖规则、Excel 日期数值。

## 隐私提醒

花名册和导出表里有成员昵称和入会时间。`.gitignore` 已经忽略 `roster.md`、`data/`、`out/` 和根目录下的 `.xlsx`，请不要把真实数据提交到公开仓库。仓库里的所有人名和数据都是虚构的。

## 关于作者

付钟浩（Alex），AI 培训讲师、FDE（前沿部署工程师）。这个工具在我的课上用作「怎么把一件重复的小事做成 Skill」的案例：先把规则写清楚，再让脚本判定，AI 只负责调度和沟通。

## 许可

[MIT](LICENSE)
