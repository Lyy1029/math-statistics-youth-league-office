# 数学与统计学院团委办公室 Skill

面向高校学院团委办公室日常工作的 Agent Skill，包含通知改写、活动与会务、逐级审批、证明制作、评优评奖、招新留任、学生信息台账及隐私归档规范。

## 快速制证

- 活动证明和突出贡献证明：从批准名单批量生成 DOCX/PDF。
- 干事任职证明：按组织和部门汇总，生成带姓名、班级、职务和学号的 DOCX/PDF。
- 部长及以上任职证明：适用于部长、主席、团委副书记等成员，保留正式 PDF 原稿的边框、徽标、水印、标题、横线和落款，直接输出一人一份 PDF。
- PDF 动态文字统一使用并嵌入随 Skill 打包的 Noto Serif SC，检测到字体替代时停止交付。

## 使用

将整个仓库放入 Agent 可读取的 Skill 目录，使用时显式触发 `$youth-league-office`。不要只复制 `SKILL.md`；`assets/`、`references/` 和 `scripts/` 都是必要组成部分。

任职证明批量生成示例：

```bash
python3 scripts/generate_appointment_certificate.py generate \
  --list 任职证明批准名单.xlsx \
  --academic-year 2025-2026 \
  --issue-date 2025-09-01 \
  --output-dir 待审核任职证明
```

## 边界

生成器只接受正式批准名单，不使用报名表直接制证，不将加分、考勤或评定信息写入对外证明，不生成、复制或合成公章。含班级和学号的干事任职证明只能受控审核和定向交付。
