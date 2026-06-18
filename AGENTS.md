# CorelDRAW VBA Skills Project

一个存放 CorelDRAW VBA 宏生成技能的 Reasonix 项目。通过自然语言描述生成 CorelDRAW VBA 宏代码，在 CorelDRAW 中运行即可自动绘图。

## Project

- **用途**：本项目本身不含代码，仅托管 Reasonix 技能包
- **核心技能**：`coreldraw-vba` — 将口头/文字描述转为 CorelDRAW VBA 宏
- **技能位置**：`.reasonix/skills/coreldraw-vba/SKILL.md`

## Commands

本项目无构建/测试命令。唯一操作：

- `/coreldraw-vba` — 调用绘图技能（或在对话中自然描述需求即可自动触发）

## Architecture

```
.reasonix/skills/coreldraw-vba/SKILL.md   ← 绘图技能主体
AGENTS.md                                  ← 本文件
```

## Conventions

- 所有 CorelDRAW VBA 代码生成工作使用 `coreldraw-vba` 技能模板
- 生成代码默认单位：毫米（mm）
- 默认页面：A4
- 始终包含撤销分组（`BeginCommandGroup/EndCommandGroup`）和错误处理

## Notes

（留空待补充）
