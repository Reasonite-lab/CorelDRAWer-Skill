# CorelDRAWer-Skill 🎨

> 用自然语言画 CorelDRAW 图——你的 AI 绘图助手

## 这是什么

这是一个 **Reasonix 技能**（AI playbook），能将你的口头/文字描述自动转换成 **CorelDRAW VBA 宏代码**。在 CorelDRAW 中运行生成的宏，即可自动绘制出你想要的图形。

### 示例

| 你说 | 它生成 |
|------|--------|
| "画一个红色圆形，直径 50mm，居中" | ✅ VBA 宏，运行即画 |
| "画一排 5 个正方形渐变填色" | ✅ VBA 宏，运行即画 |
| "画一个带阴影的标题文字在页面顶部" | ✅ VBA 宏，运行即画 |

## 快速开始

### 环境要求

- [Reasonix](https://reasonix.ai) AI 编程助手
- CorelDRAW (Windows, X3~2025，需安装 VBA 组件)

### 使用方式

**方式一：自然对话**
直接在 Reasonix 中说你想画什么，技能会自动触发并生成代码。

**方式二：手动调用**
```
/coreldraw-vba 画一个蓝色圆角矩形，100x60mm，左上角在(10,10)
```

**在 CorelDRAW 中运行：**
1. Alt+F11 打开 VBA 编辑器
2. 插入 → 模块
3. 粘贴生成的代码
4. F5 运行

## 技能内容

核心技能文件位于 `.reasonix/skills/coreldraw-vba/SKILL.md`，包含：
- CorelDRAW VBA 完整对象模型参考
- 形状创建 API（矩形、椭圆、曲线、文字等）
- 填充/轮廓/变换方法
- 最佳实践模板（撤销组、错误处理）

## 项目结构

```
.reasonix/skills/coreldraw-vba/SKILL.md    ← 绘图技能主体
AGENTS.md                                    ← 项目描述
reasonix.toml                                ← Reasonix 配置
```

## 贡献

欢迎提交 Issue 或 PR 来丰富技能支持的绘图能力！
