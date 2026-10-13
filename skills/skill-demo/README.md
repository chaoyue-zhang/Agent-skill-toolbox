# 示例技能 (skill-demo)

用于演示 Agent Skill Toolbox 卡片结构的模板技能。

## 功能

- 演示 `meta.json` 全部字段的写法
- 演示参数表（params）在卡片弹窗中的渲染效果
- 作为新增技能的复制起点

## 使用方式

1. 复制 `skills/skill-demo/` 整个目录
2. 重命名为你的技能 ID，例如 `skills/my-skill/`
3. 修改 `meta.json` 中的 id / name / category / description 等字段
4. 把截图放到 `assets/` 下并在 `screenshot` 字段引用
5. 在仓库根目录执行 `make all` 重新生成页面

## 参数

| 参数名 | 类型 | 是否必填 | 说明 |
|---|---|---|---|
| input_text | string | ✅ | 输入文本参数 |
| style | string | ❌ | 输出风格，默认 default |

## 注意事项

- `id` 必须与技能文件夹名称一致，页面用它拼接 GitHub 文档链接。
- 不要把 `docs/index.html` 手动改动，每次构建都会覆盖。
