# Agent Skill Toolbox

面向 Agent 的技能卡片工具箱：读取 `skills/` 下各技能的 `meta.json`，自动生成静态网页 `docs/index.html`，以卡片形式展示全部技能，并支持 GitHub Pages 在线访问。

> 仓库地址：https://github.com/chaoyue-zhang/Agent-skill-toolbox
> 在线地址：https://chaoyue-zhang.github.io/Agent-skill-toolbox/

## 目录结构

```
Agent-skill-toolbox/
├── docs/                       # 静态网页输出目录（自动生成，勿手改）
│   ├── index.html
│   └── assets/                 # 构建时从 skills/*/assets 复制过来的截图
├── skills/                     # 技能源数据
│   └── <skill-id>/
│       ├── meta.json           # 技能元信息【必填】
│       ├── README.md           # 技能详细文档
│       └── assets/preview.png  # 预览截图（可选）
├── scripts/
│   └── build_static_html.py    # 网页构建脚本（核心）
├── .github/workflows/
│   └── deploy-pages.yml        # 自动构建 + Pages 部署流水线
├── Makefile
└── requirements.txt
```

## 快速开始

```bash
# 1. 本地构建
make all          # 等价于 python3 scripts/build_static_html.py

# 2. 本地预览（直接打开生成的页面）
make open         # 或手动双击 docs/index.html

# 3. 提交推送
git add .
git commit -m "feat: init Agent Skill Toolbox"
git push origin main
```

## 部署到 GitHub Pages

1. 推送代码到 `main` 分支
2. 仓库 **Settings → Pages**
3. **Build and deployment → Source** 选择 `GitHub Actions`，保存
4. 之后每次 push 到 `main`，流水线自动执行「构建 → 发布」
5. 首次部署需等待约 1~3 分钟；若仍 404，用 `Ctrl+F5` 强制刷新

访问地址：`https://<用户名>.github.io/<仓库名>/`（**项目仓库必须带仓库名**，只有 `<用户名>.github.io` 这个仓库才能直接用裸域名）

## 新增一个技能

1. 在 `skills/` 下新建文件夹，名称 = 技能 id，例如 `skills/my-skill/`
2. 编写 `meta.json`（可从 `skills/skill-demo/meta.json` 复制），截图放 `assets/`
3. 执行 `make all` 重新生成页面
4. `git add . && git commit -m "feat: add my-skill" && git push`

### meta.json 字段说明

| 字段 | 类型 | 必填 | 说明 |
|---|---|---|---|
| id | string | ✅ | 技能唯一标识，须与文件夹名一致 |
| name | string | ✅ | 展示名称 |
| category | string | ✅ | 分类，用于下拉筛选 |
| version | string | ✅ | 版本号 |
| status | string | ✅ | draft / active / deprecated / archived（决定徽标颜色） |
| description | string | ✅ | 一句话简介 |
| tags | string[] | ❌ | 标签 |
| create_date | string | ❌ | 创建日期 |
| update_date | string | ❌ | 更新日期 |
| usage | string | ❌ | 使用方式说明 |
| notes | string | ❌ | 注意事项 |
| dependencies | string[] | ❌ | 依赖项 |
| params | object[] | ❌ | 参数表，每项含 key / type / required / desc |
| screenshot | string | ❌ | 相对技能目录的截图路径，如 `assets/preview.png` |

## 网页功能

- 技能卡片总览 + 截图缩略图
- 关键词搜索（名称 / ID）
- 分类筛选、仅看收藏⭐
- ⭐收藏（localStorage），支持导入 / 导出 / 重置
- 卡片拖拽排序，顺序本地持久化
- 明暗主题切换，默认跟随系统
- 弹窗查看详情：参数表、依赖、时间、GitHub 文档链接
- 截图点击放大
- 筛选结果导出 JSON

## 构建脚本配置项

编辑 `scripts/build_static_html.py` 顶部配置区：

```python
GITHUB_REPO_URL = "https://github.com/chaoyue-zhang/Agent-skill-toolbox"
GITHUB_BRANCH   = "main"
SCREENSHOT_MODE = "copy"   # "copy"=截图复制进 docs/assets（推荐）；"github"=引用 GitHub 裸链接
```

- `copy` 模式：页面自包含，`file://` 本地预览正常，Pages 带子路径也不会 404，仓库改名 / 转私有都不影响。
- `github` 模式：截图不落入 `docs/`，但仅公开仓库可用，且浏览器需能直连 GitHub。

## 常见问题

| 现象 | 排查方向 |
|---|---|
| Pages 404 | 是否已 push 到 `main`；Actions 流水线是否绿色；是否等够 1~3 分钟；`Ctrl+F5` |
| 图片不显示 | 检查 `screenshot` 路径是否正确、源文件是否存在；用 `copy` 模式最稳 |
| `make all` 报错 | Python ≥ 3.8；确认在仓库根目录执行，脚本路径为 `scripts/build_static_html.py` |
| 页面改动没生效 | `docs/index.html` 由脚本生成，改完 `meta.json` 必须重新 `make all` |

## 维护约定

- **不要手动修改 `docs/index.html`**，每次 `make all` 都会覆盖
- 所有技能信息统一维护在 `skills/<id>/meta.json`
- 修改构建脚本后重新 `make all` 并提交推送，页面即更新
