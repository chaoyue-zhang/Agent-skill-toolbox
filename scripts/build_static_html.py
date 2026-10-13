#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Agent Skill Toolbox - 技能卡片静态页构建脚本
读取 skills/*/meta.json -> 生成 docs/index.html

功能：
1. 技能卡片总览 + 截图预览 + 点击放大
2. 关键词搜索 / 分类筛选 / 仅看收藏
3. ⭐收藏导入 / 导出 / 重置（localStorage）
4. 卡片拖拽排序（localStorage 持久化）
5. 明暗主题切换（默认跟随系统）
6. 筛选结果导出 JSON

依赖：仅 Python 标准库（>=3.8）
"""
import json
import shutil
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
SKILLS_DIR = REPO_ROOT / "skills"
DOCS_DIR = REPO_ROOT / "docs"
OUTPUT_HTML = DOCS_DIR / "index.html"
ASSETS_OUT = DOCS_DIR / "assets"

# ========= 配置区：改成你自己的 GitHub 仓库地址 =========
GITHUB_REPO_URL = "https://github.com/chaoyue-zhang/Agent-skill-toolbox"
GITHUB_BRANCH = "main"

# 截图引用方式：
#   "copy"   -> 把 skills/<id>/assets 下的截图复制到 docs/assets/，页面用相对路径引用
#               （推荐：页面自包含，本地 file:// 直接打开也能看，Pages 带子路径也不会 404）
#   "github" -> 直接引用 GitHub 裸地址（仅公开仓库 + 分支名正确时可用，浏览器需能访问 GitHub）
SCREENSHOT_MODE = "copy"
# ======================================================

STATUS_COLOR_MAP = {
    "draft": "#f59e0b",
    "active": "#10b981",
    "deprecated": "#f97316",
    "archived": "#6b7280",
}


def load_all_skills():
    """扫描 skills/ 下所有含 meta.json 的技能目录"""
    skills = []
    if not SKILLS_DIR.exists():
        print(f"[WARN] skills dir not found: {SKILLS_DIR}")
        return skills
    for skill_dir in sorted(SKILLS_DIR.iterdir()):
        if not skill_dir.is_dir():
            continue
        meta_path = skill_dir / "meta.json"
        if not meta_path.exists():
            continue
        try:
            with open(meta_path, "r", encoding="utf-8") as f:
                meta = json.load(f)
            meta.setdefault("_folder", skill_dir.name)
            skills.append(meta)
        except Exception as e:
            print(f"[WARN] skip invalid meta {meta_path}: {e}")
    skills.sort(key=lambda x: x.get("id", ""))
    return skills


def resolve_screenshot(mid, rel):
    """把 meta.json 里的 screenshot 相对路径解析成页面可用 URL"""
    if not rel:
        return ""
    rel = rel.replace("\\", "/")
    if SCREENSHOT_MODE == "github":
        return f"{GITHUB_REPO_URL}/raw/{GITHUB_BRANCH}/skills/{mid}/{rel}"
    # copy 模式
    src = SKILLS_DIR / mid / rel
    if not src.exists():
        print(f"[WARN] screenshot not found: {src}")
        return ""
    dest = ASSETS_OUT / mid / Path(rel).name
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dest)
    return f"assets/{mid}/{Path(rel).name}"


def build_param_table(params):
    """生成弹窗内的参数表（真 HTML 表格，不是 markdown）"""
    if not params:
        return '<thead><tr><th>参数名</th><th>类型</th><th>是否必填</th><th>说明</th></tr></thead>' \
               '<tbody><tr><td colspan="4">无参数</td></tr></tbody>'
    head = "<thead><tr><th>参数名</th><th>类型</th><th>是否必填</th><th>说明</th></tr></thead>"
    rows = []
    for p in params:
        req = "✅" if p.get("required") else "❌"
        rows.append(
            f"<tr><td>{p.get('key','')}</td><td>{p.get('type','')}</td>"
            f"<td>{req}</td><td>{p.get('desc','')}</td></tr>"
        )
    return head + "<tbody>" + "".join(rows) + "</tbody>"


def calc_stats(skills):
    total = len(skills)
    return {
        "total": total,
        "active": sum(1 for s in skills if s.get("status") == "active"),
        "draft": sum(1 for s in skills if s.get("status") == "draft"),
        "deprecated": sum(1 for s in skills if s.get("status") == "deprecated"),
        "archived": sum(1 for s in skills if s.get("status") == "archived"),
    }


def render_html(skills):
    stats = calc_stats(skills)
    categories = sorted({s.get("category", "未分类") for s in skills})
    cat_options = "\n".join([f'<option value="{c}">{c}</option>' for c in categories])

    skill_cards = []
    modal_data_list = []

    for idx, skill in enumerate(skills):
        sid = skill.get("id", "")
        folder = skill.get("_folder", sid)
        name = skill.get("name", sid)
        category = skill.get("category", "未分类")
        ver = skill.get("version", "-")
        status = skill.get("status", "draft")
        desc = skill.get("description", "")
        tags = skill.get("tags", [])
        tag_html = " ".join([f'<span class="tag">{t}</span>' for t in tags])
        color = STATUS_COLOR_MAP.get(status, "#888888")
        github_readme_link = f"{GITHUB_REPO_URL}/blob/{GITHUB_BRANCH}/skills/{folder}/README.md"
        screenshot_url = resolve_screenshot(folder, skill.get("screenshot", ""))

        create_date = skill.get("create_date", "-")
        update_date = skill.get("update_date", "-")
        usage = skill.get("usage", "-")
        notes = skill.get("notes", "-")
        dependencies = skill.get("dependencies", [])
        dep_text = ", ".join(dependencies) if dependencies else "无"
        param_table = build_param_table(skill.get("params", []))

        img_html = ""
        if screenshot_url:
            img_html = f'<img class="card-thumb" src="{screenshot_url}" alt="preview" loading="lazy">'

        card = f"""
<div class="skill-card" data-category="{category}" data-name="{name.lower()}" data-id="{sid.lower()}" data-sid="{sid}" data-modal-idx="{idx}" draggable="true">
  {img_html}
  <div class="card-header">
    <span class="skill-id">{sid}</span>
    <span class="status-badge" style="background-color:{color}">{status}</span>
    <button class="star-btn" data-sid="{sid}" onclick="event.stopPropagation();toggleStar('{sid}')">☆</button>
  </div>
  <h3>{name}</h3>
  <div class="meta-row">
    <span class="cat">📂 {category}</span>
    <span class="ver">🔖 {ver}</span>
  </div>
  <p class="desc">{desc}</p>
  <div class="tags">
    {tag_html}
  </div>
</div>
"""
        skill_cards.append(card)

        modal_data_list.append({
            "sid": sid,
            "name": name,
            "category": category,
            "ver": ver,
            "status": status,
            "color": color,
            "create_date": create_date,
            "update_date": update_date,
            "desc": desc,
            "tags": tags,
            "dep_text": dep_text,
            "usage": usage,
            "notes": notes,
            "param_table": param_table,
            "github_readme": github_readme_link,
            "screenshot": screenshot_url,
        })

    card_html = "\n".join(skill_cards)
    modal_json_data = json.dumps(modal_data_list, ensure_ascii=False)

    html_template = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Agent Skill Toolbox - 技能工具箱</title>
<style>
:root {{
  --bg:#0f172a;
  --card-bg:#1e293b;
  --text:#f1f5f9;
  --text-secondary:#94a3b8;
  --border:#334155;
  --border-strong:#475569;
  --table-header:#334155;
}}
.light-theme {{
  --bg:#ffffff;
  --card-bg:#f8fafc;
  --text:#1e293b;
  --text-secondary:#64748b;
  --border:#cbd5e1;
  --border-strong:#94a3b8;
  --table-header:#e2e8f0;
}}
* {{
  box-sizing: border-box;
  margin: 0;
  padding: 0;
  font-family: system-ui, -apple-system, "Segoe UI", Roboto, sans-serif;
}}
body {{
  background-color: var(--bg);
  color: var(--text);
  padding: 32px max(16px, calc((100vw - 1200px)/2));
  transition: background 0.3s, color 0.3s;
}}
header {{
  margin-bottom: 24px;
}}
h1 {{
  font-size: 2rem;
  margin-bottom: 8px;
}}
.subtitle {{
  color:var(--text-secondary);
  margin-bottom:16px;
}}
.stats-bar {{
  display:flex;
  flex-wrap:wrap;
  gap:12px;
  margin-bottom:20px;
}}
.stat-item {{
  background:var(--card-bg);
  border:1px solid var(--border);
  padding:10px 14px;
  border-radius:8px;
  min-width:110px;
  text-align:center;
}}
.stat-num {{
  font-size:1.4rem;
  font-weight:bold;
}}
.stat-label {{
  font-size:0.8rem;
  color:var(--text-secondary);
}}
.controls {{
  display:flex;
  gap:12px;
  flex-wrap:wrap;
  margin-bottom:24px;
  align-items:center;
}}
#search {{
  flex:1;
  min-width:240px;
  padding:10px 14px;
  border-radius:8px;
  border:1px solid var(--border);
  background:var(--card-bg);
  color:var(--text);
  font-size:1rem;
}}
#cat-filter, #fav-filter, #theme-toggle, #export-btn, #fav-export, #fav-import, #fav-reset {{
  padding:10px 14px;
  border-radius:8px;
  border:1px solid var(--border);
  background:var(--card-bg);
  color:var(--text);
  cursor:pointer;
}}
.container {{
  display:grid;
  grid-template-columns: repeat(auto-fill, minmax(340px, 1fr));
  gap:18px;
}}
.skill-card {{
  background:var(--card-bg);
  border-radius:12px;
  padding:20px;
  border:1px solid var(--border);
  transition: transform 0.2s, border-color 0.2s;
  cursor:pointer;
  position:relative;
}}
.skill-card:hover {{
  transform: translateY(-4px);
  border-color:var(--border-strong);
}}
.skill-card[draggable=true] {{
  cursor: grab;
}}
.skill-card:active {{
  cursor: grabbing;
}}
.skill-card.dragging {{
  opacity:0.4;
}}
.skill-card.drag-over {{
  border: 2px dashed #60a5fa;
}}
.card-thumb {{
  width:100%;
  height:160px;
  object-fit:cover;
  border-radius:8px;
  margin-bottom:12px;
}}
.card-header {{
  display:flex;
  justify-content:space-between;
  align-items:center;
  gap:8px;
  margin-bottom:10px;
}}
.skill-id {{
  font-family: monospace;
  color:var(--text-secondary);
  font-size:0.9rem;
}}
.status-badge {{
  color:white;
  padding:3px 8px;
  border-radius:999px;
  font-size:0.75rem;
  font-weight:500;
}}
.star-btn {{
  background:transparent;
  border:none;
  font-size:18px;
  cursor:pointer;
  padding:2px;
}}
.star-btn.star-active {{
  filter: drop-shadow(0 0 4px gold);
}}
.skill-card h3 {{
  font-size:1.2rem;
  margin-bottom:10px;
}}
.meta-row {{
  display:flex;
  gap:12px;
  color:var(--text-secondary);
  font-size:0.9rem;
  margin-bottom:12px;
}}
.desc {{
  line-height:1.5;
  margin-bottom:14px;
}}
.tags {{
  display:flex;
  flex-wrap:wrap;
  gap:6px;
}}
.tag {{
  background:var(--border);
  color:var(--text);
  padding:3px 8px;
  border-radius:6px;
  font-size:0.75rem;
}}
.hidden {{
  display:none !important;
}}
.modal-overlay {{
  position:fixed;
  inset:0;
  background:rgba(0,0,0,0.75);
  display:none;
  align-items:center;
  justify-content:center;
  padding:16px;
  z-index:999;
}}
.modal-overlay.open {{
  display:flex;
}}
.modal {{
  background:var(--card-bg);
  border-radius:14px;
  border:1px solid var(--border-strong);
  width:100%;
  max-width:800px;
  max-height:85vh;
  overflow-y:auto;
  padding:24px;
}}
.modal-header {{
  display:flex;
  justify-content:space-between;
  align-items:flex-start;
  margin-bottom:16px;
  gap:16px;
}}
.modal-title h2 {{
  margin-bottom:4px;
}}
.modal-close {{
  background:none;
  border:none;
  color:var(--text-secondary);
  font-size:24px;
  cursor:pointer;
  padding:4px;
}}
.modal-close:hover {{
  color:var(--text);
}}
.modal-section {{
  margin:16px 0;
}}
.modal-section h4 {{
  color:#93c5fd;
  margin-bottom:8px;
  font-size:1rem;
}}
.light-theme .modal-section h4 {{
  color:#2563eb;
}}
.modal-table {{
  width:100%;
  border-collapse: collapse;
  font-size:0.9rem;
}}
.modal-table th,.modal-table td {{
  border:1px solid var(--border);
  padding:8px 10px;
  text-align:left;
}}
.modal-table th {{
  background:var(--table-header);
}}
.meta-small {{
  color:var(--text-secondary);
  font-size:0.9rem;
  line-height:1.6;
}}
a {{
  color:#60a5fa;
}}
.light-theme a {{
  color:#2563eb;
}}
.modal-img {{
  max-width:100%;
  border-radius:8px;
  cursor:zoom-in;
  margin-top:8px;
}}
.img-zoom-overlay {{
  position:fixed;
  inset:0;
  background:rgba(0,0,0,0.9);
  display:none;
  z-index:1000;
  align-items:center;
  justify-content:center;
  padding:20px;
}}
.img-zoom-overlay.open {{
  display:flex;
}}
.img-zoom {{
  max-width:95%;
  max-height:95%;
}}
footer {{
  margin-top:40px;
  color:var(--text-secondary);
  font-size:0.85rem;
  text-align:center;
}}
</style>
</head>
<body>
<header>
  <h1>🧰 Agent Skill Toolbox 技能工具箱</h1>
  <div class="subtitle">自制 Agent 技能仓库 · 技能卡片总览（点击卡片查看详情，⭐星标收藏，拖拽卡片调整顺序）</div>

  <div class="stats-bar">
    <div class="stat-item"><div class="stat-num">{stats['total']}</div><div class="stat-label">全部</div></div>
    <div class="stat-item"><div class="stat-num">{stats['active']}</div><div class="stat-label">✅ Active</div></div>
    <div class="stat-item"><div class="stat-num">{stats['draft']}</div><div class="stat-label">🟡 Draft</div></div>
    <div class="stat-item"><div class="stat-num">{stats['deprecated']}</div><div class="stat-label">🟠 Deprecated</div></div>
    <div class="stat-item"><div class="stat-num">{stats['archived']}</div><div class="stat-label">⚫ Archived</div></div>
    <div class="stat-item"><div class="stat-num" id="fav-count">0</div><div class="stat-label">⭐收藏</div></div>
  </div>

  <div class="controls">
    <input id="search" placeholder="搜索技能名称 / ID / 描述..." type="text">
    <select id="cat-filter">
      <option value="">全部分类</option>
      {cat_options}
    </select>
    <select id="fav-filter">
      <option value="">全部技能</option>
      <option value="fav">仅显示收藏⭐</option>
    </select>
    <button id="theme-toggle">🌓 切换主题</button>
    <button id="export-btn">💾 导出筛选JSON</button>
    <button id="fav-export">⭐导出收藏</button>
    <button id="fav-import">⭐导入收藏</button>
    <button id="fav-reset">⭐重置收藏</button>
  </div>
</header>
<div class="container" id="card-container">
{card_html}
</div>

<footer>共 {stats['total']} 个技能 · 页面由 scripts/build_static_html.py 自动生成，请勿手动修改 docs/index.html</footer>

<div class="modal-overlay" id="modalOverlay">
  <div class="modal" id="modalContent"></div>
</div>
<div class="img-zoom-overlay" id="imgZoomOverlay">
  <img id="zoomImg" class="img-zoom" alt="zoom">
</div>

<script>
const skillData = {modal_json_data};
const searchInput = document.getElementById('search');
const catSelect = document.getElementById('cat-filter');
const favSelect = document.getElementById('fav-filter');
const cardContainer = document.getElementById('card-container');
const overlay = document.getElementById('modalOverlay');
const modalContent = document.getElementById('modalContent');
const themeBtn = document.getElementById('theme-toggle');
const exportBtn = document.getElementById('export-btn');
const favExportBtn = document.getElementById('fav-export');
const favImportBtn = document.getElementById('fav-import');
const favResetBtn = document.getElementById('fav-reset');
const favCountEl = document.getElementById('fav-count');
const imgZoomOverlay = document.getElementById('imgZoomOverlay');
const zoomImg = document.getElementById('zoomImg');

// ========== ⭐ 收藏 ==========
let starredIds = JSON.parse(localStorage.getItem('skill_stars') || '[]');
function saveStars(){{
  localStorage.setItem('skill_stars', JSON.stringify(starredIds));
  renderStarUI();
}}
function toggleStar(sid){{
  const idx = starredIds.indexOf(sid);
  if(idx >= 0) starredIds.splice(idx,1); else starredIds.push(sid);
  saveStars();
  filterCards();
}}
function renderStarUI(){{
  favCountEl.innerText = starredIds.length;
  document.querySelectorAll('.star-btn').forEach(btn=>{{
    const sid = btn.dataset.sid;
    if(starredIds.includes(sid)){{
      btn.classList.add('star-active');
      btn.innerText = '⭐';
    }}else{{
      btn.classList.remove('star-active');
      btn.innerText = '☆';
    }}
  }});
}}
favExportBtn.onclick = ()=>{{
  const blob = new Blob([JSON.stringify(starredIds, null, 2)], {{type:'application/json'}});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = 'skill-favorites.json';
  a.click(); a.remove(); URL.revokeObjectURL(url);
}};
favImportBtn.onclick = ()=>{{
  const ipt = document.createElement('input');
  ipt.type = 'file'; ipt.accept = '.json';
  ipt.onchange = (ev)=>{{
    const f = ev.target.files[0];
    const r = new FileReader();
    r.onload = ()=>{{
      try{{
        const arr = JSON.parse(r.result);
        if(Array.isArray(arr)){{
          starredIds = arr; saveStars(); filterCards(); alert('导入成功');
        }}
      }}catch(e){{ alert('文件解析失败'); }}
    }};
    r.readAsText(f);
  }};
  ipt.click();
}};
favResetBtn.onclick = ()=>{{
  if(confirm('确认清空全部收藏？')){{ starredIds = []; saveStars(); filterCards(); }}
}};

// ========== 主题 ==========
const mediaPref = window.matchMedia('(prefers-color-scheme: light)');
function applyTheme(isLight){{
  document.body.classList.toggle('light-theme', isLight);
  localStorage.setItem('skilltoolbox_theme', isLight ? 'light' : 'dark');
}}
const storedTheme = localStorage.getItem('skilltoolbox_theme');
if(storedTheme === null){{ applyTheme(mediaPref.matches); }}
else{{ applyTheme(storedTheme === 'light'); }}
themeBtn.addEventListener('click', ()=>{{
  applyTheme(!document.body.classList.contains('light-theme'));
}});
mediaPref.addEventListener('change', (e)=>{{
  if(localStorage.getItem('skilltoolbox_theme') === null) applyTheme(e.matches);
}});

// ========== 拖拽排序 ==========
let draggedEl = null;
function getCardList(){{ return Array.from(cardContainer.querySelectorAll('.skill-card')); }}
function saveCardOrder(){{
  localStorage.setItem('skill_card_order', JSON.stringify(getCardList().map(c=>c.dataset.sid)));
}}
function loadCardOrder(){{
  const saved = localStorage.getItem('skill_card_order');
  if(!saved) return;
  JSON.parse(saved).forEach(sid=>{{
    const el = cardContainer.querySelector('.skill-card[data-sid="'+sid+'"]');
    if(el) cardContainer.appendChild(el);
  }});
}}
cardContainer.addEventListener('dragstart', (e)=>{{
  draggedEl = e.target;
  e.target.classList.add('dragging');
}});
cardContainer.addEventListener('dragend', (e)=>{{
  draggedEl = null;
  e.target.classList.remove('dragging');
  getCardList().forEach(c=>c.classList.remove('drag-over'));
  saveCardOrder();
}});
cardContainer.addEventListener('dragover', (e)=>{{
  if(!draggedEl) return;
  e.preventDefault();
  const afterEl = getDragAfterElement(cardContainer, e.clientY);
  if(afterEl == null) cardContainer.appendChild(draggedEl);
  else cardContainer.insertBefore(draggedEl, afterEl);
}});
function getDragAfterElement(container, y){{
  const draggable = [...container.querySelectorAll('.skill-card:not(.dragging)')];
  return draggable.reduce((closest, child)=>{{
    const box = child.getBoundingClientRect();
    const offset = y - box.top - box.height/2;
    if(offset < 0 && offset > closest.offset) return {{offset, element:child}};
    return closest;
  }}, {{offset: Number.NEGATIVE_INFINITY}}).element;
}}

// ========== 筛选 ==========
function filterCards(){{
  const kw = searchInput.value.trim().toLowerCase();
  const selectedCat = catSelect.value;
  const onlyFav = favSelect.value === 'fav';
  getCardList().forEach(card => {{
    const text = card.dataset.name + ' ' + card.dataset.id;
    const matchKw = text.includes(kw);
    const matchCat = !selectedCat || card.dataset.category === selectedCat;
    const matchFav = !onlyFav || starredIds.includes(card.dataset.sid);
    card.classList.toggle('hidden', !(matchKw && matchCat && matchFav));
  }});
}}
function getFilteredSkillData(){{
  return getCardList()
    .filter(c => !c.classList.contains('hidden'))
    .map(c => skillData[parseInt(c.dataset.modalIdx)]);
}}
exportBtn.addEventListener('click', ()=>{{
  const blob = new Blob([JSON.stringify(getFilteredSkillData(), null, 2)], {{type:'application/json'}});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = 'skill-toolbox-filtered.json';
  document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
}});

// ========== 弹窗与图片放大 ==========
function renderModal(item){{
  const starText = starredIds.includes(item.sid) ? '⭐已收藏' : '☆未收藏';
  let imgBlock = '';
  if(item.screenshot){{
    imgBlock = '<img class="modal-img" src="'+item.screenshot+'" onclick="openImgZoom(\\''+item.screenshot+'\\')">';
  }}
  return `
    <div class="modal-header">
      <div class="modal-title">
        <h2>${{item.name}}</h2>
        <div style="display:flex; gap:8px; align-items:center; margin-top:4px; flex-wrap:wrap;">
          <span class="skill-id">${{item.sid}}</span>
          <span class="status-badge" style="background-color:${{item.color}}">${{item.status}}</span>
          <span>${{starText}}</span>
        </div>
      </div>
      <button class="modal-close" id="closeModal">×</button>
    </div>
    ${{imgBlock}}
    <div class="meta-small" style="margin-top:12px;">
      📂 分类：${{item.category}} &nbsp;&nbsp; 🔖 版本：${{item.ver}}<br/>
      📅 创建：${{item.create_date}} &nbsp;&nbsp; 📆 更新：${{item.update_date}}<br/>
      🔗 <a target="_blank" href="${{item.github_readme}}">打开 GitHub 技能文档</a>
    </div>
    <div class="modal-section"><h4>简介</h4><p>${{item.desc}}</p></div>
    <div class="modal-section"><h4>标签</h4><div class="tags">${{item.tags.map(t=>'<span class="tag">'+t+'</span>').join('')}}</div></div>
    <div class="modal-section"><h4>参数列表</h4><table class="modal-table">${{item.param_table}}</table></div>
    <div class="modal-section"><h4>依赖</h4><p>${{item.dep_text}}</p></div>
    <div class="modal-section"><h4>使用方式</h4><p>${{item.usage}}</p></div>
    <div class="modal-section"><h4>注意事项</h4><p>${{item.notes}}</p></div>
  `;
}}
function openImgZoom(src){{
  zoomImg.src = src;
  imgZoomOverlay.classList.add('open');
}}
imgZoomOverlay.onclick = ()=> imgZoomOverlay.classList.remove('open');

getCardList().forEach(card => {{
  card.addEventListener('click', ()=>{{
    modalContent.innerHTML = renderModal(skillData[parseInt(card.dataset.modalIdx)]);
    overlay.classList.add('open');
    document.getElementById('closeModal').onclick = ()=> overlay.classList.remove('open');
  }});
}});
overlay.addEventListener('click', (e)=>{{ if(e.target === overlay) overlay.classList.remove('open'); }});
document.addEventListener('keydown', (e)=>{{
  if(e.key === 'Escape'){{ overlay.classList.remove('open'); imgZoomOverlay.classList.remove('open'); }}
}});
searchInput.addEventListener('input', filterCards);
catSelect.addEventListener('change', filterCards);
favSelect.addEventListener('change', filterCards);

// 初始化
renderStarUI();
loadCardOrder();
filterCards();
</script>
</body>
</html>
"""
    return html_template


def main():
    DOCS_DIR.mkdir(parents=True, exist_ok=True)
    skills = load_all_skills()
    html = render_html(skills)
    OUTPUT_HTML.write_text(html, encoding="utf-8")
    print(f"[OK] {len(skills)} skill(s) -> {OUTPUT_HTML}")
    print(f"     screenshot mode = {SCREENSHOT_MODE}")


if __name__ == "__main__":
    main()
