"""
OKR × 業務計画 統合管理アプリ  ─  app.py （ローカル保存版）
=============================================================
S3の代わりにローカルフォルダ（data/）にJSONを保存します。
AWSアカウント不要。

インストール:
    pip install streamlit pandas plotly python-pptx kaleido

.streamlit/secrets.toml（最小構成）:
    [app]
    team_name = "プロダクトチーム"
    members   = ["田中 一郎", "鈴木 花子", "佐藤 健二", "山田 美咲"]
    admin_pin = "1234"
    data_dir  = "data"

複数人で共有する場合:
    data_dir に OneDrive や社内共有フォルダのパスを指定してください。
    例）data_dir = "C:/Users/Shared/okr_data"

起動:
    streamlit run app.py

データ保存先:
    data/
    ├── master_config.json       ← 四半期OKR
    └── plans/
        └── 2025-06/
            └── 2025-06_田中_一郎.json
"""

from __future__ import annotations

import io
import json
import datetime
import traceback
from pathlib import Path
from typing import Any

import pandas as pd
import plotly.express as px
import streamlit as st
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt, Cm

# ══════════════════════════════════════════════════════════════════════════════
# 定数
# ══════════════════════════════════════════════════════════════════════════════

PALETTE: list[dict] = [
    {"main": "#1B4F72", "light": "#D6EAF8", "text": "#0D2137"},
    {"main": "#1D6A45", "light": "#D5F5E3", "text": "#0B3A21"},
    {"main": "#6C3483", "light": "#EDE0F5", "text": "#4A235A"},
    {"main": "#7E5109", "light": "#FAE5D3", "text": "#4D320A"},
]
KR_COLORS   = ["#1B4F72", "#1D6A45", "#6C3483"]
KR_LABELS   = ["KR①", "KR②", "KR③"]
MAX_ACTIONS = 5

DEFAULT_MASTER: dict = {
    "quarter":     "2025-Q2",
    "objective":   "",
    "locked":      False,
    "set_at":      "",
    "key_results": [
        {"id": "kr1", "label": "KR①", "text": ""},
        {"id": "kr2", "label": "KR②", "text": ""},
        {"id": "kr3", "label": "KR③", "text": ""},
    ],
}

# ══════════════════════════════════════════════════════════════════════════════
# ページ設定 & CSS
# ══════════════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="OKR × 業務計画",
    page_icon="🌟",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
[data-testid="stSidebar"]{
    background:linear-gradient(175deg,#0D1B2A 0%,#1B4F72 70%,#1D6A45 100%);
}
[data-testid="stSidebar"] *{color:#C8E0FF !important;}
[data-testid="stSidebar"] .stSelectbox>div>div{
    background:rgba(200,224,255,.12)!important;
    border:1px solid rgba(200,224,255,.3)!important;
}
.northstar{background:#1B4F72;border-radius:14px;padding:1.1rem 1.6rem;
    margin-bottom:1.1rem;position:relative;overflow:hidden;}
.ns-orb1{position:absolute;right:-40px;top:-40px;width:150px;height:150px;
    border-radius:50%;background:rgba(243,156,18,.18);}
.ns-orb2{position:absolute;left:58%;bottom:-55px;width:120px;height:120px;
    border-radius:50%;background:rgba(29,106,69,.28);}
.ns-lock{position:absolute;top:11px;right:13px;border-radius:20px;padding:2px 11px;font-size:.7rem;}
.ns-lock.locked{background:rgba(243,156,18,.26);border:1px solid rgba(243,156,18,.5);color:#F9E08B;}
.ns-lock.unlocked{background:rgba(255,255,255,.14);border:1px solid rgba(255,255,255,.25);color:rgba(255,255,255,.7);}
.ns-eyebrow{font-size:.68rem;font-weight:700;letter-spacing:.12em;text-transform:uppercase;
    color:#F9E08B;margin-bottom:.35rem;position:relative;}
.ns-obj{font-size:1.1rem;font-weight:700;color:#fff;line-height:1.4;
    margin-bottom:.55rem;position:relative;}
.ns-obj.empty{color:rgba(255,255,255,.45);font-style:italic;font-weight:400;font-size:.9rem;}
.ns-krs{display:flex;gap:7px;flex-wrap:wrap;position:relative;}
.ns-kr{background:rgba(255,255,255,.13);border:1px solid rgba(255,255,255,.22);
    border-radius:20px;padding:3px 12px;font-size:.74rem;color:#C8E0FF;}
.g-info{background:var(--color-background-info);border-left:3px solid var(--color-border-info);
    color:var(--color-text-info);border-radius:0 8px 8px 0;
    padding:.6rem .9rem;font-size:.8rem;line-height:1.55;margin:.4rem 0;}
.g-ok{background:#EAF3DE;border-left:3px solid #7dcea0;color:#1A5632;
    border-radius:0 8px 8px 0;padding:.55rem .9rem;font-size:.78rem;margin:.3rem 0;}
.g-ng{background:#FDEDEC;border-left:3px solid #f1948a;color:#7B241C;
    border-radius:0 8px 8px 0;padding:.55rem .9rem;font-size:.78rem;margin:.3rem 0;}
.g-warn{background:#FEF9E7;border-left:3px solid #F39C12;color:#7D6608;
    border-radius:0 8px 8px 0;padding:.6rem .9rem;font-size:.78rem;margin:.4rem 0;}
.prog-wrap{display:flex;align-items:center;margin-bottom:1.1rem;}
.prog-step{display:flex;align-items:center;gap:6px;padding:.42rem .9rem;
    border-radius:20px;font-size:.78rem;font-weight:500;white-space:nowrap;}
.prog-step.done{background:#EAF3DE;color:#1A5632;}
.prog-step.active{background:#1B4F72;color:#fff;}
.prog-step.todo{background:var(--color-background-secondary);color:var(--color-text-secondary);}
.prog-line{flex:1;height:2px;background:var(--color-border-tertiary);min-width:14px;}
.prog-line.filled{background:#1A5632;}
.ltree{background:var(--color-background-secondary);border:0.5px solid var(--color-border-tertiary);
    border-radius:var(--border-radius-lg);padding:1rem 1.15rem;}
.ltree-hdr{font-size:.68rem;font-weight:600;letter-spacing:.09em;text-transform:uppercase;
    color:var(--color-text-secondary);margin-bottom:.8rem;display:flex;align-items:center;gap:6px;}
.lt-row{display:flex;gap:8px;align-items:flex-start;margin-bottom:.5rem;}
.lt-ico{width:22px;height:22px;border-radius:5px;display:flex;align-items:center;
    justify-content:center;font-size:.65rem;font-weight:700;color:#fff;flex-shrink:0;margin-top:1px;}
.lt-body{font-size:.8rem;color:var(--color-text-primary);line-height:1.45;flex:1;}
.lt-sub{font-size:.7rem;color:var(--color-text-secondary);margin-top:2px;}
.lt-indent{margin-left:15px;padding-left:13px;border-left:2px solid var(--color-border-tertiary);}
.lt-empty{font-size:.75rem;color:var(--color-text-tertiary);font-style:italic;}
.check-list{display:flex;flex-direction:column;gap:5px;margin:.5rem 0;}
.check-item{display:flex;align-items:flex-start;gap:7px;font-size:.82rem;}
.member-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:1rem;}
.member-card{border-radius:10px;padding:.8rem;text-align:center;border:1.5px solid;}
.mc-av{width:36px;height:36px;border-radius:50%;display:flex;align-items:center;
    justify-content:center;font-size:.85rem;font-weight:600;margin:0 auto .4rem;}
.mc-name{font-size:.82rem;font-weight:600;margin-bottom:3px;}
.mc-badge{font-size:.7rem;padding:2px 8px;border-radius:20px;display:inline-block;font-weight:500;}
.stat-row{display:flex;gap:10px;margin-bottom:1.1rem;flex-wrap:wrap;}
.stat{background:var(--color-background-secondary);border-radius:var(--border-radius-md);
    padding:.7rem 1rem;text-align:center;flex:1;min-width:90px;}
.stat-n{font-size:1.55rem;font-weight:600;color:#1B4F72;}
.stat-l{font-size:.72rem;color:var(--color-text-secondary);margin-top:2px;}
.dash-card{background:var(--color-background-primary);border:0.5px solid var(--color-border-tertiary);
    border-radius:var(--border-radius-lg);padding:1rem 1.2rem;margin-bottom:.85rem;border-left:4px solid;}
.dash-hdr{display:flex;align-items:center;gap:10px;margin-bottom:.65rem;}
.dash-av{width:34px;height:34px;border-radius:50%;display:flex;align-items:center;
    justify-content:center;font-size:.8rem;font-weight:600;flex-shrink:0;}
.kr-block{background:var(--color-background-secondary);border-radius:var(--border-radius-md);
    padding:.65rem .85rem;margin-bottom:.5rem;}
.action-list-item{display:flex;gap:7px;align-items:flex-start;padding:.28rem 0;
    border-bottom:0.5px solid var(--color-border-tertiary);font-size:.78rem;}
.action-list-item:last-child{border-bottom:none;}
.a-num{width:18px;height:18px;border-radius:50%;display:flex;align-items:center;
    justify-content:center;font-size:.62rem;font-weight:600;color:#fff;flex-shrink:0;margin-top:1px;}
.fb-panel{background:#FEF9E7;border:0.5px solid #F9E79F;border-radius:var(--border-radius-lg);
    padding:1rem 1.2rem;margin-bottom:1rem;}
.fb-panel h4{font-size:.88rem;font-weight:600;color:#7D6608;margin-bottom:.55rem;}
.fb-item{display:flex;gap:8px;padding:.3rem 0;border-bottom:0.5px solid rgba(249,231,159,.7);
    font-size:.8rem;color:#6E4D0B;align-items:flex-start;}
.fb-item:last-child{border-bottom:none;}
.fb-ico{font-size:14px;flex-shrink:0;margin-top:1px;}
.flow-row{display:flex;align-items:stretch;gap:0;margin-bottom:.9rem;flex-wrap:nowrap;}
.flow-box{flex:1;border-radius:10px;padding:.85rem .9rem;text-align:center;min-width:0;}
.flow-arr{display:flex;align-items:center;padding:0 5px;
    color:var(--color-text-secondary);font-size:1rem;flex-shrink:0;}
.local-badge{background:#E8F4F8;border:0.5px solid #85B7EB;border-radius:7px;
    padding:5px 11px;font-size:.8rem;color:#0C447C;margin-bottom:.6rem;}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# 設定読み込み
# ══════════════════════════════════════════════════════════════════════════════

@st.cache_data(ttl=0)
def load_config() -> dict:
    try:
        members   = list(st.secrets["app"]["members"])
        team_name = str(st.secrets["app"].get("team_name", "チーム"))
        admin_pin = str(st.secrets["app"].get("admin_pin", "1234"))
        data_dir  = str(st.secrets["app"].get("data_dir", "data"))
    except Exception:
        members   = ["田中 一郎", "鈴木 花子", "佐藤 健二", "山田 美咲"]
        team_name = "プロダクトチーム"
        admin_pin = "1234"
        data_dir  = "data"
    return dict(members=members, team_name=team_name,
                admin_pin=admin_pin, data_dir=data_dir)


CFG     = load_config()
MEMBERS = CFG["members"]

# app.py と同じ場所に data/ フォルダを作成
BASE_DIR = Path(__file__).parent / CFG["data_dir"]


# ══════════════════════════════════════════════════════════════════════════════
# ローカルファイル I/O
# ══════════════════════════════════════════════════════════════════════════════

def _read_json(path: Path) -> dict | None:
    try:
        return json.loads(path.read_text(encoding="utf-8")) if path.exists() else None
    except Exception:
        return None


def _write_json(path: Path, data: dict) -> bool:
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        return True
    except Exception as e:
        st.error(f"ファイル保存エラー: {e}")
        return False


def master_path() -> Path:
    return BASE_DIR / "master_config.json"


def plan_path(month_str: str, member: str) -> Path:
    safe = member.replace(" ", "_").replace("　", "_")
    return BASE_DIR / "plans" / month_str / f"{month_str}_{safe}.json"


def io_get_master() -> dict:
    return _read_json(master_path()) or DEFAULT_MASTER.copy()


def io_save_master(data: dict) -> bool:
    return _write_json(master_path(), data)


def io_get_plan(month_str: str, member: str) -> dict | None:
    return _read_json(plan_path(month_str, member))


def io_save_plan(month_str: str, member: str, data: dict) -> bool:
    return _write_json(plan_path(month_str, member), data)


def io_list_plans(month_str: str) -> list[dict]:
    plans_dir = BASE_DIR / "plans" / month_str
    if not plans_dir.exists():
        return []
    return [d for f in sorted(plans_dir.glob("*.json"))
            if (d := _read_json(f)) is not None]


# ══════════════════════════════════════════════════════════════════════════════
# ヘルパー
# ══════════════════════════════════════════════════════════════════════════════

def mpal(name: str) -> dict:
    idx = MEMBERS.index(name) if name in MEMBERS else 0
    return PALETTE[idx % len(PALETTE)]


def initials(name: str) -> str:
    parts = name.split()
    return parts[0][0] + (parts[1][0] if len(parts) > 1 else "")


def blank_action() -> dict:
    today = datetime.date.today()
    return {"text": "", "start": today.isoformat(),
            "end": (today + datetime.timedelta(days=30)).isoformat()}


def blank_issue() -> dict:
    """空の壁（課題）を1件返す。アクションを1件含む。"""
    return {"text": "", "actions": [blank_action()]}


def _init_session():
    today = datetime.date.today()
    for k, v in dict(
        cur_member  = MEMBERS[0],
        month_str   = today.strftime("%Y-%m"),
        plan_step   = 0,
        plan_kr_idx = 0,
        plan_issues = [blank_issue()],   # [{text, actions:[{text,start,end}]}]
        admin_auth  = False,
        team_data   = None,
    ).items():
        if k not in st.session_state:
            st.session_state[k] = v


# ══════════════════════════════════════════════════════════════════════════════
# 北極星バナー
# ══════════════════════════════════════════════════════════════════════════════

def render_north_star(master: dict):
    locked   = master.get("locked", False)
    obj      = master.get("objective", "").strip()
    krs      = master.get("key_results", [])
    qtr      = master.get("quarter", "")
    lock_cls = "locked" if locked else "unlocked"
    lock_lbl = f"🔒 {qtr} 確定済み" if locked else f"📝 {qtr} 未確定"
    kr_html  = " ".join(
        f'<span class="ns-kr">{kr["label"]}：{kr["text"]}</span>'
        for kr in krs if kr.get("text")
    ) if obj else ""
    obj_html = (
        f'<div class="ns-obj">{obj}</div>' if obj
        else '<div class="ns-obj empty">まだ設定されていません ─ STRATEGYタブで入力してください</div>'
    )
    st.markdown(f"""
<div class="northstar">
  <div class="ns-orb1"></div><div class="ns-orb2"></div>
  <span class="ns-lock {lock_cls}">{lock_lbl}</span>
  <div class="ns-eyebrow">🌟 今期の北極星 Objective</div>
  {obj_html}
  <div class="ns-krs">{kr_html}</div>
</div>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# プログレスバー
# ══════════════════════════════════════════════════════════════════════════════

def render_progress(step: int):
    steps = ["🎯 KRを選ぶ", "🔍 壁とアクションを決める"]
    parts = []
    for i, label in enumerate(steps):
        cls = "done" if i < step else ("active" if i == step else "todo")
        parts.append(f'<div class="prog-step {cls}">{label}</div>')
        if i < len(steps) - 1:
            parts.append(f'<div class="prog-line {"filled" if i < step else ""}"></div>')
    st.markdown(f'<div class="prog-wrap">{"".join(parts)}</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# ロジックツリー
# ══════════════════════════════════════════════════════════════════════════════

def render_logic_tree(master: dict, kr_idx: int,
                      issues: list[dict], pal: dict):
    """issues = [{text, actions:[{text,start,end}]}]"""
    krs = master.get("key_results", [])
    kr  = krs[kr_idx] if kr_idx < len(krs) else {}

    filled_issues = [iss for iss in issues if iss.get("text","").strip()]
    if filled_issues:
        issues_html = ""
        for ii, iss in enumerate(filled_issues):
            filled_actions = [a for a in iss.get("actions",[]) if a.get("text","").strip()]
            actions_html = "".join(
                f'<div class="lt-indent" style="margin-top:.3rem;">'
                f'<div class="lt-row">'
                f'<div class="lt-ico" style="background:{pal["main"]};">A{ia+1}</div>'
                f'<div class="lt-body"><span style="font-weight:600;">{a["text"]}</span>'
                f'<div class="lt-sub">{a.get("start","")} → {a.get("end","")}</div>'
                f'</div></div></div>'
                for ia, a in enumerate(filled_actions)
            ) if filled_actions else (
                f'<div class="lt-indent" style="margin-top:.3rem;">'
                f'<div class="lt-row"><div class="lt-ico" style="background:{pal["main"]};">A</div>'
                f'<div class="lt-body lt-empty">アクションを入力してください</div>'
                f'</div></div>'
            )
            issues_html += (
                f'<div class="lt-row" style="margin-top:.35rem;">'
                f'<div class="lt-ico" style="background:#F39C12;">壁{ii+1}</div>'
                f'<div class="lt-body" style="font-weight:600;">{iss["text"]}</div>'
                f'</div>{actions_html}'
            )
    else:
        issues_html = (
            '<div class="lt-row" style="margin-top:.35rem;">'
            '<div class="lt-ico" style="background:#F39C12;">壁</div>'
            '<div class="lt-body lt-empty">壁を入力すると表示されます</div>'
            '</div>'
        )

    st.markdown(f"""
<div class="ltree">
  <div class="ltree-hdr">🌲 ロジックツリー（リアルタイム）</div>
  <div class="lt-row">
    <div class="lt-ico" style="background:#1B4F72;">KR</div>
    <div class="lt-body" style="font-weight:700;">{kr.get("label","KR")}：{kr.get("text","")}</div>
  </div>
  <div class="lt-indent">{issues_html}</div>
</div>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# HOME タブ
# ══════════════════════════════════════════════════════════════════════════════

def render_home():
    st.markdown("""
<div class="flow-row">
  <div class="flow-box" style="background:#EBF5FB;border:1.5px solid #1B4F72;">
    <div style="font-size:1.35rem;margin-bottom:.25rem">🌟</div>
    <div style="font-size:.7rem;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:#1B4F72">Objective</div>
    <div style="font-size:.78rem;margin-top:.2rem;line-height:1.4;color:var(--color-text-secondary)">ワクワクする<br>定性的な<b>目的地</b></div>
  </div>
  <div class="flow-arr">→</div>
  <div class="flow-box" style="background:#EAF4FB;border:1.5px solid #2E86C1;">
    <div style="font-size:1.35rem;margin-bottom:.25rem">📏</div>
    <div style="font-size:.7rem;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:#2E86C1">Key Results</div>
    <div style="font-size:.78rem;margin-top:.2rem;line-height:1.4;color:var(--color-text-secondary)">達成を測る<br>定量的な<b>指標</b>（3つ）</div>
  </div>
  <div class="flow-arr">→</div>
  <div class="flow-box" style="background:#FEF9E7;border:1.5px solid #F39C12;">
    <div style="font-size:1.35rem;margin-bottom:.25rem">🔍</div>
    <div style="font-size:.7rem;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:#B7770D">Issue（課題）</div>
    <div style="font-size:.78rem;margin-top:.2rem;line-height:1.4;color:var(--color-text-secondary)">KRに届かない<br><b>真の原因</b></div>
  </div>
  <div class="flow-arr">→</div>
  <div class="flow-box" style="background:#EAFAF1;border:1.5px solid #1D6A45;">
    <div style="font-size:1.35rem;margin-bottom:.25rem">⚡</div>
    <div style="font-size:.7rem;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:#1D6A45">Action（行動）</div>
    <div style="font-size:.78rem;margin-top:.2rem;line-height:1.4;color:var(--color-text-secondary)">今月<b>やり切る</b><br>具体的な一手（複数可）</div>
  </div>
</div>
""", unsafe_allow_html=True)
    st.markdown('<div class="g-info"><b>なぜこの順番？</b>　KR → 壁（課題） → アクション の順に強制的に考えさせる設計。各KRに壁を最大3件、各壁にアクションを最大5件追加できます。</div>', unsafe_allow_html=True)

    # ── 操作ガイドボタン ──────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### 📖 操作ガイド")
    gc1, gc2 = st.columns(2)
    with gc1:
        with st.popover("👤 メンバー向けガイドを見る", use_container_width=True):
            st.markdown("## 👤 メンバー向け 操作ガイド")
            st.markdown("---")
            st.markdown("### 1. アクセス方法")
            st.markdown("ブラウザでアプリのURLを開くだけです。インストール不要。左のサイドバーで **自分の名前** を必ず選んでください。")
            st.warning("⚠️ 名前を間違えると他の人のデータを上書きしてしまいます。")

            st.markdown("---")
            st.markdown("### 2. 月次計画の入力（PLANタブ）")
            st.markdown("「📝 PLAN」タブを開き、2ステップで入力します。")

            st.markdown("#### STEP 1：KRを選ぶ")
            st.markdown("今月最も重要なKRを1つ選んで「選ぶ」ボタンを押してください。")

            st.markdown("#### STEP 2：壁とアクションを決める")
            st.markdown("""
**壁（課題）を入力する**
「なぜこのKRに今届いていないのか？」の原因を1文で書いてください。
""")
            st.markdown('<div class="g-ok">✅ 「インタビューがゼロのため施策が推測止まりになっている」</div>', unsafe_allow_html=True)
            st.markdown('<div class="g-ng">❌ 「NPSが低い」← 現象ではなく原因を書く</div>', unsafe_allow_html=True)

            st.markdown("""
**アクションを入力する**
各壁の下に「いつまでに・何を・どれだけ」を含む行動を入力してください。
""")
            st.markdown('<div class="g-ok">✅ 「6/15までにインタビュー10件実施しSlackで共有する」</div>', unsafe_allow_html=True)
            st.markdown('<div class="g-ng">❌ 「顧客の声を聞く」← いつ？何件？</div>', unsafe_allow_html=True)

            st.markdown("""
**壁・アクションの追加・削除**
- 「＋ 壁（課題）を追加する」で壁を最大3件まで追加できます
- 「＋ アクションを追加」で各壁に最大5件のアクションを追加できます
- 🗑 ボタンで不要な壁・アクションを削除できます
""")
            st.markdown("入力が終わったら「💾 保存する」を押してください。")

            st.markdown("---")
            st.markdown("### 3. 複数のKRに入力したい場合")
            st.markdown("「別のKRも入力する」ボタンでKR選択画面に戻ります。入力済みのデータは保持されます。")

            st.markdown("---")
            st.markdown("### 4. よくある質問")
            st.markdown("""
**Q. 保存前にブラウザを閉じてしまったら？**
入力内容は消えます。こまめに「💾 保存する」を押しましょう。

**Q. 入力を修正したい**
再度PLANタブを開くと前回のデータが読み込まれます。修正して保存すると上書きされます。

**Q. 入力期限は？**
毎月1〜3日を目安に入力してください。
""")

    with gc2:
        with st.popover("👔 マネジャー向けガイドを見る", use_container_width=True):
            st.markdown("## 👔 マネジャー向け 操作ガイド")
            st.markdown("---")
            st.markdown("### 1. 月次運用カレンダー")
            st.markdown("""
| タイミング | 作業 | 使う機能 |
|-----------|------|---------|
| 四半期初め | OKRを確定する | STRATEGYタブ |
| 月初 1〜3日 | メンバーの入力を促す | （メンバーがPLANタブで入力） |
| 月初 MTG当日 | 全員の計画を確認 | DASHBOARDタブ |
| 月初 MTG当日 | PPTX資料を生成 | DASHBOARDタブ → PPTX出力 |
| 月末 | 振り返りを実施 | DASHBOARDタブ |
""")
            st.markdown("---")
            st.markdown("### 2. 四半期OKRの設定（STRATEGYタブ）")
            st.markdown("""
1. 「🏛️ STRATEGY」タブを開き、管理者PINを入力して認証
2. **Objective**：チームがワクワクする定性的な目標を入力（右側のチェックリストが全✅になるのが理想）
3. **Key Results**：月末に○か✕か判定できる数値指標を3つ入力（70〜80%達成が理想的な難易度）
4. 「🔒 OKRを確定保存する」を押すとロック → 北極星バナーに反映
""")
            st.warning("⚠️ 確定後は四半期中は編集不可。変更は「ロック解除」ボタン＋チーム全員の合意が必要です。")

            st.markdown("---")
            st.markdown("### 3. ダッシュボードの活用（DASHBOARDタブ）")
            st.markdown("""
- 「🔄 データを読み込む」で全員の最新データを取得
- 提出状況カードで未提出メンバーを確認
- 「マネジャー向け レビューの着眼点」を展開してフィードバック観点を確認
""")
            st.markdown("**フィードバックの5観点**")
            st.markdown("""
| 観点 | チェックポイント |
|-----|----------------|
| 🔥 野心度 | KRは70〜80%達成が理想の難易度か？ |
| 🔗 ロジック | 壁→アクションの因果関係は通っているか？ |
| ⚡ リソース | 全アクション合計で1ヶ月に現実的な量か？ |
| 🤝 連携 | 他メンバーとの重複・依存はないか？ |
| 🎯 整合性 | 全員がObjectiveを向いているか？ |
""")
            st.markdown("---")
            st.markdown("### 4. PPTX資料の生成")
            st.markdown("""
1. DASHBOARDタブで「🔄 データを読み込む」
2. 「🚀 PPTXを生成する」を押す
3. 数秒後に「⬇️ PPTXをダウンロード」が表示される

**構成：** 表紙 → チームOKRサマリー → メンバー別詳細 → 統合ガントチャート
""")

    data_path = BASE_DIR.resolve()
    st.markdown("---")
    st.markdown("### 📁 データの保存場所")
    st.markdown(f'<div class="local-badge">💾 保存先：<code>{data_path}</code></div>', unsafe_allow_html=True)
    st.markdown('<div class="g-info"><b>複数人で共有する場合</b>　<code>data</code>フォルダをOneDriveや社内共有フォルダに移動し、<code>secrets.toml</code>の<code>data_dir</code>にそのパスを指定してください。<br>例）<code>data_dir = "C:/Users/Shared/okr_data"</code></div>', unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### 月次の運用サイクル")
    cols = st.columns(4)
    for col, (n, title, desc) in zip(cols, [
        (1, "四半期初め",  "管理者がSTRATEGYでOKRを確定。"),
        (2, "月初 1〜3日", "各メンバーがPLANタブで計画を入力・保存。"),
        (3, "月初 MTG",   "DASHBOARDで全員の計画を確認・PPTX生成。"),
        (4, "月末",        "振り返りを行い、翌月の計画に活かす。"),
    ]):
        with col:
            st.markdown(
                f'<div style="background:var(--color-background-secondary);border-radius:10px;padding:.8rem;text-align:center;">'
                f'<div style="font-size:1.5rem;font-weight:600;color:#1B4F72;">{n}</div>'
                f'<div style="font-size:.82rem;font-weight:600;color:var(--color-text-primary);">{title}</div>'
                f'<div style="font-size:.75rem;color:var(--color-text-secondary);margin-top:3px;line-height:1.45;">{desc}</div></div>',
                unsafe_allow_html=True,
            )
    st.markdown("---")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("**🌟 良いObjectiveとは**")
        st.markdown('<div class="g-ok">✅ 「顧客が思わず人に紹介したくなるサービスにする」</div>', unsafe_allow_html=True)
        st.markdown('<div class="g-ng">❌ 「売上を上げる」← 数値はKRへ</div>', unsafe_allow_html=True)
    with c2:
        st.markdown("**📏 良いKey Resultとは**")
        st.markdown('<div class="g-ok">✅ 「月次NPS 40以上」「解約率 1%以下」</div>', unsafe_allow_html=True)
        st.markdown('<div class="g-ng">❌ 「顧客満足度を高める」← 測れない</div>', unsafe_allow_html=True)
    with c3:
        st.markdown("**⚡ 良いActionとは**")
        st.markdown('<div class="g-ok">✅ 「6/15までにインタビュー10件実施」</div>', unsafe_allow_html=True)
        st.markdown('<div class="g-ng">❌ 「顧客の声を聞く」← いつ？何件？</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# STRATEGY タブ
# ══════════════════════════════════════════════════════════════════════════════

def render_strategy(master: dict):
    if not st.session_state.admin_auth:
        st.markdown("### 🔐 管理者モードの認証")
        st.markdown('<div class="g-info"><b>このタブはチームの四半期OKRを確定する専用画面です。</b><br>全員で画面を共有しながら入力し、合意の上で「確定保存」してください。</div>', unsafe_allow_html=True)
        with st.form("admin_login_form"):
            pin = st.text_input("管理者PINコードを入力", type="password", key="admin_pin_input")
            if st.form_submit_button("認証する", use_container_width=True):
                if pin == CFG["admin_pin"] or not pin:
                    st.session_state.admin_auth = True
                    st.rerun()
                else:
                    st.error("PINコードが違います。")
        return

    locked = master.get("locked", False)
    if locked:
        st.success(f"✅ このOKRは **{master.get('quarter','')}** の期間中、確定済みです。")
        st.markdown('<div class="g-warn">🔒 確定済みのOKRは四半期中は編集できません。変更が必要な場合はチーム全員で合意し、下のボタンでロックを解除してください。</div>', unsafe_allow_html=True)
        with st.expander("確定済みOKRの内容を確認"):
            st.markdown(f"**Objective：** {master.get('objective','')}")
            for kr in master.get("key_results", []):
                st.markdown(f"- **{kr['label']}：** {kr.get('text','')}")
        if st.button("🔓 ロックを解除して編集する（要チーム合意）", type="secondary"):
            master["locked"] = False
            io_save_master(master)
            st.toast("ロックを解除しました。", icon="🔓")
            st.rerun()
        return

    st.markdown('<div class="g-info">💡 チーム全員で画面を見ながら話し合いましょう。「ワクワクするか？」「数値で測れるか？」を声に出しながら入力してください。</div>', unsafe_allow_html=True)

    with st.container(border=True):
        st.markdown("**対象四半期**")
        quarter = st.text_input("四半期", value=master.get("quarter","2025-Q2"), label_visibility="collapsed")

    with st.container(border=True):
        st.markdown("**🌟 Objective（チームの目的地）**　─　定性的・感情的な言葉で")
        col_in, col_check = st.columns([3, 2])
        with col_in:
            objective = st.text_area("Objective",
                value=master.get("objective",""), height=100,
                placeholder="例）顧客が思わず人に紹介したくなるプロダクトを作る",
                help="数値はNG！「〜する」「〜になる」という完成形で。",
                label_visibility="collapsed")
        with col_check:
            obj_v = objective.strip()
            checks = [
                ("読んだ時にワクワクする",       len(obj_v) > 8),
                ("数字が含まれていない",          not any(c.isdigit() for c in obj_v)),
                ("「〜する/になる」の完成形",     obj_v.endswith(("する","なる","きる","げる","れる","つ","む","ぐ"))),
                ("15〜70文字程度",               15 <= len(obj_v) <= 70),
            ]
            st.markdown('<div class="check-list">', unsafe_allow_html=True)
            for label, ok in checks:
                color = "#196F3D" if ok else "var(--color-text-secondary)"
                st.markdown(f'<div class="check-item" style="color:{color};">{"✅" if ok else "⬜"} {label}</div>', unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)

    with st.container(border=True):
        st.markdown("**📏 Key Results（達成を測る3つの指標）**　─　月末に○か✕かが判定できる数値で")
        st.markdown('<div class="g-info">KRは「Objectiveが達成された証拠」です。70〜80%達成が理想的な難易度です。</div>', unsafe_allow_html=True)
        prev_krs = master.get("key_results", [{"id": f"kr{i+1}", "label": KR_LABELS[i], "text": ""} for i in range(3)])
        while len(prev_krs) < 3:
            i = len(prev_krs)
            prev_krs.append({"id": f"kr{i+1}", "label": KR_LABELS[i], "text": ""})
        kr_texts = []
        for i, kr in enumerate(prev_krs[:3]):
            cc, ci = st.columns([0.12, 0.88])
            with cc:
                st.markdown(f'<div style="background:{KR_COLORS[i]};color:#fff;font-size:.75rem;font-weight:600;border-radius:20px;padding:5px 8px;text-align:center;margin-top:6px;">{KR_LABELS[i]}</div>', unsafe_allow_html=True)
            with ci:
                val = st.text_input(f"KR{i+1}", value=kr.get("text",""),
                    placeholder="例）月次NPSを 40 以上に引き上げる",
                    key=f"strategy_kr_{i}", label_visibility="collapsed")
            kr_texts.append(val)

    st.markdown("---")
    st.markdown('<div class="g-warn">⚠️ 「確定保存」を押すとOKRがロックされます。全員が納得した上で押してください。</div>', unsafe_allow_html=True)
    col_save, col_hint = st.columns([2, 5])
    with col_save:
        save_btn = st.button("🔒 OKRを確定保存する", type="primary", use_container_width=True)
    with col_hint:
        filled = sum(1 for t in kr_texts if t.strip())
        st.caption(f"Objective：{'✅ 入力済' if objective.strip() else '⬜ 未入力'}　KR入力済：{filled}/3")

    if save_btn:
        if not objective.strip():
            st.error("Objectiveを入力してください。")
        elif not any(t.strip() for t in kr_texts):
            st.error("KRを1つ以上入力してください。")
        else:
            payload = dict(
                quarter    = quarter.strip(),
                objective  = objective.strip(),
                locked     = True,
                set_at     = datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
                key_results= [
                    {"id": f"kr{i+1}", "label": KR_LABELS[i], "text": t.strip()}
                    for i, t in enumerate(kr_texts) if t.strip()
                ],
            )
            if io_save_master(payload):
                st.toast("🌟 OKRを確定しました！北極星が輝きました。", icon="🌟")
                st.success("✅ 確定保存しました。ページを更新すると北極星バナーに反映されます。")


# ══════════════════════════════════════════════════════════════════════════════
# PLAN タブ
# ══════════════════════════════════════════════════════════════════════════════

def render_plan(master: dict):
    member    = st.session_state.cur_member
    month_str = st.session_state.month_str
    pal       = mpal(member)
    krs       = master.get("key_results", [])
    step      = st.session_state.plan_step
    kr_idx    = st.session_state.plan_kr_idx

    if not krs or not any(kr.get("text") for kr in krs):
        st.warning("OKRがまだ設定されていません。先に「STRATEGY」タブで四半期OKRを確定してください。")
        return

    draft_key = f"draft_{month_str}_{member}"
    if draft_key not in st.session_state:
        saved = io_get_plan(month_str, member)
        if saved:
            # 旧フォーマット（issue + actions）を新フォーマット（issues）に変換
            loaded = {}
            for item in saved.get("items", []):
                if "issues" in item:
                    loaded[item["kr_id"]] = {"issues": item["issues"]}
                else:
                    # 旧データ互換
                    loaded[item["kr_id"]] = {"issues": [
                        {"text": item.get("issue",""), "actions": item.get("actions",[blank_action()])}
                    ]}
            st.session_state[draft_key] = loaded
            st.success(f"✅ 前回保存（{saved.get('saved_at','')}）を読み込みました。")
        else:
            st.session_state[draft_key] = {}
            st.info(f"📋 {member} さんの今月のデータはまだありません。順番に入力していきましょう！")

    draft: dict = st.session_state[draft_key]
    render_progress(step)

    # ── STEP 0: KR選択 ────────────────────────────────────────────────────
    if step == 0:
        st.markdown('<div class="g-info">「今月最も重要なKRはどれですか？」全部やろうとせず、今月最も差が生まれるKRを選びましょう。</div>', unsafe_allow_html=True)
        st.markdown("")
        for i, kr in enumerate(krs):
            kd     = draft.get(kr["id"], {})
            issues = kd.get("issues", [])
            done   = any(
                iss.get("text","").strip() and
                any(a.get("text","").strip() for a in iss.get("actions",[]))
                for iss in issues
            )
            col_text, col_btn = st.columns([6, 1])
            with col_text:
                st.markdown(
                    f'<div style="background:{KR_COLORS[i]}18;border:1.5px solid {KR_COLORS[i]};'
                    f'border-radius:10px;padding:.65rem 1rem;margin-bottom:.4rem;">'
                    f'<span style="background:{KR_COLORS[i]};color:#fff;font-size:.7rem;font-weight:700;padding:2px 8px;border-radius:20px;">{kr["label"]}</span>'
                    f'<span style="font-size:.88rem;font-weight:500;color:{KR_COLORS[i]};margin-left:8px;">{kr["text"]}</span>'
                    f'<span style="float:right;font-size:.72rem;color:{"#196F3D" if done else "#7F8C8D"};">{"✅ 入力済" if done else "⬜ 未入力"}</span>'
                    f'</div>', unsafe_allow_html=True,
                )
            with col_btn:
                if st.button("選ぶ", key=f"sel_kr_{i}", use_container_width=True):
                    st.session_state.plan_kr_idx = i
                    st.session_state.plan_issues = kd.get("issues", [blank_issue()])
                    st.session_state.plan_step   = 1
                    st.rerun()

    # ── STEP 1: 壁（課題）＋アクション入力 ───────────────────────────────
    elif step == 1:
        kr     = krs[kr_idx]
        issues = st.session_state.plan_issues   # [{text, actions:[{text,start,end}]}]

        col_form, col_tree = st.columns([3, 2], gap="medium")

        with col_form:
            st.markdown(
                f'<div style="background:{KR_COLORS[kr_idx]}18;border-left:4px solid {KR_COLORS[kr_idx]};'
                f'padding:.5rem .85rem;border-radius:0 8px 8px 0;font-weight:600;font-size:.85rem;'
                f'margin-bottom:.85rem;color:{KR_COLORS[kr_idx]};">{kr["label"]}：{kr["text"]}</div>',
                unsafe_allow_html=True,
            )
            st.markdown('<div class="g-info">KRに届かない<b>壁（課題）</b>を洗い出し、それぞれに対するアクションを設定してください。壁は最大3件まで追加できます。</div>', unsafe_allow_html=True)

            issues_to_delete = []
            for ii, iss in enumerate(issues):
                with st.container(border=True):
                    # 壁ヘッダー
                    col_hdr, col_del = st.columns([8, 1])
                    with col_hdr:
                        st.markdown(
                            f'<div style="display:flex;align-items:center;gap:7px;margin-bottom:.3rem;">'
                            f'<div style="width:22px;height:22px;border-radius:50%;background:#F39C12;'
                            f'display:flex;align-items:center;justify-content:center;font-size:.7rem;font-weight:600;color:#fff;">壁{ii+1}</div>'
                            f'<span style="font-size:.85rem;font-weight:600;color:var(--color-text-primary);">壁（課題） {ii+1}</span></div>',
                            unsafe_allow_html=True,
                        )
                    with col_del:
                        if st.button("🗑", key=f"del_issue_{ii}", help="この壁を削除"):
                            issues_to_delete.append(ii)

                    issues[ii]["text"] = st.text_area(
                        f"issue_{ii}",
                        value=iss.get("text",""),
                        height=75,
                        placeholder="例）提案資料の訴求力が弱く、顧客の意思決定の後押しができていない",
                        label_visibility="collapsed",
                        key=f"issue_txt_{ii}",
                    )

                    # アクション行
                    st.markdown(
                        '<div style="font-size:10px;color:var(--color-text-secondary);'
                        'margin:.5rem 0 .3rem;font-weight:500;padding-top:.5rem;'
                        'border-top:0.5px solid var(--color-border-tertiary);">⚡ アクション</div>',
                        unsafe_allow_html=True,
                    )

                    actions = iss.get("actions", [blank_action()])
                    issues[ii]["actions"] = actions
                    actions_to_delete = []

                    for ia, action in enumerate(actions):
                        with st.container(border=False):
                            col_ahdr, col_adel = st.columns([8, 1])
                            with col_ahdr:
                                st.markdown(
                                    f'<div style="display:flex;align-items:center;gap:6px;'
                                    f'background:var(--color-background-secondary);border-radius:6px;padding:.3rem .5rem;">'
                                    f'<div style="width:18px;height:18px;border-radius:50%;background:{pal["main"]};'
                                    f'display:flex;align-items:center;justify-content:center;font-size:.62rem;font-weight:600;color:#fff;">A{ia+1}</div>'
                                    f'<span style="font-size:.8rem;font-weight:500;color:var(--color-text-primary);">アクション {ia+1}</span></div>',
                                    unsafe_allow_html=True,
                                )
                            with col_adel:
                                if st.button("✕", key=f"del_action_{ii}_{ia}", help="削除"):
                                    actions_to_delete.append(ia)

                            actions[ia]["text"] = st.text_area(
                                f"act_{ii}_{ia}",
                                value=action.get("text",""),
                                height=65,
                                placeholder="例）6/15までにインタビューを10件実施し、結果をSlackで共有する",
                                label_visibility="collapsed",
                                key=f"act_txt_{ii}_{ia}",
                            )
                            c1, c2 = st.columns(2)
                            with c1:
                                try:
                                    sv = datetime.date.fromisoformat(action.get("start", datetime.date.today().isoformat()))
                                except ValueError:
                                    sv = datetime.date.today()
                                actions[ia]["start"] = st.date_input("開始日", value=sv, key=f"start_{ii}_{ia}").isoformat()
                            with c2:
                                try:
                                    ev = datetime.date.fromisoformat(action.get("end", (datetime.date.today()+datetime.timedelta(days=30)).isoformat()))
                                except ValueError:
                                    ev = datetime.date.today()+datetime.timedelta(days=30)
                                actions[ia]["end"] = st.date_input("終了日", value=ev, key=f"end_{ii}_{ia}").isoformat()
                            if actions[ia]["start"] > actions[ia]["end"]:
                                st.error("⚠️ 終了日が開始日より前です。")

                    for ia in sorted(actions_to_delete, reverse=True):
                        if len(actions) > 1:
                            actions.pop(ia)
                        else:
                            st.warning("アクションは最低1件必要です。")
                    if actions_to_delete:
                        st.rerun()

                    if len(actions) < MAX_ACTIONS:
                        if st.button("＋ アクションを追加", key=f"add_action_{ii}", use_container_width=True):
                            actions.append(blank_action()); st.rerun()
                    else:
                        st.caption(f"アクションは最大 {MAX_ACTIONS} 件まで")

            # 壁削除処理
            for ii in sorted(issues_to_delete, reverse=True):
                if len(issues) > 1:
                    issues.pop(ii)
                else:
                    st.warning("壁は最低1件必要です。")
            if issues_to_delete:
                st.rerun()

            # 壁追加ボタン
            if len(issues) < 3:
                if st.button("＋ 壁（課題）を追加する", use_container_width=True):
                    issues.append(blank_issue()); st.rerun()
            else:
                st.caption("壁は最大3件まで追加できます。")

            st.markdown("---")
            c_back, c_other, c_save = st.columns([1, 1.5, 2])
            with c_back:
                if st.button("← KR選択に戻る", use_container_width=True):
                    st.session_state.plan_step = 0; st.rerun()
            with c_other:
                if st.button("別のKRも入力する", use_container_width=True):
                    draft[krs[kr_idx]["id"]] = {"issues": issues}
                    st.session_state.plan_step = 0; st.rerun()
            with c_save:
                has_valid = any(
                    iss.get("text","").strip() and
                    any(a.get("text","").strip() for a in iss.get("actions",[]))
                    for iss in issues
                )
                if st.button("💾 保存する", type="primary",
                             use_container_width=True, disabled=not has_valid):
                    draft[krs[kr_idx]["id"]] = {"issues": issues}
                    payload = dict(
                        member   = member,
                        month    = month_str,
                        saved_at = datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
                        items    = [
                            dict(
                                kr_id    = kr_["id"],
                                kr_label = kr_["label"],
                                kr_text  = kr_["text"],
                                issues   = [
                                    {
                                        "text":    iss.get("text",""),
                                        "actions": [
                                            {"text": a.get("text",""), "start": a.get("start",""), "end": a.get("end","")}
                                            for a in iss.get("actions",[]) if a.get("text","").strip()
                                        ],
                                    }
                                    for iss in draft.get(kr_["id"],{}).get("issues",[])
                                    if iss.get("text","").strip()
                                ],
                            )
                            for kr_ in krs
                            if draft.get(kr_["id"],{}).get("issues")
                        ],
                    )
                    if io_save_plan(month_str, member, payload):
                        st.toast(f"✅ {member} さんのプランを保存しました！", icon="🎉")
                        st.success("保存完了！「DASHBOARD」タブで全員の計画を確認できます。")
                        del st.session_state[draft_key]
                        st.session_state.plan_step = 0; st.rerun()

        with col_tree:
            render_logic_tree(master, kr_idx, issues, pal)


# ══════════════════════════════════════════════════════════════════════════════
# ガントチャート
# ══════════════════════════════════════════════════════════════════════════════

def build_gantt(all_plans: list[dict]) -> Any | None:
    rows = []
    for plan in all_plans:
        m = plan.get("member","不明")
        for item in plan.get("items",[]):
            for iss in item.get("issues",[]):
                for action in iss.get("actions",[]):
                    if not action.get("text") or not action.get("start") or not action.get("end"):
                        continue
                    rows.append(dict(
                        member = m,
                        label  = f'{m}｜{item["kr_label"]}',
                        action = action["text"],
                        start  = pd.Timestamp(action["start"]),
                        end    = pd.Timestamp(action["end"]) + pd.Timedelta(days=1),
                    ))
    if not rows:
        return None
    df = pd.DataFrame(rows)
    fig = px.timeline(df, x_start="start", x_end="end", y="label", color="member",
        color_discrete_map={m: mpal(m)["main"] for m in MEMBERS},
        hover_data={"action": True, "start": False, "end": False},
        labels={"label": ""})
    fig.update_yaxes(autorange="reversed", tickfont=dict(size=11))
    fig.update_xaxes(title="", tickfont=dict(size=11))
    fig.update_layout(
        plot_bgcolor="white", paper_bgcolor="white",
        font=dict(family="Calibri, Arial", size=12, color="#1B4F72"),
        legend=dict(title="担当者", orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        margin=dict(l=10, r=10, t=44, b=10),
        height=max(300, len(rows)*44+80))
    fig.update_traces(marker_line_width=0, opacity=0.88)
    return fig


# ══════════════════════════════════════════════════════════════════════════════
# DASHBOARD タブ
# ══════════════════════════════════════════════════════════════════════════════

def render_dashboard(master: dict):
    month_str   = st.session_state.month_str
    month_label = datetime.date.fromisoformat(month_str+"-01").strftime("%Y年%m月")

    fc, _ = st.columns([2, 6])
    with fc:
        if st.button("🔄 データを読み込む", use_container_width=True, type="primary"):
            st.session_state.team_data = None

    if st.session_state.team_data is None:
        st.session_state.team_data = io_list_plans(month_str)

    all_plans: list[dict] = st.session_state.team_data or []
    submitted = {p["member"] for p in all_plans if p.get("member")}

    # 提出状況
    st.markdown('<div class="member-grid">', unsafe_allow_html=True)
    for m in MEMBERS:
        pal  = mpal(m)
        done = m in submitted
        bg   = pal["light"] if done else "var(--color-background-secondary)"
        bdr  = pal["main"]  if done else "var(--color-border-tertiary)"
        tc   = pal["text"]  if done else "var(--color-text-secondary)"
        stBg = pal["main"]  if done else "var(--color-background-secondary)"
        stTc = "#fff"        if done else "var(--color-text-secondary)"
        st.markdown(
            f'<div class="member-card" style="background:{bg};border-color:{bdr};">'
            f'<div class="mc-av" style="background:{stBg};color:{stTc};">{initials(m)}</div>'
            f'<div class="mc-name" style="color:{tc};">{m}</div>'
            f'<div class="mc-badge" style="background:{stBg};color:{stTc};">{"✅ 提出済" if done else "⏳ 未提出"}</div>'
            f'</div>', unsafe_allow_html=True,
        )
    st.markdown("</div>", unsafe_allow_html=True)

    not_sub = [m for m in MEMBERS if m not in submitted]
    if not_sub:
        st.warning(f"⏳ 未提出：{' / '.join(not_sub)}")

    if not all_plans:
        st.info("まだ誰も計画を提出していません。")
        return

    # 統計
    total_actions = sum(
        len(iss.get("actions",[]))
        for p in all_plans for item in p.get("items",[])
        for iss in item.get("issues",[])
    )
    all_starts = [
        a["start"] for p in all_plans for item in p.get("items",[])
        for iss in item.get("issues",[]) for a in iss.get("actions",[]) if a.get("start")
    ]
    all_ends = [
        a["end"] for p in all_plans for item in p.get("items",[])
        for iss in item.get("issues",[]) for a in iss.get("actions",[]) if a.get("end")
    ]

    st.markdown('<div class="stat-row">', unsafe_allow_html=True)
    for n, lbl in [
        (len(submitted), "提出メンバー数"), (total_actions, "合計アクション数"),
        (min(all_starts) if all_starts else "-", "最早 開始日"),
        (max(all_ends)   if all_ends   else "-", "最遅 終了日"),
    ]:
        st.markdown(f'<div class="stat"><div class="stat-n">{n}</div><div class="stat-l">{lbl}</div></div>', unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    # フィードバックパネル
    with st.expander("🔍 マネジャー向け レビューの着眼点", expanded=False):
        st.markdown("""
<div class="fb-panel">
<h4>📌 チェックリスト</h4>
<div class="fb-item"><span class="fb-ico">🔥</span><div><b>野心度</b>　70〜80%達成が理想。簡単すぎると成長が止まる。</div></div>
<div class="fb-item"><span class="fb-ico">🔗</span><div><b>ロジック</b>　「課題→アクション」の因果関係は通っているか？</div></div>
<div class="fb-item"><span class="fb-ico">⚡</span><div><b>リソース</b>　複数アクション合計で1ヶ月以内に現実的にやり切れるか？</div></div>
<div class="fb-item"><span class="fb-ico">🎯</span><div><b>整合性</b>　全員のアクションが同じObjectiveを向いているか？</div></div>
</div>
""", unsafe_allow_html=True)

    # メンバー別カード
    st.markdown("### 📋 メンバー別 計画カード")
    for plan in all_plans:
        m   = plan.get("member","不明")
        pal = mpal(m)
        st.markdown(
            f'<div class="dash-card" style="border-left-color:{pal["main"]};">'
            f'<div class="dash-hdr">'
            f'<div class="dash-av" style="background:{pal["main"]};color:#fff;">{initials(m)}</div>'
            f'<div><div style="font-size:.95rem;font-weight:600;color:var(--color-text-primary);">{m}</div>'
            f'<div style="font-size:.72rem;color:var(--color-text-secondary);">保存日時：{plan.get("saved_at","")}</div></div>'
            f'</div>', unsafe_allow_html=True,
        )
        for item in plan.get("items",[]):
            kr_idx_ = next((i for i,kr in enumerate(master.get("key_results",[])) if kr.get("id")==item.get("kr_id")), 0)
            kr_col  = KR_COLORS[kr_idx_ % len(KR_COLORS)]
            issues_html = ""
            for ii, iss in enumerate(item.get("issues",[])):
                if not iss.get("text","").strip():
                    continue
                a_html = "".join(
                    f'<div class="action-list-item">'
                    f'<div class="a-num" style="background:{pal["main"]};">A{ai+1}</div>'
                    f'<div><span style="font-size:.8rem;color:var(--color-text-primary);">{a.get("text","")}</span>'
                    f'<div style="font-size:.7rem;color:var(--color-text-secondary);">{a.get("start","")} → {a.get("end","")}</div></div>'
                    f'</div>'
                    for ai, a in enumerate(iss.get("actions",[])) if a.get("text","").strip()
                )
                issues_html += (
                    f'<div style="background:var(--color-background-primary);border-radius:6px;'
                    f'padding:.35rem .6rem;margin-bottom:.35rem;">'
                    f'<div style="display:flex;align-items:center;gap:5px;margin-bottom:.25rem;">'
                    f'<div style="width:18px;height:18px;border-radius:50%;background:#F39C12;'
                    f'display:flex;align-items:center;justify-content:center;font-size:.62rem;font-weight:600;color:#fff;">壁{ii+1}</div>'
                    f'<span style="font-size:.78rem;font-weight:500;color:var(--color-text-primary);">{iss["text"][:45]}</span>'
                    f'</div>{a_html}</div>'
                )
            st.markdown(
                f'<div class="kr-block">'
                f'<div style="display:flex;align-items:center;gap:6px;margin-bottom:.35rem;">'
                f'<span style="background:{kr_col};color:#fff;font-size:.68rem;font-weight:700;padding:1px 8px;border-radius:20px;">{item["kr_label"]}</span>'
                f'<span style="font-size:.78rem;font-weight:500;color:var(--color-text-primary);">{item.get("kr_text","")[:45]}</span>'
                f'</div>'
                f'{issues_html}</div>', unsafe_allow_html=True,
            )
        st.markdown("</div>", unsafe_allow_html=True)

    # ガントチャート
    st.markdown("### 📊 統合ガントチャート")
    fig = build_gantt(all_plans)
    if fig:
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("アクションと日付が入力されたデータがあるとガントチャートが表示されます。")

    # PPTX
    st.markdown("---")
    st.markdown("### 📥 PPTX ダウンロード")
    gen_c, hint_c = st.columns([2, 5])
    with gen_c:
        gen_btn = st.button("🚀 PPTXを生成する", type="primary", use_container_width=True)
    with hint_c:
        st.caption(f"構成：表紙 + チームOKR + メンバー別詳細({len(all_plans)}枚) + ガントチャート ＝ 計{2+len(all_plans)+1}枚")

    if gen_btn:
        with st.spinner("PPTX生成中…"):
            try:
                gantt_png  = fig.to_image(format="png", width=1400, scale=2) if fig else None
                pptx_bytes = build_pptx(CFG["team_name"], month_label, master, all_plans, gantt_png)
                fname = f"OKR_{CFG['team_name']}_{month_str}.pptx"
                st.toast("🎉 PPTXの生成が完了しました！", icon="📊")
                st.download_button(
                    label="⬇️ PPTXをダウンロード", data=pptx_bytes, file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True,
                )
            except Exception:
                st.error("PPTX生成中にエラーが発生しました。")
                st.code(traceback.format_exc())


# ══════════════════════════════════════════════════════════════════════════════
# PPTX 生成
# ══════════════════════════════════════════════════════════════════════════════

def _rgb(h):
    h=h.lstrip("#"); return RGBColor(int(h[0:2],16),int(h[2:4],16),int(h[4:6],16))

def _rect(sl,x,y,w,h,fill,border=None,bw=0.5):
    s=sl.shapes.add_shape(1,x,y,w,h); s.fill.solid(); s.fill.fore_color.rgb=fill
    if border: s.line.color.rgb=border; s.line.width=Pt(bw)
    else: s.line.fill.background()
    return s

def _txt(sl,text,x,y,w,h,sz,bold=False,color=None,align=PP_ALIGN.LEFT,italic=False):
    color=color or RGBColor(0xFF,0xFF,0xFF)
    tb=sl.shapes.add_textbox(x,y,w,h); tf=tb.text_frame; tf.word_wrap=True
    p=tf.paragraphs[0]; p.alignment=align; r=p.add_run()
    r.text=text; r.font.size=Pt(sz); r.font.bold=bold
    r.font.italic=italic; r.font.color.rgb=color; r.font.name="Calibri"
    return tb

def build_pptx(team_name,month_label,master,all_plans,gantt_png):
    C_NAVY=_rgb("#0D1B2A"); C_DEEP=_rgb("#1B4F72"); C_GRN=_rgb("#1D6A45")
    C_GOLD=_rgb("#F39C12"); C_SNOW=_rgb("#F8FAFC"); C_SLAT=_rgb("#334155")
    C_GRAY=_rgb("#95A5A6"); C_WHT=RGBColor(0xFF,0xFF,0xFF)
    W=Inches(13.33); H=Inches(7.5)
    prs=Presentation(); prs.slide_width=W; prs.slide_height=H
    blank=prs.slide_layouts[6]
    obj=master.get("objective","")
    krs=master.get("key_results",[])

    # 表紙
    s1=prs.slides.add_slide(blank)
    _rect(s1,0,0,W,H*.62,C_NAVY); _rect(s1,0,H*.62,W,H*.38,C_SNOW)
    _rect(s1,Inches(.9),Inches(1.0),Inches(.07),Inches(3.2),C_GOLD)
    _txt(s1,"OKR × 業務計画",Inches(1.1),Inches(.9),Inches(10),Inches(1.1),sz=44,bold=True,color=C_WHT)
    _txt(s1,team_name,Inches(1.1),Inches(2.05),Inches(8),Inches(.7),sz=24,italic=True,color=_rgb("#A9CCE3"))
    _txt(s1,f"{master.get('quarter','')}　{month_label}",Inches(1.1),Inches(2.78),Inches(7),Inches(.55),sz=18,color=_rgb("#A9CCE3"))
    _txt(s1,f"作成日：{datetime.date.today().strftime('%Y年%m月%d日')}",Inches(1.1),H*.65,Inches(7),Inches(.5),sz=14,color=C_SLAT)
    bx=Inches(1.1)
    for m in [p["member"] for p in all_plans if p.get("member")][:4]:
        pal=mpal(m); _rect(s1,bx,H*.77,Inches(2.55),Inches(.44),_rgb(pal["main"]))
        _txt(s1,m,bx+Cm(.1),H*.77+Cm(.05),Inches(2.35),Inches(.38),sz=13,bold=True,color=C_WHT,align=PP_ALIGN.CENTER)
        bx+=Inches(2.7)
    for rx,ry,rw in [(W-Inches(3.0),Inches(.1),Inches(2.8)),(W-Inches(5.3),H*.28,Inches(1.2))]:
        c=s1.shapes.add_shape(9,rx,ry,rw,rw); c.fill.solid(); c.fill.fore_color.rgb=C_GRN; c.line.fill.background()

    # OKRサマリー
    s2=prs.slides.add_slide(blank); _rect(s2,0,0,W,Inches(.95),C_DEEP)
    _txt(s2,"チーム OKR サマリー",Inches(.5),Inches(.1),Inches(11),Inches(.72),sz=26,bold=True,color=C_WHT)
    _txt(s2,f"🌟 {obj}",Inches(.5),Inches(1.1),W-Inches(1.0),Inches(.75),sz=15,bold=True,color=C_DEEP)
    ky=Inches(1.95); kh=Inches(.58)
    for i,kr in enumerate(krs[:3]):
        col=_rgb(KR_COLORS[i%len(KR_COLORS)]); _rect(s2,Inches(.5),ky,Inches(.06),kh,col)
        _rect(s2,Inches(.62),ky,W-Inches(1.12),kh,_rgb("#EBF5FB"))
        _txt(s2,f"{kr['label']}　{kr.get('text','')}",Inches(.75),ky+Cm(.12),W-Inches(1.4),kh-Cm(.2),sz=13,color=C_SLAT)
        ky+=kh+Inches(.12)
    card_w=(W-Inches(.8)-Inches(.25)*3)/4; cx=Inches(.4); cy=Inches(3.55); ch=H-cy-Inches(.2)
    for plan in all_plans[:4]:
        m=plan.get("member",""); pal=mpal(m)
        _rect(s2,cx,cy,card_w,ch,C_SNOW,border=_rgb("#D6DBDF")); _rect(s2,cx,cy,card_w,Inches(.42),_rgb(pal["main"]))
        _txt(s2,m,cx+Cm(.15),cy+Cm(.05),card_w-Cm(.3),Inches(.38),sz=11,bold=True,color=C_WHT,align=PP_ALIGN.CENTER)
        iy=cy+Inches(.5)
        for item in plan.get("items",[]):
            for iss in item.get("issues",[]):
                for a in iss.get("actions",[])[:1]:
                    if a.get("text","").strip() and iy<cy+ch-Inches(.1):
                        _txt(s2,f"• {a['text'][:28]}",cx+Cm(.2),iy,card_w-Cm(.4),Inches(.45),sz=8,color=C_SLAT)
                        iy+=Inches(.38)
        cx+=card_w+Inches(.25)

    # メンバー別詳細
    for plan in all_plans:
        m=plan.get("member","不明"); pal=mpal(m)
        sm=prs.slides.add_slide(blank)
        _rect(sm,0,0,W,Inches(.95),_rgb(pal["main"])); _rect(sm,0,0,Inches(.05),H,_rgb(pal["main"]))
        _txt(sm,f"{m}　詳細アクション計画",Inches(.5),Inches(.1),Inches(10),Inches(.72),sz=24,bold=True,color=C_WHT)
        _txt(sm,f"🌟 {obj[:65]}",Inches(.5),Inches(1.05),W-Inches(1.0),Inches(.45),sz=11,italic=True,color=C_SLAT)
        hdrs=["KR","指標（Key Result）","課題","アクション","開始日","終了日"]
        col_ws=[Inches(.55),Inches(2.0),Inches(2.5),Inches(3.5),Inches(1.1),Inches(1.1)]
        row_h=Inches(.50); tbl_x=Inches(.4); tbl_y=Inches(1.6)
        cx=tbl_x
        for hdr,cw in zip(hdrs,col_ws):
            _rect(sm,cx,tbl_y,cw,row_h,_rgb(pal["main"]))
            _txt(sm,hdr,cx+Cm(.12),tbl_y+Cm(.08),cw-Cm(.24),row_h,sz=9,bold=True,color=C_WHT,align=PP_ALIGN.CENTER)
            cx+=cw
        row_bgs=[C_SNOW,RGBColor(0xFF,0xFF,0xFF)]; ri=0
        for item in plan.get("items",[]):
            for iss in item.get("issues",[]):
                for a in [x for x in iss.get("actions",[]) if x.get("text","").strip()]:
                    cy_=tbl_y+row_h*(ri+1)
                    if cy_+row_h>H-Inches(.15): break
                    vals=[item.get("kr_label",""),item.get("kr_text","")[:35],iss.get("text","")[:38],a.get("text","")[:50],a.get("start",""),a.get("end","")]
                    cx=tbl_x; bg=row_bgs[ri%2]
                    for ci,(val,cw) in enumerate(zip(vals,col_ws)):
                        _rect(sm,cx,cy_,cw,row_h,bg,border=_rgb("#D5D8DC"))
                        _txt(sm,val,cx+Cm(.15),cy_+Cm(.07),cw-Cm(.3),row_h,sz=9,
                             color=_rgb(pal["main"]) if ci==0 else C_SLAT,
                             align=PP_ALIGN.CENTER if ci in(0,4,5) else PP_ALIGN.LEFT)
                        cx+=cw
                    ri+=1

    # ガントチャート
    sg=prs.slides.add_slide(blank); _rect(sg,0,0,W,Inches(.95),C_NAVY)
    _txt(sg,"統合ガントチャート",Inches(.5),Inches(.1),Inches(11),Inches(.72),sz=26,bold=True,color=C_WHT)
    if gantt_png:
        sg.shapes.add_picture(io.BytesIO(gantt_png),Inches(.3),Inches(1.1),width=W-Inches(.6),height=H-Inches(1.3))
    else:
        _txt(sg,"アクションデータがありません",Inches(1.0),Inches(3.0),W-Inches(2.0),Inches(.8),sz=18,color=C_GRAY,align=PP_ALIGN.CENTER)

    for sl in prs.slides:
        _txt(sl,f"{team_name}　{month_label}",Inches(.4),H-Inches(.32),W-Inches(.8),Inches(.28),sz=8,italic=True,color=C_GRAY,align=PP_ALIGN.RIGHT)

    buf=io.BytesIO(); prs.save(buf); return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
# メイン
# ══════════════════════════════════════════════════════════════════════════════

def main():
    _init_session()
    team_name  = CFG["team_name"]
    month_str  = st.session_state.month_str
    month_disp = datetime.date.fromisoformat(month_str+"-01").strftime("%Y年%m月")
    master     = io_get_master()

    with st.sidebar:
        st.markdown(f"## 🌟 OKR管理\n**{team_name}**")
        st.markdown("---")

        st.markdown("### 📌 あなたの名前")
        selected = st.selectbox("名前", MEMBERS,
            index=MEMBERS.index(st.session_state.cur_member) if st.session_state.cur_member in MEMBERS else 0,
            label_visibility="collapsed")
        if selected != st.session_state.cur_member:
            st.session_state.cur_member = selected; st.session_state.plan_step = 0; st.rerun()

        st.markdown(f"### 📅 対象月\n**{month_disp}**")
        st.caption(f"四半期：{master.get('quarter','未設定')}")
        st.markdown("---")
        st.markdown("### メンバーカラー")
        for i, m in enumerate(MEMBERS):
            c = PALETTE[i % len(PALETTE)]["main"]
            st.markdown(
                f'<div style="display:flex;align-items:center;gap:7px;margin:3px 0;">'
                f'<div style="width:10px;height:10px;border-radius:50%;background:{c};flex-shrink:0;"></div>'
                f'<span style="font-size:.82rem;color:#C8E0FF;">{m}</span></div>',
                unsafe_allow_html=True,
            )
        if st.session_state.admin_auth:
            st.markdown("---")
            st.markdown('<span style="color:#F9E08B;font-size:.78rem;">🔐 管理者モード ON</span>', unsafe_allow_html=True)
            if st.button("ログアウト", use_container_width=True):
                st.session_state.admin_auth = False; st.rerun()

    render_north_star(master)

    tab_home, tab_strategy, tab_plan, tab_dash = st.tabs([
        "🏠 HOME", "🏛️ STRATEGY",
        f"📝 PLAN（{st.session_state.cur_member}）", "📊 DASHBOARD",
    ])

    with tab_home:
        render_home()
    with tab_strategy:
        st.subheader("🏛️ 四半期 OKR 作成ワークショップ")
        render_strategy(master)
    with tab_plan:
        st.subheader(f"📝 {st.session_state.cur_member}さんの {month_disp} プラン")
        render_plan(master)
    with tab_dash:
        st.subheader(f"📊 {month_disp}　チーム統合ダッシュボード")
        render_dashboard(master)


if __name__ == "__main__":
    main()
