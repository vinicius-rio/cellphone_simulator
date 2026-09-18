from __future__ import annotations

import argparse
import html
import json
import re
from pathlib import Path

import pandas as pd

REQUIRED_COLUMNS = [
    "id_sessao",
    "seq_sessao",
    "mensagem_cliente",
    "resposta",
    "datahora",
]

OPTIONAL_COLUMNS = [
    "demanda",
    "termino_conversa",
    "nota_csat",
    "comentario_csat",
]

BUTTON_RE = re.compile(r"^Bot[aã]o\s*\d+\.\s*(.+)$", re.IGNORECASE)
LINK_RE = re.compile(r"^Link\s*\d+\.\s*(.+)$", re.IGNORECASE)


def clean_value(value):
    if pd.isna(value):
        return None
    if hasattr(value, "item"):
        try:
            return value.item()
        except Exception:
            pass
    return value


def parse_response(value) -> dict:
    """Split chatbot text into body, buttons and links."""
    if pd.isna(value):
        return {"texto": "", "botoes": [], "links": []}

    body, buttons, links = [], [], []
    for raw_line in str(value).replace("\r\n", "\n").split("\n"):
        line = raw_line.strip()
        if not line:
            continue

        button_match = BUTTON_RE.match(line)
        if button_match:
            buttons.append(button_match.group(1).strip())
            continue

        link_match = LINK_RE.match(line)
        if link_match:
            links.append(link_match.group(1).strip())
            continue

        body.append(line)

    return {
        "texto": "\n".join(body),
        "botoes": buttons,
        "links": links,
    }


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [
        str(c).strip().lower().replace(" ", "_")
        for c in df.columns
    ]
    return df


def load_conversations(xlsx_path: Path, sheet_name=0) -> list[dict]:
    df = pd.read_excel(xlsx_path, sheet_name=sheet_name, engine="openpyxl")
    df = normalize_columns(df)

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"Colunas obrigatórias ausentes: {', '.join(missing)}")

    for col in OPTIONAL_COLUMNS:
        if col not in df.columns:
            df[col] = None

    df["datahora"] = pd.to_datetime(df["datahora"], errors="coerce")
    df["seq_sessao"] = pd.to_numeric(df["seq_sessao"], errors="coerce")
    df = df.sort_values(["id_sessao", "seq_sessao", "datahora"], kind="stable")

    conversations = []
    for session_id, group in df.groupby("id_sessao", sort=False, dropna=False):
        group = group.sort_values(["seq_sessao", "datahora"], kind="stable")
        first = group.iloc[0]
        last = group.iloc[-1]

        interactions = []
        for _, row in group.iterrows():
            dt = row["datahora"]
            dt_iso = None if pd.isna(dt) else dt.isoformat()
            hour = "" if pd.isna(dt) else dt.strftime("%H:%M")

            interactions.append({
                "seq_sessao": None if pd.isna(row["seq_sessao"]) else int(row["seq_sessao"]),
                "datahora": dt_iso,
                "hora": hour,
                "cliente": "" if pd.isna(row["mensagem_cliente"]) else str(row["mensagem_cliente"]),
                "chatbot": parse_response(row["resposta"]),
            })

        csat = clean_value(last.get("nota_csat"))
        try:
            if csat is not None and float(csat).is_integer():
                csat = int(float(csat))
        except Exception:
            pass

        conversations.append({
            "id_sessao": "" if pd.isna(session_id) else str(session_id),
            "demanda": "" if pd.isna(first.get("demanda")) else str(first.get("demanda")),
            "termino_conversa": "" if pd.isna(last.get("termino_conversa")) else str(last.get("termino_conversa")),
            "nota_csat": csat,
            "comentario_csat": "" if pd.isna(last.get("comentario_csat")) else str(last.get("comentario_csat")),
            "qtd_interacoes": len(interactions),
            "interacoes": interactions,
        })

    return conversations


HTML_TEMPLATE = r'''<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Chatbot Conversation Viewer</title>
<style>
:root{
  --orange:#ff6200;
  --orange-soft:#fff0e6;
  --ink:#242833;
  --muted:#6f7580;
  --line:#e4e6ea;
  --surface:#ffffff;
  --page:#f5f6f8;
  --bubble:#f1f2f4;
  --success:#15803d;
  --warning:#b45309;
  --danger:#b91c1c;
  --shadow:0 24px 70px rgba(25, 28, 38, .14);
}
*{box-sizing:border-box}
body{margin:0;background:linear-gradient(140deg,#fafafa 0%,#f1f3f6 100%);color:var(--ink);font-family:Inter, ui-sans-serif, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;min-height:100vh}
button,input,select{font:inherit}
.app{display:grid;grid-template-columns:320px minmax(390px,1fr) 320px;gap:24px;max-width:1520px;margin:0 auto;padding:28px;height:100vh;overflow:hidden}
.panel{background:rgba(255,255,255,.9);border:1px solid rgba(220,223,229,.85);border-radius:24px;box-shadow:0 10px 35px rgba(35,38,48,.06);backdrop-filter:blur(8px)}
.sidebar{padding:20px;display:flex;flex-direction:column;min-height:0;height:calc(100vh - 56px);overflow:hidden}
.brand{display:flex;align-items:center;gap:12px;margin-bottom:20px}
.brand-mark{width:40px;height:40px;border-radius:14px;background:var(--orange);display:grid;place-items:center;color:white;font-weight:800;font-size:20px;box-shadow:0 8px 18px rgba(255,98,0,.25)}
.brand h1{font-size:16px;margin:0}.brand p{font-size:12px;color:var(--muted);margin:3px 0 0}
.search{position:relative;margin-bottom:10px}.search input{width:100%;border:1px solid var(--line);background:#fafbfc;padding:12px 14px;border-radius:14px;outline:none}.search input:focus{border-color:#ffb98c;box-shadow:0 0 0 4px #fff1e8}.filters{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:10px}.filter{width:100%;border:1px solid var(--line);background:#fafbfc;color:var(--ink);padding:10px 11px;border-radius:12px;outline:none;font-size:12px}.filter:focus{border-color:#ffb98c;box-shadow:0 0 0 3px #fff1e8}.filter-wide{grid-column:1/-1}.session-toolbar{display:flex;align-items:center;justify-content:space-between;gap:10px;margin:2px 2px 10px}.results-count{font-size:11px;color:var(--muted);font-weight:700}.clear-filters{border:0;background:transparent;color:#9a4300;font-size:11px;font-weight:800;cursor:pointer;padding:4px 0}.clear-filters:hover{text-decoration:underline}
.session-list{display:flex;flex-direction:column;gap:8px;overflow-y:auto;overflow-x:hidden;padding-right:6px;min-height:0;flex:1;overscroll-behavior:contain;scrollbar-gutter:stable}.session-list::-webkit-scrollbar{width:6px}.session-list::-webkit-scrollbar-thumb{background:#d9dde3;border-radius:10px}.session-list::-webkit-scrollbar-track{background:transparent}
.session{border:1px solid transparent;background:transparent;border-radius:16px;padding:12px;text-align:left;cursor:pointer;transition:.18s ease;color:var(--ink)}
.session:hover{background:#f7f8fa}.session.active{background:#fff4ed;border-color:#ffd4b8}.session-top{display:flex;align-items:center;justify-content:space-between;gap:8px}.session-id{font-weight:750;font-size:13px}.session-count{font-size:11px;color:var(--muted);background:#eef0f3;padding:3px 7px;border-radius:99px}.session-demand{font-size:12px;color:var(--muted);margin-top:4px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.session-meta{display:flex;align-items:center;justify-content:space-between;gap:8px;margin-top:8px}.session-date{font-size:10px;color:#8a9099}.mini-status{font-size:9px;font-weight:800;text-transform:capitalize;padding:3px 6px;border-radius:999px}.mini-status.resolutivo{background:#eaf7ee;color:var(--success)}.mini-status.transbordo{background:#fff4e5;color:var(--warning)}.mini-status.abandono{background:#fdecec;color:var(--danger)}
.stage{display:flex;align-items:center;justify-content:center;min-height:0}
.phone{width:430px;height:min(880px,calc(100vh - 56px));min-height:640px;background:white;border:9px solid #1f2228;border-radius:50px;box-shadow:var(--shadow);overflow:hidden;position:relative;display:flex;flex-direction:column}
.phone::after{content:"";position:absolute;inset:5px;border:1px solid rgba(255,255,255,.18);border-radius:38px;pointer-events:none}
.statusbar{height:46px;display:flex;align-items:center;justify-content:space-between;padding:0 24px;font-weight:750;font-size:14px;flex:0 0 auto}.status-icons{font-size:12px;letter-spacing:2px}
.chat-head{height:54px;display:flex;align-items:center;justify-content:space-between;padding:0 20px;border-bottom:1px solid #f0f1f3;flex:0 0 auto}.chat-title{display:flex;align-items:center;gap:10px}.spark{color:var(--orange);font-size:24px;line-height:1}.chat-name{font-weight:760}.chat-sub{font-size:11px;color:var(--muted);margin-top:1px}.close{font-size:28px;color:#363942;font-weight:300}
.chat{flex:1;overflow:auto;padding:18px 20px 28px;background:linear-gradient(#fff 0%,#fff 58%,#fff7f1 100%);scroll-behavior:smooth}.chat::-webkit-scrollbar{width:5px}.chat::-webkit-scrollbar-thumb{background:#d9dde3;border-radius:10px}
.day{text-align:center;color:var(--muted);font-size:12px;margin:8px 0 22px}
.turn{position:relative;margin-bottom:22px}.client-row{display:flex;justify-content:flex-end;gap:9px;align-items:flex-end}.avatar{width:38px;height:38px;border-radius:50%;background:#333840;color:white;display:grid;place-items:center;font-size:13px;flex:0 0 auto}.client-wrap{max-width:78%;display:flex;flex-direction:column;align-items:flex-end}.bubble-client{background:var(--bubble);padding:12px 15px;border-radius:18px 18px 5px 18px;line-height:1.42;font-size:14px;white-space:pre-wrap}.time{font-size:11px;color:var(--muted);margin-top:5px;display:flex;gap:5px;align-items:center}.checks{letter-spacing:-2px;color:#68707c;font-weight:700}
.bot-row{display:grid;grid-template-columns:26px 1fr;gap:9px;margin-top:20px}.bot-icon{color:var(--orange);font-size:22px;margin-top:1px}.bot-content{min-width:0}.bot-text{font-size:14px;line-height:1.55;white-space:pre-wrap}.choices{border:1px solid #dfe2e7;border-radius:16px;overflow:hidden;margin-top:13px;background:rgba(255,255,255,.92)}.choice{display:flex;align-items:center;justify-content:space-between;gap:10px;padding:13px 14px;border-bottom:1px solid #eceef1;font-weight:700;font-size:13px}.choice:last-child{border-bottom:0}.chev{font-size:20px;color:#717782;font-weight:400}.links{display:flex;flex-direction:column;gap:7px;margin-top:11px}.link-item{font-size:13px;color:#9a4300;font-weight:700;text-decoration:none;display:inline-flex;gap:6px;align-items:center}.link-item::after{content:"↗";font-size:12px}
.seq{position:absolute;left:-8px;top:-4px;background:#262a33;color:white;border-radius:8px;padding:2px 6px;font-size:9px;opacity:0;transform:translateX(-4px);transition:.15s}.turn:hover .seq{opacity:.9;transform:none}
.composer{height:86px;padding:12px 18px 18px;background:#fff3eb;flex:0 0 auto}.composer-inner{height:56px;background:white;border-radius:22px;border:1px solid #f0e5de;display:flex;align-items:center;justify-content:space-between;padding:0 16px;color:#626975}.mic{font-size:21px}.meta{padding:20px;min-height:0;overflow:auto}.meta h2{font-size:15px;margin:0 0 16px}.meta-card{border:1px solid var(--line);border-radius:17px;padding:14px;margin-bottom:12px;background:#fff}.label{text-transform:uppercase;letter-spacing:.08em;font-size:9px;color:var(--muted);font-weight:800;margin-bottom:6px}.value{font-size:14px;font-weight:750;overflow-wrap:anywhere}.badge{display:inline-flex;align-items:center;padding:5px 9px;border-radius:999px;font-size:11px;font-weight:800;text-transform:capitalize}.badge.resolutivo{background:#eaf7ee;color:var(--success)}.badge.transbordo{background:#fff4e5;color:var(--warning)}.badge.abandono{background:#fdecec;color:var(--danger)}.stars{letter-spacing:2px;color:#f59e0b}.comment{font-size:13px;line-height:1.5;color:#4f5560}.stats{display:grid;grid-template-columns:1fr 1fr;gap:10px}.stat{border:1px solid var(--line);border-radius:15px;padding:12px;background:#fafbfc}.stat strong{display:block;font-size:18px}.stat span{font-size:10px;color:var(--muted)}
.empty{color:var(--muted);font-size:13px;text-align:center;padding:30px 10px}
.footer-note{font-size:10px;color:#9aa0a8;margin-top:14px;line-height:1.45}
@media(max-width:1050px){.app{grid-template-columns:280px 1fr}.meta{display:none}}
@media(max-width:760px){body{background:white}.app{display:block;padding:0;height:auto;overflow:visible}.sidebar{display:none}.stage{display:block}.phone{width:100%;height:100dvh;min-height:100dvh;border:0;border-radius:0;box-shadow:none}.phone::after{display:none}}
</style>
</head>
<body>
<div class="app">
  <aside class="panel sidebar">
    <div class="brand"><div class="brand-mark">✦</div><div><h1>Conversation Viewer</h1><p>Leitura visual de sessões</p></div></div>
    <div class="search"><input id="search" placeholder="Buscar sessão ou demanda" /></div>
    <div class="filters">
      <select id="demandFilter" class="filter"><option value="">Todas as demandas</option></select>
      <select id="statusFilter" class="filter"><option value="">Todos os términos</option></select>
      <select id="csatFilter" class="filter"><option value="">Todos os CSATs</option><option value="5">CSAT 5</option><option value="4">CSAT 4</option><option value="3">CSAT 3</option><option value="2">CSAT 2</option><option value="1">CSAT 1</option><option value="none">Sem CSAT</option></select>
      <select id="sortOrder" class="filter"><option value="recent">Mais recentes</option><option value="oldest">Mais antigas</option><option value="id-asc">ID crescente</option><option value="id-desc">ID decrescente</option></select>
    </div>
    <div class="session-toolbar"><span id="resultsCount" class="results-count"></span><button id="clearFilters" class="clear-filters" type="button">Limpar filtros</button></div>
    <div id="sessionList" class="session-list"></div>
  </aside>

  <main class="stage">
    <section class="phone" aria-label="Simulador de conversa em celular">
      <div class="statusbar"><span id="phoneClock">21:30</span><span class="status-icons">▦ ))) ▰</span></div>
      <div class="chat-head">
        <div class="chat-title"><span class="spark">✦</span><div><div class="chat-name">Assistente virtual</div><div class="chat-sub" id="headSession">Sessão</div></div></div>
        <div class="close">×</div>
      </div>
      <div id="chat" class="chat"></div>
      <div class="composer"><div class="composer-inner"><span>Digite aqui</span><span class="mic">♩</span></div></div>
    </section>
  </main>

  <aside class="panel meta">
    <h2>Detalhes da sessão</h2>
    <div id="metadata"></div>
  </aside>
</div>
<script>
const conversations = __DATA__;
let selectedId = conversations[0]?.id_sessao ?? null;

const esc = (value='') => String(value ?? '').replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[ch]));
const sessionById = id => conversations.find(c => c.id_sessao === id);

function statusClass(value=''){
  const v = value.toLowerCase();
  if(v.includes('resol')) return 'resolutivo';
  if(v.includes('trans')) return 'transbordo';
  if(v.includes('aband')) return 'abandono';
  return '';
}

function sessionTimestamp(c){
  const raw = c.interacoes?.[0]?.datahora || '';
  const dt = raw ? new Date(raw) : null;
  return dt && !Number.isNaN(dt.getTime()) ? dt.getTime() : 0;
}

function sessionDateLabel(c){
  const raw = c.interacoes?.[0]?.datahora || '';
  const dt = raw ? new Date(raw) : null;
  return dt && !Number.isNaN(dt.getTime())
    ? dt.toLocaleDateString('pt-BR',{day:'2-digit',month:'2-digit',year:'2-digit'})
    : 'Sem data';
}

function currentFilters(){
  return {
    q: document.getElementById('search').value.trim().toLowerCase(),
    demand: document.getElementById('demandFilter').value,
    status: document.getElementById('statusFilter').value,
    csat: document.getElementById('csatFilter').value,
    sort: document.getElementById('sortOrder').value
  };
}

function renderSessions(){
  const {q,demand,status,csat,sort} = currentFilters();
  const list = document.getElementById('sessionList');
  let items = conversations.filter(c => {
    const matchesText = !q || c.id_sessao.toLowerCase().includes(q) || (c.demanda||'').toLowerCase().includes(q);
    const matchesDemand = !demand || (c.demanda||'') === demand;
    const matchesStatus = !status || (c.termino_conversa||'') === status;
    const matchesCsat = !csat || (csat === 'none' ? c.nota_csat == null : String(c.nota_csat) === csat);
    return matchesText && matchesDemand && matchesStatus && matchesCsat;
  });

  items = [...items].sort((a,b) => {
    if(sort === 'oldest') return sessionTimestamp(a) - sessionTimestamp(b);
    if(sort === 'id-asc') return a.id_sessao.localeCompare(b.id_sessao, 'pt-BR', {numeric:true});
    if(sort === 'id-desc') return b.id_sessao.localeCompare(a.id_sessao, 'pt-BR', {numeric:true});
    return sessionTimestamp(b) - sessionTimestamp(a);
  });

  document.getElementById('resultsCount').textContent = `${items.length} de ${conversations.length} sessões`;
  list.innerHTML = items.map(c => `
    <button class="session ${c.id_sessao===selectedId?'active':''}" data-id="${esc(c.id_sessao)}">
      <div class="session-top"><span class="session-id">${esc(c.id_sessao)}</span><span class="session-count">${c.qtd_interacoes} inter.</span></div>
      <div class="session-demand">${esc(c.demanda || 'Sem demanda')}</div>
      <div class="session-meta"><span class="session-date">${esc(sessionDateLabel(c))}</span><span class="mini-status ${statusClass(c.termino_conversa)}">${esc(c.termino_conversa || '—')}</span></div>
    </button>`).join('') || '<div class="empty">Nenhuma sessão encontrada.</div>';
  list.querySelectorAll('.session').forEach(btn => btn.addEventListener('click', () => selectSession(btn.dataset.id)));
}

function renderBot(interaction){
  const bot = interaction.chatbot || {};
  const buttons = (bot.botoes || []).map(x => `<div class="choice"><span>${esc(x)}</span><span class="chev">›</span></div>`).join('');
  const links = (bot.links || []).map(x => `<a class="link-item" href="#" onclick="return false">${esc(x)}</a>`).join('');
  return `
    <div class="bot-row">
      <div class="bot-icon">✦</div>
      <div class="bot-content">
        ${bot.texto ? `<div class="bot-text">${esc(bot.texto)}</div>` : ''}
        ${buttons ? `<div class="choices">${buttons}</div>` : ''}
        ${links ? `<div class="links">${links}</div>` : ''}
        <div class="time">${esc(interaction.hora || '')}</div>
      </div>
    </div>`;
}

function renderChat(c){
  const chat = document.getElementById('chat');
  const firstDate = c.interacoes?.[0]?.datahora ? new Date(c.interacoes[0].datahora) : null;
  const day = firstDate && !Number.isNaN(firstDate.getTime()) ? firstDate.toLocaleDateString('pt-BR',{day:'2-digit',month:'long',year:'numeric'}) : 'Conversa';
  chat.innerHTML = `<div class="day">${esc(day)}</div>` + c.interacoes.map(i => `
    <div class="turn">
      <span class="seq">#${esc(i.seq_sessao ?? '')}</span>
      ${i.cliente ? `<div class="client-row"><div class="client-wrap"><div class="bubble-client">${esc(i.cliente)}</div><div class="time">${esc(i.hora||'')} <span class="checks">✓✓</span></div></div><div class="avatar">VC</div></div>` : ''}
      ${renderBot(i)}
    </div>`).join('');
  chat.scrollTop = 0;
  const lastTime = c.interacoes?.at(-1)?.hora || c.interacoes?.[0]?.hora || '21:30';
  document.getElementById('phoneClock').textContent = lastTime;
  document.getElementById('headSession').textContent = `${c.id_sessao} · ${c.demanda || 'Sem demanda'}`;
}

function renderMetadata(c){
  const score = c.nota_csat == null ? '—' : c.nota_csat;
  const stars = c.nota_csat == null ? 'Sem nota' : '★'.repeat(Math.max(0,Math.min(5,Number(c.nota_csat)))) + '☆'.repeat(Math.max(0,5-Number(c.nota_csat)));
  const first = c.interacoes?.[0]?.datahora ? new Date(c.interacoes[0].datahora) : null;
  const last = c.interacoes?.at(-1)?.datahora ? new Date(c.interacoes.at(-1).datahora) : null;
  const durationMin = first && last ? Math.max(0, Math.round((last-first)/60000)) : 0;
  document.getElementById('metadata').innerHTML = `
    <div class="meta-card"><div class="label">ID da sessão</div><div class="value">${esc(c.id_sessao)}</div></div>
    <div class="meta-card"><div class="label">Demanda</div><div class="value">${esc(c.demanda || '—')}</div></div>
    <div class="stats"><div class="stat"><strong>${c.qtd_interacoes}</strong><span>interações</span></div><div class="stat"><strong>${durationMin}</strong><span>minutos</span></div></div>
    <div class="meta-card" style="margin-top:12px"><div class="label">Término</div><span class="badge ${statusClass(c.termino_conversa)}">${esc(c.termino_conversa || '—')}</span></div>
    <div class="meta-card"><div class="label">CSAT</div><div class="value">${esc(score)}</div><div class="stars">${esc(stars)}</div></div>
    <div class="meta-card"><div class="label">Comentário CSAT</div><div class="comment">${esc(c.comentario_csat || 'Sem comentário.')}</div></div>
    <div class="footer-note">Passe o mouse sobre uma interação para visualizar o <strong>seq_sessao</strong>. Os botões e links são reconstruídos automaticamente a partir dos marcadores “Botão N.” e “Link N.”.</div>`;
}

function selectSession(id){
  selectedId = id;
  const c = sessionById(id);
  if(!c) return;
  renderSessions();
  renderChat(c);
  renderMetadata(c);
}

function populateFilters(){
  const demands = [...new Set(conversations.map(c => c.demanda).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'pt-BR'));
  const statuses = [...new Set(conversations.map(c => c.termino_conversa).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'pt-BR'));
  const demandSelect = document.getElementById('demandFilter');
  const statusSelect = document.getElementById('statusFilter');
  demandSelect.insertAdjacentHTML('beforeend', demands.map(x => `<option value="${esc(x)}">${esc(x)}</option>`).join(''));
  statusSelect.insertAdjacentHTML('beforeend', statuses.map(x => `<option value="${esc(x)}">${esc(x)}</option>`).join(''));
}

['search','demandFilter','statusFilter','csatFilter','sortOrder'].forEach(id => {
  const el = document.getElementById(id);
  el.addEventListener(id === 'search' ? 'input' : 'change', renderSessions);
});

document.getElementById('clearFilters').addEventListener('click', () => {
  document.getElementById('search').value = '';
  document.getElementById('demandFilter').value = '';
  document.getElementById('statusFilter').value = '';
  document.getElementById('csatFilter').value = '';
  document.getElementById('sortOrder').value = 'recent';
  renderSessions();
});

populateFilters();
renderSessions();
if(selectedId) selectSession(selectedId);
</script>
</body>
</html>'''


def build_html(conversations: list[dict], output_path: Path) -> None:
    payload = json.dumps(conversations, ensure_ascii=False).replace("</script>", "<\\/script>")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(HTML_TEMPLATE.replace("__DATA__", payload), encoding="utf-8")


def main():
    parser = argparse.ArgumentParser(description="Gera um visualizador HTML de conversas a partir de um XLSX.")
    parser.add_argument("--input", default="data/conversas_ficticias.xlsx", help="Caminho do arquivo XLSX")
    parser.add_argument("--output", default="dist/index.html", help="Caminho do HTML de saída")
    parser.add_argument("--json", default="dist/conversas.json", help="Caminho opcional do JSON estruturado")
    args = parser.parse_args()

    input_path = Path(args.input)
    output_path = Path(args.output)
    json_path = Path(args.json)

    conversations = load_conversations(input_path)
    build_html(conversations, output_path)

    json_path.parent.mkdir(parents=True, exist_ok=True)
    json_path.write_text(json.dumps(conversations, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"{len(conversations)} conversas carregadas.")
    print(f"HTML: {output_path.resolve()}")
    print(f"JSON: {json_path.resolve()}")


if __name__ == "__main__":
    main()