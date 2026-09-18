# Chatbot Conversation Viewer — README completo

Este arquivo foi criado como uma **cópia de segurança autocontida de todo o código do projeto**. A ideia é que, mesmo que você não consiga navegar pelas pastas do repositório no GitHub corporativo, consiga abrir este README e copiar qualquer arquivo necessário.

## Visão geral

O projeto lê uma base `.xlsx` com interações de chatbot, agrupa as linhas por `id_sessao`, ordena por `seq_sessao`/`datahora` e gera um visualizador HTML que simula a conversa em uma interface de celular.

### Colunas esperadas no Excel

Obrigatórias:

- `id_sessao`
- `seq_sessao`
- `mensagem_cliente`
- `resposta`
- `datahora`

Opcionais / utilizadas pelo visualizador:

- `demanda`
- `termino_conversa`
- `nota_csat`
- `comentario_csat`
- `tipo_csat` — aceita `SMS` ou `CHAT IATI`
- `tool`
- `nome_intencao`

> `demanda`, `termino_conversa`, `nota_csat`, `comentario_csat` e `tipo_csat` são metadados de **sessão** e devem repetir o mesmo valor em todas as linhas do mesmo `id_sessao`. `tool` e `nome_intencao` pertencem à **interação** e podem variar a cada `seq_sessao`.

### Formato de botões e links na coluna `resposta`

Exemplo:

```text
Sobre o que você quer falar de pix?
Botão 1. aumentar limite do cartão
Botão 2. cadastrar chave pix
Link 1. ir pra página de pix
```

O parser transforma `Botão N.` e `Link N.` em opções visuais. Os links aparecem no mesmo bloco dos botões, usando uma seta `↗`.

## Dependências

Externas:

```bash
pip install pandas openpyxl
```

A biblioteca padrão do Python também é usada (`argparse`, `json`, `re`, `pathlib`), sem instalação adicional.

## Como executar

### Windows

```bat
gerar_viewer.bat
```

Ou manualmente:

```bat
python -m pip install -r requirements.txt
python src\build_viewer.py
```

### Linux/macOS

```bash
chmod +x gerar_viewer.sh
./gerar_viewer.sh
```

Depois, abra:

```text
dist/index.html
```

---

# Código completo do projeto

## `src/build_viewer.py`

Script principal: lê o XLSX, valida as colunas, estrutura as sessões e gera o JSON/HTML.

```python
from __future__ import annotations

import argparse
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
    "tipo_csat",
    "tool",
    "nome_intencao",
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
    """Split chatbot text into body and ordered interactive actions."""
    if pd.isna(value):
        return {"texto": "", "botoes": [], "links": [], "acoes": []}

    body, buttons, links, actions = [], [], [], []
    for raw_line in str(value).replace("\r\n", "\n").split("\n"):
        line = raw_line.strip()
        if not line:
            continue

        button_match = BUTTON_RE.match(line)
        if button_match:
            label = button_match.group(1).strip()
            buttons.append(label)
            actions.append({"tipo": "botao", "texto": label})
            continue

        link_match = LINK_RE.match(line)
        if link_match:
            label = link_match.group(1).strip()
            links.append(label)
            actions.append({"tipo": "link", "texto": label})
            continue

        body.append(line)

    return {
        "texto": "\n".join(body),
        "botoes": buttons,
        "links": links,
        "acoes": actions,
    }




def session_value(group: pd.DataFrame, column: str):
    """Retorna o valor da sessão e valida consistência entre as linhas."""
    values = []
    for value in group[column].tolist():
        if pd.isna(value):
            continue
        normalized = str(value).strip() if isinstance(value, str) else value
        if normalized == "":
            continue
        values.append(normalized)

    if not values:
        return None

    unique = []
    for value in values:
        if value not in unique:
            unique.append(value)

    if len(unique) > 1:
        session_id = "" if pd.isna(group.iloc[0]["id_sessao"]) else str(group.iloc[0]["id_sessao"])
        raise ValueError(
            f"A coluna '{column}' possui valores diferentes na sessão {session_id}: {unique}. "
            "Campos de sessão devem repetir o mesmo valor em todas as linhas do mesmo id_sessao."
        )
    return unique[0]

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

    tipos_validos = {"SMS", "CHAT IATI"}
    tipos_encontrados = {str(v).strip().upper() for v in df["tipo_csat"].dropna() if str(v).strip()}
    tipos_invalidos = sorted(tipos_encontrados - tipos_validos)
    if tipos_invalidos:
        raise ValueError(
            "Valores inválidos em tipo_csat: " + ", ".join(tipos_invalidos)
            + ". Use apenas SMS ou CHAT IATI."
        )

    df["datahora"] = pd.to_datetime(df["datahora"], errors="coerce")
    df["seq_sessao"] = pd.to_numeric(df["seq_sessao"], errors="coerce")
    df = df.sort_values(["id_sessao", "seq_sessao", "datahora"], kind="stable")

    conversations = []
    for session_id, group in df.groupby("id_sessao", sort=False, dropna=False):
        group = group.sort_values(["seq_sessao", "datahora"], kind="stable")
        first = group.iloc[0]

        # Metadados no nível da sessão: o mesmo valor deve aparecer em todas as
        # linhas do mesmo id_sessao.
        demanda = session_value(group, "demanda")
        termino_conversa = session_value(group, "termino_conversa")
        nota_csat = session_value(group, "nota_csat")
        comentario_csat = session_value(group, "comentario_csat")
        tipo_csat = session_value(group, "tipo_csat")

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
                "tool": "" if pd.isna(row.get("tool")) else str(row.get("tool")).strip(),
                "nome_intencao": "" if pd.isna(row.get("nome_intencao")) else str(row.get("nome_intencao")).strip(),
                "chatbot": parse_response(row["resposta"]),
            })

        csat = clean_value(nota_csat)
        try:
            if csat is not None and float(csat).is_integer():
                csat = int(float(csat))
        except Exception:
            pass

        conversations.append({
            "id_sessao": "" if pd.isna(session_id) else str(session_id),
            "demanda": "" if demanda is None else str(demanda),
            "termino_conversa": "" if termino_conversa is None else str(termino_conversa),
            "nota_csat": csat,
            "comentario_csat": "" if comentario_csat is None else str(comentario_csat),
            "tipo_csat": "" if tipo_csat is None else str(tipo_csat).strip().upper(),
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
.session:hover{background:#f7f8fa}.session.active{background:#fff4ed;border-color:#ffd4b8}.session-top{display:flex;align-items:center;justify-content:space-between;gap:8px}.session-id{font-weight:750;font-size:13px}.session-count{font-size:11px;color:var(--muted);background:#eef0f3;padding:3px 7px;border-radius:99px}.session-demand{font-size:12px;color:var(--muted);margin-top:4px;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;overflow-wrap:anywhere;word-break:break-word;line-height:1.35;min-height:16px}.session-meta{display:flex;align-items:center;justify-content:space-between;gap:8px;margin-top:8px}.session-date{font-size:10px;color:#8a9099}.session-meta-right{display:flex;align-items:center;gap:5px;min-width:0}.mini-type{font-size:8px;font-weight:850;padding:3px 6px;border-radius:999px;background:#f0f2f5;color:#59606b;white-space:nowrap}.mini-status{font-size:9px;font-weight:800;text-transform:capitalize;padding:3px 6px;border-radius:999px;white-space:nowrap}.mini-status.resolutivo{background:#eaf7ee;color:var(--success)}.mini-status.transbordo{background:#fff4e5;color:var(--warning)}.mini-status.abandono{background:#fdecec;color:var(--danger)}
.stage{display:flex;align-items:center;justify-content:center;min-height:0}
.phone{width:430px;height:min(880px,calc(100vh - 56px));min-height:640px;background:white;border:9px solid #1f2228;border-radius:50px;box-shadow:var(--shadow);overflow:hidden;position:relative;display:flex;flex-direction:column}
.phone::after{content:"";position:absolute;inset:5px;border:1px solid rgba(255,255,255,.18);border-radius:38px;pointer-events:none}
.statusbar{height:46px;display:flex;align-items:center;justify-content:space-between;padding:0 24px;font-weight:750;font-size:14px;flex:0 0 auto}.status-icons{font-size:12px;letter-spacing:2px}
.chat-head{height:54px;display:flex;align-items:center;justify-content:space-between;padding:0 20px;border-bottom:1px solid #f0f1f3;flex:0 0 auto}.chat-title{display:flex;align-items:center;gap:10px}.spark{color:var(--orange);font-size:24px;line-height:1}.chat-name{font-weight:760}.chat-sub{font-size:11px;color:var(--muted);margin-top:1px}.close{font-size:28px;color:#363942;font-weight:300}
.chat{flex:1;overflow:auto;padding:18px 20px 28px;background:linear-gradient(#fff 0%,#fff 58%,#fff7f1 100%);scroll-behavior:smooth}.chat::-webkit-scrollbar{width:5px}.chat::-webkit-scrollbar-thumb{background:#d9dde3;border-radius:10px}
.day{text-align:center;color:var(--muted);font-size:12px;margin:8px 0 22px}
.turn{position:relative;margin-bottom:22px}.client-row{display:flex;justify-content:flex-end;gap:9px;align-items:flex-end}.avatar{width:38px;height:38px;border-radius:50%;background:#333840;color:white;display:grid;place-items:center;font-size:13px;flex:0 0 auto}.client-wrap{max-width:78%;display:flex;flex-direction:column;align-items:flex-end}.bubble-client{background:var(--bubble);padding:12px 15px;border-radius:18px 18px 5px 18px;line-height:1.42;font-size:14px;white-space:pre-wrap}.time{font-size:11px;color:var(--muted);margin-top:5px;display:flex;gap:5px;align-items:center}.checks{letter-spacing:-2px;color:#68707c;font-weight:700}
.bot-row{display:grid;grid-template-columns:26px 1fr;gap:9px;margin-top:20px}.bot-icon{color:var(--orange);font-size:22px;margin-top:1px}.bot-content{min-width:0}.bot-text{font-size:14px;line-height:1.55;white-space:pre-wrap}.choices{border:1px solid #dfe2e7;border-radius:16px;overflow:hidden;margin-top:13px;background:rgba(255,255,255,.92)}.choice{display:flex;align-items:center;justify-content:space-between;gap:10px;padding:13px 14px;border-bottom:1px solid #eceef1;font-weight:700;font-size:13px}.choice:last-child{border-bottom:0}.choice.link-choice{background:#fffaf6}.chev{font-size:20px;color:#717782;font-weight:400;line-height:1}.link-arrow{font-size:15px;color:#9a4300;font-weight:800;line-height:1}
.turn-tooltip{position:absolute;z-index:30;left:4px;top:-8px;min-width:220px;max-width:310px;background:#242833;color:#fff;border-radius:14px;padding:11px 12px;box-shadow:0 10px 28px rgba(25,28,38,.22);opacity:0;visibility:hidden;transform:translateY(-5px);transition:.15s ease;pointer-events:none}.turn:hover .turn-tooltip{opacity:1;visibility:visible;transform:none}.tooltip-title{font-size:11px;font-weight:850;margin-bottom:8px;color:#ffd5ba}.tooltip-grid{display:grid;grid-template-columns:70px 1fr;gap:5px 9px;font-size:10px;line-height:1.35}.tooltip-label{color:#bfc4cc;font-weight:700}.tooltip-value{color:#fff;overflow-wrap:anywhere;word-break:break-word}
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
      <select id="tipoCsatFilter" class="filter"><option value="">Todos os tipos CSAT</option><option value="SMS">SMS</option><option value="CHAT IATI">CHAT IATI</option></select>
      <select id="csatFilter" class="filter"><option value="">Todos os CSATs</option><option value="5">CSAT 5</option><option value="4">CSAT 4</option><option value="3">CSAT 3</option><option value="2">CSAT 2</option><option value="1">CSAT 1</option><option value="none">Sem CSAT</option></select>
      <select id="sortOrder" class="filter filter-wide"><option value="recent">Mais recentes</option><option value="oldest">Mais antigas</option><option value="id-asc">ID crescente</option><option value="id-desc">ID decrescente</option></select>
    </div>
    <div class="session-toolbar"><span id="resultsCount" class="results-count"></span><button id="clearFilters" class="clear-filters" type="button">Limpar filtros</button></div>
    <div id="sessionList" class="session-list"></div>
  </aside>

  <main class="stage">
    <section class="phone" aria-label="Simulador de conversa em celular">
      <div class="statusbar"><span id="phoneClock">21:30</span><span class="status-icons">▦ ))) ▰</span></div>
      <div class="chat-head">
        <div class="chat-title"><span class="spark">✦</span><div><div class="chat-name">Inteligência de Atendimento Itaú</div><div class="chat-sub" id="headSession">Sessão</div></div></div>
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
    tipoCsat: document.getElementById('tipoCsatFilter').value,
    csat: document.getElementById('csatFilter').value,
    sort: document.getElementById('sortOrder').value
  };
}

function renderSessions(){
  const {q,demand,status,tipoCsat,csat,sort} = currentFilters();
  const list = document.getElementById('sessionList');
  let items = conversations.filter(c => {
    const matchesText = !q || c.id_sessao.toLowerCase().includes(q) || (c.demanda||'').toLowerCase().includes(q);
    const matchesDemand = !demand || (c.demanda||'') === demand;
    const matchesStatus = !status || (c.termino_conversa||'') === status;
    const matchesTipoCsat = !tipoCsat || (c.tipo_csat||'') === tipoCsat;
    const matchesCsat = !csat || (csat === 'none' ? c.nota_csat == null : String(c.nota_csat) === csat);
    return matchesText && matchesDemand && matchesStatus && matchesTipoCsat && matchesCsat;
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
      <div class="session-demand" title="${esc(c.demanda || 'Sem demanda')}">${esc(c.demanda || 'Sem demanda')}</div>
      <div class="session-meta"><span class="session-date">${esc(sessionDateLabel(c))}</span><span class="session-meta-right"><span class="mini-type">${esc(c.tipo_csat || '—')}</span><span class="mini-status ${statusClass(c.termino_conversa)}">${esc(c.termino_conversa || '—')}</span></span></div>
    </button>`).join('') || '<div class="empty">Nenhuma sessão encontrada.</div>';
  list.querySelectorAll('.session').forEach(btn => btn.addEventListener('click', () => selectSession(btn.dataset.id)));
}

function renderBot(interaction){
  const bot = interaction.chatbot || {};
  const fallbackActions = [
    ...(bot.botoes || []).map(texto => ({tipo:'botao', texto})),
    ...(bot.links || []).map(texto => ({tipo:'link', texto}))
  ];
  const actions = (bot.acoes && bot.acoes.length ? bot.acoes : fallbackActions);
  const choices = actions.map(action => {
    const isLink = action.tipo === 'link';
    const icon = isLink ? '<span class="link-arrow">↗</span>' : '<span class="chev">›</span>';
    return `<div class="choice ${isLink ? 'link-choice' : ''}"><span>${esc(action.texto)}</span>${icon}</div>`;
  }).join('');
  return `
    <div class="bot-row">
      <div class="bot-icon">✦</div>
      <div class="bot-content">
        ${bot.texto ? `<div class="bot-text">${esc(bot.texto)}</div>` : ''}
        ${choices ? `<div class="choices">${choices}</div>` : ''}
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
      <div class="turn-tooltip"><div class="tooltip-title">Interação #${esc(i.seq_sessao ?? '—')}</div><div class="tooltip-grid"><span class="tooltip-label">Tool</span><span class="tooltip-value">${esc(i.tool || '—')}</span><span class="tooltip-label">Intenção</span><span class="tooltip-value">${esc(i.nome_intencao || '—')}</span></div></div>
      ${i.cliente ? `<div class="client-row"><div class="client-wrap"><div class="bubble-client">${esc(i.cliente)}</div><div class="time">${esc(i.hora||'')} <span class="checks">✓✓</span></div></div><div class="avatar">VC</div></div>` : ''}
      ${renderBot(i)}
    </div>`).join('');
  chat.scrollTop = 0;
  const lastTime = c.interacoes?.at(-1)?.hora || c.interacoes?.[0]?.hora || '21:30';
  document.getElementById('phoneClock').textContent = lastTime;
  document.getElementById('headSession').textContent = c.id_sessao || 'Sessão';
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
    <div class="meta-card"><div class="label">Tipo CSAT</div><div class="value">${esc(c.tipo_csat || '—')}</div></div>
    <div class="meta-card"><div class="label">CSAT</div><div class="value">${esc(score)}</div><div class="stars">${esc(stars)}</div></div>
    <div class="meta-card"><div class="label">Comentário CSAT</div><div class="comment">${esc(c.comentario_csat || 'Sem comentário.')}</div></div>
    <div class="footer-note">Passe o mouse sobre uma interação para visualizar <strong>seq_sessao</strong>, <strong>tool</strong> e <strong>nome_intencao</strong>. Os campos demanda, término, CSAT, comentário e tipo de CSAT pertencem à sessão e se repetem nas linhas do mesmo <strong>id_sessao</strong>.</div>`;
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

['search','demandFilter','statusFilter','tipoCsatFilter','csatFilter','sortOrder'].forEach(id => {
  const el = document.getElementById(id);
  el.addEventListener(id === 'search' ? 'input' : 'change', renderSessions);
});

document.getElementById('clearFilters').addEventListener('click', () => {
  document.getElementById('search').value = '';
  document.getElementById('demandFilter').value = '';
  document.getElementById('statusFilter').value = '';
  document.getElementById('tipoCsatFilter').value = '';
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

```

## `dist/index.html`

HTML final gerado pelo projeto. Pode ser aberto diretamente no navegador.

```html
<!doctype html>
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
.session:hover{background:#f7f8fa}.session.active{background:#fff4ed;border-color:#ffd4b8}.session-top{display:flex;align-items:center;justify-content:space-between;gap:8px}.session-id{font-weight:750;font-size:13px}.session-count{font-size:11px;color:var(--muted);background:#eef0f3;padding:3px 7px;border-radius:99px}.session-demand{font-size:12px;color:var(--muted);margin-top:4px;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;overflow-wrap:anywhere;word-break:break-word;line-height:1.35;min-height:16px}.session-meta{display:flex;align-items:center;justify-content:space-between;gap:8px;margin-top:8px}.session-date{font-size:10px;color:#8a9099}.session-meta-right{display:flex;align-items:center;gap:5px;min-width:0}.mini-type{font-size:8px;font-weight:850;padding:3px 6px;border-radius:999px;background:#f0f2f5;color:#59606b;white-space:nowrap}.mini-status{font-size:9px;font-weight:800;text-transform:capitalize;padding:3px 6px;border-radius:999px;white-space:nowrap}.mini-status.resolutivo{background:#eaf7ee;color:var(--success)}.mini-status.transbordo{background:#fff4e5;color:var(--warning)}.mini-status.abandono{background:#fdecec;color:var(--danger)}
.stage{display:flex;align-items:center;justify-content:center;min-height:0}
.phone{width:430px;height:min(880px,calc(100vh - 56px));min-height:640px;background:white;border:9px solid #1f2228;border-radius:50px;box-shadow:var(--shadow);overflow:hidden;position:relative;display:flex;flex-direction:column}
.phone::after{content:"";position:absolute;inset:5px;border:1px solid rgba(255,255,255,.18);border-radius:38px;pointer-events:none}
.statusbar{height:46px;display:flex;align-items:center;justify-content:space-between;padding:0 24px;font-weight:750;font-size:14px;flex:0 0 auto}.status-icons{font-size:12px;letter-spacing:2px}
.chat-head{height:54px;display:flex;align-items:center;justify-content:space-between;padding:0 20px;border-bottom:1px solid #f0f1f3;flex:0 0 auto}.chat-title{display:flex;align-items:center;gap:10px}.spark{color:var(--orange);font-size:24px;line-height:1}.chat-name{font-weight:760}.chat-sub{font-size:11px;color:var(--muted);margin-top:1px}.close{font-size:28px;color:#363942;font-weight:300}
.chat{flex:1;overflow:auto;padding:18px 20px 28px;background:linear-gradient(#fff 0%,#fff 58%,#fff7f1 100%);scroll-behavior:smooth}.chat::-webkit-scrollbar{width:5px}.chat::-webkit-scrollbar-thumb{background:#d9dde3;border-radius:10px}
.day{text-align:center;color:var(--muted);font-size:12px;margin:8px 0 22px}
.turn{position:relative;margin-bottom:22px}.client-row{display:flex;justify-content:flex-end;gap:9px;align-items:flex-end}.avatar{width:38px;height:38px;border-radius:50%;background:#333840;color:white;display:grid;place-items:center;font-size:13px;flex:0 0 auto}.client-wrap{max-width:78%;display:flex;flex-direction:column;align-items:flex-end}.bubble-client{background:var(--bubble);padding:12px 15px;border-radius:18px 18px 5px 18px;line-height:1.42;font-size:14px;white-space:pre-wrap}.time{font-size:11px;color:var(--muted);margin-top:5px;display:flex;gap:5px;align-items:center}.checks{letter-spacing:-2px;color:#68707c;font-weight:700}
.bot-row{display:grid;grid-template-columns:26px 1fr;gap:9px;margin-top:20px}.bot-icon{color:var(--orange);font-size:22px;margin-top:1px}.bot-content{min-width:0}.bot-text{font-size:14px;line-height:1.55;white-space:pre-wrap}.choices{border:1px solid #dfe2e7;border-radius:16px;overflow:hidden;margin-top:13px;background:rgba(255,255,255,.92)}.choice{display:flex;align-items:center;justify-content:space-between;gap:10px;padding:13px 14px;border-bottom:1px solid #eceef1;font-weight:700;font-size:13px}.choice:last-child{border-bottom:0}.choice.link-choice{background:#fffaf6}.chev{font-size:20px;color:#717782;font-weight:400;line-height:1}.link-arrow{font-size:15px;color:#9a4300;font-weight:800;line-height:1}
.turn-tooltip{position:absolute;z-index:30;left:4px;top:-8px;min-width:220px;max-width:310px;background:#242833;color:#fff;border-radius:14px;padding:11px 12px;box-shadow:0 10px 28px rgba(25,28,38,.22);opacity:0;visibility:hidden;transform:translateY(-5px);transition:.15s ease;pointer-events:none}.turn:hover .turn-tooltip{opacity:1;visibility:visible;transform:none}.tooltip-title{font-size:11px;font-weight:850;margin-bottom:8px;color:#ffd5ba}.tooltip-grid{display:grid;grid-template-columns:70px 1fr;gap:5px 9px;font-size:10px;line-height:1.35}.tooltip-label{color:#bfc4cc;font-weight:700}.tooltip-value{color:#fff;overflow-wrap:anywhere;word-break:break-word}
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
      <select id="tipoCsatFilter" class="filter"><option value="">Todos os tipos CSAT</option><option value="SMS">SMS</option><option value="CHAT IATI">CHAT IATI</option></select>
      <select id="csatFilter" class="filter"><option value="">Todos os CSATs</option><option value="5">CSAT 5</option><option value="4">CSAT 4</option><option value="3">CSAT 3</option><option value="2">CSAT 2</option><option value="1">CSAT 1</option><option value="none">Sem CSAT</option></select>
      <select id="sortOrder" class="filter filter-wide"><option value="recent">Mais recentes</option><option value="oldest">Mais antigas</option><option value="id-asc">ID crescente</option><option value="id-desc">ID decrescente</option></select>
    </div>
    <div class="session-toolbar"><span id="resultsCount" class="results-count"></span><button id="clearFilters" class="clear-filters" type="button">Limpar filtros</button></div>
    <div id="sessionList" class="session-list"></div>
  </aside>

  <main class="stage">
    <section class="phone" aria-label="Simulador de conversa em celular">
      <div class="statusbar"><span id="phoneClock">21:30</span><span class="status-icons">▦ ))) ▰</span></div>
      <div class="chat-head">
        <div class="chat-title"><span class="spark">✦</span><div><div class="chat-name">Inteligência de Atendimento Itaú</div><div class="chat-sub" id="headSession">Sessão</div></div></div>
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
const conversations = [{"id_sessao": "SES-1001", "demanda": "Cadastro e gerenciamento de chave Pix + análise basal e recursiva", "termino_conversa": "resolutivo", "nota_csat": 5, "comentario_csat": "Consegui cadastrar a chave sem dificuldade.", "tipo_csat": "CHAT IATI", "qtd_interacoes": 5, "interacoes": [{"seq_sessao": 1, "datahora": "2026-09-17T08:17:00", "hora": "08:17", "cliente": "pix", "tool": "kb-global", "nome_intencao": "menu_pix", "chatbot": {"texto": "Sobre o que você quer falar de pix?", "botoes": ["aumentar limite do pix", "cadastrar chave pix"], "links": ["ir pra página de pix"], "acoes": [{"tipo": "botao", "texto": "aumentar limite do pix"}, {"tipo": "botao", "texto": "cadastrar chave pix"}, {"tipo": "link", "texto": "ir pra página de pix"}]}}, {"seq_sessao": 2, "datahora": "2026-09-17T08:19:00", "hora": "08:19", "cliente": "cadastrar chave pix", "tool": "pix-orchestrator", "nome_intencao": "cadastro_chave_pix", "chatbot": {"texto": "Você pode cadastrar uma nova chave Pix pelo app.", "botoes": ["cadastrar agora"], "links": ["entender tipos de chave"], "acoes": [{"tipo": "botao", "texto": "cadastrar agora"}, {"tipo": "link", "texto": "entender tipos de chave"}]}}, {"seq_sessao": 3, "datahora": "2026-09-17T08:21:00", "hora": "08:21", "cliente": "quero usar meu celular", "tool": "pix-key-api", "nome_intencao": "chave_pix_celular", "chatbot": {"texto": "Perfeito. O número de celular pode ser usado como chave Pix.", "botoes": ["continuar cadastro"], "links": [], "acoes": [{"tipo": "botao", "texto": "continuar cadastro"}]}}, {"seq_sessao": 4, "datahora": "2026-09-17T08:23:00", "hora": "08:23", "cliente": "continuar", "tool": "pix-key-api", "nome_intencao": "confirmacao_dados", "chatbot": {"texto": "Antes de concluir, confirme se o número exibido no app está correto.", "botoes": ["confirmar número", "alterar número"], "links": [], "acoes": [{"tipo": "botao", "texto": "confirmar número"}, {"tipo": "botao", "texto": "alterar número"}]}}, {"seq_sessao": 5, "datahora": "2026-09-17T08:25:00", "hora": "08:25", "cliente": "confirmar número", "tool": "pix-key-api", "nome_intencao": "cadastro_pix_sucesso", "chatbot": {"texto": "Tudo certo. A solicitação de cadastro foi iniciada e você acompanhará o status pelo app.", "botoes": [], "links": [], "acoes": []}}]}, {"id_sessao": "SES-1002", "demanda": "Contestação de compra não reconhecida no cartão de crédito", "termino_conversa": "transbordo", "nota_csat": 3, "comentario_csat": "Precisou chamar um atendente.", "tipo_csat": "SMS", "qtd_interacoes": 4, "interacoes": [{"seq_sessao": 1, "datahora": "2026-09-17T09:17:00", "hora": "09:17", "cliente": "não reconheço uma compra", "tool": "cards-api", "nome_intencao": "contestacao_compra", "chatbot": {"texto": "Posso te ajudar a analisar essa compra.", "botoes": ["ver compras recentes", "falar com atendente"], "links": [], "acoes": [{"tipo": "botao", "texto": "ver compras recentes"}, {"tipo": "botao", "texto": "falar com atendente"}]}}, {"seq_sessao": 2, "datahora": "2026-09-17T09:19:00", "hora": "09:19", "cliente": "ver compras recentes", "tool": "cards-api", "nome_intencao": "lista_compras", "chatbot": {"texto": "Encontrei compras recentes no seu cartão.", "botoes": ["compra de hoje", "compra de ontem"], "links": [], "acoes": [{"tipo": "botao", "texto": "compra de hoje"}, {"tipo": "botao", "texto": "compra de ontem"}]}}, {"seq_sessao": 3, "datahora": "2026-09-17T09:21:00", "hora": "09:21", "cliente": "compra de ontem", "tool": "cards-api", "nome_intencao": "detalhe_compra", "chatbot": {"texto": "Essa compra foi feita em estabelecimento físico. Você reconhece?", "botoes": ["sim", "não"], "links": [], "acoes": [{"tipo": "botao", "texto": "sim"}, {"tipo": "botao", "texto": "não"}]}}, {"seq_sessao": 4, "datahora": "2026-09-17T09:23:00", "hora": "09:23", "cliente": "não", "tool": "handoff-api", "nome_intencao": "transbordo_contestacao", "chatbot": {"texto": "Vou encaminhar seu caso para um especialista continuar a contestação.", "botoes": [], "links": [], "acoes": []}}]}, {"id_sessao": "SES-1003", "demanda": "Emissão de segunda via de boleto de financiamento", "termino_conversa": "resolutivo", "nota_csat": 5, "comentario_csat": "Boleto encontrado rapidamente.", "tipo_csat": "CHAT IATI", "qtd_interacoes": 3, "interacoes": [{"seq_sessao": 1, "datahora": "2026-09-17T10:17:00", "hora": "10:17", "cliente": "preciso do boleto", "tool": "billing-router", "nome_intencao": "menu_boleto", "chatbot": {"texto": "Qual boleto você procura?", "botoes": ["financiamento", "cartão"], "links": [], "acoes": [{"tipo": "botao", "texto": "financiamento"}, {"tipo": "botao", "texto": "cartão"}]}}, {"seq_sessao": 2, "datahora": "2026-09-17T10:19:00", "hora": "10:19", "cliente": "financiamento", "tool": "financing-api", "nome_intencao": "boleto_financiamento", "chatbot": {"texto": "Encontrei seu contrato ativo.", "botoes": ["boleto deste mês"], "links": ["ver parcelas do contrato"], "acoes": [{"tipo": "botao", "texto": "boleto deste mês"}, {"tipo": "link", "texto": "ver parcelas do contrato"}]}}, {"seq_sessao": 3, "datahora": "2026-09-17T10:21:00", "hora": "10:21", "cliente": "boleto deste mês", "tool": "financing-api", "nome_intencao": "segunda_via_boleto", "chatbot": {"texto": "Seu boleto está disponível.", "botoes": ["copiar código de barras"], "links": ["abrir boleto"], "acoes": [{"tipo": "link", "texto": "abrir boleto"}, {"tipo": "botao", "texto": "copiar código de barras"}]}}]}, {"id_sessao": "SES-1004", "demanda": "Aumento de limite do cartão considerando análise de crédito", "termino_conversa": "abandono", "nota_csat": null, "comentario_csat": "", "tipo_csat": "SMS", "qtd_interacoes": 3, "interacoes": [{"seq_sessao": 1, "datahora": "2026-09-17T11:17:00", "hora": "11:17", "cliente": "aumentar limite", "tool": "credit-api", "nome_intencao": "consulta_limite", "chatbot": {"texto": "Posso verificar as opções disponíveis para seu cartão.", "botoes": ["consultar limite disponível", "entender análise de crédito"], "links": [], "acoes": [{"tipo": "botao", "texto": "consultar limite disponível"}, {"tipo": "botao", "texto": "entender análise de crédito"}]}}, {"seq_sessao": 2, "datahora": "2026-09-17T11:19:00", "hora": "11:19", "cliente": "consultar limite disponível", "tool": "credit-api", "nome_intencao": "limite_sem_oferta", "chatbot": {"texto": "No momento não há aumento pré-aprovado.", "botoes": ["receber dicas para aumentar limite"], "links": ["atualizar renda"], "acoes": [{"tipo": "link", "texto": "atualizar renda"}, {"tipo": "botao", "texto": "receber dicas para aumentar limite"}]}}, {"seq_sessao": 3, "datahora": "2026-09-17T11:21:00", "hora": "11:21", "cliente": "receber dicas", "tool": "credit-api", "nome_intencao": "orientacao_limite", "chatbot": {"texto": "Manter renda atualizada e bom histórico de pagamento pode ajudar em futuras análises.", "botoes": [], "links": [], "acoes": []}}]}, {"id_sessao": "SES-1005", "demanda": "Seguro Pet: coberturas, contratação e assistência veterinária", "termino_conversa": "resolutivo", "nota_csat": 4, "comentario_csat": "Explicação clara sobre as coberturas.", "tipo_csat": "CHAT IATI", "qtd_interacoes": 5, "interacoes": [{"seq_sessao": 1, "datahora": "2026-09-17T12:17:00", "hora": "12:17", "cliente": "seguro pet", "tool": "insurance-kb", "nome_intencao": "menu_seguro_pet", "chatbot": {"texto": "O Seguro Pet oferece assistência para cães e gatos.", "botoes": ["ver coberturas", "contratar"], "links": ["página do Seguro Pet"], "acoes": [{"tipo": "botao", "texto": "ver coberturas"}, {"tipo": "botao", "texto": "contratar"}, {"tipo": "link", "texto": "página do Seguro Pet"}]}}, {"seq_sessao": 2, "datahora": "2026-09-17T12:19:00", "hora": "12:19", "cliente": "ver coberturas", "tool": "insurance-api", "nome_intencao": "coberturas_pet", "chatbot": {"texto": "As coberturas variam conforme o plano e podem incluir consultas, exames e vacinas.", "botoes": ["comparar planos"], "links": [], "acoes": [{"tipo": "botao", "texto": "comparar planos"}]}}, {"seq_sessao": 3, "datahora": "2026-09-17T12:21:00", "hora": "12:21", "cliente": "comparar planos", "tool": "insurance-api", "nome_intencao": "comparacao_planos_pet", "chatbot": {"texto": "Você pode comparar os planos disponíveis para o seu pet.", "botoes": [], "links": ["comparar planos"], "acoes": [{"tipo": "link", "texto": "comparar planos"}]}}, {"seq_sessao": 4, "datahora": "2026-09-17T12:23:00", "hora": "12:23", "cliente": "quero contratar", "tool": "insurance-api", "nome_intencao": "contratacao_pet", "chatbot": {"texto": "Certo. Vou iniciar a jornada de contratação.", "botoes": ["começar"], "links": [], "acoes": [{"tipo": "botao", "texto": "começar"}]}}, {"seq_sessao": 5, "datahora": "2026-09-17T12:25:00", "hora": "12:25", "cliente": "começar", "tool": "insurance-api", "nome_intencao": "contratacao_pet_sucesso", "chatbot": {"texto": "Pronto. A contratação foi iniciada e você pode continuar no app.", "botoes": [], "links": [], "acoes": []}}]}, {"id_sessao": "SES-1006", "demanda": "Simulação de empréstimo pessoal e condições de pagamento", "termino_conversa": "resolutivo", "nota_csat": 5, "comentario_csat": "Gostei da simulação.", "tipo_csat": "SMS", "qtd_interacoes": 5, "interacoes": [{"seq_sessao": 1, "datahora": "2026-09-17T13:17:00", "hora": "13:17", "cliente": "quero empréstimo", "tool": "loans-api", "nome_intencao": "menu_emprestimo", "chatbot": {"texto": "Posso consultar opções de empréstimo pessoal.", "botoes": ["simular empréstimo"], "links": ["entender taxas"], "acoes": [{"tipo": "botao", "texto": "simular empréstimo"}, {"tipo": "link", "texto": "entender taxas"}]}}, {"seq_sessao": 2, "datahora": "2026-09-17T13:19:00", "hora": "13:19", "cliente": "simular empréstimo", "tool": "loans-api", "nome_intencao": "valor_simulacao", "chatbot": {"texto": "Qual valor você deseja simular?", "botoes": ["R$ 2.000", "R$ 5.000", "outro valor"], "links": [], "acoes": [{"tipo": "botao", "texto": "R$ 2.000"}, {"tipo": "botao", "texto": "R$ 5.000"}, {"tipo": "botao", "texto": "outro valor"}]}}, {"seq_sessao": 3, "datahora": "2026-09-17T13:21:00", "hora": "13:21", "cliente": "R$ 5.000", "tool": "loans-api", "nome_intencao": "prazo_simulacao", "chatbot": {"texto": "Você pode escolher o prazo.", "botoes": ["12 meses", "24 meses"], "links": [], "acoes": [{"tipo": "botao", "texto": "12 meses"}, {"tipo": "botao", "texto": "24 meses"}]}}, {"seq_sessao": 4, "datahora": "2026-09-17T13:23:00", "hora": "13:23", "cliente": "24 meses", "tool": "loans-pricing", "nome_intencao": "resultado_simulacao", "chatbot": {"texto": "A simulação foi calculada com as condições disponíveis no momento.", "botoes": ["contratar"], "links": ["ver simulação completa"], "acoes": [{"tipo": "link", "texto": "ver simulação completa"}, {"tipo": "botao", "texto": "contratar"}]}}, {"seq_sessao": 5, "datahora": "2026-09-17T13:25:00", "hora": "13:25", "cliente": "não vou contratar agora", "tool": "loans-api", "nome_intencao": "encerramento_simulacao", "chatbot": {"texto": "Sem problema. Você pode refazer a simulação quando quiser.", "botoes": [], "links": [], "acoes": []}}]}, {"id_sessao": "SES-1007", "demanda": "Recuperação de senha e desbloqueio de acesso ao aplicativo", "termino_conversa": "resolutivo", "nota_csat": 4, "comentario_csat": "Consegui redefinir a senha.", "tipo_csat": "CHAT IATI", "qtd_interacoes": 4, "interacoes": [{"seq_sessao": 1, "datahora": "2026-09-17T14:17:00", "hora": "14:17", "cliente": "esqueci minha senha", "tool": "auth-api", "nome_intencao": "menu_acesso", "chatbot": {"texto": "Posso te orientar na recuperação de acesso.", "botoes": ["recuperar senha", "desbloquear acesso"], "links": [], "acoes": [{"tipo": "botao", "texto": "recuperar senha"}, {"tipo": "botao", "texto": "desbloquear acesso"}]}}, {"seq_sessao": 2, "datahora": "2026-09-17T14:19:00", "hora": "14:19", "cliente": "recuperar senha", "tool": "auth-api", "nome_intencao": "recuperacao_senha", "chatbot": {"texto": "Vamos validar alguns dados de segurança pelo app.", "botoes": ["continuar"], "links": [], "acoes": [{"tipo": "botao", "texto": "continuar"}]}}, {"seq_sessao": 3, "datahora": "2026-09-17T14:21:00", "hora": "14:21", "cliente": "continuar", "tool": "auth-api", "nome_intencao": "validacao_seguranca", "chatbot": {"texto": "Validação concluída.", "botoes": ["criar nova senha"], "links": [], "acoes": [{"tipo": "botao", "texto": "criar nova senha"}]}}, {"seq_sessao": 4, "datahora": "2026-09-17T14:23:00", "hora": "14:23", "cliente": "criar nova senha", "tool": "auth-api", "nome_intencao": "senha_redefinida", "chatbot": {"texto": "Sua senha foi redefinida com sucesso.", "botoes": [], "links": [], "acoes": []}}]}, {"id_sessao": "SES-1008", "demanda": "Consulta de investimentos e comparação entre produtos de renda fixa", "termino_conversa": "abandono", "nota_csat": null, "comentario_csat": "", "tipo_csat": "CHAT IATI", "qtd_interacoes": 3, "interacoes": [{"seq_sessao": 1, "datahora": "2026-09-17T15:17:00", "hora": "15:17", "cliente": "quero investir", "tool": "investments-router", "nome_intencao": "menu_investimentos", "chatbot": {"texto": "O que você quer fazer?", "botoes": ["ver investimentos", "comparar renda fixa"], "links": [], "acoes": [{"tipo": "botao", "texto": "ver investimentos"}, {"tipo": "botao", "texto": "comparar renda fixa"}]}}, {"seq_sessao": 2, "datahora": "2026-09-17T15:19:00", "hora": "15:19", "cliente": "comparar renda fixa", "tool": "investments-api", "nome_intencao": "comparacao_renda_fixa", "chatbot": {"texto": "Posso mostrar opções considerando prazo e liquidez.", "botoes": ["liquidez diária", "prazo acima de 1 ano"], "links": [], "acoes": [{"tipo": "botao", "texto": "liquidez diária"}, {"tipo": "botao", "texto": "prazo acima de 1 ano"}]}}, {"seq_sessao": 3, "datahora": "2026-09-17T15:21:00", "hora": "15:21", "cliente": "liquidez diária", "tool": "investments-api", "nome_intencao": "renda_fixa_liquidez", "chatbot": {"texto": "Existem produtos com liquidez diária disponíveis para consulta.", "botoes": [], "links": ["ver opções no app"], "acoes": [{"tipo": "link", "texto": "ver opções no app"}]}}]}, {"id_sessao": "SES-1009", "demanda": "Atualização de dados cadastrais e endereço da conta corrente", "termino_conversa": "resolutivo", "nota_csat": 5, "comentario_csat": "Atualização concluída.", "tipo_csat": "SMS", "qtd_interacoes": 4, "interacoes": [{"seq_sessao": 1, "datahora": "2026-09-17T16:17:00", "hora": "16:17", "cliente": "mudar meu endereço", "tool": "profile-api", "nome_intencao": "menu_dados_cadastrais", "chatbot": {"texto": "Você pode atualizar seus dados cadastrais pelo app.", "botoes": ["atualizar endereço"], "links": ["ver dados cadastrais"], "acoes": [{"tipo": "botao", "texto": "atualizar endereço"}, {"tipo": "link", "texto": "ver dados cadastrais"}]}}, {"seq_sessao": 2, "datahora": "2026-09-17T16:19:00", "hora": "16:19", "cliente": "atualizar endereço", "tool": "profile-api", "nome_intencao": "atualizacao_endereco", "chatbot": {"texto": "Digite o novo CEP na etapa seguinte do app.", "botoes": ["continuar"], "links": [], "acoes": [{"tipo": "botao", "texto": "continuar"}]}}, {"seq_sessao": 3, "datahora": "2026-09-17T16:21:00", "hora": "16:21", "cliente": "continuar", "tool": "address-api", "nome_intencao": "confirmacao_endereco", "chatbot": {"texto": "Endereço localizado. Confira os dados antes de confirmar.", "botoes": ["confirmar endereço", "corrigir"], "links": [], "acoes": [{"tipo": "botao", "texto": "confirmar endereço"}, {"tipo": "botao", "texto": "corrigir"}]}}, {"seq_sessao": 4, "datahora": "2026-09-17T16:23:00", "hora": "16:23", "cliente": "confirmar endereço", "tool": "profile-api", "nome_intencao": "endereco_atualizado", "chatbot": {"texto": "Seu endereço foi atualizado com sucesso.", "botoes": [], "links": [], "acoes": []}}]}, {"id_sessao": "SES-1010", "demanda": "Rastreamento e prazo de entrega de novo cartão físico", "termino_conversa": "transbordo", "nota_csat": 2, "comentario_csat": "O rastreio não mostrou onde estava o cartão.", "tipo_csat": "CHAT IATI", "qtd_interacoes": 5, "interacoes": [{"seq_sessao": 1, "datahora": "2026-09-17T17:17:00", "hora": "17:17", "cliente": "onde está meu cartão", "tool": "delivery-api", "nome_intencao": "menu_entrega_cartao", "chatbot": {"texto": "Posso consultar a entrega do seu cartão.", "botoes": ["rastrear cartão", "prazo de entrega"], "links": [], "acoes": [{"tipo": "botao", "texto": "rastrear cartão"}, {"tipo": "botao", "texto": "prazo de entrega"}]}}, {"seq_sessao": 2, "datahora": "2026-09-17T17:19:00", "hora": "17:19", "cliente": "rastrear cartão", "tool": "delivery-api", "nome_intencao": "rastreio_cartao", "chatbot": {"texto": "Seu cartão saiu para entrega.", "botoes": ["não recebi"], "links": ["acompanhar entrega"], "acoes": [{"tipo": "link", "texto": "acompanhar entrega"}, {"tipo": "botao", "texto": "não recebi"}]}}, {"seq_sessao": 3, "datahora": "2026-09-17T17:21:00", "hora": "17:21", "cliente": "não recebi", "tool": "delivery-api", "nome_intencao": "falha_entrega", "chatbot": {"texto": "A entrega consta como tentativa sem sucesso.", "botoes": ["confirmar endereço", "falar com atendente"], "links": [], "acoes": [{"tipo": "botao", "texto": "confirmar endereço"}, {"tipo": "botao", "texto": "falar com atendente"}]}}, {"seq_sessao": 4, "datahora": "2026-09-17T17:23:00", "hora": "17:23", "cliente": "confirmar endereço", "tool": "delivery-api", "nome_intencao": "confirmacao_endereco_entrega", "chatbot": {"texto": "O endereço cadastrado está válido. Deseja tentar uma nova entrega?", "botoes": ["sim", "falar com atendente"], "links": [], "acoes": [{"tipo": "botao", "texto": "sim"}, {"tipo": "botao", "texto": "falar com atendente"}]}}, {"seq_sessao": 5, "datahora": "2026-09-17T17:25:00", "hora": "17:25", "cliente": "falar com atendente", "tool": "handoff-api", "nome_intencao": "transbordo_entrega", "chatbot": {"texto": "Vou encaminhar você para um especialista verificar a entrega.", "botoes": [], "links": [], "acoes": []}}]}];
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
    tipoCsat: document.getElementById('tipoCsatFilter').value,
    csat: document.getElementById('csatFilter').value,
    sort: document.getElementById('sortOrder').value
  };
}

function renderSessions(){
  const {q,demand,status,tipoCsat,csat,sort} = currentFilters();
  const list = document.getElementById('sessionList');
  let items = conversations.filter(c => {
    const matchesText = !q || c.id_sessao.toLowerCase().includes(q) || (c.demanda||'').toLowerCase().includes(q);
    const matchesDemand = !demand || (c.demanda||'') === demand;
    const matchesStatus = !status || (c.termino_conversa||'') === status;
    const matchesTipoCsat = !tipoCsat || (c.tipo_csat||'') === tipoCsat;
    const matchesCsat = !csat || (csat === 'none' ? c.nota_csat == null : String(c.nota_csat) === csat);
    return matchesText && matchesDemand && matchesStatus && matchesTipoCsat && matchesCsat;
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
      <div class="session-demand" title="${esc(c.demanda || 'Sem demanda')}">${esc(c.demanda || 'Sem demanda')}</div>
      <div class="session-meta"><span class="session-date">${esc(sessionDateLabel(c))}</span><span class="session-meta-right"><span class="mini-type">${esc(c.tipo_csat || '—')}</span><span class="mini-status ${statusClass(c.termino_conversa)}">${esc(c.termino_conversa || '—')}</span></span></div>
    </button>`).join('') || '<div class="empty">Nenhuma sessão encontrada.</div>';
  list.querySelectorAll('.session').forEach(btn => btn.addEventListener('click', () => selectSession(btn.dataset.id)));
}

function renderBot(interaction){
  const bot = interaction.chatbot || {};
  const fallbackActions = [
    ...(bot.botoes || []).map(texto => ({tipo:'botao', texto})),
    ...(bot.links || []).map(texto => ({tipo:'link', texto}))
  ];
  const actions = (bot.acoes && bot.acoes.length ? bot.acoes : fallbackActions);
  const choices = actions.map(action => {
    const isLink = action.tipo === 'link';
    const icon = isLink ? '<span class="link-arrow">↗</span>' : '<span class="chev">›</span>';
    return `<div class="choice ${isLink ? 'link-choice' : ''}"><span>${esc(action.texto)}</span>${icon}</div>`;
  }).join('');
  return `
    <div class="bot-row">
      <div class="bot-icon">✦</div>
      <div class="bot-content">
        ${bot.texto ? `<div class="bot-text">${esc(bot.texto)}</div>` : ''}
        ${choices ? `<div class="choices">${choices}</div>` : ''}
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
      <div class="turn-tooltip"><div class="tooltip-title">Interação #${esc(i.seq_sessao ?? '—')}</div><div class="tooltip-grid"><span class="tooltip-label">Tool</span><span class="tooltip-value">${esc(i.tool || '—')}</span><span class="tooltip-label">Intenção</span><span class="tooltip-value">${esc(i.nome_intencao || '—')}</span></div></div>
      ${i.cliente ? `<div class="client-row"><div class="client-wrap"><div class="bubble-client">${esc(i.cliente)}</div><div class="time">${esc(i.hora||'')} <span class="checks">✓✓</span></div></div><div class="avatar">VC</div></div>` : ''}
      ${renderBot(i)}
    </div>`).join('');
  chat.scrollTop = 0;
  const lastTime = c.interacoes?.at(-1)?.hora || c.interacoes?.[0]?.hora || '21:30';
  document.getElementById('phoneClock').textContent = lastTime;
  document.getElementById('headSession').textContent = c.id_sessao || 'Sessão';
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
    <div class="meta-card"><div class="label">Tipo CSAT</div><div class="value">${esc(c.tipo_csat || '—')}</div></div>
    <div class="meta-card"><div class="label">CSAT</div><div class="value">${esc(score)}</div><div class="stars">${esc(stars)}</div></div>
    <div class="meta-card"><div class="label">Comentário CSAT</div><div class="comment">${esc(c.comentario_csat || 'Sem comentário.')}</div></div>
    <div class="footer-note">Passe o mouse sobre uma interação para visualizar <strong>seq_sessao</strong>, <strong>tool</strong> e <strong>nome_intencao</strong>. Os campos demanda, término, CSAT, comentário e tipo de CSAT pertencem à sessão e se repetem nas linhas do mesmo <strong>id_sessao</strong>.</div>`;
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

['search','demandFilter','statusFilter','tipoCsatFilter','csatFilter','sortOrder'].forEach(id => {
  const el = document.getElementById(id);
  el.addEventListener(id === 'search' ? 'input' : 'change', renderSessions);
});

document.getElementById('clearFilters').addEventListener('click', () => {
  document.getElementById('search').value = '';
  document.getElementById('demandFilter').value = '';
  document.getElementById('statusFilter').value = '';
  document.getElementById('tipoCsatFilter').value = '';
  document.getElementById('csatFilter').value = '';
  document.getElementById('sortOrder').value = 'recent';
  renderSessions();
});

populateFilters();
renderSessions();
if(selectedId) selectSession(selectedId);
</script>
</body>
</html>
```

## `requirements.txt`

Dependências Python externas.

```text
pandas>=2.2,<3
openpyxl>=3.1,<4

```

## `gerar_viewer.bat`

Atalho para Windows.

```bat
@echo off
setlocal

if not exist .venv (
    py -m venv .venv
)

call .venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt
python src\build_viewer.py

start "" dist\index.html
endlocal

```

## `gerar_viewer.sh`

Atalho para Linux/macOS.

```bash
#!/usr/bin/env bash
set -e

python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
pip install -r requirements.txt
python src/build_viewer.py

echo "Pronto: abra dist/index.html no navegador."

```

## `.gitignore`

Arquivos/pastas ignorados pelo Git.

```gitignore
__pycache__/
*.py[cod]
.venv/
venv/
.DS_Store
.idea/
.vscode/

```

---

## Estrutura do repositório

```text
chatbot_conversation_viewer/
├── data/
│   └── conversas_ficticias.xlsx
├── dist/
│   ├── conversas.json
│   └── index.html
├── src/
│   └── build_viewer.py
├── .gitignore
├── gerar_viewer.bat
├── gerar_viewer.sh
├── requirements.txt
├── README.md
└── README_COMPLETO.md
```

## Observação sobre `conversas.json`

`dist/conversas.json` é um **arquivo gerado**, não um código-fonte. Ele é recriado automaticamente a partir do Excel toda vez que `src/build_viewer.py` é executado. Por isso, o conteúdo integral dele não foi duplicado aqui.

## Fluxo do projeto

```text
Excel (.xlsx)
    ↓
Python / pandas + openpyxl
    ↓
Validação e agrupamento por id_sessao
    ↓
Parser de texto / botões / links
    ↓
conversas.json
    ↓
HTML + CSS + JavaScript puros
    ↓
Conversation Viewer
```
