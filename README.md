<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Painel de Carregamento de Docas</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/sql.js/1.10.3/sql-wasm.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.4/chart.umd.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.4.1/papaparse.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.2/jspdf.plugin.autotable.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
:root{
  --bg:#0d1117;
  --bg-panel:#141a22;
  --bg-panel2:#1b2430;
  --border:#26313f;
  --text:#e6edf3;
  --text-dim:#8b96a5;
  --accent:#ff8a3d;
  --accent-dim:#c9662a;
  --ok:#3ddc84;
  --err:#ff5566;
  --mono:'JetBrains Mono', 'Consolas', monospace;
  --sans:'Inter', -apple-system, sans-serif;
}
*{box-sizing:border-box;margin:0;padding:0;}
body{
  background:radial-gradient(1200px 600px at 10% -10%, #16202b 0%, var(--bg) 55%);
  color:var(--text);
  font-family:var(--sans);
  min-height:100vh;
}
@font-face{font-family:'Inter';src:local('Inter');}
.app{display:flex;min-height:100vh;}

/* ---- MENU LATERAL ---- */
nav.side{
  width:220px;
  background:var(--bg-panel);
  border-right:1px solid var(--border);
  padding:20px 0;
  position:sticky;top:0;height:100vh;
  display:flex;flex-direction:column;
}
.brand{
  padding:0 20px 20px 20px;
  border-bottom:1px solid var(--border);
  margin-bottom:14px;
}
.brand h1{font-size:15px;letter-spacing:.06em;color:var(--accent);text-transform:uppercase;font-weight:800;}
.brand span{font-size:11px;color:var(--text-dim);font-family:var(--mono);}
.menu-item{
  display:flex;align-items:center;gap:10px;
  padding:11px 20px;cursor:pointer;
  color:var(--text-dim);font-size:13.5px;font-weight:600;
  border-left:3px solid transparent;
  transition:.15s;
}
.menu-item:hover{background:var(--bg-panel2);color:var(--text);}
.menu-item.active{color:var(--accent);border-left-color:var(--accent);background:var(--bg-panel2);}
.menu-item .ic{width:18px;text-align:center;}
.side-footer{margin-top:auto;padding:16px 20px;border-top:1px solid var(--border);font-size:10.5px;color:var(--text-dim);font-family:var(--mono);}

/* ---- MAIN ---- */
main{flex:1;padding:26px 32px 60px 32px;max-width:1500px;}
.view{display:none;}
.view.active{display:block;animation:fade .2s ease;}
@keyframes fade{from{opacity:0;transform:translateY(4px);}to{opacity:1;transform:translateY(0);}}

.view-header{display:flex;justify-content:space-between;align-items:flex-end;margin-bottom:22px;flex-wrap:wrap;gap:12px;}
.view-header h2{font-size:20px;font-weight:800;}
.view-header p{color:var(--text-dim);font-size:12.5px;margin-top:4px;}

.btn{
  background:var(--accent);color:#151008;border:none;border-radius:6px;
  padding:9px 16px;font-weight:700;font-size:12.5px;cursor:pointer;
  transition:.15s; letter-spacing:.01em;
}
.btn:hover{background:#ff9c5c;}
.btn.secondary{background:transparent;border:1px solid var(--border);color:var(--text);}
.btn.secondary:hover{border-color:var(--accent);color:var(--accent);}
.btn.danger{background:transparent;border:1px solid var(--err);color:var(--err);}
.btn.danger:hover{background:rgba(255,85,102,.1);}
.btn:disabled{opacity:.4;cursor:not-allowed;}
.btn-row{display:flex;gap:10px;flex-wrap:wrap;}

/* upload */
.dropzone{
  border:2px dashed var(--border);border-radius:12px;
  padding:44px 20px;text-align:center;background:var(--bg-panel);
  cursor:pointer;transition:.15s;
}
.dropzone:hover, .dropzone.drag{border-color:var(--accent);background:var(--bg-panel2);}
.dropzone .big{font-size:15px;font-weight:700;margin-bottom:6px;}
.dropzone .small{font-size:12px;color:var(--text-dim);}
input[type=file]{display:none;}

.card{background:var(--bg-panel);border:1px solid var(--border);border-radius:10px;padding:18px 20px;}
.grid{display:grid;gap:16px;}
.grid.kpis{grid-template-columns:repeat(auto-fit,minmax(160px,1fr));margin:20px 0;}
.kpi{background:var(--bg-panel);border:1px solid var(--border);border-radius:10px;padding:16px 18px;position:relative;overflow:hidden;}
.kpi::before{content:'';position:absolute;left:0;top:0;bottom:0;width:3px;background:var(--accent);}
.kpi .label{font-size:11px;color:var(--text-dim);text-transform:uppercase;letter-spacing:.05em;font-weight:700;}
.kpi .value{font-size:26px;font-weight:800;margin-top:6px;font-family:var(--mono);}
.kpi.ok::before{background:var(--ok);}
.kpi.err::before{background:var(--err);}

.charts-grid{display:grid;grid-template-columns:1fr 1fr;gap:18px;margin-top:18px;}
@media(max-width:1000px){.charts-grid{grid-template-columns:1fr;}}
.chart-box{background:var(--bg-panel);border:1px solid var(--border);border-radius:10px;padding:16px;}
.chart-box h3{font-size:13px;color:var(--text-dim);text-transform:uppercase;letter-spacing:.04em;margin-bottom:12px;}
canvas{max-height:280px;}

/* funcionarios */
.emp-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:12px;margin-top:16px;}
.emp-card{
  background:var(--bg-panel);border:1px solid var(--border);border-radius:10px;
  padding:16px 12px;text-align:center;cursor:pointer;transition:.15s;
}
.emp-card:hover{border-color:var(--accent);transform:translateY(-2px);}
.emp-avatar{
  width:46px;height:46px;border-radius:50%;background:var(--bg-panel2);
  display:flex;align-items:center;justify-content:center;margin:0 auto 10px auto;
  font-weight:800;color:var(--accent);border:1px solid var(--border);font-size:15px;
}
.emp-card .name{font-size:12.5px;font-weight:700;margin-bottom:4px;}
.emp-card .sub{font-size:10.5px;color:var(--text-dim);font-family:var(--mono);}
.emp-card .sub.err{color:var(--err);}

/* modal detalhe funcionario */
.modal-overlay{
  position:fixed;inset:0;background:rgba(0,0,0,.6);backdrop-filter:blur(3px);
  display:none;align-items:flex-start;justify-content:center;z-index:100;padding:30px 16px;overflow-y:auto;
}
.modal-overlay.active{display:flex;}
.modal{
  background:var(--bg-panel);border:1px solid var(--border);border-radius:14px;
  width:100%;max-width:920px;padding:26px;
}
.modal-head{display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;}
.modal-head h2{font-size:19px;font-weight:800;}
.close-x{cursor:pointer;color:var(--text-dim);font-size:22px;line-height:1;}
.close-x:hover{color:var(--err);}

table{width:100%;border-collapse:collapse;font-size:12.5px;margin-top:10px;}
th{text-align:left;color:var(--text-dim);text-transform:uppercase;font-size:10.5px;letter-spacing:.04em;padding:8px 10px;border-bottom:1px solid var(--border);}
td{padding:8px 10px;border-bottom:1px solid var(--border);}
tr:hover td{background:var(--bg-panel2);}
.badge{padding:2px 8px;border-radius:20px;font-size:10.5px;font-weight:700;}
.badge.ok{background:rgba(61,220,132,.15);color:var(--ok);}
.badge.err{background:rgba(255,85,102,.15);color:var(--err);}
.table-wrap{max-height:340px;overflow-y:auto;border:1px solid var(--border);border-radius:8px;margin-top:14px;}

.empty-state{text-align:center;padding:60px 20px;color:var(--text-dim);}
.empty-state .ic{font-size:38px;margin-bottom:12px;}

.hist-row{display:flex;justify-content:space-between;align-items:center;padding:12px 16px;border:1px solid var(--border);border-radius:8px;margin-bottom:10px;background:var(--bg-panel);}
.hist-row .meta{font-size:12px;color:var(--text-dim);font-family:var(--mono);}
.hist-row .name{font-weight:700;font-size:13.5px;}

.toast{
  position:fixed;bottom:22px;right:22px;background:var(--bg-panel2);border:1px solid var(--accent);
  padding:12px 18px;border-radius:8px;font-size:13px;z-index:200;display:none;box-shadow:0 6px 24px rgba(0,0,0,.4);
}
.toast.show{display:block;animation:fade .2s ease;}
.filter-bar{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:14px;}
.filter-bar select, .filter-bar input{
  background:var(--bg-panel2);border:1px solid var(--border);color:var(--text);
  padding:8px 10px;border-radius:6px;font-size:12.5px;font-family:var(--sans);
}

/* ---- FUNCIONARIOS: MASTER-DETAIL ---- */
#view-funcionarios{padding:0;max-width:none;margin-top:-26px;margin-left:-32px;margin-right:-32px;margin-bottom:-60px;}
.split-view{display:flex;height:calc(100vh - 0px);min-height:600px;}
.split-list{
  width:280px;background:var(--bg-panel);border-right:1px solid var(--border);
  display:flex;flex-direction:column;flex-shrink:0;
}
.list-header{padding:18px;border-bottom:1px solid var(--border);}
.list-header input{
  width:100%;background:var(--bg-panel2);border:1px solid var(--border);color:var(--text);
  padding:8px 12px;border-radius:6px;font-size:13px;margin-top:10px;
}
.emp-items{flex:1;overflow-y:auto;padding:10px;}
.emp-item{
  padding:12px 14px;border-radius:8px;cursor:pointer;margin-bottom:4px;
  display:flex;justify-content:space-between;align-items:center;border:1px solid transparent;
}
.emp-item:hover{background:var(--bg-panel2);}
.emp-item.active{background:var(--bg-panel2);border-color:var(--accent);}
.emp-item .name{font-size:12.5px;font-weight:700;margin-bottom:4px;}
.emp-item .meta{font-size:11px;color:var(--text-dim);font-family:var(--mono);}
.emp-item .err-badge{
  background:rgba(255,85,102,.15);color:var(--err);font-size:10px;font-weight:800;
  padding:2px 7px;border-radius:10px;white-space:nowrap;
}
.split-detail{flex:1;padding:26px 32px 60px 32px;overflow-y:auto;background:var(--bg);}
.detail-header{display:flex;justify-content:space-between;align-items:flex-end;border-bottom:1px solid var(--border);padding-bottom:16px;margin-bottom:20px;flex-wrap:wrap;gap:12px;}
.detail-header h2{font-size:22px;font-weight:800;color:var(--accent);}
@media(max-width:900px){
  .split-view{flex-direction:column;height:auto;}
  .split-list{width:100%;max-height:280px;}
}
</style>
</head>
<body>
<div class="app">

  <nav class="side">
    <div class="brand"><h1>Docas • Análise</h1><span>carregamento &amp; rotas</span></div>
    <div class="menu-item active" data-view="upload"><span class="ic">⭱</span> Importar CSV</div>
    <div class="menu-item" data-view="resumo"><span class="ic">▤</span> Resumo Geral</div>
    <div class="menu-item" data-view="funcionarios"><span class="ic">☰</span> Funcionários</div>
    <div class="menu-item" data-view="analises"><span class="ic">📈</span> Análises</div>
    <div class="menu-item" data-view="historico"><span class="ic">◷</span> Histórico</div>
    <div class="side-footer">SQLite (sql.js) local no navegador. Nenhum dado sai do seu computador.</div>
  </nav>

  <main>

    <!-- ===================== UPLOAD ===================== -->
    <section class="view active" id="view-upload">
      <div class="view-header">
        <div>
          <h2>Importar planilha</h2>
          <p>Envie o CSV exportado do sistema de separação/expedição.</p>
        </div>
      </div>

      <div class="card" style="margin-bottom:18px;">
        <b style="font-size:12.5px;">Colunas esperadas (nomes flexíveis):</b>
        <p style="color:var(--text-dim);font-size:12px;margin-top:8px;line-height:1.6;">
          <b>Funcionário</b> (nome, colaborador, funcionario) · <b>Rota</b> (rota, route) ·
          <b>ID Pacote</b> (id, id_pacote, pacote, codigo) · <b>Horário</b> (horario, hora, timestamp) ·
          <b>Status</b> (status, correto, ok, situacao) — valores tipo "correto/incorreto", "sim/não", "1/0", "ok/erro" são aceitos.
        </p>
      </div>

      <label class="dropzone" id="dropzone">
        <input type="file" id="fileInput" accept=".csv,.xlsx,.xls">
        <div class="big">📄 Clique ou arraste o arquivo CSV ou Excel (.xlsx) aqui</div>
        <div class="small">Um novo lote será somado ao histórico do painel</div>
      </label>

      <div id="uploadStatus" style="margin-top:16px;"></div>
    </section>

    <!-- ===================== RESUMO GERAL ===================== -->
    <section class="view" id="view-resumo">
      <div class="view-header">
        <div>
          <h2>Resumo Geral</h2>
          <p id="resumoSub">Visão consolidada de todos os funcionários e rotas</p>
        </div>
        <div class="btn-row">
          <button class="btn secondary" onclick="exportCSVGeral()">⭳ CSV</button>
          <button class="btn secondary" onclick="exportPDFGeral()">⭳ PDF</button>
          <button class="btn danger" onclick="confirmReset()">✕ Resetar tudo</button>
        </div>
      </div>

      <div id="resumoEmpty" class="empty-state">
        <div class="ic">📊</div>
        <div>Nenhum dado importado ainda. Vá em <b>Importar CSV</b>.</div>
      </div>

      <div id="resumoContent" style="display:none;">
        <div class="grid kpis">
          <div class="kpi"><div class="label">Rotas</div><div class="value" id="kpiRotas">0</div></div>
          <div class="kpi"><div class="label">Pacotes</div><div class="value" id="kpiPacotes">0</div></div>
          <div class="kpi"><div class="label">Funcionários</div><div class="value" id="kpiFuncionarios">0</div></div>
          <div class="kpi"><div class="label">Tempo médio</div><div class="value" id="kpiTempo">--</div></div>
          <div class="kpi ok"><div class="label">Pacotes corretos</div><div class="value" id="kpiCorretos">0</div></div>
          <div class="kpi err"><div class="label">Pacotes incorretos</div><div class="value" id="kpiIncorretos">0</div></div>
        </div>

        <div class="charts-grid">
          <div class="chart-box"><h3>Pacotes por funcionário</h3><canvas id="chartPorFunc"></canvas></div>
          <div class="chart-box"><h3>Corretos vs Incorretos</h3><canvas id="chartCorretoIncorreto"></canvas></div>
          <div class="chart-box"><h3>Pacotes por rota</h3><canvas id="chartPorRota"></canvas></div>
          <div class="chart-box"><h3>Volume por horário</h3><canvas id="chartPorHora"></canvas></div>
        </div>

        <h3 style="margin-top:26px;font-size:14px;">Pacotes incorretos (detalhe)</h3>
        <div class="table-wrap">
          <table id="tblErros">
            <thead><tr><th>Funcionário</th><th>Rota</th><th>ID Pacote</th><th>Horário</th><th>Status</th></tr></thead>
            <tbody></tbody>
          </table>
        </div>
      </div>
    </section>

    <!-- ===================== FUNCIONARIOS (MASTER-DETAIL) ===================== -->
    <section class="view" id="view-funcionarios">
      <div class="split-view">
        <div class="split-list">
          <div class="list-header">
            <h3 style="font-size:14px;">Equipe</h3>
            <input type="text" id="searchEmp" placeholder="Buscar funcionário...">
          </div>
          <div class="emp-items" id="empList">
            <div style="padding:20px;color:var(--text-dim);font-size:12px;text-align:center;">Nenhum dado. Importe o CSV.</div>
          </div>
        </div>

        <div class="split-detail" id="empDetail" style="display:none;">
          <div class="detail-header">
            <div>
              <h2 id="detName">Nome do Funcionário</h2>
              <p id="detSub" style="color:var(--text-dim);font-size:13px;margin-top:4px;"></p>
            </div>
            <div class="btn-row">
              <select id="detFilterStatus" style="background:var(--bg-panel2);border:1px solid var(--border);color:var(--text);padding:8px 10px;border-radius:6px;font-size:12.5px;">
                <option value="">Todos os status</option>
                <option value="1">Somente corretos</option>
                <option value="0">Somente incorretos</option>
              </select>
              <button class="btn secondary" onclick="exportCSVFuncionario()">⭳ CSV</button>
              <button class="btn secondary" onclick="exportPDFFuncionario()">⭳ PDF</button>
            </div>
          </div>

          <div class="grid kpis" id="detKpis"></div>

          <div class="charts-grid" style="margin-top:20px;grid-template-columns:2fr 1fr;">
            <div class="chart-box"><h3>Produtividade (pacotes por hora)</h3><canvas id="chartDetHora"></canvas></div>
            <div class="chart-box"><h3>Taxa de acerto</h3><canvas id="chartDetStatus"></canvas></div>
          </div>
          <div class="charts-grid" style="margin-top:20px;grid-template-columns:1fr;">
            <div class="chart-box"><h3>Volume de pacotes por rota (rotas únicas trabalhadas)</h3><canvas id="chartDetRota"></canvas></div>
          </div>

          <h3 style="margin-top:26px;font-size:14px;border-bottom:1px solid var(--border);padding-bottom:10px;">Registros do funcionário</h3>
          <div class="table-wrap">
            <table id="tblDet">
              <thead><tr><th>Rota</th><th>ID Pacote</th><th>Horário</th><th>Status</th></tr></thead>
              <tbody></tbody>
            </table>
          </div>
        </div>

        <div class="split-detail" id="empDetailEmpty" style="display:flex;align-items:center;justify-content:center;color:var(--text-dim);">
          Selecione um funcionário na lateral para ver os indicadores.
        </div>
      </div>
    </section>

    <!-- ===================== ANALISES ===================== -->
    <section class="view" id="view-analises">
      <div class="view-header">
        <div>
          <h2>Análises</h2>
          <p>Ranking, comparativos entre importações e rotas problemáticas</p>
        </div>
      </div>

      <div id="analisesEmpty" class="empty-state">
        <div class="ic">📈</div>
        <div>Nenhum dado importado ainda. Vá em <b>Importar CSV</b>.</div>
      </div>

      <div id="analisesContent" style="display:none;">

        <h3 style="font-size:14px;margin-bottom:4px;">🏆 Ranking de funcionários</h3>
        <p style="color:var(--text-dim);font-size:12px;margin-bottom:12px;">Ordene por volume, taxa de acerto ou tempo médio</p>
        <div class="filter-bar">
          <select id="rankSortBy">
            <option value="total">Ordenar por: Volume de pacotes</option>
            <option value="taxa">Ordenar por: Taxa de acerto</option>
            <option value="tempo">Ordenar por: Tempo médio (mais rápido primeiro)</option>
            <option value="rotas">Ordenar por: Rotas únicas</option>
          </select>
        </div>
        <div class="table-wrap" style="margin-bottom:10px;">
          <table id="tblRanking">
            <thead><tr><th>#</th><th>Funcionário</th><th>Pacotes</th><th>Rotas únicas</th><th>Tempo médio</th><th>Taxa de acerto</th><th>Incorretos</th></tr></thead>
            <tbody></tbody>
          </table>
        </div>

        <div class="charts-grid">
          <div class="chart-box"><h3>Top 10 — Volume de pacotes</h3><canvas id="chartRankVolume"></canvas></div>
          <div class="chart-box"><h3>Top 10 — Taxa de acerto</h3><canvas id="chartRankTaxa"></canvas></div>
        </div>

        <h3 style="margin-top:32px;font-size:14px;margin-bottom:4px;">📊 Comparativo entre importações (lotes/dias)</h3>
        <p style="color:var(--text-dim);font-size:12px;margin-bottom:12px;">Cada importação de CSV vira um ponto de comparação — útil pra ver evolução dia a dia</p>
        <div class="charts-grid">
          <div class="chart-box"><h3>Pacotes por lote importado</h3><canvas id="chartLoteVolume"></canvas></div>
          <div class="chart-box"><h3>Taxa de acerto por lote importado</h3><canvas id="chartLoteTaxa"></canvas></div>
        </div>

        <h3 style="margin-top:32px;font-size:14px;margin-bottom:4px;">⚠️ Rotas com mais problemas</h3>
        <p style="color:var(--text-dim);font-size:12px;margin-bottom:12px;">Rotas ordenadas pela maior taxa de erro (mínimo 2 pacotes na rota)</p>
        <div class="table-wrap">
          <table id="tblRotasProblema">
            <thead><tr><th>Rota</th><th>Total de pacotes</th><th>Incorretos</th><th>Taxa de erro</th></tr></thead>
            <tbody></tbody>
          </table>
        </div>

      </div>
    </section>

    <!-- ===================== HISTORICO ===================== -->
    <section class="view" id="view-historico">
      <div class="view-header">
        <div>
          <h2>Histórico de importações</h2>
          <p>Lotes de CSV importados nesta sessão do navegador</p>
        </div>
      </div>
      <div id="histEmpty" class="empty-state">
        <div class="ic">◷</div>
        <div>Nenhuma importação registrada ainda.</div>
      </div>
      <div id="histList"></div>
    </section>

  </main>
</div>

<div class="toast" id="toast"></div>

<script>
/* =========================================================
   ESTADO / BANCO SQLite (sql.js)
========================================================= */
let SQL = null;
let db = null;
let charts = {}; // guarda instâncias Chart.js pra destruir ao redesenhar
let currentEmployee = null;

const LS_KEY = 'docas_sqlite_db_v1';
const LS_HIST_KEY = 'docas_historico_v1';

function toast(msg){
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.classList.add('show');
  setTimeout(()=>t.classList.remove('show'), 2600);
}

async function initSQL(){
  SQL = await initSqlJs({ locateFile: f => `https://cdnjs.cloudflare.com/ajax/libs/sql.js/1.10.3/${f}` });
  const saved = localStorage.getItem(LS_KEY);
  if(saved){
    try{
      const bin = Uint8Array.from(atob(saved), c=>c.charCodeAt(0));
      db = new SQL.Database(bin);
    }catch(e){ db = new SQL.Database(); createSchema(); }
  } else {
    db = new SQL.Database();
    createSchema();
  }
  ensureSchema();
  refreshAllViews();
}

function createSchema(){
  db.run(`
    CREATE TABLE IF NOT EXISTS pacotes(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      lote_id INTEGER,
      funcionario TEXT,
      rota TEXT,
      id_pacote TEXT,
      horario TEXT,
      correto INTEGER
    );
    CREATE TABLE IF NOT EXISTS lotes(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      arquivo TEXT,
      data_importacao TEXT,
      total_linhas INTEGER
    );
  `);
}
function ensureSchema(){
  try{ db.run("SELECT 1 FROM pacotes LIMIT 1"); }catch(e){ createSchema(); }
}

function persist(){
  const data = db.export();
  let binary = '';
  for(let i=0;i<data.length;i++) binary += String.fromCharCode(data[i]);
  localStorage.setItem(LS_KEY, btoa(binary));
}

function q(sql, params=[]){
  const stmt = db.prepare(sql);
  if(params.length) stmt.bind(params);
  const rows = [];
  while(stmt.step()) rows.push(stmt.getAsObject());
  stmt.free();
  return rows;
}

/* =========================================================
   MENU NAVIGATION
========================================================= */
document.querySelectorAll('.menu-item').forEach(item=>{
  item.addEventListener('click', ()=> goToView(item.dataset.view));
});

/* =========================================================
   UPLOAD / PARSE CSV
========================================================= */
const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');

dropzone.addEventListener('dragover', e=>{e.preventDefault();dropzone.classList.add('drag');});
dropzone.addEventListener('dragleave', ()=>dropzone.classList.remove('drag'));
dropzone.addEventListener('drop', e=>{
  e.preventDefault(); dropzone.classList.remove('drag');
  if(e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener('change', e=>{
  if(e.target.files.length) handleFile(e.target.files[0]);
});

const COLUMN_ALIASES = {
  funcionario: ['funcionario','funcionário','colaborador','nome','operador','empregado'],
  rota: ['rota','route','trajeto'],
  id_pacote: ['id','id_pacote','idpacote','pacote','codigo','código','tracking','nf','id do pacote'],
  horario: ['horario','horário','hora','timestamp','data_hora','datahora'],
  status: ['status','correto','ok','situacao','situação','conferencia','conferência']
};

function guessColumn(headers, aliasList){
  const norm = h => h.toLowerCase().trim();
  for(const h of headers){
    if(aliasList.includes(norm(h))) return h;
  }
  // partial match fallback
  for(const h of headers){
    const nh = norm(h);
    if(aliasList.some(a => nh.includes(a))) return h;
  }
  return null;
}

function handleFile(file){
  document.getElementById('uploadStatus').innerHTML = `<p style="color:var(--text-dim);font-size:12.5px;">Lendo arquivo "${file.name}"...</p>`;
  const ext = file.name.toLowerCase().split('.').pop();

  if(ext === 'xlsx' || ext === 'xls'){
    readExcelFile(file);
  } else if(ext === 'csv') {
    readCSVFile(file);
  } else {
    document.getElementById('uploadStatus').innerHTML = `<p style="color:var(--err);font-size:12.5px;">Formato não suportado (.${ext}). Envie um arquivo .csv, .xlsx ou .xls.</p>`;
  }
}

function readCSVFile(file){
  Papa.parse(file, {
    header:true,
    skipEmptyLines:true,
    encoding:'UTF-8',
    complete: res => {
      if(!res.data.length){
        document.getElementById('uploadStatus').innerHTML = `<p style="color:var(--err);font-size:12.5px;">Arquivo vazio ou ilegível.</p>`;
        return;
      }
      processParsedRows(res.data, res.meta.fields, file.name);
    },
    error: err => {
      document.getElementById('uploadStatus').innerHTML = `<p style="color:var(--err);font-size:12.5px;">Erro ao ler CSV: ${err.message}</p>`;
    }
  });
}

function readExcelFile(file){
  const reader = new FileReader();
  reader.onload = e => {
    try{
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, { type: 'array', cellDates:true });
      const sheetName = wb.SheetNames[0];
      const sheet = wb.Sheets[sheetName];
      // defval:'' garante que células vazias virem string vazia em vez de sumir da linha
      const json = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });
      if(!json.length){
        document.getElementById('uploadStatus').innerHTML = `<p style="color:var(--err);font-size:12.5px;">A planilha "${sheetName}" está vazia.</p>`;
        return;
      }
      const headers = Object.keys(json[0]);
      processParsedRows(json, headers, file.name);
    }catch(err){
      document.getElementById('uploadStatus').innerHTML = `<p style="color:var(--err);font-size:12.5px;">Erro ao ler Excel: ${err.message}</p>`;
    }
  };
  reader.onerror = () => {
    document.getElementById('uploadStatus').innerHTML = `<p style="color:var(--err);font-size:12.5px;">Não foi possível ler o arquivo.</p>`;
  };
  reader.readAsArrayBuffer(file);
}

function processParsedRows(rows, headers, filename){
  const mapping = {
    funcionario: guessColumn(headers, COLUMN_ALIASES.funcionario),
    rota: guessColumn(headers, COLUMN_ALIASES.rota),
    id_pacote: guessColumn(headers, COLUMN_ALIASES.id_pacote),
    horario: guessColumn(headers, COLUMN_ALIASES.horario),
    status: guessColumn(headers, COLUMN_ALIASES.status)
  };
  if(!mapping.funcionario || !mapping.id_pacote){
    document.getElementById('uploadStatus').innerHTML = `<p style="color:var(--err);font-size:12.5px;">Não consegui identificar as colunas de Funcionário e/ou ID do Pacote. Colunas encontradas: ${headers.join(', ')}.</p>`;
    return;
  }
  importData(rows, mapping, filename);
}

function parseStatus(val){
  if(val === undefined || val === null) return 1;
  const v = String(val).trim().toLowerCase();
  if(['correto','ok','sim','1','true','certo','verdadeiro'].includes(v)) return 1;
  if(['incorreto','erro','nao','não','0','false','errado','falso'].includes(v)) return 0;
  return 1; // default: considera correto se não reconhecido
}

function importData(rows, mapping, filename){
  ensureSchema();
  const now = new Date().toISOString();
  db.run("INSERT INTO lotes (arquivo, data_importacao, total_linhas) VALUES (?,?,?)", [filename, now, rows.length]);
  const loteId = q("SELECT last_insert_rowid() as id")[0].id;

  const stmt = db.prepare(`INSERT INTO pacotes (lote_id, funcionario, rota, id_pacote, horario, correto) VALUES (?,?,?,?,?,?)`);
  db.run("BEGIN TRANSACTION");
  let count = 0;
  rows.forEach(r=>{
    const funcionario = (r[mapping.funcionario] || '').toString().trim();
    if(!funcionario) return;
    const rota = mapping.rota ? (r[mapping.rota]||'').toString().trim() : '';
    const idPacote = mapping.id_pacote ? (r[mapping.id_pacote]||'').toString().trim() : '';
    const horario = mapping.horario ? (r[mapping.horario]||'').toString().trim() : '';
    const status = mapping.status ? parseStatus(r[mapping.status]) : 1;
    stmt.run([loteId, funcionario, rota, idPacote, horario, status]);
    count++;
  });
  db.run("COMMIT");
  stmt.free();
  persist();

  document.getElementById('uploadStatus').innerHTML = `<p style="color:var(--ok);font-size:13px;">✔ ${count} pacotes importados com sucesso (lote #${loteId}).</p>`;
  fileInput.value = '';
  toast(`Importação concluída: ${count} linhas. Abrindo resumo...`);
  refreshAllViews();
  goToView('resumo');
}

function goToView(viewName){
  document.querySelectorAll('.menu-item').forEach(m=>m.classList.toggle('active', m.dataset.view === viewName));
  document.querySelectorAll('.view').forEach(v=>v.classList.toggle('active', v.id === 'view-'+viewName));
  if(viewName === 'resumo') renderResumo();
  if(viewName === 'funcionarios') renderFuncionarios();
  if(viewName === 'analises') renderAnalises();
  if(viewName === 'historico') renderHistorico();
}

/* =========================================================
   TEMPO MÉDIO — parsing flexível de horário
========================================================= */
function parseTimeToMinutes(str){
  if(!str) return null;
  str = str.trim();
  // tenta HH:MM ou HH:MM:SS
  let m = str.match(/(\d{1,2}):(\d{2})(:(\d{2}))?/);
  if(m){
    return parseInt(m[1])*60 + parseInt(m[2]) + (m[4]?parseInt(m[4])/60:0);
  }
  // tenta data completa
  const d = new Date(str);
  if(!isNaN(d.getTime())) return d.getHours()*60 + d.getMinutes();
  return null;
}
function minutesToHHMM(mins){
  if(mins === null || isNaN(mins)) return '--';
  const h = Math.floor(mins/60)%24;
  const m = Math.round(mins%60);
  return String(h).padStart(2,'0')+'h'+String(m).padStart(2,'0');
}

/* =========================================================
   RESUMO GERAL
========================================================= */
function destroyChart(id){
  if(charts[id]){ charts[id].destroy(); delete charts[id]; }
}

function refreshAllViews(){
  renderResumo();
  renderFuncionarios();
  renderAnalises();
  renderHistorico();
}

function renderResumo(){
  const total = q("SELECT COUNT(*) as c FROM pacotes")[0].c;
  if(total === 0){
    document.getElementById('resumoEmpty').style.display='block';
    document.getElementById('resumoContent').style.display='none';
    return;
  }
  document.getElementById('resumoEmpty').style.display='none';
  document.getElementById('resumoContent').style.display='block';

  const rotas = q("SELECT COUNT(DISTINCT rota) as c FROM pacotes WHERE rota != ''")[0].c;
  const pacotes = q("SELECT COUNT(DISTINCT id_pacote) as c FROM pacotes WHERE id_pacote != ''")[0].c;
  const funcionarios = q("SELECT COUNT(DISTINCT funcionario) as c FROM pacotes")[0].c;
  const corretos = q("SELECT COUNT(*) as c FROM pacotes WHERE correto=1")[0].c;
  const incorretos = q("SELECT COUNT(*) as c FROM pacotes WHERE correto=0")[0].c;

  const horarios = q("SELECT horario FROM pacotes WHERE horario != ''").map(r=>parseTimeToMinutes(r.horario)).filter(v=>v!==null);
  const tempoMedio = horarios.length ? horarios.reduce((a,b)=>a+b,0)/horarios.length : null;

  document.getElementById('kpiRotas').textContent = rotas;
  document.getElementById('kpiPacotes').textContent = pacotes || total;
  document.getElementById('kpiFuncionarios').textContent = funcionarios;
  document.getElementById('kpiTempo').textContent = minutesToHHMM(tempoMedio);
  document.getElementById('kpiCorretos').textContent = corretos;
  document.getElementById('kpiIncorretos').textContent = incorretos;

  // gráfico por funcionário
  const porFunc = q(`SELECT funcionario, COUNT(*) as total FROM pacotes GROUP BY funcionario ORDER BY total DESC LIMIT 15`);
  destroyChart('porFunc');
  charts.porFunc = new Chart(document.getElementById('chartPorFunc'), {
    type:'bar',
    data:{ labels: porFunc.map(r=>r.funcionario), datasets:[{ label:'Pacotes', data: porFunc.map(r=>r.total), backgroundColor:'#ff8a3d' }] },
    options: baseChartOpts()
  });

  // corretos vs incorretos
  destroyChart('corretoIncorreto');
  charts.corretoIncorreto = new Chart(document.getElementById('chartCorretoIncorreto'), {
    type:'doughnut',
    data:{ labels:['Corretos','Incorretos'], datasets:[{ data:[corretos, incorretos], backgroundColor:['#3ddc84','#ff5566'] }] },
    options:{ plugins:{ legend:{ labels:{ color:'#e6edf3' } } } }
  });

  // por rota
  const porRota = q(`SELECT rota, COUNT(*) as total FROM pacotes WHERE rota != '' GROUP BY rota ORDER BY total DESC LIMIT 15`);
  destroyChart('porRota');
  charts.porRota = new Chart(document.getElementById('chartPorRota'), {
    type:'bar',
    data:{ labels: porRota.map(r=>r.rota), datasets:[{ label:'Pacotes', data: porRota.map(r=>r.total), backgroundColor:'#4d9fff' }] },
    options: baseChartOpts()
  });

  // por hora
  const horasList = q("SELECT horario FROM pacotes WHERE horario != ''").map(r=>parseTimeToMinutes(r.horario)).filter(v=>v!==null).map(m=>Math.floor(m/60));
  const horaCount = {};
  horasList.forEach(h=>{ horaCount[h]=(horaCount[h]||0)+1; });
  const horaLabels = Object.keys(horaCount).sort((a,b)=>a-b);
  destroyChart('porHora');
  charts.porHora = new Chart(document.getElementById('chartPorHora'), {
    type:'line',
    data:{ labels: horaLabels.map(h=>h+'h'), datasets:[{ label:'Pacotes', data: horaLabels.map(h=>horaCount[h]), borderColor:'#ff8a3d', backgroundColor:'rgba(255,138,61,.15)', fill:true, tension:.3 }] },
    options: baseChartOpts()
  });

  // tabela de erros
  const erros = q(`SELECT funcionario, rota, id_pacote, horario FROM pacotes WHERE correto=0 ORDER BY funcionario LIMIT 500`);
  const tb = document.querySelector('#tblErros tbody');
  tb.innerHTML = erros.map(r=>`<tr><td>${r.funcionario}</td><td>${r.rota||'-'}</td><td>${r.id_pacote||'-'}</td><td>${r.horario||'-'}</td><td><span class="badge err">Incorreto</span></td></tr>`).join('') || `<tr><td colspan="5" style="color:var(--text-dim);">Nenhum pacote incorreto 🎉</td></tr>`;
}

function baseChartOpts(){
  return {
    responsive:true,
    plugins:{ legend:{ display:false } },
    scales:{
      x:{ ticks:{ color:'#8b96a5', font:{size:10} }, grid:{ color:'#26313f' } },
      y:{ ticks:{ color:'#8b96a5' }, grid:{ color:'#26313f' } }
    }
  };
}

/* =========================================================
   FUNCIONARIOS (MASTER-DETAIL)
========================================================= */
let listData = [];

function renderFuncionarios(){
  listData = q(`SELECT funcionario, COUNT(*) as total, COUNT(DISTINCT rota) as rotas, SUM(CASE WHEN correto=0 THEN 1 ELSE 0 END) as erros FROM pacotes GROUP BY funcionario ORDER BY total DESC`);
  const listEl = document.getElementById('empList');
  if(!listData.length){
    listEl.innerHTML = '<div style="padding:20px;color:var(--text-dim);font-size:12px;text-align:center;">Nenhum dado. Importe o CSV.</div>';
    document.getElementById('empDetail').style.display = 'none';
    document.getElementById('empDetailEmpty').style.display = 'flex';
    currentEmployee = null;
    return;
  }
  updateEmpListHTML(listData);
  // se o funcionário selecionado sumiu (ex: após reset+import), ou nada selecionado, seleciona o primeiro
  if(!currentEmployee || !listData.some(r=>r.funcionario===currentEmployee)){
    openEmployee(listData[0].funcionario);
  }
}

function updateEmpListHTML(data){
  document.getElementById('empList').innerHTML = data.map(r=>{
    return `<div class="emp-item ${currentEmployee===r.funcionario?'active':''}" onclick="openEmployee('${r.funcionario.replace(/'/g,"\\'")}')">
      <div>
        <div class="name">${r.funcionario}</div>
        <div class="meta">${r.total} pct · ${r.rotas} rota(s)</div>
      </div>
      ${r.erros>0 ? `<div class="err-badge">${r.erros} err</div>` : '<span style="color:var(--ok);font-size:13px;">✔</span>'}
    </div>`;
  }).join('') || '<div style="padding:20px;color:var(--text-dim);font-size:12px;text-align:center;">Nenhum funcionário encontrado.</div>';
}

document.getElementById('searchEmp').addEventListener('input', e=>{
  const term = e.target.value.toLowerCase();
  updateEmpListHTML(listData.filter(r=>r.funcionario.toLowerCase().includes(term)));
});

function openEmployee(name){
  currentEmployee = name;
  updateEmpListHTML(listData.some(r=>r.funcionario===name) ? listData : q(`SELECT funcionario, COUNT(*) as total, COUNT(DISTINCT rota) as rotas, SUM(CASE WHEN correto=0 THEN 1 ELSE 0 END) as erros FROM pacotes GROUP BY funcionario ORDER BY total DESC`));
  document.getElementById('empDetailEmpty').style.display = 'none';
  document.getElementById('empDetail').style.display = 'block';

  const rows = q(`SELECT rota, id_pacote, horario, correto FROM pacotes WHERE funcionario=?`, [name]);
  const total = rows.length;
  const corretos = rows.filter(r=>r.correto===1).length;
  const incorretos = total - corretos;
  const rotasUnicas = new Set(rows.map(r=>r.rota).filter(Boolean)).size;
  const tempos = rows.map(r=>parseTimeToMinutes(r.horario)).filter(v=>v!==null);
  const tempoMedio = tempos.length ? tempos.reduce((a,b)=>a+b,0)/tempos.length : null;
  const taxa = total > 0 ? ((corretos/total)*100).toFixed(1) : 0;

  document.getElementById('detName').textContent = name;
  document.getElementById('detSub').textContent = `Esse funcionário processou pacotes em ${rotasUnicas} rota(s) única(s) · tempo médio ${minutesToHHMM(tempoMedio)}`;

  document.getElementById('detKpis').innerHTML = `
    <div class="kpi"><div class="label">Total de pacotes</div><div class="value">${total}</div></div>
    <div class="kpi" style="border-left:3px solid #4d9fff;"><div class="label">Rotas únicas</div><div class="value" style="color:#4d9fff;">${rotasUnicas}</div></div>
    <div class="kpi"><div class="label">Tempo médio</div><div class="value">${minutesToHHMM(tempoMedio)}</div></div>
    <div class="kpi ok"><div class="label">Taxa de acerto</div><div class="value">${taxa}%</div></div>
    <div class="kpi err"><div class="label">Incorretos</div><div class="value">${incorretos}</div></div>
  `;

  // Gráfico: volume por rota (rotas únicas)
  const porRota = {};
  rows.forEach(r=>{ const k = r.rota||'(sem rota)'; porRota[k]=(porRota[k]||0)+1; });
  destroyChart('detRota');
  charts.detRota = new Chart(document.getElementById('chartDetRota'), {
    type:'bar',
    data:{ labels:Object.keys(porRota), datasets:[{ label:'Pacotes', data:Object.values(porRota), backgroundColor:'#4d9fff', borderRadius:4 }] },
    options: baseChartOpts()
  });

  // Gráfico: taxa de acerto
  destroyChart('detStatus');
  charts.detStatus = new Chart(document.getElementById('chartDetStatus'), {
    type:'doughnut',
    data:{ labels:['Corretos','Incorretos'], datasets:[{ data:[corretos, incorretos], backgroundColor:['#3ddc84','#ff5566'], borderWidth:0 }] },
    options:{ plugins:{ legend:{ position:'bottom', labels:{ color:'#e6edf3' } } } }
  });

  // Gráfico: produtividade por hora (usa parseTimeToMinutes, robusto a HH:MM, HH:MM:SS e data completa)
  const horaCount = {};
  rows.forEach(r=>{
    const mins = parseTimeToMinutes(r.horario);
    if(mins === null) return;
    const h = Math.floor(mins/60);
    horaCount[h] = (horaCount[h]||0)+1;
  });
  const horaLabels = Object.keys(horaCount).sort((a,b)=>a-b);
  destroyChart('detHora');
  charts.detHora = new Chart(document.getElementById('chartDetHora'), {
    type:'line',
    data:{ labels: horaLabels.map(h=>h+'h'), datasets:[{ label:'Pacotes', data: horaLabels.map(h=>horaCount[h]), borderColor:'#ff8a3d', backgroundColor:'rgba(255,138,61,.2)', fill:true, tension:.4 }] },
    options: baseChartOpts()
  });

  document.getElementById('detFilterStatus').value = '';
  renderDetTable();
}

document.getElementById('detFilterStatus').addEventListener('change', renderDetTable);

function renderDetTable(){
  if(!currentEmployee) return;
  const filter = document.getElementById('detFilterStatus').value;
  let sql = `SELECT rota, id_pacote, horario, correto FROM pacotes WHERE funcionario=?`;
  const params = [currentEmployee];
  if(filter !== ''){ sql += ` AND correto=?`; params.push(parseInt(filter)); }
  sql += ` ORDER BY horario`;
  const rows = q(sql, params);
  const tb = document.querySelector('#tblDet tbody');
  tb.innerHTML = rows.map(r=>`<tr>
    <td>${r.rota||'-'}</td><td>${r.id_pacote||'-'}</td><td>${r.horario||'-'}</td>
    <td>${r.correto===1?'<span class="badge ok">Correto</span>':'<span class="badge err">Incorreto</span>'}</td>
  </tr>`).join('') || `<tr><td colspan="4" style="color:var(--text-dim);">Nenhum registro.</td></tr>`;
}

/* =========================================================
   ANÁLISES — Ranking, comparativo entre lotes, rotas problemáticas
========================================================= */
function renderAnalises(){
  const total = q("SELECT COUNT(*) as c FROM pacotes")[0].c;
  if(total === 0){
    document.getElementById('analisesEmpty').style.display = 'block';
    document.getElementById('analisesContent').style.display = 'none';
    return;
  }
  document.getElementById('analisesEmpty').style.display = 'none';
  document.getElementById('analisesContent').style.display = 'block';

  renderRanking();
  renderComparativoLotes();
  renderRotasProblema();
}

function computeRankingData(){
  const funcionarios = q(`SELECT DISTINCT funcionario FROM pacotes`).map(r=>r.funcionario);
  return funcionarios.map(nome=>{
    const rows = q(`SELECT rota, horario, correto FROM pacotes WHERE funcionario=?`, [nome]);
    const total = rows.length;
    const corretos = rows.filter(r=>r.correto===1).length;
    const incorretos = total - corretos;
    const rotasUnicas = new Set(rows.map(r=>r.rota).filter(Boolean)).size;
    const tempos = rows.map(r=>parseTimeToMinutes(r.horario)).filter(v=>v!==null);
    const tempoMedio = tempos.length ? tempos.reduce((a,b)=>a+b,0)/tempos.length : null;
    const taxa = total > 0 ? (corretos/total)*100 : 0;
    return { nome, total, incorretos, rotasUnicas, tempoMedio, taxa };
  });
}

function renderRanking(){
  const sortBy = document.getElementById('rankSortBy').value;
  let data = computeRankingData();
  if(sortBy === 'total') data.sort((a,b)=>b.total-a.total);
  else if(sortBy === 'taxa') data.sort((a,b)=>b.taxa-a.taxa);
  else if(sortBy === 'tempo') data.sort((a,b)=>(a.tempoMedio??Infinity)-(b.tempoMedio??Infinity));
  else if(sortBy === 'rotas') data.sort((a,b)=>b.rotasUnicas-a.rotasUnicas);

  const medals = ['🥇','🥈','🥉'];
  const tb = document.querySelector('#tblRanking tbody');
  tb.innerHTML = data.map((r,i)=>`<tr>
    <td>${medals[i]||(i+1)}</td>
    <td><b>${r.nome}</b></td>
    <td>${r.total}</td>
    <td>${r.rotasUnicas}</td>
    <td>${minutesToHHMM(r.tempoMedio)}</td>
    <td>${r.taxa.toFixed(1)}%</td>
    <td>${r.incorretos>0?`<span class="badge err">${r.incorretos}</span>`:'<span class="badge ok">0</span>'}</td>
  </tr>`).join('') || `<tr><td colspan="7" style="color:var(--text-dim);">Sem dados.</td></tr>`;

  const top10Volume = [...data].sort((a,b)=>b.total-a.total).slice(0,10);
  destroyChart('rankVolume');
  charts.rankVolume = new Chart(document.getElementById('chartRankVolume'), {
    type:'bar',
    data:{ labels: top10Volume.map(r=>r.nome), datasets:[{ label:'Pacotes', data: top10Volume.map(r=>r.total), backgroundColor:'#ff8a3d', borderRadius:4 }] },
    options: { ...baseChartOpts(), indexAxis:'y' }
  });

  const top10Taxa = [...data].filter(r=>r.total>0).sort((a,b)=>b.taxa-a.taxa).slice(0,10);
  destroyChart('rankTaxa');
  charts.rankTaxa = new Chart(document.getElementById('chartRankTaxa'), {
    type:'bar',
    data:{ labels: top10Taxa.map(r=>r.nome), datasets:[{ label:'Taxa de acerto (%)', data: top10Taxa.map(r=>r.taxa.toFixed(1)), backgroundColor:'#3ddc84', borderRadius:4 }] },
    options: { ...baseChartOpts(), indexAxis:'y', scales:{ x:{ min:0, max:100, ticks:{ color:'#8b96a5' }, grid:{ color:'#26313f' } }, y:{ ticks:{ color:'#8b96a5', font:{size:10} }, grid:{ color:'#26313f' } } } }
  });
}
document.getElementById('rankSortBy').addEventListener('change', renderRanking);

function renderComparativoLotes(){
  const lotes = q(`SELECT id, arquivo, data_importacao FROM lotes ORDER BY id ASC`);
  const labels = [];
  const volumeData = [];
  const taxaData = [];
  lotes.forEach(l=>{
    const rows = q(`SELECT correto FROM pacotes WHERE lote_id=?`, [l.id]);
    if(rows.length === 0) return;
    const corretos = rows.filter(r=>r.correto===1).length;
    const d = new Date(l.data_importacao);
    labels.push(`#${l.id} ${d.toLocaleDateString('pt-BR')}`);
    volumeData.push(rows.length);
    taxaData.push(((corretos/rows.length)*100).toFixed(1));
  });

  destroyChart('loteVolume');
  charts.loteVolume = new Chart(document.getElementById('chartLoteVolume'), {
    type:'bar',
    data:{ labels, datasets:[{ label:'Pacotes', data: volumeData, backgroundColor:'#4d9fff', borderRadius:4 }] },
    options: baseChartOpts()
  });

  destroyChart('loteTaxa');
  charts.loteTaxa = new Chart(document.getElementById('chartLoteTaxa'), {
    type:'line',
    data:{ labels, datasets:[{ label:'Taxa de acerto (%)', data: taxaData, borderColor:'#3ddc84', backgroundColor:'rgba(61,220,132,.15)', fill:true, tension:.3 }] },
    options:{ ...baseChartOpts(), scales:{ x:{ ticks:{ color:'#8b96a5', font:{size:10} }, grid:{ color:'#26313f' } }, y:{ min:0, max:100, ticks:{ color:'#8b96a5' }, grid:{ color:'#26313f' } } } }
  });
}

function renderRotasProblema(){
  const rows = q(`SELECT rota, COUNT(*) as total, SUM(CASE WHEN correto=0 THEN 1 ELSE 0 END) as erros FROM pacotes WHERE rota != '' GROUP BY rota HAVING total >= 2 ORDER BY (CAST(erros as FLOAT)/total) DESC, erros DESC LIMIT 20`);
  const tb = document.querySelector('#tblRotasProblema tbody');
  tb.innerHTML = rows.map(r=>{
    const taxaErro = ((r.erros/r.total)*100).toFixed(1);
    return `<tr>
      <td><b>${r.rota}</b></td>
      <td>${r.total}</td>
      <td>${r.erros>0?`<span class="badge err">${r.erros}</span>`:'<span class="badge ok">0</span>'}</td>
      <td>${taxaErro}%</td>
    </tr>`;
  }).join('') || `<tr><td colspan="4" style="color:var(--text-dim);">Nenhuma rota com pacotes suficientes para análise.</td></tr>`;
}

/* =========================================================
   HISTORICO
========================================================= */
function renderHistorico(){
  const lotes = q(`SELECT * FROM lotes ORDER BY id DESC`);
  const list = document.getElementById('histList');
  if(lotes.length === 0){
    document.getElementById('histEmpty').style.display='block';
    list.innerHTML='';
    return;
  }
  document.getElementById('histEmpty').style.display='none';
  list.innerHTML = lotes.map(l=>{
    const d = new Date(l.data_importacao);
    return `<div class="hist-row">
      <div>
        <div class="name">${l.arquivo}</div>
        <div class="meta">Lote #${l.id} · ${d.toLocaleString('pt-BR')}</div>
      </div>
      <div class="meta">${l.total_linhas} linhas</div>
    </div>`;
  }).join('');
}

/* =========================================================
   EXPORTAÇÃO CSV / PDF
========================================================= */
function downloadBlob(content, filename, mime){
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function rowsToCSV(rows, headers){
  const lines = [headers.join(';')];
  rows.forEach(r=> lines.push(headers.map(h=>`"${(r[h]??'').toString().replace(/"/g,'""')}"`).join(';')));
  return lines.join('\n');
}

function exportCSVGeral(){
  const rows = q(`SELECT funcionario, rota, id_pacote, horario, CASE WHEN correto=1 THEN 'Correto' ELSE 'Incorreto' END as status FROM pacotes`);
  const csv = rowsToCSV(rows, ['funcionario','rota','id_pacote','horario','status']);
  downloadBlob(csv, `relatorio_geral_${dateStamp()}.csv`, 'text/csv;charset=utf-8;');
  toast('CSV geral exportado.');
}

function exportCSVFuncionario(){
  if(!currentEmployee) return;
  const rows = q(`SELECT rota, id_pacote, horario, CASE WHEN correto=1 THEN 'Correto' ELSE 'Incorreto' END as status FROM pacotes WHERE funcionario=?`, [currentEmployee]);
  const csv = rowsToCSV(rows, ['rota','id_pacote','horario','status']);
  downloadBlob(csv, `relatorio_${currentEmployee.replace(/\s+/g,'_')}_${dateStamp()}.csv`, 'text/csv;charset=utf-8;');
  toast('CSV do funcionário exportado.');
}

function dateStamp(){
  const d = new Date();
  return d.toISOString().slice(0,10);
}

function exportPDFGeral(){
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF();
  doc.setFontSize(16); doc.text('Relatório Geral - Carregamento de Docas', 14, 16);
  doc.setFontSize(10); doc.setTextColor(120);
  doc.text(`Gerado em ${new Date().toLocaleString('pt-BR')}`, 14, 22);

  const rotas = q("SELECT COUNT(DISTINCT rota) as c FROM pacotes WHERE rota != ''")[0].c;
  const pacotes = q("SELECT COUNT(DISTINCT id_pacote) as c FROM pacotes WHERE id_pacote != ''")[0].c;
  const funcionarios = q("SELECT COUNT(DISTINCT funcionario) as c FROM pacotes")[0].c;
  const corretos = q("SELECT COUNT(*) as c FROM pacotes WHERE correto=1")[0].c;
  const incorretos = q("SELECT COUNT(*) as c FROM pacotes WHERE correto=0")[0].c;
  const horarios = q("SELECT horario FROM pacotes WHERE horario != ''").map(r=>parseTimeToMinutes(r.horario)).filter(v=>v!==null);
  const tempoMedio = horarios.length ? horarios.reduce((a,b)=>a+b,0)/horarios.length : null;

  doc.setTextColor(0);
  doc.setFontSize(11);
  doc.text(`Rotas: ${rotas}   |   Pacotes: ${pacotes}   |   Funcionários: ${funcionarios}`, 14, 32);
  doc.text(`Tempo médio: ${minutesToHHMM(tempoMedio)}   |   Corretos: ${corretos}   |   Incorretos: ${incorretos}`, 14, 39);

  const porFunc = q(`SELECT funcionario, COUNT(*) as total, SUM(CASE WHEN correto=0 THEN 1 ELSE 0 END) as erros FROM pacotes GROUP BY funcionario ORDER BY total DESC`);
  doc.autoTable({
    startY: 46,
    head: [['Funcionário','Total de pacotes','Incorretos']],
    body: porFunc.map(r=>[r.funcionario, r.total, r.erros]),
    styles:{ fontSize:9 }, headStyles:{ fillColor:[255,138,61] }
  });

  const erros = q(`SELECT funcionario, rota, id_pacote, horario FROM pacotes WHERE correto=0`);
  if(erros.length){
    doc.addPage();
    doc.setFontSize(13); doc.text('Pacotes incorretos', 14, 16);
    doc.autoTable({
      startY: 22,
      head: [['Funcionário','Rota','ID Pacote','Horário']],
      body: erros.map(r=>[r.funcionario, r.rota||'-', r.id_pacote||'-', r.horario||'-']),
      styles:{ fontSize:9 }, headStyles:{ fillColor:[255,85,102] }
    });
  }
  doc.save(`relatorio_geral_${dateStamp()}.pdf`);
  toast('PDF geral exportado.');
}

function exportPDFFuncionario(){
  if(!currentEmployee) return;
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF();
  const rows = q(`SELECT rota, id_pacote, horario, correto FROM pacotes WHERE funcionario=?`, [currentEmployee]);
  const total = rows.length;
  const corretos = rows.filter(r=>r.correto===1).length;
  const incorretos = total - corretos;
  const tempos = rows.map(r=>parseTimeToMinutes(r.horario)).filter(v=>v!==null);
  const tempoMedio = tempos.length ? tempos.reduce((a,b)=>a+b,0)/tempos.length : null;

  doc.setFontSize(16); doc.text(`Relatório - ${currentEmployee}`, 14, 16);
  doc.setFontSize(10); doc.setTextColor(120);
  doc.text(`Gerado em ${new Date().toLocaleString('pt-BR')}`, 14, 22);
  doc.setTextColor(0); doc.setFontSize(11);
  doc.text(`Total de pacotes: ${total}   |   Tempo médio: ${minutesToHHMM(tempoMedio)}`, 14, 32);
  doc.text(`Corretos: ${corretos}   |   Incorretos: ${incorretos}`, 14, 39);

  doc.autoTable({
    startY: 46,
    head: [['Rota','ID Pacote','Horário','Status']],
    body: rows.map(r=>[r.rota||'-', r.id_pacote||'-', r.horario||'-', r.correto===1?'Correto':'Incorreto']),
    styles:{ fontSize:9 }, headStyles:{ fillColor:[255,138,61] }
  });
  doc.save(`relatorio_${currentEmployee.replace(/\s+/g,'_')}_${dateStamp()}.pdf`);
  toast('PDF do funcionário exportado.');
}

/* =========================================================
   RESET
========================================================= */
function confirmReset(){
  if(confirm('Isso vai apagar TODOS os dados importados (pacotes e histórico). Deseja continuar?')){
    db.run("DELETE FROM pacotes; DELETE FROM lotes;");
    currentEmployee = null;
    persist();
    refreshAllViews();
    toast('Todos os dados foram resetados.');
  }
}

/* =========================================================
   INIT
========================================================= */
initSQL();
</script>
</body>
</html>

