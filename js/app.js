/* ============================================================
   B.I.G.O.D.E — lógica da aplicação
   Depende de: js/data.js (dados fixos) e, para exportar Excel, da
   biblioteca SheetJS (carregada no index.html).
   ============================================================ */
/* ============================================================
   CAPACIDADES DA PLATAFORMA
   ============================================================ */
let dbNs = null;
let userNs = null;
let canManage = false;
let myId = null;
let saveTimers = {};
let localOnlyFichas = {};
let localOnlyInv = {};
let localOnlyColab = {};


function regiaoDeColab(c){
  if(!c) return '';
  if(c.regiao !== undefined && c.regiao !== null) return c.regiao;   // escolha manual no cadastro
  const cid = (c.cidade || CIDADE_POR_MATRICULA[c.matricula] || '').toUpperCase();
  return REGIAO_POR_CIDADE[cid] || '';
}
function regiaoDaFicha(f){
  const c = colaboradores.find(x => (x.nome||'').toLowerCase() === (f.colaborador||'').toLowerCase());
  return regiaoDeColab(c);
}
let colabRegiao = 'todas';
let fichaRegiao = 'todas';
let fichaCC = 'todas';
let colabCC = 'todas';
let invPosse = 'todas';
let invLocacao = 'todas';

function centroCustoDeColab(c){
  if(!c) return '';
  if(c.centroCusto !== undefined && c.centroCusto !== null) return c.centroCusto;   // escolha manual no cadastro
  return CC_POR_MATRICULA[c.matricula] || '';
}
function centroCustoDaFicha(f){
  const c = colaboradores.find(x => (x.nome||'').toLowerCase() === (f.colaborador||'').toLowerCase());
  return centroCustoDeColab(c);
}
function listaCentrosCusto(){
  const s = new Set(Object.values(CC_POR_MATRICULA).filter(Boolean));
  colaboradores.forEach(c => { const v = centroCustoDeColab(c); if(v) s.add(v); });
  return [...s].sort((a,b)=> a.localeCompare(b));
}

/* Locação: dias até o vencimento e situação (vencida / vencendo em até 15 dias / ok) */
function diasParaVencer(i){
  if(!i.vencimentoLocacao) return null;
  const hoje = new Date(); hoje.setHours(0,0,0,0);
  const [y,m,d] = i.vencimentoLocacao.split('-').map(Number);
  return Math.round((new Date(y, m-1, d) - hoje) / 86400000);
}
function locacaoInfo(i){
  if(i.posse !== 'locado' || !i.vencimentoLocacao || i.baixado) return null;
  const d = diasParaVencer(i);
  if(d < 0)  return {status:'vencida',  cls:'manutencao', label:`Locação vencida há ${-d} dia(s)`};
  if(d === 0) return {status:'vencendo', cls:'alocado',    label:'Locação vence hoje'};
  if(d <= 15) return {status:'vencendo', cls:'alocado',    label:`Locação vence em ${d} dia(s)`};
  return {status:'ok', cls:'', label:`Locação até ${fmtDataBR(i.vencimentoLocacao)}`};
}
function locacaoBadge(i){
  const info = locacaoInfo(i);
  if(!info) return '';
  return info.cls ? `<span class="pill ${info.cls}">${info.label}</span>` : `<span>${info.label}</span>`;
}

let fichas = [];
let inventario = [];
let colaboradores = [];
let currentFichaId = null;
let currentSection = 'fichas';
let filtroAtual = "todas";
let searchTerm = "";
let invFiltro = "todos";
let invSearchTerm = "";
let colabSearchTerm = "";
let expandedInvId = null;
let expandedColabId = null;

function uid(prefix){ return prefix + '_' + Date.now() + '_' + Math.random().toString(36).slice(2,7); }
function getFichaById(id){ return fichas.find(f => f.id === id); }
function getInvById(id){ return inventario.find(i => i.id === id); }
function fmtDate(iso){
  if(!iso) return "—";
  const d = new Date(iso);
  return d.toLocaleDateString('pt-BR') + ' às ' + d.toLocaleTimeString('pt-BR', {hour:'2-digit', minute:'2-digit'});
}
function fmtDateShort(iso){ if(!iso) return "—"; return new Date(iso).toLocaleDateString('pt-BR'); }
function escapeHtml(s){ return (s||'').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }

function novaFichaParaFuncao(funcao){
  const lista = FUNCOES[funcao] || [];
  return {
    id: uid('f'), colaborador: "", funcao: funcao, tipo: "padrao",
    itens: lista.map(nome => ({
      nome, patrimonio: "", inventarioId: null,
      condEntrega: null, dataRetirada: "", obsEntrega: "",
      condDevolucao: null, dataDevolucao: "", obsDevolucao: ""
    })),
    assinaturaRetirada: null, dataAssinaturaRetirada: null,
    assinaturaDevolucao: null, dataAssinaturaDevolucao: null,
    status: "rascunho", criadoEm: new Date().toISOString()
  };
}
function novaFichaEsporadica(){
  return {
    id: uid('f'), colaborador: "", funcao: "Entrega esporádica", tipo: "esporadica",
    itens: [],
    assinaturaRetirada: null, dataAssinaturaRetirada: null,
    assinaturaDevolucao: null, dataAssinaturaDevolucao: null,
    status: "rascunho", criadoEm: new Date().toISOString()
  };
}
function novoItemInventario(){
  return {
    id: uid('inv'), nome:"", patrimonio:"", quantidade:1, posse:"proprio", fornecedor:"",
    quantidadeAlocada:0, quantidadeManutencao:0, baixado:false,
    periodos:[], manutencoes:[],
    historico:[{tipo:'criacao', data:new Date().toISOString(), detalhes:'Item cadastrado', por: myId || null}],
    criadoEm: new Date().toISOString()
  };
}
/* Compatibilidade: lê itens antigos (campo "situacao") e garante as listas novas. */
function invNormalize(i){
  if(i.quantidadeAlocada === undefined){
    i.quantidade = i.quantidade || 1;
    i.quantidadeAlocada = i.situacao === 'alocado' ? 1 : 0;
    i.quantidadeManutencao = i.situacao === 'manutencao' ? 1 : 0;
    i.baixado = i.situacao === 'baixado';
  }
  i.periodos = i.periodos || [];
  i.manutencoes = i.manutencoes || [];
  const abertas = i.manutencoes.filter(m => !m.dataRetorno).length;
  for(let k = abertas; k < (i.quantidadeManutencao||0); k++){
    i.manutencoes.push({id:'man_legado_'+k, dataEnvio:(i.criadoEm||'').slice(0,10) || hojeISO(), dataRetorno:null, valor:null, oficina:'', motivo:'Registro anterior ao controle detalhado'});
  }
  return i;
}
function invDisponivel(i){ return Math.max(0, (i.quantidade||1) - (i.quantidadeAlocada||0) - (i.quantidadeManutencao||0)); }
function invPillInfo(i){
  const un = (i.quantidade||1) === 1;
  if(i.baixado) return {cls:'baixado', label:'Baixado'};
  if(invDisponivel(i) > 0) return {cls:'disponivel', label: un ? 'Disponível' : `Disponível ${invDisponivel(i)}/${i.quantidade}`};
  if((i.quantidadeManutencao||0) > 0) return {cls:'manutencao', label:'Em manutenção'};
  return {cls:'alocado', label: un ? 'Alocado' : 'Todo alocado'};
}

/* ---------- datas, duração e dinheiro ---------- */
const pad2 = n => String(n).padStart(2,'0');
function hojeISO(){ const d = new Date(); return `${d.getFullYear()}-${pad2(d.getMonth()+1)}-${pad2(d.getDate())}`; }
function parseISO(s){ const [y,m,d] = s.split('-').map(Number); return new Date(y, m-1, d); }
function diffDias(a, b){ return Math.round((parseISO(b) - parseISO(a)) / 86400000); }
function fmtDuracao(dias){
  if(!dias || dias <= 0) return '0 dia';
  const m = Math.floor(dias/30), r = dias % 30, p = [];
  if(m) p.push(`${m} ${m > 1 ? 'meses' : 'mês'}`);
  if(r) p.push(`${r} ${r > 1 ? 'dias' : 'dia'}`);
  return p.join(' e ');
}
function fmtBRL(v){ return (Number(v)||0).toLocaleString('pt-BR', {style:'currency', currency:'BRL'}); }
function round2(v){ return Math.round((Number(v)||0) * 100) / 100; }

/* ---------- histórico de locação (períodos com valor) ----------
   Cada período: {id, inicio, fim (vazio = em curso), valorMensal, obs}.
   Mudança de valor = fecha o período e abre outro na mesma data (sem intervalo).
   Intervalo entre o fim de um período e o início do próximo = "sem locação".
   Custo do período = valor mensal × dias ÷ 30. */
function periodosDe(i){ return [...(i.periodos||[])].sort((a,b)=> a.inicio.localeCompare(b.inicio)); }
function periodoDias(p){ return Math.max(0, diffDias(p.inicio, p.fim || hojeISO())); }
function periodoValor(p){ return (Number(p.valorMensal)||0) * periodoDias(p) / 30; }
function periodoAberto(i){ return periodosDe(i).find(p => !p.fim) || null; }
function inicioContinuo(i){   // início da locação em curso, somando períodos emendados (mudanças de valor)
  const ps = periodosDe(i), ab = periodoAberto(i);
  if(!ab) return null;
  let ini = ab.inicio;
  for(let k = ps.indexOf(ab) - 1; k >= 0 && ps[k].fim === ini; k--) ini = ps[k].inicio;
  return ini;
}
function locacaoResumo(i){
  const ps = periodosDe(i);
  let dias = 0, total = 0, ocioso = 0;
  ps.forEach((p,k)=>{
    dias += periodoDias(p); total += periodoValor(p);
    if(k > 0 && ps[k-1].fim) ocioso += Math.max(0, diffDias(ps[k-1].fim, p.inicio));
  });
  const ult = ps[ps.length-1];
  if(ult && ult.fim) ocioso += Math.max(0, diffDias(ult.fim, hojeISO()));
  return {dias, total, ocioso, aberto: periodoAberto(i), n: ps.length};
}
function manutResumo(i){
  const l = i.manutencoes || [];
  return {total: l.reduce((s,m)=> s + (Number(m.valor)||0), 0), abertas: l.filter(m => !m.dataRetorno).length, n: l.length};
}
function syncContadorManut(i){ i.quantidadeManutencao = (i.manutencoes||[]).filter(m => !m.dataRetorno).length; }
function abrirManutencaoAuto(inv, motivo){
  inv.manutencoes = (inv.manutencoes||[]).concat([{id:uid('man'), dataEnvio:hojeISO(), dataRetorno:null, valor:null, oficina:'', motivo}]);
  syncContadorManut(inv);
  logInv(inv, 'manutencao', `Enviado para manutenção — ${motivo}`);
}
function locacaoLinha(i){
  const ab = periodoAberto(i);
  if(!ab) return '';
  const ini = inicioContinuo(i);
  const partes = [`Locado desde ${fmtDataBR(ini)} (${fmtDuracao(diffDias(ini, hojeISO()))})`];
  if(Number(ab.valorMensal)) partes.push(`${fmtBRL(ab.valorMensal)}/mês`);
  if(i.vencimentoLocacao) partes.push(`vence ${fmtDataBR(i.vencimentoLocacao)}`);
  return `<span>${partes.join(' · ')}</span>`;
}
/* Lista de ferramentas agrupada por nome, escolhendo o patrimônio de cada uma */
function invOptionsGrouped(lista, selectedId, extraFn){
  const g = {};
  lista.forEach(i => { const n = i.nome || '(sem nome)'; (g[n] = g[n] || []).push(i); });
  return Object.keys(g).sort((a,b)=> a.localeCompare(b)).map(n =>
    `<optgroup label="${escapeHtml(n)}">` +
    g[n].sort((a,b)=> (a.patrimonio||'').localeCompare(b.patrimonio||'', 'pt-BR', {numeric:true}))
        .map(i => `<option value="${i.id}" ${i.id === selectedId ? 'selected' : ''}>Patr. ${escapeHtml(i.patrimonio)}${i.posse === 'locado' ? ' · locado' : ''}${extraFn ? extraFn(i) : ''}</option>`).join('') +
    `</optgroup>`).join('');
}
function novoColaborador(){
  return { id: uid('colab'), nome:"", funcao: FUNCOES_NOMES[0], matricula:"", criadoEm: new Date().toISOString() };
}

/* ============================================================
   PERSISTÊNCIA
   ============================================================ */
function scheduleSave(f){
  clearTimeout(saveTimers[f.id]);
  saveTimers[f.id] = setTimeout(()=> persistFicha(f), 500);
}
async function persistFicha(f){
  if(!dbNs) return;
  try{ const body = Object.assign({}, f); delete body.id; await dbNs.collection('fichas').doc(f.id).set(body); }
  catch(e){ console.error(e); showToast('Não foi possível salvar agora.'); }
}
async function persistFichaNow(f){ clearTimeout(saveTimers[f.id]); await persistFicha(f); }
async function deleteFichaRemote(id){
  if(!dbNs) return;
  try{ await dbNs.collection('fichas').doc(id).delete(); }catch(e){ console.error(e); showToast('Não foi possível excluir.'); }
}
async function persistInv(item){
  if(!dbNs) return;
  item.atualizadoEm = new Date().toISOString(); item.atualizadoPor = myId || null;   // data (e autor) da última alteração
  try{ const body = Object.assign({}, item); delete body.id; await dbNs.collection('inventario').doc(item.id).set(body); }
  catch(e){ console.error(e); showToast('Não foi possível salvar o item.'); }
}
async function persistColab(c){
  if(!dbNs) return;
  try{ const body = Object.assign({}, c); delete body.id; await dbNs.collection('colaboradores').doc(c.id).set(body); }
  catch(e){ console.error(e); showToast('Não foi possível salvar o colaborador.'); }
}
async function deleteInvRemote(id){
  if(!dbNs) return;
  try{ await dbNs.collection('inventario').doc(id).delete(); }
  catch(e){ console.error(e); showToast('Não foi possível excluir o item.'); }
}
async function deleteColabRemote(id){
  if(!dbNs) return;
  try{ await dbNs.collection('colaboradores').doc(id).delete(); }
  catch(e){ console.error(e); showToast('Não foi possível excluir o colaborador.'); }
}

/* O banco entrega cada documento CONGELADO (não pode ser alterado no lugar). Para editar registros
   internos (manutenções, períodos, itens da ficha...), trabalhamos numa cópia profunda.
   Documento que não mudou chega como o mesmo objeto, então a cópia é reaproveitada. */
const _copiasDoc = new WeakMap();
function docClone(d){
  const raw = d.data();
  let copia = _copiasDoc.get(raw);
  if(!copia){ copia = JSON.parse(JSON.stringify(raw)); _copiasDoc.set(raw, copia); }
  return Object.assign({id: d.id}, copia);
}
function subscribeAll(){
  dbNs.collection('fichas').orderBy('criadoEm','desc').onSnapshot((snap)=>{
    const fromServer = snap.docs.map(d => docClone(d));
    const ids = new Set(fromServer.map(f=>f.id));
    Object.keys(localOnlyFichas).forEach(id=>{ if(ids.has(id)) delete localOnlyFichas[id]; });
    fichas = [...Object.values(localOnlyFichas), ...fromServer];
    if(currentSection==='fichas' && viewVisible('view-dashboard')) renderDashboard();
    if(currentFichaId && !fichas.some(f=>f.id===currentFichaId) && viewVisible('view-ficha')){
      showToast('Esta ficha não está mais disponível.'); showDashboard();
    }
  }, (err)=> console.error(err));

  dbNs.collection('inventario').orderBy('criadoEm','desc').onSnapshot((snap)=>{
    const fromServer = snap.docs.map(d => docClone(d));
    const ids = new Set(fromServer.map(i=>i.id));
    Object.keys(localOnlyInv).forEach(id=>{ if(ids.has(id)) delete localOnlyInv[id]; });
    inventario = [...Object.values(localOnlyInv), ...fromServer.map(invNormalize)];
    if(viewVisible('view-inventario')) renderInventario();
    if(viewVisible('view-ficha')) renderEsporadicaAdder(getFichaById(currentFichaId));
    if(viewVisible('view-dashboard')) renderStats();
  }, (err)=> console.error(err));

  dbNs.collection('colaboradores').orderBy('nome').onSnapshot((snap)=>{
    const fromServer = snap.docs.map(d => docClone(d));
    const ids = new Set(fromServer.map(c=>c.id));
    Object.keys(localOnlyColab).forEach(id=>{ if(ids.has(id)) delete localOnlyColab[id]; });
    colaboradores = [...Object.values(localOnlyColab), ...fromServer];
    if(viewVisible('view-colaboradores')) renderColaboradores();
    renderColabDatalist();
    populateFuncaoSelect();
    renderAllFilterMenus();
  }, (err)=> console.error(err));
}
function viewVisible(id){ return !document.getElementById(id).classList.contains('hidden'); }

/* ============================================================
   INICIALIZAÇÃO
   ============================================================ */
let downloadsNs = null;
async function initApp(){
  try{ dbNs = await claude.use('db'); }catch(e){ console.error(e); }
  try{ userNs = await claude.use('user'); }catch(e){ console.error(e); }
  try{ downloadsNs = await claude.use('downloads'); }catch(e){ console.error(e); }
  if(userNs){
    try{ canManage = await userNs.canEdit(); myId = await userNs.id(); }catch(e){ console.error(e); }
  }
  if(!dbNs){ renderUnavailable(); return; }
  if(!canManage){   // perfil Visualização: esconde os botões de cadastro (só consulta e exporta Excel)
    ['btn-toggle-inv-form','btn-toggle-colab-form'].forEach(id => { const el = document.getElementById(id); if(el) el.classList.add('hidden'); });
  }
  if(downloadsNs && window.XLSX){
    document.getElementById('btn-export-excel').classList.remove('hidden');
  }
  subscribeAll();
  populateFuncaoSelect();
  renderAllFilterMenus();
  showDashboard();
}

/* ============================================================
   EXPORTAR EXCEL
   ============================================================ */
function statusFichaLabel(s){ return s==='concluida' ? 'Concluída' : (s==='aberta' ? 'Em aberto' : 'Rascunho'); }
function fmtDataBR(d){
  if(!d) return '';
  if(/^\d{4}-\d{2}-\d{2}$/.test(d)){ const [y,m,dd] = d.split('-'); return `${dd}/${m}/${y}`; }
  return new Date(d).toLocaleDateString('pt-BR');
}
/* ---------- Menu suspenso de filtros (um por aba) ---------- */
const OPC_REGIAO = ()=> [['todas','Todas'],['Interior','Interior'],['Vale','Vale']];
const OPC_CC = ()=> [['todas','Todos'], ...listaCentrosCusto().map(c => [c, c])];
const FILTER_SCHEMAS = {
  fichas: [
    {key:'status', label:'Status da ficha', def:'todas', get:()=>filtroAtual, set:v=>{ filtroAtual = v; }, options:()=>[['todas','Todas'],['aberta','Em aberto'],['concluida','Concluídas']]},
    {key:'regiao', label:'Região', def:'todas', get:()=>fichaRegiao, set:v=>{ fichaRegiao = v; }, options:OPC_REGIAO},
    {key:'cc', label:'Centro de custo', def:'todas', get:()=>fichaCC, set:v=>{ fichaCC = v; }, options:OPC_CC}
  ],
  colab: [
    {key:'regiao', label:'Região', def:'todas', get:()=>colabRegiao, set:v=>{ colabRegiao = v; }, options:OPC_REGIAO},
    {key:'cc', label:'Centro de custo', def:'todas', get:()=>colabCC, set:v=>{ colabCC = v; }, options:OPC_CC}
  ],
  inv: [
    {key:'situacao', label:'Situação', def:'todos', get:()=>invFiltro, set:v=>{ invFiltro = v; }, options:()=>[['todos','Todos'],['disponivel','Disponível'],['alocado','Alocado'],['manutencao','Manutenção'],['baixado','Baixados']]},
    {key:'posse', label:'Posse', def:'todas', get:()=>invPosse, set:v=>{ invPosse = v; }, options:()=>[['todas','Todas'],['proprio','Próprio'],['locado','Locado']]},
    {key:'locacao', label:'Vencimento da locação', def:'todas', get:()=>invLocacao, set:v=>{ invLocacao = v; }, options:()=>[['todas','Todas'],['vencida','Vencida'],['vencendo','Vence em até 15 dias']]}
  ]
};
let filterMenuOpen = null;
function activeFilterCount(kind){ return FILTER_SCHEMAS[kind].filter(g => g.get() !== g.def).length; }
function renderFilterMenu(kind){
  const el = document.getElementById('filter-menu-' + kind);
  if(!el) return;
  const n = activeFilterCount(kind);
  let html = `<button type="button" class="btn btn-ghost filter-btn ${n ? 'has-active' : ''}" data-fm-toggle="${kind}">☰ Filtros${n ? ` (${n})` : ''} ▾</button>`;
  if(filterMenuOpen === kind){
    html += `<div class="filter-panel">` + FILTER_SCHEMAS[kind].map(g => `
      <div class="filter-group">
        <div class="filter-group-title">${g.label}</div>
        <div class="filter-opts">${g.options().map(([v,l]) => `<button type="button" class="filter-opt ${g.get() === v ? 'active' : ''}" data-fm-kind="${kind}" data-fm-key="${g.key}" data-fm-val="${escapeHtml(v)}">${escapeHtml(l)}</button>`).join('')}</div>
      </div>`).join('') + `
      <div class="filter-panel-foot">
        <button type="button" class="btn btn-ghost btn-sm" data-fm-clear="${kind}">Limpar filtros</button>
        <button type="button" class="btn btn-primary btn-sm" data-fm-close="1">Fechar</button>
      </div>
    </div>`;
  }
  el.innerHTML = html;
}
function renderAllFilterMenus(){
  ['fichas','colab','inv'].forEach(renderFilterMenu);
  const dl = document.getElementById('cc-datalist');
  if(dl) dl.innerHTML = listaCentrosCusto().map(c => `<option value="${escapeHtml(c)}">`).join('');
}
document.addEventListener('click', (e)=>{
  const t = e.target;
  if(!t || !t.closest) return;
  const tog = t.closest('[data-fm-toggle]');
  if(tog){ const k = tog.dataset.fmToggle; filterMenuOpen = (filterMenuOpen === k) ? null : k; renderAllFilterMenus(); return; }
  const opt = t.closest('[data-fm-val]');
  if(opt){
    const k = opt.dataset.fmKind;
    const g = FILTER_SCHEMAS[k].find(x => x.key === opt.dataset.fmKey);
    if(g){ g.set(opt.dataset.fmVal); }
    renderFilterMenu(k); rerenderKind(k);
    return;
  }
  const clr = t.closest('[data-fm-clear]');
  if(clr){ const k = clr.dataset.fmClear; FILTER_SCHEMAS[k].forEach(g => g.set(g.def)); renderFilterMenu(k); rerenderKind(k); return; }
  if(t.closest('[data-fm-close]')){ filterMenuOpen = null; renderAllFilterMenus(); return; }
  if(filterMenuOpen && !t.closest('.filter-menu')){ filterMenuOpen = null; renderAllFilterMenus(); }
});

/* ---------- Seleção de linhas para exportar ---------- */
const sel = {fichas:new Set(), inv:new Set(), colab:new Set()};
const lastVisible = {fichas:[], inv:[], colab:[]};
function selBox(kind, id){
  return `<input type="checkbox" class="sel-box" data-kind="${kind}" data-id="${id}" ${sel[kind].has(id) ? 'checked' : ''} title="Selecionar para exportar">`;
}
function totalSelecionado(){ return sel.fichas.size + sel.inv.size + sel.colab.size; }
function updateExportButton(){
  const b = document.getElementById('btn-export-excel');
  if(!b) return;
  const n = totalSelecionado();
  b.textContent = n ? `⬇ Exportar Excel (${n} selecionado${n>1?'s':''})` : '⬇ Exportar Excel';
}
function renderSelBar(kind, ids){
  lastVisible[kind] = ids;
  const el = document.getElementById('sel-bar-' + kind);
  if(el){
    const n = sel[kind].size;
    const todos = ids.length > 0 && ids.every(id => sel[kind].has(id));
    el.innerHTML = ids.length
      ? `<button type="button" class="btn btn-ghost btn-sm" data-selall="${kind}">${todos ? 'Desmarcar todos' : 'Selecionar todos'} (${ids.length})</button>`
        + (n ? `<button type="button" class="btn btn-ghost btn-sm" data-selclear="${kind}">Limpar seleção</button><span class="sel-count">${n} selecionado${n>1?'s':''}</span>` : '')
      : '';
  }
  updateExportButton();
}
function rerenderKind(kind){
  if(kind === 'fichas') renderDashboard(); else if(kind === 'inv') renderInventario(); else renderColaboradores();
}
// clique na caixinha: marca/desmarca sem abrir a ficha nem expandir a linha
document.addEventListener('click', (e)=>{
  const t = e.target;
  if(t && t.classList && t.classList.contains('sel-box')){
    e.stopPropagation();
    const k = t.dataset.kind;
    const ids = t.dataset.ids ? t.dataset.ids.split(',') : [t.dataset.id];
    ids.forEach(id => { if(t.checked) sel[k].add(id); else sel[k].delete(id); });
    if(k === 'inv') renderInventario(); else if(k === 'fichas' && t.dataset.ids) renderDashboard(); else renderSelBar(k, lastVisible[k]);
  }
}, true);
document.addEventListener('click', (e)=>{
  const all = e.target.closest && e.target.closest('[data-selall]');
  if(all){
    const k = all.dataset.selall;
    const ids = lastVisible[k];
    const todos = ids.length > 0 && ids.every(id => sel[k].has(id));
    ids.forEach(id => { if(todos) sel[k].delete(id); else sel[k].add(id); });
    rerenderKind(k);
    return;
  }
  const clr = e.target.closest && e.target.closest('[data-selclear]');
  if(clr){ const k = clr.dataset.selclear; sel[k].clear(); rerenderKind(k); }
});

/* ---------- Montagem das planilhas ---------- */
function linhasFichasDe(lista){
  const out = [];
  [...lista]
    .filter(f => !(f.status==='rascunho' && !f.colaborador && !f.itens.some(i=>i.condEntrega)))
    .sort((a,b)=> new Date(b.criadoEm) - new Date(a.criadoEm))
    .forEach(f=>{
      const base = {
        'Colaborador': f.colaborador || '',
        'Função': f.funcao || '',
        'Tipo': f.tipo === 'esporadica' ? 'Esporádica' : 'Padrão',
        'Região': regiaoDaFicha(f),
        'Centro de custo': centroCustoDaFicha(f),
        'Status da ficha': statusFichaLabel(f.status),
        'Criada em': fmtDataBR(f.criadoEm),
        'Retirada assinada em': f.dataAssinaturaRetirada ? fmtDate(f.dataAssinaturaRetirada) : '',
        'Última renovação do prazo': f.renovadoEm ? fmtDate(f.renovadoEm) : '',
        'Renovações': (f.renovacoes||[]).length || '',
        'Dias com os itens': (f.status === 'aberta' && f.dataAssinaturaRetirada) ? Math.floor((Date.now() - new Date(f.dataAssinaturaRetirada).getTime())/86400000) : '',
        'Prazo de 30 dias': f.status === 'aberta' ? (fichaVencida(f) ? 'Vencido' : 'Em dia') : '',
        'Devolução assinada em': f.dataAssinaturaDevolucao ? fmtDate(f.dataAssinaturaDevolucao) : ''
      };
      if(!f.itens.length){ out.push(Object.assign({}, base)); return; }
      f.itens.forEach(it=>{
        out.push(Object.assign({}, base, {
          'Equipamento': it.nome || '',
          'Patrimônio': it.patrimonio || '',
          'Condição na retirada': it.condEntrega || '',
          'Data de retirada': fmtDataBR(it.dataRetirada),
          'Obs. retirada': it.obsEntrega || '',
          'Condição na devolução': it.condDevolucao || '',
          'Data de devolução': fmtDataBR(it.dataDevolucao),
          'Obs. devolução': it.obsDevolucao || '',
          'Trocado': it.trocado ? 'Sim' : ''
        }));
      });
    });
  return out;
}
function linhasInvDe(lista){
  return [...lista].sort((a,b)=> (a.nome||'').localeCompare(b.nome||'')).map(i=>({
    'Nome / descrição': i.nome || '',
    'Patrimônio': i.patrimonio || '',
    'Setor': i.setor || '',
    'Fabricante': i.fabricante || '',
    'Modelo': i.modelo || '',
    'Alocado a': (i.alocadoA && (i.quantidadeAlocada||0) > 0) ? i.alocadoA : '',
    'Quantidade total': i.quantidade || 1,
    'Alocado': i.quantidadeAlocada || 0,
    'Em manutenção': i.quantidadeManutencao || 0,
    'Disponível': i.baixado ? 0 : invDisponivel(i),
    'Posse': i.posse === 'locado' ? 'Locado' : 'Próprio',
    'Fornecedor / locadora': i.fornecedor || '',
    'Vencimento da locação': fmtDataBR(i.vencimentoLocacao),
    'Locado desde': inicioContinuo(i) ? fmtDataBR(inicioContinuo(i)) : '',
    'Valor mensal atual (R$)': periodoAberto(i) ? (Number(periodoAberto(i).valorMensal) || 0) : '',
    'Locação acumulada (R$)': round2(locacaoResumo(i).total),
    'Manutenção acumulada (R$)': round2(manutResumo(i).total),
    'Total gasto (R$)': round2(locacaoResumo(i).total + manutResumo(i).total),
    'Situação': i.baixado ? 'Baixado' : invPillInfo(i).label,
    'Cadastrado em': fmtDataBR(i.criadoEm),
    'Última alteração': i.atualizadoEm ? fmtDate(i.atualizadoEm) : ''
  }));
}
function linhasColabDe(lista){
  return [...lista].sort((a,b)=> (a.nome||'').localeCompare(b.nome||'')).map(c=>{
    const dele = fichas.filter(f => (f.colaborador||'').toLowerCase() === (c.nome||'').toLowerCase());
    return {
      'Nome': c.nome || '',
      'Função': c.funcao || '',
      'Matrícula': c.matricula || '',
      'Região': regiaoDeColab(c),
      'Centro de custo': centroCustoDeColab(c),
      'Fichas registradas': dele.length,
      'Fichas em aberto': dele.filter(f=>f.status==='aberta').length
    };
  });
}
function linhasLocacoesDe(lista){
  const out = [];
  [...lista].sort((a,b)=> (a.nome||'').localeCompare(b.nome||'')).forEach(i=>{
    const ps = periodosDe(i);
    const base = {'Equipamento': i.nome || '', 'Patrimônio': i.patrimonio || '', 'Posse': i.posse === 'locado' ? 'Locado' : 'Próprio'};
    const linha = (tipo, ini, fim, dias, valor, sub, obs) => out.push(Object.assign({}, base, {'Tipo de linha': tipo, 'Início': ini, 'Fim': fim, 'Duração (dias)': dias, 'Duração': fmtDuracao(dias), 'Valor mensal (R$)': valor, 'Subtotal (R$)': sub, 'Observação': obs || ''}));
    ps.forEach((p,k)=>{
      if(k > 0 && ps[k-1].fim){ const gap = diffDias(ps[k-1].fim, p.inicio); if(gap > 0) linha('Sem locação', fmtDataBR(ps[k-1].fim), fmtDataBR(p.inicio), gap, '', ''); }
      linha('Locação', fmtDataBR(p.inicio), p.fim ? fmtDataBR(p.fim) : 'em curso', periodoDias(p), Number(p.valorMensal) || 0, round2(periodoValor(p)), p.obs);
    });
    const ult = ps[ps.length-1];
    if(ult && ult.fim){ const parado = diffDias(ult.fim, hojeISO()); if(parado > 0) linha('Sem locação', fmtDataBR(ult.fim), 'até hoje', parado, '', ''); }
  });
  return out;
}
function linhasManutencoesDe(lista){
  const out = [];
  [...lista].sort((a,b)=> (a.nome||'').localeCompare(b.nome||'')).forEach(i=>{
    (i.manutencoes||[]).slice().sort((a,b)=> (a.dataEnvio||'').localeCompare(b.dataEnvio||'')).forEach(m=>{
      out.push({
        'Equipamento': i.nome || '', 'Patrimônio': i.patrimonio || '',
        'Enviado em': fmtDataBR(m.dataEnvio), 'Retornou em': m.dataRetorno ? fmtDataBR(m.dataRetorno) : 'em manutenção',
        'Dias na manutenção': Math.max(0, diffDias(m.dataEnvio, m.dataRetorno || hojeISO())),
        'Motivo': m.motivo || '', 'Oficina': m.oficina || '', 'Custo (R$)': Number(m.valor) || 0, 'Observação': m.obsRetorno || ''
      });
    });
  });
  return out;
}
function slugNome(s){
  return (s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]+/g,'-').replace(/^-|-$/g,'').slice(0,40);
}
async function gerarExcel(dados, sufixo){
  const wb = XLSX.utils.book_new();
  let abas = 0;
  const add = (linhas, nome, vazio) => {
    if(!linhas.length && !dados.vazios) return;
    const ws = XLSX.utils.json_to_sheet(linhas.length ? linhas : [vazio]);
    const cols = [...new Set((linhas.length ? linhas : [vazio]).flatMap(l => Object.keys(l)))];
    ws['!cols'] = cols.map(c => ({ wch: Math.min(40, Math.max(c.length + 2, ...linhas.map(l => String(l[c] ?? '').length + 2))) }));
    XLSX.utils.book_append_sheet(wb, ws, nome);
    abas++;
  };
  add(linhasFichasDe(dados.fichasList), 'Fichas', {'Aviso':'Nenhuma ficha registrada'});
  add(linhasInvDe(dados.invList), 'Inventário', {'Aviso':'Nenhum item cadastrado'});
  add(linhasLocacoesDe(dados.invList), 'Histórico de locação', {'Aviso':'Nenhum período de locação registrado'});
  add(linhasManutencoesDe(dados.invList), 'Manutenções', {'Aviso':'Nenhuma manutenção registrada'});
  add(linhasColabDe(dados.colabList), 'Colaboradores', {'Aviso':'Nenhum colaborador cadastrado'});
  if(!abas){ showToast('Não há dados para exportar com essa escolha.'); return; }
  const array = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([array], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const hoje = new Date().toISOString().slice(0,10);
  await downloadsNs.save({ filename: `bigode-${sufixo}-${hoje}.xlsx`, data: blob });
  showToast('Relatório exportado.');
}

/* ---------- Botão "Exportar Excel": pergunta o que exportar ---------- */
async function exportarExcel(){
  if(!downloadsNs || !window.XLSX){ showToast('Exportação indisponível neste ambiente.'); return; }
  try{
    const nomeAba = {fichas:'Fichas', inventario:'Inventário', colaboradores:'Colaboradores'}[currentSection] || 'Fichas';
    const listaAtual = currentSection === 'inventario' ? getInvFiltrado()
                     : currentSection === 'colaboradores' ? getColabFiltrados()
                     : getFichasFiltradas();
    const opcoes = [], modos = [];
    if(totalSelecionado()){
      const partes = [];
      if(sel.fichas.size) partes.push(`${sel.fichas.size} ficha(s)`);
      if(sel.inv.size) partes.push(`${sel.inv.size} item(ns) do inventário`);
      if(sel.colab.size) partes.push(`${sel.colab.size} colaborador(es) + as fichas deles`);
      opcoes.push(`Somente os selecionados: ${partes.join(', ')}`); modos.push('sel');
    }
    opcoes.push(`Tudo que está listado agora em ${nomeAba} (${listaAtual.length})`); modos.push('lista');
    opcoes.push('Tudo: fichas, inventário e colaboradores'); modos.push('tudo');

    const pick = await askChoice('O que exportar?', 'Dica: marque as caixinhas nas listas para exportar só alguns colaboradores, fichas ou itens.', opcoes);
    if(pick === null) return;
    const modo = modos[pick];

    const nomesDe = (cs)=> new Set(cs.map(c => (c.nome||'').toLowerCase()));
    const fichasDe = (cs, extra)=>{ const nm = nomesDe(cs); return fichas.filter(f => (extra && extra.has(f.id)) || nm.has((f.colaborador||'').toLowerCase())); };

    let dados, sufixo = 'relatorio';
    if(modo === 'tudo'){
      dados = { fichasList: fichas, invList: inventario, colabList: colaboradores, vazios: true };
    } else if(modo === 'sel'){
      const cs = colaboradores.filter(c => sel.colab.has(c.id));
      dados = { fichasList: fichasDe(cs, sel.fichas), invList: inventario.filter(i => sel.inv.has(i.id)), colabList: cs };
      sufixo = (cs.length === 1 && !sel.fichas.size && !sel.inv.size) ? 'colaborador-' + slugNome(cs[0].nome) : 'selecao';
    } else {
      if(currentSection === 'inventario') dados = { fichasList: [], invList: listaAtual, colabList: [] };
      else if(currentSection === 'colaboradores') dados = { fichasList: fichasDe(listaAtual), invList: [], colabList: listaAtual };
      else dados = { fichasList: listaAtual, invList: [], colabList: [] };
      sufixo = 'lista-' + currentSection;
    }
    await gerarExcel(dados, sufixo);
  }catch(e){
    if(e && e.code === 'declined') return;
    console.error('Erro ao exportar:', e);
    showToast('Não foi possível gerar o arquivo Excel.');
  }
}
document.getElementById('btn-export-excel').addEventListener('click', exportarExcel);
function renderUnavailable(){
  document.getElementById('top-actions').innerHTML = "";
  document.getElementById('ficha-list').innerHTML = `<div class="empty"><b>Não foi possível conectar ao banco de dados</b>Abra esta página pelo link publicado no Claude.</div>`;
  document.getElementById('view-dashboard').classList.remove('hidden');
}
function populateFuncaoSelect(){
  // Opções = funções padrão do sistema + todos os cargos já cadastrados nos colaboradores
  const sel = document.getElementById('colab-f-funcao');
  const atual = sel.value;
  const nomes = [...new Set([...FUNCOES_NOMES, ...colaboradores.map(c=>c.funcao).filter(Boolean)])].sort((a,b)=>a.localeCompare(b));
  sel.innerHTML = nomes.map(n=>`<option value="${escapeHtml(n)}">${escapeHtml(n)}</option>`).join('');
  if(atual && nomes.includes(atual)) sel.value = atual;
}

/* ============================================================
   TOAST / CONFIRM
   ============================================================ */
let toastTimer;
function showToast(msg){
  const t = document.getElementById('toast');
  t.textContent = msg; t.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(()=> t.classList.remove('show'), 2800);
}
function askChoice(title, msg, options){
  return new Promise((resolve)=>{
    const root = document.getElementById('confirm-root');
    root.innerHTML = `<div class="confirm-overlay"><div class="confirm-box">
      <h3>${title}</h3><p>${msg}</p>
      <div style="display:flex; flex-direction:column; gap:8px; margin-bottom:12px;">
        ${options.map((o,i)=>`<button class="btn btn-ghost btn-sm" style="text-align:left;" data-choice="${i}">${o}</button>`).join('')}
      </div>
      <div class="confirm-actions"><button class="btn btn-ghost btn-sm" id="choice-cancel">Cancelar</button></div>
    </div></div>`;
    root.querySelectorAll('[data-choice]').forEach(b=>{ b.onclick = ()=>{ root.innerHTML=''; resolve(+b.dataset.choice); }; });
    document.getElementById('choice-cancel').onclick = ()=>{ root.innerHTML=''; resolve(null); };
  });
}
function askConfirm(title, msg, okLabel){
  return new Promise((resolve)=>{
    const root = document.getElementById('confirm-root');
    root.innerHTML = `<div class="confirm-overlay"><div class="confirm-box">
      <h3>${title}</h3><p>${msg}</p>
      <div class="confirm-actions">
        <button class="btn btn-ghost btn-sm" id="conf-cancel">Cancelar</button>
        <button class="btn btn-danger btn-sm" id="conf-ok">${okLabel}</button>
      </div></div></div>`;
    document.getElementById('conf-cancel').onclick = ()=>{ root.innerHTML=""; resolve(false); };
    document.getElementById('conf-ok').onclick = ()=>{ root.innerHTML=""; resolve(true); };
  });
}

/* ============================================================
   ROTEAMENTO
   ============================================================ */
function hideAllViews(){
  ['view-dashboard','view-inventario','view-colaboradores','view-select-funcao','view-ficha'].forEach(id=>{
    document.getElementById(id).classList.add('hidden');
  });
}
function showSectionTabs(show){ document.getElementById('section-tabs').style.display = show ? 'flex' : 'none'; }

function showDashboard(){
  currentFichaId = null; currentSection = 'fichas';
  hideAllViews();
  document.getElementById('view-dashboard').classList.remove('hidden');
  showSectionTabs(true); updateSectionTabsUI();
  renderTopActionsDashboard();
  renderStats();
  renderDashboard();
}
function showInventario(){
  currentSection = 'inventario';
  hideAllViews();
  document.getElementById('view-inventario').classList.remove('hidden');
  showSectionTabs(true); updateSectionTabsUI();
  document.getElementById('top-actions').innerHTML = canManage ? '' : `<span class="funcao-badge" style="background:var(--line); color:var(--ink-soft);">MODO CONSULTA</span>`;
  renderInventario();
}
function showColaboradores(){
  currentSection = 'colaboradores';
  hideAllViews();
  document.getElementById('view-colaboradores').classList.remove('hidden');
  showSectionTabs(true); updateSectionTabsUI();
  document.getElementById('top-actions').innerHTML = canManage ? '' : `<span class="funcao-badge" style="background:var(--line); color:var(--ink-soft);">MODO CONSULTA</span>`;
  renderColaboradores();
}
function updateSectionTabsUI(){
  document.querySelectorAll('.section-tab').forEach(b=> b.classList.toggle('active', b.dataset.sec===currentSection));
}
document.querySelectorAll('.section-tab').forEach(btn=>{
  btn.addEventListener('click', ()=>{
    if(btn.dataset.sec==='fichas') showDashboard();
    else if(btn.dataset.sec==='inventario') showInventario();
    else showColaboradores();
  });
});

function showSelectFuncao(){
  if(!canManage){ showDashboard(); return; }
  hideAllViews(); showSectionTabs(false);
  document.getElementById('view-select-funcao').classList.remove('hidden');
  document.getElementById('top-actions').innerHTML = "";
  renderFuncaoGrid();
}
function showFicha(id){
  currentFichaId = id;
  hideAllViews(); showSectionTabs(false);
  document.getElementById('view-ficha').classList.remove('hidden');
  renderTopActionsFicha();
  renderFicha();
  window.scrollTo({top:0, behavior:'instant'});
}

function renderTopActionsDashboard(){
  if(canManage){
    document.getElementById('top-actions').innerHTML = `<button class="btn btn-primary" id="btn-nova-ficha">+ Nova ficha</button>`;
    document.getElementById('btn-nova-ficha').onclick = showSelectFuncao;
  } else {
    document.getElementById('top-actions').innerHTML = `<span class="funcao-badge" style="background:var(--line); color:var(--ink-soft);">MODO CONSULTA</span>`;
  }
}
/* ---------- Prazo das fichas em aberto: aviso aos 30 dias e renovação ---------- */
const LIMITE_DIAS_FICHA = 30;
function fichaBasePrazo(f){ return f.renovadoEm || f.dataAssinaturaRetirada || null; }
function fichaDiasPrazo(f){ const b = fichaBasePrazo(f); return b ? Math.floor((Date.now() - new Date(b).getTime())/86400000) : 0; }
function fichaVencida(f){ return f.status === 'aberta' && !!fichaBasePrazo(f) && fichaDiasPrazo(f) >= LIMITE_DIAS_FICHA; }
function fichaProximoAviso(f){ const b = fichaBasePrazo(f); return b ? new Date(new Date(b).getTime() + LIMITE_DIAS_FICHA*86400000).toISOString() : null; }
async function renovarFicha(id){
  if(!canManage) return;
  const f = getFichaById(id);
  if(!f || f.status !== 'aberta') return;
  const ativos = f.itens.filter(i => !i.trocado && !i.liberadoManual).length;
  const v = await askForm('Renovar prazo da ficha',
    `${escapeHtml(f.colaborador || 'Colaborador')} continua com ${ativos} item${ativos===1?'':'ns'}. O aviso volta em ${LIMITE_DIAS_FICHA} dias.`,
    [{key:'obs', label:'Observação (opcional)', placeholder:'Ex.: obra em andamento até dezembro'}], 'Renovar por 30 dias');
  if(!v) return;
  const agora = new Date().toISOString();
  f.renovadoEm = agora;
  f.renovacoes = (f.renovacoes || []).concat([{data: agora, por: myId || null, obs: v.obs || ''}]);
  await persistFichaNow(f);
  showToast('Prazo renovado por 30 dias.');
  if(viewVisible('view-ficha')){ renderTopActionsFicha(); renderItems(f); }
  else if(viewVisible('view-dashboard')){ renderDashboard(); renderStats(); }
}
function renderTopActionsFicha(){
  const f = getFichaById(currentFichaId);
  let html = "";
  if(f && f.status === 'concluida'){ html = `<button class="btn btn-ghost" id="btn-print">Imprimir / PDF</button>`; }
  if(f && f.status === 'aberta' && canManage){ html += ` <button class="btn ${fichaVencida(f) ? 'btn-primary' : 'btn-ghost'}" id="btn-renovar">Renovar prazo (30 dias)</button>`; }
  if(!canManage){ html += ` <span class="funcao-badge" style="background:var(--line); color:var(--ink-soft);">SOMENTE LEITURA</span>`; }
  document.getElementById('top-actions').innerHTML = html;
  const p = document.getElementById('btn-print');
  if(p) p.onclick = ()=> window.print();
  const rn = document.getElementById('btn-renovar');
  if(rn) rn.onclick = ()=> renovarFicha(currentFichaId);
}

/* ============================================================
   PAINEL: INDICADORES
   ============================================================ */
let dashPeriodo = '30d';
function fichasNoPeriodo(){
  if(dashPeriodo === 'todos') return fichas;
  const agora = new Date();
  let desde;
  if(dashPeriodo === '7d') desde = new Date(agora.getTime() - 7*86400000);
  else if(dashPeriodo === '30d') desde = new Date(agora.getTime() - 30*86400000);
  else if(dashPeriodo === 'mes') desde = new Date(agora.getFullYear(), agora.getMonth(), 1);
  else desde = new Date(0);
  return fichas.filter(f => new Date(f.criadoEm) >= desde);
}

function renderStats(){
  const grid = document.getElementById('stat-grid');
  if(!grid) return;
  const fichasPeriodo = fichasNoPeriodo();
  const abertas = fichasPeriodo.filter(f=>f.status==='aberta');
  const concluidas = fichasPeriodo.filter(f=>f.status==='concluida');
  const rascunhos = fichasPeriodo.filter(f=>f.status==='rascunho' && (f.colaborador || f.itens.some(i=>i.condEntrega)));
  const disponiveis = inventario.filter(i=>!i.baixado && invDisponivel(i) > 0).length;
  const manutencao = inventario.filter(i=>(i.quantidadeManutencao||0) > 0 && !i.baixado).length;
  const locados = inventario.filter(i=>i.posse==='locado' && !i.baixado).length;

  grid.innerHTML = `
    <div class="stat-card warn"><div class="stat-num">${abertas.length}</div><div class="stat-label">Fichas em aberto</div></div>
    <div class="stat-card ok"><div class="stat-num">${disponiveis}</div><div class="stat-label">Itens disponíveis</div></div>
    <div class="stat-card alert"><div class="stat-num">${manutencao}</div><div class="stat-label">Em manutenção</div></div>
    <div class="stat-card brand"><div class="stat-num">${locados}</div><div class="stat-label">Itens locados</div></div>
  `;

  renderCharts(rascunhos.length, abertas.length, concluidas.length, disponiveis, manutencao, inventario.filter(i=>!i.baixado && invDisponivel(i)<=0 && (i.quantidadeManutencao||0)<=0).length);

  const todasAbertas = fichas.filter(f=>f.status==='aberta');
  const vencidasTodas = todasAbertas.filter(fichaVencida).sort((a,b)=> fichaDiasPrazo(b) - fichaDiasPrazo(a));
  const atrasadas = vencidasTodas.slice(0,8).map(f=>({f, dias: fichaDiasPrazo(f)}));
  const maisVencidas = Math.max(0, vencidasTodas.length - atrasadas.length);

  const alertBox = document.getElementById('top-alert-list');
  const locAlertas = inventario
    .map(i => ({i, info: locacaoInfo(i)}))
    .filter(x => x.info && (x.info.status === 'vencida' || x.info.status === 'vencendo'))
    .sort((a,b)=> diasParaVencer(a.i) - diasParaVencer(b.i))
    .slice(0,5);
  alertBox.innerHTML = locAlertas.map(x=>`
      <div class="top-alert-row" ${x.info.status==='vencendo' ? 'style="border-color:var(--regular); background:var(--regular-bg);"' : ''}>
        <span>🔧 ${escapeHtml(x.i.nome)} (Patr. ${escapeHtml(x.i.patrimonio)}) — ${x.info.label}${x.i.fornecedor ? ' — ' + escapeHtml(x.i.fornecedor) : ''}</span>
        <button class="btn btn-ghost btn-sm" data-goinv="${x.i.id}">Ver item</button>
      </div>`).join('')
    + atrasadas.map(x=>`
      <div class="top-alert-row">
        <span>${escapeHtml(x.f.colaborador||'Colaborador não informado')} — ${escapeHtml(x.f.funcao)} — prazo de ${LIMITE_DIAS_FICHA} dias vencido (${x.dias} dias com os itens${x.f.renovacoes && x.f.renovacoes.length ? ', renovada ' + x.f.renovacoes.length + 'x' : ''})</span>
        <span style="display:flex; gap:6px;"><button class="btn btn-ghost btn-sm" data-open="${x.f.id}">Abrir</button>${canManage ? `<button class="btn btn-primary btn-sm" data-renovar="${x.f.id}">Renovar</button>` : ''}</span>
      </div>`).join('')
    + (maisVencidas ? `<div class="top-alert-row"><span>… e mais ${maisVencidas} ficha${maisVencidas>1?'s':''} com prazo vencido (veja na lista de fichas)</span></div>` : '');
  alertBox.querySelectorAll('[data-open]').forEach(b=> b.addEventListener('click', ()=> showFicha(b.dataset.open)));
  alertBox.querySelectorAll('[data-renovar]').forEach(b=> b.addEventListener('click', ()=> renovarFicha(b.dataset.renovar)));
  alertBox.querySelectorAll('[data-goinv]').forEach(b=> b.addEventListener('click', ()=>{ const u = getInvById(b.dataset.goinv); expandedInvId = b.dataset.goinv; expandedGroup = u ? grupoKey(u) : null; showInventario(); }));
}

function barRow(label, count, max, color){
  const pct = max > 0 ? Math.round((count/max)*100) : 0;
  return `<div class="bar-row">
    <span class="bar-label">${label}</span>
    <div class="bar-track"><div class="bar-fill" style="width:${pct}%; background:${color};"></div></div>
    <span class="bar-count">${count}</span>
  </div>`;
}
function renderCharts(rascunhos, abertas, concluidas, disponiveis, manutencao, alocados){
  const el = document.getElementById('chart-grid');
  if(!el) return;
  const maxF = Math.max(rascunhos, abertas, concluidas, 1);
  const maxI = Math.max(disponiveis, manutencao, alocados, 1);
  el.innerHTML = `
    <div class="chart-card">
      <h3>Fichas por status (no período)</h3>
      ${barRow('Rascunho', rascunhos, maxF, 'var(--line-strong)')}
      ${barRow('Em aberto', abertas, maxF, 'var(--regular)')}
      ${barRow('Concluídas', concluidas, maxF, 'var(--good)')}
    </div>
    <div class="chart-card">
      <h3>Inventário por situação</h3>
      ${barRow('Disponível', disponiveis, maxI, 'var(--good)')}
      ${barRow('Alocado', alocados, maxI, 'var(--regular)')}
      ${barRow('Manutenção', manutencao, maxI, 'var(--ruim)')}
    </div>
  `;
}
document.getElementById('dash-periodo').addEventListener('change', (e)=>{ dashPeriodo = e.target.value; renderStats(); });

/* ============================================================
   SELEÇÃO DE FUNÇÃO
   ============================================================ */
function renderFuncaoGrid(){
  const grid = document.getElementById('funcao-grid');
  let html = `<button type="button" class="funcao-card" data-esp="1" style="border-color:var(--brand);">
      <div class="fc-name">Entrega esporádica <span class="fc-arrow">→</span></div>
      <div class="fc-count">Itens avulsos, fora do padrão da função</div>
    </button>`;
  html += FUNCOES_NOMES.map(nome => `
    <button type="button" class="funcao-card" data-funcao="${nome}">
      <div class="fc-name">${nome} <span class="fc-arrow">→</span></div>
      <div class="fc-count">${FUNCOES[nome].length} equipamentos no termo</div>
    </button>`).join('');
  grid.innerHTML = html;
  grid.querySelectorAll('.funcao-card').forEach(card=>{
    card.addEventListener('click', ()=>{
      if(!canManage) return;
      const f = card.dataset.esp ? novaFichaEsporadica() : novaFichaParaFuncao(card.dataset.funcao);
      localOnlyFichas[f.id] = f; fichas = [f, ...fichas];
      showFicha(f.id);
      persistFichaNow(f);
    });
  });
}
document.getElementById('btn-back-funcao').addEventListener('click', showDashboard);

/* ============================================================
   DASHBOARD: LISTA DE FICHAS
   ============================================================ */
function getFichasFiltradas(){
  let items = [...fichas].filter(f => !(f.status === 'rascunho' && !f.colaborador && !f.itens.some(i=>i.condEntrega)));
  if(searchTerm.trim()){
    const q = searchTerm.trim().toLowerCase();
    items = items.filter(f => (f.colaborador||"").toLowerCase().includes(q) || (f.funcao||"").toLowerCase().includes(q));
  }
  if(filtroAtual !== 'todas'){
    items = items.filter(f => f.status === filtroAtual || (filtroAtual==='aberta' && f.status==='rascunho'));
  }
  if(fichaRegiao !== 'todas') items = items.filter(f => regiaoDaFicha(f) === fichaRegiao);
  if(fichaCC !== 'todas') items = items.filter(f => centroCustoDaFicha(f) === fichaCC);
  items.sort((a,b)=> new Date(b.criadoEm) - new Date(a.criadoEm));
  return items;
}
const fichasAbertas = new Set();   // colaboradores cuja lista de fichas está aberta
function chaveColab(f){ return (f.colaborador||'').trim().toLowerCase().replace(/\s+/g, ' '); }
function renderDashboard(){
  const list = document.getElementById('ficha-list');
  const items = getFichasFiltradas();
  renderSelBar('fichas', items.map(x=>x.id));

  if(items.length === 0){
    list.innerHTML = `<div class="empty"><b>Nenhuma ficha encontrada</b>${canManage ? 'Crie uma nova ficha selecionando a função do colaborador.' : 'Ainda não há fichas registradas.'}</div>`;
    return;
  }
  // Uma linha por colaborador (menu suspenso); dentro dela ficam todas as fichas dele
  const grupos = new Map();
  items.forEach(f => {
    const k = chaveColab(f);
    if(!grupos.has(k)) grupos.set(k, {key:k, nome:(f.colaborador||'').trim(), fichas:[]});
    grupos.get(k).fichas.push(f);
  });
  const lista = [...grupos.values()].sort((a,b)=> (a.nome||'~').localeCompare(b.nome||'~', 'pt-BR'));
  const abrirTudo = searchTerm.trim() !== '' && lista.length <= 3;

  const linhaFicha = (f) => {
    const statusLabel = f.status === 'concluida' ? 'Concluída' : (f.status === 'aberta' ? 'Em aberto' : 'Rascunho');
    const statusClass = f.status === 'concluida' ? 'concluida' : 'aberta';
    const qtdItens = f.itens.filter(i => i.condEntrega || f.importada).length;
    return `<div class="ficha-row status-${statusClass}" data-id="${f.id}">
      <div class="fr-main">
        <div class="fr-name">${selBox('fichas', f.id)}${escapeHtml(f.funcao || 'Ficha')}${f.importada ? ' <span class="pill modelo">importada do inventário</span>' : ''}</div>
        <div class="fr-meta">
          <span>${qtdItens}/${f.itens.length||0} itens registrados</span>
          <span>Criada em ${fmtDate(f.criadoEm).split(' às')[0]}</span>
        </div>
      </div>
      <div class="fr-right"><span class="status-badge ${statusClass}">${statusLabel.toUpperCase()}</span></div>
    </div>`;
  };

  list.innerHTML = lista.map(g => {
    const ids = g.fichas.map(f => f.id);
    const todos = ids.every(id => sel.fichas.has(id)), algum = ids.some(id => sel.fichas.has(id));
    const chk = `<input type="checkbox" class="sel-box" data-kind="fichas" data-ids="${ids.join(',')}" ${todos ? 'checked' : ''} ${algum && !todos ? 'data-indet="1"' : ''} title="Selecionar todas as fichas deste colaborador para exportar">`;
    const abertas = g.fichas.filter(f => f.status !== 'concluida').length;
    const concl = g.fichas.length - abertas;
    const aberto = abrirTudo || fichasAbertas.has(g.key);
    const f0 = g.fichas[0];
    const nItens = g.fichas.filter(f => f.status !== 'concluida').reduce((n,f)=> n + f.itens.filter(i=>!i.trocado && !i.liberadoManual).length, 0);
    return `<div class="list-row ${abertas ? 'situacao-alocado' : 'situacao-disponivel'}" data-gkey="${escapeHtml(g.key)}">
      <div class="list-row-head" style="cursor:pointer;">
        <div>
          <div class="lr-name">${chk}<span class="fc-arrow">${aberto ? '▾' : '▸'}</span> ${escapeHtml(g.nome) || 'Colaborador não informado'} <span class="pill modelo">${g.fichas.length} ficha${g.fichas.length>1?'s':''}</span></div>
          <div class="lr-meta">
            ${regiaoDaFicha(f0) ? `<span>${regiaoDaFicha(f0)}</span>` : ''}
            ${centroCustoDaFicha(f0) ? `<span>${escapeHtml(centroCustoDaFicha(f0))}</span>` : ''}
            ${g.fichas.some(fichaVencida) ? `<span style="color:var(--ruim); font-weight:700;">${g.fichas.filter(fichaVencida).length} com prazo de 30 dias vencido</span>` : ''}
            ${abertas ? `<span>${abertas} em aberto${nItens ? ' · ' + nItens + (nItens>1 ? ' itens' : ' item') + ' com o colaborador' : ''}</span>` : ''}
            ${concl ? `<span>${concl} concluída${concl>1?'s':''}</span>` : ''}
          </div>
        </div>
        <span class="status-badge ${abertas ? 'aberta' : 'concluida'}">${abertas ? abertas + ' EM ABERTO' : 'CONCLUÍDAS'}</span>
      </div>
      ${aberto ? `<div class="row-detail" style="display:flex; flex-direction:column; gap:8px;">${g.fichas.map(linhaFicha).join('')}</div>` : ''}
    </div>`;
  }).join('');
  list.querySelectorAll('[data-indet]').forEach(el => { el.indeterminate = true; });
  list.querySelectorAll('.list-row-head').forEach(head=> head.addEventListener('click', (e)=>{
    if(e.target.closest('button, input, label')) return;
    const k = head.parentElement.dataset.gkey;
    if(fichasAbertas.has(k)) fichasAbertas.delete(k); else fichasAbertas.add(k);
    renderDashboard();
  }));
  list.querySelectorAll('.ficha-row').forEach(row=> row.addEventListener('click', ()=> showFicha(row.dataset.id)));
}
document.getElementById('search-input').addEventListener('input', (e)=>{ searchTerm = e.target.value; renderDashboard(); });
document.getElementById('btn-back').addEventListener('click', ()=>{
  const f = getFichaById(currentFichaId);
  if(canManage && f && f.status === 'rascunho' && !f.colaborador && !f.itens.some(i=>i.condEntrega)){
    fichas = fichas.filter(x => x.id !== f.id);
    delete localOnlyFichas[f.id];
    deleteFichaRemote(f.id);
  }
  showDashboard();
});

/* ============================================================
   INVENTÁRIO
   ============================================================ */
let invEditingId = null;
function parsePatrimonios(txt){
  return (txt||'').split(/[\n,;]+/).map(s => s.trim()).filter(Boolean);
}
function atualizaDicaPatrimonios(){
  const n = parsePatrimonios(document.getElementById('inv-f-patrimonios').value).length;
  document.getElementById('inv-f-pat-hint').textContent = invEditingId ? 'Na edição, informe um único patrimônio.' : (n ? `Serão cadastradas ${n} ferramenta${n>1?'s':''} (uma por patrimônio).` : 'Cada patrimônio vira uma ferramenta própria no inventário.');
}
document.getElementById('inv-f-patrimonios').addEventListener('input', atualizaDicaPatrimonios);
document.getElementById('btn-toggle-inv-form').addEventListener('click', ()=>{
  if(!canManage) return;
  invEditingId = null;
  clearInvForm();
  document.getElementById('inv-form-title').textContent = 'Novo item de inventário';
  document.getElementById('inv-new-form').classList.toggle('hidden');
});
document.getElementById('btn-cancel-inv-form').addEventListener('click', ()=>{
  invEditingId = null;
  document.getElementById('inv-new-form').classList.add('hidden');
});
document.getElementById('inv-f-posse').addEventListener('change', (e)=>{
  document.getElementById('inv-f-fornecedor-wrap').classList.toggle('hidden', e.target.value !== 'locado');
});
function clearInvForm(){
  ['inv-f-nome','inv-f-patrimonios','inv-f-fornecedor','inv-f-vencimento','inv-f-inicio','inv-f-valor'].forEach(id=> document.getElementById(id).value = '');
  document.getElementById('inv-f-posse').value = 'proprio';
  document.getElementById('inv-f-fornecedor-wrap').classList.add('hidden');
  document.getElementById('inv-f-novo-wrap').classList.remove('hidden');
  document.getElementById('inv-f-edit-note').classList.add('hidden');
  atualizaDicaPatrimonios();
}
function abrirEdicaoInv(id){
  const item = getInvById(id);
  if(!item || !canManage) return;
  invEditingId = id;
  clearInvForm();
  document.getElementById('inv-f-nome').value = item.nome || '';
  document.getElementById('inv-f-patrimonios').value = item.patrimonio || '';
  document.getElementById('inv-f-posse').value = item.posse || 'proprio';
  document.getElementById('inv-f-fornecedor').value = item.fornecedor || '';
  document.getElementById('inv-f-vencimento').value = item.vencimentoLocacao || '';
  document.getElementById('inv-f-fornecedor-wrap').classList.toggle('hidden', item.posse !== 'locado');
  document.getElementById('inv-f-novo-wrap').classList.add('hidden');
  document.getElementById('inv-f-edit-note').classList.remove('hidden');
  document.getElementById('inv-form-title').textContent = 'Editar item de inventário';
  atualizaDicaPatrimonios();
  document.getElementById('inv-new-form').classList.remove('hidden');
  document.getElementById('inv-new-form').scrollIntoView({behavior:'smooth', block:'center'});
}
document.getElementById('btn-save-inv-form').addEventListener('click', ()=>{
  if(!canManage) return;
  const nome = document.getElementById('inv-f-nome').value.trim();
  const pats = parsePatrimonios(document.getElementById('inv-f-patrimonios').value);
  const posse = document.getElementById('inv-f-posse').value;
  const fornecedor = document.getElementById('inv-f-fornecedor').value.trim();
  const vencimento = document.getElementById('inv-f-vencimento').value;
  const inicio = document.getElementById('inv-f-inicio').value;
  const valor = Number(String(document.getElementById('inv-f-valor').value).replace(',', '.')) || 0;

  if(!nome){ showToast('Informe o nome ou descrição da ferramenta.'); return; }
  if(!pats.length){ showToast('Informe pelo menos um patrimônio.'); return; }
  if(invEditingId && pats.length !== 1){ showToast('Na edição, informe apenas um patrimônio.'); return; }
  const minusc = pats.map(p => p.toLowerCase());
  const repetido = pats.find((p,k) => minusc.indexOf(p.toLowerCase()) !== k);
  if(repetido){ showToast(`O patrimônio "${repetido}" aparece mais de uma vez na lista.`); return; }
  const usados = new Set(inventario.filter(x => x.id !== invEditingId).map(x => (x.patrimonio||'').toLowerCase()));
  const jaExiste = pats.filter(p => usados.has(p.toLowerCase()));
  if(jaExiste.length){ showToast(`Patrimônio já cadastrado: ${jaExiste.slice(0,3).join(', ')}.`); return; }
  if(posse === 'locado'){
    if(inicio && vencimento && vencimento <= inicio){ showToast('O vencimento precisa ser depois do início da locação.'); return; }
    if(!invEditingId && valor && !inicio){ showToast('Informe o início da locação para registrar o valor mensal.'); return; }
  }

  if(invEditingId){
    const item = getInvById(invEditingId);
    if(!item) return;
    const antes = {nome: item.nome || '', pat: item.patrimonio || '', posse: item.posse || 'proprio', forn: item.fornecedor || '', venc: item.vencimentoLocacao || ''};
    const depois = {nome, pat: pats[0], posse, forn: posse === 'locado' ? fornecedor : '', venc: posse === 'locado' ? vencimento : ''};
    const q = v => v ? `"${v}"` : '(vazio)';
    const nomePosse = v => v === 'locado' ? 'Locado' : 'Próprio';
    const mud = [];
    if(antes.nome !== depois.nome) mud.push(`Nome: ${q(antes.nome)} → ${q(depois.nome)}`);
    if(antes.pat !== depois.pat) mud.push(`Patrimônio: ${q(antes.pat)} → ${q(depois.pat)}`);
    if(antes.posse !== depois.posse) mud.push(`Posse: ${nomePosse(antes.posse)} → ${nomePosse(depois.posse)}`);
    if(antes.forn !== depois.forn) mud.push(`Locadora: ${q(antes.forn)} → ${q(depois.forn)}`);
    if(antes.venc !== depois.venc) mud.push(`Vencimento: ${antes.venc ? fmtDataBR(antes.venc) : '(vazio)'} → ${depois.venc ? fmtDataBR(depois.venc) : '(vazio)'}`);
    if(!mud.length){ showToast('Nenhuma alteração para salvar.'); return; }
    item.nome = depois.nome; item.patrimonio = depois.pat; item.posse = depois.posse;
    item.fornecedor = depois.forn; item.vencimentoLocacao = depois.venc;
    logInv(item, 'edicao', mud.join('; '));
    persistInv(item);
    showToast('Item atualizado.');
  } else {
    const novos = pats.map(p => {
      const item = novoItemInventario();
      item.nome = nome; item.patrimonio = p; item.posse = posse;
      item.fornecedor = posse === 'locado' ? fornecedor : ''; item.vencimentoLocacao = posse === 'locado' ? vencimento : '';
      if(posse === 'locado' && inicio){
        item.periodos = [{id: uid('per'), inicio, fim: null, valorMensal: valor, obs: ''}];
        logInv(item, 'locacao', `Locação iniciada em ${fmtDataBR(inicio)}${valor ? ' — ' + fmtBRL(valor) + '/mês' : ''}`);
      }
      localOnlyInv[item.id] = item;
      persistInv(item);
      return item;
    });
    inventario = [...novos, ...inventario];
    showToast(novos.length > 1 ? `${novos.length} ferramentas adicionadas ao inventário.` : 'Item adicionado ao inventário.');
  }
  invEditingId = null;
  clearInvForm();
  document.getElementById('inv-new-form').classList.add('hidden');
  renderInventario();
});
document.getElementById('inv-search-input').addEventListener('input', (e)=>{ invSearchTerm = e.target.value; renderInventario(); });

function getInvFiltrado(){
  let items = [...inventario];
  if(invSearchTerm.trim()){
    const q = invSearchTerm.trim().toLowerCase();
    items = items.filter(i => (i.nome||'').toLowerCase().includes(q) || (i.patrimonio||'').toLowerCase().includes(q));
  }
  if(invFiltro !== 'todos'){
    items = items.filter(i => {
      const p = invPillInfo(i);
      if(invFiltro === 'disponivel') return p.cls === 'disponivel';
      if(invFiltro === 'alocado') return p.cls === 'alocado';
      if(invFiltro === 'manutencao') return (i.quantidadeManutencao||0) > 0 && !i.baixado;
      if(invFiltro === 'baixado') return i.baixado;
      return true;
    });
  }
  if(invPosse !== 'todas') items = items.filter(i => i.posse === invPosse);
  if(invLocacao !== 'todas') items = items.filter(i => { const l = locacaoInfo(i); return l && l.status === invLocacao; });
  items.sort((a,b)=> (a.nome||'').localeCompare(b.nome||''));
  return items;
}
/* Registro de alteração: guarda data/hora e quem fez (só o id; o nome é resolvido na hora de exibir) */
function logInv(item, tipo, detalhes){
  item.historico = (item.historico||[]).concat([{tipo, data:new Date().toISOString(), detalhes, por: myId || null}]);
}
const nomesUsuarios = {};
async function carregarNomesUsuarios(ids){
  const novos = [...new Set(ids.filter(id => id && !(id in nomesUsuarios)))];
  if(!novos.length || !userNs || typeof userNs.profiles !== 'function') return;
  try{
    const ps = await userNs.profiles(novos);
    novos.forEach(id => { nomesUsuarios[id] = (ps && ps[id] && ps[id].name) || ''; });
  }catch(e){ novos.forEach(id => { nomesUsuarios[id] = ''; }); }
  if(novos.some(id => nomesUsuarios[id])) renderInventario();
}

/* Visualização: ferramentas com o mesmo nome aparecem como UM modelo; cada patrimônio continua sendo uma unidade
   própria e é contada normalmente (estatísticas, totais e Excel usam as unidades). */
let expandedGroup = null;
function grupoKey(i){ return (i.nome||'').trim().toLowerCase().replace(/\s+/g, ' '); }
function agruparInv(items){
  const g = new Map();
  items.forEach(i => { const k = grupoKey(i); if(!g.has(k)) g.set(k, {key:k, nome:(i.nome||'').trim(), units:[]}); g.get(k).units.push(i); });
  return [...g.values()].sort((a,b)=> a.nome.localeCompare(b.nome));
}
function situacaoTexto(u){
  if(!u.baixado && (u.quantidade||1) === 1 && (u.quantidadeAlocada||0) > 0 && u.alocadoA) return `Alocado a ${u.alocadoA}`;
  return invPillInfo(u).label;
}
function resumoGrupo(g){
  const r = {tot:0, disp:0, aloc:0, man:0, baix:0, prop:0, loc:0, venc:0, vencendo:0, ult:''};
  g.units.forEach(u=>{
    const q = u.quantidade || 1;
    r.tot += q;
    if(u.baixado) r.baix += q; else { r.disp += invDisponivel(u); r.aloc += (u.quantidadeAlocada||0); r.man += (u.quantidadeManutencao||0); }
    if(u.posse === 'locado') r.loc += q; else r.prop += q;
    const li = locacaoInfo(u);
    if(li && li.status === 'vencida') r.venc++; else if(li && li.status === 'vencendo') r.vencendo++;
    if(u.atualizadoEm && u.atualizadoEm > r.ult) r.ult = u.atualizadoEm;
  });
  return r;
}
function pillGrupo(r){
  if(r.disp > 0) return {cls:'disponivel', label:`${r.disp} disponíve${r.disp > 1 ? 'is' : 'l'}`};
  if(r.aloc > 0) return {cls:'alocado', label: r.tot === 1 ? 'Alocado' : 'Todos alocados'};
  if(r.man > 0) return {cls:'manutencao', label:'Em manutenção'};
  return {cls:'baixado', label:'Baixado'};
}
function renderUnidadeCard(u){
  const p = invPillInfo(u);
  const por = (u.atualizadoPor && nomesUsuarios[u.atualizadoPor]) ? ` · por ${escapeHtml(nomesUsuarios[u.atualizadoPor])}` : '';
  return `<div class="unidade-card">
    <div class="unidade-head">
      <div><b>Patrimônio ${escapeHtml(u.patrimonio)}</b> <span class="pill ${p.cls}">${escapeHtml(situacaoTexto(u))}</span> <span class="pill ${u.posse}">${u.posse === 'locado' ? 'Locado' : 'Próprio'}</span> ${locacaoBadge(u)}</div>
      <label class="sel-inline">${selBox('inv', u.id)} incluir na exportação</label>
    </div>
    <div class="lr-meta">${u.setor ? `<span>Setor: ${escapeHtml(u.setor)}</span>` : ''}${(u.fabricante || u.modelo) ? `<span>${escapeHtml([u.fabricante, u.modelo].filter(Boolean).join(' · '))}</span>` : ''}${locacaoLinha(u)}${u.atualizadoEm ? `<span>Última alteração: ${fmtDate(u.atualizadoEm)}${por}</span>` : ''}</div>
    ${renderInvDetail(u)}
  </div>`;
}
function renderGrupoAberto(g){
  const us = [...g.units].sort((a,b)=> (a.patrimonio||'').localeCompare(b.patrimonio||'', 'pt-BR', {numeric:true}));
  let und = us.find(u => u.id === expandedInvId);
  if(!und && us.length === 1) und = us[0];
  const opcoes = us.map(u => `<option value="${u.id}" ${und && u.id === und.id ? 'selected' : ''}>${escapeHtml(u.patrimonio)} — ${escapeHtml(situacaoTexto(u))}${u.posse === 'locado' ? ' · locado' : ''}</option>`).join('');
  return `<div class="row-detail">
    <div class="field" style="max-width:440px; margin-bottom:10px;">
      <label>Patrimônio</label>
      <select class="sel-unit" data-gkey="${escapeHtml(g.key)}"><option value="">Selecione o patrimônio...</option>${opcoes}</select>
    </div>
    ${und ? renderUnidadeCard(und) : '<div class="locked-note">Selecione um patrimônio para ver a situação e o histórico daquela unidade.</div>'}
  </div>`;
}
function renderInventario(){
  const list = document.getElementById('inv-list');
  const items = getInvFiltrado();
  renderSelBar('inv', items.map(x=>x.id));
  renderInvTotais(items);

  if(items.length === 0){
    list.innerHTML = `<div class="empty"><b>Nenhum item encontrado</b>${canManage ? 'Cadastre o primeiro item do inventário acima.' : 'Ainda não há itens cadastrados.'}</div>`;
    return;
  }
  const totalDoModelo = {};
  inventario.forEach(u => { const k = grupoKey(u); totalDoModelo[k] = (totalDoModelo[k]||0) + (u.quantidade||1); });

  list.innerHTML = agruparInv(items).map(g=>{
    const r = resumoGrupo(g), pill = pillGrupo(r);
    const ids = g.units.map(u => u.id);
    const todos = ids.every(id => sel.inv.has(id)), algum = ids.some(id => sel.inv.has(id));
    const chk = `<input type="checkbox" class="sel-box" data-kind="inv" data-ids="${ids.join(',')}" ${todos ? 'checked' : ''} ${algum && !todos ? 'data-indet="1"' : ''} title="Selecionar todas as unidades deste modelo para exportar">`;
    const deTotal = r.tot < totalDoModelo[g.key] ? ` de ${totalDoModelo[g.key]}` : '';
    return `<div class="list-row situacao-${pill.cls}" data-gkey="${escapeHtml(g.key)}">
      <div class="list-row-head">
        <div>
          <div class="lr-name">${chk}${escapeHtml(g.nome)} <span class="pill modelo">${r.tot}${deTotal} unidade${(r.tot > 1 || deTotal) ? 's' : ''}</span></div>
          <div class="lr-meta">
            ${g.units.length === 1 ? `<span>Patrimônio: ${escapeHtml(g.units[0].patrimonio)}</span>` : ''}
            ${r.disp ? `<span>${r.disp} disponíve${r.disp > 1 ? 'is' : 'l'}</span>` : ''}
            ${r.aloc ? `<span>${r.aloc} alocado${r.aloc > 1 ? 's' : ''}</span>` : ''}
            ${r.man ? `<span>${r.man} em manutenção</span>` : ''}
            ${r.baix ? `<span>${r.baix} baixado${r.baix > 1 ? 's' : ''}</span>` : ''}
            ${r.prop ? `<span class="pill proprio">Próprio${r.loc ? ' ×' + r.prop : ''}</span>` : ''}
            ${r.loc ? `<span class="pill locado">Locado${r.prop ? ' ×' + r.loc : ''}</span>` : ''}
            ${r.venc ? `<span class="pill manutencao">${r.venc} locação vencida${r.venc > 1 ? 's' : ''}</span>` : ''}
            ${r.vencendo ? `<span class="pill alocado">${r.vencendo} vence${r.vencendo > 1 ? 'm' : ''} em breve</span>` : ''}
            ${r.ult ? `<span>Alterado em ${fmtDate(r.ult)}</span>` : ''}
          </div>
        </div>
        <span class="pill ${pill.cls}">${pill.label}</span>
      </div>
      ${expandedGroup === g.key ? renderGrupoAberto(g) : ''}
    </div>`;
  }).join('');

  list.querySelectorAll('[data-indet]').forEach(el => { el.indeterminate = true; });
  list.querySelectorAll('.list-row-head').forEach(head=>{
    head.addEventListener('click', (e)=>{
      if(e.target.closest('button, input, label')) return;
      const k = head.parentElement.dataset.gkey;
      expandedGroup = (expandedGroup === k) ? null : k;
      expandedInvId = null;
      renderInventario();
    });
  });
  list.querySelectorAll('.sel-unit').forEach(s => s.addEventListener('change', ()=>{ expandedInvId = s.value || null; renderInventario(); }));
  if(canManage){
    list.querySelectorAll('[data-act]').forEach(btn=>{
      btn.addEventListener('click', (e)=>{ e.stopPropagation(); invAction(btn.dataset.id, btn.dataset.act); });
    });
    list.querySelectorAll('[data-edit-inv]').forEach(btn=>{
      btn.addEventListener('click', (e)=>{ e.stopPropagation(); abrirEdicaoInv(btn.dataset.editInv); });
    });
  }
  const aberta = expandedInvId ? getInvById(expandedInvId) : null;
  if(aberta) carregarNomesUsuarios([aberta.atualizadoPor, ...(aberta.historico||[]).map(h => h.por)]);
}
function renderInvTotais(items){
  const el = document.getElementById('inv-totais');
  if(!el) return;
  let loc = 0, man = 0;
  items.forEach(i => { loc += locacaoResumo(i).total; man += manutResumo(i).total; });
  el.innerHTML = (loc || man)
    ? `<span>Locação acumulada: <b>${fmtBRL(loc)}</b></span><span>Manutenção acumulada: <b>${fmtBRL(man)}</b></span><span>Total: <b>${fmtBRL(loc + man)}</b></span><span class="lr-meta">(${items.length} unidade(s) listada(s))</span>`
    : '';
}
function renderInvDetail(i){
  const rl = locacaoResumo(i), rm = manutResumo(i);
  const btn = (cls, iact, label, pid) => `<button class="btn ${cls} btn-sm" data-iact="${iact}" data-id="${i.id}"${pid ? ` data-pid="${pid}"` : ''}>${label}</button>`;

  /* --- resumo financeiro --- */
  const stats = `<div class="inv-stats">
    <div class="inv-stat"><span>Locação acumulada</span><b>${fmtBRL(rl.total)}</b></div>
    <div class="inv-stat"><span>Tempo locado</span><b>${fmtDuracao(rl.dias)}</b></div>
    <div class="inv-stat"><span>Sem locação</span><b>${fmtDuracao(rl.ocioso)}</b></div>
    <div class="inv-stat"><span>Manutenção acumulada</span><b>${fmtBRL(rm.total)}</b></div>
    <div class="inv-stat total"><span>Total gasto</span><b>${fmtBRL(rl.total + rm.total)}</b></div>
  </div>`;

  /* --- contrato atual + barra de prazo --- */
  let contrato = '';
  if(i.posse === 'locado'){
    contrato = `<div class="lr-meta" style="margin-bottom:8px;">Locadora: ${escapeHtml(i.fornecedor || 'não informada')} · Vencimento: ${i.vencimentoLocacao ? fmtDataBR(i.vencimentoLocacao) : 'não informado'}</div>`;
    const ini = inicioContinuo(i);
    if(ini && i.vencimentoLocacao){
      const total = diffDias(ini, i.vencimentoLocacao), dec = diffDias(ini, hojeISO());
      if(total > 0){
        const pct = Math.max(0, Math.min(100, Math.round(dec / total * 100)));
        const cor = dec > total ? 'var(--ruim)' : (pct >= 85 ? 'var(--regular)' : 'var(--good)');
        contrato += `<div class="prazo"><div class="lr-meta">Prazo do contrato: ${Math.max(0, dec)} de ${total} dias (${pct}%) · ${dec > total ? 'vencido há ' + fmtDuracao(dec - total) : 'faltam ' + fmtDuracao(total - dec)}</div><div class="bar-track" style="height:8px; margin:4px 0 10px;"><div class="bar-fill" style="width:${pct}%; background:${cor};"></div></div></div>`;
      }
    }
  }

  /* --- histórico de locação --- */
  const ps = periodosDe(i);
  let locHtml = '';
  ps.forEach((p,k)=>{
    if(k > 0 && ps[k-1].fim){
      const gap = diffDias(ps[k-1].fim, p.inicio);
      if(gap > 0) locHtml += `<div class="per-row gap">⏸ Sem locação: ${fmtDuracao(gap)} (${fmtDataBR(ps[k-1].fim)} → ${fmtDataBR(p.inicio)})</div>`;
    }
    locHtml += `<div class="per-row"><div><b>${fmtDataBR(p.inicio)} → ${p.fim ? fmtDataBR(p.fim) : 'em curso'}</b> · ${fmtDuracao(periodoDias(p))} · ${Number(p.valorMensal) ? fmtBRL(p.valorMensal) + '/mês' : 'sem valor informado'} · subtotal <b>${fmtBRL(periodoValor(p))}</b>${p.obs ? `<div class="lr-meta">${escapeHtml(p.obs)}</div>` : ''}</div>${canManage ? `<div class="per-actions">${btn('btn-ghost','loc-edit','Editar',p.id)}${btn('btn-danger','loc-del','Excluir',p.id)}</div>` : ''}</div>`;
  });
  const ult = ps[ps.length-1];
  if(ult && ult.fim){
    const parado = diffDias(ult.fim, hojeISO());
    if(parado > 0) locHtml += `<div class="per-row gap">⏸ Sem locação há ${fmtDuracao(parado)} (desde ${fmtDataBR(ult.fim)})</div>`;
  }
  if(!locHtml) locHtml = '<div class="hist-item">Nenhum período de locação registrado.</div>';
  let locAct = '';
  if(canManage){
    if(!rl.aberto) locAct = btn('btn-ghost', 'loc-add', ps.length ? '+ Retomar locação' : '+ Registrar locação');
    else locAct = btn('btn-ghost', 'loc-value', 'Alterar valor mensal') + btn('btn-ghost', 'loc-end', 'Encerrar locação');
  }

  /* --- manutenções --- */
  const mans = [...(i.manutencoes||[])].sort((a,b)=> (b.dataEnvio||'').localeCompare(a.dataEnvio||''));
  const manHtml = mans.length ? mans.map(m => `<div class="per-row"><div><b>${fmtDataBR(m.dataEnvio)} → ${m.dataRetorno ? fmtDataBR(m.dataRetorno) : 'em manutenção'}</b> · ${fmtDuracao(Math.max(0, diffDias(m.dataEnvio, m.dataRetorno || hojeISO())))}${m.motivo ? ' · ' + escapeHtml(m.motivo) : ''}${m.oficina ? ' · ' + escapeHtml(m.oficina) : ''} · ${Number(m.valor) ? 'custo <b>' + fmtBRL(m.valor) + '</b>' : 'custo não informado'}${m.obsRetorno ? `<div class="lr-meta">${escapeHtml(m.obsRetorno)}</div>` : ''}</div>${canManage ? `<div class="per-actions">${!m.dataRetorno ? btn('btn-ghost','man-return','Registrar retorno',m.id) : ''}${btn('btn-ghost','man-edit','Editar',m.id)}${btn('btn-danger','man-del','Excluir',m.id)}</div>` : ''}</div>`).join('')
    : '<div class="hist-item">Nenhuma manutenção registrada.</div>';
  const manAct = (canManage && !i.baixado && invDisponivel(i) > 0) ? btn('btn-ghost', 'man-add', '+ Enviar para manutenção') : '';

  /* --- eventos + ações gerais --- */
  const hist = (i.historico||[]).slice().reverse().map(h=>`<div class="hist-item">${fmtDate(h.data)}${h.por && nomesUsuarios[h.por] ? ' · ' + escapeHtml(nomesUsuarios[h.por]) : ''} — ${escapeHtml(h.detalhes||h.tipo)}</div>`).join('') || '<div class="hist-item">Sem histórico ainda.</div>';
  let actions = '';
  if(canManage){
    actions += `<button class="btn btn-ghost btn-sm" data-edit-inv="${i.id}">Editar item</button>`;
    if(!i.baixado && (i.quantidadeAlocada||0) > 0) actions += `<button class="btn btn-ghost btn-sm" data-act="liberar" data-id="${i.id}">Liberar unidade alocada (voltar p/ disponível)</button>`;
    if(i.baixado) actions += `<button class="btn btn-ghost btn-sm" data-act="reativar" data-id="${i.id}">Reativar item (voltar p/ disponível)</button>`;
    if(!i.baixado) actions += `<button class="btn btn-danger btn-sm" data-act="baixar" data-id="${i.id}">Dar baixa / descarte</button>`;
    actions += `<button class="btn btn-danger btn-sm" data-act="excluir" data-id="${i.id}">Excluir item</button>`;
  }
  return `<div class="row-detail">
    ${stats}${contrato}
    <div class="sec-title">Histórico de locação <span class="lr-meta">(valor mensal × dias ÷ 30)</span></div>${locHtml}<div class="row-actions" style="margin-bottom:12px;">${locAct}</div>
    <div class="sec-title">Manutenções</div>${manHtml}<div class="row-actions" style="margin-bottom:12px;">${manAct}</div>
    <div class="sec-title">Histórico de alterações</div>${hist}
    <div class="row-actions">${actions}</div>
  </div>`;
}

/* ---------- formulário em janela (modal) ---------- */
function askForm(title, intro, fields, okLabel, validate){
  return new Promise((resolve)=>{
    const root = document.getElementById('confirm-root');
    root.innerHTML = `<div class="confirm-overlay"><div class="confirm-box" style="max-width:460px; max-height:90vh; overflow:auto;">
      <h3>${title}</h3>${intro ? `<p>${intro}</p>` : ''}
      ${fields.map(f => `<div class="field" style="margin-bottom:10px;"><label for="af-${f.key}">${f.label}${f.required ? ' *' : ''}</label>
        <input id="af-${f.key}" type="${f.type || 'text'}" value="${escapeHtml(f.value == null ? '' : String(f.value))}" ${f.step ? `step="${f.step}"` : ''} ${f.min != null ? `min="${f.min}"` : ''} placeholder="${escapeHtml(f.placeholder || '')}"></div>`).join('')}
      <div id="af-erro" style="color:var(--ruim); font-size:12.5px; min-height:16px; margin-bottom:8px;"></div>
      <div class="confirm-actions"><button class="btn btn-ghost btn-sm" id="af-cancel">Cancelar</button><button class="btn btn-primary btn-sm" id="af-ok">${okLabel}</button></div>
    </div></div>`;
    const fechar = (v)=>{ root.innerHTML = ''; resolve(v); };
    const enviar = ()=>{
      const vals = {};
      fields.forEach(f => { vals[f.key] = document.getElementById('af-' + f.key).value.trim(); });
      const faltando = fields.find(f => f.required && !vals[f.key]);
      if(faltando){ document.getElementById('af-erro').textContent = `Preencha: ${faltando.label}.`; return; }
      const erro = validate ? validate(vals) : '';
      if(erro){ document.getElementById('af-erro').textContent = erro; return; }
      fechar(vals);
    };
    document.getElementById('af-cancel').onclick = ()=> fechar(null);
    document.getElementById('af-ok').onclick = enviar;
    root.querySelectorAll('input').forEach(inp => inp.addEventListener('keydown', e => { if(e.key === 'Enter') enviar(); }));
    const primeiro = root.querySelector('input'); if(primeiro) primeiro.focus();
  });
}

/* ---------- ações de locação e manutenção (botões do detalhe) ---------- */
const numBR = v => (v === '' || v == null) ? 0 : (Number(String(v).replace(',', '.')) || 0);
document.getElementById('inv-list').addEventListener('click', (e)=>{
  const b = e.target.closest('[data-iact]');
  if(b && canManage) invOp(b.dataset.iact, b.dataset.id, b.dataset.pid);
});
async function invOp(op, id, pid){
  const i = getInvById(id);
  if(!i) return;
  const evento = (txt)=>{ logInv(i, op, txt); };
  const salvar = (msg)=>{ persistInv(i); renderInventario(); renderStats(); showToast(msg); };
  const ps = periodosDe(i);
  const ult = ps[ps.length-1];
  const ab = periodoAberto(i);

  if(op === 'loc-add'){
    const v = await askForm(ps.length ? 'Retomar locação' : 'Registrar locação',
      ps.length ? `A última locação terminou em ${fmtDataBR(ult.fim)}. O intervalo até a nova data fica registrado como "sem locação".` : 'Informe quando a locação começou e o valor combinado.',
      [ {key:'inicio', label:'Data de início', type:'date', value:hojeISO(), required:true},
        {key:'valor', label:'Valor mensal (R$)', type:'number', step:'0.01', min:'0', placeholder:'0,00'},
        {key:'fornecedor', label:'Locadora / fornecedor', value:i.fornecedor || ''},
        {key:'vencimento', label:'Vencimento do contrato', type:'date', value:i.vencimentoLocacao || ''},
        {key:'obs', label:'Observação'} ],
      'Registrar',
      (x)=> (ult && ult.fim && x.inicio < ult.fim) ? `A nova locação não pode começar antes do fim da anterior (${fmtDataBR(ult.fim)}).`
          : (x.vencimento && x.vencimento <= x.inicio) ? 'O vencimento precisa ser depois do início.' : '');
    if(!v) return;
    i.periodos = (i.periodos||[]).concat([{id:uid('per'), inicio:v.inicio, fim:null, valorMensal:numBR(v.valor), obs:v.obs}]);
    if(v.fornecedor) i.fornecedor = v.fornecedor;
    if(v.vencimento) i.vencimentoLocacao = v.vencimento;
    evento(`Locação iniciada em ${fmtDataBR(v.inicio)}${numBR(v.valor) ? ' — ' + fmtBRL(numBR(v.valor)) + '/mês' : ''}`);
    return salvar('Locação registrada.');
  }
  if(op === 'loc-end'){
    if(!ab) return;
    const v = await askForm('Encerrar locação', `Locação em curso desde ${fmtDataBR(ab.inicio)}.`,
      [ {key:'fim', label:'Data de término', type:'date', value:hojeISO(), required:true} ], 'Encerrar',
      (x)=> x.fim < ab.inicio ? 'O término não pode ser antes do início.' : '');
    if(!v) return;
    ab.fim = v.fim;
    evento(`Locação encerrada em ${fmtDataBR(v.fim)}`);
    return salvar('Locação encerrada.');
  }
  if(op === 'loc-value'){
    if(!ab) return;
    const v = await askForm('Alterar valor mensal', `Valor atual: ${fmtBRL(ab.valorMensal)}/mês. O período atual é encerrado na data informada e um novo período começa com o novo valor, guardando o histórico.`,
      [ {key:'valor', label:'Novo valor mensal (R$)', type:'number', step:'0.01', min:'0', required:true},
        {key:'desde', label:'Vale a partir de', type:'date', value:hojeISO(), required:true} ], 'Salvar',
      (x)=> x.desde < ab.inicio ? `A data não pode ser antes do início do período atual (${fmtDataBR(ab.inicio)}).` : '');
    if(!v) return;
    const antigo = Number(ab.valorMensal) || 0;
    if(v.desde === ab.inicio){ ab.valorMensal = numBR(v.valor); }
    else { ab.fim = v.desde; i.periodos = (i.periodos||[]).concat([{id:uid('per'), inicio:v.desde, fim:null, valorMensal:numBR(v.valor), obs:''}]); }
    evento(`Valor mensal alterado de ${fmtBRL(antigo)} para ${fmtBRL(numBR(v.valor))} a partir de ${fmtDataBR(v.desde)}`);
    return salvar('Valor atualizado e histórico guardado.');
  }
  if(op === 'loc-edit'){
    const p = (i.periodos||[]).find(x => x.id === pid);
    if(!p) return;
    const v = await askForm('Editar período de locação', 'Deixe o término em branco se a locação ainda está em curso.',
      [ {key:'inicio', label:'Início', type:'date', value:p.inicio, required:true},
        {key:'fim', label:'Término', type:'date', value:p.fim || ''},
        {key:'valor', label:'Valor mensal (R$)', type:'number', step:'0.01', min:'0', value:p.valorMensal || ''},
        {key:'obs', label:'Observação', value:p.obs || ''} ], 'Salvar',
      (x)=>{
        if(x.fim && x.fim < x.inicio) return 'O término não pode ser antes do início.';
        const outros = (i.periodos||[]).filter(q => q.id !== pid);
        if(!x.fim && outros.some(q => !q.fim)) return 'Já existe outra locação em curso. Informe o término desta.';
        const sobrepoe = outros.find(q => x.inicio < (q.fim || '9999-12-31') && (x.fim || '9999-12-31') > q.inicio);
        return sobrepoe ? `Esse período se sobrepõe a outro (${fmtDataBR(sobrepoe.inicio)} → ${sobrepoe.fim ? fmtDataBR(sobrepoe.fim) : 'em curso'}).` : '';
      });
    if(!v) return;
    p.inicio = v.inicio; p.fim = v.fim || null; p.valorMensal = numBR(v.valor); p.obs = v.obs;
    evento('Período de locação editado');
    return salvar('Período atualizado.');
  }
  if(op === 'loc-del'){
    const p = (i.periodos||[]).find(x => x.id === pid);
    if(!p) return;
    const ok = await askConfirm('Excluir período de locação?', `Remove o período ${fmtDataBR(p.inicio)} → ${p.fim ? fmtDataBR(p.fim) : 'em curso'} e o valor dele deixa de entrar no acumulado.`, 'Excluir');
    if(!ok) return;
    i.periodos = i.periodos.filter(x => x.id !== pid);
    evento('Período de locação excluído');
    return salvar('Período excluído.');
  }
  if(op === 'man-add'){
    if(invDisponivel(i) <= 0){ showToast('Não há unidade disponível para enviar à manutenção.'); return; }
    const v = await askForm('Enviar para manutenção', `${escapeHtml(i.nome)} (Patr. ${escapeHtml(i.patrimonio)})`,
      [ {key:'dataEnvio', label:'Data de envio', type:'date', value:hojeISO(), required:true},
        {key:'motivo', label:'Motivo / defeito'},
        {key:'oficina', label:'Oficina / assistência'},
        {key:'valor', label:'Custo previsto (R$) — se já souber', type:'number', step:'0.01', min:'0'} ], 'Enviar');
    if(!v) return;
    i.manutencoes = (i.manutencoes||[]).concat([{id:uid('man'), dataEnvio:v.dataEnvio, dataRetorno:null, valor:v.valor ? numBR(v.valor) : null, oficina:v.oficina, motivo:v.motivo}]);
    syncContadorManut(i);
    evento(`Enviado para manutenção em ${fmtDataBR(v.dataEnvio)}${v.motivo ? ' — ' + v.motivo : ''}`);
    return salvar('Enviado para manutenção.');
  }
  const m = (i.manutencoes||[]).find(x => x.id === pid);
  if(op === 'man-return'){
    if(!m) return;
    const v = await askForm('Registrar retorno da manutenção', `Enviado em ${fmtDataBR(m.dataEnvio)}${m.motivo ? ' — ' + escapeHtml(m.motivo) : ''}.`,
      [ {key:'dataRetorno', label:'Data de retorno', type:'date', value:hojeISO(), required:true},
        {key:'valor', label:'Custo final (R$)', type:'number', step:'0.01', min:'0', value:m.valor || ''},
        {key:'obs', label:'Observação (serviço realizado)'} ], 'Registrar retorno',
      (x)=> x.dataRetorno < m.dataEnvio ? 'O retorno não pode ser antes do envio.' : '');
    if(!v) return;
    m.dataRetorno = v.dataRetorno; m.valor = v.valor ? numBR(v.valor) : null; m.obsRetorno = v.obs;
    syncContadorManut(i);
    evento(`Retornou da manutenção em ${fmtDataBR(v.dataRetorno)}${numBR(v.valor) ? ' — custo ' + fmtBRL(numBR(v.valor)) : ''}`);
    return salvar('Retorno registrado — item disponível.');
  }
  if(op === 'man-edit'){
    if(!m) return;
    const v = await askForm('Editar manutenção', 'Deixe o retorno em branco se o item ainda está na manutenção.',
      [ {key:'dataEnvio', label:'Data de envio', type:'date', value:m.dataEnvio, required:true},
        {key:'dataRetorno', label:'Data de retorno', type:'date', value:m.dataRetorno || ''},
        {key:'motivo', label:'Motivo / defeito', value:m.motivo || ''},
        {key:'oficina', label:'Oficina / assistência', value:m.oficina || ''},
        {key:'valor', label:'Custo (R$)', type:'number', step:'0.01', min:'0', value:m.valor || ''} ], 'Salvar',
      (x)=> (x.dataRetorno && x.dataRetorno < x.dataEnvio) ? 'O retorno não pode ser antes do envio.' : '');
    if(!v) return;
    m.dataEnvio = v.dataEnvio; m.dataRetorno = v.dataRetorno || null; m.motivo = v.motivo; m.oficina = v.oficina; m.valor = v.valor ? numBR(v.valor) : null;
    syncContadorManut(i);
    evento('Registro de manutenção editado');
    return salvar('Manutenção atualizada.');
  }
  if(op === 'man-del'){
    if(!m) return;
    const ok = await askConfirm('Excluir registro de manutenção?', `Remove o registro de ${fmtDataBR(m.dataEnvio)} e o custo dele deixa de entrar no acumulado.`, 'Excluir');
    if(!ok) return;
    i.manutencoes = i.manutencoes.filter(x => x.id !== pid);
    syncContadorManut(i);
    evento('Registro de manutenção excluído');
    return salvar('Registro excluído.');
  }
}
async function invAction(id, act){
  const item = getInvById(id);
  if(!item) return;
  if(act === 'baixar'){
    const ok = await askConfirm('Dar baixa neste item?', `Todo o estoque de <b>${escapeHtml(item.nome)}</b> sai de circulação, mas o histórico é mantido.`, 'Dar baixa');
    if(!ok) return;
    item.baixado = true;
    logInv(item, 'baixa', 'Baixado / descartado');
    persistInv(item); renderInventario(); renderStats();
    showToast('Item baixado.');
    return;
  }
  if(act === 'excluir'){
    const ok = await askConfirm('Excluir item do inventário?', `Isso apaga <b>${escapeHtml(item.nome)}</b> (patrimônio ${escapeHtml(item.patrimonio)}) permanentemente, incluindo o histórico. Não é possível desfazer.`, 'Excluir');
    if(!ok) return;
    inventario = inventario.filter(x => x.id !== id);
    delete localOnlyInv[id];
    expandedInvId = null;
    renderInventario();
    showToast('Item excluído.');
    await deleteInvRemote(id);
    return;
  }
  if(act === 'liberar'){
    if((item.quantidadeAlocada||0) <= 0){ showToast('Não há unidade alocada para liberar.'); return; }
    const holders = [];
    fichas.filter(x=>x.status==='aberta').forEach(fx=> fx.itens.forEach((it,ix)=>{ if(it.inventarioId===item.id && !it.trocado && !it.liberadoManual) holders.push({f:fx, ix}); }));
    if(holders.length){
      const labels = holders.map(h=>`${escapeHtml(h.f.colaborador||'Sem nome')} — ${escapeHtml(h.f.funcao)} (retirado em ${fmtDataBR(h.f.itens[h.ix].dataRetirada)})`);
      labels.push('Liberar sem alterar nenhuma ficha');
      const pick = await askChoice('Liberar unidade alocada', `Qual ficha em aberto está com 1 unidade de <b>${escapeHtml(item.nome)}</b>? Ela será registrada como <b>devolvida em condição Bom</b> nessa ficha e voltará para disponível.`, labels);
      if(pick === null) return;
      if(pick < holders.length){
        const h = holders[pick]; const it = h.f.itens[h.ix];
        it.condDevolucao = 'Bom'; it.dataDevolucao = new Date().toISOString().slice(0,10);
        it.obsDevolucao = 'Liberado manualmente pelo inventário'; it.liberadoManual = true;
        item.quantidadeAlocada -= 1;
        logInv(item, 'liberado', `1 unidade liberada manualmente — devolvida na ficha de ${h.f.colaborador||'colaborador'}`);
        await persistFichaNow(h.f);
        persistInv(item); renderInventario(); renderStats();
        showToast('Unidade liberada e ficha atualizada.');
        return;
      }
    }
    const ok = await askConfirm('Liberar unidade alocada?', `1 unidade de <b>${escapeHtml(item.nome)}</b> volta para <b>disponível</b> e poderá ser realocada. Atenção: a ficha do colaborador que estava com o item não é alterada — se ele ainda constar como responsável, registre a devolução ou a troca na ficha.`, 'Liberar');
    if(!ok) return;
    item.quantidadeAlocada -= 1;
    logInv(item, 'liberado', '1 unidade liberada manualmente (voltou para disponível)');
  } else if(act === 'reativar'){
    item.baixado = false;
    logInv(item, 'reativado', 'Item reativado (voltou para disponível)');
  } else if(act === 'manutencao'){
    if(invDisponivel(item) <= 0){ showToast('Não há unidade disponível para enviar à manutenção.'); return; }
    item.quantidadeManutencao = (item.quantidadeManutencao||0) + 1;
    logInv(item, 'manutencao', '1 unidade enviada para manutenção');
  } else if(act === 'disponivel'){
    if((item.quantidadeManutencao||0) <= 0){ showToast('Não há unidade em manutenção.'); return; }
    item.quantidadeManutencao -= 1;
    logInv(item, 'disponivel', '1 unidade retornou da manutenção');
  }
  persistInv(item);
  renderInventario();
  renderStats();
  showToast('Inventário atualizado.');
}

/* ============================================================
   COLABORADORES
   ============================================================ */
let colabEditingId = null;
document.getElementById('btn-toggle-colab-form').addEventListener('click', ()=>{
  if(!canManage) return;
  colabEditingId = null;
  document.getElementById('colab-f-nome').value=''; document.getElementById('colab-f-matricula').value='';
  document.getElementById('colab-f-funcao').value = FUNCOES_NOMES[0];
  document.getElementById('colab-f-regiao').value = '';
  document.getElementById('colab-f-cc').value = '';
  document.getElementById('colab-form-title').textContent = 'Novo colaborador';
  document.getElementById('colab-new-form').classList.toggle('hidden');
});
document.getElementById('btn-cancel-colab-form').addEventListener('click', ()=>{
  colabEditingId = null;
  document.getElementById('colab-new-form').classList.add('hidden');
});
function abrirEdicaoColab(id){
  const c = colaboradores.find(x=>x.id===id);
  if(!c || !canManage) return;
  colabEditingId = id;
  document.getElementById('colab-f-nome').value = c.nome || '';
  document.getElementById('colab-f-funcao').value = c.funcao || FUNCOES_NOMES[0];
  document.getElementById('colab-f-matricula').value = c.matricula || '';
  document.getElementById('colab-f-regiao').value = regiaoDeColab(c);
  document.getElementById('colab-f-cc').value = centroCustoDeColab(c);
  document.getElementById('colab-form-title').textContent = 'Editar colaborador';
  document.getElementById('colab-new-form').classList.remove('hidden');
  document.getElementById('colab-new-form').scrollIntoView({behavior:'smooth', block:'center'});
}
document.getElementById('btn-save-colab-form').addEventListener('click', ()=>{
  if(!canManage) return;
  const nome = document.getElementById('colab-f-nome').value.trim();
  const funcao = document.getElementById('colab-f-funcao').value;
  const matricula = document.getElementById('colab-f-matricula').value.trim();
  const regiao = document.getElementById('colab-f-regiao').value;
  const cc = document.getElementById('colab-f-cc').value.trim();
  if(!nome){ showToast('Informe o nome do colaborador.'); return; }
  if(colabEditingId){
    const c = colaboradores.find(x=>x.id===colabEditingId);
    if(c){ c.nome = nome; c.funcao = funcao; c.matricula = matricula; c.regiao = regiao; c.centroCusto = cc; persistColab(c); showToast('Colaborador atualizado.'); }
  } else {
    const c = novoColaborador();
    c.nome = nome; c.funcao = funcao; c.matricula = matricula; c.regiao = regiao; c.centroCusto = cc;
    localOnlyColab[c.id] = c; colaboradores = [c, ...colaboradores];
    persistColab(c);
    showToast('Colaborador cadastrado.');
  }
  colabEditingId = null;
  document.getElementById('colab-f-nome').value=''; document.getElementById('colab-f-matricula').value='';
  document.getElementById('colab-new-form').classList.add('hidden');
  renderColaboradores();
});
document.getElementById('colab-search-input').addEventListener('input', (e)=>{ colabSearchTerm = e.target.value; renderColaboradores(); });

function renderColabDatalist(){
  const dl = document.getElementById('colab-datalist');
  if(dl) dl.innerHTML = colaboradores.map(c=>`<option value="${escapeHtml(c.nome)}">`).join('');
}
function getColabFiltrados(){
  let items = [...colaboradores];
  if(colabSearchTerm.trim()){
    const q = colabSearchTerm.trim().toLowerCase();
    items = items.filter(c => (c.nome||'').toLowerCase().includes(q) || (c.matricula||'').toLowerCase().includes(q));
  }
  if(colabRegiao !== 'todas') items = items.filter(c => regiaoDeColab(c) === colabRegiao);
  if(colabCC !== 'todas') items = items.filter(c => centroCustoDeColab(c) === colabCC);
  items.sort((a,b)=> (a.nome||'').localeCompare(b.nome||''));
  return items;
}
function renderColaboradores(){
  const list = document.getElementById('colab-list');
  const items = getColabFiltrados();
  renderSelBar('colab', items.map(x=>x.id));
  if(items.length===0){
    list.innerHTML = `<div class="empty"><b>Nenhum colaborador encontrado</b>${canManage ? 'Cadastre o primeiro colaborador acima.' : 'Ainda não há colaboradores cadastrados.'}</div>`;
    return;
  }
  list.innerHTML = items.map(c=>{
    const expanded = expandedColabId === c.id;
    const fichasDele = fichas.filter(f => (f.colaborador||'').toLowerCase() === (c.nome||'').toLowerCase());
    return `<div class="list-row" data-id="${c.id}">
      <div class="list-row-head">
        <div>
          <div class="lr-name">${selBox('colab', c.id)}${escapeHtml(c.nome)}</div>
          <div class="lr-meta"><span>${escapeHtml(c.funcao)}</span>${regiaoDeColab(c) ? `<span class="pill ${regiaoDeColab(c)==='Vale' ? 'locado' : 'proprio'}">${regiaoDeColab(c)}</span>` : ''}${centroCustoDeColab(c) ? `<span>${escapeHtml(centroCustoDeColab(c))}</span>` : ''}${c.matricula ? `<span>Matrícula: ${escapeHtml(c.matricula)}</span>` : ''}<span>${fichasDele.length} ficha(s)</span></div>
        </div>
      </div>
      ${expanded ? renderColabDetail(c, fichasDele) : ''}
    </div>`;
  }).join('');
  list.querySelectorAll('.list-row').forEach(row=>{
    row.addEventListener('click', (e)=>{
      if(e.target.closest('button')) return;
      expandedColabId = expandedColabId === row.dataset.id ? null : row.dataset.id;
      renderColaboradores();
    });
  });
  if(canManage){
    list.querySelectorAll('[data-del-colab]').forEach(btn=>{
      btn.addEventListener('click', (e)=>{ e.stopPropagation(); excluirColaborador(btn.dataset.delColab); });
    });
    list.querySelectorAll('[data-edit-colab]').forEach(btn=>{
      btn.addEventListener('click', (e)=>{ e.stopPropagation(); abrirEdicaoColab(btn.dataset.editColab); });
    });
  }
}
function renderColabDetail(c, fichasDele){
  const histHtml = fichasDele.length
    ? fichasDele.slice(0,8).map(f=>`<div class="hist-item">${fmtDateShort(f.criadoEm)} — ${escapeHtml(f.funcao)} — ${f.status==='concluida'?'Concluída':(f.status==='aberta'?'Em aberto':'Rascunho')}</div>`).join('')
    : '<div class="hist-item">Nenhuma ficha registrada ainda para este colaborador.</div>';
  const actions = canManage ? `<div class="row-actions">
      <button class="btn btn-ghost btn-sm" data-edit-colab="${c.id}">Editar colaborador</button>
      <button class="btn btn-danger btn-sm" data-del-colab="${c.id}">Excluir colaborador</button>
    </div>` : '';
  return `<div class="row-detail"><b style="color:var(--ink);display:block;margin-bottom:4px;">Histórico de fichas</b>${histHtml}${actions}</div>`;
}
async function excluirColaborador(id){
  const c = colaboradores.find(x=>x.id===id);
  if(!c) return;
  const ok = await askConfirm('Excluir colaborador?', `Isso apaga o cadastro de <b>${escapeHtml(c.nome)}</b>. As fichas já criadas para ele continuam no histórico, mas ele sai da lista de colaboradores. Não é possível desfazer.`, 'Excluir');
  if(!ok) return;
  colaboradores = colaboradores.filter(x => x.id !== id);
  delete localOnlyColab[id];
  expandedColabId = null;
  renderColaboradores();
  renderColabDatalist();
  showToast('Colaborador excluído.');
  await deleteColabRemote(id);
}

/* ============================================================
   FICHA: RENDER
   ============================================================ */
function renderFicha(){
  const f = getFichaById(currentFichaId);
  if(!f){ showDashboard(); return; }
  document.getElementById('f-nome').value = f.colaborador || "";
  document.getElementById('f-nome').disabled = (f.status !== 'rascunho') || !canManage;
  document.getElementById('f-nome').oninput = (e)=>{ f.colaborador = e.target.value; scheduleSave(f); };
  document.getElementById('f-funcao-display').textContent = f.funcao;
  renderEsporadicaAdder(f);
  renderItems(f);
  renderSigRetirada(f);
  renderSigDevolucao(f);
  renderBottomBar(f);
  const cardDevolucao = document.getElementById('card-sig-devolucao');
  if(f.assinaturaRetirada) cardDevolucao.classList.remove('hidden'); else cardDevolucao.classList.add('hidden');
}

function renderEsporadicaAdder(f){
  const wrap = document.getElementById('esporadica-adder-wrap');
  if(!f || f.tipo !== 'esporadica' || f.assinaturaRetirada || !canManage){ wrap.innerHTML = ''; return; }
  const disponiveis = inventario.filter(i=>!i.baixado && invDisponivel(i) > 0);
  wrap.innerHTML = `<div class="esporadica-adder">
    <label>Adicionar equipamento do inventário</label>
    <select id="esp-select">
      <option value="">Selecione um item disponível...</option>
      ${invOptionsGrouped(disponiveis, '', i => invDisponivel(i) > 1 ? ` (${invDisponivel(i)} disponíveis)` : '')}
    </select>
    <button class="btn btn-primary btn-sm" id="esp-add-btn">+ Adicionar item</button>
  </div>`;
  document.getElementById('esp-add-btn').addEventListener('click', ()=>{
    const sel = document.getElementById('esp-select');
    const inv = getInvById(sel.value);
    if(!inv){ showToast('Escolha um item da lista.'); return; }
    f.itens.push({
      nome: inv.nome, patrimonio: inv.patrimonio, inventarioId: inv.id,
      condEntrega: null, dataRetirada: "", obsEntrega: "",
      condDevolucao: null, dataDevolucao: "", obsDevolucao: ""
    });
    scheduleSave(f);
    renderItems(f);
    renderEsporadicaAdder(f);
  });
}

function obsBoxHtml(role, idx, value, cond, disabled){
  if(cond !== 'Regular' && cond !== 'Ruim') return '';
  if(disabled){
    if(!value) return '';
    return `<div class="obs-readonly"><b>Observação (${cond})</b>${escapeHtml(value)}</div>`;
  }
  const cls = cond === 'Ruim' ? 'ruim' : '';
  return `<div class="obs-box ${cls}" data-obsbox="${role}" data-idx="${idx}">
    <label>Motivo da condição "${cond}"</label>
    <textarea data-role="obs-${role}" data-idx="${idx}" placeholder="Descreva o problema observado...">${escapeHtml(value||'')}</textarea>
  </div>`;
}

function renderItems(f){
  const container = document.getElementById('items-container');
  const entregaEditable = !f.assinaturaRetirada && canManage;
  const devolucaoEditable = !!f.assinaturaRetirada && !f.assinaturaDevolucao && canManage;
  const importBanner = f.importada ? '<div class="locked-note" style="margin-bottom:10px;">Ficha criada pela importação do inventário geral: estes itens já estavam com o colaborador antes do sistema, por isso a retirada não tem condição, data nem assinatura. Registre a devolução normalmente quando eles voltarem.</div>' : '';
  const prazoBanner = (f.status === 'aberta' && fichaBasePrazo(f)) ? (fichaVencida(f)
    ? `<div class="locked-note" style="margin-bottom:10px; border-color:var(--ruim); color:var(--ruim);">Prazo de ${LIMITE_DIAS_FICHA} dias vencido há ${fichaDiasPrazo(f) - LIMITE_DIAS_FICHA} dia(s). Confirme com o colaborador e ${canManage ? 'renove o prazo (botão no topo) ou registre a devolução' : 'peça ao gerenciador para renovar ou registrar a devolução'}.${(f.renovacoes||[]).length ? ' Renovada ' + f.renovacoes.length + 'x.' : ''}</div>`
    : `<div class="locked-note" style="margin-bottom:10px;">Itens com o colaborador há ${fichaDiasPrazo(f)} dia(s) · próximo aviso em ${fmtDataBR(fichaProximoAviso(f))}${f.renovadoEm ? ' (renovada em ' + fmtDataBR(f.renovadoEm) + ')' : ''}.</div>`) : '';
  const readOnlyBanner = prazoBanner + importBanner + (!canManage ? '<div class="locked-note" style="margin-bottom:10px;">Modo consulta: você pode visualizar esta ficha, mas não editá-la.</div>' : '');

  if(f.itens.length === 0){
    container.innerHTML = readOnlyBanner + `<div class="empty">Nenhum item adicionado ainda${f.tipo==='esporadica' ? ' — use o seletor acima para adicionar equipamentos do inventário.' : '.'}</div>`;
    return;
  }

  const podeTrocar = canManage && f.status === 'aberta' && f.assinaturaRetirada && !f.assinaturaDevolucao;

  container.innerHTML = readOnlyBanner + f.itens.map((item, idx) => {
    if(item.trocado || item.liberadoManual){
      return `<div class="item-card" style="opacity:.75;">
        <div class="item-head">
          <span class="item-name-input" style="padding:2px;">${escapeHtml(item.nome)} <span class="pill baixado" style="margin-left:6px;">${item.liberadoManual ? 'Liberado' : 'Trocado'}</span></span>
          <span class="item-tag">ITEM ${String(idx+1).padStart(2,'0')}</span>
        </div>
        <div class="item-body">
          <div class="locked-note">${item.liberadoManual ? 'Liberado pelo inventário' : 'Devolvido na troca'} em ${fmtDateShort(item.dataDevolucao)} — condição registrada: ${escapeHtml(item.condDevolucao||'—')}${item.obsDevolucao ? ' ('+escapeHtml(item.obsDevolucao)+')' : ''}.${item.liberadoManual ? '' : ' Substituído por um novo item nesta ficha.'}</div>
        </div>
      </div>`;
    }
    return `
    <div class="item-card">
      <div class="item-head">
        <input type="text" class="item-name-input" data-idx="${idx}" value="${escapeHtml(item.nome)}" ${!entregaEditable?'disabled':''}>
        <span class="item-tag">ITEM ${String(idx+1).padStart(2,'0')}</span>
      </div>
      <div class="item-body">
        <div class="patrimonio-row" style="display:flex; gap:12px; flex-wrap:wrap;">
          <div class="mini-field">
            <label>ID / Patrimônio</label>
            <input type="text" class="inp-patrimonio" data-idx="${idx}" value="${escapeHtml(item.patrimonio||'')}" placeholder="Nº do patrimônio" ${!entregaEditable?'disabled':''}>
          </div>
          <div class="mini-field" style="flex:1; min-width:220px; max-width:440px;">
            <label>Vínculo com inventário</label>
            ${entregaEditable
              ? `<select class="sel-inv-link" data-idx="${idx}">${invLinkOptions(f, idx)}</select>`
              : `<div class="locked-note" style="margin:0; padding-top:8px;">${(item.inventarioId && getInvById(item.inventarioId)) ? escapeHtml(getInvById(item.inventarioId).nome)+' (Patr. '+escapeHtml(getInvById(item.inventarioId).patrimonio)+')' : 'Sem vínculo com o inventário'}</div>`}
          </div>
        </div>
        <div class="subgrid">
          <div>
            <div class="subgroup-title"><span class="dot entrega"></span>Na retirada</div>
            <div class="cond-options" data-role="cond-entrega" data-idx="${idx}">
              ${['Bom','Regular','Ruim'].map(v => `<button type="button" class="cond-btn ${item.condEntrega===v?'selected':''} ${!entregaEditable?'locked':''}" data-v="${v}" ${!entregaEditable?'disabled':''}>${v}</button>`).join('')}
            </div>
            ${obsBoxHtml('entrega', idx, item.obsEntrega, item.condEntrega, !entregaEditable)}
            <div class="mini-field" style="margin-top:10px;">
              <label>Data de retirada</label>
              <input type="date" class="inp-data-retirada" data-idx="${idx}" value="${item.dataRetirada || ''}" ${!entregaEditable?'disabled':''}>
            </div>
          </div>
          <div>
            <div class="subgroup-title"><span class="dot devolucao"></span>Na devolução</div>
            <div class="cond-options" data-role="cond-devolucao" data-idx="${idx}">
              ${['Bom','Regular','Ruim'].map(v => `<button type="button" class="cond-btn ${item.condDevolucao===v?'selected':''} ${!devolucaoEditable?'locked':''}" data-v="${v}" ${!devolucaoEditable?'disabled':''}>${v}</button>`).join('')}
            </div>
            ${obsBoxHtml('devolucao', idx, item.obsDevolucao, item.condDevolucao, !devolucaoEditable)}
            <div class="mini-field" style="margin-top:10px;">
              <label>Data de devolução</label>
              <input type="date" class="inp-data-devolucao" data-idx="${idx}" value="${item.dataDevolucao || ''}" ${!devolucaoEditable?'disabled':''}>
            </div>
            ${!f.assinaturaRetirada ? '<div class="locked-note">Disponível após a assinatura de retirada.</div>' : ''}
            ${f.assinaturaDevolucao ? '<div class="locked-note">Devolução já confirmada.</div>' : ''}
          </div>
        </div>
        ${podeTrocar ? `<div class="row-actions" style="margin-top:12px; border-top:1px dashed var(--line-strong); padding-top:10px;">
          <button type="button" class="btn btn-ghost btn-sm" data-troca-idx="${idx}">🔁 Trocar este equipamento</button>
        </div>` : ''}
        ${trocaEmAndamentoIdx === idx ? renderTrocaForm(f, idx) : ''}
      </div>
    </div>`;
  }).join('');

  if(entregaEditable){
    container.querySelectorAll('.item-name-input').forEach(inp=> inp.addEventListener('input', (e)=>{ f.itens[+inp.dataset.idx].nome = e.target.value; scheduleSave(f); }));
    container.querySelectorAll('.inp-patrimonio').forEach(inp=> inp.addEventListener('input', (e)=>{ f.itens[+inp.dataset.idx].patrimonio = e.target.value; scheduleSave(f); }));
    container.querySelectorAll('.sel-inv-link').forEach(sel=> sel.addEventListener('change', (e)=>{
      const it = f.itens[+sel.dataset.idx];
      const inv = getInvById(e.target.value);
      it.inventarioId = inv ? inv.id : null;
      if(inv) it.patrimonio = inv.patrimonio;
      scheduleSave(f); renderItems(f);
    }));
    container.querySelectorAll('[data-role="cond-entrega"]').forEach(group=>{
      group.querySelectorAll('.cond-btn').forEach(btn=> btn.addEventListener('click', ()=>{
        const idx = +group.dataset.idx;
        f.itens[idx].condEntrega = btn.dataset.v;
        if(btn.dataset.v === 'Bom') f.itens[idx].obsEntrega = '';
        scheduleSave(f); renderItems(f); renderBottomBar(f);
      }));
    });
    container.querySelectorAll('[data-role="obs-entrega"]').forEach(ta=> ta.addEventListener('input', (e)=>{ f.itens[+ta.dataset.idx].obsEntrega = e.target.value; scheduleSave(f); }));
    container.querySelectorAll('.inp-data-retirada').forEach(inp=> inp.addEventListener('change', (e)=>{ f.itens[+inp.dataset.idx].dataRetirada = e.target.value; scheduleSave(f); renderBottomBar(f); }));
  }
  if(devolucaoEditable){
    container.querySelectorAll('[data-role="cond-devolucao"]').forEach(group=>{
      group.querySelectorAll('.cond-btn').forEach(btn=> btn.addEventListener('click', ()=>{
        const idx = +group.dataset.idx;
        f.itens[idx].condDevolucao = btn.dataset.v;
        if(btn.dataset.v === 'Bom') f.itens[idx].obsDevolucao = '';
        scheduleSave(f); renderItems(f); renderBottomBar(f);
      }));
    });
    container.querySelectorAll('[data-role="obs-devolucao"]').forEach(ta=> ta.addEventListener('input', (e)=>{ f.itens[+ta.dataset.idx].obsDevolucao = e.target.value; scheduleSave(f); }));
    container.querySelectorAll('.inp-data-devolucao').forEach(inp=> inp.addEventListener('change', (e)=>{ f.itens[+inp.dataset.idx].dataDevolucao = e.target.value; scheduleSave(f); renderBottomBar(f); }));
  }
  container.querySelectorAll('[data-troca-idx]').forEach(btn=>{
    btn.addEventListener('click', ()=>{
      trocaEmAndamentoIdx = trocaEmAndamentoIdx === +btn.dataset.trocaIdx ? null : +btn.dataset.trocaIdx;
      renderItems(f);
    });
  });
}

/* ------------- TROCA DE EQUIPAMENTO (ficha em aberto) ------------- */
/* Opções do seletor de vínculo: itens do inventário com unidade livre (descontando as já usadas por outros itens desta mesma ficha) */
function invLinkOptions(f, idx){
  const item = f.itens[idx];
  const usados = {};
  f.itens.forEach((it,i)=>{ if(i!==idx && it.inventarioId && !it.trocado) usados[it.inventarioId] = (usados[it.inventarioId]||0) + 1; });
  const livres = (i)=> Math.max(0, invDisponivel(i) - (usados[i.id]||0));
  const opts = inventario.filter(i => !i.baixado && (i.id === item.inventarioId || livres(i) > 0));
  return `<option value="">— sem vínculo —</option>` + invOptionsGrouped(opts, item.inventarioId, i => livres(i) > 1 ? ` (${livres(i)} livres)` : '');
}
let trocaEmAndamentoIdx = null;
function renderTrocaForm(f, idx){
  const item = f.itens[idx];
  const disponiveis = inventario.filter(i=>!i.baixado && invDisponivel(i) > 0 && i.id !== item.inventarioId);
  return `<div class="esporadica-adder" style="margin-top:10px;">
    <label style="color:var(--ruim);">Fechar este item (condição na devolução)</label>
    <div class="cond-options" id="troca-cond-devolucao" style="margin-bottom:8px;">
      ${['Bom','Regular','Ruim'].map(v=>`<button type="button" class="cond-btn" data-v="${v}">${v}</button>`).join('')}
    </div>
    <div id="troca-obs-wrap" class="hidden" style="margin-bottom:8px;">
      <textarea id="troca-obs" placeholder="Motivo da condição..."></textarea>
    </div>
    <label>Novo equipamento (do inventário)</label>
    <select id="troca-select">
      <option value="">Selecione um item disponível...</option>
      ${invOptionsGrouped(disponiveis, '', i => invDisponivel(i) > 1 ? ` (${invDisponivel(i)} disponíveis)` : '')}
      <option value="__livre__">Outro (digitar manualmente)</option>
    </select>
    <div id="troca-livre-wrap" class="hidden field-grid" style="margin-top:8px;">
      <div class="field"><label>Nome do novo item</label><input type="text" id="troca-livre-nome"></div>
      <div class="field"><label>Patrimônio</label><input type="text" id="troca-livre-patr"></div>
    </div>
    <div style="display:flex; gap:8px; justify-content:flex-end; margin-top:10px;">
      <button class="btn btn-ghost btn-sm" id="troca-cancelar">Cancelar</button>
      <button class="btn btn-primary btn-sm" id="troca-confirmar">Confirmar troca</button>
    </div>
  </div>`;
}
document.getElementById('items-container').addEventListener('click', (e)=>{
  const btn = e.target.closest('#troca-cond-devolucao .cond-btn');
  if(btn){
    btn.parentElement.querySelectorAll('.cond-btn').forEach(b=>b.classList.remove('selected'));
    btn.classList.add('selected');
    document.getElementById('troca-obs-wrap').classList.toggle('hidden', btn.dataset.v === 'Bom');
  }
});
document.getElementById('items-container').addEventListener('change', (e)=>{
  if(e.target.id === 'troca-select'){
    document.getElementById('troca-livre-wrap').classList.toggle('hidden', e.target.value !== '__livre__');
  }
});
document.getElementById('items-container').addEventListener('click', async (e)=>{
  if(e.target.id === 'troca-cancelar'){ trocaEmAndamentoIdx = null; renderItems(getFichaById(currentFichaId)); }
  if(e.target.id === 'troca-confirmar'){ await confirmarTroca(getFichaById(currentFichaId), trocaEmAndamentoIdx); }
});
async function confirmarTroca(f, idx){
  if(!f || idx === null) return;
  const item = f.itens[idx];
  const condBtn = document.querySelector('#troca-cond-devolucao .cond-btn.selected');
  if(!condBtn){ showToast('Selecione a condição de devolução deste item.'); return; }
  const cond = condBtn.dataset.v;
  const obs = document.getElementById('troca-obs').value.trim();
  if((cond === 'Regular' || cond === 'Ruim') && !obs){ showToast(`Descreva o motivo da condição "${cond}".`); return; }
  const sel = document.getElementById('troca-select').value;
  let novoNome, novoPatr, novoInvId = null;
  if(sel === '__livre__'){
    novoNome = document.getElementById('troca-livre-nome').value.trim();
    novoPatr = document.getElementById('troca-livre-patr').value.trim();
    if(!novoNome){ showToast('Informe o nome do novo equipamento.'); return; }
  } else if(sel){
    const inv = getInvById(sel);
    if(!inv){ showToast('Escolha um item válido.'); return; }
    novoNome = inv.nome; novoPatr = inv.patrimonio; novoInvId = inv.id;
  } else {
    showToast('Escolha o novo equipamento, ou "Outro" para digitar manualmente.'); return;
  }

  const hoje = new Date().toISOString().slice(0,10);
  item.condDevolucao = cond; item.obsDevolucao = obs; item.dataDevolucao = hoje; item.trocado = true;
  if(item.inventarioId){
    const invAntigo = getInvById(item.inventarioId);
    if(invAntigo){
      invAntigo.quantidadeAlocada = Math.max(0, (invAntigo.quantidadeAlocada||0) - 1);
      if(cond === 'Ruim') abrirManutencaoAuto(invAntigo, `Devolvido por troca em condição Ruim (${f.colaborador || 'colaborador'})`);
      logInv(invAntigo, 'devolucao', `Devolvido por troca — condição: ${cond}`);
      await persistInv(invAntigo);
    }
  }
  const novoItem = {
    nome: novoNome, patrimonio: novoPatr, inventarioId: novoInvId,
    condEntrega: 'Bom', dataRetirada: hoje, obsEntrega: '',
    condDevolucao: null, dataDevolucao: '', obsDevolucao: ''
  };
  f.itens.push(novoItem);
  if(novoInvId){
    const invNovo = getInvById(novoInvId);
    if(invNovo){
      invNovo.quantidadeAlocada = (invNovo.quantidadeAlocada||0) + 1;
      invNovo.alocadoA = f.colaborador || '';
      logInv(invNovo, 'alocado', `Retirado por troca — ${f.colaborador}`);
      await persistInv(invNovo);
    }
  }
  trocaEmAndamentoIdx = null;
  await persistFichaNow(f);
  renderItems(f);
  renderStats();
  showToast('Equipamento trocado com sucesso.');
}

/* ------------- ASSINATURA (canvas) ------------- */
function setupSignaturePad(canvasId){
  const canvas = document.getElementById(canvasId);
  const ctx = canvas.getContext('2d');
  function resize(){
    const rect = canvas.getBoundingClientRect();
    const ratio = window.devicePixelRatio || 1;
    canvas.width = rect.width * ratio; canvas.height = rect.height * ratio;
    ctx.scale(ratio, ratio); ctx.lineWidth = 2.2; ctx.lineCap = 'round'; ctx.strokeStyle = '#1B1B1B';
  }
  resize();
  window.addEventListener('resize', resize);
  let drawing=false,lastX=0,lastY=0;
  function pos(e){ const rect=canvas.getBoundingClientRect(); const t=e.touches?e.touches[0]:e; return {x:t.clientX-rect.left,y:t.clientY-rect.top}; }
  function start(e){ drawing=true; const p=pos(e); lastX=p.x; lastY=p.y; e.preventDefault(); }
  function move(e){ if(!drawing) return; const p=pos(e); ctx.beginPath(); ctx.moveTo(lastX,lastY); ctx.lineTo(p.x,p.y); ctx.stroke(); lastX=p.x; lastY=p.y; canvas._hasContent=true; e.preventDefault(); }
  function end(){ drawing=false; }
  canvas.addEventListener('mousedown', start);
  canvas.addEventListener('mousemove', move);
  window.addEventListener('mouseup', end);
  canvas.addEventListener('touchstart', start, {passive:false});
  canvas.addEventListener('touchmove', move, {passive:false});
  canvas.addEventListener('touchend', end);
  canvas._clear = function(){ const rect=canvas.getBoundingClientRect(); ctx.clearRect(0,0,rect.width,rect.height); canvas._hasContent=false; };
  canvas._isEmpty = function(){ return !canvas._hasContent; };
  return canvas;
}
function exportSignature(canvas){
  const rect = canvas.getBoundingClientRect();
  const maxW = 500; const scale = Math.min(1, maxW/rect.width);
  const tmp = document.createElement('canvas');
  tmp.width = Math.max(1, Math.round(rect.width*scale)); tmp.height = Math.max(1, Math.round(rect.height*scale));
  const tctx = tmp.getContext('2d');
  tctx.fillStyle = '#fff'; tctx.fillRect(0,0,tmp.width,tmp.height);
  tctx.drawImage(canvas, 0, 0, tmp.width, tmp.height);
  return tmp.toDataURL('image/png');
}

let canvasRetirada, canvasDevolucao;

function renderSigRetirada(f){
  const emptyBlock = document.getElementById('sig-retirada-empty');
  const savedBlock = document.getElementById('sig-retirada-saved');
  if(f.assinaturaRetirada){
    emptyBlock.classList.add('hidden'); savedBlock.classList.remove('hidden');
    document.getElementById('img-retirada').src = f.assinaturaRetirada;
    document.getElementById('meta-retirada').textContent = f.importada ? 'Retirada anterior ao sistema — importada do inventário geral (sem assinatura)' : 'Assinado em ' + fmtDate(f.dataAssinaturaRetirada);
    return;
  }
  savedBlock.classList.add('hidden'); emptyBlock.classList.remove('hidden');
  if(canManage){
    emptyBlock.innerHTML = `<div class="sig-wrap"><canvas class="sigpad" id="canvas-retirada"></canvas></div>
      <div class="sig-actions"><span class="sig-hint">Assine com o dedo ou o mouse na área acima.</span>
      <div style="display:flex;gap:8px;"><button class="btn btn-ghost btn-sm" id="btn-clear-retirada">Limpar</button></div></div>`;
    document.getElementById('btn-clear-retirada').addEventListener('click', ()=> canvasRetirada && canvasRetirada._clear());
    setTimeout(()=>{ canvasRetirada = setupSignaturePad('canvas-retirada'); }, 0);
  } else {
    emptyBlock.innerHTML = `<div class="locked-note">Assinatura de retirada pendente.</div>`;
  }
}
function renderSigDevolucao(f){
  const emptyBlock = document.getElementById('sig-devolucao-empty');
  const savedBlock = document.getElementById('sig-devolucao-saved');
  if(f.assinaturaDevolucao){
    emptyBlock.classList.add('hidden'); savedBlock.classList.remove('hidden');
    document.getElementById('img-devolucao').src = f.assinaturaDevolucao;
    document.getElementById('meta-devolucao').textContent = 'Assinado em ' + fmtDate(f.dataAssinaturaDevolucao);
    return;
  }
  if(!f.assinaturaRetirada) return;
  savedBlock.classList.add('hidden'); emptyBlock.classList.remove('hidden');
  if(canManage){
    emptyBlock.innerHTML = `<div class="sig-wrap"><canvas class="sigpad" id="canvas-devolucao"></canvas></div>
      <div class="sig-actions"><span class="sig-hint">Assine com o dedo ou o mouse na área acima.</span>
      <div style="display:flex;gap:8px;"><button class="btn btn-ghost btn-sm" id="btn-clear-devolucao">Limpar</button></div></div>`;
    document.getElementById('btn-clear-devolucao').addEventListener('click', ()=> canvasDevolucao && canvasDevolucao._clear());
    setTimeout(()=>{ canvasDevolucao = setupSignaturePad('canvas-devolucao'); }, 0);
  } else {
    emptyBlock.innerHTML = `<div class="locked-note">Assinatura de devolução pendente.</div>`;
  }
}

/* ------------- BARRA INFERIOR ------------- */
function renderBottomBar(f){
  const bar = document.getElementById('bottom-bar');
  bar.innerHTML = "";
  if(canManage && (f.status === 'rascunho' || (f.status === 'aberta' && !f.assinaturaRetirada))){
    const btn = document.createElement('button');
    btn.className = 'btn btn-primary'; btn.textContent = 'Confirmar retirada e assinar';
    btn.addEventListener('click', ()=> confirmarRetirada(f));
    bar.appendChild(btn);
    const del = document.createElement('button');
    del.className = 'btn btn-danger'; del.textContent = 'Excluir ficha';
    del.addEventListener('click', ()=> excluirFicha(f));
    bar.appendChild(del);
  } else if(canManage && f.assinaturaRetirada && !f.assinaturaDevolucao){
    const btn = document.createElement('button');
    btn.className = 'btn btn-primary'; btn.textContent = 'Confirmar devolução e assinar';
    btn.addEventListener('click', ()=> confirmarDevolucao(f));
    bar.appendChild(btn);
    const del = document.createElement('button');
    del.className = 'btn btn-danger'; del.textContent = 'Excluir ficha';
    del.addEventListener('click', ()=> excluirFicha(f));
    bar.appendChild(del);
  } else if(f.status === 'concluida'){
    const btn = document.createElement('button');
    btn.className = 'btn btn-ghost'; btn.textContent = 'Imprimir / gerar PDF';
    btn.addEventListener('click', ()=> window.print());
    bar.appendChild(btn);
    if(canManage){
      const del = document.createElement('button');
      del.className = 'btn btn-danger'; del.textContent = 'Excluir ficha';
      del.addEventListener('click', ()=> excluirFicha(f));
      bar.appendChild(del);
    }
  }
}

async function excluirFicha(f){
  if(!canManage) return;
  const ok = await askConfirm('Excluir ficha?', `Esta ação vai apagar permanentemente a ficha de <b>${escapeHtml(f.colaborador) || 'colaborador não informado'}</b>. Não é possível desfazer.`, 'Excluir');
  if(!ok) return;
  fichas = fichas.filter(x => x.id !== f.id);
  delete localOnlyFichas[f.id];
  showToast('Ficha excluída.');
  showDashboard();
  await deleteFichaRemote(f.id);
}

async function syncInventarioNaRetirada(f){
  for(const item of f.itens){
    if(item.inventarioId && !item.trocado){
      const inv = getInvById(item.inventarioId);
      if(inv){
        inv.quantidadeAlocada = (inv.quantidadeAlocada||0) + 1;
        inv.alocadoA = f.colaborador || '';
        logInv(inv, 'alocado', `1 unidade retirada por ${f.colaborador}`);
        await persistInv(inv);
      }
    }
  }
}
async function syncInventarioNaDevolucao(f){
  for(const item of f.itens){
    if(item.inventarioId && !item.trocado && !item.liberadoManual){
      const inv = getInvById(item.inventarioId);
      if(inv){
        inv.quantidadeAlocada = Math.max(0, (inv.quantidadeAlocada||0) - 1);
        if(item.condDevolucao === 'Ruim') abrirManutencaoAuto(inv, `Devolvido em condição Ruim por ${f.colaborador || 'colaborador'}`);
        logInv(inv, 'devolucao', `Devolvido por ${f.colaborador} — condição: ${item.condDevolucao}`);
        await persistInv(inv);
      }
    }
  }
}

async function confirmarRetirada(f){
  if(!canManage) return;
  if(!f.colaborador || !f.colaborador.trim()){ showToast('Informe o nome do colaborador.'); document.getElementById('f-nome').focus(); return; }
  if(f.itens.length === 0){ showToast('Adicione pelo menos um equipamento antes de confirmar.'); return; }
  for(const item of f.itens){
    if(!item.condEntrega || !item.dataRetirada){ showToast(`Preencha condição e data de retirada de todos os itens (faltando: ${item.nome}).`); return; }
    if((item.condEntrega === 'Regular' || item.condEntrega === 'Ruim') && !item.obsEntrega.trim()){ showToast(`Descreva o motivo da condição "${item.condEntrega}" em: ${item.nome}.`); return; }
  }
  if(!canvasRetirada || canvasRetirada._isEmpty()){ showToast('Colete a assinatura do colaborador antes de confirmar.'); return; }
  f.assinaturaRetirada = exportSignature(canvasRetirada);
  f.dataAssinaturaRetirada = new Date().toISOString();
  f.status = 'aberta';
  await persistFichaNow(f);
  await syncInventarioNaRetirada(f);
  const semVinculo = f.itens.filter(it => !it.inventarioId && !it.trocado).length;
  showToast(semVinculo ? `Retirada registrada. ${semVinculo} item(ns) sem vínculo com o inventário não alteram o estoque.` : 'Retirada registrada com sucesso.');
  renderFicha();
}
async function confirmarDevolucao(f){
  if(!canManage) return;
  for(const item of f.itens){
    if(!item.condDevolucao || !item.dataDevolucao){ showToast(`Preencha condição e data de devolução de todos os itens (faltando: ${item.nome}).`); return; }
    if((item.condDevolucao === 'Regular' || item.condDevolucao === 'Ruim') && !item.obsDevolucao.trim()){ showToast(`Descreva o motivo da condição "${item.condDevolucao}" em: ${item.nome}.`); return; }
  }
  if(!canvasDevolucao || canvasDevolucao._isEmpty()){ showToast('Colete a assinatura do colaborador antes de confirmar.'); return; }
  f.assinaturaDevolucao = exportSignature(canvasDevolucao);
  f.dataAssinaturaDevolucao = new Date().toISOString();
  f.status = 'concluida';
  await persistFichaNow(f);
  await syncInventarioNaDevolucao(f);
  showToast('Devolução registrada com sucesso.');
  renderFicha();
}

/* ============================================================
   INÍCIO
   ============================================================ */
initApp();

