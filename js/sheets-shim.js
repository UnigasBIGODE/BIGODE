/* ============================================================
   CONEXÃO COM O GOOGLE PLANILHAS
   Implementa as mesmas 3 funções que o sistema usa (db, user, downloads),
   guardando os dados na planilha por meio do Apps Script (apps-script/Codigo.gs).
   Só ativa se window.BIGODE_SHEETS_URL estiver preenchida e window.claude não existir.
   ============================================================ */
(function(){
  const URL_API = window.BIGODE_SHEETS_URL;
  if(!URL_API || (window.claude && typeof window.claude.use === 'function')) return;

  const COLS = ['fichas','inventario','colaboradores'];
  const LS = 'bigode:sheets:cred';
  const cache = { fichas:{}, inventario:{}, colaboradores:{} };
  const listeners = {};
  let cred = null, papel = 'consulta', lastV = -1, pendentes = 0, fila = Promise.resolve();

  /* ---------- chamadas ao Apps Script ---------- */
  async function api(acao, extra){
    const r = await fetch(URL_API, {
      method:'POST',
      headers:{'Content-Type':'text/plain;charset=utf-8'},   // evita "preflight" CORS
      body: JSON.stringify(Object.assign({acao, codigo: cred && cred.codigo, matricula: cred && cred.matricula}, extra||{}))
    });
    const j = await r.json();
    if(!j.ok){ const e = new Error(j.erro || 'erro'); e.codigo = j.erro; throw e; }
    return j;
  }

  /* ---------- tela de entrada: Visualização ou Gerenciador ---------- */
  const NOME_PAPEL = { consulta: 'Visualização', responsavel: 'Gerenciador' };
  function telaLogin(msg, modoInicial){
    return new Promise(resolve=>{
      let modo = modoInicial || 'consulta';
      const ov = document.createElement('div');
      ov.style.cssText = 'position:fixed;inset:0;z-index:1000;background:#fff;display:flex;align-items:center;justify-content:center;padding:16px;font-family:Inter,system-ui,sans-serif;overflow:auto;';
      ov.innerHTML = `
        <form style="width:100%;max-width:400px;border:1px solid #E5E5E5;border-radius:14px;padding:26px 22px;box-shadow:0 8px 30px rgba(0,0,0,.08);">
          <img src="assets/logo-unigas.jpg" alt="Unigás" style="height:46px;display:block;margin:0 auto 14px;">
          <div style="text-align:center;font-weight:800;font-size:18px;color:#E30613;">B.I.G.O.D.E</div>
          <div style="text-align:center;font-size:12px;color:#666;margin:2px 0 16px;">Base Integrada de Gestão Operacional De Equipamentos</div>
          <div style="font-size:12px;font-weight:600;color:#333;margin-bottom:6px;">Como você vai entrar?</div>
          <div id="bg-modos" style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:14px;">
            <button type="button" data-m="consulta" style="text-align:left;padding:10px;border-radius:10px;cursor:pointer;font-family:inherit;">
              <div style="font-weight:700;font-size:13px;">Visualização</div>
              <div style="font-size:11px;margin-top:2px;opacity:.8;">Só consulta e exporta Excel</div></button>
            <button type="button" data-m="responsavel" style="text-align:left;padding:10px;border-radius:10px;cursor:pointer;font-family:inherit;">
              <div style="font-weight:700;font-size:13px;">Gerenciador</div>
              <div style="font-size:11px;margin-top:2px;opacity:.8;">Cadastra, edita e exclui tudo</div></button>
          </div>
          <label style="font-size:12px;font-weight:600;color:#333;">Matrícula</label>
          <input id="bg-mat" required maxlength="8" placeholder="F0000000" autocapitalize="characters" autocomplete="username" style="width:100%;box-sizing:border-box;padding:10px;border:1px solid #ccc;border-radius:8px;margin:4px 0 12px;font-size:14px;">
          <label style="font-size:12px;font-weight:600;color:#333;">Código de acesso <span id="bg-modo-txt"></span></label>
          <input id="bg-cod" type="password" required style="width:100%;box-sizing:border-box;padding:10px;border:1px solid #ccc;border-radius:8px;margin:4px 0 8px;font-size:14px;">
          <div id="bg-msg" style="min-height:18px;font-size:12px;color:#B00020;margin-bottom:8px;">${msg||''}</div>
          <button style="width:100%;padding:11px;border:0;border-radius:8px;background:#E30613;color:#fff;font-weight:700;font-size:14px;cursor:pointer;">Entrar</button>
        </form>`;
      document.body.appendChild(ov);
      const pintar = ()=>{
        ov.querySelectorAll('#bg-modos button').forEach(bt=>{
          const on = bt.dataset.m === modo;
          bt.style.border = on ? '2px solid #E30613' : '1px solid #ccc';
          bt.style.background = on ? '#FFF1F2' : '#fff';
          bt.style.color = '#1B1B1B';
        });
        ov.querySelector('#bg-modo-txt').textContent = '(' + NOME_PAPEL[modo] + ')';
      };
      pintar();
      ov.querySelectorAll('#bg-modos button').forEach(bt=> bt.addEventListener('click', ()=>{ modo = bt.dataset.m; pintar(); }));
      try{ const old = JSON.parse(localStorage.getItem(LS)||'null'); if(old) ov.querySelector('#bg-mat').value = old.matricula || ''; }catch(e){}
      ov.querySelector('form').addEventListener('submit', ev=>{
        ev.preventDefault();
        const matricula = ov.querySelector('#bg-mat').value.replace(/\s/g, '').toUpperCase();
        const codigo = ov.querySelector('#bg-cod').value.trim();
        if(!/^F\d{7}$/.test(matricula)){ ov.querySelector('#bg-msg').textContent = 'Digite a matrícula no formato F0000000 (letra F + 7 números).'; return; }
        ov.remove(); resolve({matricula, codigo, modo});
      });
    });
  }

  async function entrar(){
    let msg = '', modoAnterior = null;
    try{ cred = JSON.parse(localStorage.getItem(LS)||'null'); }catch(e){ cred = null; }
    for(;;){
      if(cred && cred.codigo && cred.matricula){
        try{
          const r = await api('auth');
          if(cred.modo && r.papel !== cred.modo){
            msg = 'Este código é de ' + NOME_PAPEL[r.papel] + ', não de ' + NOME_PAPEL[cred.modo] + '. Escolha a opção correta.';
            modoAnterior = r.papel; throw Object.assign(new Error('modo'), {codigo:'modo'});
          }
          papel = r.papel; lastV = r.v; cred.nome = r.nome || cred.matricula; localStorage.setItem(LS, JSON.stringify(cred)); return true;
        }catch(e){
          if(e.codigo === 'codigo') msg = 'Código incorreto.';
          else if(e.codigo === 'matricula') msg = 'Matrícula não encontrada. Confira o formato F0000000 e se você está na lista de colaboradores.';
          else if(e.codigo !== 'modo') msg = 'Não foi possível conectar à planilha. Verifique a internet e o link do Apps Script.';
          if(!modoAnterior && cred) modoAnterior = cred.modo;
          try{ localStorage.removeItem(LS); }catch(_){}
        }
      }
      cred = await telaLogin(msg, modoAnterior);
    }
  }

  /* Botão fixo "Sair" com o nome e o perfil de quem está usando */
  function chipUsuario(){
    const c = document.createElement('div');
    c.style.cssText = 'position:fixed;right:8px;bottom:30px;z-index:101;background:#fff;border:1px solid #E5E5E5;border-radius:999px;box-shadow:0 2px 8px rgba(0,0,0,.12);padding:5px 6px 5px 12px;font:600 11.5px Inter,system-ui,sans-serif;color:#333;display:flex;gap:8px;align-items:center;';
    const t = document.createElement('span'); t.textContent = cred.nome + ' · ' + NOME_PAPEL[papel];
    const b = document.createElement('button'); b.textContent = 'Sair';
    b.style.cssText = 'border:0;border-radius:999px;background:#E30613;color:#fff;font:700 11px Inter,system-ui,sans-serif;padding:4px 10px;cursor:pointer;';
    b.addEventListener('click', ()=> window.bigodeSair());
    c.append(t, b); document.body.appendChild(c);
  }

  /* ---------- cache e avisos às telas ---------- */
  function snapshot(col, order){
    const docs = Object.keys(cache[col]).map(id => { const copia = JSON.parse(JSON.stringify(cache[col][id])); return { id, data: ()=> copia }; });
    if(order){
      docs.sort((a,b)=>{
        const x = a.data()[order.field], y = b.data()[order.field];
        const c = String(x == null ? '' : x).localeCompare(String(y == null ? '' : y));
        return order.dir === 'desc' ? -c : c;
      });
    }
    return { docs, size: docs.length };
  }
  function notify(col){ (listeners[col] || []).forEach(l => l.cb(snapshot(col, l.order))); }
  function notifyAll(){ COLS.forEach(notify); }

  async function carregarTudo(){
    const r = await api('tudo');
    COLS.forEach(c => { cache[c] = r.dados[c] || {}; });
    lastV = r.v;
    notifyAll();
  }

  /* ---------- gravações (em fila, uma de cada vez) ---------- */
  function avisoFaixa(texto, erro){
    let b = document.getElementById('bg-faixa');
    if(!b){
      b = document.createElement('div'); b.id = 'bg-faixa';
      b.style.cssText = 'position:fixed;left:0;right:0;bottom:0;z-index:100;font:600 11.5px Inter,sans-serif;text-align:center;padding:5px 8px;pointer-events:none;';
      document.body.appendChild(b);
    }
    b.textContent = texto;
    b.style.background = erro ? '#B00020' : '#1B1B1B'; b.style.color = '#fff';
    b.style.opacity = texto ? '.92' : '0';
  }
  function gravarRemoto(acao, extra){
    pendentes++; avisoFaixa('Salvando na planilha…');
    const p = fila.then(async ()=>{
      const r = await api(acao, extra);
      if(r.v === lastV + 1) lastV = r.v;     // ninguém mais gravou nesse meio-tempo
    });
    fila = p.catch(()=>{});
    return p.then(()=>{ pendentes--; avisoFaixa(pendentes ? 'Salvando na planilha…' : ''); },
                  async (e)=>{
                    pendentes--;
                    avisoFaixa('Falha ao salvar na planilha — dados recarregados', true);
                    setTimeout(()=> avisoFaixa(''), 4000);
                    try{ await carregarTudo(); }catch(_){}
                    throw e;
                  });
  }

  const db = {
    collection(col){
      const query = {
        order: null,
        orderBy(field, dir){ const q = Object.create(query); q.order = { field, dir: dir || 'asc' }; return q; },
        onSnapshot(cb){
          const l = { cb, order: this.order };
          (listeners[col] = listeners[col] || []).push(l);
          setTimeout(()=> cb(snapshot(col, l.order)), 0);
          return ()=>{ listeners[col] = listeners[col].filter(x => x !== l); };
        },
        doc(id){
          return {
            async set(body){
              if(papel !== 'responsavel') throw new Error('somente leitura');
              const doc = JSON.parse(JSON.stringify(body));
              cache[col][id] = doc; setTimeout(()=> notify(col), 0);
              await gravarRemoto('set', { col, id, doc });
            },
            async delete(){
              if(papel !== 'responsavel') throw new Error('somente leitura');
              delete cache[col][id]; setTimeout(()=> notify(col), 0);
              await gravarRemoto('del', { col, id });
            }
          };
        }
      };
      return query;
    }
  };

  const user = {
    async canEdit(){ return papel === 'responsavel'; },
    async isOwner(){ return papel === 'responsavel'; },
    async id(){ return (cred && cred.nome) || 'usuario'; },
    async profiles(ids){ const out = {}; [].concat(ids).forEach(id => { out[id] = { id, name: id }; }); return out; }
  };

  const downloads = {
    async save({ filename, data }){
      const blob = (data instanceof Blob) ? data : new Blob([data]);
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = filename;
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(()=> URL.revokeObjectURL(url), 1500);
    }
  };

  /* ---------- inicialização (feita na primeira chamada de claude.use) ---------- */
  let pronto = null;
  async function iniciar(){
    await entrar();
    chipUsuario();
    await carregarTudo();
    // Primeira vez com a planilha vazia: envia os 258 colaboradores iniciais
    const seed = (window.BIGODE_SEED || {}).colaboradores;
    if(papel === 'responsavel' && seed && !Object.keys(cache.colaboradores).length){
      const docs = seed.map(d => { const c = Object.assign({}, d); delete c.id; return { id: d.id, doc: c }; });
      docs.forEach(x => { cache.colaboradores[x.id] = x.doc; });
      notify('colaboradores');
      try{ await gravarRemoto('setmany', { col:'colaboradores', docs }); }catch(e){ console.error(e); }
    }
    // Atualização: a cada 10 s, vê se alguém gravou algo novo
    setInterval(async ()=>{
      if(document.hidden || pendentes) return;
      try{ const r = await api('versao'); if(r.v !== lastV) await carregarTudo(); }catch(e){ /* tenta de novo no próximo ciclo */ }
    }, 10000);
    document.addEventListener('visibilitychange', ()=>{ if(!document.hidden && !pendentes) api('versao').then(r=>{ if(r.v !== lastV) return carregarTudo(); }).catch(()=>{}); });
  }

  window.claude = { async use(name){
    if(!pronto) pronto = iniciar();
    await pronto;
    return ({ db, user, downloads })[name] || null;
  }};

  window.bigodeSair = function(){ try{ localStorage.removeItem(LS); }catch(e){} location.reload(); };
})();
