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
  const LS = 'bigode:sheets:sessao';
  const cache = { fichas:{}, inventario:{}, colaboradores:{} };
  const listeners = {};
  const NOME_PERFIL = { consulta:'Visualização', responsavel:'Gerenciador', admin:'Administrador' };
  let sessao = null, lastV = -1, pendentes = 0, fila = Promise.resolve();
  const podeEscrever = () => !!sessao && (sessao.perfil === 'responsavel' || sessao.perfil === 'admin');

  /* ---------- chamadas ao Apps Script ---------- */
  async function api(acao, extra){
    const r = await fetch(URL_API, {
      method:'POST',
      headers:{'Content-Type':'text/plain;charset=utf-8'},   // evita "preflight" CORS
      body: JSON.stringify(Object.assign({acao, token: sessao && sessao.token}, extra||{}))
    });
    const j = await r.json();
    if(!j.ok){
      const e = new Error(j.erro || 'erro'); e.codigo = j.erro;
      if(j.erro === 'sessao' && sessao){ sairComAviso('Sua sessão expirou ou foi encerrada. Entre novamente.'); }
      throw e;
    }
    return j;
  }
  function guardar(){ try{ localStorage.setItem(LS, JSON.stringify(sessao)); }catch(e){} }
  function sairComAviso(msg){ try{ localStorage.removeItem(LS); sessionStorage.setItem('bigode:aviso', msg); }catch(e){} location.reload(); }

  const MSG_ERRO = {
    credenciais:'Matrícula ou senha incorretas.', bloqueado:'Muitas tentativas erradas. Aguarde 10 minutos e tente de novo.',
    inativo:'Este usuário está desativado. Fale com o administrador.', senha_fraca:'A senha precisa ter 8 ou mais caracteres, com letras e números.',
    senha_atual:'A senha atual não confere.', senha_igual:'A nova senha deve ser diferente da atual.', matricula_formato:'Matrícula no formato F0000000.',
    ultimo_admin:'Precisa existir pelo menos um administrador ativo.', proprio:'Você não pode excluir o seu próprio usuário.',
    somente_admin:'Só o administrador pode fazer isso.', somente_leitura:'Seu perfil é somente leitura.', nome:'Informe o nome.', perfil:'Escolha o perfil.',
    sem_usuario:'Crie primeiro o seu usuário na tela Usuários.', nao_encontrado:'Usuário não encontrado.'
  };
  const msgErro = (e)=> MSG_ERRO[e && e.codigo] || 'Não foi possível conectar à planilha. Verifique a internet e o link do Apps Script.';

  /* ---------- telas: entrada (duas colunas) e troca de senha ---------- */
  const SVG_USER = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><circle cx="12" cy="8" r="4"/><path d="M4 21c.8-4 4-6 8-6s7.2 2 8 6"/></svg>';
  const SVG_LOCK = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><rect x="4" y="11" width="16" height="10" rx="2"/><path d="M8 11V8a4 4 0 0 1 8 0v3"/></svg>';
  const olho = '<button type="button" class="pw-eye" aria-label="Mostrar senha" title="Mostrar/ocultar senha">👁</button>';
  function ligarOlhos(raiz){
    raiz.querySelectorAll('.pw-eye').forEach(b=> b.addEventListener('click', ()=>{ const i = b.parentElement.querySelector('input'); i.type = i.type === 'password' ? 'text' : 'password'; }));
  }
  function telaLogin(msg){
    return new Promise(resolve=>{
      const ov = document.createElement('div'); ov.className = 'login';
      const barras = [38,52,44,68,57,74,63,82].map((h,i)=>`<span style="height:${h}%;animation-delay:${400+i*70}ms"></span>`).join('');
      ov.innerHTML = `
        <aside class="login-brand">
          <div class="blob b1"></div><div class="blob b2"></div><div class="dots"></div>
          <img class="logo" src="assets/logo-unigas.jpg" alt="Unigás">
          <div>
            <div class="eyebrow">B.I.G.O.D.E</div>
            <h2>Cada ferramenta no lugar, cada responsável identificado.</h2>
            <p class="lead">Controle fichas de retirada e devolução, inventário, locações e manutenções em um só lugar, com histórico de tudo.</p>
            <div class="mini-cards">
              <div class="mini-card"><div class="t">Devoluções em dia</div><div class="bars">${barras}</div></div>
              <div class="mini-stack">
                <div class="mini-card"><div class="t">Ferramentas controladas</div><div class="n">100%</div></div>
                <div class="mini-alert"><span class="ping"><i></i><i></i></span>Alertas de vencimento</div>
              </div>
            </div>
          </div>
          <div class="copy">© ${new Date().getFullYear()} Unigás – Sistemas de Aquecimento</div>
        </aside>
        <main class="login-form-side">
          <div style="width:100%;max-width:420px;">
            <form class="login-card" novalidate>
              <img class="logo-m" src="assets/logo-unigas.jpg" alt="Unigás">
              <h1>Entrar</h1>
              <p class="sub">Bem-vindo(a) de volta. Use sua matrícula e senha.</p>
              <div class="field"><label for="lg-mat">Matrícula</label>
                <div class="ic">${SVG_USER}<input id="lg-mat" maxlength="9" placeholder="F0000000" autocapitalize="characters" autocomplete="username" autofocus></div></div>
              <div class="field"><label for="lg-senha">Senha</label>
                <div class="ic pw-wrap">${SVG_LOCK}<input id="lg-senha" type="password" autocomplete="current-password">${olho}</div></div>
              <div id="lg-erro" class="erro" style="${msg ? '' : 'display:none'}">${msg || ''}</div>
              <button class="go" type="submit">Entrar</button>
              <button class="esq" type="button" id="lg-esq">Esqueci minha senha</button>
            </form>
            <p class="login-note">B.I.G.O.D.E · Unigás</p>
          </div>
        </main>`;
      document.body.appendChild(ov);
      ligarOlhos(ov);
      const card = ov.querySelector('.login-card'), erro = ov.querySelector('#lg-erro'), go = ov.querySelector('.go');
      const mostrar = (t)=>{ erro.style.display = t ? '' : 'none'; erro.textContent = t || ''; if(t){ card.classList.remove('tremer'); void card.offsetWidth; card.classList.add('tremer'); } };
      try{ const old = localStorage.getItem('bigode:ultima-mat'); if(old) ov.querySelector('#lg-mat').value = old; }catch(e){}
      ov.querySelector('#lg-esq').addEventListener('click', ()=> mostrar('Para redefinir a senha, procure o administrador do sistema — ele define uma senha provisória para você.'));
      card.addEventListener('submit', async ev=>{
        ev.preventDefault();
        const matricula = ov.querySelector('#lg-mat').value.replace(/\s/g,'').toUpperCase();
        const senha = ov.querySelector('#lg-senha').value;
        if(!/^F\d{7,8}$/.test(matricula)){ mostrar('Digite a matrícula no formato F0000000 (letra F + 7 números).'); return; }
        if(!senha){ mostrar('Digite a senha.'); return; }
        go.disabled = true; go.textContent = 'Entrando…'; mostrar('');
        try{
          const r = await fetch(URL_API, { method:'POST', headers:{'Content-Type':'text/plain;charset=utf-8'}, body: JSON.stringify({acao:'login', matricula, senha}) }).then(x=>x.json());
          if(!r.ok){ const e = new Error(r.erro); e.codigo = r.erro; throw e; }
          try{ localStorage.setItem('bigode:ultima-mat', matricula); }catch(e){}
          go.classList.add('ok'); go.textContent = '✓ Bem-vindo(a)!';
          setTimeout(()=>{ ov.remove(); resolve({ r, senha }); }, 700);
        }catch(e){
          go.disabled = false; go.textContent = 'Entrar'; mostrar(msgErro(e));
        }
      });
    });
  }

  /* Troca de senha obrigatória (primeiro acesso ou senha redefinida pelo administrador) */
  function telaTrocaObrigatoria(senhaAtual){
    return new Promise(resolve=>{
      const ov = document.createElement('div'); ov.className = 'login';
      ov.style.gridTemplateColumns = '1fr';
      ov.innerHTML = `<main class="login-form-side"><form class="login-card" novalidate>
          <img class="logo-m" style="display:block" src="assets/logo-unigas.jpg" alt="Unigás">
          <h1>Crie sua senha</h1>
          <p class="sub">Por segurança, troque a senha provisória por uma só sua.</p>
          <div class="field"><label>Nova senha</label><div class="ic pw-wrap">${SVG_LOCK}<input id="tn-nova" type="password" autocomplete="new-password">${olho}</div></div>
          <div class="field"><label>Repita a nova senha</label><div class="ic pw-wrap">${SVG_LOCK}<input id="tn-conf" type="password" autocomplete="new-password">${olho}</div></div>
          <div class="hint" style="margin:-4px 0 12px;">Mínimo de 8 caracteres, com letras e números.</div>
          <div id="tn-erro" class="erro" style="display:none"></div>
          <button class="go" type="submit">Salvar e entrar</button>
        </form></main>`;
      document.body.appendChild(ov); ligarOlhos(ov);
      const erro = ov.querySelector('#tn-erro'), go = ov.querySelector('.go');
      const mostrar = t=>{ erro.style.display = t ? '' : 'none'; erro.textContent = t || ''; };
      ov.querySelector('form').addEventListener('submit', async ev=>{
        ev.preventDefault();
        const nova = ov.querySelector('#tn-nova').value, conf = ov.querySelector('#tn-conf').value;
        if(nova !== conf){ mostrar('As duas senhas não são iguais.'); return; }
        go.disabled = true; mostrar('');
        try{ await trocarSenha(senhaAtual, nova); ov.remove(); resolve(); }
        catch(e){ go.disabled = false; mostrar(msgErro(e)); }
      });
    });
  }

  async function trocarSenha(atual, nova){
    const r = await api('conta.senha', { atual, nova });
    sessao.token = r.token; sessao.trocar = false; guardar();
    return true;
  }

  async function entrar(){
    let aviso = '';
    try{ aviso = sessionStorage.getItem('bigode:aviso') || ''; sessionStorage.removeItem('bigode:aviso'); }catch(e){}
    try{ sessao = JSON.parse(localStorage.getItem(LS) || 'null'); }catch(e){ sessao = null; }
    if(sessao && sessao.token){
      try{
        const r = await api('auth');
        sessao = { token:r.token, matricula:r.matricula, nome:r.nome, perfil:r.perfil, boot:r.boot, trocar:false }; lastV = r.v; guardar();
        return;
      }catch(e){
        if(e.codigo !== 'sessao'){ aviso = msgErro(e); }
        sessao = null; try{ localStorage.removeItem(LS); }catch(_){}
      }
    }
    const { r, senha } = await telaLogin(aviso);
    sessao = { token:r.token, matricula:r.matricula, nome:r.nome, perfil:r.perfil, boot:r.boot, trocar:r.trocar }; lastV = r.v; guardar();
    if(r.trocar) await telaTrocaObrigatoria(senha);
  }

  /* API de contas usada pela aplicação (menu, tela Usuários e Minha conta) */
  window.bigodeAuth = {
    get perfil(){ return sessao && sessao.perfil; }, get nome(){ return sessao && sessao.nome; },
    get matricula(){ return sessao && sessao.matricula; }, get boot(){ return !!(sessao && sessao.boot); },
    nomePerfil: p => NOME_PERFIL[p || (sessao && sessao.perfil)] || '',
    msgErro,
    trocarSenha,
    listarUsuarios: async ()=> (await api('usuarios.listar')).usuarios,
    salvarUsuario: async (u)=>{
      const r = await api('usuarios.salvar', u);
      if(r.token && sessao){ sessao.token = r.token; sessao.boot = false; sessao.perfil = 'admin'; guardar(); }
      return r.usuario;
    },
    removerUsuario: async (matricula)=> { await api('usuarios.remover', { matricula }); },
    sair(){ try{ localStorage.removeItem(LS); }catch(e){} location.reload(); }
  };
  window.bigodeSair = ()=> window.bigodeAuth.sair();

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
        async setMany(lista){
          if(!podeEscrever()) throw new Error('somente leitura');
          const docs = lista.map(x => ({ id: x.id, doc: JSON.parse(JSON.stringify(x.doc)) }));
          docs.forEach(x => { cache[col][x.id] = x.doc; });
          setTimeout(()=> notify(col), 0);
          await gravarRemoto('setmany', { col, docs });
        },
        doc(id){
          return {
            async set(body){
              if(!podeEscrever()) throw new Error('somente leitura');
              const doc = JSON.parse(JSON.stringify(body));
              cache[col][id] = doc; setTimeout(()=> notify(col), 0);
              await gravarRemoto('set', { col, id, doc });
            },
            async delete(){
              if(!podeEscrever()) throw new Error('somente leitura');
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
    async canEdit(){ return podeEscrever(); },
    async isOwner(){ return podeEscrever(); },
    async isAdmin(){ return !!sessao && sessao.perfil === 'admin'; },
    async id(){ return (sessao && sessao.nome) || 'usuario'; },
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
    await carregarTudo();
    // Primeira vez com a planilha vazia: envia os 258 colaboradores iniciais
    const seed = (window.BIGODE_SEED || {}).colaboradores;
    if(podeEscrever() && seed && !Object.keys(cache.colaboradores).length){
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

})();
