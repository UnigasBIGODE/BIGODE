/* ============================================================
   MODO LOCAL
   Dentro do Claude, a página usa o banco compartilhado, as permissões e o
   download de arquivos da plataforma (window.claude). Fora dele (VS Code,
   duplo clique no index.html, outro servidor) esse objeto não existe.

   Este arquivo só entra em ação quando window.claude NÃO existe e imita
   as 3 funções que o sistema usa:
     - db         → grava no localStorage do navegador (dados só neste PC)
     - user       → você é sempre "responsável" (acesso total)
     - downloads  → baixa o arquivo direto pelo navegador

   Para ligar um backend de verdade (Firebase, Supabase etc.), substitua
   este arquivo por um que implemente as mesmas funções.
   ============================================================ */
(function(){
  if(window.claude && typeof window.claude.use === 'function') return;

  const PREFIX = 'bigode:db:';
  const listeners = {};   // coleção -> [{cb, order}]

  const read  = (col)=>{ try{ return JSON.parse(localStorage.getItem(PREFIX + col)); }catch(e){ return null; } };
  const write = (col, obj)=> localStorage.setItem(PREFIX + col, JSON.stringify(obj));

  // Primeira abertura: carrega os dados iniciais (colaboradores da planilha)
  const seed = window.BIGODE_SEED || {};
  Object.keys(seed).forEach(col=>{
    if(read(col) === null){
      const obj = {};
      seed[col].forEach(d=>{ const copy = Object.assign({}, d); delete copy.id; obj[d.id] = copy; });
      write(col, obj);
    }
  });

  function snapshot(col, order){
    const obj = read(col) || {};
    // como no banco real, os dados entregues são congelados (não podem ser alterados no lugar)
    const congela = (o)=>{ Object.values(o).forEach(v => { if(v && typeof v === 'object') congela(v); }); return Object.freeze(o); };
    const docs = Object.keys(obj).map(id => { const copia = congela(JSON.parse(JSON.stringify(obj[id]))); return { id, data: ()=> copia }; });
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
            async set(body){ const o = read(col) || {}; o[id] = JSON.parse(JSON.stringify(body)); write(col, o); setTimeout(()=> notify(col), 0); },
            async delete(){ const o = read(col) || {}; delete o[id]; write(col, o); setTimeout(()=> notify(col), 0); }
          };
        }
      };
      return query;
    }
  };

  // (os avisos às telas são assíncronos, como no banco real)
  // Mantém várias abas do navegador sincronizadas
  window.addEventListener('storage', (e)=>{
    if(e.key && e.key.indexOf(PREFIX) === 0) notify(e.key.slice(PREFIX.length));
  });

  const user = {
    async canEdit(){ return true; },
    async isOwner(){ return true; },
    async id(){ return 'usuario-local'; },
    async profiles(ids){ const out = {}; [].concat(ids).forEach(id => { out[id] = { id, name: id === 'usuario-local' ? 'Usuário local' : '' }; }); return out; }
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

  window.claude = { async use(name){ return ({ db, user, downloads })[name] || null; } };

  // Apaga tudo que foi salvo neste navegador e recarrega (volta aos dados iniciais).
  // Uso: abra o console do navegador (F12) e digite  bigodeLocalReset()
  window.bigodeLocalReset = function(){
    Object.keys(localStorage).filter(k => k.indexOf(PREFIX) === 0).forEach(k => localStorage.removeItem(k));
    location.reload();
  };

  // Faixa discreta avisando que está em modo local
  document.addEventListener('DOMContentLoaded', ()=>{
    const b = document.createElement('div');
    b.textContent = 'Modo local — os dados ficam salvos só neste navegador';
    b.style.cssText = 'position:fixed;left:0;right:0;bottom:0;z-index:100;background:#1B1B1B;color:#fff;font:600 11.5px Inter,sans-serif;text-align:center;padding:5px 8px;opacity:.92;pointer-events:none;';
    document.body.appendChild(b);
  });
})();
