// Returns the bookmarklet automation script with jobId and apiUrl interpolated.
// This script runs in the user's real browser on muisca.dian.gov.co.
export function buildBookmarkletScript(jobId: string, apiUrl: string): string {
  return `
(function() {
  'use strict';

  var JOB_ID = ${JSON.stringify(jobId)};
  var API_URL = ${JSON.stringify(apiUrl)};
  var BATCH = 5;
  var PREFIX = 'vistaConsultaEstadoRUT:formConsultaEstadoRUT';

  // ── Guard: prevent double-run ──────────────────────────────────────────────
  if (window.__cgRut) {
    window.__cgRut.show();
    return;
  }

  // ── Overlay UI ─────────────────────────────────────────────────────────────
  var overlay = document.createElement('div');
  overlay.id = '_cg_overlay';
  overlay.style.cssText = [
    'position:fixed', 'top:16px', 'right:16px', 'z-index:2147483647',
    'background:#0f0a1e', 'color:#fff', 'padding:18px 22px',
    'border-radius:14px', 'font-family:system-ui,sans-serif', 'font-size:14px',
    'min-width:300px', 'max-width:360px',
    'box-shadow:0 12px 40px rgba(0,0,0,0.5)',
    'border:1px solid rgba(124,58,237,0.4)',
    'backdrop-filter:blur(8px)',
  ].join(';');

  overlay.innerHTML = [
    '<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:12px">',
      '<span style="font-weight:800;font-size:15px;color:#a78bfa">ContaGO · RUT</span>',
      '<button id="_cg_close" style="background:none;border:none;color:rgba(255,255,255,0.4);font-size:18px;cursor:pointer;padding:0;line-height:1">&times;</button>',
    '</div>',
    '<div id="_cg_status" style="color:rgba(255,255,255,0.85);margin-bottom:10px">Iniciando...</div>',
    '<div style="height:6px;background:rgba(255,255,255,0.08);border-radius:99px;overflow:hidden;margin-bottom:8px">',
      '<div id="_cg_bar" style="height:100%;background:linear-gradient(90deg,#7c3aed,#a855f7);border-radius:99px;width:0%;transition:width 0.4s"></div>',
    '</div>',
    '<div id="_cg_detail" style="font-size:12px;color:rgba(255,255,255,0.5)"></div>',
  ].join('');

  document.body.appendChild(overlay);
  window.__cgRut = { show: function() { overlay.style.display = 'block'; } };

  document.getElementById('_cg_close').onclick = function() {
    overlay.style.display = 'none';
  };

  function setStatus(txt, detail) {
    document.getElementById('_cg_status').textContent = txt;
    document.getElementById('_cg_detail').textContent = detail || '';
  }

  function setProgress(done, total) {
    var pct = total > 0 ? Math.round(done / total * 100) : 0;
    document.getElementById('_cg_bar').style.width = pct + '%';
    document.getElementById('_cg_detail').textContent = done + ' / ' + total + ' NITs procesados';
  }

  // ── API helpers ────────────────────────────────────────────────────────────
  function apiFetch(path, opts) {
    return fetch(API_URL + path, Object.assign({
      headers: { 'Content-Type': 'application/json' },
    }, opts));
  }

  function getPending() {
    return apiFetch('/rut-consulta/browser-job/' + JOB_ID + '/pending?batch=' + BATCH)
      .then(function(r) { return r.json(); });
  }

  function sendResult(nit, data) {
    return apiFetch('/rut-consulta/browser-job/' + JOB_ID + '/result', {
      method: 'POST',
      body: JSON.stringify({ nit: nit, data: data }),
    }).catch(function() {});
  }

  // ── Turnstile detection ────────────────────────────────────────────────────
  // Wait until hddToken has a FRESH value (different from previousToken)
  function waitForToken(previousToken, maxMs) {
    return new Promise(function(resolve) {
      var deadline = Date.now() + maxMs;
      function poll() {
        var el = document.getElementById(PREFIX + ':hddToken');
        var val = el ? el.value : '';
        if (val && val.length > 50 && val !== previousToken) {
          resolve(val);
          return;
        }
        if (Date.now() >= deadline) { resolve(null); return; }
        setTimeout(poll, 600);
      }
      poll();
    });
  }

  // ── Result detection ───────────────────────────────────────────────────────
  // Wait until the estado field is non-empty (JSF AJAX updated the DOM)
  function waitForResult(previousEstado, maxMs) {
    return new Promise(function(resolve) {
      var deadline = Date.now() + maxMs;
      function poll() {
        var el = document.getElementById(PREFIX + ':estado');
        var val = el ? (el.textContent || '').trim() : '';
        if (val && val !== previousEstado) { resolve(val); return; }
        // Also check for error message
        var tbl = document.getElementById('tblMensajes');
        if (tbl) {
          var msg = (tbl.textContent || '').toLowerCase();
          if (msg.indexOf('no se encontr') !== -1 || msg.indexOf('no existe') !== -1 ||
              msg.indexOf('error') !== -1) {
            resolve('__error__'); return;
          }
        }
        if (Date.now() >= deadline) { resolve(null); return; }
        setTimeout(poll, 400);
      }
      poll();
    });
  }

  // ── Data extraction ────────────────────────────────────────────────────────
  function extractData() {
    var P = PREFIX + ':';
    function g(id) {
      var el = document.getElementById(P + id);
      return el ? (el.textContent || '').trim() : '';
    }

    var bodyLow = ((document.body && document.body.textContent) || '').toLowerCase();
    if (bodyLow.indexOf('error validando token') !== -1 ||
        bodyLow.indexOf('error validando captcha') !== -1) {
      return { error: 'CAPTCHA inválido — recarga la página y vuelve a intentar' };
    }

    var tbl = document.getElementById('tblMensajes');
    if (tbl) {
      var mt = (tbl.textContent || '').toLowerCase();
      if (mt.indexOf('no se encontr') !== -1 || mt.indexOf('no existe') !== -1) {
        return { error: 'NIT no registrado en el RUT' };
      }
      if (mt.indexOf('token') !== -1 || mt.indexOf('captcha') !== -1) {
        return { error: 'CAPTCHA requerido' };
      }
    }

    var estado = g('estado');
    if (!estado) return { error: 'Sin datos — posible error de sesión' };

    return {
      primerApellido:  g('primerApellido'),
      segundoApellido: g('segundoApellido'),
      primerNombre:    g('primerNombre'),
      otrosNombres:    g('otrosNombres'),
      razonSocial:     g('razonSocial') || g('denominacionRazonSocial'),
      estado:          estado,
      dv:              g('dv'),
    };
  }

  // ── Single NIT processing ──────────────────────────────────────────────────
  function sleep(ms) { return new Promise(function(r) { setTimeout(r, ms); }); }

  function processNit(nit, usedToken) {
    return new Promise(function(resolve, reject) {
      var nitEl = document.getElementById(PREFIX + ':numNit');
      if (!nitEl) { reject(new Error('Campo NIT no encontrado')); return; }

      // Capture current estado before submitting (to detect DOM change)
      var prevEstado = (function() {
        var el = document.getElementById(PREFIX + ':estado');
        return el ? (el.textContent || '').trim() : '';
      })();

      // Clear and fill NIT field
      nitEl.focus();
      nitEl.value = '';
      nitEl.dispatchEvent(new Event('input', { bubbles: true }));
      nitEl.dispatchEvent(new Event('change', { bubbles: true }));

      (function typeChar(i) {
        if (i >= nit.length) {
          nitEl.dispatchEvent(new Event('change', { bubbles: true }));
          // Wait for a fresh Turnstile token
          setStatus('Esperando verificación (' + nit + ')...', 'Cloudflare Turnstile...');
          waitForToken(usedToken, 90000).then(function(token) {
            if (!token) {
              resolve({ token: usedToken, data: { error: 'Turnstile no se renovó en 90s' } });
              return;
            }
            setStatus('Consultando ' + nit + '...', '');
            // Click the search button
            var btn = document.querySelector('[name="' + PREFIX + ':btnBuscar"]');
            if (!btn) { resolve({ token: token, data: { error: 'Botón no encontrado' } }); return; }
            btn.click();
            // Wait for DOM to update with results
            waitForResult(prevEstado, 12000).then(function() {
              sleep(600).then(function() {
                resolve({ token: token, data: extractData() });
              });
            });
          });
          return;
        }
        nitEl.value += nit[i];
        nitEl.dispatchEvent(new Event('input', { bubbles: true }));
        setTimeout(function() { typeChar(i + 1); }, 40);
      })(0);
    });
  }

  // ── Main loop ──────────────────────────────────────────────────────────────
  function run() {
    if (!window.location.hostname.includes('muisca.dian.gov.co')) {
      setStatus('Error: abre muisca.dian.gov.co primero', '');
      return;
    }

    setStatus('Conectando con ContaGO...', '');

    var processedCount = 0;
    var totalCount = 0;
    var lastToken = ''; // track the last used token to detect refresh

    function nextBatch() {
      getPending().then(function(pending) {
        if (!pending || !pending.nits) {
          setStatus('Error conectando con la API de ContaGO', '');
          return;
        }
        totalCount = pending.total;
        var nits = pending.nits;

        if (nits.length === 0) {
          setStatus('¡Listo! ' + processedCount + ' NITs procesados.', 'Ya puedes cerrar esta pestaña.');
          document.getElementById('_cg_bar').style.background = '#22c55e';
          window.__cgRut = null;
          return;
        }

        (function nextNit(i) {
          if (i >= nits.length) {
            if (!pending.done) {
              sleep(800).then(nextBatch);
            } else {
              setStatus('¡Listo! ' + processedCount + ' NITs procesados.', 'Ya puedes cerrar esta pestaña.');
              document.getElementById('_cg_bar').style.background = '#22c55e';
              window.__cgRut = null;
            }
            return;
          }
          var nit = nits[i];
          setStatus('Procesando ' + nit + '...', (i + 1) + ' de ' + nits.length + ' en este lote');
          processNit(nit, lastToken).then(function(res) {
            lastToken = res.token || lastToken;
            processedCount++;
            setProgress(processedCount, totalCount);
            sendResult(nit, res.data).then(function() { nextNit(i + 1); });
          }).catch(function(err) {
            processedCount++;
            setProgress(processedCount, totalCount);
            sendResult(nit, { error: err.message || 'Error desconocido' }).then(function() {
              nextNit(i + 1);
            });
          });
        })(0);
      }).catch(function(err) {
        setStatus('Error: ' + (err.message || 'conexión fallida'), '');
      });
    }

    nextBatch();
  }

  run();
})();
`.trimStart();
}

// Returns a fully self-contained javascript: URI with all automation code inlined.
// No external script load — avoids mixed-content blocks when DIAN is on HTTPS.
export function buildBookmarkletHref(jobId: string, apiUrl: string): string {
  const J = JSON.stringify(jobId);
  const A = JSON.stringify(apiUrl);
  const P = "'vistaConsultaEstadoRUT:formConsultaEstadoRUT'";

  const code = `(function(){
var J=${J},A=${A},B=5,P=${P};
if(window.__cgRut){window.__cgRut.show();return;}
var o=document.createElement('div');
o.style='position:fixed;top:16px;right:16px;z-index:2147483647;background:#0f0a1e;color:#fff;padding:18px 22px;border-radius:14px;font-family:system-ui,sans-serif;font-size:14px;min-width:300px;box-shadow:0 12px 40px rgba(0,0,0,.5);border:1px solid rgba(124,58,237,.4)';
o.innerHTML='<b style="color:#a78bfa">ContaGO · RUT</b><button id="_cgX" style="float:right;background:none;border:none;color:#fff;cursor:pointer;font-size:18px">&times;</button><div id="_cgS" style="margin:10px 0">Iniciando...</div><div style="height:6px;background:rgba(255,255,255,.08);border-radius:99px"><div id="_cgB" style="height:100%;background:#7c3aed;border-radius:99px;width:0%;transition:width .4s"></div></div><div id="_cgD" style="margin-top:6px;font-size:12px;color:rgba(255,255,255,.5)"></div>';
document.body.appendChild(o);
window.__cgRut={show:function(){o.style.display='block'}};
document.getElementById('_cgX').onclick=function(){o.style.display='none'};
function ss(t,d){document.getElementById('_cgS').textContent=t;if(d!=null)document.getElementById('_cgD').textContent=d}
function sp(c,t){document.getElementById('_cgB').style.width=(t>0?Math.round(c/t*100):0)+'%';document.getElementById('_cgD').textContent=c+' / '+t+' NITs'}
function af(p,op){return fetch(A+p,Object.assign({headers:{'Content-Type':'application/json'}},op))}
function gp(){return af('/rut-consulta/browser-job/'+J+'/pending?batch='+B).then(function(r){return r.json()})}
function sr(n,d){return af('/rut-consulta/browser-job/'+J+'/result',{method:'POST',body:JSON.stringify({nit:n,data:d})}).catch(function(){})}
function wt(prev,ms){return new Promise(function(res){var dl=Date.now()+ms;(function p(){var el=document.getElementById(P+':hddToken'),v=el?el.value:'';if(v&&v.length>50&&v!==prev){res(v);return;}if(Date.now()>=dl){res(null);return;}setTimeout(p,600)})()})}
function wr(prev,ms){return new Promise(function(res){var dl=Date.now()+ms;(function p(){var el=document.getElementById(P+':estado'),v=el?(el.textContent||'').trim():'';if(v&&v!==prev){res(v);return;}var t=document.getElementById('tblMensajes');if(t&&/(no se encontr|no existe|error)/i.test(t.textContent||'')){res('__err__');return;}if(Date.now()>=dl){res(null);return;}setTimeout(p,400)})()})}
function ex(){var g=function(id){var el=document.getElementById(P+':'+id);return el?(el.textContent||'').trim():''};var bl=((document.body&&document.body.textContent)||'').toLowerCase();if(bl.indexOf('error validando token')!==-1)return{error:'CAPTCHA inválido'};var t=document.getElementById('tblMensajes');if(t){var mt=(t.textContent||'').toLowerCase();if(mt.indexOf('no se encontr')!==-1||mt.indexOf('no existe')!==-1)return{error:'NIT no registrado'};if(mt.indexOf('token')!==-1)return{error:'CAPTCHA requerido'}}var e=g('estado');if(!e)return{error:'Sin datos'};return{primerApellido:g('primerApellido'),segundoApellido:g('segundoApellido'),primerNombre:g('primerNombre'),otrosNombres:g('otrosNombres'),razonSocial:g('razonSocial')||g('denominacionRazonSocial'),estado:e,dv:g('dv')}}
function sl(ms){return new Promise(function(r){setTimeout(r,ms)})}
function pn(n,ut){return new Promise(function(res,rej){var ne=document.getElementById(P+':numNit');if(!ne){rej(new Error('Campo NIT no encontrado'));return;}var pe=(function(){var el=document.getElementById(P+':estado');return el?(el.textContent||'').trim():''})();ne.focus();ne.value='';ne.dispatchEvent(new Event('input',{bubbles:true}));ne.dispatchEvent(new Event('change',{bubbles:true}));(function tc(i){if(i>=n.length){ne.dispatchEvent(new Event('change',{bubbles:true}));ss('Esperando verificación ('+n+')...');wt(ut,90000).then(function(tok){if(!tok){res({token:ut,data:{error:'Turnstile no renovado en 90s'}});return;}ss('Consultando '+n+'...');var b=document.querySelector('[name="'+P+':btnBuscar"]');if(!b){res({token:tok,data:{error:'Botón no encontrado'}});return;}b.click();wr(pe,12000).then(function(){sl(600).then(function(){res({token:tok,data:ex()})})})});return;}ne.value+=n[i];ne.dispatchEvent(new Event('input',{bubbles:true}));setTimeout(function(){tc(i+1)},40)})(0)})}
if(!window.location.hostname.includes('muisca.dian.gov.co')){ss('Error: abre muisca.dian.gov.co primero');return;}
ss('Conectando con ContaGO...');
var pc=0,tot=0,lt='';
(function nb(){gp().then(function(pd){if(!pd||!pd.nits){ss('Error conectando con la API');return;}tot=pd.total;var ns=pd.nits;if(ns.length===0){ss('¡Listo! '+pc+' NITs procesados.','Puedes cerrar esta pestaña.');document.getElementById('_cgB').style.background='#22c55e';window.__cgRut=null;return;}(function nn(i){if(i>=ns.length){if(!pd.done)setTimeout(nb,800);else{ss('¡Listo! '+pc+' NITs procesados.','Puedes cerrar esta pestaña.');document.getElementById('_cgB').style.background='#22c55e';window.__cgRut=null;}return;}var n=ns[i];ss('Procesando '+n+'...',(i+1)+' de '+ns.length);pn(n,lt).then(function(r){lt=r.token||lt;pc++;sp(pc,tot);sr(n,r.data).then(function(){nn(i+1)})}).catch(function(e){pc++;sp(pc,tot);sr(n,{error:e.message||'Error'}).then(function(){nn(i+1)})});})(0)}).catch(function(e){ss('Error: '+(e.message||'conexión fallida'))})})()
})()`;

  return `javascript:${encodeURIComponent(code)}`;
}
