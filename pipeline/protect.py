"""Protege as páginas do painel com senha.

O conteúdo de cada página é criptografado (AES-256-GCM, chave derivada da senha com
PBKDF2-SHA256). O arquivo publicado só contém o texto cifrado e uma tela de senha:
sem a senha, os dados não podem ser lidos, mesmo abrindo o código da página.
Todas as páginas de um mesmo build usam o mesmo sal, então a senha é pedida uma vez
e a chave fica guardada na aba (ou no navegador, se a pessoa marcar "lembrar").
"""
import base64, json, os
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives.ciphers.aead import AESGCM

ITER = 310_000

def new_salt():
    return os.urandom(16)

def derive_key(password, salt):
    return PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=ITER).derive(password.encode('utf-8'))

def full_doc(content):
    if content.lstrip().lower().startswith('<!doctype'):
        return content
    return ('<!doctype html><html lang="pt-BR"><head><meta charset="utf-8">'
            '<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">'
            '<style>body{margin:0}</style></head><body>' + content + '</body></html>')

def protect(content, key, salt, title, standalone):
    """content: HTML da página; standalone=True para arquivos de apoio (documento completo)."""
    iv = os.urandom(12)
    ct = AESGCM(key).encrypt(iv, full_doc(content).encode('utf-8'), None)
    payload = json.dumps({'s': base64.b64encode(salt).decode(), 'i': base64.b64encode(iv).decode(),
                          'c': base64.b64encode(ct).decode(), 'n': ITER})
    page = LOADER.replace('__TITLE__', title).replace('__PAYLOAD__', payload)
    return full_doc(page) if standalone else page

LOADER = r'''<meta charset="utf-8">
<title>__TITLE__</title>
<meta name="claude:viewport-fit-sides" content="cover">
<style>
/* Tela de senha do Painel Estratégico: só o texto cifrado fica no arquivo. */
:root{--bg:#f4f4f2;--panel:#fff;--line:#e1e0dc;--ink:#111;--ink-2:#4d4c48;--ink-3:#7a7974;--err:#b3261e;--focus:#2a78d6}
@media (prefers-color-scheme: dark){:root:not([data-theme="light"]){--bg:#141517;--panel:#1c1d20;--line:#33353a;--ink:#f1f0ed;--ink-2:#bdbbb5;--ink-3:#8c8a85;--err:#f2b8b5;--focus:#4e86da;color-scheme:dark}}
:root[data-theme="dark"]{--bg:#141517;--panel:#1c1d20;--line:#33353a;--ink:#f1f0ed;--ink-2:#bdbbb5;--ink-3:#8c8a85;--err:#f2b8b5;--focus:#4e86da;color-scheme:dark}
*{box-sizing:border-box}
html,body{height:100%}
body{background:var(--bg);color:var(--ink);font:15px/1.5 system-ui,-apple-system,"Segoe UI",sans-serif;margin:0;display:flex;flex-direction:column;
 padding-left:max(16px,env(safe-area-inset-left,0px));padding-right:max(16px,env(safe-area-inset-right,0px))}
.bar{margin:0 calc(-1*max(16px,env(safe-area-inset-right,0px))) 0 calc(-1*max(16px,env(safe-area-inset-left,0px)));background:#111;color:#f1f0ed;font-weight:800;
 padding:12px max(16px,env(safe-area-inset-right,0px)) 12px max(16px,env(safe-area-inset-left,0px))}
main{flex:1;display:grid;place-items:center;padding-block:32px}
form{width:100%;max-width:380px;background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:24px;display:flex;flex-direction:column;gap:12px}
h1{font-size:22px;line-height:1.2;margin:0}
p{margin:0;color:var(--ink-2);font-size:14px}
label{font-size:13px;font-weight:600;color:var(--ink-2)}
input[type=password]{font:inherit;color:var(--ink);background:var(--bg);border:1px solid var(--line);border-radius:8px;padding:10px 12px;width:100%}
input:focus-visible,button:focus-visible{outline:2px solid var(--focus);outline-offset:2px}
.row{display:flex;align-items:center;gap:8px;font-size:13px;color:var(--ink-2)}
button{font:600 15px system-ui,sans-serif;background:var(--ink);color:var(--bg);border:0;border-radius:8px;padding:11px 14px;cursor:pointer}
button[disabled]{opacity:.6;cursor:progress}
.msg{font-size:13px;min-height:1.2em;color:var(--err)}
</style>
<div class="bar">Painel Estratégico Dexco</div>
<main>
<form id="f" autocomplete="on">
  <h1>Acesso do time</h1>
  <p>Digite a senha do painel para ver os dados.</p>
  <label for="pw">Senha</label>
  <input type="password" id="pw" autocomplete="current-password" required>
  <label class="row" for="lembrar"><input type="checkbox" id="lembrar"> Lembrar neste navegador</label>
  <button id="go" type="submit">Entrar</button>
  <div class="msg" id="msg" role="status"></div>
</form>
</main>
<script>
(function(){
  const P=__PAYLOAD__;
  const b64=s=>Uint8Array.from(atob(s),c=>c.charCodeAt(0));
  const toB64=buf=>btoa(String.fromCharCode(...new Uint8Array(buf)));
  const KEY='painel-chave-'+P.s;
  const get=st=>{try{return st.getItem(KEY)}catch(e){return null}};
  const put=(st,v)=>{try{st.setItem(KEY,v)}catch(e){}};
  const del=st=>{try{st.removeItem(KEY)}catch(e){}};
  async function decrypt(raw){
    const key=await crypto.subtle.importKey('raw',raw,{name:'AES-GCM'},false,['decrypt']);
    return new TextDecoder().decode(await crypto.subtle.decrypt({name:'AES-GCM',iv:b64(P.i)},key,b64(P.c)));
  }
  // troca a página só depois de terminar de carregar; antes disso o document.write mistura as duas
  function show(html){const go=()=>{document.open();document.write(html);document.close()};if(document.readyState==='complete')go();else addEventListener('load',go,{once:true})}
  async function derive(pw){
    const base=await crypto.subtle.importKey('raw',new TextEncoder().encode(pw),'PBKDF2',false,['deriveBits']);
    return crypto.subtle.deriveBits({name:'PBKDF2',hash:'SHA-256',salt:b64(P.s),iterations:P.n},base,256);
  }
  const msg=document.getElementById('msg'),go=document.getElementById('go');
  if(!(window.crypto&&crypto.subtle)){msg.textContent='Este navegador não consegue abrir o painel. Use uma versão atual do Chrome, Edge, Safari ou Firefox.';go.disabled=true;return}
  const saved=get(sessionStorage)||get(localStorage);
  if(saved){decrypt(b64(saved)).then(show).catch(()=>{del(sessionStorage);del(localStorage)})}
  document.getElementById('f').addEventListener('submit',async ev=>{
    ev.preventDefault();msg.textContent='';go.disabled=true;go.textContent='Abrindo…';
    try{
      const raw=await derive(document.getElementById('pw').value);
      const html=await decrypt(raw);
      const s=toB64(raw);put(sessionStorage,s);if(document.getElementById('lembrar').checked)put(localStorage,s);
      show(html);
    }catch(e){
      go.disabled=false;go.textContent='Entrar';msg.textContent='Senha incorreta. Confira e tente de novo.';
      document.getElementById('pw').select();
    }
  });
})();
</script>'''
