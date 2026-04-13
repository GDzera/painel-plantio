"""
╔══════════════════════════════════════════════════════════════╗
║  auth_dashboard.py — Módulo de Segurança Reutilizável        ║
║  Importa em qualquer robô para proteger o HTML gerado        ║
╚══════════════════════════════════════════════════════════════╝

COMO USAR:
    1. Coloque este arquivo na mesma pasta dos seus robôs
    2. No seu robô, importe e use:

        from auth_dashboard import proteger_html, gerenciar_usuarios

        # No início do script, antes de tudo:
        if gerenciar_usuarios():
            exit()  # Se foi um comando de usuário, para aqui

        # Depois de gerar o HTML:
        html = build_html(...)
        html_protegido = proteger_html(html)
        with open('meu_dash.html', 'w') as f:
            f.write(html_protegido)

COMANDOS (iguais pra todos os robôs):
    python seu_robo.py --add usuario senha
    python seu_robo.py --remove usuario
    python seu_robo.py --users
    python seu_robo.py --no-auth
"""

import hashlib, base64, json, os, sys, argparse

# ═══════════════════════════════════════════════════════════
#  CONFIGURAÇÃO
# ═══════════════════════════════════════════════════════════
USUARIOS_FILE = "usuarios.json"  # Compartilhado entre todos os robôs


# ═══════════════════════════════════════════════════════════
#  CRIPTOGRAFIA AES-256-CBC (compatível com CryptoJS)
# ═══════════════════════════════════════════════════════════
def _evp_bytes_to_key(password, salt, key_len=32, iv_len=16):
    """Deriva key+IV no padrão EVP_BytesToKey (OpenSSL/CryptoJS)"""
    dtot, d = b'', b''
    while len(dtot) < key_len + iv_len:
        d = hashlib.md5(d + password + salt).digest()
        dtot += d
    return dtot[:key_len], dtot[key_len:key_len + iv_len]


def _aes_encrypt(plaintext, passphrase):
    """Criptografa texto com AES-256-CBC, compatível com CryptoJS"""
    try:
        from Crypto.Cipher import AES
    except ImportError:
        print("\n  ❌ ERRO: Instale pycryptodome:")
        print("     pip install pycryptodome\n")
        sys.exit(1)

    salt = os.urandom(8)
    key, iv = _evp_bytes_to_key(passphrase.encode(), salt)
    raw = plaintext.encode('utf-8')
    pad_len = 16 - (len(raw) % 16)
    padded = raw + bytes([pad_len] * pad_len)
    cipher = AES.new(key, AES.MODE_CBC, iv)
    ct = cipher.encrypt(padded)
    return base64.b64encode(b'Salted__' + salt + ct).decode()


def _aes_decrypt(ciphertext_b64, passphrase):
    """Descriptografa texto criptografado com _aes_encrypt"""
    try:
        from Crypto.Cipher import AES
    except ImportError:
        print("\n  ❌ ERRO: Instale pycryptodome:")
        print("     pip install pycryptodome\n")
        sys.exit(1)

    raw = base64.b64decode(ciphertext_b64)
    assert raw[:8] == b'Salted__', "Formato inválido"
    salt = raw[8:16]
    ct = raw[16:]
    key, iv = _evp_bytes_to_key(passphrase.encode(), salt)
    cipher = AES.new(key, AES.MODE_CBC, iv)
    decrypted = cipher.decrypt(ct)
    pad_len = decrypted[-1]
    return decrypted[:-pad_len].decode('utf-8')


# ═══════════════════════════════════════════════════════════
#  GESTÃO DE USUÁRIOS
# ═══════════════════════════════════════════════════════════
def _load_db():
    if os.path.exists(USUARIOS_FILE):
        with open(USUARIOS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {"master_key": base64.b64encode(os.urandom(32)).decode(), "users": {}}


def _save_db(data):
    with open(USUARIOS_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def add_usuario(username, password):
    """Cadastra um novo usuário"""
    db = _load_db()
    mk = db["master_key"]
    pw_hash = hashlib.sha256(password.encode()).hexdigest()
    mk_encrypted = _aes_encrypt(mk, password)
    db["users"][username.lower().strip()] = {
        "hash": pw_hash,
        "mk_enc": mk_encrypted
    }
    _save_db(db)
    print(f"  ✅ Usuário '{username}' cadastrado com sucesso.")


def remove_usuario(username):
    """Remove um usuário"""
    db = _load_db()
    user = username.lower().strip()
    if user in db["users"]:
        del db["users"][user]
        _save_db(db)
        print(f"  ✅ Usuário '{username}' removido.")
    else:
        print(f"  ❌ Usuário '{username}' não encontrado.")


def list_usuarios():
    """Lista todos os usuários cadastrados"""
    db = _load_db()
    if not db["users"]:
        print("\n  Nenhum usuário cadastrado.")
        print("  Use: python seu_robo.py --add usuario senha\n")
        return
    print(f"\n  {'Usuário':<25} {'Status'}")
    print(f"  {'-'*25} {'-'*10}")
    for u in sorted(db["users"].keys()):
        print(f"  {u:<25} ativo")
    print(f"\n  Total: {len(db['users'])} usuário(s)\n")


def alterar_senha(username, nova_senha):
    """Altera a senha de um usuário existente"""
    db = _load_db()
    user = username.lower().strip()
    if user not in db["users"]:
        print(f"  ❌ Usuário '{username}' não encontrado.")
        return
    mk = db["master_key"]
    pw_hash = hashlib.sha256(nova_senha.encode()).hexdigest()
    mk_encrypted = _aes_encrypt(mk, nova_senha)
    db["users"][user] = {"hash": pw_hash, "mk_enc": mk_encrypted}
    _save_db(db)
    print(f"  ✅ Senha de '{username}' alterada com sucesso.")


# ═══════════════════════════════════════════════════════════
#  DESCRIPTOGRAFAR (para você recuperar dados se necessário)
# ═══════════════════════════════════════════════════════════
def descriptografar_html(html_path, username, password):
    """
    Descriptografa um HTML protegido e salva a versão aberta.
    Uso: python auth_dashboard.py --decrypt dash.html usuario senha
    """
    db = _load_db()
    user = username.lower().strip()

    if user not in db["users"]:
        print(f"  ❌ Usuário '{user}' não encontrado.")
        return

    pw_hash = hashlib.sha256(password.encode()).hexdigest()
    if pw_hash != db["users"][user]["hash"]:
        print(f"  ❌ Senha incorreta.")
        return

    # Recuperar master key
    mk = _aes_decrypt(db["users"][user]["mk_enc"], password)

    # Ler HTML criptografado
    with open(html_path, 'r', encoding='utf-8') as f:
        html = f.read()

    # Extrair blob criptografado
    import re
    match = re.search(r'_E=["\']([^"\']+)["\']', html)
    if not match:
        print(f"  ❌ Arquivo não parece estar criptografado.")
        return

    encrypted_blob = match.group(1)
    body_decrypted = _aes_decrypt(encrypted_blob, mk)

    # Reconstruir HTML limpo
    # Pegar o head original
    head_end = html.index('</head>') + 7
    head = html[:head_end]

    # Remover scripts de crypto e login do head
    output_html = head + "\n<body>\n" + body_decrypted + "\n</body></html>"

    output_path = html_path.replace('.html', '_ABERTO.html')
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(output_html)

    print(f"  ✅ Arquivo descriptografado salvo em: {output_path}")
    print(f"     ⚠️  CUIDADO: este arquivo contém todos os dados sem proteção!")


# ═══════════════════════════════════════════════════════════
#  PROTEGER HTML (função principal que você chama no robô)
# ═══════════════════════════════════════════════════════════
def proteger_html(html):
    """
    Recebe o HTML completo do dashboard e retorna a versão criptografada
    com tela de login. Se não houver usuários cadastrados, retorna sem proteção.
    """
    db = _load_db()
    if not db["users"]:
        print("  ⚠️  Nenhum usuário cadastrado! HTML gerado SEM proteção.")
        print("     Use: python seu_robo.py --add admin suasenha")
        return html

    mk = db["master_key"]

    # Montar dict de usuários para o JS
    users_js = {}
    for u, info in db["users"].items():
        users_js[u] = {"h": info["hash"], "mk": info["mk_enc"]}

    # Extrair body
    body_start = html.index('<body>') + 6
    body_end = html.index('</body>')
    body_content = html[body_start:body_end]

    # Criptografar
    encrypted_body = _aes_encrypt(body_content, mk)
    users_json = json.dumps(users_js, ensure_ascii=False)

    # Tela de login
    login_html = '''
<style>
.login-overlay{position:fixed;top:0;left:0;width:100%;height:100%;background:linear-gradient(135deg,#004D23 0%,#00843D 50%,#006830 100%);display:flex;align-items:center;justify-content:center;z-index:9999;font-family:'Plus Jakarta Sans',sans-serif}
.login-box{background:#fff;border-radius:20px;padding:48px 40px;width:380px;box-shadow:0 20px 60px rgba(0,0,0,.3);text-align:center}
.login-logo{width:64px;height:64px;background:#00843D;border-radius:16px;display:inline-flex;align-items:center;justify-content:center;font-size:28px;font-weight:900;color:#fff;margin-bottom:20px;font-family:'JetBrains Mono',monospace}
.login-title{font-size:22px;font-weight:800;color:#1a2e22;margin-bottom:6px}
.login-sub{font-size:13px;color:#5a7265;margin-bottom:28px}
.login-input{width:100%;padding:14px 16px;border:2px solid #e0e8e3;border-radius:10px;font-size:14px;font-family:'Plus Jakarta Sans',sans-serif;margin-bottom:12px;outline:none;transition:border-color .2s;box-sizing:border-box}
.login-input:focus{border-color:#00843D}
.login-btn{width:100%;padding:14px;border:none;border-radius:10px;background:#00843D;color:#fff;font-size:15px;font-weight:700;cursor:pointer;font-family:'Plus Jakarta Sans',sans-serif;transition:background .2s;margin-top:4px}
.login-btn:hover{background:#006830}
.login-btn:active{transform:scale(0.98)}
.login-error{color:#d63031;font-size:12px;font-weight:600;margin-top:12px;min-height:18px}
.login-footer{font-size:10px;color:rgba(255,255,255,.5);position:fixed;bottom:16px;left:0;right:0;text-align:center}
.login-lock{font-size:11px;color:#5a7265;margin-top:16px;display:flex;align-items:center;justify-content:center;gap:4px}
.login-loading{display:none;margin-top:12px;font-size:12px;color:#00843D;font-weight:600}
</style>
<script src="https://cdnjs.cloudflare.com/ajax/libs/crypto-js/4.2.0/crypto-js.min.js"></script>

<div class="login-overlay" id="loginOverlay">
  <div class="login-box">
    <div class="login-logo">R</div>
    <div class="login-title">CCT Ipaussu</div>
    <div class="login-sub">Painel Integrado — Acesso Restrito</div>
    <input class="login-input" id="loginUser" type="text" placeholder="Usuário" autocomplete="username"
           onkeydown="if(event.key==='Enter')document.getElementById('loginPass').focus()">
    <input class="login-input" id="loginPass" type="password" placeholder="Senha" autocomplete="current-password"
           onkeydown="if(event.key==='Enter')doLogin()">
    <button class="login-btn" id="loginBtn" onclick="doLogin()">Entrar</button>
    <div class="login-loading" id="loginLoading">Descriptografando dados...</div>
    <div class="login-error" id="loginError"></div>
    <div class="login-lock">🔒 Dados criptografados com AES-256</div>
  </div>
  <div class="login-footer">Ambiente protegido • Dados criptografados end-to-end</div>
</div>

<div id="dashContent" style="display:none"></div>

<script>
const _U=''' + users_json + ''';
const _E="''' + encrypted_body + '''";

function sha256(s){return CryptoJS.SHA256(s).toString()}

function doLogin(){
  const user=document.getElementById('loginUser').value.trim().toLowerCase();
  const pass=document.getElementById('loginPass').value;
  const errEl=document.getElementById('loginError');
  const loadEl=document.getElementById('loginLoading');
  const btnEl=document.getElementById('loginBtn');

  errEl.textContent='';
  if(!user||!pass){errEl.textContent='Preencha usuário e senha.';return}

  const uData=_U[user];
  if(!uData){errEl.textContent='Usuário não encontrado.';return}

  const passHash=sha256(pass);
  if(passHash!==uData.h){errEl.textContent='Senha incorreta.';return}

  btnEl.disabled=true;
  btnEl.textContent='Aguarde...';
  loadEl.style.display='block';

  setTimeout(function(){
    try{
      var mk=CryptoJS.AES.decrypt(uData.mk,pass).toString(CryptoJS.enc.Utf8);
      if(!mk){errEl.textContent='Erro na descriptografia.';btnEl.disabled=false;btnEl.textContent='Entrar';loadEl.style.display='none';return}

      var body=CryptoJS.AES.decrypt(_E,mk).toString(CryptoJS.enc.Utf8);
      if(!body){errEl.textContent='Erro ao carregar dados.';btnEl.disabled=false;btnEl.textContent='Entrar';loadEl.style.display='none';return}

      document.getElementById('loginOverlay').style.display='none';
      var dc=document.getElementById('dashContent');
      dc.style.display='block';
      dc.innerHTML=body;

      // Re-create ONLY executable scripts (skip type=application/json data containers)
      dc.querySelectorAll('script').forEach(function(old){
        if(old.type && old.type !== 'text/javascript' && old.type !== '') return; // skip JSON data tags
        var s=document.createElement('script');
        if(old.src){s.src=old.src}else{s.textContent=old.textContent}
        old.parentNode.replaceChild(s,old);
      });

      // Fire DOMContentLoaded manually for scripts that listen to it
      window.dispatchEvent(new Event('DOMContentLoaded'));

      sessionStorage.setItem('cct_auth','1');
    }catch(e){
      console.error(e);
      errEl.textContent='Falha na autenticação.';
      btnEl.disabled=false;
      btnEl.textContent='Entrar';
      loadEl.style.display='none';
    }
  },100);
}

document.getElementById('loginUser').focus();
</script>
'''

    return html[:body_start] + login_html + html[body_end:]


# ═══════════════════════════════════════════════════════════
#  GERENCIADOR CLI (chame no início do seu robô)
# ═══════════════════════════════════════════════════════════
def gerenciar_usuarios():
    """
    Processa argumentos de linha de comando para gestão de usuários.
    Retorna True se processou um comando (e o robô deve parar).
    Retorna False se deve continuar normalmente.

    Uso no seu robô:
        from auth_dashboard import gerenciar_usuarios, proteger_html
        if gerenciar_usuarios():
            exit()
    """
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument('--add', nargs=2, metavar=('USER', 'SENHA'))
    parser.add_argument('--remove', metavar='USER')
    parser.add_argument('--users', action='store_true')
    parser.add_argument('--passwd', nargs=2, metavar=('USER', 'NOVA_SENHA'))
    parser.add_argument('--decrypt', nargs=3, metavar=('ARQUIVO', 'USER', 'SENHA'))
    parser.add_argument('--no-auth', action='store_true', dest='no_auth')

    args, _ = parser.parse_known_args()

    if args.add:
        add_usuario(args.add[0], args.add[1])
        return True
    elif args.remove:
        remove_usuario(args.remove)
        return True
    elif args.users:
        list_usuarios()
        return True
    elif args.passwd:
        alterar_senha(args.passwd[0], args.passwd[1])
        return True
    elif args.decrypt:
        descriptografar_html(args.decrypt[0], args.decrypt[1], args.decrypt[2])
        return True

    return False


def deve_proteger():
    """Retorna True se deve criptografar (não passou --no-auth)"""
    return '--no-auth' not in sys.argv


# ═══════════════════════════════════════════════════════════
#  CLI DIRETO (para gerenciar usuários sem precisar de robô)
# ═══════════════════════════════════════════════════════════
if __name__ == "__main__":
    print("\n  🔐 Auth Dashboard — Gerenciador de Acesso")
    print("  " + "=" * 45)

    if len(sys.argv) < 2:
        print("""
  Comandos disponíveis:

    python auth_dashboard.py --add usuario senha      Cadastrar
    python auth_dashboard.py --remove usuario         Remover
    python auth_dashboard.py --users                  Listar todos
    python auth_dashboard.py --passwd usuario novasenha  Trocar senha
    python auth_dashboard.py --decrypt arquivo.html usuario senha
        """)
    else:
        gerenciar_usuarios()
