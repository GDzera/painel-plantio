"""
╔══════════════════════════════════════════════════════════════╗
║  EXEMPLO: Como proteger QUALQUER robô com auth_dashboard.py  ║
║                                                              ║
║  Antes:  seu robô gera HTML aberto                           ║
║  Depois: seu robô gera HTML criptografado com login          ║
╚══════════════════════════════════════════════════════════════╝

ARQUIVOS NECESSÁRIOS NA PASTA:
  - auth_dashboard.py   (o módulo de segurança)
  - usuarios.json       (gerado automaticamente)
  - seu_robo.py         (qualquer robô que gere HTML)
"""

# ============================================================
# EXEMPLO 1: gerar_dash.py (mínimo necessário)
# ============================================================
"""
# Adicione estas 2 linhas NO TOPO do seu script:
from auth_dashboard import gerenciar_usuarios, proteger_html, deve_proteger

# Adicione esta linha ANTES de qualquer outra lógica:
if gerenciar_usuarios():
    exit()

# ... todo o código do seu robô aqui ...
# ... no final, onde você salva o HTML:

html = build_html(...)  # sua função que gera o HTML

# Adicione estas 2 linhas ANTES do f.write:
if deve_proteger():
    html = proteger_html(html)

with open('meu_dashboard.html', 'w', encoding='utf-8') as f:
    f.write(html)
"""


# ============================================================
# EXEMPLO 2: Integração completa (robo_checkapontamentos.py)
# ============================================================
"""
import pandas as pd
import json
# ... seus imports normais ...

# ── SEGURANÇA (adicionar no topo) ──
from auth_dashboard import gerenciar_usuarios, proteger_html, deve_proteger
if gerenciar_usuarios():
    exit()
# ── FIM da parte de segurança do topo ──


def processar_dados():
    # ... sua lógica de processamento ...
    return dados

def build_html(dados):
    # ... sua lógica de geração HTML ...
    return '<html>...</html>'

def main():
    dados = processar_dados()
    html = build_html(dados)

    # ── SEGURANÇA (adicionar antes de salvar) ──
    if deve_proteger():
        html = proteger_html(html)
    # ── FIM ──

    with open('check_apontamentos.html', 'w', encoding='utf-8') as f:
        f.write(html)
    print('Dashboard gerado!')

if __name__ == '__main__':
    main()
"""


# ============================================================
# EXEMPLO 3: Se seu robô tem loop de monitoramento
# ============================================================
"""
from auth_dashboard import gerenciar_usuarios, proteger_html, deve_proteger
if gerenciar_usuarios():
    exit()

def monitorar():
    while True:
        # ... detecta mudanças ...

        html = gerar_html()

        if deve_proteger():
            html = proteger_html(html)

        with open('dash.html', 'w') as f:
            f.write(html)

        time.sleep(5)
"""

print("""
╔══════════════════════════════════════════════════════════════╗
║  RESUMO: Para proteger qualquer robô, você precisa de       ║
║  apenas 4 linhas de código:                                  ║
║                                                              ║
║  TOPO DO ARQUIVO:                                            ║
║    from auth_dashboard import gerenciar_usuarios,            ║
║         proteger_html, deve_proteger                         ║
║    if gerenciar_usuarios(): exit()                           ║
║                                                              ║
║  ANTES DE SALVAR O HTML:                                     ║
║    if deve_proteger():                                       ║
║        html = proteger_html(html)                            ║
║                                                              ║
║  COMANDOS (funcionam em qualquer robô):                      ║
║    python seu_robo.py --add usuario senha                    ║
║    python seu_robo.py --remove usuario                       ║
║    python seu_robo.py --users                                ║
║    python seu_robo.py --passwd usuario novasenha             ║
║    python seu_robo.py --no-auth        (gerar sem proteção)  ║
║                                                              ║
║  DESCRIPTOGRAFAR (recuperar dados):                          ║
║    python auth_dashboard.py --decrypt dash.html admin senha  ║
╚══════════════════════════════════════════════════════════════╝
""")
