import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import json, os, urllib.request

st.set_page_config(
    page_title="Painel Governança TASY | FHEMIG",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS para ocultar barras padrão do Streamlit e garantir fundo 100% branco sem scroll
st.markdown("""
    <style>
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}
        /* Ocultar marca d'água e toolbar do Streamlit Cloud */
        [data-testid="stToolbar"] {visibility: hidden !important;}
        [data-testid="stDecoration"] {display: none !important;}
        [data-testid="stDeployButton"] {display: none !important;}
        .stDeployButton {display: none !important;}
        #stDecoration {display: none !important;}
        button[kind="header"] {display: none !important;}
        html, body, [data-testid="stAppViewContainer"], .stApp {
            background-color: #FAF8F3 !important;
            margin: 0 !important;
            padding: 0 !important;
        }
        .block-container {
            padding: 0rem !important;
            margin: 0rem !important;
            max-width: 100% !important;
        }
        iframe {
            width: 100% !important;
            border: none !important;
            display: block !important;
        }
    </style>
""", unsafe_allow_html=True)

ONEDRIVE_DIRECT_URL = "https://1drv.ms/x/c/2C62B039F7F27235/IQC0sPfxiHBLRommd2UGb8aLAaGrLnl-5TF_topDMBLFWQQ?download=1"
LOCAL_JSON = os.path.join(os.path.dirname(__file__), "Governanca_TASY_FHEMIG.json")
HTML_FILE = os.path.join(os.path.dirname(__file__), "Painel_Governanca_TASY_FHEMIG_Bordas_Mais_Vivas.html")

import requests, unicodedata

def norm(s):
    return unicodedata.normalize('NFD', str(s)).encode('ascii', 'ignore').decode('utf-8').lower().strip()

@st.cache_data(ttl=60)
def carregar_dados_onedrive():
    data = {"pessoas": [], "modulos": [], "uas": [], "uassist": []}
    
    # 1. Tentar baixar a planilha AO VIVO do OneDrive público
    try:
        import io
        hdrs = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
        res = requests.get(ONEDRIVE_DIRECT_URL, headers=hdrs, allow_redirects=True, timeout=15)
        if res.status_code == 200:
            xls = pd.ExcelFile(io.BytesIO(res.content), engine='openpyxl')

            for sheet_name in xls.sheet_names:
                name = norm(sheet_name)

                # ─── ABA: Cadastrar Pessoas ─────────────────────────────────
                if 'pessoa' in name or 'agente' in name:
                    df = pd.read_excel(xls, sheet_name=sheet_name, header=5).fillna('')
                    # Colunas da planilha → chaves do JS
                    col_map = {
                        'Nome': 'nome',
                        'MASP': 'masp',
                        'Vínculo': 'vinculo',
                        'Setor / Unidade Administrativa': 'setor',
                        'Módulo': 'modulo',
                        'Tipo de Responsabilidade': 'responsabilidade',
                        'Unidade Assistencial': 'unidade_assistencial',
                    }
                    df = df.rename(columns=col_map)
                    keep = [c for c in col_map.values() if c in df.columns]
                    data['pessoas'] = df[keep].to_dict(orient='records')

                # ─── ABA: Cadastrar UAssist ─────────────────────────────────
                elif 'uassist' in name:
                    df = pd.read_excel(xls, sheet_name=sheet_name, header=5).fillna('')
                    col_map = {
                        'ID_UnidadeAssist': 'id',
                        'Sigla': 'sigla',
                        'Nome': 'nome',
                    }
                    df = df.rename(columns=col_map)
                    keep = [c for c in col_map.values() if c in df.columns]
                    data['uassist'] = df[keep].to_dict(orient='records')

                # ─── ABA: Cadastrar UA (unidades administrativas) ──────────────
                elif 'cadastrar ua' == name:
                    df = pd.read_excel(xls, sheet_name=sheet_name, header=5).fillna('')
                    col_map = {
                        'ID_UnidadeAdm': 'id',
                        'Sigla': 'sigla',
                        'Nome': 'nome',
                    }
                    df = df.rename(columns=col_map)
                    keep = [c for c in col_map.values() if c in df.columns]
                    data['uas'] = df[keep].to_dict(orient='records')

                # ─── ABA: Cadastrar Módulos ─────────────────────────────────
                elif 'modulo' in name:
                    df = pd.read_excel(xls, sheet_name=sheet_name, header=5).fillna('')
                    col_map = {
                        'ID_Modulo': 'id',
                        'Sigla UA': 'sigla_ua',
                        'ID_UnidadeAdm': 'id_ua',
                        'Unidade Administrativa': 'ua',
                        'Nome do Módulo': 'nome',
                        'Detalhamento': 'detalhamento',
                    }
                    df = df.rename(columns=col_map)
                    keep = [c for c in col_map.values() if c in df.columns]
                    data['modulos'] = df[keep].to_dict(orient='records')

            if data['pessoas'] or data['modulos']:
                return data
    except Exception as e:
        pass  # cai para o fallback local

    # 2. Fallback: Ler do JSON local se o OneDrive estiver offline
    if os.path.exists(LOCAL_JSON):
        try:
            with open(LOCAL_JSON, 'r', encoding='utf-8') as f:
                content = json.load(f)
                if content and content.get('pessoas'):
                    return content
        except Exception:
            pass

    return data

def main():
    dados = carregar_dados_onedrive()
    json_str = json.dumps(dados, ensure_ascii=False)
    
    if os.path.exists(HTML_FILE):
        with open(HTML_FILE, 'r', encoding='utf-8') as f:
            html_content = f.read()
        
        # Injeta os dados dinâmicos de forma 100% segura contra erros de sintaxe JS
        injection_block = f"""
<script id="injected-data" type="application/json">
{json_str}
</script>
<script>
  try {{
    window.INJECTED_DATA = JSON.parse(document.getElementById('injected-data').textContent);
  }} catch(e) {{
    console.error("Erro ao ler INJECTED_DATA:", e);
  }}
</script>
"""
        full_html = html_content.replace("</head>", f"{injection_block}\n</head>")
        
        # Renderiza a interface visual em HTML/CSS com 100% da fidelidade
        components.html(full_html, height=860, scrolling=True)
    else:
        st.error("Arquivo HTML do Painel não encontrado.")

if __name__ == "__main__":
    main()
