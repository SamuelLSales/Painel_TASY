import requests, pandas as pd, io, unicodedata, json

def norm(s):
    return unicodedata.normalize('NFD', str(s)).encode('ascii', 'ignore').decode('utf-8').lower().strip()

def handler(event, context):
    ONEDRIVE_URL = "https://1drv.ms/x/c/2C62B039F7F27235/IQC0sPfxiHBLRommd2UGb8aLAaGrLnl-5TF_topDMBLFWQQ?download=1"
    data = {"pessoas": [], "modulos": [], "uas": [], "uassist": []}
    
    try:
        hdrs = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
        res = requests.get(ONEDRIVE_URL, headers=hdrs, timeout=12)
        if res.status_code == 200:
            xls = pd.ExcelFile(io.BytesIO(res.content), engine='openpyxl')
            for sheet_name in xls.sheet_names:
                name = norm(sheet_name)
                if 'pessoa' in name:
                    df = pd.read_excel(xls, sheet_name=sheet_name, header=5).fillna('')
                    df = df.rename(columns={'Nome':'nome','MASP':'masp','Vínculo':'vinculo','Setor / Unidade Administrativa':'setor','Módulo':'modulo','Tipo de Responsabilidade':'responsabilidade','Unidade Assistencial':'unidade_assistencial'})
                    data['pessoas'] = df.to_dict(orient='records')
                elif 'uassist' in name:
                    df = pd.read_excel(xls, sheet_name=sheet_name, header=5).fillna('')
                    df = df.rename(columns={'ID_UnidadeAssist':'id','Sigla':'sigla','Nome':'nome'})
                    data['uassist'] = df.to_dict(orient='records')
                elif 'cadastrar ua' == name:
                    df = pd.read_excel(xls, sheet_name=sheet_name, header=5).fillna('')
                    df = df.rename(columns={'ID_UnidadeAdm':'id','Sigla':'sigla','Nome':'nome'})
                    data['uas'] = df.to_dict(orient='records')
                elif 'modulo' in name:
                    df = pd.read_excel(xls, sheet_name=sheet_name, header=5).fillna('')
                    df = df.rename(columns={'ID_Modulo':'id','Sigla UA':'sigla_ua','ID_UnidadeAdm':'id_ua','Unidade Administrativa':'ua','Nome do Módulo':'nome','Detalhamento':'detalhamento'})
                    data['modulos'] = df.to_dict(orient='records')
    except Exception as e:
        print("Erro ao processar planilha:", e)

    return {
        "statusCode": 200,
        "headers": {
            "Content-Type": "application/json",
            "Access-Control-Allow-Origin": "*"
        },
        "body": json.dumps(data, ensure_ascii=False)
    }
