from __future__ import annotations

import io
import json
import os
import re
import sys
import unicodedata
import zipfile
from datetime import datetime
from pathlib import Path
from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit
from zoneinfo import ZoneInfo

import pandas as pd


DEFAULT_ONEDRIVE_URL = (
    "https://1drv.ms/x/c/2C62B039F7F27235/"
    "IQBiLMh61Pn1SZGH0PGCZjLNAWnyqMmXevmXjoQD05R0YGQ?e=zEdRU9"
)
OUTPUT_PATH = Path("data.json")


def normalizar(valor: object) -> str:
    texto = unicodedata.normalize("NFD", str(valor))
    texto = texto.encode("ascii", "ignore").decode("utf-8").lower().strip()
    return re.sub(r"[^a-z0-9]+", "_", texto).strip("_")


def nome_logico_aba(nome: str) -> str:
    return re.sub(r"^\d+_", "", normalizar(nome))


def url_download(url: str) -> str:
    partes = urlsplit(url.strip())
    parametros = dict(parse_qsl(partes.query, keep_blank_values=True))
    parametros["download"] = "1"
    return urlunsplit(
        (partes.scheme, partes.netloc, partes.path, urlencode(parametros), partes.fragment)
    )


def baixar_planilha(url: str) -> io.BytesIO:
    import requests

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 Chrome/150 Safari/537.36"
        )
    }
    resposta = requests.get(
        url_download(url),
        headers=headers,
        timeout=60,
        allow_redirects=True,
    )

    print("Status do download:", resposta.status_code)
    print("Tipo de conteúdo:", resposta.headers.get("content-type", "não informado"))
    print("Tamanho recebido:", len(resposta.content), "bytes")

    resposta.raise_for_status()
    conteudo = io.BytesIO(resposta.content)
    if not zipfile.is_zipfile(conteudo):
        inicio = resposta.content[:120].decode("utf-8", errors="replace")
        raise RuntimeError(
            "O OneDrive não retornou um arquivo XLSX válido. "
            f"Início da resposta: {inicio!r}"
        )

    conteudo.seek(0)
    return conteudo


def ler_aba(
    excel: pd.ExcelFile,
    nome_aba: str,
    mapa_colunas: dict[str, str],
    colunas_obrigatorias: list[str],
    chave_linha: str,
) -> list[dict]:
    tabela = pd.read_excel(
        excel,
        sheet_name=nome_aba,
        header=5,
        dtype=object,
    ).fillna("")
    tabela.columns = [normalizar(coluna) for coluna in tabela.columns]
    tabela = tabela.rename(columns=mapa_colunas)

    ausentes = [coluna for coluna in colunas_obrigatorias if coluna not in tabela.columns]
    if ausentes:
        raise RuntimeError(
            f"A aba {nome_aba!r} não contém as colunas obrigatórias: {ausentes}"
        )

    tabela = tabela[colunas_obrigatorias]
    tabela = tabela[
        tabela[chave_linha].map(lambda valor: str(valor).strip() not in ("", "0"))
    ]

    registros: list[dict] = []
    for registro in tabela.to_dict(orient="records"):
        limpo: dict = {}
        for chave, valor in registro.items():
            if pd.isna(valor):
                valor = ""
            elif hasattr(valor, "item"):
                valor = valor.item()
            if isinstance(valor, float) and valor.is_integer():
                valor = int(valor)
            if isinstance(valor, str):
                valor = valor.strip()
            limpo[chave] = valor
        registros.append(limpo)
    return registros


def extrair_dados(planilha: io.BytesIO) -> dict:
    excel = pd.ExcelFile(planilha, engine="openpyxl")
    abas = {nome_logico_aba(nome): nome for nome in excel.sheet_names}

    esperadas = {
        "cadastro_pessoas",
        "cadastro_ua",
        "cadastro_uassist",
        "cadastro_modulos",
    }
    ausentes = sorted(esperadas - set(abas))
    if ausentes:
        raise RuntimeError(f"Abas obrigatórias não encontradas: {ausentes}")

    pessoas = ler_aba(
        excel,
        abas["cadastro_pessoas"],
        {
            "nome": "nome",
            "masp": "masp",
            "vinculo": "vinculo",
            "setor_unidade_administrativa": "setor",
            "modulo": "modulo",
            "tipo_de_responsabilidade": "responsabilidade",
            "unidade_assistencial": "unidade_assistencial",
        },
        [
            "nome",
            "masp",
            "vinculo",
            "setor",
            "modulo",
            "responsabilidade",
            "unidade_assistencial",
        ],
        "nome",
    )

    unidades_administrativas = ler_aba(
        excel,
        abas["cadastro_ua"],
        {
            "id_unidadeadm": "id",
            "sigla": "sigla",
            "nome": "nome",
        },
        ["id", "sigla", "nome"],
        "sigla",
    )

    unidades_assistenciais = ler_aba(
        excel,
        abas["cadastro_uassist"],
        {
            "id_unidadeassist": "id",
            "sigla": "sigla",
            "nome": "nome",
        },
        ["id", "sigla", "nome"],
        "sigla",
    )

    modulos = ler_aba(
        excel,
        abas["cadastro_modulos"],
        {
            "id_modulo": "id",
            "sigla_ua": "sigla_ua",
            "id_unidadeadm": "id_ua",
            "unidade_administrativa": "ua",
            "nome_do_modulo": "nome",
            "detalhamento": "detalhamento",
        },
        ["id", "sigla_ua", "id_ua", "ua", "nome", "detalhamento"],
        "nome",
    )

    dados = {
        "meta": {
            "atualizado_em": datetime.now(
                ZoneInfo("America/Sao_Paulo")
            ).isoformat(timespec="seconds"),
            "origem": "OneDrive",
            "quantidades": {
                "pessoas": len(pessoas),
                "modulos": len(modulos),
                "uas": len(unidades_administrativas),
                "uassist": len(unidades_assistenciais),
            },
        },
        "pessoas": pessoas,
        "modulos": modulos,
        "uas": unidades_administrativas,
        "uassist": unidades_assistenciais,
    }

    vazias = [
        chave
        for chave in ("pessoas", "modulos", "uas", "uassist")
        if not dados[chave]
    ]
    if vazias:
        raise RuntimeError(
            "A atualização foi cancelada porque estas coleções ficaram vazias: "
            + ", ".join(vazias)
        )
    return dados


def main() -> int:
    url = os.environ.get("ONEDRIVE_URL", "").strip() or DEFAULT_ONEDRIVE_URL
    try:
        dados = extrair_dados(baixar_planilha(url))
        temporario = OUTPUT_PATH.with_suffix(".json.tmp")
        temporario.write_text(
            json.dumps(dados, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        temporario.replace(OUTPUT_PATH)
    except Exception as erro:
        print(f"ERRO: {erro}", file=sys.stderr)
        return 1

    quantidades = dados["meta"]["quantidades"]
    print(
        "data.json atualizado com sucesso:",
        ", ".join(f"{chave}={valor}" for chave, valor in quantidades.items()),
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
