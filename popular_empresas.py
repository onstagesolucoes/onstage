"""
Busca todas as empresas cadastradas no escritório via API da OneFlow,
percorrendo todas as páginas disponíveis, monta um DataFrame com os dados
e sobe apphash, razão social e CNPJ para a tabela 'oneflow_empresas' no Postgres.

Setup necessário (apenas uma vez):
  1. Adicione ao arquivo .env: ONEFLOW_TOKEN=<seu token de autorização>
  2. DATABASE_URL já deve estar definido no .env (mesmo banco usado pelas outras rotinas)
"""

import json
import os
from datetime import datetime

import pandas as pd
import requests
from dotenv import load_dotenv
from sqlalchemy import create_engine, MetaData, Table, select

load_dotenv()

URL_LISTAR_EMPRESAS = "https://rest.oneflow.com.br/api/oneflow/escritorio/empresas/listar"
URL_TOKEN_OMIE = "https://app.omie.com.br/api/portal/apps/{apphash}/token/"
APPHASH_ADMIN = "accsim_admin"
TOKEN = os.getenv("ONEFLOW_TOKEN")
DATABASE_URL = os.getenv("DATABASE_URL")
TABLE_NAME_EMPRESAS = "oneflow_empresas"
COLUNAS_EMPRESA = ("apphash", "razao", "cnpj")


def extrair_corpo(resposta_json: dict):
    """A API normalmente responde {"code", "result", ...}. Em erros de auth
    (ex.: token ausente/inválido) vem no formato {"statusCode", "body", ...},
    onde "body" pode vir como string JSON."""
    if "result" in resposta_json:
        codigo = resposta_json.get("code")
        if codigo is not None and str(codigo) != "200":
            raise RuntimeError(f"API retornou code={codigo}: {resposta_json}")
        return resposta_json["result"]

    status_code = resposta_json.get("statusCode")
    corpo = resposta_json.get("body", resposta_json)
    if isinstance(corpo, str):
        corpo = json.loads(corpo)
    if status_code is not None and status_code != 200:
        raise RuntimeError(f"API retornou statusCode={status_code}: {corpo}")
    return corpo


def extrair_lista_empresas(corpo):
    """Localiza a lista de empresas dentro do corpo da resposta,
    independente da chave usada pela API."""
    if isinstance(corpo, list):
        return corpo
    if isinstance(corpo, dict):
        for chave in ("empresas", "data", "items", "registros", "resultado", "lista"):
            valor = corpo.get(chave)
            if isinstance(valor, list):
                return valor
    raise ValueError(f"Não foi possível localizar a lista de empresas na resposta: {corpo}")


def buscar_todas_empresas() -> pd.DataFrame:
    if not TOKEN:
        raise RuntimeError("Defina ONEFLOW_TOKEN no arquivo .env antes de executar.")

    headers = {"Authorization": TOKEN}
    todas_empresas = []
    pagina = 1

    while True:
        resp = requests.get(URL_LISTAR_EMPRESAS, headers=headers, params={"pagina": pagina})
        resp.raise_for_status()
        corpo = extrair_corpo(resp.json())
        empresas = extrair_lista_empresas(corpo)

        if not empresas:
            break

        print(f"Página {pagina}: {len(empresas)} empresa(s)")
        todas_empresas.extend(empresas)

        total_paginas = corpo.get("totalPaginas") if isinstance(corpo, dict) else None
        if total_paginas is not None and pagina >= total_paginas:
            break

        pagina += 1

    return pd.DataFrame(todas_empresas)


def limpar_cnpj(cnpj) -> str:
    """Remove tudo que não for dígito (pontos, barra, hífen)."""
    if not cnpj:
        return cnpj
    return "".join(ch for ch in str(cnpj) if ch.isdigit())


def upsert_empresa(conn, tabela: Table, dados: dict) -> str:
    existente = conn.execute(
        select(tabela).where(tabela.c.apphash == dados["apphash"])
    ).fetchone()

    if existente is None:
        conn.execute(tabela.insert().values(**dados))
        return "inserido"

    valores_update = {
        campo: dados[campo]
        for campo in ("razao", "cnpj")
        if existente._mapping.get(campo) != dados[campo]
    }
    if valores_update:
        conn.execute(
            tabela.update().where(tabela.c.apphash == dados["apphash"]).values(**valores_update)
        )
        return "atualizado"

    return "sem alterações"


def salvar_no_banco(df: pd.DataFrame) -> list[str]:
    """Faz upsert das empresas no banco e retorna os apphashes recém-inseridos."""
    if not DATABASE_URL:
        raise RuntimeError("Defina DATABASE_URL no arquivo .env antes de executar.")

    engine = create_engine(DATABASE_URL, echo=False)
    tabela = Table(TABLE_NAME_EMPRESAS, MetaData(), autoload_with=engine)

    contadores = {"inserido": 0, "atualizado": 0, "sem_alteracoes": 0, "sem_apphash": 0}
    apphashes_inseridos = []

    for registro in df.to_dict(orient="records"):
        dados = {coluna: registro.get(coluna) for coluna in COLUNAS_EMPRESA}
        if not dados["apphash"]:
            contadores["sem_apphash"] += 1
            continue

        dados["cnpj"] = limpar_cnpj(dados["cnpj"])

        with engine.begin() as conn:
            resultado = upsert_empresa(conn, tabela, dados)

        if resultado == "inserido":
            contadores["inserido"] += 1
            apphashes_inseridos.append(dados["apphash"])
        elif resultado == "atualizado":
            contadores["atualizado"] += 1
        else:
            contadores["sem_alteracoes"] += 1

    print(
        f"\nBanco ({TABLE_NAME_EMPRESAS}): "
        f"{contadores['inserido']} inserido(s), "
        f"{contadores['atualizado']} atualizado(s), "
        f"{contadores['sem_alteracoes']} sem alterações, "
        f"{contadores['sem_apphash']} sem apphash (ignorado(s))."
    )

    return apphashes_inseridos


def buscar_token_usuario(engine, tabela: Table) -> str:
    """Busca o token do usuário admin (linha apphash='accsim_admin'), usado para
    autenticar a chamada de coleta de token de cada empresa na API da Omie."""
    with engine.connect() as conn:
        linha = conn.execute(
            select(tabela.c.token).where(tabela.c.apphash == APPHASH_ADMIN)
        ).fetchone()

    if linha is None or not linha.token:
        raise RuntimeError(f"Não encontrei token do usuário admin (apphash='{APPHASH_ADMIN}').")

    return linha.token


def coletar_token_empresa(token_usuario: str, apphash: str) -> dict:
    resp = requests.get(
        URL_TOKEN_OMIE.format(apphash=apphash),
        headers={"Authorization": f"Bearer {token_usuario}"},
    )
    resp.raise_for_status()
    dados = resp.json()

    if "token" not in dados or "refresh_token" not in dados:
        raise RuntimeError(f"Resposta inesperada ao buscar token de '{apphash}': {dados}")

    return dados


def atualizar_tokens_novas_empresas(apphashes_novos: list[str]):
    """Para cada empresa nova, busca token/refresh_token na API da Omie e grava no banco."""
    if not apphashes_novos:
        print("\nNenhuma empresa nova: nada a fazer na coleta de tokens.")
        return

    engine = create_engine(DATABASE_URL, echo=False)
    tabela = Table(TABLE_NAME_EMPRESAS, MetaData(), autoload_with=engine)
    token_usuario = buscar_token_usuario(engine, tabela)

    print(f"\nColetando tokens para {len(apphashes_novos)} empresa(s) nova(s)...")
    sucesso = 0
    falha = 0

    for apphash in apphashes_novos:
        try:
            dados = coletar_token_empresa(token_usuario, apphash)
        except Exception as exc:
            print(f"  Falha ao coletar token de '{apphash}': {exc}")
            falha += 1
            continue

        with engine.begin() as conn:
            conn.execute(
                tabela.update().where(tabela.c.apphash == apphash).values(
                    token=dados["token"],
                    refresh_token=dados["refresh_token"],
                    data_renovacao=datetime.now(),
                )
            )
        print(f"  Token coletado para '{apphash}'.")
        sucesso += 1

    print(f"\nColeta de tokens: {sucesso} sucesso(s), {falha} falha(s).")


def main():
    df = buscar_todas_empresas()
    print(f"\nTotal de empresas coletadas: {len(df)}")
    apphashes_novos = salvar_no_banco(df)
    atualizar_tokens_novas_empresas(apphashes_novos)


if __name__ == "__main__":
    main()
