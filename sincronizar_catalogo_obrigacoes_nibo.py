"""
Sincroniza o CATÁLOGO de tipos de obrigação do Nibo para o Supabase.

Endpoint: GET https://api.nibo.com.br/accountant/api/v1/accountingfirms/{id}/obligations
(https://nibo.readme.io/reference/listar-obrigacoes)

Atenção: esse endpoint devolve o catálogo de TIPOS de obrigação do escritório,
não os vencimentos por cliente. Os vencimentos (quem deve o quê, quando) já
são coletados pelo workflow n8n "FISCAL 01" via
GET /accountingfirms/{id}/reports/obligations/complete, e ficam em
'calendario_obrigacoes'. Não duplicar essa carga aqui.

O que este script resolve: a tabela 'obrigacoes_depara' (que traduz o ID de
obrigação do Nibo para o ob_key interno usado por perfis/ClickUp/OneFlow) só
cobre 12 dos códigos que o Nibo retorna. Este script grava o catálogo bruto em
'stg_nibo_obrigacoes_catalogo' e reporta quais IDs ainda não têm de-para, para
alguém preencher nome_amigavel/frequencia/mes_competencia com critério humano
-- a tradução não é inferível automaticamente.

Setup necessário (apenas uma vez):
  Adicione ao .env:
    NIBO_API_KEY=<chave do escritório, header X-Api-Key>
    NIBO_ACCOUNTING_FIRM_ID=<id do escritório no Nibo>
    DATABASE_URL=<mesmo banco usado pelas outras rotinas>
"""

import os

import requests
from dotenv import load_dotenv
from sqlalchemy import create_engine, func, MetaData, select, Table, text

load_dotenv()

NIBO_API_KEY = os.getenv("NIBO_API_KEY")
ACCOUNTING_FIRM_ID = os.getenv("NIBO_ACCOUNTING_FIRM_ID")
DATABASE_URL = os.getenv("DATABASE_URL")

URL_OBRIGACOES = "https://api.nibo.com.br/accountant/api/v1/accountingfirms/{firm}/obligations"
TABLE_NAME = "stg_nibo_obrigacoes_catalogo"
PAGE_SIZE = 100


def buscar_catalogo_obrigacoes() -> list[dict]:
    if not NIBO_API_KEY or not ACCOUNTING_FIRM_ID:
        raise RuntimeError(
            "Defina NIBO_API_KEY e NIBO_ACCOUNTING_FIRM_ID no .env antes de executar."
        )

    url = URL_OBRIGACOES.format(firm=ACCOUNTING_FIRM_ID)
    headers = {"X-Api-Key": NIBO_API_KEY}
    itens: list[dict] = []
    skip = 0

    while True:
        resp = requests.get(url, headers=headers, params={"$top": PAGE_SIZE, "$skip": skip})
        resp.raise_for_status()
        corpo = resp.json()
        pagina = corpo.get("items", corpo) if isinstance(corpo, dict) else corpo

        if not isinstance(pagina, list):
            raise ValueError(f"Formato de resposta inesperado em skip={skip}: {corpo}")
        if not pagina:
            break

        print(f"  skip={skip}: {len(pagina)} obrigação(ões)")
        itens.extend(pagina)

        if len(pagina) < PAGE_SIZE:
            break
        skip += PAGE_SIZE

    return itens


def extrair_id_nome(item: dict) -> tuple[str, str | None]:
    """Os nomes reais de campo não são conhecidos até a primeira chamada real;
    tenta as variações mais prováveis e falha alto se nenhuma bater."""
    ident = item.get("id") or item.get("obligationId") or item.get("code")
    nome = item.get("name") or item.get("description") or item.get("nome")
    if ident is None:
        raise ValueError(f"Não encontrei identificador na obrigação: {item}")
    return str(ident), nome


def salvar_no_banco(itens: list[dict]) -> list[str]:
    if not DATABASE_URL:
        raise RuntimeError("Defina DATABASE_URL no .env antes de executar.")

    engine = create_engine(DATABASE_URL, echo=False)
    tabela = Table(TABLE_NAME, MetaData(), autoload_with=engine)

    novos, atualizados, ids_vistos = 0, 0, []

    for item in itens:
        nibo_obligation_id, nome = extrair_id_nome(item)
        ids_vistos.append(nibo_obligation_id)

        with engine.begin() as conn:
            existente = conn.execute(
                select(tabela.c.nibo_obligation_id).where(
                    tabela.c.nibo_obligation_id == nibo_obligation_id
                )
            ).fetchone()

            if existente is None:
                conn.execute(
                    tabela.insert().values(
                        nibo_obligation_id=nibo_obligation_id, nome=nome, payload=item
                    )
                )
                novos += 1
            else:
                conn.execute(
                    tabela.update()
                    .where(tabela.c.nibo_obligation_id == nibo_obligation_id)
                    .values(nome=nome, payload=item, sincronizado_em=func.now())
                )
                atualizados += 1

    print(f"\nCatálogo: {novos} novo(s), {atualizados} atualizado(s).")
    return ids_vistos


def relatorio_pendentes_no_depara(engine, ids_vistos: list[str]) -> None:
    """Lista, no catálogo recém-sincronizado, quais IDs ainda não têm
    correspondência em obrigacoes_depara -- é a fila de revisão manual."""
    if not ids_vistos:
        return

    with engine.connect() as conn:
        mapeados = set(
            conn.execute(text("select nibo_obligation_id from obrigacoes_depara")).scalars()
        )

    pendentes = [i for i in ids_vistos if i not in mapeados]
    print(f"\n{len(pendentes)} de {len(ids_vistos)} obrigação(ões) do catálogo sem de-para.")
    if pendentes:
        print(
            "Revise 'stg_nibo_obrigacoes_catalogo' para esses nibo_obligation_id "
            "e complete 'obrigacoes_depara' (ob_key, nome_amigavel, frequencia, "
            "mes_competencia) com critério humano antes de usar em qualquer perfil."
        )
        print("IDs pendentes:", ", ".join(pendentes))


def main():
    print("Buscando catálogo de obrigações no Nibo...")
    itens = buscar_catalogo_obrigacoes()
    print(f"\nTotal coletado: {len(itens)}")

    ids_vistos = salvar_no_banco(itens)

    engine = create_engine(DATABASE_URL, echo=False)
    relatorio_pendentes_no_depara(engine, ids_vistos)


if __name__ == "__main__":
    main()
