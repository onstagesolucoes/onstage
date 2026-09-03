"""
Revisão cruzada do fechamento mensal, usando o perfil fiscal do cliente como
elo entre Nibo, ClickUp e OneFlow.

Para cada cliente ativo, `clientes.perfil_id` aponta para `perfis.codigos`, a
lista de obrigações (ob_key) que aquele perfil deve entregar todo mês (mais as
anuais/trimestrais que caem na competência). Esta é a régua contra a qual as
três pontas já alimentadas no Supabase são comparadas:

  Nibo     -> calendario_obrigacoes  (alimentado pelo workflow FISCAL 01)
  ClickUp  -> fechamento_clickup     (alimentado pelos workflows FISCAL 00/02)
  OneFlow  -> entregas               (alimentado pelo workflow NÚCLEO)

O resultado -- o que bate e o que não bate com o perfil -- é gravado em
'diagnostico_fechamento_mensal', uma linha por cliente por competência
(upsert por cnpj+competencia). Essa tabela já existe no schema do projeto,
criada para justamente esse diagnóstico, mas nunca foi populada.

Este script não chama nenhuma API externa: revisa o que as automações
existentes já sincronizaram no Supabase. Não substitui o FISCAL 01 (que é
quem de fato consulta o Nibo e grava o calendário) -- ele é o consumidor
seguinte, que audita o resultado.

Setup necessário (apenas uma vez):
  DATABASE_URL no .env (mesmo banco usado pelas outras rotinas)

Uso:
  python revisar_fechamento_obrigacoes.py [AAAA-MM]
  Sem argumento, revisa a competência do mês apurado (mês anterior ao atual).
"""

import json
import os
import sys
from datetime import date

from dotenv import load_dotenv
from sqlalchemy import create_engine, text

load_dotenv()
DATABASE_URL = os.getenv("DATABASE_URL")

# Perfis que existem no cadastro mas ficam fora da automação de obrigações
# fiscais (ex.: cliente atendido só para comissão) -- não fazem sentido nesta
# revisão porque não têm codigos associados em `perfis`.
PERFIS_FORA_DA_AUTOMACAO = {"COMISSAO_NIBO"}


def competencia_padrao() -> str:
    hoje = date.today()
    ano, mes = hoje.year, hoje.month - 1
    if mes == 0:
        ano, mes = ano - 1, 12
    return f"{ano:04d}-{mes:02d}"


def mes_da_competencia(competencia: str) -> int:
    return int(competencia.split("-")[1])


def carregar_perfis(conn) -> dict[str, list[str]]:
    linhas = conn.execute(text("select perfil_id, codigos from perfis")).mappings().all()
    return {
        l["perfil_id"]: [c.strip() for c in (l["codigos"] or "").split(",") if c.strip()]
        for l in linhas
    }


def carregar_depara(conn) -> dict[str, dict]:
    linhas = conn.execute(
        text("select ob_key, frequencia, mes_competencia from obrigacoes_depara where ativo")
    ).mappings().all()
    return {l["ob_key"]: dict(l) for l in linhas}


def obrigacoes_esperadas(ob_keys: list[str], depara: dict, mes: int) -> list[str]:
    """Filtra a lista de codigos do perfil pela frequência: mensais sempre
    entram; anuais/trimestrais só entram se `mes_competencia` bater com a
    competência revisada. Um ob_key sem de-para entra mesmo assim (não dá
    para saber a frequência), mas isso é sinalizado à parte."""
    esperadas = []
    for ob_key in ob_keys:
        info = depara.get(ob_key)
        if info is None or info["mes_competencia"] is None or info["mes_competencia"] == mes:
            esperadas.append(ob_key)
    return esperadas


def carregar_clientes(conn) -> list[dict]:
    linhas = conn.execute(
        text(
            "select cnpj, nome_cliente, perfil_id from clientes "
            "where ativo and perfil_id is not null "
            "and not (perfil_id = any(:excluidos))"
        ),
        {"excluidos": list(PERFIS_FORA_DA_AUTOMACAO)},
    ).mappings().all()
    return [dict(l) for l in linhas]


def carregar_por_cnpj(conn, tabela: str, colunas: str, competencia: str) -> dict[str, dict[str, dict]]:
    linhas = conn.execute(
        text(f"select {colunas} from {tabela} where competencia = :comp"),
        {"comp": competencia},
    ).mappings().all()
    mapa: dict[str, dict[str, dict]] = {}
    for l in linhas:
        mapa.setdefault(l["cnpj"], {})[l["ob_key"]] = dict(l)
    return mapa


def revisar_cliente(
    cliente: dict,
    ob_keys_esperadas: list[str],
    depara: dict,
    calendario: dict,
    clickup: dict,
    entregas: dict,
) -> dict:
    cnpj = cliente["cnpj"]
    cal_cliente = calendario.get(cnpj, {})
    click_cliente = clickup.get(cnpj, {})
    entrega_cliente = entregas.get(cnpj, {})

    sem_depara = [k for k in ob_keys_esperadas if k not in depara]
    faltando_nibo = [k for k in ob_keys_esperadas if k not in cal_cliente]
    faltando_clickup = [k for k in ob_keys_esperadas if k not in click_cliente]
    fora_do_perfil_clickup = [k for k in click_cliente if k not in ob_keys_esperadas]
    faltando_oneflow = [
        k for k in ob_keys_esperadas if not entrega_cliente.get(k, {}).get("protocolado_nibo")
    ]

    nibo_ok = not faltando_nibo
    clickup_ok = not faltando_clickup and not fora_do_perfil_clickup
    oneflow_ok = not faltando_oneflow

    pendencias = {}
    if sem_depara:
        pendencias["perfil_com_ob_key_sem_depara"] = sem_depara
    if faltando_nibo:
        pendencias["nibo_sem_vencimento"] = faltando_nibo
    if faltando_clickup:
        pendencias["clickup_sem_tarefa"] = faltando_clickup
    if fora_do_perfil_clickup:
        pendencias["clickup_fora_do_perfil"] = fora_do_perfil_clickup
    if faltando_oneflow:
        pendencias["oneflow_nao_protocolado"] = faltando_oneflow

    return {
        "cnpj": cnpj,
        "nome_cliente": cliente["nome_cliente"],
        "nibo_ok": nibo_ok,
        "nibo_detalhe": "ok" if nibo_ok else f"{len(faltando_nibo)} obrigação(ões) sem vencimento no calendário",
        "clickup_ok": clickup_ok,
        "clickup_detalhe": (
            "ok"
            if clickup_ok
            else f"{len(faltando_clickup)} sem tarefa, {len(fora_do_perfil_clickup)} fora do perfil"
        ),
        "oneflow_ok": oneflow_ok,
        "oneflow_detalhe": "ok" if oneflow_ok else f"{len(faltando_oneflow)} obrigação(ões) não protocolada(s) no Nibo",
        "tem_pendencia": bool(pendencias),
        "pendencias": pendencias,
    }


UPSERT_SQL = text(
    """
    insert into diagnostico_fechamento_mensal
        (cnpj, competencia, nome_cliente, nibo_ok, nibo_detalhe,
         clickup_ok, clickup_detalhe, oneflow_ok, oneflow_detalhe,
         tem_pendencia, pendencias, gerado_em)
    values
        (:cnpj, :competencia, :nome_cliente, :nibo_ok, :nibo_detalhe,
         :clickup_ok, :clickup_detalhe, :oneflow_ok, :oneflow_detalhe,
         :tem_pendencia, :pendencias, now())
    on conflict (cnpj, competencia) do update set
        nome_cliente = excluded.nome_cliente,
        nibo_ok = excluded.nibo_ok,
        nibo_detalhe = excluded.nibo_detalhe,
        clickup_ok = excluded.clickup_ok,
        clickup_detalhe = excluded.clickup_detalhe,
        oneflow_ok = excluded.oneflow_ok,
        oneflow_detalhe = excluded.oneflow_detalhe,
        tem_pendencia = excluded.tem_pendencia,
        pendencias = excluded.pendencias,
        gerado_em = now()
    """
)


def salvar_diagnostico(conn, resultado: dict, competencia: str) -> None:
    conn.execute(
        UPSERT_SQL,
        {
            **resultado,
            "competencia": competencia,
            "pendencias": json.dumps(resultado["pendencias"], ensure_ascii=False),
        },
    )


def main():
    if not DATABASE_URL:
        raise RuntimeError("Defina DATABASE_URL no .env antes de executar.")

    competencia = sys.argv[1] if len(sys.argv) > 1 else competencia_padrao()
    mes = mes_da_competencia(competencia)
    print(f"Revisando competência {competencia}...")

    engine = create_engine(DATABASE_URL, echo=False)

    with engine.connect() as conn:
        perfis = carregar_perfis(conn)
        depara = carregar_depara(conn)
        clientes = carregar_clientes(conn)
        calendario = carregar_por_cnpj(
            conn, "calendario_obrigacoes", "cnpj, ob_key, data_vencimento", f"{competencia}-01"
        )
        clickup = carregar_por_cnpj(conn, "fechamento_clickup", "cnpj, ob_key, clickup_task_id", competencia)
        entregas = carregar_por_cnpj(conn, "entregas", "cnpj, ob_key, protocolado_nibo", competencia)

    print(f"{len(clientes)} cliente(s) ativo(s) a revisar.")

    resultados, com_pendencia = [], 0
    for cliente in clientes:
        ob_keys = perfis.get(cliente["perfil_id"], [])
        if not ob_keys:
            print(f"  aviso: perfil '{cliente['perfil_id']}' sem codigos ({cliente['cnpj']})")
            continue

        esperadas = obrigacoes_esperadas(ob_keys, depara, mes)
        resultado = revisar_cliente(cliente, esperadas, depara, calendario, clickup, entregas)
        resultados.append(resultado)
        if resultado["tem_pendencia"]:
            com_pendencia += 1

    with engine.begin() as conn:
        for resultado in resultados:
            salvar_diagnostico(conn, resultado, competencia)

    print(
        f"\nDiagnóstico gravado: {len(resultados)} cliente(s), "
        f"{com_pendencia} com pendência, {len(resultados) - com_pendencia} ok em tudo."
    )


if __name__ == "__main__":
    main()
