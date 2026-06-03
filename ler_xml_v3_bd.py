import xml.etree.ElementTree as ET
import json
import os
from datetime import datetime, timezone

from dotenv import load_dotenv
from sqlalchemy import (
    create_engine, Column, String, Numeric, DateTime, Text, MetaData, Table, select
)

load_dotenv()

# =========================
# CONFIGURAÇÃO
# =========================

DATABASE_URL = os.getenv("DATABASE_URL")
TABLE_NAME   = os.getenv("TABLE_NAME", "notas_fiscais")

PATH_PY  = os.getcwd()
BASE_DIR = os.path.join(PATH_PY, "notasfiscais")
NS = {"nfse": "http://www.sped.fazenda.gov.br/nfse"}

COLUNAS_NUMERICAS = {
    "valor_servico", "base_calculo", "aliquota_iss", "valor_iss",
    "valor_liquido", "valor_total_retido", "valor_pis", "valor_cofins",
    "valor_irrf", "valor_csll",
}

# =========================
# FUNÇÕES AUXILIARES DE XML
# =========================

def get_text(root, *paths):
    for path in paths:
        el = root.find(path, NS)
        if el is not None and el.text:
            return el.text.strip()
    return None


def extrair_dados(root, clinica, ano_mes, tipo):
    dados = {
        "clinica":               clinica,
        "ano_mes":               ano_mes,
        "tipo":                  tipo,
        "chave_nfse":            None,
        "numero_nf":             get_text(root, ".//nfse:nNFSe"),
        "numero_dfse":           get_text(root, ".//nfse:nDFSe"),
        "data_emissao":          get_text(root, ".//nfse:dhProc", ".//nfse:dhEmi"),
        "data_competencia":      get_text(root, ".//nfse:dCompet"),
        "status":                get_text(root, ".//nfse:cStat"),
        "ambiente":              get_text(root, ".//nfse:ambGer"),
        "cnpj_emitente":         get_text(root, ".//nfse:emit/nfse:CNPJ", ".//nfse:prest/nfse:CNPJ"),
        "razao_emitente":        get_text(root, ".//nfse:emit/nfse:xNome", ".//nfse:prest/nfse:xNome"),
        "cidade_emitente":       get_text(root, ".//nfse:emit/nfse:enderNac/nfse:cMun"),
        "uf_emitente":           get_text(root, ".//nfse:emit/nfse:enderNac/nfse:UF"),
        "telefone_emitente":     get_text(root, ".//nfse:emit/nfse:fone"),
        "email_emitente":        get_text(root, ".//nfse:emit/nfse:email"),
        "cnpj_tomador":          get_text(root, ".//nfse:toma/nfse:CNPJ"),
        "nome_tomador":          get_text(root, ".//nfse:toma/nfse:xNome"),
        "cidade_tomador":        get_text(root, ".//nfse:toma/nfse:end/nfse:endNac/nfse:cMun"),
        "email_tomador":         get_text(root, ".//nfse:toma/nfse:email"),
        "codigo_servico":        get_text(root, ".//nfse:cTribNac"),
        "descricao_servico":     get_text(root, ".//nfse:xDescServ"),
        "info_complementar":     get_text(root, ".//nfse:xInfComp"),
        "valor_servico":         get_text(root, ".//nfse:vServ",      ".//nfse:DPS//nfse:vServ"),
        "base_calculo":          get_text(root, ".//nfse:vBC",        ".//nfse:valores/nfse:vBC"),
        "aliquota_iss":          get_text(root, ".//nfse:pAliqAplic", ".//nfse:trib//nfse:pAliq"),
        "valor_iss":             get_text(root, ".//nfse:vISSQN",     ".//nfse:valores/nfse:vISSQN"),
        "valor_liquido":         get_text(root, ".//nfse:vLiq"),
        "valor_total_retido":    get_text(root, ".//nfse:vTotalRet"),
        "valor_pis":             get_text(root, ".//nfse:vPis"),
        "valor_cofins":          get_text(root, ".//nfse:vCofins"),
        "valor_irrf":            get_text(root, ".//nfse:vRetIRRF"),
        "valor_csll":            get_text(root, ".//nfse:vRetCSLL"),
        "municipio_prestacao":   get_text(root, ".//nfse:cLocPrestacao"),
        "municipio_incidencia":  get_text(root, ".//nfse:cLocIncid"),
        "desc_local_incidencia": get_text(root, ".//nfse:xLocIncid"),
    }

    inf_nfse = root.find(".//nfse:infNFSe", NS)
    if inf_nfse is not None:
        dados["chave_nfse"] = inf_nfse.attrib.get("Id")

    return dados


# =========================
# BANCO DE DADOS
# =========================

def criar_tabela(engine):
    meta = MetaData()
    tabela = Table(
        TABLE_NAME, meta,
        # Chave primária: chave_nfse quando disponível, senão composta (clinica__tipo__numero_nf)
        Column("id",                    String(200), primary_key=True),
        Column("clinica",               String(200)),
        Column("ano_mes",               String(7)),
        Column("tipo",                  String(20)),
        Column("chave_nfse",            String(200)),
        Column("numero_nf",             String(30)),
        Column("numero_dfse",           String(30)),
        Column("data_emissao",          String(30)),
        Column("data_competencia",      String(30)),
        Column("status",                String(10)),
        Column("ambiente",              String(5)),
        Column("cnpj_emitente",         String(20)),
        Column("razao_emitente",        String(200)),
        Column("cidade_emitente",       String(10)),
        Column("uf_emitente",           String(2)),
        Column("telefone_emitente",     String(30)),
        Column("email_emitente",        String(150)),
        Column("cnpj_tomador",          String(20)),
        Column("nome_tomador",          String(200)),
        Column("cidade_tomador",        String(10)),
        Column("email_tomador",         String(150)),
        Column("codigo_servico",        String(10)),
        Column("descricao_servico",     Text),
        Column("info_complementar",     Text),
        Column("valor_servico",         Numeric(15, 2)),
        Column("base_calculo",          Numeric(15, 2)),
        Column("aliquota_iss",          Numeric(10, 4)),
        Column("valor_iss",             Numeric(15, 2)),
        Column("valor_liquido",         Numeric(15, 2)),
        Column("valor_total_retido",    Numeric(15, 2)),
        Column("valor_pis",             Numeric(15, 2)),
        Column("valor_cofins",          Numeric(15, 2)),
        Column("valor_irrf",            Numeric(15, 2)),
        Column("valor_csll",            Numeric(15, 2)),
        Column("municipio_prestacao",   String(10)),
        Column("municipio_incidencia",  String(10)),
        Column("desc_local_incidencia", String(200)),
        # Controle de carga
        Column("created_at",            DateTime),
        Column("updated_at",            DateTime),
    )
    meta.create_all(engine, checkfirst=True)
    return tabela


def gerar_id(dados):
    """Chave primária: chave_nfse se existir, senão clinica__tipo__numero_nf."""
    if dados.get("chave_nfse"):
        return dados["chave_nfse"]
    return f"{dados.get('clinica', '')}_{dados.get('tipo', '')}_{dados.get('numero_nf', '')}"


def normalizar(campo, valor):
    """Converte para tipo comparável; campos numéricos viram float."""
    if campo in COLUNAS_NUMERICAS:
        try:
            return float(valor) if valor is not None else None
        except (ValueError, TypeError):
            return None
    return valor


def processar_registro(conn, tabela, dados):
    id_nf = gerar_id(dados)
    agora = datetime.now(timezone.utc).replace(tzinfo=None)

    dados_norm = {campo: normalizar(campo, valor) for campo, valor in dados.items()}

    existente = conn.execute(
        select(tabela).where(tabela.c.id == id_nf)
    ).fetchone()

    if existente is None:
        conn.execute(tabela.insert().values(
            id=id_nf,
            created_at=agora,
            updated_at=agora,
            **dados_norm,
        ))
        return "inserido"

    valores_update = {}

    for campo, novo_valor in dados_norm.items():
        valor_atual = normalizar(campo, existente._mapping.get(campo))
        if valor_atual != novo_valor:
            valores_update[campo] = novo_valor

    if valores_update:
        valores_update["updated_at"] = agora
        conn.execute(
            tabela.update()
            .where(tabela.c.id == id_nf)
            .values(**valores_update)
        )
        return f"atualizado: {', '.join(valores_update.keys())}"

    return "sem alterações"


# =========================
# EXECUÇÃO PRINCIPAL
# =========================

if not DATABASE_URL:
    raise SystemExit("❌  DATABASE_URL não definido no .env")

engine = create_engine(DATABASE_URL, echo=False)
tabela = criar_tabela(engine)

contadores = {"inserido": 0, "atualizado": 0, "sem_alteracoes": 0, "erro": 0}
erros_log = []

for clinica in os.listdir(BASE_DIR):
    caminho_clinica = os.path.join(BASE_DIR, clinica)
    if not os.path.isdir(caminho_clinica):
        continue

    for ano_mes in os.listdir(caminho_clinica):
        caminho_ano_mes = os.path.join(caminho_clinica, ano_mes)
        if not os.path.isdir(caminho_ano_mes):
            continue

        for tipo in os.listdir(caminho_ano_mes):
            caminho_tipo = os.path.join(caminho_ano_mes, tipo)
            if not os.path.isdir(caminho_tipo):
                continue

            tipo_norm = tipo.lower()
            if tipo_norm not in ("prestados", "tomados"):
                print(f"[AVISO] Pasta ignorada (não é prestados/tomados): {caminho_tipo}")
                continue

            for arquivo in os.listdir(caminho_tipo):
                if not arquivo.endswith(".xml"):
                    continue

                caminho_xml = os.path.join(caminho_tipo, arquivo)
                try:
                    tree = ET.parse(caminho_xml)
                    root = tree.getroot()
                    dados = extrair_dados(root, clinica, ano_mes, tipo_norm)

                    # Cada arquivo em sua própria transação para não perder o progresso em caso de erro
                    with engine.begin() as conn:
                        resultado = processar_registro(conn, tabela, dados)

                    if resultado == "inserido":
                        contadores["inserido"] += 1
                    elif resultado.startswith("atualizado"):
                        contadores["atualizado"] += 1
                    else:
                        contadores["sem_alteracoes"] += 1

                except Exception as e:
                    contadores["erro"] += 1
                    erros_log.append({"arquivo": caminho_xml, "erro": str(e)})
                    print(f"[ERRO] {caminho_xml}: {e}")

print(f"\nConcluído:")
print(f"  Inseridos:      {contadores['inserido']}")
print(f"  Atualizados:    {contadores['atualizado']}")
print(f"  Sem alterações: {contadores['sem_alteracoes']}")
print(f"  Erros:          {contadores['erro']}")

if erros_log:
    log_path = os.path.join(PATH_PY, f"erros_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json")
    with open(log_path, "w", encoding="utf-8") as f:
        json.dump(erros_log, f, ensure_ascii=False, indent=2)
    print(f"\n  Log de erros salvo em: {log_path}")
