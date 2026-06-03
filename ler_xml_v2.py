import xml.etree.ElementTree as ET
import os
import pandas as pd
from datetime import datetime

# =========================
# CONFIGURAÇÃO
# =========================

PATH_PY = os.getcwd()
BASE_DIR = os.path.join(PATH_PY, "notasfiscais")

NS = {"nfse": "http://www.sped.fazenda.gov.br/nfse"}

# Campos obrigatórios para considerar a NF válida
CAMPOS_OBRIGATORIOS = ["cnpj_emitente", "valor_servico", "data_emissao"]

# =========================
# FUNÇÕES AUXILIARES
# =========================

def get_text(root, *paths):
    """Tenta múltiplos caminhos XPath e retorna o texto do primeiro que encontrar."""
    for path in paths:
        el = root.find(path, NS)
        if el is not None and el.text:
            return el.text.strip()
    return None


def extrair_dados(root, clinica, ano_mes, tipo):
    """Extrai todos os campos relevantes de um XML de NFS-e."""

    dados = {
        # Metadados de origem
        "clinica":              clinica,
        "ano_mes":              ano_mes,
        "tipo":                 tipo,  # "prestados" ou "tomados"

        # Identificação da nota
        "chave_nfse":           None,
        "numero_nf":            get_text(root, ".//nfse:nNFSe"),
        "numero_dfse":          get_text(root, ".//nfse:nDFSe"),
        "data_emissao":         get_text(root, ".//nfse:dhProc", ".//nfse:dhEmi"),
        "data_competencia":     get_text(root, ".//nfse:dCompet"),
        "status":               get_text(root, ".//nfse:cStat"),
        "ambiente":             get_text(root, ".//nfse:ambGer"),

        # Emitente
        "cnpj_emitente":        get_text(root, ".//nfse:emit/nfse:CNPJ", ".//nfse:prest/nfse:CNPJ"),
        "razao_emitente":       get_text(root, ".//nfse:emit/nfse:xNome", ".//nfse:prest/nfse:xNome"),
        "cidade_emitente":      get_text(root, ".//nfse:emit/nfse:enderNac/nfse:cMun"),
        "uf_emitente":          get_text(root, ".//nfse:emit/nfse:enderNac/nfse:UF"),
        "telefone_emitente":    get_text(root, ".//nfse:emit/nfse:fone"),
        "email_emitente":       get_text(root, ".//nfse:emit/nfse:email"),

        # Tomador
        "cnpj_tomador":         get_text(root, ".//nfse:toma/nfse:CNPJ"),
        "nome_tomador":         get_text(root, ".//nfse:toma/nfse:xNome"),
        "cidade_tomador":       get_text(root, ".//nfse:toma/nfse:end/nfse:endNac/nfse:cMun"),
        "email_tomador":        get_text(root, ".//nfse:toma/nfse:email"),

        # Serviço
        "codigo_servico":       get_text(root, ".//nfse:cTribNac"),
        "descricao_servico":    get_text(root, ".//nfse:xDescServ"),
        "info_complementar":    get_text(root, ".//nfse:xInfComp"),

        # Valores — tenta caminho direto (versão recente) e dentro do DPS (versão 2023)
        "valor_servico":        get_text(root, ".//nfse:vServ",       ".//nfse:DPS//nfse:vServ"),
        "base_calculo":         get_text(root, ".//nfse:vBC",         ".//nfse:valores/nfse:vBC"),
        "aliquota_iss":         get_text(root, ".//nfse:pAliqAplic",  ".//nfse:trib//nfse:pAliq"),
        "valor_iss":            get_text(root, ".//nfse:vISSQN",      ".//nfse:valores/nfse:vISSQN"),
        "valor_liquido":        get_text(root, ".//nfse:vLiq"),
        "valor_total_retido":   get_text(root, ".//nfse:vTotalRet"),

        # Retenções federais (podem estar ausentes em NFs mais antigas)
        "valor_pis":            get_text(root, ".//nfse:vPis"),
        "valor_cofins":         get_text(root, ".//nfse:vCofins"),
        "valor_irrf":           get_text(root, ".//nfse:vRetIRRF"),
        "valor_csll":           get_text(root, ".//nfse:vRetCSLL"),

        # Localização
        "municipio_prestacao":  get_text(root, ".//nfse:cLocPrestacao"),
        "municipio_incidencia": get_text(root, ".//nfse:cLocIncid"),
        "desc_local_incidencia":get_text(root, ".//nfse:xLocIncid"),
    }

    # Chave da NFSe vem como atributo, não como texto
    inf_nfse = root.find(".//nfse:infNFSe", NS)
    if inf_nfse is not None:
        dados["chave_nfse"] = inf_nfse.attrib.get("Id")

    return dados


# =========================
# ITERAÇÃO NAS PASTAS
# =========================

registros = []
erros = []
falhas = []
chaves_canceladas = set()


def tag_local(element):
    """Retorna o nome da tag sem namespace."""
    tag = element.tag
    return tag.split("}")[-1] if "}" in tag else tag


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

            # Normaliza para "prestados" ou "tomados" (case-insensitive)
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

                    # Evento de cancelamento — coleta a chave e ignora como nota
                    if tag_local(root) == "evento":
                        if root.find(".//nfse:e101101", NS) is not None:
                            ch = get_text(root, ".//nfse:chNFSe")
                            if ch:
                                chaves_canceladas.add("NFS" + ch)
                        continue

                    dados = extrair_dados(root, clinica, ano_mes, tipo_norm)

                    campos_vazios = [c for c in CAMPOS_OBRIGATORIOS if not dados.get(c)]
                    if campos_vazios:
                        falhas.append({
                            "arquivo": caminho_xml,
                            "clinica": clinica,
                            "ano_mes": ano_mes,
                            "tipo": tipo_norm,
                            "motivo": "campos obrigatórios ausentes",
                            "campos_vazios": ", ".join(campos_vazios),
                        })
                        print(f"[INCOMPLETO] {caminho_xml} — campos vazios: {campos_vazios}")
                    else:
                        registros.append(dados)

                except Exception as e:
                    falhas.append({
                        "arquivo": caminho_xml,
                        "clinica": clinica,
                        "ano_mes": ano_mes,
                        "tipo": tipo_norm,
                        "motivo": "erro de leitura XML",
                        "campos_vazios": "",
                    })
                    erros.append({"arquivo": caminho_xml, "erro": str(e)})
                    print(f"[ERRO] {caminho_xml}: {e}")


# =========================
# RESULTADO
# =========================

df = pd.DataFrame(registros)

# Marca notas canceladas cruzando com os eventos de cancelamento coletados
if chaves_canceladas and "chave_nfse" in df.columns:
    mask = df["chave_nfse"].isin(chaves_canceladas)
    df.loc[mask, "status"] = "CANCELADA"
    print(f"🚫 {mask.sum()} nota(s) marcada(s) como CANCELADA com base nos eventos.")

# Converte colunas de valor/alíquota para numérico (ponto como separador decimal)
colunas_numericas = [
    "valor_servico", "base_calculo", "aliquota_iss", "valor_iss",
    "valor_liquido", "valor_total_retido", "valor_pis", "valor_cofins",
    "valor_irrf", "valor_csll",
]
for col in colunas_numericas:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")

print(f"\n✅ {len(df)} notas processadas com sucesso.")
if falhas:
    print(f"⚠️  {len(falhas)} arquivo(s) com falha ({sum(1 for f in falhas if f['motivo'] == 'erro de leitura XML')} erro(s) de XML, {sum(1 for f in falhas if f['motivo'] == 'campos obrigatórios ausentes')} incompleto(s)).")

print(df.head())

# Salva consolidado em Excel
output_path = os.path.join(PATH_PY, "notas_fiscais_consolidado.xlsx")
df.to_excel(output_path, index=False)
print(f"\n📁 Consolidado salvo em: {output_path}")

# Gera relatório de falhas
if falhas:
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    report_path = os.path.join(PATH_PY, f"report_falhas_{timestamp}.xlsx")
    df_falhas = pd.DataFrame(falhas)
    df_falhas.to_excel(report_path, index=False)
    print(f"📋 Relatório de falhas salvo em: {report_path}")