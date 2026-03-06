import xml.etree.ElementTree as ET
import os

dir = r"C:\Users\tiago\OneDrive\Documentos\Python\onstage\notasfiscais\MEDF Serviços Médicos\2026-03\prestados"

tree = ET.parse(rf"{dir}\31693562249899508000162000000000001626032170572186.xml")

root = tree.getroot()

ns = {"nfse": "http://www.sped.fazenda.gov.br/nfse"}


def get_text(path):
    el = root.find(path, ns)
    return el.text if el is not None else None


# =========================
# IDENTIFICAÇÃO DA NOTA
# =========================

chave_nfse = root.find(".//nfse:infNFSe", ns).attrib.get("Id")
numero_nf = get_text(".//nfse:nNFSe")
numero_dfse = get_text(".//nfse:nDFSe")
data_emissao = get_text(".//nfse:dhProc")
status = get_text(".//nfse:cStat")
ambiente = get_text(".//nfse:ambGer")

# =========================
# EMITENTE
# =========================

cnpj_emitente = get_text(".//nfse:emit/nfse:CNPJ")
razao_emitente = get_text(".//nfse:emit/nfse:xNome")
cidade_emitente = get_text(".//nfse:emit/nfse:enderNac/nfse:cMun")
uf_emitente = get_text(".//nfse:emit/nfse:enderNac/nfse:UF")
telefone_emitente = get_text(".//nfse:emit/nfse:fone")
email_emitente = get_text(".//nfse:emit/nfse:email")

# =========================
# TOMADOR
# =========================

cnpj_tomador = get_text(".//nfse:toma/nfse:CNPJ")
nome_tomador = get_text(".//nfse:toma/nfse:xNome")
cidade_tomador = get_text(".//nfse:toma/nfse:end/nfse:endNac/nfse:cMun")
email_tomador = get_text(".//nfse:toma/nfse:email")

# =========================
# SERVIÇO
# =========================

codigo_servico_nacional = get_text(".//nfse:cTribNac")
descricao_servico = get_text(".//nfse:xDescServ")
informacao_complementar = get_text(".//nfse:xInfComp")

# =========================
# VALORES
# =========================

valor_servico = get_text(".//nfse:vServ")
base_calculo = get_text(".//nfse:vBC")
aliquota_iss = get_text(".//nfse:pAliqAplic")
valor_iss = get_text(".//nfse:vISSQN")
valor_liquido = get_text(".//nfse:vLiq")
valor_total_retido = get_text(".//nfse:vTotalRet")

# =========================
# RETENÇÕES FEDERAIS
# =========================

valor_pis = get_text(".//nfse:vPis")
valor_cofins = get_text(".//nfse:vCofins")
valor_irrf = get_text(".//nfse:vRetIRRF")
valor_csll = get_text(".//nfse:vRetCSLL")

# =========================
# LOCALIZAÇÃO
# =========================

municipio_prestacao = get_text(".//nfse:cLocPrestacao")
municipio_incidencia = get_text(".//nfse:cLocIncid")
descricao_local_incidencia = get_text(".//nfse:xLocIncid")

# =========================
# PRINT
# =========================

print("Chave NFSe:", chave_nfse)
print("Número NF:", numero_nf)
print("Número DFSe:", numero_dfse)
print("Data emissão:", data_emissao)
print("Status:", status)
print("Ambiente:", ambiente)

print("\nEmitente")
print("CNPJ:", cnpj_emitente)
print("Razão:", razao_emitente)
print("Cidade:", cidade_emitente)
print("UF:", uf_emitente)
print("Telefone:", telefone_emitente)
print("Email:", email_emitente)

print("\nTomador")
print("CNPJ:", cnpj_tomador)
print("Nome:", nome_tomador)
print("Cidade:", cidade_tomador)
print("Email:", email_tomador)

print("\nServiço")
print("Código serviço:", codigo_servico_nacional)
print("Descrição:", descricao_servico)
print("Info complementar:", informacao_complementar)

print("\nValores")
print("Valor serviço:", valor_servico)
print("Base cálculo:", base_calculo)
print("Alíquota ISS:", aliquota_iss)
print("ISS:", valor_iss)
print("Valor líquido:", valor_liquido)
print("Total retido:", valor_total_retido)

print("\nRetenções federais")
print("PIS:", valor_pis)
print("COFINS:", valor_cofins)
print("IRRF:", valor_irrf)
print("CSLL:", valor_csll)

print("\nLocalização")
print("Município prestação:", municipio_prestacao)
print("Município incidência:", municipio_incidencia)
print("Descrição incidência:", descricao_local_incidencia)