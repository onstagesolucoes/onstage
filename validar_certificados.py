import sys
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd  # já usado no NovaBusca.py
from cryptography.hazmat.primitives.serialization import pkcs12

# Reaproveita o que já existe no NovaBusca.py em vez de duplicar lógica
from NovaBusca import EXCEL_PATH, CERT_DIR, only_digits, encontrar_certificado_pkcs12

DIAS_ALERTA = 30  # a partir de quantos dias antes do vencimento mostrar aviso


def obter_validade_certificado(cert_path: str, senha: str | None):
    """
    Abre o .pfx/.p12 e retorna (not_before, not_after) em UTC.
    """
    with open(cert_path, "rb") as f:
        dados = f.read()

    senha_bytes = senha.encode("utf-8") if senha else None

    _, certificate, _ = pkcs12.load_key_and_certificates(dados, senha_bytes)

    if certificate is None:
        raise ValueError("Não foi possível extrair o certificado do arquivo PKCS#12.")

    # cryptography >= 42 expõe diretamente em UTC; fallback para versões antigas
    not_after = getattr(certificate, "not_valid_after_utc", None) or certificate.not_valid_after.replace(
        tzinfo=timezone.utc
    )
    not_before = getattr(certificate, "not_valid_before_utc", None) or certificate.not_valid_before.replace(
        tzinfo=timezone.utc
    )

    return not_before, not_after


def validar_certificado_empresa(nome_empresa: str, cnpj: str, senha: str | None):
    try:
        cert_path = encontrar_certificado_pkcs12(cnpj)
    except FileNotFoundError as e:
        return {
            "Nome empresa": nome_empresa,
            "cnpj": cnpj,
            "status": "CERTIFICADO NAO ENCONTRADO",
            "arquivo": None,
            "validade": None,
            "dias_restantes": None,
            "erro": str(e),
        }

    try:
        _, not_after = obter_validade_certificado(cert_path, senha)
    except Exception as e:
        return {
            "Nome empresa": nome_empresa,
            "cnpj": cnpj,
            "status": "ERRO AO ABRIR CERTIFICADO",
            "arquivo": cert_path,
            "validade": None,
            "dias_restantes": None,
            "erro": str(e),
        }

    agora = datetime.now(timezone.utc)
    dias_restantes = (not_after - agora).days

    if dias_restantes < 0:
        status = "VENCIDO"
    elif dias_restantes <= DIAS_ALERTA:
        status = "VENCE EM BREVE"
    else:
        status = "OK"

    return {
        "Nome empresa": nome_empresa,
        "cnpj": cnpj,
        "status": status,
        "arquivo": cert_path,
        "validade": not_after.strftime("%d/%m/%Y %H:%M:%S"),
        "dias_restantes": dias_restantes,
        "erro": None,
    }


def main():
    if not Path(EXCEL_PATH).exists():
        print(f"Arquivo Excel não encontrado: {EXCEL_PATH}")
        sys.exit(1)

    df = pd.read_excel(EXCEL_PATH, dtype={"cnpj": str})

    required_cols = {"Nome empresa", "cnpj", "senha"}
    missing = required_cols - set(df.columns)
    if missing:
        raise ValueError(f"Colunas ausentes no Excel: {missing}")

    resultados = []
    for _, row in df.iterrows():
        nome_empresa = str(row["Nome empresa"])
        cnpj = only_digits(str(row["cnpj"]))
        senha = str(row["senha"]) if str(row["senha"]).lower() != "nan" else ""

        resultado = validar_certificado_empresa(nome_empresa, cnpj, senha)
        resultados.append(resultado)

        linha = f"[{resultado['status']}] {nome_empresa} (CNPJ {cnpj})"
        if resultado["validade"]:
            linha += f" - vence em {resultado['validade']} ({resultado['dias_restantes']} dia(s))"
        if resultado["erro"]:
            linha += f" - {resultado['erro']}"
        print(linha)

    df_resultado = pd.DataFrame(resultados)
    saida = "validade_certificados.xlsx"
    df_resultado.to_excel(saida, index=False)
    print(f"\nRelatório salvo em: {saida}")


if __name__ == "__main__":
    main()
