import pandas as pd
import re
import unicodedata
from pathlib import Path

# ============================================================
# CONFIGURAÇÃO
# ============================================================

# Pasta onde a planilha de origem será colocada.
PASTA_DADOS = Path(__file__).resolve().parent / "Dados"

# Nome do arquivo final.
ARQUIVO_SAIDA = "RELAÇÃO_DE_PLACAS_PADRONIZADA.xlsx"

# Colunas que serão mantidas.
COLUNAS_NECESSARIAS = [
    "NOME DO TRANSPORTADOR",
    "TIPO DO VEÍCULO",
    "PLACA"
]

# ============================================================
# REGRAS DE PADRONIZAÇÃO
# ============================================================

REGRAS_TRANSPORTADOR = [
    # Específicas primeiro
    ("JET EXPRESS YUNYI", "JET EXPRESS (YUNYI)"),
    ("SOUSA CHIC", "TRANS CHIC"),

    # Transportadores
    ("PACHECO", "PACHECO E PACHECO"),
    ("CEOS", "CEOS EXPRESS"),
    ("EVEZEN", "EVEZEN LTDA"),
    ("FS VLMR", "FS VLMR"),
    ("SAMY", "Samy Transportes Ltda"),
    ("EASY CARGO", "EASY CARGO SOLUÇOES"),
    # FAP é uma sigla específica; não usar apenas "FAP" como
    # substring para evitar falsos positivos.
    ("FAP", "Fernando Augusto Pires LTDA"),
    ("FERNANDO AUGUSTO PIRES", "Fernando Augusto Pires LTDA"),
    ("DC FLEX", "DC FLEX"),
    ("3ZX", "3ZX Transportes"),
    ("LALALOG", "LALALOG EXPRESS LTDA"),
    ("MULLER ENXOVAIS", "MULLER ENXOVAIS LTDA"),
    ("DEBORA", "DEBORA TRANSPORTES"),
    ("BBR AGROLOG", "BBR Agrolog Ltda"),
    ("VLX", "VLX LOGISTICA LTDA"),
    ("GF LOGISTICA", "GF LOGISTICA LTDA"),
    ("MAXXLOG", "MAXXLOG LOGISTICA"),
    ("JET EXPRESS", "JET EXPRESS"),
    ("GTI EXPRESS", "GTI EXPRESS"),
    ("TGA", "TGA TRANSPORTES"),
    ("JT EXPRESS", "JT EXPRESS BRAZIL"),
    ("FLIGHTCARGO", "FLIGHTCARGO TRANSPORTES"),
    ("TRANS CHIC", "TRANS CHIC"),
    ("LDJ", "LDJ TRANSPORTES LTDS"),
    ("FRANQUIA JET", "FRANQUIA JET"),
    ("F S-MOOCA", "F S-MOOCA-SP"),
    ("FSMOCA", "F S-MOOCA-SP"),
    ("JOSIAS VITORINO", "JOSIAS VITORINO DO NASCIMENTO"),
    ("BLIXX", "BLIXX INVESTIMENTOS"),
    ("FSMGUE02SP", "FSMGUE02SP"),
    ("JND", "JND TRANSPORTES LTDA"),
    ("SAMUEL MESSIAS", "SAMUEL MESSIAS DE ANDRADE"),
    ("LUIS FERNANDO", "LUIS FERNANDO DE MOURA"),
    ("PLLUGO", "PLLUGO"),
    ("YANES", "YANES LOGÍSTICA LTDA"),
    ("CSM LOGISTICA", "CSM LOGISTICA E TRANSPORTES"),
    ("LUIS MENEZES", "LUIS MENEZES OLIVEIRA"),
    ("CAPITAO EXPRESS", "CAPITAO EXPRESS"),
    ("SAO JOSE", "São José transportes"),
    ("WILLIAN PAULO", "WILLIAN PAULO VIANA DE BRITTO"),
    ("F DDM", "F DDM 02-SP"),
]


# ============================================================
# FUNÇÕES
# ============================================================

def remover_acentos(texto):
    texto = unicodedata.normalize("NFKD", texto)
    return "".join(
        c for c in texto
        if not unicodedata.combining(c)
    )


def limpar_texto(texto):
    """
    Remove:
    - espaços no início/fim;
    - espaços duplicados;
    - quebras de linha.
    """
    if pd.isna(texto):
        return ""

    texto = str(texto)
    texto = texto.replace("\n", " ").replace("\r", " ")
    texto = re.sub(r"\s+", " ", texto)

    return texto.strip()


def chave_comparacao(texto):
    """
    Cria uma versão para comparação:
    - remove acentos;
    - coloca em maiúsculas;
    - corrige espaços.
    """
    texto = limpar_texto(texto)
    texto = remover_acentos(texto)

    return texto.upper()


def padronizar_transportador(nome):
    """
    Procura uma palavra-chave dentro do nome.

    Exemplo:
        PACHECO TRANSPORTES LTDA
        PACHECO & PACHECO
        PACHECO E PACHECO TRANSPORTE

    Todos viram:
        PACHECO E PACHECO

    Se nenhuma regra for encontrada, o nome original é mantido,
    apenas com os espaços corrigidos.
    """

    nome = limpar_texto(nome)

    if not nome:
        return ""

    nome_busca = chave_comparacao(nome)

    for palavra_chave, nome_padrao in REGRAS_TRANSPORTADOR:

        palavra_busca = chave_comparacao(palavra_chave)

        # Siglas curtas (até 3 caracteres) devem ser palavras
        # independentes para evitar falsos positivos.
        if len(palavra_busca) <= 3:
            padrao = rf"(?<![A-Z0-9]){re.escape(palavra_busca)}(?![A-Z0-9])"

            if re.search(padrao, nome_busca):
                return nome_padrao

        # Para nomes/palavras maiores, mantém a busca por conteúdo.
        elif palavra_busca in nome_busca:
            return nome_padrao

    return nome


def encontrar_arquivo_entrada():
    """
    Procura arquivos .xlsx dentro da pasta configurada.

    O arquivo de saída é ignorado.

    Se houver mais de uma fonte, o script para para evitar
    processar a planilha errada.
    """

    if not PASTA_DADOS.exists():
        raise FileNotFoundError(
            f"\nA pasta não foi encontrada:\n"
            f"{PASTA_DADOS}\n\n"
            "Verifique se o caminho está correto."
        )

    arquivos_excel = []

    for arquivo in PASTA_DADOS.glob("*.xlsx"):

        # Ignora o arquivo que o próprio script gera.
        if arquivo.name.lower() == ARQUIVO_SAIDA.lower():
            continue

        # Ignora arquivos temporários do Excel.
        if arquivo.name.startswith("~$"):
            continue

        arquivos_excel.append(arquivo)

    if not arquivos_excel:
        raise FileNotFoundError(
            f"\nNenhuma planilha .xlsx encontrada em:\n"
            f"{PASTA_DADOS}\n\n"
            "Coloque a planilha de origem nessa pasta."
        )

    if len(arquivos_excel) > 1:

        nomes = "\n".join(
            f"  - {arquivo.name}"
            for arquivo in arquivos_excel
        )

        raise ValueError(
            "\nForam encontradas várias planilhas na pasta:\n\n"
            f"{nomes}\n\n"
            "Deixe somente a planilha de origem na pasta "
            "e execute novamente."
        )

    return arquivos_excel[0]


# ============================================================
# INÍCIO
# ============================================================

print("=" * 60)
print("PADRONIZAÇÃO DE RELAÇÃO DE PLACAS")
print("=" * 60)

print(f"\nPasta de origem:")
print(f"  {PASTA_DADOS}")

arquivo_entrada = encontrar_arquivo_entrada()

print(f"\nArquivo encontrado:")
print(f"  {arquivo_entrada.name}")

print("\nLendo planilha...")

df = pd.read_excel(arquivo_entrada)


# ============================================================
# VALIDAR COLUNAS
# ============================================================

colunas_faltantes = [
    coluna
    for coluna in COLUNAS_NECESSARIAS
    if coluna not in df.columns
]

if colunas_faltantes:

    raise ValueError(
        "\nAs seguintes colunas não foram encontradas:\n"
        + "\n".join(
            f"  - {coluna}"
            for coluna in colunas_faltantes
        )
        + "\n\nColunas encontradas na planilha:\n"
        + "\n".join(
            f"  - {coluna}"
            for coluna in df.columns
        )
    )


# ============================================================
# MANTER SOMENTE AS 3 COLUNAS
# ============================================================

df = df[COLUNAS_NECESSARIAS].copy()


# ============================================================
# LIMPAR AS COLUNAS
# ============================================================

for coluna in COLUNAS_NECESSARIAS:
    df[coluna] = df[coluna].apply(limpar_texto)


# ============================================================
# PADRONIZAR TRANSPORTADOR
# ============================================================

df["NOME DO TRANSPORTADOR"] = (
    df["NOME DO TRANSPORTADOR"]
    .apply(padronizar_transportador)
)


# ============================================================
# PADRONIZAR PLACA
# ============================================================
#
# Aqui todos os espaços são removidos.
#
# ABC 1D23 -> ABC1D23
# abc 1234 -> ABC1234
# ============================================================

df["PLACA"] = (
    df["PLACA"]
    .astype(str)
    .str.replace(r"\s+", "", regex=True)
    .str.upper()
)

df["PLACA"] = df["PLACA"].replace("NAN", "")


# ============================================================
# SALVAR
# ============================================================

arquivo_saida = PASTA_DADOS / ARQUIVO_SAIDA

df.to_excel(
    arquivo_saida,
    index=False
)


# ============================================================
# RESULTADO
# ============================================================

print("\n" + "=" * 60)
print("PROCESSAMENTO CONCLUÍDO COM SUCESSO")
print("=" * 60)

print(f"\nArquivo de origem:")
print(f"  {arquivo_entrada}")

print(f"\nArquivo gerado:")
print(f"  {arquivo_saida}")

print(f"\nLinhas processadas: {len(df)}")

print("\nColunas mantidas:")
for coluna in COLUNAS_NECESSARIAS:
    print(f"  - {coluna}")

print("\nTratamentos realizados:")
print("  ✓ Espaços duplicados removidos")
print("  ✓ Espaços no início/fim removidos")
print("  ✓ Quebras de linha removidas")
print("  ✓ Transportadores padronizados por palavra-chave e siglas")
print("  ✓ Placas sem espaços")
print("  ✓ Placas convertidas para maiúsculas")
print("  ✓ Transportadores não reconhecidos mantidos")

print("\nTudo pronto!")
