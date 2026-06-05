#!/usr/bin/env python3
"""
xlsx_to_yml.py — Converte publications.xlsx em publications.yml
para uso como fonte de dados de uma listagem customizada no Quarto.

Uso:
    python xlsx_to_yml.py
        Usa os caminhos padrão definidos nas constantes DEFAULT_INPUT e DEFAULT_OUTPUT.

    python xlsx_to_yml.py input.xlsx output.yml
        Permite informar manualmente o arquivo de entrada e o arquivo de saída.

    python xlsx_to_yml.py --force
        Força a conversão mesmo que o arquivo de saída já exista e esteja atualizado.

Requisitos:
    pip install openpyxl

Observação:
    Se o pacote PyYAML estiver instalado, o script também faz uma validação
    simples do YAML gerado ao final da conversão.
"""

import sys
from pathlib import Path
import openpyxl


# =============================================================================
# Caminhos padrão
# =============================================================================
# Estes caminhos são usados quando o script é executado sem argumentos.
# Exemplo:
#   python xlsx_to_yml.py
# Nesse caso, o script procura "publications.xlsx" e gera "publications.yml".
DEFAULT_INPUT = "publications.xlsx"
DEFAULT_OUTPUT = "publications.yml"


# =============================================================================
# Mapeamento: nome da coluna no Excel -> nome da chave no YAML
# =============================================================================
# Esta tabela traduz os nomes mais "humanos" da planilha para as chaves usadas
# pelo template de listagem do Quarto.
#
# Exemplo:
#   coluna Excel: "Paper Link"
#   chave YAML:   "path"
#
# Se uma coluna não estiver aqui, o script tenta criar automaticamente uma chave
# a partir do nome da coluna, convertendo para minúsculas e trocando espaços por "-".
KEY_MAP = {
    "Section": "section",
    "Authors": "authors",
    "Year": "year",
    "Date": "date",
    "Title": "title",
    "Paper Link": "path",
    "Journal": "journal",
    "Volume": "volume",
    "Issue": "issue",
    "Pages": "pages",
    "DOI": "doi",
    "PDF": "pdf",
    "Preprint": "preprint",
    "ShareIt": "shareit",
    "Supplemental Information": "supplemental",
    "GitHub": "github",
    "Code": "code",
    "Data": "data",
    "Highly Cited": "highlycited",
    "Hot Paper": "hotpaper",
    "Awards": "awards",
    "Media Coverage": "mediacoverage",
    "Invited Presentation": "invitedpresentation",
    "Categories": "categories",
}


# =============================================================================
# Conjunto de campos que devem ser convertidos para inteiro, quando possível
# =============================================================================
# Hoje apenas "year" é tratado assim, mas o conjunto facilita expansão futura.
INT_KEYS = {"year"}


def yaml_scalar(value: str) -> str:
    """
    Retorna um valor escalar formatado corretamente para YAML.

    O objetivo desta função é evitar que certos caracteres especiais quebrem
    a sintaxe do YAML.

    Exemplo:
        Entrada:  Hello: world
        Saída:    "Hello: world"

    Regras:
    - Se o valor contém caracteres especiais do YAML, ele será colocado entre aspas.
    - Aspas duplas e barras invertidas são escapadas quando necessário.
    - Se não houver necessidade, o valor é retornado "como está".
    """
    # Caracteres que costumam exigir aspas no YAML
    SPECIAL = set(':#|>[]{}*!,%@`')

    if any(c in value for c in SPECIAL) or value.startswith(('"', "'", '-', '?')):
        escaped = value.replace('\\', '\\\\').replace('"', '\\"')
        return f'"{escaped}"'

    return value


def yaml_inline_list(values: list[str]) -> str:
    """
    Constrói uma lista YAML inline no formato:

        [a, b, "c d"]

    Essa representação é usada principalmente para o campo "categories",
    que o Quarto entende bem como lista simples.
    """
    return "[" + ", ".join(yaml_scalar(v) for v in values) + "]"


def parse_categories(value: str) -> list[str]:
    """
    Converte o conteúdo da célula "Categories" em uma lista de categorias.

    A planilha pode usar diferentes separadores, como:
    - ponto e vírgula ;
    - barra vertical |
    - vírgula ,

    Esta função normaliza tudo para vírgula e devolve uma lista limpa,
    sem espaços extras e sem itens vazios.
    """
    s = str(value).strip()
    if not s:
        return []

    # Normaliza separadores comuns de planilhas para um formato único
    for sep in (";", "|", ","):
        s = s.replace(sep, ",")

    items = [x.strip() for x in s.split(",")]
    return [x for x in items if x]


def parse_row(headers: list, row: tuple) -> dict:
    """
    Converte uma linha da planilha em um dicionário com campos prontos
    para serem serializados em YAML.

    Parâmetros:
        headers:
            Lista com os nomes das colunas da primeira linha da planilha.
        row:
            Tupla com os valores da linha atual.

    Retorno:
        dict com os campos relevantes da publicação.

    Regras importantes:
    - Colunas sem nome são ignoradas.
    - Células vazias são ignoradas.
    - Campos em INT_KEYS tentam ser convertidos para int.
    - O campo "categories" é convertido para lista.
    """
    rec = {}

    for header, cell_value in zip(headers, row):
        # Ignora colunas sem cabeçalho definido
        if header is None:
            continue

        header_text = str(header).strip()

        # Ignora cabeçalhos vazios ou só com espaços
        if not header_text:
            continue

        # Traduz nome da coluna para a chave YAML esperada
        # Se não houver mapeamento explícito, cria uma chave automática
        key = KEY_MAP.get(header_text, header_text.lower().replace(" ", "-"))

        # Ignora células vazias
        if cell_value is None:
            continue

        # Tenta converter campos numéricos definidos em INT_KEYS
        if key in INT_KEYS:
            try:
                rec[key] = int(cell_value)
            except (ValueError, TypeError):
                s = str(cell_value).strip()
                if s:
                    rec[key] = s

        # Trata o campo de categorias como lista
        elif key == "categories":
            cats = parse_categories(cell_value)
            if cats:
                rec[key] = cats

        # Para os demais campos, salva como string se houver conteúdo
        else:
            s = str(cell_value).strip()
            if s:
                rec[key] = s

    return rec


def record_to_yaml(rec: dict) -> str:
    """
    Converte um registro de publicação (dict) para um bloco YAML.

    O resultado tem formato de item de lista YAML, por exemplo:

        - title: "Meu artigo"
          year: 2024
          authors: "Autor A; Autor B"

    A ordem dos campos é controlada manualmente para manter a saída organizada,
    estável e compatível com o template do site.
    """
    # Ordem fixa dos campos no YAML de saída
    FIELD_ORDER = [
        "section", "authors", "year", "date", "title", "path", "journal",
        "volume", "issue", "pages", "doi", "pdf", "preprint", "shareit", "supplemental",
        "github", "code", "data", "highlycited", "hotpaper", "awards", "mediacoverage", "invitedpresentation",
        "categories",
    ]

    lines = []
    first = True

    for key in FIELD_ORDER:
        if key not in rec:
            continue

        value = rec[key]

        # O primeiro campo recebe o prefixo "- " para marcar o item da lista YAML
        if first:
            prefix = "- "
            indent = " "
            first = False
        else:
            prefix = " "
            indent = " "

        # Inteiros são gravados sem aspas
        if isinstance(value, int):
            lines.append(f"{prefix}{key}: {value}")

        # Categorias são gravadas como lista inline YAML
        elif isinstance(value, list) and key == "categories":
            lines.append(
                f"{prefix}{key}: {yaml_inline_list([str(v).strip() for v in value if str(v).strip()])}"
            )

        # Demais valores são tratados como escalares YAML
        else:
            lines.append(f"{prefix}{key}: {yaml_scalar(value)}")

    return "\n".join(lines)


def convert(input_path: str, output_path: str) -> None:
    """
    Executa a conversão completa do arquivo Excel para YAML.

    Etapas:
    1. Abre a planilha de entrada.
    2. Lê os cabeçalhos da primeira linha.
    3. Percorre as linhas de dados.
    4. Converte cada linha em um registro estruturado.
    5. Gera o texto YAML final.
    6. Salva o arquivo de saída.
    7. Tenta validar o YAML gerado com PyYAML, se disponível.
    """
    print(f"Reading : {input_path}")

    wb = openpyxl.load_workbook(input_path)
    ws = wb.active

    # Usa a primeira linha da planilha como cabeçalho
    headers = [cell.value for cell in ws[1]]
    print(f"Columns : {headers}")

    records = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        # Ignora linhas totalmente vazias
        if all(v is None for v in row):
            continue

        rec = parse_row(headers, row)
        if rec:
            records.append(rec)

    print(f"Records : {len(records)}")

    # Converte todos os registros para blocos YAML separados por linha em branco
    yaml_blocks = [record_to_yaml(r) for r in records]
    output_text = "\n\n".join(yaml_blocks) + "\n"

    # Escreve o YAML final em UTF-8
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(output_text)

    print(f"Written : {output_path}")

    # Validação opcional usando PyYAML
    # Isso ajuda a detectar rapidamente se o YAML gerado ficou inválido.
    try:
        import yaml

        with open(output_path, encoding="utf-8") as f:
            parsed = yaml.safe_load(f)

        print(f"Validated: {len(parsed)} entries parsed successfully by PyYAML")

    except ImportError:
        print("Tip: install PyYAML (`pip install pyyaml`) for automatic validation")

    except Exception as e:
        print(f"WARNING: YAML validation failed — {e}")


def should_convert(input_path: str, output_path: str, force: bool = False) -> bool:
    """
    Decide se a conversão deve ser executada.

    Retorna True quando:
    - o usuário passou --force;
    - o arquivo de saída ainda não existe;
    - o arquivo de entrada é mais novo que o arquivo de saída.

    Retorna False quando:
    - o arquivo de saída já existe e está atualizado em relação ao de entrada.

    Isso evita rodar a conversão desnecessariamente a cada render do site.
    """
    if force:
        return True

    src = Path(input_path)
    dst = Path(output_path)

    # Se o arquivo de entrada não existir, o script não pode continuar
    if not src.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    # Se o arquivo de saída não existir, é necessário converter
    if not dst.exists():
        print(f"Output missing: {output_path}; running conversion.")
        return True

    # Compara datas de modificação dos arquivos
    src_mtime = src.stat().st_mtime
    dst_mtime = dst.stat().st_mtime

    # Se a entrada for mais recente que a saída, refaz a conversão
    if src_mtime > dst_mtime:
        print(f"Detected update in {input_path}; running conversion.")
        return True

    print(f"No update in {input_path}; skip conversion.")
    return False


if __name__ == "__main__":
    # Remove a flag --force da lista principal de argumentos
    # para que sobrem apenas os caminhos de entrada e saída.
    args = [a for a in sys.argv[1:] if a != "--force"]

    # Detecta se o usuário pediu conversão forçada
    force = "--force" in sys.argv[1:]

    # Usa arquivos passados por linha de comando, se existirem;
    # caso contrário, usa os caminhos padrão definidos no topo do script.
    input_file = args[0] if len(args) > 0 else DEFAULT_INPUT
    output_file = args[1] if len(args) > 1 else DEFAULT_OUTPUT

    # Só executa a conversão se realmente houver necessidade
    if should_convert(input_file, output_file, force=force):
        convert(input_file, output_file)
