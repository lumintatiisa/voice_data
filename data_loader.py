# -*- coding: utf-8 -*-
"""
data_loader.py
--------------
Camada de ingestão e normalização dos arquivos TFP (Título/Fatura Pendente).

Regras de negócio importantes (ver README.md / metodologia no app para o texto completo):

1) Cada arquivo .xlsx representa um "período" (uma ou mais semanas consolidadas).
   O período é inferido a partir do NOME DO ARQUIVO, no padrão:
       TFP_SEMANA_<n>.xlsx            -> período com 1 semana
       TFP_SEMANA_<n>_E_<m>.xlsx      -> período consolidando as semanas n e m
   Novos arquivos seguindo esse padrão são descobertos automaticamente.

2) Cada arquivo possui (pelo menos) duas abas: uma de linhas MÓVEIS e outra de
   linhas FIXAS, cada uma listando os clientes que POSSUEM fatura em aberto
   (débito). Não existe, nos arquivos, a base total de clientes por
   consultor/parceiro — logo o "%" calculado neste dashboard é a
   PARTICIPAÇÃO de cada consultor/parceiro dentro do total de inadimplentes
   daquele tipo de linha e período (e não uma taxa de inadimplência sobre a
   carteira total). Essa decisão foi validada com o time de negócio.

3) CORREÇÃO CONHECIDA DE DADOS: no arquivo `TFP_SEMANA_1_E_2.xlsx`, os rótulos
   das abas "TFP MÓVEL" e "TFP FIXA" estavam TROCADOS em relação ao conteúdo
   real (confirmado comparando CNPJs com o arquivo `TFP_SEMANA_3.xlsx`, que
   serve de referência). A correção abaixo inverte os rótulos apenas para
   este arquivo específico. Caso os arquivos futuros venham com a mesma
   inconsistência, adicione uma entrada em `SHEET_TYPE_OVERRIDES`.
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import pandas as pd

DATA_DIR = Path(__file__).parent / "data"

# ---------------------------------------------------------------------------
# Correções manuais conhecidas: (nome_do_arquivo, nome_da_aba_normalizado) -> tipo correto
# tipo correto em {"MOVEL", "FIXA"}
# ---------------------------------------------------------------------------
SHEET_TYPE_OVERRIDES = {
    ("TFP_SEMANA_1_E_2.xlsx", "TFP FIXA"): "MOVEL",
    ("TFP_SEMANA_1_E_2.xlsx", "TFP MOVEL"): "FIXA",  # acento removido na normalização
}

EXPECTED_COLUMNS = ["CNPJ", "NOME CLIENTE", "CONSULTOR", "PARCEIRO"]


def _strip_accents(text: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFKD", text) if not unicodedata.combining(c)
    )


def _normalize(text: str) -> str:
    return _strip_accents(str(text)).upper().strip()


def _guess_sheet_type(filename: str, sheet_name: str) -> Optional[str]:
    """Determina se uma aba é MOVEL ou FIXA, aplicando overrides conhecidos primeiro."""
    norm_sheet = _normalize(sheet_name)
    key = (filename, norm_sheet)
    if key in SHEET_TYPE_OVERRIDES:
        return SHEET_TYPE_OVERRIDES[key]

    if "MOV" in norm_sheet:
        return "MOVEL"
    if "FIX" in norm_sheet:
        return "FIXA"
    return None


def parse_period_from_filename(filename: str) -> Optional[dict]:
    """
    Extrai informações de período do nome do arquivo.
    Aceita padrões como:
        TFP_SEMANA_3.xlsx            -> semanas=[3]
        TFP_SEMANA_1_E_2.xlsx        -> semanas=[1, 2]
        TFP_SEMANA_4_5.xlsx          -> semanas=[4, 5]
    """
    stem = Path(filename).stem
    norm = _normalize(stem)
    numbers = [int(n) for n in re.findall(r"\d+", norm)]
    if not numbers:
        return None

    if len(numbers) == 1:
        label = f"Semana {numbers[0]}"
    else:
        label = "Semana " + " e ".join(str(n) for n in numbers)

    return {
        "weeks": numbers,
        "label": label,
        "sort_key": max(numbers),
    }


@dataclass
class LoadIssue:
    file: str
    message: str


def load_all_periods(data_dir: Path = DATA_DIR) -> tuple[pd.DataFrame, list[LoadIssue]]:
    """
    Varre `data_dir` em busca de arquivos TFP_SEMANA_*.xlsx, lê as abas de
    MÓVEL e FIXA de cada um, aplica as correções conhecidas e retorna um único
    DataFrame consolidado, junto com uma lista de avisos de qualidade de dados.
    """
    issues: list[LoadIssue] = []
    frames = []

    if not data_dir.exists():
        return pd.DataFrame(), [LoadIssue("-", f"Pasta de dados não encontrada: {data_dir}")]

    files = sorted(data_dir.glob("*.xlsx"))
    files = [f for f in files if not f.name.startswith("~$")]

    if not files:
        issues.append(LoadIssue("-", f"Nenhum arquivo .xlsx encontrado em {data_dir}"))
        return pd.DataFrame(), issues

    for fpath in files:
        period = parse_period_from_filename(fpath.name)
        if period is None:
            issues.append(
                LoadIssue(fpath.name, "Nome de arquivo fora do padrão TFP_SEMANA_*; ignorado.")
            )
            continue

        try:
            xls = pd.ExcelFile(fpath, engine="openpyxl")
        except Exception as e:  # noqa: BLE001
            issues.append(LoadIssue(fpath.name, f"Falha ao abrir o arquivo: {e}"))
            continue

        found_types = set()
        for sheet_name in xls.sheet_names:
            tipo = _guess_sheet_type(fpath.name, sheet_name)
            if tipo is None:
                issues.append(
                    LoadIssue(
                        fpath.name,
                        f"Aba '{sheet_name}' não reconhecida como MÓVEL ou FIXA; ignorada.",
                    )
                )
                continue

            df = xls.parse(sheet_name)
            # remove linhas totalmente vazias e colunas fora do padrão
            df = df.dropna(how="all")
            missing = [c for c in EXPECTED_COLUMNS if c not in df.columns]
            if missing:
                issues.append(
                    LoadIssue(
                        fpath.name,
                        f"Aba '{sheet_name}' sem colunas esperadas {missing}; ignorada.",
                    )
                )
                continue

            df = df[df["CNPJ"].notna()].copy()
            df = df[EXPECTED_COLUMNS].copy()

            df["CNPJ"] = (
                df["CNPJ"]
                .astype(str)
                .str.replace(r"\.0$", "", regex=True)
                .str.replace(r"\D", "", regex=True)
                .str.zfill(14)
            )
            for col in ["NOME CLIENTE", "CONSULTOR", "PARCEIRO"]:
                df[col] = df[col].astype(str).str.strip()
                df[col] = df[col].replace({"nan": "Não informado", "": "Não informado"})

            df["TIPO"] = tipo
            df["ARQUIVO"] = fpath.name
            df["PERIODO"] = period["label"]
            df["PERIODO_SORT"] = period["sort_key"]
            df["SEMANAS"] = ", ".join(str(w) for w in period["weeks"])

            frames.append(df)
            found_types.add(tipo)

        for expected in ("MOVEL", "FIXA"):
            if expected not in found_types:
                issues.append(
                    LoadIssue(fpath.name, f"Nenhuma aba do tipo {expected} encontrada.")
                )

    if not frames:
        return pd.DataFrame(), issues

    full = pd.concat(frames, ignore_index=True)

    # Deduplicação: um mesmo cliente pode aparecer em mais de uma linha na
    # mesma aba (múltiplos contratos/linhas). Para a métrica de "clientes"
    # mantemos uma linha por CNPJ + CONSULTOR + PARCEIRO + TIPO + PERIODO,
    # preservando a contagem de contratos em coluna separada.
    full["QTD_REGISTROS"] = 1
    dedup_keys = ["CNPJ", "CONSULTOR", "PARCEIRO", "TIPO", "PERIODO"]
    agg = (
        full.groupby(dedup_keys, as_index=False)
        .agg(
            NOME_CLIENTE=("NOME CLIENTE", "first"),
            QTD_REGISTROS=("QTD_REGISTROS", "sum"),
            ARQUIVO=("ARQUIVO", "first"),
            PERIODO_SORT=("PERIODO_SORT", "first"),
            SEMANAS=("SEMANAS", "first"),
        )
    )
    agg = agg.rename(columns={"NOME_CLIENTE": "NOME CLIENTE"})
    return agg, issues


def period_order(df: pd.DataFrame) -> list[str]:
    """Retorna os rótulos de período ordenados cronologicamente."""
    if df.empty:
        return []
    return (
        df[["PERIODO", "PERIODO_SORT"]]
        .drop_duplicates()
        .sort_values("PERIODO_SORT")["PERIODO"]
        .tolist()
    )
