from __future__ import annotations

from pathlib import Path

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent / "bases"
EVENTOS_PATH = BASE_DIR / "Eventos.xls"
QUANTIFICACAO_PATH = BASE_DIR / "Quantificacao.xls"
MATERIAIS_PATH = BASE_DIR / "Materiais.xls"
FATURADO_PATH = BASE_DIR / "Faturado.xlsx"
OUTPUT_PATH = Path(__file__).resolve().parent / "relatorio_b2b.xlsx"


def read_table(path: Path) -> pd.DataFrame:
    if path.suffix.lower() == ".xls":
        df = pd.read_html(path)[0]
    else:
        df = pd.read_excel(path, engine="openpyxl")
    return normalize_header(df)


def normalize_header(df: pd.DataFrame) -> pd.DataFrame:
    if len(df) == 0:
        return df

    first_row = df.iloc[0]
    if all(isinstance(col, int) for col in df.columns) and first_row.notna().all():
        df = df.copy()
        df.columns = [str(col).strip() for col in first_row]
        df = df.iloc[1:].reset_index(drop=True)
    else:
        df.columns = [str(col).strip() for col in df.columns]
    return df


def normalize_id(series: pd.Series) -> pd.Series:
    return series.astype(str).str.strip().str.replace(r"\.0$", "", regex=True)


def parse_number(value: object) -> float:
    if pd.isna(value):
        return 0.0

    text = str(value).strip()
    if not text:
        return 0.0

    text = "".join(ch for ch in text if ch.isdigit() or ch in ",.-")
    if not text:
        return 0.0

    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(".", "").replace(",", ".")

    try:
        return float(text)
    except ValueError:
        return 0.0


def main() -> None:
    eventos = read_table(EVENTOS_PATH)
    quantificacao = read_table(QUANTIFICACAO_PATH)
    materiais = read_table(MATERIAIS_PATH)
    faturado = read_table(FATURADO_PATH)

    for df in [eventos, quantificacao, materiais, faturado]:
        if "ID" in df.columns:
            df["ID"] = normalize_id(df["ID"])

    eventos_filtrado = eventos[
        (eventos["TIPO DE CABO"].astype(str).str.strip().str.upper() == "CORPORATIVO")
        & (eventos["EMPREITEIRA"].astype(str).str.strip().str.upper() == "TLP")
    ].copy()

    ids_b2b = set(eventos_filtrado["ID"])

    servico = quantificacao[quantificacao["ID"].isin(ids_b2b)].copy()
    materiais_filtrado = materiais[materiais["ID"].isin(ids_b2b)].copy()

    servico["CUSTO TOTAL NUM"] = servico["CUSTO TOTAL"].apply(parse_number)
    custo_por_id = servico.groupby("ID", dropna=False)["CUSTO TOTAL NUM"].sum()

    eventos_filtrado["CUSTO TOTAL"] = eventos_filtrado["ID"].map(custo_por_id).fillna(0.0)

    ids_faturado = set(faturado["ID"]) if "ID" in faturado.columns else set()
    ionix_col = "IONIX" if "IONIX" in eventos_filtrado.columns else "Ionix"
    if ionix_col in eventos_filtrado.columns:
        eventos_filtrado["Faturado"] = eventos_filtrado.apply(
            lambda row: "Sim"
            if (row["ID"] in ids_faturado) or (str(row[ionix_col]).strip() in ids_faturado)
            else "Não",
            axis=1,
        )
    else:
        eventos_filtrado["Faturado"] = eventos_filtrado["ID"].apply(
            lambda value: "Sim" if value in ids_faturado else "Não"
        )
    if "STATUS" in eventos_filtrado.columns:
        eventos_filtrado.loc[eventos_filtrado["Faturado"] == "Sim", "STATUS"] = "Faturado"
    elif "Status" in eventos_filtrado.columns:
        eventos_filtrado.loc[eventos_filtrado["Faturado"] == "Sim", "Status"] = "Faturado"

    servico["CUSTO TOTAL"] = servico["CUSTO TOTAL NUM"].astype(float)
    servico = servico.drop(columns=["CUSTO TOTAL NUM"])

    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
        eventos_filtrado.to_excel(writer, index=False, sheet_name="B2B")
        servico.to_excel(writer, index=False, sheet_name="Serviços")
        materiais_filtrado.to_excel(writer, index=False, sheet_name="Materiais")
        for sheet_name in ["B2B", "Serviços"]:
            sheet = writer.book[sheet_name]
            custo_col_idx = None
            for idx, cell in enumerate(sheet[1], start=1):
                if cell.value == "CUSTO TOTAL":
                    custo_col_idx = idx
                    break
            if custo_col_idx is None:
                continue
            for row_idx in range(2, sheet.max_row + 1):
                sheet.cell(row=row_idx, column=custo_col_idx).number_format = "#,##0.00"

    print(f"Relatório gerado em: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
