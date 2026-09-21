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


def create_first_pivot_table() -> None:
    import win32com.client  # type: ignore

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        workbook = excel.Workbooks.Open(str(OUTPUT_PATH))

        for sheet in workbook.Worksheets:
            if sheet.Name == "Dinamicas":
                sheet.Delete()
                break

        pivot_sheet = workbook.Worksheets.Add()
        pivot_sheet.Name = "Dinamicas"

        base_sheet = workbook.Worksheets("B2B")
        xl_up = -4162
        xl_to_left = -4159
        last_row = base_sheet.Cells(base_sheet.Rows.Count, 1).End(xl_up).Row
        last_col = base_sheet.Cells(1, base_sheet.Columns.Count).End(xl_to_left).Column
        source_data = f"'B2B'!R1C1:R{last_row}C{last_col}"

        pivot_cache = workbook.PivotCaches().Create(SourceType=1, SourceData=source_data)
        pivot_table = pivot_cache.CreatePivotTable(
            TableDestination="'Dinamicas'!R1C1", TableName="ResumoStatus"
        )

        xl_row_field = 1
        pivot_table.PivotFields("UF").Orientation = xl_row_field
        pivot_table.PivotFields("UF").Position = 1
        pivot_table.PivotFields("STATUS").Orientation = xl_row_field
        pivot_table.PivotFields("STATUS").Position = 2

        xl_count = -4112
        xl_sum = -4157
        xl_average = -4106
        campo_total_id = pivot_table.AddDataField(
            pivot_table.PivotFields("ID"), "TOTAL_ID", xl_count
        )
        campo_valor_total = pivot_table.AddDataField(
            pivot_table.PivotFields("CUSTO TOTAL"), "VALOR TOTAL", xl_sum
        )
        campo_media = pivot_table.AddDataField(
            pivot_table.PivotFields("CUSTO TOTAL"), "MÉDIA", xl_average
        )

        pivot_table.TableStyle2 = "PivotStyleMedium9"
        pivot_table.RowAxisLayout(0)
        pivot_table.CompactLayoutRowHeader = "STATUS"
        pivot_table.PivotFields("UF").LayoutSubtotalLocation = 1

        table_range = pivot_table.TableRange1
        first_row = table_range.Row
        first_col = table_range.Column
        last_table_row = first_row + table_range.Rows.Count - 1
        last_table_col = first_col + table_range.Columns.Count - 1

        data_start_row = first_row + 2
        if data_start_row <= last_table_row:
            pivot_sheet.Range(
                pivot_sheet.Cells(data_start_row, first_col + 1),
                pivot_sheet.Cells(last_table_row, first_col + 1),
            ).NumberFormat = "0"
            pivot_sheet.Range(
                pivot_sheet.Cells(data_start_row, first_col + 2),
                pivot_sheet.Cells(last_table_row, first_col + 3),
            ).NumberFormatLocal = "R$ #.##0,00"

        for row_idx in range(first_row + 1, last_table_row + 1):
            row_values = [
                pivot_sheet.Cells(row_idx, col_idx).Value
                for col_idx in range(first_col, last_table_col + 1)
            ]
            if any(isinstance(value, str) and "Faturado" in value for value in row_values):
                pivot_sheet.Range(
                    pivot_sheet.Cells(row_idx, first_col),
                    pivot_sheet.Cells(row_idx, last_table_col),
                ).Interior.ColorIndex = 35

        pivot_sheet.Columns("A:D").AutoFit()
        if pivot_sheet.Columns("A").ColumnWidth < 18:
            pivot_sheet.Columns("A").ColumnWidth = 18
        if pivot_sheet.Columns("B").ColumnWidth < 10:
            pivot_sheet.Columns("B").ColumnWidth = 10
        if pivot_sheet.Columns("C").ColumnWidth < 15:
            pivot_sheet.Columns("C").ColumnWidth = 15
        if pivot_sheet.Columns("D").ColumnWidth < 12:
            pivot_sheet.Columns("D").ColumnWidth = 12

        workbook.Save()
        workbook.Close(SaveChanges=True)
    finally:
        excel.Quit()


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

    create_first_pivot_table()
    print(f"Relatório gerado em: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
