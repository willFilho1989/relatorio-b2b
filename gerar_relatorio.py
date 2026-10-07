from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd


APP_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", APP_DIR))


def resolve_base_dir() -> Path:
    candidates = [
        APP_DIR / "bases",
        APP_DIR.parent / "bases",
        Path.cwd() / "bases",
        RESOURCE_DIR / "bases",
    ]
    for candidate in candidates:
        if candidate.exists() and candidate.is_dir():
            return candidate
    raise FileNotFoundError(
        "Pasta 'bases' não encontrada. Coloque-a ao lado do executável ou no diretório atual."
    )


BASE_DIR = resolve_base_dir()
EVENTOS_PATH = BASE_DIR / "Eventos.xls"
QUANTIFICACAO_PATH = BASE_DIR / "Quantificacao.xls"
MATERIAIS_PATH = BASE_DIR / "Materiais.xls"
FATURADO_PATH = BASE_DIR / "Faturado.xlsx"
OUTPUT_PATH = APP_DIR / "relatorio_b2b.xlsx"
SERVICO_DISTINCT_TOTAL = 0
MATERIAIS_DISTINCT_TOTAL = 0


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


def cleanup_generated_reports() -> None:
    for report_path in APP_DIR.glob("relatorio_b2b_atualizado_*.xlsx"):
        try:
            report_path.unlink()
        except OSError:
            pass


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


def create_pivot_tables() -> None:
    import win32com.client  # type: ignore

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        workbook = excel.Workbooks.Open(str(OUTPUT_PATH))
        theme = workbook.Theme.ThemeColorScheme
        theme.Colors(5).RGB = 0xF4B183
        theme.Colors(6).RGB = 0xED7D31
        theme.Colors(7).RGB = 0xC55A11
        theme.Colors(8).RGB = 0xA64D2A
        theme.Colors(9).RGB = 0xFCE4D6
        theme.Colors(10).RGB = 0xF8CBAD

        for sheet in list(workbook.Worksheets):
            if sheet.Name == "Dinamicas":
                sheet.Delete()
                break

        pivot_sheet = workbook.Worksheets.Add()
        pivot_sheet.Name = "Dinamicas"

        xl_up = -4162
        xl_to_left = -4159
        xl_row_field = 1
        xl_count = -4112
        xl_sum = -4157
        xl_average = -4106
        xl_tabular = 1
        xl_left = -4131
        xl_database = 1

        def style_pivot(pivot_table):
            pivot_table.TableStyle = "PivotStyleMedium2"
            pivot_table.TableStyle2 = "PivotStyleMedium2"
            table_range = pivot_table.TableRange1
            table_range.HorizontalAlignment = xl_left
            table_range.Interior.Color = 0xFFFFFF
            table_range.Font.Color = 0x000000

            header_rows = 1
            if pivot_table.Name == "ResumoStatus":
                header_rows = 2

            for row_offset in range(header_rows):
                header_row = table_range.Row + row_offset
                header = pivot_sheet.Range(
                    pivot_sheet.Cells(header_row, table_range.Column),
                    pivot_sheet.Cells(
                        header_row, table_range.Column + table_range.Columns.Count - 1
                    ),
                )
                header.Interior.Color = 0xF4B183
                header.Font.Color = 0xFFFFFF
                header.Font.Bold = True

        # B2B
        b2b = workbook.Worksheets("B2B")
        last_row = b2b.Cells(b2b.Rows.Count, 1).End(xl_up).Row
        last_col = b2b.Cells(1, b2b.Columns.Count).End(xl_to_left).Column
        source = f"'B2B'!R1C1:R{last_row}C{last_col}"
        cache = workbook.PivotCaches().Create(SourceType=xl_database, SourceData=source)
        pt = cache.CreatePivotTable(TableDestination="'Dinamicas'!R1C1", TableName="ResumoStatus")
        pt.PivotFields("UF").Orientation = xl_row_field
        pt.PivotFields("UF").Position = 1
        pt.PivotFields("STATUS").Orientation = xl_row_field
        pt.PivotFields("STATUS").Position = 2
        pt.AddDataField(pt.PivotFields("ID"), "TOTAL_ID", xl_count)
        pt.AddDataField(pt.PivotFields("CUSTO TOTAL"), "VALOR TOTAL", xl_sum)
        pt.AddDataField(pt.PivotFields("CUSTO TOTAL"), "MÉDIA", xl_average)
        pt.RowAxisLayout(0)
        pt.CompactLayoutRowHeader = "STATUS"
        pt.PivotFields("UF").LayoutSubtotalLocation = 1
        
        # Fixar ordem dos status
        status_order = ["Cancelado", "Eps/Provisorio", "Supervisor Cidade", "Concluido", "Faturado"]
        status_field = pt.PivotFields("STATUS")
        for idx, status_value in enumerate(status_order):
            try:
                status_field.PivotItems(status_value).Position = idx + 1
            except:
                pass
        
        style_pivot(pt)
        b2b_range = pt.TableRange1
        pivot_sheet.Range(
            pivot_sheet.Cells(b2b_range.Row + 1, b2b_range.Column + 1),
            pivot_sheet.Cells(b2b_range.Row + b2b_range.Rows.Count - 1, b2b_range.Column + 1),
        ).NumberFormat = "0"
        pivot_sheet.Range(
            pivot_sheet.Cells(b2b_range.Row + 1, b2b_range.Column + 2),
            pivot_sheet.Cells(b2b_range.Row + b2b_range.Rows.Count - 1, b2b_range.Column + 3),
        ).NumberFormatLocal = "R$ #.##0,00"
        for row_idx in range(b2b_range.Row + 1, b2b_range.Row + b2b_range.Rows.Count):
            vals = [
                pivot_sheet.Cells(row_idx, col_idx).Value
                for col_idx in range(b2b_range.Column, b2b_range.Column + b2b_range.Columns.Count)
            ]
            if any(isinstance(v, str) and "Faturado" in v for v in vals):
                pivot_sheet.Range(
                    pivot_sheet.Cells(row_idx, b2b_range.Column),
                    pivot_sheet.Cells(row_idx, b2b_range.Column + b2b_range.Columns.Count - 1),
                ).Interior.ColorIndex = 35

        # Serviços
        serv = workbook.Worksheets("_ServicosPivot")
        last_row = serv.Cells(serv.Rows.Count, 1).End(xl_up).Row
        last_col = serv.Cells(1, serv.Columns.Count).End(xl_to_left).Column
        source = f"'_ServicosPivot'!R1C1:R{last_row}C{last_col}"
        cache = workbook.PivotCaches().Create(SourceType=xl_database, SourceData=source)
        pt = cache.CreatePivotTable(TableDestination="'Dinamicas'!R1C6", TableName="ResumoServicos")
        pt.PivotFields("CODIGO").Orientation = xl_row_field
        pt.PivotFields("CODIGO").Position = 1
        pt.PivotFields("DESCRICAO").Orientation = xl_row_field
        pt.PivotFields("DESCRICAO").Position = 2
        pt.PivotFields("CODIGO").Subtotals = (False,) * 12
        pt.PivotFields("DESCRICAO").Subtotals = (False,) * 12
        pt.AddDataField(pt.PivotFields("ID"), "Contagem Distinta de ID", xl_count)
        pt.RowAxisLayout(xl_tabular)
        pt.RowGrand = False
        style_pivot(pt)
        pt.TableRange1.HorizontalAlignment = xl_left
        serv_grand_total_row = pt.TableRange1.Row + pt.TableRange1.Rows.Count - 1
        pivot_sheet.Rows(serv_grand_total_row).Hidden = True
        serv_end_row = serv_grand_total_row + 1
        pivot_sheet.Cells(serv_end_row, 6).Value = "Total Geral"
        pivot_sheet.Cells(serv_end_row, 8).Value = SERVICO_DISTINCT_TOTAL

        # Materiais
        mat = workbook.Worksheets("_MateriaisPivot")
        last_row = mat.Cells(mat.Rows.Count, 1).End(xl_up).Row
        last_col = mat.Cells(1, mat.Columns.Count).End(xl_to_left).Column
        source = f"'_MateriaisPivot'!R1C1:R{last_row}C{last_col}"
        cache = workbook.PivotCaches().Create(SourceType=xl_database, SourceData=source)
        pt = cache.CreatePivotTable(TableDestination="'Dinamicas'!R1C10", TableName="ResumoMateriais")
        pt.PivotFields("CODIGO").Orientation = xl_row_field
        pt.PivotFields("CODIGO").Position = 1
        pt.PivotFields("MATERIAL").Orientation = xl_row_field
        pt.PivotFields("MATERIAL").Position = 2
        pt.PivotFields("CODIGO").Subtotals = (False,) * 12
        pt.PivotFields("MATERIAL").Subtotals = (False,) * 12
        pt.AddDataField(pt.PivotFields("ID"), "Contagem Distinta de ID", xl_count)
        pt.RowAxisLayout(xl_tabular)
        pt.RowGrand = False
        style_pivot(pt)
        pt.TableRange1.HorizontalAlignment = xl_left
        mat_grand_total_row = pt.TableRange1.Row + pt.TableRange1.Rows.Count - 1
        pivot_sheet.Rows(mat_grand_total_row).Hidden = True
        mat_end_row = mat_grand_total_row + 1
        pivot_sheet.Cells(mat_end_row, 10).Value = "Total Geral"
        pivot_sheet.Cells(mat_end_row, 12).Value = MATERIAIS_DISTINCT_TOTAL

        if workbook.Windows.Count > 0:
            workbook.Windows(1).DisplayGridlines = False

        pivot_sheet.Range("A:O").HorizontalAlignment = xl_left
        pivot_sheet.Columns("A:O").AutoFit()

        workbook.Save()
        workbook.Close(SaveChanges=True)
    finally:
        excel.Quit()


def main() -> None:
    global SERVICO_DISTINCT_TOTAL, MATERIAIS_DISTINCT_TOTAL
    cleanup_generated_reports()
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
    SERVICO_DISTINCT_TOTAL = int(servico["ID"].nunique())
    MATERIAIS_DISTINCT_TOTAL = int(materiais_filtrado["ID"].nunique())
    servico_pivot_base = servico[["CODIGO", "DESCRICAO", "ID"]].drop_duplicates().copy()
    materiais_pivot_base = materiais_filtrado[["CODIGO", "MATERIAL", "ID"]].drop_duplicates().copy()

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
        servico_pivot_base.to_excel(writer, index=False, sheet_name="_ServicosPivot")
        materiais_pivot_base.to_excel(writer, index=False, sheet_name="_MateriaisPivot")
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
        for sheet_name in ["_ServicosPivot", "_MateriaisPivot"]:
            writer.book[sheet_name].sheet_state = "hidden"

    create_pivot_tables()
    print(f"Relatório gerado em: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
