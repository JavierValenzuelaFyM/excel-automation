import pandas as pd
from pathlib import Path
from datetime import datetime

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.chart import BarChart, Reference
from openpyxl.worksheet.table import Table, TableStyleInfo


# ============================================================
# CONFIGURATION
# ============================================================

INPUT_FOLDER = Path("input")
OUTPUT_FOLDER = Path("output")
OUTPUT_FOLDER.mkdir(exist_ok=True)


# ============================================================
# DATA CLEANING
# ============================================================

def clean_data(df):
    original_rows = len(df)
    duplicates = int(df.duplicated().sum())
    missing_values = int(df.isna().sum().sum())

    # Remove duplicates
    df = df.drop_duplicates().copy()

    # Handle missing values
    for column in df.columns:
        if df[column].dtype == "object":
            df[column] = df[column].fillna("Unknown")
        else:
            df[column] = df[column].fillna(0)

    # Sort automatically when order_id exists
    if "order_id" in df.columns:
        df = df.sort_values("order_id")

    return df, original_rows, duplicates, missing_values


# ============================================================
# STYLE HELPERS
# ============================================================

NAVY = "0F172A"
BLUE = "2563EB"
ORANGE = "F59E0B"
GREEN = "16A34A"
LIGHT_BG = "F8FAFC"
LIGHT_BORDER = "E2E8F0"
WHITE = "FFFFFF"
TEXT = "1E293B"
MUTED = "64748B"


def add_kpi_card(ws, cell_range, title, value, accent):
    ws.merge_cells(cell_range)

    first_cell = cell_range.split(":")[0]
    cell = ws[first_cell]

    # Slightly smaller title size for long labels such as MISSING VALUES FOUND.
    font_size = 15 if len(title) > 18 else 16

    cell.value = f"{value:,}\n{title}"
    cell.font = Font(
        name="Aptos",
        size=font_size,
        bold=True,
        color=WHITE,
    )
    cell.fill = PatternFill("solid", fgColor=accent)
    cell.alignment = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True,
    )

    thin = Side(style="thin", color=accent)
    for row in ws[cell_range]:
        for c in row:
            c.border = Border(left=thin, right=thin, top=thin, bottom=thin)


def set_column_widths(ws, widths):
    for column, width in widths.items():
        ws.column_dimensions[column].width = width


# ============================================================
# INPUT FILES
# ============================================================

files = [
    file
    for file in (
        list(INPUT_FOLDER.glob("*.csv"))
        + list(INPUT_FOLDER.glob("*.xlsx"))
    )
    if not file.name.startswith("~$")
]

all_cleaned_data = []
summary = []
errors = []


# ============================================================
# PROCESS FILES
# ============================================================

for file in files:
    print(f"\n===== Procesando {file.name} =====")

    try:
        # ----------------------------------------------------
        # CSV
        # ----------------------------------------------------
        if file.suffix.lower() == ".csv":
            df = pd.read_csv(file)
            df, original, duplicates, missing = clean_data(df)

            master_df = df.copy()
            master_df["_source_file"] = file.name
            master_df["_source_sheet"] = "CSV"
            all_cleaned_data.append(master_df)

            summary.append({
                "Worksheet": "CSV",
                "Original rows": original,
                "Duplicates removed": duplicates,
                "Missing values found": missing,
                "Final rows": len(df),
            })

            output_file = OUTPUT_FOLDER / f"{file.stem}_clean.xlsx"
            df.to_excel(output_file, index=False)

            print("Filas originales:", original)
            print("Duplicados eliminados:", duplicates)
            print("Valores ausentes:", missing)
            print("Filas finales:", len(df))

        # ----------------------------------------------------
        # EXCEL
        # ----------------------------------------------------
        elif file.suffix.lower() == ".xlsx":
            sheets = pd.read_excel(file, sheet_name=None)
            output_file = OUTPUT_FOLDER / f"{file.stem}_clean.xlsx"

            with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                for sheet_name, df in sheets.items():
                    print(f"\nHoja: {sheet_name}")

                    df, original, duplicates, missing = clean_data(df)

                    # Individual clean workbook
                    df.to_excel(
                        writer,
                        sheet_name=sheet_name,
                        index=False,
                    )

                    # Master dataset
                    master_df = df.copy()
                    master_df["_source_file"] = file.name
                    master_df["_source_sheet"] = sheet_name
                    all_cleaned_data.append(master_df)

                    summary.append({
                        "Worksheet": sheet_name,
                        "Original rows": original,
                        "Duplicates removed": duplicates,
                        "Missing values found": missing,
                        "Final rows": len(df),
                    })

                    print("Filas originales:", original)
                    print("Duplicados eliminados:", duplicates)
                    print("Valores ausentes:", missing)
                    print("Filas finales:", len(df))

            print(f"\nArchivo creado: {output_file}")

    except Exception as error:
        errors.append({
            "File": file.name,
            "Error": str(error),
        })
        print(f"ERROR procesando {file.name}: {error}")


# ============================================================
# MASTER REPORT
# ============================================================

if all_cleaned_data:
    master_df = pd.concat(
        all_cleaned_data,
        ignore_index=True,
        sort=False,
    )

    summary_df = pd.DataFrame(summary)

    total_original = int(summary_df["Original rows"].sum())
    total_duplicates = int(summary_df["Duplicates removed"].sum())
    total_missing = int(summary_df["Missing values found"].sum())
    total_final = int(summary_df["Final rows"].sum())

    report_file = OUTPUT_FOLDER / "master_report.xlsx"

    # First create workbook with pandas
    with pd.ExcelWriter(report_file, engine="openpyxl") as writer:
        master_df.to_excel(
            writer,
            sheet_name="Cleaned_Data",
            index=False,
        )

        summary_df.to_excel(
            writer,
            sheet_name="Dashboard",
            startrow=11,
            startcol=0,
            index=False,
        )

    # ========================================================
    # PROFESSIONAL FORMATTING
    # ========================================================

    wb = load_workbook(report_file)
    dashboard = wb["Dashboard"]
    data_sheet = wb["Cleaned_Data"]

    # --------------------------------------------------------
    # DASHBOARD
    # --------------------------------------------------------

    dashboard.sheet_view.showGridLines = False

    # Number of worksheets/months controls the vertical layout.
    worksheet_count = len(summary_df)
    header_row = 12
    first_data_row = 13
    final_data_row = first_data_row + worksheet_count - 1
    total_row = final_data_row + 1
    final_table_row = total_row
    chart_start_row = final_table_row + 3

    # The dashboard background grows with the number of worksheets.
    dashboard_background_rows = chart_start_row + max(28, worksheet_count * 2 + 20)

    for row in dashboard.iter_rows(
        min_row=1,
        max_row=dashboard_background_rows,
        min_col=1,
        max_col=12,
    ):
        for cell in row:
            cell.fill = PatternFill("solid", fgColor=LIGHT_BG)

    # Main title
    dashboard.merge_cells("A1:H2")
    dashboard["A1"] = "AUTOMATIC EXCEL PROCESSING REPORT"
    dashboard["A1"].font = Font(
        name="Aptos Display",
        size=24,
        bold=True,
        color=NAVY,
    )
    dashboard["A1"].alignment = Alignment(vertical="center")

    # Subtitle
    dashboard.merge_cells("A3:H3")
    dashboard["A3"] = "Automated cleaning, consolidation and data-quality summary"
    dashboard["A3"].font = Font(
        name="Aptos",
        size=11,
        color=MUTED,
    )
    dashboard["A3"].alignment = Alignment(vertical="center")

    # KPI cards — all four cards have exactly the same width.
    add_kpi_card(dashboard, "A5:B7", "ROWS PROCESSED", total_original, NAVY)
    add_kpi_card(dashboard, "C5:D7", "DUPLICATES REMOVED", total_duplicates, BLUE)
    add_kpi_card(dashboard, "E5:F7", "MISSING VALUES FOUND", total_missing, ORANGE)
    add_kpi_card(dashboard, "G5:H7", "CLEAN ROWS GENERATED", total_final, GREEN)

    # Status line
    dashboard.merge_cells("A9:H9")
    dashboard["A9"] = (
        f"{worksheet_count} worksheets processed successfully  •  "
        f"Report generated {datetime.now().strftime('%d %b %Y')}"
    )
    dashboard["A9"].font = Font(
        name="Aptos",
        size=10,
        color=MUTED,
        italic=True,
    )

    # Details title
    dashboard.merge_cells("A11:E11")
    dashboard["A11"] = "PROCESSING DETAILS"
    dashboard["A11"].font = Font(
        name="Aptos",
        size=13,
        bold=True,
        color=NAVY,
    )

    # Header styling
    for cell in dashboard[header_row][:5]:
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.font = Font(
            name="Aptos",
            size=10,
            bold=True,
            color=WHITE,
        )
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )

    # Detail rows — count depends directly on number of worksheets/months.
    for row_number in range(first_data_row, final_data_row + 1):
        dashboard.row_dimensions[row_number].height = 22

        for cell in dashboard[row_number][:5]:
            cell.font = Font(name="Aptos", size=10, color=TEXT)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = Border(
                bottom=Side(style="thin", color=LIGHT_BORDER)
            )

        if (row_number - first_data_row) % 2 == 1:
            for cell in dashboard[row_number][:5]:
                cell.fill = PatternFill("solid", fgColor="F1F5F9")

    # Dynamic total row
    dashboard.cell(total_row, 1, "TOTAL")
    dashboard.cell(total_row, 2, total_original)
    dashboard.cell(total_row, 3, total_duplicates)
    dashboard.cell(total_row, 4, total_missing)
    dashboard.cell(total_row, 5, total_final)

    for cell in dashboard[total_row][:5]:
        cell.fill = PatternFill("solid", fgColor="E2E8F0")
        cell.font = Font(name="Aptos", size=10, bold=True, color=NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(
            top=Side(style="thin", color=NAVY),
            bottom=Side(style="thin", color=NAVY),
        )

    dashboard.row_dimensions[total_row].height = 23

    # Make all four KPI cards equal width. This fixes the orange card overflow.
    set_column_widths(
        dashboard,
        {
            "A": 18,
            "B": 18,
            "C": 18,
            "D": 18,
            "E": 18,
            "F": 18,
            "G": 18,
            "H": 18,
            "I": 14,
            "J": 14,
            "K": 14,
            "L": 14,
        },
    )

    dashboard.row_dimensions[1].height = 30
    dashboard.row_dimensions[2].height = 20
    dashboard.row_dimensions[5].height = 34
    dashboard.row_dimensions[6].height = 34
    dashboard.row_dimensions[7].height = 34

    dashboard.freeze_panes = "A12"

    # Turn the details area into a real Excel table.
    summary_table_ref = f"A12:E{final_data_row}"
    summary_table = Table(
        displayName="ProcessingDetailsTable",
        ref=summary_table_ref,
    )
    summary_table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=False,
        showColumnStripes=False,
    )
    dashboard.add_table(summary_table)

    # --------------------------------------------------------
    # CHART — BELOW THE TABLE, SIZE ALSO ADAPTS TO MONTH COUNT
    # --------------------------------------------------------

    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "Data Quality Issues by Worksheet"
    chart.y_axis.title = "Issues"
    chart.x_axis.title = ""

    # More months -> more horizontal room; bounded so it never becomes absurd.
    chart.width = min(30, max(22, 16 + worksheet_count * 1.8))
    chart.height = min(15, max(10, 8 + worksheet_count * 0.5))

    chart_data = Reference(
        dashboard,
        min_col=3,
        max_col=4,
        min_row=header_row,
        max_row=final_data_row,
    )

    chart_categories = Reference(
        dashboard,
        min_col=1,
        min_row=first_data_row,
        max_row=final_data_row,
    )

    chart.add_data(chart_data, titles_from_data=True)
    chart.set_categories(chart_categories)
    chart.legend.position = "b"

    # No data labels: keeps the chart clean and avoids text collisions.
    dashboard.add_chart(chart, f"A{chart_start_row}")

    # --------------------------------------------------------
    # CLEANED DATA SHEET
    # --------------------------------------------------------

    data_sheet.freeze_panes = "A2"
    data_sheet.sheet_view.showGridLines = False

    max_row = data_sheet.max_row
    max_col = data_sheet.max_column

    for cell in data_sheet[1]:
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.font = Font(
            name="Aptos",
            size=10,
            bold=True,
            color=WHITE,
        )
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
        )

    data_sheet.row_dimensions[1].height = 24

    table_reference = (
        f"A1:"
        f"{data_sheet.cell(row=1, column=max_col).column_letter}"
        f"{max_row}"
    )

    table = Table(
        displayName="CleanedDataTable",
        ref=table_reference,
    )

    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )

    data_sheet.add_table(table)

    # Intelligent column widths
    for column_cells in data_sheet.columns:
        column_letter = column_cells[0].column_letter
        max_length = 0

        for cell in column_cells[:200]:
            if cell.value is not None:
                max_length = max(max_length, len(str(cell.value)))

        data_sheet.column_dimensions[column_letter].width = min(
            max(max_length + 3, 12),
            28,
        )

    # Numeric formatting
    headers = {
        cell.value: cell.column
        for cell in data_sheet[1]
    }

    if "price" in headers:
        col = data_sheet.cell(
            row=1,
            column=headers["price"],
        ).column_letter

        for cell in data_sheet[col][1:]:
            cell.number_format = '#,##0.00'

    # --------------------------------------------------------
    # SAVE
    # --------------------------------------------------------

    wb.save(report_file)

    # ========================================================
    # FINAL TERMINAL SUMMARY
    # ========================================================

    print("\n========================================")
    print("        PROCESSING COMPLETED")
    print("========================================")
    print(f"Rows processed:          {total_original}")
    print(f"Duplicates removed:      {total_duplicates}")
    print(f"Missing values found:    {total_missing}")
    print(f"Clean rows generated:    {total_final}")
    print(f"Worksheets processed:    {worksheet_count}")
    print("----------------------------------------")
    print(f"Professional report: {report_file}")
    print("========================================")

else:
    print("No valid CSV or Excel files were found.")
