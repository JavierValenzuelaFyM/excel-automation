import pandas as pd
from pathlib import Path

input_folder = Path("input")
output_folder = Path("output")

output_folder.mkdir(exist_ok=True)


def clean_data(df):
    original_rows = len(df)
    duplicates = int(df.duplicated().sum())
    missing_values = int(df.isna().sum().sum())

    # Eliminar duplicados
    df = df.drop_duplicates().copy()

    # Tratar valores ausentes
    for column in df.columns:
        if df[column].dtype == "object":
            df[column] = df[column].fillna("Unknown")
        else:
            df[column] = df[column].fillna(0)

    # Ordenar si existe order_id
    if "order_id" in df.columns:
        df = df.sort_values("order_id")

    return df, original_rows, duplicates, missing_values


files = list(input_folder.glob("*.csv")) + list(input_folder.glob("*.xlsx"))

all_cleaned_data = []
summary = []


for file in files:

    print(f"\n===== Procesando {file.name} =====")

    # -------------------------
    # CSV
    # -------------------------
    if file.suffix == ".csv":

        df = pd.read_csv(file)

        df, original, duplicates, missing = clean_data(df)

        # Guardamos de dónde viene cada fila
        df["_source_file"] = file.name
        df["_source_sheet"] = "CSV"

        all_cleaned_data.append(df)

        output_file = output_folder / f"{file.stem}_clean.xlsx"
        df.to_excel(output_file, index=False)

        summary.append({
            "file": file.name,
            "sheet": "CSV",
            "original_rows": original,
            "duplicates_removed": duplicates,
            "missing_values_found": missing,
            "final_rows": len(df)
        })

        print("Filas originales:", original)
        print("Duplicados eliminados:", duplicates)
        print("Valores ausentes:", missing)
        print("Filas finales:", len(df))


    # -------------------------
    # EXCEL
    # -------------------------
    elif file.suffix == ".xlsx":

        sheets = pd.read_excel(file, sheet_name=None)

        output_file = output_folder / f"{file.stem}_clean.xlsx"

        with pd.ExcelWriter(output_file) as writer:

            for sheet_name, df in sheets.items():

                print(f"\nHoja: {sheet_name}")

                df, original, duplicates, missing = clean_data(df)

                # Copia para archivo individual
                df.to_excel(
                    writer,
                    sheet_name=sheet_name,
                    index=False
                )

                # Copia para tabla maestra
                master_df = df.copy()
                master_df["_source_file"] = file.name
                master_df["_source_sheet"] = sheet_name

                all_cleaned_data.append(master_df)

                summary.append({
                    "file": file.name,
                    "sheet": sheet_name,
                    "original_rows": original,
                    "duplicates_removed": duplicates,
                    "missing_values_found": missing,
                    "final_rows": len(df)
                })

                print("Filas originales:", original)
                print("Duplicados eliminados:", duplicates)
                print("Valores ausentes:", missing)
                print("Filas finales:", len(df))

        print(f"\nArchivo creado: {output_file}")


# ==================================================
# CREAR INFORME MAESTRO
# ==================================================

if all_cleaned_data:

    master_df = pd.concat(
        all_cleaned_data,
        ignore_index=True,
        sort=False
    )

    summary_df = pd.DataFrame(summary)

    report_file = output_folder / "master_report.xlsx"

    with pd.ExcelWriter(report_file) as writer:

        master_df.to_excel(
            writer,
            sheet_name="Cleaned_Data",
            index=False
        )

        summary_df.to_excel(
            writer,
            sheet_name="Processing_Summary",
            index=False
        )

    print("\n================================")
    print("PROCESAMIENTO COMPLETADO")
    print("================================")

    print("Total de filas limpias:", len(master_df))
    print("Archivos/hojas procesados:", len(summary_df))
    print(f"Informe maestro creado: {report_file}")

else:
    print("No se encontraron archivos para procesar.")