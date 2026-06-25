import streamlit as st
import zipfile
import io
import pandas as pd
import re
import unicodedata
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def normalize(text):
    text = str(text).lower()
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-z0-9]", "", text)


def find_column(columns, keywords):
    normalized_cols = {normalize(col): col for col in columns}
    for key in keywords:
        key_norm = normalize(key)
        for col_norm, original_col in normalized_cols.items():
            if key_norm in col_norm or col_norm in key_norm:
                return original_col
    return None


def clean_money(series):
    return (
        series.astype(str)
        .str.replace("$", "", regex=False)
        .str.replace(",", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.strip()
        .pipe(pd.to_numeric, errors="coerce")
        .fillna(0)
    )


def is_header_row(row):
    text_cells = [str(x).strip() for x in row if isinstance(x, str) and len(str(x).strip()) > 0]
    short_texts = [t for t in text_cells if len(t) < 30]
    return len(short_texts) >= 4


def read_real_excel(file):
    df_raw = pd.read_excel(file, header=None)
    header_row = None
    for i, row in df_raw.iterrows():
        if is_header_row(row):
            header_row = i
            break
    if header_row is None:
        header_row = 0
    return pd.read_excel(file, header=header_row), header_row


COLUMN_MAP = {
    "categoria": ["categoria"],
    "concepto": ["concepto"],
    "valor": ["vlr flujo", "valor dc", "valor d/c", "vlr", "valor", "valor de la compra"],
    "fecha": ["fecha", "date", "fecha movimiento", "fecha valor", "fecha transaccion"],
}

MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
    5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
    9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}


def standardize_df(df, filename):
    cols = df.columns.tolist()
    n = len(df)

    categoria_col = find_column(cols, COLUMN_MAP["categoria"])
    concepto_col = find_column(cols, COLUMN_MAP["concepto"])
    valor_col = find_column(cols, COLUMN_MAP["valor"])
    fecha_col = find_column(cols, COLUMN_MAP["fecha"])

    categoria = df[categoria_col] if categoria_col else pd.Series([None] * n)
    concepto = df[concepto_col] if concepto_col else pd.Series([None] * n)
    total = clean_money(df[valor_col]) if valor_col else pd.Series([0] * n)

    if fecha_col:
        parsed = pd.to_datetime(df[fecha_col], errors="coerce", dayfirst=True)
        if hasattr(parsed.dt, "tz") and parsed.dt.tz is not None:
            parsed = parsed.dt.tz_convert(None)
        fecha = parsed.dt.normalize()
    else:
        fecha = pd.Series([pd.NaT] * n)

    banco = filename.split("/")[-1].replace(".xlsx", "")

    return pd.DataFrame({
        "Categoria": categoria.values,
        "Concepto": concepto.values,
        "Banco": banco,
        "Fecha": fecha.values,
        "Valor": total.values,
    })


def build_excel_flujo(pivot_df, meses_cols):
    wb = Workbook()
    ws = wb.active
    ws.title = "Flujo de Tesorería"

    header_font = Font(bold=True, color="FFFFFF", size=11)
    subtotal_fill = PatternFill("solid", fgColor="D6E4F0")
    subtotal_font = Font(bold=True, size=11)
    total_fill = PatternFill("solid", fgColor="BDD7EE")
    total_font = Font(bold=True, size=12)
    alt_fill = PatternFill("solid", fgColor="F2F7FB")
    border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    money_fmt = '#,##0.00'

    cols = list(pivot_df.columns)
    for col_idx, col_name in enumerate(cols, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = header_font
        cell.fill = PatternFill("solid", fgColor="1F4E79")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    row_idx = 2
    for _, row in pivot_df.iterrows():
        categoria = str(row.get("Categoria", "")).strip().upper()
        is_subtotal = categoria in ("TOTAL INGRESOS", "TOTAL EGRESOS", "FLUJO NETO", "SALDO ACUMULADO", "TOTAL OTROS")
        is_total_general = categoria in ("FLUJO NETO", "SALDO ACUMULADO")

        for col_idx, col_name in enumerate(cols, start=1):
            value = row[col_name]
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.border = border
            if is_total_general:
                cell.fill = total_fill
                cell.font = total_font
            elif is_subtotal:
                cell.fill = subtotal_fill
                cell.font = subtotal_font
            elif row_idx % 2 == 0:
                cell.fill = alt_fill
            if col_name not in ("Categoria", "Concepto") and isinstance(value, (int, float)):
                cell.number_format = money_fmt
                cell.alignment = Alignment(horizontal="right")

        row_idx += 1

    for col_idx in range(1, len(cols) + 1):
        max_len = max(
            len(str(ws.cell(row=r, column=col_idx).value or ""))
            for r in range(1, ws.max_row + 1)
        )
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 45)

    ws.freeze_panes = "C2"
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def run():
    st.title("Flujo de Tesorería")
    st.markdown(
        "Sube el ZIP con los archivos Excel de todos los bancos. "
        "El sistema consolida **todos los conceptos por mes** y genera el flujo de tesorería."
    )

    zip_file = st.file_uploader("Suba el ZIP con los Excel", type="zip", key="ft_zip")

    if zip_file is not None:
        all_data = []
        report = []

        with zipfile.ZipFile(io.BytesIO(zip_file.read())) as z:
            excel_files = [f for f in z.namelist() if f.lower().endswith(".xlsx") and not f.startswith("__")]
            st.info(f"Archivos encontrados: {len(excel_files)}")

            for file in excel_files:
                with z.open(file) as f:
                    try:
                        df, _ = read_real_excel(f)
                        clean_df = standardize_df(df, file)
                        report.append({"Archivo": file.split("/")[-1], "Filas": len(clean_df), "Estado": "✅ OK"})
                        all_data.append(clean_df)
                    except Exception as e:
                        report.append({"Archivo": file.split("/")[-1], "Filas": 0, "Estado": f"❌ {e}"})

        with st.expander("Reporte de lectura por archivo"):
            st.dataframe(pd.DataFrame(report), use_container_width=True)

        if all_data:
            final_df = pd.concat(all_data, ignore_index=True)
            final_df = final_df.dropna(subset=["Fecha"])

            if final_df.empty:
                st.warning("No hay datos con fechas válidas en los archivos.")
                return

            final_df["Mes_num"] = final_df["Fecha"].dt.month
            final_df["Año"] = final_df["Fecha"].dt.year
            final_df["Mes"] = final_df["Mes_num"].map(MESES_ES)

            meses_presentes = (
                final_df[["Año", "Mes_num", "Mes"]]
                .drop_duplicates()
                .sort_values(["Año", "Mes_num"])
            )
            meses_cols = [f"{row['Mes']} {row['Año']}" for _, row in meses_presentes.iterrows()]
            final_df["Mes_label"] = final_df.apply(lambda r: f"{r['Mes']} {r['Año']}", axis=1)

            grouped = (
                final_df
                .groupby(["Categoria", "Concepto", "Mes_label"], dropna=False)["Valor"]
                .sum()
                .reset_index()
            )

            pivot_df = grouped.pivot_table(
                index=["Categoria", "Concepto"],
                columns="Mes_label",
                values="Valor",
                fill_value=0,
            ).reset_index()
            pivot_df.columns.name = None

            existing_mes_cols = [m for m in meses_cols if m in pivot_df.columns]
            pivot_df = pivot_df[["Categoria", "Concepto"] + existing_mes_cols]
            pivot_df["Total"] = pivot_df[existing_mes_cols].sum(axis=1)

            mask_ingreso = pivot_df["Categoria"].apply(
                lambda x: normalize(str(x)) == normalize("Ingreso")
            )
            mask_egreso = pivot_df["Categoria"].apply(
                lambda x: normalize(str(x)) in [
                    normalize("Egreso"), normalize("Egresos"),
                    normalize("Gasto"), normalize("Gastos"), normalize("Salida")
                ]
            )

            ingresos_df = pivot_df[mask_ingreso].copy()
            egresos_df = pivot_df[mask_egreso].copy()
            otros_df = pivot_df[~mask_ingreso & ~mask_egreso].copy()

            num_cols = existing_mes_cols + ["Total"]

            def totals_row(df, label_cat, label_conc=""):
                row = {c: df[c].sum() for c in num_cols}
                row["Categoria"] = label_cat
                row["Concepto"] = label_conc
                return pd.DataFrame([row])

            sections = []
            if not ingresos_df.empty:
                sections.append(ingresos_df)
                sections.append(totals_row(ingresos_df, "TOTAL INGRESOS"))
            if not egresos_df.empty:
                sections.append(egresos_df)
                sections.append(totals_row(egresos_df, "TOTAL EGRESOS"))
            if not otros_df.empty:
                sections.append(otros_df)
                sections.append(totals_row(otros_df, "OTROS", "TOTAL OTROS"))

            if not ingresos_df.empty or not egresos_df.empty:
                flujo_row = {}
                for c in num_cols:
                    ing_val = ingresos_df[c].sum() if not ingresos_df.empty else 0
                    egr_val = egresos_df[c].sum() if not egresos_df.empty else 0
                    flujo_row[c] = ing_val - egr_val
                flujo_row["Categoria"] = "FLUJO NETO"
                flujo_row["Concepto"] = "Ingresos - Egresos"
                sections.append(pd.DataFrame([flujo_row]))

                saldo_row = {"Categoria": "SALDO ACUMULADO", "Concepto": "Acumulado del período"}
                acum = 0
                for c in existing_mes_cols:
                    acum += flujo_row[c]
                    saldo_row[c] = acum
                saldo_row["Total"] = flujo_row["Total"]
                sections.append(pd.DataFrame([saldo_row]))

            final_pivot = pd.concat(sections, ignore_index=True)
            col_order = ["Categoria", "Concepto"] + existing_mes_cols + ["Total"]
            final_pivot = final_pivot[col_order]

            st.success(f"Flujo de Tesorería listo: {len(existing_mes_cols)} mes(es) — {len(pivot_df)} conceptos")

            df_show = final_pivot.copy()
            for c in existing_mes_cols + ["Total"]:
                df_show[c] = df_show[c].apply(
                    lambda x: f"${x:,.2f}" if isinstance(x, (int, float)) else x
                )
            st.dataframe(df_show, use_container_width=True, hide_index=True)

            excel_buf = build_excel_flujo(final_pivot, existing_mes_cols)
            st.download_button(
                "⬇️ Descargar Flujo de Tesorería",
                excel_buf,
                "Flujo_de_Tesoreria.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.error("Ningún archivo aportó datos útiles.")
