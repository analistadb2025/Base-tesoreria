# -*- coding: utf-8 -*-
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


def cargar_excels_desde_uploads(uploads):
    result = []
    for upload in uploads:
        name = upload.name
        data = upload.read()
        if name.lower().endswith(".zip"):
            with zipfile.ZipFile(io.BytesIO(data)) as z:
                for fname in z.namelist():
                    if fname.lower().endswith((".xlsx", ".xls")) and not fname.startswith("__"):
                        result.append((fname.split("/")[-1], z.read(fname)))
        elif name.lower().endswith((".xlsx", ".xls")):
            result.append((name, data))
    return result


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

    banco = str(filename).replace(".xlsx", "").replace(".xls", "")

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
    ws.title = "Flujo de Tesoreria"

    header_font = Font(bold=True, color="FFFFFF", size=11)
    subtotal_fill = PatternFill("solid", fgColor="D6E4F0")
    subtotal_font = Font(bold=True, size=11)
    total_fill = PatternFill("solid", fgColor="BDD7EE")
    total_font = Font(bold=True, size=12)
    alt_fill = PatternFill("solid", fgColor="F2F7FB")
    border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    money_fmt = "#,##0.00"

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
        is_total_general = categoria in (
            "FLUJO NETO", "SALDO ACUMULADO", "TOTAL INGRESOS", "TOTAL EGRESOS"
        )

        for col_idx, col_name in enumerate(cols, start=1):
            value = row[col_name]
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.border = border

            if is_total_general:
                cell.fill = total_fill
                cell.font = total_font
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
    st.title("Flujo de Tesoreria")
    st.markdown(
        "Sube los archivos Excel de los bancos **o un ZIP** que los contenga. "
        "El sistema consolida todos los conceptos por mes y genera el flujo de tesoreria."
    )

    uploads = st.file_uploader(
        "Suba los Excel o el ZIP con los archivos",
        type=["xlsx", "xls", "zip"],
        accept_multiple_files=True,
        key="ft_files",
    )

    if uploads:
        archivos = cargar_excels_desde_uploads(uploads)
        st.info("Archivos encontrados: {}".format(len(archivos)))

        all_data = []
        report = []

        for nombre, data in archivos:
            try:
                df, _ = read_real_excel(io.BytesIO(data))
                clean_df = standardize_df(df, nombre)
                report.append({"Archivo": nombre, "Filas": len(clean_df), "Estado": "OK"})
                all_data.append(clean_df)
            except Exception as e:
                report.append({"Archivo": nombre, "Filas": 0, "Estado": "Error: {}".format(str(e))})

        with st.expander("Reporte de lectura por archivo"):
            st.dataframe(pd.DataFrame(report), use_container_width=True)

        if all_data:
            final_df = pd.concat(all_data, ignore_index=True)
            final_df = final_df.dropna(subset=["Fecha"])

            if final_df.empty:
                st.warning("No hay datos con fechas validas en los archivos.")
                return

            final_df["Mes_num"] = final_df["Fecha"].dt.month
            final_df["Anio"] = final_df["Fecha"].dt.year
            final_df["Mes"] = final_df["Mes_num"].map(MESES_ES)

            meses_presentes = (
                final_df[["Anio", "Mes_num", "Mes"]]
                .drop_duplicates()
                .sort_values(["Anio", "Mes_num"])
            )
            meses_cols = [
                "{} {}".format(row["Mes"], int(row["Anio"]))
                for _, row in meses_presentes.iterrows()
            ]
            final_df["Mes_label"] = final_df.apply(
                lambda r: "{} {}".format(r["Mes"], int(r["Anio"])), axis=1
            )

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
                    normalize("Egreso"),
                    normalize("Egresos"),
                    normalize("Gasto"),
                    normalize("Gastos"),
                    normalize("Salida"),
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

                saldo_row = {"Categoria": "SALDO ACUMULADO", "Concepto": "Acumulado del periodo"}
                acum = 0
                for c in existing_mes_cols:
                    acum += flujo_row[c]
                    saldo_row[c] = acum
                saldo_row["Total"] = flujo_row["Total"]
                sections.append(pd.DataFrame([saldo_row]))

            final_pivot = pd.concat(sections, ignore_index=True)
            col_order = ["Categoria", "Concepto"] + existing_mes_cols + ["Total"]
            final_pivot = final_pivot[col_order]

            st.success(
                "Flujo de Tesoreria listo: {} mes(es) - {} conceptos".format(
                    len(existing_mes_cols), len(pivot_df)
                )
            )

            df_show = final_pivot.copy()
            for c in existing_mes_cols + ["Total"]:
                df_show[c] = df_show[c].apply(
                    lambda x: "${:,.2f}".format(x) if isinstance(x, (int, float)) else x
                )
            st.dataframe(df_show, use_container_width=True, hide_index=True)

            excel_buf = build_excel_flujo(final_pivot, existing_mes_cols)
            st.download_button(
                "Descargar Flujo de Tesoreria",
                excel_buf,
                "Flujo_de_Tesoreria.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.error("Ningun archivo aporto datos utiles.")