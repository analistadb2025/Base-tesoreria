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

NOMBRES_SALIDA = ["consolidadobancos", "flujotesoreria", "seguimientorecaudo"]

MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
    5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
    9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}

COLUMN_MAP = {
    "categoria": ["categoria"],
    "concepto": ["concepto"],
    "valor": ["vlr flujo", "valor dc", "valor d/c", "vlr", "valor", "valor de la compra"],
    "fecha": ["fecha", "date", "fecha movimiento", "fecha valor", "fecha transaccion"],
}


def normalize(text):
    text = str(text).lower()
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-z0-9]", "", text)


def es_archivo_salida(nombre):
    n = normalize(nombre)
    return any(s in n for s in NOMBRES_SALIDA)


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
        if es_archivo_salida(name):
            continue
        data = upload.read()
        if name.lower().endswith(".zip"):
            with zipfile.ZipFile(io.BytesIO(data)) as z:
                for fname in z.namelist():
                    if fname.lower().endswith((".xlsx", ".xls")) and not fname.startswith("__"):
                        fname_base = fname.split("/")[-1]
                        if not es_archivo_salida(fname_base):
                            result.append((fname_base, z.read(fname)))
        elif name.lower().endswith((".xlsx", ".xls")):
            result.append((name, data))
    return result


def standardize_df(df, filename):
    cols = df.columns.tolist()
    n = len(df)

    categoria_col = find_column(cols, COLUMN_MAP["categoria"])
    concepto_col = find_column(cols, COLUMN_MAP["concepto"])
    valor_col = find_column(cols, COLUMN_MAP["valor"])
    fecha_col = find_column(cols, COLUMN_MAP["fecha"])

    categoria = df[categoria_col] if categoria_col else pd.Series([None] * n)
    concepto = df[concepto_col] if concepto_col else pd.Series([None] * n)
    total = clean_money(df[valor_col]) if valor_col else pd.Series([0.0] * n)

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
        "Total": total.values,
    })


def build_pivot_mes(df_mes, banco_cols):
    pivot = (
        df_mes
        .groupby(["Categoria", "Concepto", "Banco"])["Total"]
        .sum()
        .reset_index()
        .pivot_table(
            index=["Categoria", "Concepto"],
            columns="Banco",
            values="Total",
            fill_value=0,
        )
        .reset_index()
    )
    pivot.columns.name = None
    for b in banco_cols:
        if b not in pivot.columns:
            pivot[b] = 0.0
    pivot = pivot[["Categoria", "Concepto"] + banco_cols]
    pivot["Total"] = pivot[banco_cols].sum(axis=1)

    totals = {c: pivot[c].sum() for c in banco_cols + ["Total"]}
    totals["Categoria"] = "TOTAL"
    totals["Concepto"] = ""
    return pd.concat([pivot, pd.DataFrame([totals])], ignore_index=True)


def write_sheet(ws, df, header_color="1F4E79"):
    header_fill = PatternFill("solid", fgColor=header_color)
    header_font = Font(bold=True, color="FFFFFF", size=11)
    total_fill = PatternFill("solid", fgColor="D6E4F0")
    total_font = Font(bold=True, size=11)
    alt_fill = PatternFill("solid", fgColor="EBF3FB")
    border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    money_fmt = "#,##0.00"

    cols = list(df.columns)
    for col_idx, col_name in enumerate(cols, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    for row_idx, row in enumerate(df.itertuples(index=False), start=2):
        is_total = str(row[0]).upper() == "TOTAL"
        for col_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.border = border
            if is_total:
                cell.fill = total_fill
                cell.font = total_font
            elif row_idx % 2 == 0:
                cell.fill = alt_fill
            col_name = cols[col_idx - 1]
            if col_name not in ("Categoria", "Concepto") and isinstance(value, (int, float)):
                cell.number_format = money_fmt
                cell.alignment = Alignment(horizontal="right")

    for col_idx in range(1, len(cols) + 1):
        max_len = max(
            len(str(ws.cell(row=r, column=col_idx).value or ""))
            for r in range(1, ws.max_row + 1)
        )
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 40)

    ws.freeze_panes = "A2"


def build_excel_consolidado(data_por_mes, banco_cols):
    wb = Workbook()
    wb.remove(wb.active)

    for mes_label, df_mes in data_por_mes.items():
        pivot = build_pivot_mes(df_mes, banco_cols)
        ws = wb.create_sheet(title=mes_label[:31])
        write_sheet(ws, pivot)

    df_todo = pd.concat(list(data_por_mes.values()), ignore_index=True)
    pivot_total = build_pivot_mes(df_todo, banco_cols)
    ws_total = wb.create_sheet(title="Total")
    write_sheet(ws_total, pivot_total, header_color="163755")

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def run():
    st.title("Consolidado Bancos")
    st.markdown(
        "Sube los archivos Excel de los bancos **o un ZIP** que los contenga. "
        "Puedes subir varios meses a la vez. El Excel resultante tendra una hoja por mes."
    )

    uploads = st.file_uploader(
        "Suba los Excel o el ZIP con los archivos",
        type=["xlsx", "xls", "zip"],
        accept_multiple_files=True,
        key="cb_files",
    )

    if uploads:
        archivos = cargar_excels_desde_uploads(uploads)
        if not archivos:
            st.warning("No se encontraron archivos de bancos validos (se omiten archivos de consolidado/flujo).")
            return

        st.info("Archivos de banco encontrados: {}".format(len(archivos)))

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

        if not all_data:
            st.error("Ningun archivo aporto datos utiles.")
            return

        final_df = pd.concat(all_data, ignore_index=True)
        banco_cols = sorted(final_df["Banco"].dropna().unique().tolist())

        tiene_fechas = final_df["Fecha"].notna().any()
        if tiene_fechas:
            final_df = final_df.copy()
            final_df["Fecha_val"] = pd.to_datetime(final_df["Fecha"], errors="coerce")
            final_df["Mes_num"] = final_df["Fecha_val"].dt.month
            final_df["Anio"] = final_df["Fecha_val"].dt.year
            final_df["Mes"] = final_df["Mes_num"].map(MESES_ES)
            final_df["Mes_label"] = final_df.apply(
                lambda r: "{} {}".format(r["Mes"], int(r["Anio"]))
                if pd.notna(r["Fecha_val"]) else "Sin fecha",
                axis=1,
            )

            orden = (
                final_df[["Anio", "Mes_num", "Mes_label"]]
                .dropna()
                .drop_duplicates()
                .sort_values(["Anio", "Mes_num"])
            )
            meses_ordenados = orden["Mes_label"].tolist()
            if "Sin fecha" in final_df["Mes_label"].values:
                meses_ordenados.append("Sin fecha")

            data_por_mes = {}
            for mes in meses_ordenados:
                df_mes = final_df[final_df["Mes_label"] == mes]
                if not df_mes.empty:
                    data_por_mes[mes] = df_mes
        else:
            data_por_mes = {"Sin fecha": final_df}

        num_meses = len(data_por_mes)
        st.success(
            "Consolidado listo: {} mes(es), {} banco(s)".format(num_meses, len(banco_cols))
        )

        if num_meses <= 12:
            tabs = st.tabs(list(data_por_mes.keys()))
            for tab, (mes_label, df_mes) in zip(tabs, data_por_mes.items()):
                with tab:
                    pivot = build_pivot_mes(df_mes, banco_cols)
                    st.dataframe(pivot, use_container_width=True, hide_index=True)
        else:
            pivot_total = build_pivot_mes(final_df, banco_cols)
            st.dataframe(pivot_total, use_container_width=True, hide_index=True)

        excel_buf = build_excel_consolidado(data_por_mes, banco_cols)
        st.download_button(
            "Descargar Consolidado Bancos (hoja por mes)",
            excel_buf,
            "Consolidado_Bancos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )