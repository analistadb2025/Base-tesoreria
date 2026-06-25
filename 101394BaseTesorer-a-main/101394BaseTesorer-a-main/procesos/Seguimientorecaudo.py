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

RECAUDO_KEYWORDS = [
    "recaudo ventas",
    "recaudo de ventas",
    "recaudo",
    "ventas",
    "recaudo venta",
]


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


def es_recaudo_ventas(concepto_str):
    norm = normalize(str(concepto_str))
    return any(normalize(kw) in norm for kw in RECAUDO_KEYWORDS)


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
        "Valor": total.values,
    })


def build_daily_for_month(df_month):
    # Crear tabla dinámica: filas = Fecha, columnas = Banco, valores = Valor
    daily = df_month.pivot_table(
        index="Fecha",
        columns="Banco",
        values="Valor",
        aggfunc="sum",
        fill_value=0
    ).reset_index()

    daily = daily.sort_values("Fecha").reset_index(drop=True)

    # Identificar las columnas de los bancos (todas las que no sean 'Fecha')
    bancos_cols = [c for c in daily.columns if c != "Fecha"]

    # Calcular Totales diarios y acumulados
    daily["Total Diario"] = daily[bancos_cols].sum(axis=1)
    daily["Total Acumulado"] = daily["Total Diario"].cumsum()

    # Construir la fila de Totales al final
    totals = {"Fecha": "TOTAL"}
    for banco in bancos_cols:
        totals[banco] = daily[banco].sum()
    totals["Total Diario"] = daily["Total Diario"].sum()
    totals["Total Acumulado"] = daily["Total Acumulado"].iloc[-1] if not daily.empty else 0

    return pd.concat([daily, pd.DataFrame([totals])], ignore_index=True)


def write_seguimiento_sheet(ws, df_display):
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    total_fill = PatternFill("solid", fgColor="BDD7EE")
    total_font = Font(bold=True, size=11)
    alt_fill = PatternFill("solid", fgColor="EBF3FB")
    border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    money_fmt = "#,##0.00"
    date_fmt = "DD/MM/YYYY"

    cols = list(df_display.columns)
    for col_idx, col_name in enumerate(cols, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    for row_idx, row in enumerate(df_display.itertuples(index=False), start=2):
        is_total = any(str(v).upper() == "TOTAL" for v in row)
        use_alt = (row_idx % 2 == 0) and not is_total
        for col_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.border = border
            if is_total:
                cell.fill = total_fill
                cell.font = total_font
            elif use_alt:
                cell.fill = alt_fill
            
            col_name = cols[col_idx - 1]
            
            if col_name == "Fecha" and not is_total and isinstance(value, pd.Timestamp):
                cell.number_format = date_fmt
                cell.alignment = Alignment(horizontal="center")
            # Aplicar formato de moneda a CUALQUIER columna que no sea Fecha
            elif col_name != "Fecha" and isinstance(value, (int, float)):
                cell.number_format = money_fmt
                cell.alignment = Alignment(horizontal="right")

    for col_idx in range(1, len(cols) + 1):
        max_len = max(
            len(str(ws.cell(row=r, column=col_idx).value or ""))
            for r in range(1, ws.max_row + 1)
        )
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 45)

    ws.freeze_panes = "A2"


def build_excel_seguimiento(data_por_mes):
    wb = Workbook()
    wb.remove(wb.active)

    all_daily = []
    for mes_label, df_mes in data_por_mes.items():
        df_display = build_daily_for_month(df_mes)
        ws = wb.create_sheet(title=mes_label[:31])
        write_seguimiento_sheet(ws, df_display)
        all_daily.append(df_mes)

    df_todo = pd.concat(all_daily, ignore_index=True)
    df_total_display = build_daily_for_month(df_todo)
    ws_total = wb.create_sheet(title="Total")
    write_seguimiento_sheet(ws_total, df_total_display)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def run():
    st.title("Seguimiento Recaudo Diario")
    st.markdown(
        "Sube los archivos Excel de los bancos **o un ZIP**. "
        "Filtra categoria = Ingreso y concepto = Recaudo Ventas. "
        "El Excel resultante tendra una hoja por mes + hoja Total con detalle por banco."
    )

    uploads = st.file_uploader(
        "Suba los Excel o el ZIP con los archivos",
        type=["xlsx", "xls", "zip"],
        accept_multiple_files=True,
        key="sr_files",
    )

    if uploads:
        archivos = cargar_excels_desde_uploads(uploads)
        if not archivos:
            st.warning("No se encontraron archivos de bancos validos.")
            return

        st.info("Archivos de banco encontrados: {}".format(len(archivos)))

        all_data = []
        report = []

        for nombre, data in archivos:
            try:
                df, _ = read_real_excel(io.BytesIO(data))
                clean_df = standardize_df(df, nombre)

                mask_cat = clean_df["Categoria"].apply(
                    lambda x: normalize(str(x)) == normalize("Ingreso")
                )
                mask_concepto = clean_df["Concepto"].apply(es_recaudo_ventas)
                filtered = clean_df[mask_cat & mask_concepto].copy()

                report.append({
                    "Archivo": nombre,
                    "Filas totales": len(clean_df),
                    "Filas recaudo": len(filtered),
                    "Estado": "OK",
                })
                if not filtered.empty:
                    all_data.append(filtered)
            except Exception as e:
                report.append({
                    "Archivo": nombre,
                    "Filas totales": 0,
                    "Filas recaudo": 0,
                    "Estado": "Error: {}".format(str(e)),
                })

        with st.expander("Reporte de lectura por archivo"):
            st.dataframe(pd.DataFrame(report), use_container_width=True)

        if not all_data:
            st.warning(
                "No se encontraron movimientos con categoria Ingreso y concepto Recaudo Ventas. "
                "Revisa el reporte de lectura para ver que columnas se detectaron."
            )
            return

        final_df = pd.concat(all_data, ignore_index=True)
        final_df = final_df.dropna(subset=["Fecha"])

        if final_df.empty:
            st.warning("No hay registros de recaudo con fechas validas.")
            return

        final_df = final_df.copy()
        final_df["Mes_num"] = final_df["Fecha"].dt.month
        final_df["Anio"] = final_df["Fecha"].dt.year
        final_df["Mes"] = final_df["Mes_num"].map(MESES_ES)
        final_df["Mes_label"] = final_df.apply(
            lambda r: "{} {}".format(r["Mes"], int(r["Anio"])), axis=1
        )

        orden = (
            final_df[["Anio", "Mes_num", "Mes_label"]]
            .drop_duplicates()
            .sort_values(["Anio", "Mes_num"])
        )
        meses_ordenados = orden["Mes_label"].tolist()

        data_por_mes = {}
        for mes in meses_ordenados:
            df_mes = final_df[final_df["Mes_label"] == mes]
            if not df_mes.empty:
                data_por_mes[mes] = df_mes

        total_recaudo = final_df["Valor"].sum()
        st.success(
            "Seguimiento listo: {} mes(es), {} dias con recaudo - Total: ${:,.0f}".format(
                len(data_por_mes),
                final_df["Fecha"].nunique(),
                total_recaudo,
            )
        )

        if len(data_por_mes) <= 12:
            tabs = st.tabs(list(data_por_mes.keys()))
            for tab, (mes_label, df_mes) in zip(tabs, data_por_mes.items()):
                with tab:
                    df_display = build_daily_for_month(df_mes)
                    df_show = df_display.copy()
                    
                    # Formatear la fecha
                    df_show["Fecha"] = df_show["Fecha"].apply(
                        lambda x: x.strftime("%d/%m/%Y") if isinstance(x, pd.Timestamp) else x
                    )
                    
                    # Formatear a moneda todas las columnas numéricas (Bancos y Totales)
                    cols_numericas = [c for c in df_show.columns if c != "Fecha"]
                    for col in cols_numericas:
                        df_show[col] = df_show[col].apply(
                            lambda x: "${:,.2f}".format(x) if isinstance(x, (int, float)) else x
                        )
                        
                    st.dataframe(df_show, use_container_width=True, hide_index=True)

        excel_buf = build_excel_seguimiento(data_por_mes)
        st.download_button(
            "Descargar Seguimiento Recaudo Diario (hoja por mes)",
            excel_buf,
            "Seguimiento_Recaudo_Diario.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )