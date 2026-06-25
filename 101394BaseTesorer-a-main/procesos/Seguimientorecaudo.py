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

RECAUDO_KEYWORDS = [
    "recaudo ventas", "recaudo de ventas", "recaudo", "ventas", "recaudo venta",
]


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


def build_excel_seguimiento(df_display):
    wb = Workbook()
    ws = wb.active
    ws.title = "Seguimiento Recaudo Diario"

    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    total_fill = PatternFill("solid", fgColor="BDD7EE")
    total_font = Font(bold=True, size=11)
    alt_fill = PatternFill("solid", fgColor="EBF3FB")
    border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    money_fmt = '#,##0.00'

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
                cell.number_format = 'DD/MM/YYYY'
                cell.alignment = Alignment(horizontal="center")
            elif col_name in ("Valor", "Total Diario", "Total Acumulado") and isinstance(value, (int, float)):
                cell.number_format = money_fmt
                cell.alignment = Alignment(horizontal="right")

    for col_idx in range(1, len(cols) + 1):
        max_len = max(
            len(str(ws.cell(row=r, column=col_idx).value or ""))
            for r in range(1, ws.max_row + 1)
        )
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 45)

    ws.freeze_panes = "A2"
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def run():
    st.title("Seguimiento Recaudo Diario")
    st.markdown(
        "Sube el ZIP con los archivos Excel de los bancos. "
        "El sistema filtra automáticamente los movimientos con **categoría = Ingreso** "
        "y **concepto = Recaudo Ventas**, agrupa por fecha y calcula el total diario."
    )

    zip_file = st.file_uploader("Suba el ZIP con los Excel", type="zip", key="sr_zip")

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

                        mask_cat = clean_df["Categoria"].apply(
                            lambda x: normalize(str(x)) == normalize("Ingreso")
                        )
                        mask_concepto = clean_df["Concepto"].apply(es_recaudo_ventas)
                        filtered = clean_df[mask_cat & mask_concepto].copy()

                        report.append({
                            "Archivo": file.split("/")[-1],
                            "Filas totales": len(clean_df),
                            "Filas recaudo": len(filtered),
                            "Estado": "✅ OK",
                        })
                        if not filtered.empty:
                            all_data.append(filtered)
                    except Exception as e:
                        report.append({
                            "Archivo": file.split("/")[-1],
                            "Filas totales": 0,
                            "Filas recaudo": 0,
                            "Estado": f"❌ {e}",
                        })

        with st.expander("Reporte de lectura por archivo"):
            st.dataframe(pd.DataFrame(report), use_container_width=True)

        if all_data:
            final_df = pd.concat(all_data, ignore_index=True)
            final_df = final_df.dropna(subset=["Fecha"])

            if final_df.empty:
                st.warning("No se encontraron movimientos de recaudo ventas con fechas válidas.")
            else:
                daily = (
                    final_df
                    .groupby("Fecha")["Valor"]
                    .sum()
                    .reset_index()
                    .rename(columns={"Valor": "Total Diario"})
                    .sort_values("Fecha")
                    .reset_index(drop=True)
                )
                daily["Concepto"] = "Recaudo Ventas"
                daily["Total Acumulado"] = daily["Total Diario"].cumsum()
                daily = daily[["Fecha", "Concepto", "Total Diario", "Total Acumulado"]]

                totals = {
                    "Fecha": "TOTAL",
                    "Concepto": "",
                    "Total Diario": daily["Total Diario"].sum(),
                    "Total Acumulado": daily["Total Acumulado"].iloc[-1],
                }
                df_display = pd.concat([daily, pd.DataFrame([totals])], ignore_index=True)

                st.success(
                    f"Seguimiento listo: {len(daily)} días con recaudo — "
                    f"Total: ${daily['Total Diario'].sum():,.0f}"
                )

                df_show = df_display.copy()
                df_show["Fecha"] = df_show["Fecha"].apply(
                    lambda x: x.strftime("%d/%m/%Y") if isinstance(x, pd.Timestamp) else x
                )
                df_show["Total Diario"] = df_show["Total Diario"].apply(
                    lambda x: f"${x:,.2f}" if isinstance(x, float) else x
                )
                df_show["Total Acumulado"] = df_show["Total Acumulado"].apply(
                    lambda x: f"${x:,.2f}" if isinstance(x, float) else x
                )
                st.dataframe(df_show, use_container_width=True, hide_index=True)

                excel_buf = build_excel_seguimiento(df_display)
                st.download_button(
                    "⬇️ Descargar Seguimiento Recaudo Diario",
                    excel_buf,
                    "Seguimiento_Recaudo_Diario.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        else:
            st.warning(
                "No se encontraron movimientos con categoría **Ingreso** y concepto **Recaudo Ventas**. "
                "Verifica que los Excel tengan esas columnas y valores."
            )