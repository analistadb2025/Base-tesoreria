import streamlit as st

st.set_page_config(
    page_title="Procesos",
    page_icon="📊",
    layout="wide",
)

with st.sidebar:
    st.title("📊 Procesos")
    st.markdown("¿Qué proceso quieres usar?")
    proceso = st.selectbox(
        label="proceso",
        options=["Consolidado Bancos", "Seguimiento Recaudo Diario", "Flujo de Tesorería"],
        label_visibility="collapsed",
    )

if proceso == "Consolidado Bancos":
    from procesos.consolidadobancos import run
    run()
elif proceso == "Seguimiento Recaudo Diario":
    from procesos.Seguimientorecaudo import run
    run()
elif proceso == "Flujo de Tesorería":
    from procesos.Flujodetesoreria import run
    run()