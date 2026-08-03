import streamlit as st
from utils.migracion_data import cargar_referencias, cargar_params

st.title("Test conexión Supabase")

refs, errores = cargar_referencias()
st.write("Errores:", errores)
for k, v in refs.items():
    st.write(f"**{k}**: {len(v)} filas")

st.subheader("Parámetros")
df = cargar_params()
st.write(f"{len(df)} períodos")
st.dataframe(df.tail(5))