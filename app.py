import streamlit as st
import pandas as pd

def procesar_excel_flexible(file):
    # Leer el excel completo inicialmente
    df_raw = pd.read_excel(file)
    
    # BUSCAR EL ENCABEZADO REAL:
    # Recorremos las primeras 10 filas buscando la palabra 'Artículo' o 'Stock'
    for i in range(len(df_raw)):
        fila_valores = df_raw.iloc[i].astype(str).values
        if any('Artículo' in x or 'Articulo' in x or 'Stock' in x for x in fila_valores):
            # Encontramos la fila del encabezado
            df_final = df_raw.iloc[i+1:].copy()
            df_final.columns = fila_valores
            return df_final.reset_index(drop=True)
    return df_raw

# --- INTERFAZ DE CARGA ---
st.title("Auditoría Grupo Cenoa - Clasificación ABC")
archivo_subido = st.file_uploader("Subir Reporte de Stock Jujuy", type=['xlsx'])

if archivo_subido:
    df = procesar_excel_flexible(archivo_subido)
    
    st.write("### Columnas detectadas:")
    cols = df.columns.tolist()
    
    # El usuario confirma cuáles son las columnas (por si cambian mañana)
    col_art = st.selectbox("Columna de Artículo", cols, index=cols.index('Artículo') if 'Artículo' in cols else 0)
    col_stock = st.selectbox("Columna de Stock", cols, index=cols.index('Stock') if 'Stock' in cols else 0)
    col_costo = st.selectbox("Columna de Costo", cols, index=cols.index('Cto.Rep.') if 'Cto.Rep.' in cols else 0)

    if st.button("Ejecutar Análisis ABC"):
        # Limpieza de datos: Asegurar que sean números
        df[col_stock] = pd.to_numeric(df[col_stock], errors='coerce').fillna(0)
        df[col_costo] = pd.to_numeric(df[col_costo], errors='coerce').fillna(0)
        
        # Cálculo del ABC
        df['Valor_Total'] = df[col_stock] * df[col_costo]
        # ... (aquí sigue el resto de tu lógica de clasificación)
        st.success("¡Análisis completado sin errores!")
        st.dataframe(df[[col_art, col_stock, 'Valor_Total']].head())
