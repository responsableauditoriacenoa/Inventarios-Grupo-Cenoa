import streamlit as st
import pandas as pd
import numpy as np

def clasificar_y_muestrear(df):
    # 1. Calcular valor total por artículo
    df['Valor_Total'] = df['Cantidad'] * df['Costo_Reposicion']
    df = df.sort_values(by='Valor_Total', ascending=False)
    
    # 2. Calcular porcentajes acumulados
    df['Pct_Acumulado'] = df['Valor_Total'].cumsum() / df['Valor_Total'].sum()
    
    # 3. Asignar ABC
    def asignar_abc(pct):
        if pct <= 0.80: return 'A'
        elif pct <= 0.95: return 'B'
        else: return 'C'
        
    df['Categoria'] = df['Pct_Acumulado'].apply(asignar_abc)
    
    # 4. Selección Aleatoria (80A, 15B, 5C)
    muestra_a = df[df['Categoria'] == 'A'].sample(n=min(80, len(df[df['Categoria'] == 'A'])))
    muestra_b = df[df['Categoria'] == 'B'].sample(n=min(15, len(df[df['Categoria'] == 'B'])))
    muestra_c = df[df['Categoria'] == 'C'].sample(n=min(5, len(df[df['Categoria'] == 'C'])))
    
    return pd.concat([muestra_a, muestra_b, muestra_c])

# Interfaz en Streamlit
st.title("Sistema de Inventarios Rotativos - Cenoa")

archivo = st.file_uploader("Subir reporte de stock (Excel)", type=['xlsx'])

if archivo:
    df_stock = pd.read_excel(archivo)
    if st.button("Generar Muestra Aleatoria ABC"):
        muestra = clasificar_y_muestrear(df_stock)
        st.write("Muestra Generada con Éxito")
        st.dataframe(muestra)
        # Aquí guardaríamos 'muestra' en Google Sheets para el siguiente paso
