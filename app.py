import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
import datetime

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Inventarios Rotativos - Cenoa", layout="wide", page_icon="📦")

# --- CONEXIÓN A BASE DE DATOS ---
# Importante: Asegúrate de que los Secrets en Streamlit tengan la misma URL
URL_SHEET = "https://docs.google.com/spreadsheets/d/1Dwn-uXcsT8CKFKwL0kZ4WyeVSwOGzXGcxMTW1W1bTe4/edit#gid=1078564738"
conn = st.connection("gsheets", type=GSheetsConnection)

# --- FUNCIONES DE PERSISTENCIA (GOOGLE SHEETS) ---

def registrar_en_historial(id_inv, usuario):
    """Crea una nueva fila en la pestaña de historial general"""
    nueva_entrada = pd.DataFrame([{
        "ID_Inventario": id_inv,
        "Fecha": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "Auditor": usuario,
        "Estado": "Abierto"
    }])
    # Append a la hoja Historial_Inventarios
    conn.create(worksheet="Historial_Inventarios", data=nueva_entrada)

def guardar_detalle_articulos(df_detalle, id_inv):
    """Guarda los 100 artículos seleccionados en la base de datos"""
    df_detalle["ID_Inventario"] = id_inv
    # Append a la hoja Detalle_Articulos
    conn.create(worksheet="Detalle_Articulos", data=df_detalle)

# --- FUNCIONES DE LIMPIEZA ---

def limpiar_excel_cenoa(df_raw):
    for i in range(len(df_raw)):
        fila = [str(x).strip() for x in df_raw.iloc[i].tolist()]
        if 'Artículo' in fila or 'Articulo' in fila:
            df_limpio = df_raw.iloc[i+1:].copy()
            df_limpio.columns = fila
            return df_limpio.reset_index(drop=True)
    return df_raw

# --- SISTEMA DE LOGUEO ---

def login():
    with st.sidebar:
        st.title("🔐 Acceso Cenoa")
        user = st.selectbox("Usuario", 
                           ["Seleccionar", "Diego Guantay", "Nancy Fernandez", "Gustavo Zambrano", "Admin", "Jefe de Repuestos"])
        if user == "Seleccionar": return None, None
        rol = "Auditor" if user != "Jefe de Repuestos" else "Deposito"
        return user, rol

usuario_actual, rol_actual = login()

# --- INTERFAZ PRINCIPAL ---

if not usuario_actual:
    st.info("👋 Por favor, selecciona tu usuario en la barra lateral para comenzar.")
else:
    st.title(f"Sistema de Inventarios Rotativos - Grupo Cenoa")
    
    tab1, tab2, tab3, tab4 = st.tabs([
        "📂 1. Nuevo Inventario", 
        "📝 2. Conteo Físico", 
        "🛠 3. Justificaciones", 
        "📊 4. KPIs e Historial"
    ])

    c_art, c_loc, c_desc, c_stock, c_costo = 'Artículo', 'Locación', 'Descripción', 'Stock', 'Cto.Rep.'

    # --- PESTAÑA 1: GENERACIÓN ---
    with tab1:
        if rol_actual == "Auditor":
            st.header("Generar Muestra ABC")
            archivo = st.file_uploader("Subir Reporte de Stock (.xlsx)", type=['xlsx'])
            
            if archivo:
                df_input = pd.read_excel(archivo, header=None)
                df_base = limpiar_excel_cenoa(df_input)
                
                if st.button("Generar y Guardar en Plataforma"):
                    # Lógica ABC
                    df_base[c_stock] = pd.to_numeric(df_base[c_stock], errors='coerce').fillna(0)
                    df_base[c_costo] = pd.to_numeric(df_base[c_costo], errors='coerce').fillna(0)
                    df_base['Valor_T'] = df_base[c_stock] * df_base[c_costo]
                    df_base = df_base.sort_values('Valor_T', ascending=False)
                    df_base['Acc'] = df_base['Valor_T'].cumsum() / df_base['Valor_T'].sum()
                    df_base['Cat'] = df_base['Acc'].apply(lambda x: 'A' if x<=0.8 else ('B' if x<=0.95 else 'C'))
                    
                    # Muestra 85-10-5
                    m_a = df_base[df_base['Cat']=='A'].sample(n=min(85, len(df_base[df_base['Cat']=='A'])))
                    m_b = df_base[df_base['Cat']=='B'].sample(n=min(10, len(df_base[df_base['Cat']=='B'])))
                    m_c = df_base[df_base['Cat']=='C'].sample(n=min(5, len(df_base[df_base['Cat']=='C'])))
                    muestra = pd.concat([m_a, m_b, m_c]).reset_index(drop=True)
                    
                    # Preparar campos de seguimiento
                    muestra['Conteo_Fisico'] = 0
                    muestra['Diferencia'] = 0
                    muestra['Justificacion'] = ""
                    
                    id_inv = datetime.datetime.now().strftime("INV-%Y%m%d-%H%M")
                    
                    # GUARDADO REAL EN GOOGLE SHEETS
                    registrar_en_historial(id_inv, usuario_actual)
                    guardar_detalle_articulos(muestra, id_inv)
                    
                    st.session_state['id_inv'] = id_inv
                    st.success(f"Inventario {id_inv} guardado correctamente en la nube.")
                    st.dataframe(muestra[[c_loc, c_art, c_desc, 'Cat']])

    # --- PESTAÑA 2: CONTEO ---
    with tab2:
        if rol_actual == "Auditor":
            st.header("Carga de Conteo")
            # Aquí podrías leer los inventarios abiertos de la hoja
            st.write("Seleccione el inventario generado para cargar los conteos.")
            # (Lógica de edición similar al paso anterior pero conectada a conn.read)

    # --- PESTAÑA 3: JUSTIFICACIONES ---
    with tab3:
        if rol_actual == "Deposito":
            st.header("Justificación de Diferencias")
            st.info("El Jefe de Repuestos verá aquí solo las diferencias para justificar.")

    # --- PESTAÑA 4: KPIs ---
    with tab4:
        st.header("Panel de Control (Histórico)")
        if st.button("Actualizar Reporte de KPIs"):
            datos_historicos = conn.read(worksheet="Detalle_Articulos")
            st.write("Total de artículos auditados históricamente:", len(datos_historicos))
            # Aquí se pueden agregar gráficos de barras por Categoria o Diferencias
