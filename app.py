import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
import datetime

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Inventarios Rotativos - Cenoa", layout="wide", page_icon="📦")

# --- CONEXIÓN A GOOGLE SHEETS ---
# Nota: La URL no se usa directamente por conn; la usa la config en secrets.toml (connection "gsheets").
# Podés dejarla igual o eliminarla.
URL_SHEET = "https://docs.google.com/spreadsheets/d/1Dwn-uXcsT8CKFKwL0kZ4WyeVSwOGzXGcxMTW1W1bTe4/edit#gid=1078564738"
conn = st.connection("gsheets", type=GSheetsConnection)

# =========================
#  FUNCIONES GOOGLE SHEETS
# =========================

def _append_df_to_worksheet(worksheet: str, df_nuevo: pd.DataFrame):
    """
    Emula un append a una worksheet existente usando streamlit_gsheets:
    1) Lee lo existente
    2) Concatena
    3) Update (sobrescribe la hoja completa)
    """
    df_nuevo = df_nuevo.copy()

    try:
        df_existente = conn.read(worksheet=worksheet)

        # Si está vacía o sin estructura, escribimos directamente
        if df_existente is None or df_existente.empty:
            conn.update(worksheet=worksheet, data=df_nuevo)
            return

        # Normalizar columnas para que concat no desordene
        for col in df_existente.columns:
            if col not in df_nuevo.columns:
                df_nuevo[col] = ""
        for col in df_nuevo.columns:
            if col not in df_existente.columns:
                df_existente[col] = ""

        df_final = pd.concat([df_existente, df_nuevo[df_existente.columns]], ignore_index=True)
        conn.update(worksheet=worksheet, data=df_final)

    except Exception as e:
        st.error(f"Error guardando en la hoja '{worksheet}'.")
        st.exception(e)
        raise


def registrar_en_historial(id_inv: str, usuario: str):
    """Agrega una fila al historial general."""
    nueva_entrada = pd.DataFrame([{
        "ID_Inventario": id_inv,
        "Fecha": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "Auditor": usuario,
        "Estado": "Abierto"
    }])

    # Anti-duplicado por rerun/doble click
    try:
        df_hist = conn.read(worksheet="Historial_Inventarios")
        if df_hist is not None and not df_hist.empty and "ID_Inventario" in df_hist.columns:
            if (df_hist["ID_Inventario"].astype(str) == str(id_inv)).any():
                return
    except Exception:
        # Si falla la lectura, seguimos e intentamos guardar igual
        pass

    _append_df_to_worksheet("Historial_Inventarios", nueva_entrada)


def guardar_detalle_articulos(df_detalle: pd.DataFrame, id_inv: str):
    """Agrega al detalle los artículos muestreados para un inventario."""
    df_detalle = df_detalle.copy()
    df_detalle["ID_Inventario"] = id_inv

    _append_df_to_worksheet("Detalle_Articulos", df_detalle)

# =========================
#       UTILIDADES
# =========================

def limpiar_excel_cenoa(df_raw: pd.DataFrame) -> pd.DataFrame:
    """
    Busca la fila que contiene 'Artículo'/'Articulo' y la toma como encabezado.
    Devuelve la tabla limpia a partir de la fila siguiente.
    """
    for i in range(len(df_raw)):
        fila = [str(x).strip() for x in df_raw.iloc[i].tolist()]
        if "Artículo" in fila or "Articulo" in fila:
            df_limpio = df_raw.iloc[i + 1 :].copy()
            df_limpio.columns = fila
            return df_limpio.reset_index(drop=True)
    return df_raw

def login():
    with st.sidebar:
        st.title("🔐 Acceso Cenoa")
        user = st.selectbox(
            "Usuario",
            ["Seleccionar", "Diego Guantay", "Nancy Fernandez", "Gustavo Zambrano", "Admin", "Jefe de Repuestos"]
        )
        if user == "Seleccionar":
            return None, None
        rol = "Auditor" if user != "Jefe de Repuestos" else "Deposito"
        return user, rol

# =========================
#          APP
# =========================

usuario_actual, rol_actual = login()

if not usuario_actual:
    st.info("👋 Por favor, selecciona tu usuario en la barra lateral para comenzar.")
    st.stop()

st.title("Sistema de Inventarios Rotativos - Grupo Cenoa")

tab1, tab2, tab3, tab4 = st.tabs([
    "📂 1. Nuevo Inventario",
    "📝 2. Conteo Físico",
    "🛠 3. Justificaciones",
    "📊 4. KPIs e Historial"
])

# Columnas esperadas (según tu Excel)
c_art, c_loc, c_desc, c_stock, c_costo = "Artículo", "Locación", "Descripción", "Stock", "Cto.Rep."

# -------------------------
# TAB 1 - NUEVO INVENTARIO
# -------------------------
with tab1:
    if rol_actual != "Auditor":
        st.info("Esta sección es solo para Auditores.")
    else:
        st.header("Generar Muestra ABC")
        archivo = st.file_uploader("Subir Reporte de Stock (.xlsx)", type=["xlsx"])

        if archivo:
            df_input = pd.read_excel(archivo, header=None)
            df_base = limpiar_excel_cenoa(df_input)

            st.write("Vista previa del archivo interpretado:")
            st.dataframe(df_base.head(20))

            if st.button("Generar y Guardar en Plataforma"):
                # Validaciones mínimas
                faltantes = [c for c in [c_art, c_loc, c_desc, c_stock, c_costo] if c not in df_base.columns]
                if faltantes:
                    st.error(f"Faltan columnas en el Excel: {', '.join(faltantes)}")
                    st.stop()

                # Lógica ABC
                df_base = df_base.copy()
                df_base[c_stock] = pd.to_numeric(df_base[c_stock], errors="coerce").fillna(0)
                df_base[c_costo] = pd.to_numeric(df_base[c_costo], errors="coerce").fillna(0)

                # Evitar división por cero si el archivo viene vacío
                df_base["Valor_T"] = df_base[c_stock] * df_base[c_costo]
                total_valor = df_base["Valor_T"].sum()
                if total_valor <= 0:
                    st.error("No se puede calcular ABC: el total de Valor_T es 0. Revisá Stock y Cto.Rep.")
                    st.stop()

                df_base = df_base.sort_values("Valor_T", ascending=False)
                df_base["Acc"] = df_base["Valor_T"].cumsum() / total_valor
                df_base["Cat"] = df_base["Acc"].apply(lambda x: "A" if x <= 0.8 else ("B" if x <= 0.95 else "C"))

                # Muestra 85-10-5 (como lo tenías)
                df_a = df_base[df_base["Cat"] == "A"]
                df_b = df_base[df_base["Cat"] == "B"]
                df_c = df_base[df_base["Cat"] == "C"]

                m_a = df_a.sample(n=min(85, len(df_a))) if len(df_a) else df_a
                m_b = df_b.sample(n=min(10, len(df_b))) if len(df_b) else df_b
                m_c = df_c.sample(n=min(5, len(df_c))) if len(df_c) else df_c

                muestra = pd.concat([m_a, m_b, m_c], ignore_index=True)

                # Campos de seguimiento
                muestra["Conteo_Fisico"] = 0
                muestra["Diferencia"] = 0
                muestra["Justificacion"] = ""

                id_inv = datetime.datetime.now().strftime("INV-%Y%m%d-%H%M")

                # Guardado en Google Sheets (sin create)
                registrar_en_historial(id_inv, usuario_actual)
                guardar_detalle_articulos(muestra, id_inv)

                st.session_state["id_inv"] = id_inv
                st.success(f"✅ Inventario {id_inv} guardado correctamente en la nube.")
                st.dataframe(muestra[[c_loc, c_art, c_desc, "Cat"]])

# -------------------------
# TAB 2 - CONTEO
# -------------------------
with tab2:
    if rol_actual != "Auditor":
        st.info("Esta sección es solo para Auditores.")
    else:
        st.header("Carga de Conteo")
        st.write("Pendiente: leer inventarios abiertos y permitir edición de conteo.")

# -------------------------
# TAB 3 - JUSTIFICACIONES
# -------------------------
with tab3:
    if rol_actual != "Deposito":
        st.info("Esta sección es para Depósito / Jefe de Repuestos.")
    else:
        st.header("Justificación de Diferencias")
        st.info("Pendiente: mostrar solo artículos con diferencia para justificar.")

# -------------------------
# TAB 4 - KPIs
# -------------------------
with tab4:
    st.header("Panel de Control (Histórico)")

    if st.button("Actualizar Reporte de KPIs"):
        try:
            datos_historicos = conn.read(worksheet="Detalle_Articulos")
            if datos_historicos is None or datos_historicos.empty:
                st.info("No hay datos en Detalle_Articulos todavía.")
            else:
                st.write("Total de artículos auditados históricamente:", len(datos_historicos))
                # Acá podés agregar gráficos / pivots
                st.dataframe(datos_historicos.tail(50))
        except Exception as e:
            st.error("No se pudo leer Detalle_Articulos.")
            st.exception(e)

