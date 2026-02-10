import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
import gspread
import datetime
import io
import bcrypt
from usuarios_config import USUARIOS_CREDENCIALES, CREDENCIALES_INICIALES

# Version: 2.3 - Force rebuild with gspread (Feb 10, 2026)

# ----------------------------
# CONFIG
# ----------------------------
st.set_page_config(page_title="Inventarios Rotativos - Grupo Cenoa", layout="wide", page_icon="📦")
conn = st.connection("gsheets", type=GSheetsConnection)

# Obtener cliente gspread
client = conn.client

# ID del spreadsheet
SPREADSHEET_ID = "1Dwn-uXcsT8CKFKwL0kZ4WyeVSwOGzXGcxMTW1W1bTe4"

SHEET_HIST = "Historial_Inventarios"
SHEET_DET = "Detalle_Articulos"

# Columnas esperadas del Excel
C_ART = "Artículo"
C_LOC = "Locación"
C_DESC = "Descripción"
C_STOCK = "Stock"
C_COSTO = "Cto.Rep."

# Concesionarias y sucursales
CONCESIONARIAS = {
    "Autolux": ["Ax Jujuy", "Ax Salta", "Ax Tartagal", "Ax Lajitas", "Ax Taller Movil"],
    "Autosol": ["As Jujuy", "As Salta", "As Tartagal", "As Taller Express", "As Taller Movil"],
    "Ciel": ["Ac Jujuy"],
    "Portico": ["Las Lomas", "Brown"],
}

# ----------------------------
# HELPERS GSHEETS
# ----------------------------
def _read_ws(ws: str) -> pd.DataFrame:
    """
    Lee una worksheet usando gspread y la convierte a DataFrame
    """
    try:
        spreadsheet = client.open_by_key(SPREADSHEET_ID)
        worksheet = spreadsheet.worksheet(ws)
        data = worksheet.get_all_records()
        if not data:
            return pd.DataFrame()
        return pd.DataFrame(data)
    except Exception as e:
        st.error(f"Error al leer {ws}: {str(e)}")
        return pd.DataFrame()

def _update_ws(ws: str, df: pd.DataFrame):
    """
    Actualiza una worksheet con los datos de un DataFrame
    """
    try:
        spreadsheet = client.open_by_key(SPREADSHEET_ID)
        worksheet = spreadsheet.worksheet(ws)
        
        # Limpiar la worksheet
        worksheet.clear()
        
        # Escribir datos
        worksheet.update([df.columns.values.tolist()] + df.values.tolist())
    except Exception as e:
        st.error(f"Error al actualizar {ws}: {str(e)}")

def _append_df(ws: str, df_nuevo: pd.DataFrame):
    """Append emulado: read + concat + update"""
    df_exist = _read_ws(ws)
    if df_exist.empty:
        _update_ws(ws, df_nuevo)
        return

    df_nuevo = df_nuevo.copy()

    # Normalizar columnas
    for col in df_exist.columns:
        if col not in df_nuevo.columns:
            df_nuevo[col] = ""
    for col in df_nuevo.columns:
        if col not in df_exist.columns:
            df_exist[col] = ""

    df_final = pd.concat([df_exist, df_nuevo[df_exist.columns]], ignore_index=True)
    _update_ws(ws, df_final)

# ----------------------------
# AUTH CON USUARIO Y CONTRASEÑA
# ----------------------------
def verificar_password(password: str, password_hash: str) -> bool:
    """Verifica la contraseña contra el hash bcrypt"""
    return bcrypt.checkpw(password.encode(), password_hash.encode())

def login():
    """Sistema de login con usuario y contraseña"""
    # Mostrar logo en la parte superior
    col1, col2 = st.columns([1, 4])
    with col1:
        try:
            st.image("assets/logo_grupo_cenoa.png", width=100)
        except:
            # Si no existe la imagen local, mostrar un placeholder
            st.write("🏢 **GRUPO CENOA**")
    
    with col2:
        st.write("")  # Espacios en blanco
    
    st.title("🔐 Inventarios Rotativos - Grupo Cenoa")
    
    with st.form("login_form"):
        usuario = st.text_input("Usuario (ID):", placeholder="Ej: diego_guantay")
        contrasena = st.text_input("Contraseña:", type="password")
        submit = st.form_submit_button("Ingresar", use_container_width=True)
        
        if submit:
            if usuario in USUARIOS_CREDENCIALES:
                creds = USUARIOS_CREDENCIALES[usuario]
                # Verificar contraseña
                if verificar_password(contrasena, creds["password_hash"]):
                    st.session_state["logged_in"] = True
                    st.session_state["usuario"] = usuario
                    st.session_state["nombre_usuario"] = creds["nombre"]
                    st.session_state["rol"] = creds["rol"]
                    st.success(f"✅ Bienvenido {creds['nombre']}!")
                    st.rerun()
                else:
                    st.error("❌ Contraseña incorrecta")
            else:
                st.error("❌ Usuario no encontrado")
    
    # Mostrar tabla de credenciales (solo para pruebas - ELIMINAR EN PRODUCCIÓN)
    st.divider()
    st.write("**📋 CREDENCIALES DE PRUEBA** (Eliminar después de la primera vez)")
    
    creds_data = []
    for user_id, password in CREDENCIALES_INICIALES.items():
        rol = USUARIOS_CREDENCIALES[user_id]["rol"]
        nombre = USUARIOS_CREDENCIALES[user_id]["nombre"]
        creds_data.append({
            "Usuario (ID)": user_id,
            "Contraseña": password,
            "Rol": rol,
            "Nombre": nombre
        })
    
    df_creds = pd.DataFrame(creds_data)
    st.dataframe(df_creds, use_container_width=True, hide_index=True)
    st.info("⚠️ Estas credenciales son para pruebas. Cámbialas en producción.")
    
    return None, None

# Verificar si está logueado
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

if not st.session_state["logged_in"]:
    login()
    st.stop()

usuario_actual = st.session_state.get("usuario")
nombre_actual = st.session_state.get("nombre_usuario")
rol_actual = st.session_state.get("rol")

# ----------------------------
# DATA ACCESS
# ----------------------------
def listar_inventarios_abiertos():
    df_hist = _read_ws(SHEET_HIST)
    if df_hist.empty or "Estado" not in df_hist.columns:
        return pd.DataFrame()
    return df_hist[df_hist["Estado"].astype(str).str.lower() == "abierto"].copy()

def cargar_detalle(id_inv: str) -> pd.DataFrame:
    df = _read_ws(SHEET_DET)
    if df.empty or "ID_Inventario" not in df.columns:
        return pd.DataFrame()
    return df[df["ID_Inventario"].astype(str) == str(id_inv)].copy()

def calcular_resultados_inventario(df_det: pd.DataFrame) -> dict:
    """
    Calcula los resultados del inventario desde el detalle.
    Retorna dict con: cantidad_muestra, valor_muestra, 
    cantidad_faltantes, valor_faltantes,
    cantidad_sobrantes, valor_sobrantes,
    cantidad_dif_neta, valor_dif_neta,
    cantidad_dif_absoluta, valor_dif_absoluta,
    pct_absoluto, grado
    """
    if df_det.empty:
        return {}
    
    # Convertir a números
    df_r = df_det.copy()
    stock_col = C_STOCK if C_STOCK in df_r.columns else None
    costo_col = C_COSTO if C_COSTO in df_r.columns else None
    dif_col = "Diferencia"
    
    if not stock_col or not costo_col or dif_col not in df_r.columns:
        return {}
    
    df_r["_stock"] = pd.to_numeric(df_r[stock_col], errors="coerce").fillna(0)
    df_r["_costo"] = pd.to_numeric(df_r[costo_col], errors="coerce").fillna(0)
    df_r["_dif"] = pd.to_numeric(df_r[dif_col], errors="coerce").fillna(0)
    
    # Muestra
    cant_muestra = int(df_r["_stock"].sum())
    valor_muestra = (df_r["_stock"] * df_r["_costo"]).sum()
    
    # Faltantes (Diferencia < 0)
    mask_falt = df_r["_dif"] < 0
    cant_faltantes = int((df_r.loc[mask_falt, "_dif"].abs()).sum())
    valor_faltantes = (df_r.loc[mask_falt, "_dif"].abs() * df_r.loc[mask_falt, "_costo"]).sum()
    
    # Sobrantes (Diferencia > 0)
    mask_sobr = df_r["_dif"] > 0
    cant_sobrantes = int(df_r.loc[mask_sobr, "_dif"].sum())
    valor_sobrantes = (df_r.loc[mask_sobr, "_dif"] * df_r.loc[mask_sobr, "_costo"]).sum()
    
    # Diferencia neta (suma algebraica)
    cant_dif_neta = int(df_r["_dif"].sum())
    valor_dif_neta = (df_r["_dif"] * df_r["_costo"]).sum()
    
    # Diferencia absoluta (suma de valores absolutos)
    cant_dif_absoluta = int(df_r["_dif"].abs().sum())
    valor_dif_absoluta = (df_r["_dif"].abs() * df_r["_costo"]).sum()
    
    # Porcentaje absoluto y grado de cumplimiento
    pct_absoluto = (valor_dif_absoluta / valor_muestra * 100) if valor_muestra else 0
    
    escala = [(0.00, 100), (0.10, 94), (0.80, 82), (1.60, 65), (2.40, 35), (3.30, 0)]
    escala_sorted = sorted(escala, key=lambda x: x[0])
    grado = 0
    for th, g in escala_sorted:
        if pct_absoluto >= th:
            grado = g
    
    return {
        "cant_muestra": cant_muestra,
        "valor_muestra": valor_muestra,
        "cant_faltantes": cant_faltantes,
        "valor_faltantes": valor_faltantes,
        "cant_sobrantes": cant_sobrantes,
        "valor_sobrantes": valor_sobrantes,
        "cant_dif_neta": cant_dif_neta,
        "valor_dif_neta": valor_dif_neta,
        "cant_dif_absoluta": cant_dif_absoluta,
        "valor_dif_absoluta": valor_dif_absoluta,
        "pct_absoluto": pct_absoluto,
        "grado": grado,
        "escala": escala_sorted
    }

def guardar_detalle_modificado(id_inv: str, df_mod: pd.DataFrame):
    """Reemplaza solo las filas de ese inventario dentro del Detalle_Articulos."""
    df_all = _read_ws(SHEET_DET)
    if df_all.empty:
        _update_ws(SHEET_DET, df_mod)
        return

    df_all = df_all.copy()
    mask = df_all["ID_Inventario"].astype(str) == str(id_inv)
    df_rest = df_all.loc[~mask].copy()
    df_final = pd.concat([df_rest, df_mod], ignore_index=True)
    _update_ws(SHEET_DET, df_final)

def cerrar_inventario(id_inv: str, usuario: str):
    df_hist = _read_ws(SHEET_HIST)
    if df_hist.empty or "ID_Inventario" not in df_hist.columns:
        return
    df_hist = df_hist.copy()
    mask = df_hist["ID_Inventario"].astype(str) == str(id_inv)
    if mask.sum() == 0:
        return
    df_hist.loc[mask, "Estado"] = "Cerrado"
    df_hist.loc[mask, "Cierre_Fecha"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    df_hist.loc[mask, "Cierre_Usuario"] = usuario
    _update_ws(SHEET_HIST, df_hist)

# ----------------------------
# UI
# ----------------------------
# Sidebar con info de usuario y logout
with st.sidebar:
    st.write("---")
    st.write(f"**👤 Logueado como:** {nombre_actual}")
    st.write(f"**🎯 Rol:** {rol_actual}")
    
    if st.button("🚪 Cerrar sesión", use_container_width=True):
        st.session_state["logged_in"] = False
        st.session_state.clear()
        st.rerun()

st.title("📦 Inventarios Rotativos - Auditoría Interna (Grupo Cenoa)")

tab1, tab2, tab3, tab4 = st.tabs([
    "1) Nuevo inventario",
    "2) Conteo físico (Auditor)",
    "3) Justificaciones (Depósito / Auditor)",
    "4) Cierre + Reporte"
])

# ----------------------------
# TAB 1: NUEVO INVENTARIO
# ----------------------------
with tab1:
    st.subheader("Panel de control del Auditor (antes de importar)")

    c1, c2 = st.columns(2)
    with c1:
        concesionaria = st.selectbox("Concesionaria", list(CONCESIONARIAS.keys()))
    with c2:
        sucursal = st.selectbox("Sucursal", CONCESIONARIAS[concesionaria])

    st.divider()

    if rol_actual != "Auditor":
        st.info("Solo Auditores pueden generar inventarios.")
    else:
        st.subheader("Importar Excel → ABC → Muestra 80/15/5 → Guardar")

        archivo = st.file_uploader("Subir reporte de stock (.xlsx)", type=["xlsx"])

        if archivo:
            df_base = pd.read_excel(archivo)

            st.write("Vista previa del reporte:")
            st.dataframe(df_base.head(15), use_container_width=True)

            if st.button("✅ Generar y guardar inventario"):
                falt = [c for c in [C_ART, C_LOC, C_DESC, C_STOCK, C_COSTO] if c not in df_base.columns]
                if falt:
                    st.error(f"Faltan columnas en el Excel: {', '.join(falt)}")
                    st.stop()

                df = df_base.copy()
                df[C_STOCK] = pd.to_numeric(df[C_STOCK], errors="coerce").fillna(0)
                df[C_COSTO] = pd.to_numeric(df[C_COSTO], errors="coerce").fillna(0)

                df["Valor_T"] = df[C_STOCK] * df[C_COSTO]
                total = df["Valor_T"].sum()
                if total <= 0:
                    st.error("No se puede calcular ABC: Valor_T total es 0. Revisá Stock y Cto.Rep.")
                    st.stop()

                df = df.sort_values("Valor_T", ascending=False)
                df["Acc"] = df["Valor_T"].cumsum() / total
                df["Cat"] = df["Acc"].apply(lambda x: "A" if x <= 0.8 else ("B" if x <= 0.95 else "C"))

                df_a = df[df["Cat"] == "A"]
                df_b = df[df["Cat"] == "B"]
                df_c = df[df["Cat"] == "C"]

                m_a = df_a.sample(n=min(80, len(df_a))) if len(df_a) else df_a
                m_b = df_b.sample(n=min(15, len(df_b))) if len(df_b) else df_b
                m_c = df_c.sample(n=min(5, len(df_c))) if len(df_c) else df_c

                muestra = pd.concat([m_a, m_b, m_c], ignore_index=True)

                muestra["Concesionaria"] = concesionaria
                muestra["Sucursal"] = sucursal
                muestra["Conteo_Fisico"] = ""
                muestra["Diferencia"] = ""
                muestra["Justificacion"] = ""
                muestra["Justif_Validada"] = ""
                muestra["Validador"] = ""
                muestra["Fecha_Validacion"] = ""

                id_inv = datetime.datetime.now().strftime("INV-%Y%m%d-%H%M")

                df_hist = _read_ws(SHEET_HIST)
                if not df_hist.empty and "ID_Inventario" in df_hist.columns:
                    if (df_hist["ID_Inventario"].astype(str) == id_inv).any():
                        st.warning("Este ID ya existe (rerun). Probá otra vez.")
                        st.stop()

                nueva_fila = pd.DataFrame([{
                    "ID_Inventario": id_inv,
                    "Fecha": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "Concesionaria": concesionaria,
                    "Sucursal": sucursal,
                    "Auditor": usuario_actual,
                    "Estado": "Abierto",
                    "Cierre_Fecha": "",
                    "Cierre_Usuario": ""
                }])
                _append_df(SHEET_HIST, nueva_fila)

                muestra["ID_Inventario"] = id_inv
                _append_df(SHEET_DET, muestra)

                st.success(f"✅ Inventario {id_inv} creado y guardado.")
                st.session_state["id_inv"] = id_inv
                st.dataframe(muestra[[C_LOC, C_ART, C_DESC, "Cat"]], use_container_width=True)

                # Forzar refresco general
                st.rerun()

# ----------------------------
# TAB 2: CONTEO (AUDITOR)
# ----------------------------
with tab2:
    st.subheader("Carga de conteo físico")

    if rol_actual != "Auditor":
        st.info("Solo Auditores pueden cargar conteos.")
    else:
        df_abiertos = listar_inventarios_abiertos()
        if df_abiertos.empty:
            st.info("No hay inventarios abiertos.")
        else:
            id_sel = st.selectbox("Seleccionar inventario", df_abiertos["ID_Inventario"].astype(str).tolist())

            df_det = cargar_detalle(id_sel)
            if df_det.empty:
                st.warning("No hay detalle para ese inventario.")
            else:
                st.caption("Cargá Conteo_Fisico y guardá para recalcular diferencias.")

                cols_show = ["Concesionaria","Sucursal",C_LOC,C_ART,C_DESC,C_STOCK,C_COSTO,"Cat","Conteo_Fisico","Diferencia"]
                cols_show = [c for c in cols_show if c in df_det.columns]
                df_edit = df_det[cols_show].copy()

                edited = st.data_editor(
                    df_edit,
                    use_container_width=True,
                    num_rows="fixed",
                    disabled=[c for c in df_edit.columns if c != "Conteo_Fisico"],
                    key=f"conteo_{id_sel}",
                )

                if st.button("💾 Guardar conteo y recalcular diferencias"):
                    df_det2 = df_det.copy()

                    key_cols = [C_ART, C_LOC]
                    if not all(c in df_det2.columns for c in key_cols):
                        st.error("No encuentro columnas para matchear (Artículo/Locación).")
                        st.stop()

                    edited2 = edited.copy()
                    for c in key_cols:
                        edited2[c] = edited2[c].astype(str)
                        df_det2[c] = df_det2[c].astype(str)

                    df_merge = df_det2.merge(
                        edited2[key_cols + ["Conteo_Fisico"]],
                        on=key_cols,
                        how="left",
                        suffixes=("", "_new")
                    )

                    df_merge["Conteo_Fisico"] = df_merge["Conteo_Fisico_new"].combine_first(df_merge.get("Conteo_Fisico"))
                    if "Conteo_Fisico_new" in df_merge.columns:
                        df_merge = df_merge.drop(columns=["Conteo_Fisico_new"])

                    stock_num = pd.to_numeric(df_merge[C_STOCK], errors="coerce").fillna(0)
                    conteo_num = pd.to_numeric(df_merge["Conteo_Fisico"], errors="coerce").fillna(0)
                    df_merge["Diferencia"] = conteo_num - stock_num

                    guardar_detalle_modificado(id_sel, df_merge)
                    st.success("✅ Conteo guardado y diferencias recalculadas.")

                    # CLAVE: forzar refresco para que Tab 3 lea las diferencias
                    st.rerun()

# ----------------------------
# TAB 3: JUSTIFICACIONES (DEPÓSITO + VALIDACIÓN AUDITOR)
# ----------------------------
with tab3:
    st.subheader("Justificaciones y validación")

    if st.button("🔄 Refrescar diferencias"):
        st.rerun()

    df_abiertos = listar_inventarios_abiertos()
    if df_abiertos.empty:
        st.info("No hay inventarios abiertos.")
    else:
        id_sel = st.selectbox("Seleccionar inventario", df_abiertos["ID_Inventario"].astype(str).tolist(), key="sel_just")
        df_det = cargar_detalle(id_sel)

        if df_det.empty:
            st.warning("No hay detalle para ese inventario.")
        else:
            dif_num = pd.to_numeric(df_det.get("Diferencia", 0), errors="coerce").fillna(0)
            df_dif = df_det.loc[dif_num != 0].copy()

            if df_dif.empty:
                st.success("No hay diferencias para justificar (o todavía no guardaste conteo).")
            else:
                base_cols = ["Concesionaria","Sucursal",C_LOC,C_ART,C_DESC,C_STOCK,"Conteo_Fisico","Diferencia","Justificacion","Justif_Validada"]
                base_cols = [c for c in base_cols if c in df_dif.columns]
                df_view = df_dif[base_cols].copy()

                if rol_actual == "Deposito":
                    st.caption("Depósito: completá Justificacion y guardá.")
                    
                    # Mostrar tabla sin editar, con columnas de entrada separadas
                    display_cols = [c for c in base_cols if c != "Justificacion"]
                    st.dataframe(df_view[display_cols], use_container_width=True, hide_index=True)
                    
                    st.write("**Ingresá las justificaciones:**")
                    justificaciones_dict = {}
                    
                    for idx, row in df_view.iterrows():
                        col_key = f"{row[C_ART]}_{row[C_LOC]}"
                        art = row[C_ART]
                        loc = row[C_LOC]
                        dif = row["Diferencia"]
                        just_actual = row.get("Justificacion", "")
                        
                        st.write(f"**{art} ({loc}) - Diferencia: {dif}**")
                        justificacion = st.text_area(
                            f"Justificación",
                            value=just_actual,
                            height=80,
                            key=f"just_{col_key}"
                        )
                        justificaciones_dict[col_key] = {
                            "Articulo": art,
                            "Locacion": loc,
                            "Justificacion": justificacion
                        }
                        st.divider()
                    
                    if st.button("💾 Guardar justificaciones (Depósito)"):
                        df_det2 = df_det.copy()
                        
                        for col_key, data in justificaciones_dict.items():
                            art = data["Articulo"]
                            loc = data["Locacion"]
                            just = data["Justificacion"]
                            
                            mask = (df_det2[C_ART].astype(str) == str(art)) & (df_det2[C_LOC].astype(str) == str(loc))
                            df_det2.loc[mask, "Justificacion"] = just
                        
                        guardar_detalle_modificado(id_sel, df_det2)
                        st.success("✅ Justificaciones guardadas.")
                        st.rerun()

                else:
                    st.caption("Auditor: marcá Justif_Validada (SI/NO) y guardá.")
                    
                    # Mostrar tabla sin editar, con columnas de entrada separadas
                    display_cols = [c for c in base_cols if c != "Justif_Validada"]
                    st.dataframe(df_view[display_cols], use_container_width=True, hide_index=True)
                    
                    st.write("**Validá las justificaciones:**")
                    validaciones_dict = {}
                    
                    for idx, row in df_view.iterrows():
                        col_key = f"{row[C_ART]}_{row[C_LOC]}"
                        art = row[C_ART]
                        loc = row[C_LOC]
                        dif = row["Diferencia"]
                        just = row.get("Justificacion", "")
                        val_actual = row.get("Justif_Validada", "")
                        
                        st.write(f"**{art} ({loc}) - Diferencia: {dif}**")
                        st.write(f"*Justificación: {just if just else '(sin justificación)'}*")
                        
                        validacion = st.selectbox(
                            "¿Está validada?",
                            options=["", "SI", "NO"],
                            index=(["", "SI", "NO"].index(val_actual) if val_actual in ["SI", "NO"] else 0),
                            key=f"val_{col_key}"
                        )
                        validaciones_dict[col_key] = {
                            "Articulo": art,
                            "Locacion": loc,
                            "Justif_Validada": validacion
                        }
                        st.divider()
                    
                    if st.button("💾 Guardar validación (Auditor)"):
                        df_det2 = df_det.copy()
                        
                        for col_key, data in validaciones_dict.items():
                            art = data["Articulo"]
                            loc = data["Locacion"]
                            val = data["Justif_Validada"]
                            
                            mask = (df_det2[C_ART].astype(str) == str(art)) & (df_det2[C_LOC].astype(str) == str(loc))
                            df_det2.loc[mask, "Justif_Validada"] = val
                            df_det2.loc[mask, "Validador"] = usuario_actual
                            df_det2.loc[mask, "Fecha_Validacion"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
                        
                        guardar_detalle_modificado(id_sel, df_det2)
                        st.success("✅ Validación guardada.")
                        st.rerun()

# ----------------------------
# TAB 4: CIERRE + REPORTE
# ----------------------------
with tab4:
    st.subheader("Cierre de inventario + reporte")

    if rol_actual != "Auditor":
        st.info("Solo Auditores pueden cerrar inventarios.")
    else:
        df_abiertos = listar_inventarios_abiertos()
        if df_abiertos.empty:
            st.info("No hay inventarios abiertos.")
        else:
            id_sel = st.selectbox("Seleccionar inventario a cerrar", df_abiertos["ID_Inventario"].astype(str).tolist(), key="sel_cierre")
            df_det = cargar_detalle(id_sel)

            if df_det.empty:
                st.warning("No hay detalle para ese inventario.")
            else:
                # Calcular resultados desde el inventario
                resultados = calcular_resultados_inventario(df_det)
                
                if not resultados:
                    st.error("No se pudieron calcular los resultados. Verifica las columnas del inventario.")
                    st.stop()
                
                # Mostrar preview visual de los resultados
                st.write("### 📊 Paneo Visual de Resultados")
                
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("Muestra (Qty)", resultados["cant_muestra"])
                col2.metric("Faltantes (Qty)", resultados["cant_faltantes"])
                col3.metric("Sobrantes (Qty)", resultados["cant_sobrantes"])
                col4.metric("Grado de Cumplimiento", f"{resultados['grado']}%")
                
                # Tabla de resultados
                st.write("### 📋 Tabla Detallada")
                tabla_resultados = pd.DataFrame([
                    {
                        "Detalle": "Muestra",
                        "Cant. de Art.": resultados["cant_muestra"],
                        "$": resultados["valor_muestra"],
                        "%": 1.0
                    },
                    {
                        "Detalle": "Faltantes",
                        "Cant. de Art.": resultados["cant_faltantes"],
                        "$": resultados["valor_faltantes"],
                        "%": resultados["valor_faltantes"] / resultados["valor_muestra"] if resultados["valor_muestra"] else 0
                    },
                    {
                        "Detalle": "Sobrantes",
                        "Cant. de Art.": resultados["cant_sobrantes"],
                        "$": resultados["valor_sobrantes"],
                        "%": resultados["valor_sobrantes"] / resultados["valor_muestra"] if resultados["valor_muestra"] else 0
                    },
                    {
                        "Detalle": "Diferencia Neta",
                        "Cant. de Art.": resultados["cant_dif_neta"],
                        "$": resultados["valor_dif_neta"],
                        "%": resultados["valor_dif_neta"] / resultados["valor_muestra"] if resultados["valor_muestra"] else 0
                    },
                    {
                        "Detalle": "Diferencia Absoluta",
                        "Cant. de Art.": resultados["cant_dif_absoluta"],
                        "$": resultados["valor_dif_absoluta"],
                        "%": resultados["valor_dif_absoluta"] / resultados["valor_muestra"] if resultados["valor_muestra"] else 0
                    }
                ])
                
                # Formatear la tabla para mostrar
                tabla_display = tabla_resultados.copy()
                tabla_display["$"] = tabla_display["$"].apply(lambda x: f"${x:,.2f}")
                tabla_display["%"] = tabla_display["%"].apply(lambda x: f"{x*100:.2f}%")
                
                st.dataframe(tabla_display, use_container_width=True, hide_index=True)
                
                # Validación
                dif = pd.to_numeric(df_det.get("Diferencia", 0), errors="coerce").fillna(0)
                difmask = dif != 0
                val = df_det.get("Justif_Validada", "").astype(str)
                ok_validacion = bool(((~difmask) | (val.str.strip() != "")).all())

                if not ok_validacion:
                    st.warning("⚠️ Hay diferencias sin validar (Justif_Validada vacío).")

                st.divider()
                st.write("### 📥 Descargar reporte (.xlsx):")

                def build_report_xlsx():
                    from openpyxl import Workbook
                    from openpyxl.utils.dataframe import dataframe_to_rows
                    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

                    buffer = io.BytesIO()
                    
                    # Usar los valores calculados
                    muestra_cnt = resultados["cant_muestra"]
                    valor_muestra = resultados["valor_muestra"]
                    cant_faltantes = resultados["cant_faltantes"]
                    value_faltantes = resultados["valor_faltantes"]
                    cant_sobrantes = resultados["cant_sobrantes"]
                    value_sobrantes = resultados["valor_sobrantes"]
                    cant_dif_neta = resultados["cant_dif_neta"]
                    value_neta = resultados["valor_dif_neta"]
                    cant_dif_absoluta = resultados["cant_dif_absoluta"]
                    value_absoluta = resultados["valor_dif_absoluta"]
                    pct_absoluto = resultados["pct_absoluto"]
                    grado = resultados["grado"]
                    escala_sorted = resultados["escala"]

                    # Crear workbook
                    wb = Workbook()
                    ws = wb.active
                    ws.title = "Resultado"

                    # Estilos
                    title_font = Font(size=14, bold=True)
                    light_red = PatternFill(start_color="FFF2F2", end_color="FFF2F2", fill_type="solid")
                    bold = Font(bold=True)
                    center = Alignment(horizontal="center", vertical="center")
                    money_fmt = "#,##0.00"
                    thin = Side(border_style="thin", color="000000")
                    border = Border(left=thin, right=thin, top=thin, bottom=thin)

                    # Título
                    ws.merge_cells("A1:D1")
                    ws["A1"] = "4. Resultado Inventario Rotativo"
                    ws["A1"].font = title_font
                    ws["A1"].alignment = center

                    # Descripción
                    ws["A3"] = "El resultado del inventario rotativo es el siguiente:"

                    # Encabezados
                    start_row = 5
                    ws[f"A{start_row}"] = "Detalle"
                    ws[f"B{start_row}"] = "Cant. de Art."
                    ws[f"C{start_row}"] = "$"
                    ws[f"D{start_row}"] = "%"
                    for col in ["A","B","C","D"]:
                        cell = ws[f"{col}{start_row}"]
                        cell.font = bold
                        cell.alignment = center
                        cell.border = border

                    # Filas de datos
                    rows = [
                        ("Muestra", muestra_cnt, valor_muestra, 1.0),
                        ("Faltantes", cant_faltantes, value_faltantes, (value_faltantes / valor_muestra) if valor_muestra else 0),
                        ("Sobrantes", cant_sobrantes, value_sobrantes, (value_sobrantes / valor_muestra) if valor_muestra else 0),
                        ("Diferencia Neta", cant_dif_neta, value_neta, (value_neta / valor_muestra) if valor_muestra else 0),
                        ("Diferencia Absoluta", cant_dif_absoluta, value_absoluta, (value_absoluta / valor_muestra) if valor_muestra else 0),
                    ]

                    for i, r in enumerate(rows, start=start_row + 1):
                        ws[f"A{i}"] = r[0]
                        ws[f"B{i}"] = r[1]
                        ws[f"C{i}"] = r[2]
                        ws[f"D{i}"] = r[3]
                        ws[f"B{i}"].alignment = center
                        ws[f"C{i}"].number_format = money_fmt
                        ws[f"D{i}"].number_format = "0.00%"
                        ws[f"A{i}"].border = border
                        ws[f"B{i}"].border = border
                        ws[f"C{i}"].border = border
                        ws[f"D{i}"].border = border

                    # Resaltar Diferencia Neta y Absoluta
                    diff_rows = [start_row + 3, start_row + 4]
                    for r in diff_rows:
                        for col in ["B","C","D"]:
                            ws[f"{col}{r}"].fill = light_red

                    # Grado de cumplimiento
                    pct_cell_row = start_row + 1
                    ws.merge_cells(f"F{pct_cell_row}:G{pct_cell_row}")
                    ws[f"F{pct_cell_row}"] = f"{grado}%"
                    ws[f"F{pct_cell_row}"].font = Font(size=12, bold=True)
                    ws[f"F{pct_cell_row}"].alignment = center

                    # Tabla de escala
                    escala_start = start_row + 7
                    ws[f"B{escala_start}"] = "Dif. Abs. desde"
                    ws[f"C{escala_start}"] = "Grado de cumplim."
                    ws[f"B{escala_start}"].font = bold
                    ws[f"C{escala_start}"].font = bold

                    for j, (th, g) in enumerate(escala_sorted, start=escala_start + 1):
                        ws[f"B{j}"] = f"{th:.2f}%"
                        ws[f"C{j}"] = f"{g}%"
                        ws[f"B{j}"].alignment = center
                        ws[f"C{j}"].alignment = center

                    # Hoja detalle
                    ws2 = wb.create_sheet(title="Detalle")
                    for r in dataframe_to_rows(df_det, index=False, header=True):
                        ws2.append(r)

                    # Ajustes de ancho
                    for column_cells in ws.columns:
                        try:
                            length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells)
                            col_letter = column_cells[0].column_letter
                            ws.column_dimensions[col_letter].width = min(40, length + 4)
                        except:
                            pass

                    wb.save(buffer)
                    buffer.seek(0)
                    return buffer

                xlsx_data = build_report_xlsx()
                st.download_button(
                    "⬇️ Descargar Reporte XLSX",
                    data=xlsx_data,
                    file_name=f"Reporte_{id_sel}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                st.divider()
                if st.button("✅ Cerrar inventario (marcar Cerrado)"):
                    cerrar_inventario(id_sel, usuario_actual)
                    st.success("Inventario cerrado en Historial_Inventarios.")
                    st.rerun()



