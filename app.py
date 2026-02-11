import streamlit as st
import pandas as pd
import gspread
import datetime
import io
import bcrypt
from usuarios_config import USUARIOS_CREDENCIALES, CREDENCIALES_INICIALES

# Version: 3.0 - Complete rewrite with gspread client
# Using gspread directly instead of st-gsheets-connection wrapper

# ----------------------------
# CONFIG
# ----------------------------
st.set_page_config(page_title="Inventarios Rotativos - Grupo Cenoa", layout="wide", page_icon="📦")

# Construir cliente gspread directamente desde st.secrets
try:
    gs_creds = st.secrets.get("connections", {}).get("gsheets")
    if not gs_creds:
        raise KeyError("connections.gsheets not found in secrets")
    client = gspread.service_account_from_dict(gs_creds)
except Exception as e:
    st.error(f"Google Sheets credentials missing or invalid: {e}")
    st.stop()

# Spreadsheet ID (fijo)
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
# GSPREAD FUNCTIONS
# ----------------------------
@st.cache_resource
def get_spreadsheet():
    """Get spreadsheet by ID"""
    return client.open_by_key(SPREADSHEET_ID)

@st.cache_data(ttl=30)
def read_gspread_worksheet(ws_name: str) -> pd.DataFrame:
    """Read worksheet using gspread with short caching to avoid hitting API quotas.

    Cached for 30s; writers will clear the cache after updates.
    """
    try:
        spreadsheet = get_spreadsheet()
        worksheet = spreadsheet.worksheet(ws_name)
        data = worksheet.get_all_records()
        return pd.DataFrame(data) if data else pd.DataFrame()
    except Exception as e:
        # Detect quota errors and show a friendly message
        msg = str(e)
        if "RATE_LIMIT_EXCEEDED" in msg or "quota" in msg.lower() or "RESOURCE_EXHAUSTED" in msg:
            st.error(f"Error reading {ws_name}: cuota de Google Sheets excedida. Esperá unos segundos y volvé a intentar.")
        else:
            st.error(f"Error reading {ws_name}: {msg}")
        return pd.DataFrame()

def write_gspread_worksheet(ws_name: str, df: pd.DataFrame):
    """Write worksheet using gspread"""
    try:
        spreadsheet = get_spreadsheet()
        worksheet = spreadsheet.worksheet(ws_name)
        worksheet.clear()
        worksheet.update([df.columns.values.tolist()] + df.values.tolist())
        # Invalidate read cache so subsequent reads fetch fresh data
        try:
            st.cache_data.clear()
        except Exception:
            pass
    except Exception as e:
        msg = str(e)
        if "RATE_LIMIT_EXCEEDED" in msg or "quota" in msg.lower() or "RESOURCE_EXHAUSTED" in msg:
            st.error(f"Error writing {ws_name}: cuota de Google Sheets excedida. Intentá de nuevo más tarde.")
        else:
            st.error(f"Error writing {ws_name}: {msg}")

def append_gspread_worksheet(ws_name: str, df_new: pd.DataFrame):
    """Append to worksheet"""
    df_exist = read_gspread_worksheet(ws_name)
    if df_exist.empty:
        write_gspread_worksheet(ws_name, df_new)
        return
    
    df_new = df_new.copy()
    for col in df_exist.columns:
        if col not in df_new.columns:
            df_new[col] = ""
    for col in df_new.columns:
        if col not in df_exist.columns:
            df_exist[col] = ""
    
    df_final = pd.concat([df_exist, df_new[df_exist.columns]], ignore_index=True)
    write_gspread_worksheet(ws_name, df_final)
    # Ensure cache invalidation after append
    try:
        st.cache_data.clear()
    except Exception:
        pass

# ----------------------------
# AUTH
# ----------------------------
def verify_password(password: str, password_hash: str) -> bool:
    """Verify password against bcrypt hash"""
    return bcrypt.checkpw(password.encode(), password_hash.encode())

def login():
    """Login form"""
    col1, col2 = st.columns([1, 4])
    with col1:
        try:
                st.image("assets/logo_grupo_cenoa.png", width=100)
        except:
            st.write("🏢 **GRUPO CENOA**")
    
    with col2:
        st.write("")
    
    st.title("🔐 Inventarios Rotativos - Grupo Cenoa")
    
    with st.form("login_form"):
        usuario = st.text_input("Usuario (ID):", placeholder="Ej: diego_guantay")
        contrasena = st.text_input("Contraseña:", type="password")
        submit = st.form_submit_button("Ingresar", use_container_width=True)
        
        if submit:
            if usuario in USUARIOS_CREDENCIALES:
                creds = USUARIOS_CREDENCIALES[usuario]
                if verify_password(contrasena, creds["password_hash"]):
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
    
    st.divider()
    st.write("**📋 CREDENCIALES DE PRUEBA** (Eliminar después)")
    
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
    st.info("⚠️ Test credentials only.")

if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

if not st.session_state["logged_in"]:
    login()
    st.stop()

usuario_actual = st.session_state.get("usuario")
nombre_actual = st.session_state.get("nombre_usuario")
rol_actual = st.session_state.get("rol")

# --- Admin debug: mostrar estado de las hojas (solo para admin)
def _admin_debug_show():
    try:
        dfh = read_gspread_worksheet(SHEET_HIST)
        dfd = read_gspread_worksheet(SHEET_DET)
    except Exception as e:
        st.sidebar.error(f"Debug read error: {e}")
        return

    with st.sidebar.expander("GSheets debug (admin)", expanded=False):
        st.write("**Historial_Inventarios**")
        st.write("Rows:", 0 if dfh is None else len(dfh))
        st.write("Columns:", list(dfh.columns) if not (dfh is None or dfh.empty) else [])
        st.write("---")
        st.write("**Detalle_Articulos**")
        st.write("Rows:", 0 if dfd is None else len(dfd))
        st.write("Columns:", list(dfd.columns) if not (dfd is None or dfd.empty) else [])
        st.write("---")
        st.info("Este panel muestra solo conteos y nombres de columnas para depuración.")

if usuario_actual == "admin":
    _admin_debug_show()

# ----------------------------
# DATA FUNCTIONS
# ----------------------------
def listar_inventarios_abiertos():
    df_hist = read_gspread_worksheet(SHEET_HIST)
    if df_hist.empty or "Estado" not in df_hist.columns:
        return pd.DataFrame()
    return df_hist[df_hist["Estado"].astype(str).str.lower() == "abierto"].copy()

def cargar_detalle(id_inv: str) -> pd.DataFrame:
    df = read_gspread_worksheet(SHEET_DET)
    if df.empty or "ID_Inventario" not in df.columns:
        return pd.DataFrame()
    return df[df["ID_Inventario"].astype(str) == str(id_inv)].copy()

def calcular_resultados_inventario(df_det: pd.DataFrame) -> dict:
    """Calculate inventory results"""
    if df_det.empty:
        return {}
    
    df_r = df_det.copy()
    stock_col = C_STOCK if C_STOCK in df_r.columns else None
    costo_col = C_COSTO if C_COSTO in df_r.columns else None
    dif_col = "Diferencia"
    
    if not stock_col or not costo_col or dif_col not in df_r.columns:
        return {}
    
    df_r["_stock"] = pd.to_numeric(df_r[stock_col], errors="coerce").fillna(0)
    df_r["_costo"] = pd.to_numeric(df_r[costo_col], errors="coerce").fillna(0)
    df_r["_dif"] = pd.to_numeric(df_r[dif_col], errors="coerce").fillna(0)
    
    cant_muestra = int(df_r["_stock"].sum())
    valor_muestra = (df_r["_stock"] * df_r["_costo"]).sum()
    
    mask_falt = df_r["_dif"] < 0
    cant_faltantes = int((df_r.loc[mask_falt, "_dif"].abs()).sum())
    valor_faltantes = (df_r.loc[mask_falt, "_dif"].abs() * df_r.loc[mask_falt, "_costo"]).sum()
    
    mask_sobr = df_r["_dif"] > 0
    cant_sobrantes = int(df_r.loc[mask_sobr, "_dif"].sum())
    valor_sobrantes = (df_r.loc[mask_sobr, "_dif"] * df_r.loc[mask_sobr, "_costo"]).sum()
    
    cant_dif_neta = int(df_r["_dif"].sum())
    valor_dif_neta = (df_r["_dif"] * df_r["_costo"]).sum()
    
    cant_dif_absoluta = int(df_r["_dif"].abs().sum())
    valor_dif_absoluta = (df_r["_dif"].abs() * df_r["_costo"]).sum()
    
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
    """Update inventory details"""
    df_all = read_gspread_worksheet(SHEET_DET)
    if df_all.empty:
        write_gspread_worksheet(SHEET_DET, df_mod)
        return
    
    df_all = df_all.copy()
    mask = df_all["ID_Inventario"].astype(str) == str(id_inv)
    df_rest = df_all.loc[~mask].copy()
    df_final = pd.concat([df_rest, df_mod], ignore_index=True)
    write_gspread_worksheet(SHEET_DET, df_final)

def cerrar_inventario(id_inv: str, usuario: str):
    """Close inventory"""
    df_hist = read_gspread_worksheet(SHEET_HIST)
    if df_hist.empty or "ID_Inventario" not in df_hist.columns:
        return
    df_hist = df_hist.copy()
    mask = df_hist["ID_Inventario"].astype(str) == str(id_inv)
    if mask.sum() == 0:
        return
    df_hist.loc[mask, "Estado"] = "Cerrado"
    df_hist.loc[mask, "Cierre_Fecha"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    df_hist.loc[mask, "Cierre_Usuario"] = usuario
    write_gspread_worksheet(SHEET_HIST, df_hist)

# ----------------------------
# UI
# ----------------------------
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
# TAB 1
# ----------------------------
with tab1:
    st.subheader("Panel de control del Auditor")

    c1, c2 = st.columns(2)
    with c1:
        concesionaria = st.selectbox("Concesionaria", list(CONCESIONARIAS.keys()))
    with c2:
        sucursal = st.selectbox("Sucursal", CONCESIONARIAS[concesionaria])

    st.divider()

    if rol_actual != "Auditor":
        st.info("Solo Auditores pueden generar inventarios.")
    else:
        st.subheader("Importar Excel → ABC → Muestra 80/15/5")

        archivo = st.file_uploader("Subir reporte de stock (.xlsx)", type=["xlsx"])

        if archivo:
            df_base = pd.read_excel(archivo)
            st.write("Vista previa:")
            st.dataframe(df_base.head(15), use_container_width=True)

            if st.button("✅ Generar y guardar inventario"):
                falt = [c for c in [C_ART, C_LOC, C_DESC, C_STOCK, C_COSTO] if c not in df_base.columns]
                if falt:
                    st.error(f"Faltan columnas: {', '.join(falt)}")
                    st.stop()

                df = df_base.copy()
                df[C_STOCK] = pd.to_numeric(df[C_STOCK], errors="coerce").fillna(0)
                df[C_COSTO] = pd.to_numeric(df[C_COSTO], errors="coerce").fillna(0)

                df["Valor_T"] = df[C_STOCK] * df[C_COSTO]
                total = df["Valor_T"].sum()
                if total <= 0:
                    st.error("No se puede calcular ABC")
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

                df_hist = read_gspread_worksheet(SHEET_HIST)
                if not df_hist.empty and "ID_Inventario" in df_hist.columns:
                    if (df_hist["ID_Inventario"].astype(str) == id_inv).any():
                        st.warning("ID ya existe")
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
                append_gspread_worksheet(SHEET_HIST, nueva_fila)

                muestra["ID_Inventario"] = id_inv
                append_gspread_worksheet(SHEET_DET, muestra)

                st.success(f"✅ Inventario {id_inv} creado.")
                st.rerun()

# ----------------------------
# TAB 2
# ----------------------------
with tab2:
    st.subheader("Carga de conteo físico")

    if rol_actual != "Auditor":
        st.info("Solo Auditores")
    else:
        df_abiertos = listar_inventarios_abiertos()
        if df_abiertos.empty:
            st.info("No hay inventarios abiertos")
        else:
            id_sel = st.selectbox("Seleccionar inventario", df_abiertos["ID_Inventario"].astype(str).tolist())
            df_det = cargar_detalle(id_sel)
            if df_det.empty:
                st.warning("No hay detalle")
            else:
                cols_show = ["Concesionaria","Sucursal",C_LOC,C_ART,C_DESC,C_STOCK,C_COSTO,"Cat","Conteo_Fisico","Diferencia"]
                cols_show = [c for c in cols_show if c in df_det.columns]
                df_edit = df_det[cols_show].copy()

                edited = st.data_editor(
                    df_edit,
                    use_container_width=True,
                    num_rows="fixed",
                    disabled=[c for c in df_edit.columns if c != "Conteo_Fisico"],
                )

                if st.button("💾 Guardar conteo"):
                    df_det2 = df_det.copy()
                    
                    key_cols = [C_ART, C_LOC]
                    if not all(c in df_det2.columns for c in key_cols):
                        st.error("Columnas no encontradas")
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
                    st.success("✅ Conteo guardado")
                    st.rerun()

# ----------------------------
# TAB 3
# ----------------------------
with tab3:
    st.subheader("Justificaciones")

    df_abiertos = listar_inventarios_abiertos()
    if df_abiertos.empty:
        st.info("No hay inventarios abiertos")
    else:
        id_sel = st.selectbox("Seleccionar", df_abiertos["ID_Inventario"].astype(str).tolist(), key="tab3")
        df_det = cargar_detalle(id_sel)

        if df_det.empty:
            st.warning("No hay detalle")
        else:
            dif_num = pd.to_numeric(df_det.get("Diferencia", 0), errors="coerce").fillna(0)
            df_dif = df_det.loc[dif_num != 0].copy()

            if df_dif.empty:
                st.success("Sin diferencias")
            else:
                if rol_actual == "Deposito":
                    st.write("**Ingresá justificaciones:**")
                    justificaciones_dict = {}
                    
                    for idx, row in df_dif.iterrows():
                        art = row[C_ART]
                        loc = row[C_LOC]
                        dif = row["Diferencia"]
                        just_actual = row.get("Justificacion", "")
                        
                        st.write(f"**{art} ({loc}) - Diferencia: {dif}**")
                        just = st.text_area(
                            f"Justificación",
                            value=just_actual,
                            height=80,
                            key=f"just_{idx}"
                        )
                        justificaciones_dict[idx] = just
                        st.divider()
                    
                    if st.button("💾 Guardar justificaciones"):
                        df_det2 = df_det.copy()
                        for idx, just in justificaciones_dict.items():
                            df_det2.loc[df_det2.index == idx, "Justificacion"] = just
                        
                        guardar_detalle_modificado(id_sel, df_det2)
                        st.success("✅ Guardado")
                        st.rerun()
                else:
                    st.write("**Validá justificaciones:**")
                    validaciones_dict = {}
                    
                    for idx, row in df_dif.iterrows():
                        art = row[C_ART]
                        loc = row[C_LOC]
                        just = row.get("Justificacion", "")
                        val_actual = row.get("Justif_Validada", "")
                        
                        st.write(f"**{art} ({loc})**")
                        st.write(f"*{just if just else '(sin justificación)'}*")
                        
                        val = st.selectbox(
                            "¿Validada?",
                            options=["", "SI", "NO"],
                            index=(["", "SI", "NO"].index(val_actual) if val_actual in ["SI", "NO"] else 0),
                            key=f"val_{idx}"
                        )
                        validaciones_dict[idx] = val
                        st.divider()
                    
                    if st.button("💾 Guardar validación"):
                        df_det2 = df_det.copy()
                        for idx, val in validaciones_dict.items():
                            df_det2.loc[df_det2.index == idx, "Justif_Validada"] = val
                            df_det2.loc[df_det2.index == idx, "Validador"] = usuario_actual
                            df_det2.loc[df_det2.index == idx, "Fecha_Validacion"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
                        
                        guardar_detalle_modificado(id_sel, df_det2)
                        st.success("✅ Guardado")
                        st.rerun()

# ----------------------------
# TAB 4
# ----------------------------
with tab4:
    st.subheader("Cierre + Reporte")

    if rol_actual != "Auditor":
        st.info("Solo Auditores")
    else:
        df_abiertos = listar_inventarios_abiertos()
        if df_abiertos.empty:
            st.info("No hay inventarios abiertos")
        else:
            id_sel = st.selectbox("Seleccionar para cerrar", df_abiertos["ID_Inventario"].astype(str).tolist(), key="tab4")
            df_det = cargar_detalle(id_sel)

            if df_det.empty:
                st.warning("No hay detalle")
            else:
                resultados = calcular_resultados_inventario(df_det)
                
                if not resultados:
                    st.error("Error en cálculos")
                    st.stop()
                
                st.write("### 📊 Resultados")
                
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("Muestra", resultados["cant_muestra"])
                col2.metric("Faltantes", resultados["cant_faltantes"])
                col3.metric("Sobrantes", resultados["cant_sobrantes"])
                col4.metric("Grado", f"{resultados['grado']}%")
                
                tabla_resultados = pd.DataFrame([
                    {"Detalle": "Muestra", "Cant": resultados["cant_muestra"], "$": resultados["valor_muestra"]},
                    {"Detalle": "Faltantes", "Cant": resultados["cant_faltantes"], "$": resultados["valor_faltantes"]},
                    {"Detalle": "Sobrantes", "Cant": resultados["cant_sobrantes"], "$": resultados["valor_sobrantes"]},
                    {"Detalle": "Dif Neta", "Cant": resultados["cant_dif_neta"], "$": resultados["valor_dif_neta"]},
                    {"Detalle": "Dif Absoluta", "Cant": resultados["cant_dif_absoluta"], "$": resultados["valor_dif_absoluta"]},
                ])
                
                st.dataframe(tabla_resultados, use_container_width=True, hide_index=True)
                
                st.divider()
                st.write("### 📥 Descargar reporte:")

                def build_report_xlsx():
                    from openpyxl import Workbook
                    from openpyxl.utils.dataframe import dataframe_to_rows
                    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

                    buffer = io.BytesIO()
                    wb = Workbook()
                    ws = wb.active
                    ws.title = "Resultado"

                    title_font = Font(size=14, bold=True)
                    light_red = PatternFill(start_color="FFF2F2", end_color="FFF2F2", fill_type="solid")
                    bold = Font(bold=True)
                    center = Alignment(horizontal="center", vertical="center")
                    thin = Side(border_style="thin", color="000000")
                    border = Border(left=thin, right=thin, top=thin, bottom=thin)

                    ws.merge_cells("A1:D1")
                    ws["A1"] = "4. Resultado Inventario Rotativo"
                    ws["A1"].font = title_font
                    ws["A1"].alignment = center

                    ws["A3"] = "Resultado:"

                    start_row = 5
                    ws[f"A{start_row}"] = "Detalle"
                    ws[f"B{start_row}"] = "Cant"
                    ws[f"C{start_row}"] = "$"
                    
                    rows = [
                        ("Muestra", resultados["cant_muestra"], resultados["valor_muestra"]),
                        ("Faltantes", resultados["cant_faltantes"], resultados["valor_faltantes"]),
                        ("Sobrantes", resultados["cant_sobrantes"], resultados["valor_sobrantes"]),
                        ("Dif Neta", resultados["cant_dif_neta"], resultados["valor_dif_neta"]),
                        ("Dif Absoluta", resultados["cant_dif_absoluta"], resultados["valor_dif_absoluta"]),
                    ]

                    for i, r in enumerate(rows, start=start_row + 1):
                        ws[f"A{i}"] = r[0]
                        ws[f"B{i}"] = r[1]
                        ws[f"C{i}"] = r[2]
                        ws[f"C{i}"].number_format = "#,##0.00"

                    ws2 = wb.create_sheet(title="Detalle")
                    for r in dataframe_to_rows(df_det, index=False, header=True):
                        ws2.append(r)

                    wb.save(buffer)
                    buffer.seek(0)
                    return buffer

                xlsx_data = build_report_xlsx()
                st.download_button(
                    "⬇️ Descargar XLSX",
                    data=xlsx_data,
                    file_name=f"Reporte_{id_sel}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                st.divider()
                if st.button("✅ Cerrar inventario"):
                    cerrar_inventario(id_sel, usuario_actual)
                    st.success("Cerrado")
                    st.rerun()
