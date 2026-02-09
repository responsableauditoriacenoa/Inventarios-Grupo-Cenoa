import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
import datetime
import io

# ----------------------------
# CONFIG
# ----------------------------
st.set_page_config(page_title="Inventarios Rotativos - Grupo Cenoa", layout="wide", page_icon="📦")
conn = st.connection("gsheets", type=GSheetsConnection)

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
    df = conn.read(worksheet=ws)
    if df is None:
        return pd.DataFrame()
    return df

def _update_ws(ws: str, df: pd.DataFrame):
    conn.update(worksheet=ws, data=df)

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
# AUTH SIMPLE
# ----------------------------
USUARIOS = ["Seleccionar", "Diego Guantay", "Nancy Fernandez", "Gustavo Zambrano", "Admin", "Jefe de Repuestos"]

def login():
    with st.sidebar:
        st.title("🔐 Acceso")
        user = st.selectbox("Usuario", USUARIOS)
        if user == "Seleccionar":
            return None, None
        rol = "Deposito" if user == "Jefe de Repuestos" else "Auditor"
        return user, rol

usuario_actual, rol_actual = login()
if not usuario_actual:
    st.info("Seleccioná tu usuario para comenzar.")
    st.stop()

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
            # Tu archivo viene con headers en la primera fila, así que lo leemos directo
            df_base = pd.read_excel(archivo)

            st.write("Vista previa del reporte:")
            st.dataframe(df_base.head(15), use_container_width=True)

            if st.button("✅ Generar y guardar inventario"):
                # Validación de columnas mínimas
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

                # Muestra 80 / 15 / 5
                df_a = df[df["Cat"] == "A"]
                df_b = df[df["Cat"] == "B"]
                df_c = df[df["Cat"] == "C"]

                m_a = df_a.sample(n=min(80, len(df_a))) if len(df_a) else df_a
                m_b = df_b.sample(n=min(15, len(df_b))) if len(df_b) else df_b
                m_c = df_c.sample(n=min(5, len(df_c))) if len(df_c) else df_c

                muestra = pd.concat([m_a, m_b, m_c], ignore_index=True)

                # Campos del circuito
                muestra["Concesionaria"] = concesionaria
                muestra["Sucursal"] = sucursal
                muestra["Conteo_Fisico"] = ""
                muestra["Diferencia"] = ""
                muestra["Justificacion"] = ""
                muestra["Justif_Validada"] = ""
                muestra["Validador"] = ""
                muestra["Fecha_Validacion"] = ""

                id_inv = datetime.datetime.now().strftime("INV-%Y%m%d-%H%M")

                # Guardar historial (anti-duplicado por rerun)
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

                # Guardar detalle
                muestra["ID_Inventario"] = id_inv
                _append_df(SHEET_DET, muestra)

                st.success(f"✅ Inventario {id_inv} creado y guardado.")
                st.session_state["id_inv"] = id_inv
                st.dataframe(muestra[[C_LOC, C_ART, C_DESC, "Cat"]], use_container_width=True)

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

# ----------------------------
# TAB 3: JUSTIFICACIONES (DEPÓSITO + VALIDACIÓN AUDITOR)
# ----------------------------
with tab3:
    st.subheader("Justificaciones y validación")

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
                st.success("No hay diferencias para justificar.")
            else:
                base_cols = ["Concesionaria","Sucursal",C_LOC,C_ART,C_DESC,C_STOCK,"Conteo_Fisico","Diferencia","Justificacion","Justif_Validada"]
                base_cols = [c for c in base_cols if c in df_dif.columns]
                df_view = df_dif[base_cols].copy()

                if rol_actual == "Deposito":
                    st.caption("Depósito: completá Justificacion y guardá.")
                    edited = st.data_editor(
                        df_view,
                        use_container_width=True,
                        num_rows="fixed",
                        disabled=[c for c in df_view.columns if c != "Justificacion"],
                        key=f"dep_{id_sel}"
                    )
                    if st.button("💾 Guardar justificaciones (Depósito)"):
                        df_det2 = df_det.copy()

                        key_cols = [C_ART, C_LOC]
                        edited2 = edited.copy()
                        for c in key_cols:
                            edited2[c] = edited2[c].astype(str)
                            df_det2[c] = df_det2[c].astype(str)

                        df_merge = df_det2.merge(
                            edited2[key_cols + ["Justificacion"]],
                            on=key_cols,
                            how="left",
                            suffixes=("", "_new")
                        )
                        df_merge["Justificacion"] = df_merge["Justificacion_new"].combine_first(df_merge.get("Justificacion"))
                        if "Justificacion_new" in df_merge.columns:
                            df_merge = df_merge.drop(columns=["Justificacion_new"])

                        guardar_detalle_modificado(id_sel, df_merge)
                        st.success("✅ Justificaciones guardadas.")
                else:
                    st.caption("Auditor: marcá Justif_Validada (SI/NO) y guardá.")
                    edited = st.data_editor(
                        df_view,
                        use_container_width=True,
                        num_rows="fixed",
                        disabled=[c for c in df_view.columns if c != "Justif_Validada"],
                        key=f"audit_{id_sel}"
                    )
                    if st.button("💾 Guardar validación (Auditor)"):
                        df_det2 = df_det.copy()

                        key_cols = [C_ART, C_LOC]
                        edited2 = edited.copy()
                        for c in key_cols:
                            edited2[c] = edited2[c].astype(str)
                            df_det2[c] = df_det2[c].astype(str)

                        df_merge = df_det2.merge(
                            edited2[key_cols + ["Justif_Validada"]],
                            on=key_cols,
                            how="left",
                            suffixes=("", "_new")
                        )
                        df_merge["Justif_Validada"] = df_merge["Justif_Validada_new"].combine_first(df_merge.get("Justif_Validada"))
                        df_merge["Validador"] = usuario_actual
                        df_merge["Fecha_Validacion"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

                        if "Justif_Validada_new" in df_merge.columns:
                            df_merge = df_merge.drop(columns=["Justif_Validada_new"])

                        guardar_detalle_modificado(id_sel, df_merge)
                        st.success("✅ Validación guardada.")

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
                dif = pd.to_numeric(df_det.get("Diferencia", 0), errors="coerce").fillna(0)

                faltantes = dif[dif < 0].sum()
                sobrantes = dif[dif > 0].sum()
                neta = dif.sum()
                absoluta = dif.abs().sum()

                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Faltantes (sum)", f"{faltantes:,.2f}")
                c2.metric("Sobrantes (sum)", f"{sobrantes:,.2f}")
                c3.metric("Diferencia neta", f"{neta:,.2f}")
                c4.metric("Diferencia absoluta", f"{absoluta:,.2f}")

                difmask = dif != 0
                val = df_det.get("Justif_Validada", "").astype(str)
                ok_validacion = bool(((~difmask) | (val.str.strip() != "")).all())

                if not ok_validacion:
                    st.warning("⚠️ Hay diferencias sin validar (Justif_Validada vacío).")

                st.divider()
                st.write("Descargar reporte (.xlsx):")

                def build_report_xlsx():
                    buffer = io.BytesIO()

                    # Datos del encabezado
                    conces = df_det["Concesionaria"].iloc[0] if "Concesionaria" in df_det.columns and len(df_det) else ""
                    sucu = df_det["Sucursal"].iloc[0] if "Sucursal" in df_det.columns and len(df_det) else ""

                    resumen = pd.DataFrame([{
                        "ID_Inventario": id_sel,
                        "Fecha_Reporte": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
                        "Concesionaria": conces,
                        "Sucursal": sucu,
                        "Faltantes_Sum": faltantes,
                        "Sobrantes_Sum": sobrantes,
                        "Diferencia_Neta": neta,
                        "Diferencia_Absoluta": absoluta,
                    }])

                    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                        resumen.to_excel(writer, index=False, sheet_name="Resumen")
                        df_det.to_excel(writer, index=False, sheet_name="Detalle")

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


