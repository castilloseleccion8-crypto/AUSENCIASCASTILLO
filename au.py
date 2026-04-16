import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import gspread
from google.oauth2.service_account import Credentials
import io

st.set_page_config(page_title="Gestión de Ausencias", layout="wide", page_icon="📋")

st.markdown("""
<style>
html, body, [class*="css"] { font-family: 'Segoe UI', sans-serif; }
.stApp { background-color: #f0f2f5; }
section[data-testid="stSidebar"] { background-color: #0f1f3d !important; }
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] div.stRadio > label,
section[data-testid="stSidebar"] div.stSelectbox label,
section[data-testid="stSidebar"] div.stTextInput label,
section[data-testid="stSidebar"] .stMarkdown { color: white !important; }
section[data-testid="stSidebar"] input {
    background-color: #1e3a6e !important;
    color: white !important;
    border-color: #3b5ea6 !important;
}
section[data-testid="stSidebar"] div[data-baseweb="select"] > div {
    background-color: #1e3a6e !important;
    color: white !important;
}
div.stButton > button {
    background-color: #0f1f3d; color: white; border-radius: 8px;
    padding: 8px 20px; border: none; font-weight: 600; width: 100%;
}
div.stButton > button:hover { background-color: #1e3a6e; }
.metric-box {
    background: white; border-radius: 10px; padding: 14px;
    text-align: center; box-shadow: 0 2px 8px rgba(0,0,0,0.06);
}
.metric-num   { font-size: 30px; font-weight: 800; }
.metric-label { font-size: 12px; color: #666; margin-top: 3px; }
</style>
""", unsafe_allow_html=True)

# ── CONSTANTES ──────────────────────────────────────────────────────────────
TIPIFICACIONES = [
    "Tardanza / Cambio de horario",
    "Vacaciones",
    "Compensatorio",
    "Licencia",
    "Ausencia injustificada",
]
NAALOO_OPTS = [
    "Sí, está cargado en Naaloo",
    "No, falta cargarlo en Naaloo",
    "No aplica",
]
PASSWORD_RRHH  = "rrhh2026"
HOJA_REPORTE   = "REPORTE"
HOJA_GESTIONES = "GESTIONES"
HOJA_HISTORIAL = "HISTORIAL"
COLS_REPORTE   = ["FECHA","LEGAJO","SUCURSAL","NOMBRE","POSICION","MAÑANA","TARDE","SITUACION","LICENCIA"]
COLS_GESTIONES = ["KEY","TIPIFICACION","NAALOO","OBSERVACIONES","RESPONSABLE","FECHA_GESTION","SUCURSAL_ORIGEN"]

# ── GOOGLE SHEETS ────────────────────────────────────────────────────────────
@st.cache_resource
def get_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets",
              "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
    return gspread.authorize(creds)

def get_sheet(nombre):
    gc = get_client()
    sh = gc.open_by_key(st.secrets["SHEET_ID"])
    try:
        return sh.worksheet(nombre)
    except gspread.WorksheetNotFound:
        return sh.add_worksheet(title=nombre, rows=3000, cols=20)

@st.cache_data(ttl=60)
def cargar_reporte():
    ws = get_sheet(HOJA_REPORTE)
    data = ws.get_all_records()
    if not data:
        return pd.DataFrame(columns=COLS_REPORTE)
    df = pd.DataFrame(data).fillna("")
    for col in COLS_REPORTE:
        if col not in df.columns:
            df[col] = ""
    return df

@st.cache_data(ttl=30)
def cargar_gestiones():
    ws = get_sheet(HOJA_GESTIONES)
    data = ws.get_all_records()
    if not data:
        return {}
    return {
        row["KEY"]: {k.lower(): v for k, v in row.items() if k != "KEY"}
        for row in data if row.get("KEY")
    }

def subir_reporte(df):
    ws = get_sheet(HOJA_REPORTE)
    ws.clear()
    for col in COLS_REPORTE:
        if col not in df.columns:
            df[col] = ""
    df = df[COLS_REPORTE].fillna("").astype(str)
    ws.update([df.columns.tolist()] + df.values.tolist())

def guardar_gestion(key, datos):
    ws = get_sheet(HOJA_GESTIONES)
    todas = ws.get_all_records()
    keys_ex = [r.get("KEY") for r in todas]
    fila = [
        key,
        datos.get("tipificacion",""),
        datos.get("naaloo",""),
        datos.get("observaciones",""),
        datos.get("responsable",""),
        datos.get("fecha_gestion",""),
        datos.get("sucursal_origen",""),
    ]
    if key in keys_ex:
        idx = keys_ex.index(key) + 2
        ws.update(f"A{idx}:G{idx}", [fila])
    else:
        if not todas:
            ws.update("A1:G1", [COLS_GESTIONES])
        ws.append_row(fila)

def archivar_y_eliminar(keys_a_archivar, indices_reporte, df_completo, gestiones_dict):
    ws_hist = get_sheet(HOJA_HISTORIAL)
    ws_gest = get_sheet(HOJA_GESTIONES)

    hist_data = ws_hist.get_all_records()
    if not hist_data:
        ws_hist.update("A1:O1", [["KEY","FECHA_ARCHIVO","FECHA","LEGAJO","SUCURSAL","NOMBRE","POSICION",
                                   "MAÑANA","TARDE","LICENCIA_NAALOO",
                                   "TIPIFICACION","NAALOO_CARGADO","OBSERVACIONES",
                                   "RESPONSABLE","FECHA_GESTION"]])

    filas_hist = []
    for k in keys_a_archivar:
        g = gestiones_dict.get(k, {})
        fila_rep = df_completo[df_completo.apply(lambda r: clave(r), axis=1) == k]
        nombre_r = fila_rep["NOMBRE"].values[0]   if not fila_rep.empty else ""
        pos_r    = fila_rep["POSICION"].values[0]  if not fila_rep.empty else ""
        manana_r = fila_rep["MAÑANA"].values[0]   if not fila_rep.empty else ""
        tarde_r  = fila_rep["TARDE"].values[0]    if not fila_rep.empty else ""
        lic_r    = fila_rep["LICENCIA"].values[0]  if not fila_rep.empty else ""
        fecha_r  = fila_rep["FECHA"].values[0]    if not fila_rep.empty else ""
        legajo_r = fila_rep["LEGAJO"].values[0]   if not fila_rep.empty else ""
        suc_r    = fila_rep["SUCURSAL"].values[0] if not fila_rep.empty else ""
        filas_hist.append([
            k, datetime.now().strftime("%d/%m/%Y %H:%M"),
            fecha_r, legajo_r, suc_r, nombre_r, pos_r, manana_r, tarde_r, lic_r,
            g.get("tipificacion",""), g.get("naaloo",""),
            g.get("observaciones",""), g.get("responsable",""), g.get("fecha_gestion",""),
        ])

    if filas_hist:
        ws_hist.append_rows(filas_hist)

    df_nuevo = df_completo.drop(index=indices_reporte).reset_index(drop=True)
    subir_reporte(df_nuevo)

    todas_g = ws_gest.get_all_records()
    keys_g  = [r.get("KEY") for r in todas_g]
    filas_borrar = sorted(
        [keys_g.index(k) + 2 for k in keys_a_archivar if k in keys_g],
        reverse=True
    )
    for fila_idx in filas_borrar:
        ws_gest.delete_rows(fila_idx)

    return df_nuevo

# ── HELPERS ──────────────────────────────────────────────────────────────────
def fmt_fichada(v):
    v = str(v).strip().upper()
    if v == "A": return "AUSENTE"
    if v == "P": return "PRESENTE"
    if v == "-": return "—"
    return v

def es_falso_ausente(row):
    """Sábado a la tarde con mañana presente = no es ausencia real."""
    try:
        fecha = datetime.strptime(str(row["FECHA"]).strip(), "%d/%m/%Y").date()
        if fecha.weekday() != 5:  # 5 = sábado
            return False
        manana = str(row["MAÑANA"]).strip().upper()
        tarde  = str(row["TARDE"]).strip().upper()
        return manana == "P" and tarde == "A"
    except Exception:
        return False

def calcular_plazo(fecha_str):
    try:
        fecha = datetime.strptime(str(fecha_str).strip(), "%d/%m/%Y").date()
    except Exception:
        return "PENDIENTE", 99
    limite = fecha + timedelta(days=2)
    dias = (limite - date.today()).days
    return ("VENCIDO" if date.today() > limite else "PENDIENTE"), dias

def clave(row):
    return f"{str(row['FECHA']).strip()}_{str(row['LEGAJO']).strip()}_{str(row['SUCURSAL']).strip()}"

def estado_gestion_fn(key, gestiones):
    g = gestiones.get(key, {})
    return ("RESUELTO", g) if g.get("tipificacion") else ("PENDIENTE", g)

def estado_fila(row, gestiones):
    k = clave(row)
    est, _ = estado_gestion_fn(k, gestiones)
    if est == "RESUELTO":
        return "RESUELTO"
    plazo, _ = calcular_plazo(row["FECHA"])
    return plazo

def excel_a_df(archivo_bytes):
    sheets = pd.read_excel(io.BytesIO(archivo_bytes), sheet_name=None)
    frames = []
    for _, df in sheets.items():
        df.columns = [str(c).strip() for c in df.columns]
        df = df.fillna("")
        for col in COLS_REPORTE:
            if col not in df.columns:
                df[col] = ""
        frames.append(df[COLS_REPORTE])
    return pd.concat(frames, ignore_index=True).astype(str) if frames else pd.DataFrame(columns=COLS_REPORTE)

def slug_pass(sucursal):
    return (sucursal.lower()
            .replace(" ","").replace(".","")
            .replace("ü","u").replace("é","e").replace("ó","o")
            .replace("á","a").replace("í","i").replace("ú","u"))

# ── FORMULARIO TIPIFICACIÓN ───────────────────────────────────────────────────
def form_tipificacion(key, gestion_actual, gestiones_dict, sucursal_origen):
    c1, c2 = st.columns(2)
    with c1:
        tip_idx = TIPIFICACIONES.index(gestion_actual.get("tipificacion", TIPIFICACIONES[0])) \
                  if gestion_actual.get("tipificacion") in TIPIFICACIONES else 0
        tipificacion = st.selectbox("Tipo de ausencia *", TIPIFICACIONES, index=tip_idx, key=f"tip_{key}")
    with c2:
        n_idx = NAALOO_OPTS.index(gestion_actual.get("naaloo", NAALOO_OPTS[0])) \
                if gestion_actual.get("naaloo") in NAALOO_OPTS else 0
        naaloo = st.selectbox("¿Cargado en Naaloo? *", NAALOO_OPTS, index=n_idx, key=f"naaloo_{key}")

    obs  = st.text_area("Observaciones", value=gestion_actual.get("observaciones",""),
                        placeholder="Ej: Avisó por WhatsApp, presentó certificado el...",
                        key=f"obs_{key}", height=68)
    resp = st.text_input("Responsable *", value=gestion_actual.get("responsable",""),
                         placeholder="Nombre de quien tipifica", key=f"resp_{key}")

    cb, _ = st.columns([1, 3])
    with cb:
        if st.button("💾 Guardar", key=f"btn_{key}"):
            if not resp.strip():
                st.error("El campo Responsable es obligatorio.")
            else:
                datos = {
                    "tipificacion": tipificacion, "naaloo": naaloo,
                    "observaciones": obs.strip(), "responsable": resp.strip(),
                    "fecha_gestion": datetime.now().strftime("%d/%m/%Y %H:%M"),
                    "sucursal_origen": sucursal_origen,
                }
                guardar_gestion(key, datos)
                gestiones_dict[key] = datos
                st.success("✅ Guardado")
                st.cache_data.clear()
                st.rerun()

    if gestion_actual.get("fecha_gestion"):
        st.caption(f"Último guardado: {gestion_actual['fecha_gestion']} — {gestion_actual.get('responsable','')} ({gestion_actual.get('sucursal_origen','')})")

# ── SIDEBAR ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🔐 Acceso")
    modo = st.radio("Tipo de acceso", ["Sucursal", "RRHH (Global)"])

    df_temp = cargar_reporte()
    sucursales_disponibles = sorted(df_temp["SUCURSAL"].dropna().unique().tolist()) if not df_temp.empty else []

    if modo == "Sucursal":
        sucursal_sel = st.selectbox("Sucursal", sucursales_disponibles) if sucursales_disponibles else None
        password     = st.text_input("Contraseña", type="password")
        es_rrhh      = False
        autenticado  = bool(sucursal_sel and password == f"{slug_pass(sucursal_sel)}{date.today().year}")
    else:
        sucursal_sel = None
        password     = st.text_input("Contraseña RRHH", type="password")
        es_rrhh      = True
        autenticado  = password == PASSWORD_RRHH

    if password and not autenticado:
        st.error("❌ Contraseña incorrecta")
    st.markdown("---")
    st.caption(f"📅 {date.today().strftime('%d/%m/%Y')}")

# ── HEADER ───────────────────────────────────────────────────────────────────
st.markdown('<h1 style="text-align:center;color:#0f1f3d;font-size:34px;font-weight:800;margin-bottom:4px;">📋 GESTIÓN DE AUSENCIAS</h1>', unsafe_allow_html=True)
st.markdown('<p style="text-align:center;color:#666;font-size:15px;margin-bottom:20px;">Tipificación y seguimiento diario</p>', unsafe_allow_html=True)

if not password:
    st.info("👈 Seleccioná tu sucursal e ingresá la contraseña para continuar.")
    st.stop()
if not autenticado:
    st.stop()

df_global = cargar_reporte()
gestiones = cargar_gestiones()

# ── PANEL RRHH ───────────────────────────────────────────────────────────────
if es_rrhh:
    with st.expander("⚙️ Administración", expanded=df_global.empty):
        tab_subir, tab_limpiar = st.tabs(["📤 Subir nuevo reporte", "🗂️ Archivar casos"])

        with tab_subir:
            # ── Borrar reporte actual ──
            if not df_global.empty:
                with st.container():
                    st.markdown("**🗑️ Borrar reporte actual**")
                    st.caption("Elimina todos los datos del reporte. Las tipificaciones guardadas NO se borran.")
                    col_del, _ = st.columns([1, 3])
                    with col_del:
                        if st.button("🗑️ Borrar todo el reporte", key="btn_borrar_reporte"):
                            with st.spinner("Borrando..."):
                                ws_rep = get_sheet(HOJA_REPORTE)
                                ws_rep.clear()
                            st.success("✅ Reporte borrado. Ya podés subir el nuevo Excel.")
                            st.cache_data.clear()
                            st.rerun()
                st.markdown("---")

            # ── Subir nuevo Excel ──
            st.markdown("**⬆️ Subir nuevo Excel**")
            st.caption("Las tipificaciones ya cargadas NO se tocan.")
            archivo = st.file_uploader("Archivo .xlsx", type=["xlsx"], key="up_reporte")
            if archivo is not None:
                try:
                    df_nuevo = excel_a_df(archivo.getvalue())
                    aus_n = df_nuevo[df_nuevo["SITUACION"].str.contains("AUSENTE", case=False, na=False)]
                    st.success(f"✅ {len(df_nuevo)} filas — {len(aus_n)} ausentes")
                    st.dataframe(df_nuevo.head(8), use_container_width=True, hide_index=True)
                    c_btn, _ = st.columns([1, 3])
                    with c_btn:
                        if st.button("⬆️ Confirmar y subir", key="btn_subir"):
                            with st.spinner("Subiendo..."):
                                subir_reporte(df_nuevo)
                            st.success("✅ Reporte actualizado. Las gestiones previas se conservan.")
                            st.cache_data.clear()
                            st.rerun()
                except Exception as e:
                    st.error(f"Error: {e}")

        with tab_limpiar:
            if df_global.empty:
                st.info("No hay datos cargados.")
            else:
                df_adm = df_global[df_global["SITUACION"].str.contains("AUSENTE|MEDIO", case=False, na=False)].copy()
                df_adm = df_adm[~df_adm.apply(es_falso_ausente, axis=1)]
                df_adm["_ESTADO"] = df_adm.apply(lambda r: estado_fila(r, gestiones), axis=1)
                df_adm["_KEY"]    = df_adm.apply(clave, axis=1)

                tipificados_keys = [row["_KEY"] for _, row in df_adm.iterrows() if row["_ESTADO"] == "RESUELTO"]
                tipificados_idx  = df_adm[df_adm["_ESTADO"] == "RESUELTO"].index.tolist()

                # ── Botón archivar todos los tipificados ──
                c_limpiar, c_info = st.columns([2, 3])
                with c_limpiar:
                    if tipificados_keys:
                        if st.button(f"✅ Archivar todos los tipificados ({len(tipificados_keys)})", key="btn_limpiar_todos"):
                            with st.spinner("Archivando..."):
                                archivar_y_eliminar(tipificados_keys, tipificados_idx, df_global, gestiones)
                            st.success(f"✅ {len(tipificados_keys)} casos archivados en HISTORIAL.")
                            st.cache_data.clear()
                            st.rerun()
                    else:
                        st.info("No hay casos tipificados para archivar.")
                with c_info:
                    st.caption("Los casos archivados van a la hoja HISTORIAL del Drive. No se pierden.")

                st.markdown("---")
                st.markdown("**O elegí casos individuales:**")

                # Filtro por sucursal
                suc_adm = st.selectbox("Filtrar por sucursal", ["Todas"] + sucursales_disponibles, key="adm_suc")
                df_adm_v = df_adm.copy()
                if suc_adm != "Todas":
                    df_adm_v = df_adm_v[df_adm_v["SUCURSAL"] == suc_adm]

                # ── SELECCIONAR TODO ──
                sel_todo = st.checkbox(
                    f"☑️ Seleccionar todos ({len(df_adm_v)})",
                    key="chk_todo"
                )

                seleccionados = []
                for i, (orig_idx, row) in enumerate(df_adm_v.iterrows()):
                    k   = row["_KEY"]
                    est = row["_ESTADO"]
                    g   = gestiones.get(k, {})
                    if est == "RESUELTO":
                        est_txt = f"✅ {g.get('tipificacion','')}"
                    elif est == "VENCIDO":
                        est_txt = "🚨 VENCIDO"
                    else:
                        est_txt = "⏳ PENDIENTE"

                    pos_txt = f" · {row.get('POSICION','').strip()}" if str(row.get('POSICION','')).strip() else ""

                    cc, ci = st.columns([0.5, 9.5])
                    with cc:
                        # Si "seleccionar todo" está marcado, forzar True
                        elegido = st.checkbox(
                            "", key=f"chk_{k}_{i}",
                            value=sel_todo,
                            label_visibility="collapsed"
                        )
                    with ci:
                        st.markdown(
                            f"**{row['NOMBRE'].strip()}**{pos_txt} · "
                            f"Leg. {row['LEGAJO']} · {row['SUCURSAL']} · {row['FECHA']} · {est_txt}"
                        )
                    if elegido:
                        seleccionados.append((orig_idx, k))

                if seleccionados:
                    st.warning(f"{len(seleccionados)} caso(s) seleccionados → se moverán al HISTORIAL.")
                    cd, _ = st.columns([1, 4])
                    with cd:
                        if st.button("🗂️ Archivar seleccionados", key="btn_arch"):
                            with st.spinner("Archivando..."):
                                archivar_y_eliminar(
                                    [s[1] for s in seleccionados],
                                    [s[0] for s in seleccionados],
                                    df_global, gestiones
                                )
                            st.success("✅ Archivados.")
                            st.cache_data.clear()
                            st.rerun()

    st.markdown("---")

if df_global.empty:
    st.warning("No hay datos. RRHH debe subir el reporte primero.")
    st.stop()

sucursales_disponibles = sorted(df_global["SUCURSAL"].dropna().unique().tolist())
df_vista = df_global.copy() if es_rrhh else df_global[df_global["SUCURSAL"] == sucursal_sel].copy()

df_ausentes = df_vista[df_vista["SITUACION"].str.contains("AUSENTE|MEDIO", case=False, na=False)].copy()
df_ausentes = df_ausentes[~df_ausentes.apply(es_falso_ausente, axis=1)]
df_ausentes["_ESTADO"] = df_ausentes.apply(lambda r: estado_fila(r, gestiones), axis=1)
df_ausentes["_KEY"]    = df_ausentes.apply(clave, axis=1)

titulo_vista = "Vista Global — Todas las Sucursales" if es_rrhh else f"Sucursal: {sucursal_sel}"
st.markdown(f"### {titulo_vista}")

# ── MÉTRICAS ─────────────────────────────────────────────────────────────────
total      = len(df_ausentes)
resueltos  = int((df_ausentes["_ESTADO"] == "RESUELTO").sum())
vencidos   = int((df_ausentes["_ESTADO"] == "VENCIDO").sum())
pendientes = int((df_ausentes["_ESTADO"] == "PENDIENTE").sum())

c1, c2, c3, c4 = st.columns(4)
c1.markdown(f'<div class="metric-box"><div class="metric-num" style="color:#0f1f3d">{total}</div><div class="metric-label">Registros totales</div></div>', unsafe_allow_html=True)
c2.markdown(f'<div class="metric-box"><div class="metric-num" style="color:#f59e0b">{pendientes}</div><div class="metric-label">⏳ Pendientes</div></div>', unsafe_allow_html=True)
c3.markdown(f'<div class="metric-box"><div class="metric-num" style="color:#dc2626">{vencidos}</div><div class="metric-label">{"🚨 Urgentes" if not es_rrhh else "🚨 Vencidos"}</div></div>', unsafe_allow_html=True)
c4.markdown(f'<div class="metric-box"><div class="metric-num" style="color:#10b981">{resueltos}</div><div class="metric-label">✅ Tipificados</div></div>', unsafe_allow_html=True)
st.markdown("<br>", unsafe_allow_html=True)

# ── FILTROS ───────────────────────────────────────────────────────────────────
cf1, cf2, cf3 = st.columns([2, 2, 2])
with cf1:
    buscar = st.text_input("🔎 Buscar nombre o legajo")
with cf2:
    filtro_estado = st.selectbox("Estado", ["Todos", "Pendientes", "Urgentes" if not es_rrhh else "Vencidos", "Tipificados"])
with cf3:
    filtro_suc = st.selectbox("Sucursal", ["Todas"] + sucursales_disponibles) if es_rrhh else sucursal_sel

df_filtrado = df_ausentes.copy()
if buscar:
    df_filtrado = df_filtrado[
        df_filtrado["NOMBRE"].str.contains(buscar, case=False, na=False) |
        df_filtrado["LEGAJO"].astype(str).str.contains(buscar, na=False)
    ]
if es_rrhh and filtro_suc != "Todas":
    df_filtrado = df_filtrado[df_filtrado["SUCURSAL"] == filtro_suc]
if filtro_estado in ("Vencidos", "Urgentes"):
    df_filtrado = df_filtrado[df_filtrado["_ESTADO"] == "VENCIDO"]
elif filtro_estado == "Pendientes":
    df_filtrado = df_filtrado[df_filtrado["_ESTADO"] == "PENDIENTE"]
elif filtro_estado == "Tipificados":
    df_filtrado = df_filtrado[df_filtrado["_ESTADO"] == "RESUELTO"]

# ── AGRUPAR POR EMPLEADO ─────────────────────────────────────────────────────
orden_map = {"VENCIDO": 0, "PENDIENTE": 1, "RESUELTO": 2}
df_filtrado = df_filtrado.copy()
df_filtrado["_ORDEN"] = df_filtrado["_ESTADO"].map(orden_map)

peor_estado = df_filtrado.groupby("LEGAJO")["_ORDEN"].min().reset_index().rename(columns={"_ORDEN": "_PEOR"})
df_filtrado = df_filtrado.merge(peor_estado, on="LEGAJO", how="left")
df_filtrado = df_filtrado.sort_values(["_PEOR", "SUCURSAL", "NOMBRE", "FECHA"]).reset_index(drop=True)

dias_por_legajo = df_filtrado.groupby("LEGAJO").size().to_dict()

n_empleados = df_filtrado["LEGAJO"].nunique()
st.markdown(f"**{n_empleados} empleados** · {len(df_filtrado)} registros")
st.markdown("---")

# ── VISTA POR EMPLEADO ────────────────────────────────────────────────────────
if df_filtrado.empty:
    st.info("Sin registros para los filtros seleccionados.")
else:
    grupos = df_filtrado.groupby(["LEGAJO", "NOMBRE", "SUCURSAL"], sort=False)

    for (legajo, nombre, sucursal), grupo in grupos:
        dias_grupo    = len(grupo)
        estados_grupo = grupo["_ESTADO"].tolist()

        # Posición — tomamos la primera que aparezca en el grupo
        posicion = str(grupo["POSICION"].values[0]).strip() if "POSICION" in grupo.columns else ""

        if "VENCIDO" in estados_grupo:
            color_header   = "#dc2626"
            label_urgencia = "URGENTE" if not es_rrhh else "VENCIDO"
            bg_urgencia    = "#fee2e2"
            fg_urgencia    = "#991b1b"
        elif "PENDIENTE" in estados_grupo:
            color_header   = "#f59e0b"
            label_urgencia = "PENDIENTE"
            bg_urgencia    = "#fef3c7"
            fg_urgencia    = "#92400e"
        else:
            color_header   = "#10b981"
            label_urgencia = "TIPIFICADO"
            bg_urgencia    = "#d1fae5"
            fg_urgencia    = "#065f46"

        dias_txt = f"{dias_grupo} día{'s' if dias_grupo > 1 else ''} en falta"
        suc_txt  = f"  ·  📍 {sucursal}" if es_rrhh else ""
        pos_txt  = f"  ·  🏷️ {posicion}" if posicion and posicion not in ("nan","") else ""

        col_b, col_c = st.columns([0.015, 0.985])
        with col_b:
            st.markdown(
                f'<div style="background:{color_header};height:100%;min-height:60px;border-radius:3px;"></div>',
                unsafe_allow_html=True
            )
        with col_c:
            ch1, _ = st.columns([3, 1])
            with ch1:
                st.markdown(f"**👤 {nombre.strip()}** &nbsp; Leg. {legajo}{suc_txt}{pos_txt}")
                st.markdown(
                    f'<span style="background:{bg_urgencia};color:{fg_urgencia};'
                    f'padding:2px 10px;border-radius:20px;font-size:12px;font-weight:600;">'
                    f'{label_urgencia}</span>'
                    f'&nbsp;&nbsp;<span style="color:#888;font-size:13px;">📅 {dias_txt}</span>',
                    unsafe_allow_html=True
                )

        st.markdown("<div style='margin-left:2rem'>", unsafe_allow_html=True)

        for _, row in grupo.iterrows():
            k = row["_KEY"]
            est_actual, gestion_actual = estado_gestion_fn(k, gestiones)
            plazo_str, dias_restantes  = calcular_plazo(str(row["FECHA"]))

            if est_actual == "RESUELTO":
                tip_txt  = gestion_actual.get("tipificacion","")
                dia_badge = f'<span style="background:#d1fae5;color:#065f46;padding:1px 8px;border-radius:12px;font-size:11px;">✅ {tip_txt}</span>'
            elif plazo_str == "VENCIDO":
                etiqueta  = "URGENTE" if not es_rrhh else "VENCIDO"
                dia_badge = f'<span style="background:#fee2e2;color:#991b1b;padding:1px 8px;border-radius:12px;font-size:11px;">🚨 {etiqueta}</span>'
            else:
                dr_txt    = f"{dias_restantes}d restante{'s' if dias_restantes != 1 else ''}" if dias_restantes > 0 else "vence hoy"
                dia_badge = f'<span style="background:#fef3c7;color:#92400e;padding:1px 8px;border-radius:12px;font-size:11px;">⏳ {dr_txt}</span>'

            lic = str(row.get("LICENCIA","")).strip()
            lic_html = f'&nbsp;<span style="color:#999;font-size:12px;">Naaloo: {lic}</span>' if lic and lic not in ("nan","") else ""

            st.markdown(
                f'&nbsp;&nbsp;&nbsp;📅 **{row["FECHA"]}** &nbsp; '
                f'Mañana: **{fmt_fichada(row["MAÑANA"])}** · Tarde: **{fmt_fichada(row["TARDE"])}** '
                f'&nbsp; {dia_badge}{lic_html}',
                unsafe_allow_html=True
            )

            label_exp = f"Ver tipificación — {row['FECHA']}" if est_actual == "RESUELTO" else f"➕ Tipificar — {row['FECHA']}"
            with st.expander(label_exp):
                form_tipificacion(k, gestion_actual, gestiones, "RRHH" if es_rrhh else sucursal)

        st.markdown("</div>", unsafe_allow_html=True)
        st.markdown('<hr style="border:none;border-top:1.5px solid #e5e7eb;margin:12px 0 18px 0;">', unsafe_allow_html=True)

# ── EXPORTAR (RRHH) ──────────────────────────────────────────────────────────
if es_rrhh and not df_ausentes.empty:
    st.markdown("---")
    st.markdown("### 📊 Exportar reporte completo")

    rows_exp = []
    for _, row in df_ausentes.iterrows():
        k = row["_KEY"]
        est, g = estado_gestion_fn(k, gestiones)
        plazo, _ = calcular_plazo(str(row["FECHA"]))
        rows_exp.append({
            "FECHA":           row["FECHA"],
            "LEGAJO":          row["LEGAJO"],
            "SUCURSAL":        row["SUCURSAL"],
            "NOMBRE":          row["NOMBRE"],
            "POSICION":        row.get("POSICION",""),
            "MAÑANA":          fmt_fichada(row["MAÑANA"]),
            "TARDE":           fmt_fichada(row["TARDE"]),
            "LICENCIA_NAALOO": row.get("LICENCIA",""),
            "DIAS_EN_FALTA":   dias_por_legajo.get(row["LEGAJO"], 1),
            "ESTADO":          est if est == "RESUELTO" else plazo,
            "TIPIFICACION":    g.get("tipificacion",""),
            "NAALOO_CARGADO":  g.get("naaloo",""),
            "OBSERVACIONES":   g.get("observaciones",""),
            "RESPONSABLE":     g.get("responsable",""),
            "FECHA_GESTION":   g.get("fecha_gestion",""),
        })

    df_exp = pd.DataFrame(rows_exp)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_exp.to_excel(writer, index=False, sheet_name="Ausencias")
        resumen = df_ausentes.groupby("SUCURSAL")["_ESTADO"].value_counts().unstack(fill_value=0)
        for cn in ["VENCIDO","PENDIENTE","RESUELTO"]:
            if cn not in resumen.columns:
                resumen[cn] = 0
        resumen = resumen[["VENCIDO","PENDIENTE","RESUELTO"]].reset_index()
        resumen.columns = ["Sucursal","Vencidos","Pendientes","Tipificados"]
        resumen.to_excel(writer, index=False, sheet_name="Resumen por sucursal")

    ce, cr = st.columns(2)
    with ce:
        st.download_button(
            "⬇️ Descargar Excel completo",
            data=buf.getvalue(),
            file_name=f"ausencias_{date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    with cr:
        resumen_v = df_ausentes.groupby("SUCURSAL")["_ESTADO"].value_counts().unstack(fill_value=0)
        for cn in ["VENCIDO","PENDIENTE","RESUELTO"]:
            if cn not in resumen_v.columns:
                resumen_v[cn] = 0
        resumen_v = resumen_v[["VENCIDO","PENDIENTE","RESUELTO"]].reset_index()
        resumen_v.columns = ["Sucursal","🚨 Vencidos","⏳ Pendientes","✅ Tipificados"]
        st.dataframe(resumen_v, use_container_width=True, hide_index=True)
                resumen_v[cn] = 0
        resumen_v = resumen_v[["VENCIDO","PENDIENTE","RESUELTO"]].reset_index()
        resumen_v.columns = ["Sucursal","🚨 Vencidos","⏳ Pendientes","✅ Tipificados"]
        st.dataframe(resumen_v, use_container_width=True, hide_index=True)
