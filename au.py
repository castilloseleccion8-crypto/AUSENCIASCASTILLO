import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import json
import gspread
from google.oauth2.service_account import Credentials
import io

st.set_page_config(page_title="Gestión de Ausencias", layout="wide", page_icon="📋")

st.markdown("""
<style>
html, body, [class*="css"] { font-family: 'Segoe UI', sans-serif; }
.stApp { background-color: #f0f2f5; }

/* Sidebar oscuro SIN pisar el área principal */
section[data-testid="stSidebar"] { background-color: #0f1f3d !important; }
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] div.stRadio > label,
section[data-testid="stSidebar"] div.stSelectbox label,
section[data-testid="stSidebar"] div.stTextInput label,
section[data-testid="stSidebar"] .stMarkdown { color: white !important; }

/* Inputs del sidebar con fondo oscuro */
section[data-testid="stSidebar"] input { 
    background-color: #1e3a6e !important; 
    color: white !important; 
    border-color: #3b5ea6 !important;
}
section[data-testid="stSidebar"] select { 
    background-color: #1e3a6e !important; 
    color: white !important;
}
section[data-testid="stSidebar"] div[data-baseweb="select"] > div {
    background-color: #1e3a6e !important;
    color: white !important;
}

/* Botones generales */
div.stButton > button { 
    background-color: #0f1f3d; color: white; border-radius: 8px;
    padding: 8px 20px; border: none; font-weight: 600; width: 100%; 
}
div.stButton > button:hover { background-color: #1e3a6e; }

/* Métricas */
.metric-box { 
    background: white; border-radius: 10px; padding: 14px;
    text-align: center; box-shadow: 0 2px 8px rgba(0,0,0,0.06); 
}
.metric-num   { font-size: 30px; font-weight: 800; }
.metric-label { font-size: 12px; color: #666; margin-top: 3px; }
</style>
""", unsafe_allow_html=True)

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
PASSWORD_RRHH = "rrhh2026"
HOJA_REPORTE   = "REPORTE"
HOJA_GESTIONES = "GESTIONES"
COLS_REPORTE = ["FECHA", "LEGAJO", "SUCURSAL", "NOMBRE", "MAÑANA", "TARDE", "SITUACION", "LICENCIA"]
COLS_GESTIONES = ["KEY", "TIPIFICACION", "NAALOO", "OBSERVACIONES", "RESPONSABLE", "FECHA_GESTION", "SUCURSAL_ORIGEN"]


def render_card(nombre, legajo, fecha, manana, tarde, licencia, suc_label, border_color, badge_html):
    """Renderiza una card de ausencia usando componentes nativos de Streamlit (sin HTML crudo)."""
    color_map = {"#dc2626": "🚨", "#f59e0b": "⏳", "#10b981": "✅"}
    
    with st.container():
        # Línea de color simulada con barra lateral usando columnas
        col_borde, col_contenido = st.columns([0.015, 0.985])
        with col_borde:
            st.markdown(
                f'<div style="background:{border_color};height:80px;border-radius:4px;"></div>',
                unsafe_allow_html=True
            )
        with col_contenido:
            c_nombre, c_badge = st.columns([3, 1])
            with c_nombre:
                suc_txt = f"  ·  📍 {suc_label}" if suc_label else ""
                st.markdown(f"**👤 {nombre}** &nbsp; Leg. {legajo}{suc_txt}")
            with c_badge:
                st.markdown(
                    f'<div style="text-align:right">{badge_html}</div>',
                    unsafe_allow_html=True
                )
            lic_txt = f"&nbsp;&nbsp;📄 Naaloo previo: *{licencia}*" if licencia and licencia != "nan" else ""
            st.markdown(
                f"📅 **{fecha}** &nbsp;&nbsp; 🌅 Mañana: **{manana}** &nbsp;&nbsp; 🌆 Tarde: **{tarde}**{lic_txt}"
            )
        st.markdown('<hr style="margin:4px 0 12px 0;border:none;border-top:1px solid #e5e7eb;">', unsafe_allow_html=True)


def badge_html(texto, bg, color):
    return (f'<span style="background:{bg};color:{color};padding:3px 12px;'
            f'border-radius:20px;font-size:12px;font-weight:600;">{texto}</span>')


@st.cache_resource
def get_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
    return gspread.authorize(creds)


def get_sheet(nombre_hoja):
    gc = get_client()
    sh = gc.open_by_key(st.secrets["SHEET_ID"])
    try:
        return sh.worksheet(nombre_hoja)
    except gspread.WorksheetNotFound:
        return sh.add_worksheet(title=nombre_hoja, rows=2000, cols=20)


@st.cache_data(ttl=60)
def cargar_reporte():
    ws = get_sheet(HOJA_REPORTE)
    data = ws.get_all_records()
    if not data:
        return pd.DataFrame(columns=COLS_REPORTE)
    return pd.DataFrame(data).fillna("")


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
    keys_existentes = [r.get("KEY") for r in todas]
    fila = [key, datos.get("tipificacion",""), datos.get("naaloo",""),
            datos.get("observaciones",""), datos.get("responsable",""),
            datos.get("fecha_gestion",""), datos.get("sucursal_origen","")]
    if key in keys_existentes:
        idx = keys_existentes.index(key) + 2
        ws.update(f"A{idx}:G{idx}", [fila])
    else:
        if not todas:
            ws.update("A1:G1", [COLS_GESTIONES])
        ws.append_row(fila)


def eliminar_gestion(key):
    ws = get_sheet(HOJA_GESTIONES)
    todas = ws.get_all_records()
    keys_ex = [r.get("KEY") for r in todas]
    if key in keys_ex:
        ws.delete_rows(keys_ex.index(key) + 2)


def eliminar_filas_reporte(indices, df_completo):
    df_nuevo = df_completo.drop(index=indices).reset_index(drop=True)
    subir_reporte(df_nuevo)
    return df_nuevo


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


def form_tipificacion(key, gestion_actual, gestiones_dict, row, es_rrhh):
    c1, c2 = st.columns(2)
    with c1:
        tip_actual = gestion_actual.get("tipificacion", TIPIFICACIONES[0])
        tip_idx = TIPIFICACIONES.index(tip_actual) if tip_actual in TIPIFICACIONES else 0
        tipificacion = st.selectbox("Tipo de ausencia *", TIPIFICACIONES, index=tip_idx, key=f"tip_{key}")
    with c2:
        n_actual = gestion_actual.get("naaloo", NAALOO_OPTS[0])
        n_idx = NAALOO_OPTS.index(n_actual) if n_actual in NAALOO_OPTS else 0
        naaloo = st.selectbox("¿Cargado en Naaloo? *", NAALOO_OPTS, index=n_idx, key=f"naaloo_{key}")

    obs = st.text_area("Observaciones", value=gestion_actual.get("observaciones",""),
                       placeholder="Ej: Avisó por WhatsApp, presentó certificado el...",
                       key=f"obs_{key}", height=70)
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
                    "sucursal_origen": "RRHH" if es_rrhh else str(row["SUCURSAL"]),
                }
                guardar_gestion(key, datos)
                gestiones_dict[key] = datos
                st.success("✅ Guardado en Drive")
                st.cache_data.clear()
                st.rerun()

    if gestion_actual.get("fecha_gestion"):
        st.caption(f"Guardado: {gestion_actual['fecha_gestion']} — {gestion_actual.get('responsable','')} ({gestion_actual.get('sucursal_origen','')})")


# ====== SIDEBAR ======
with st.sidebar:
    st.markdown("## 🔐 Acceso")
    modo = st.radio("Tipo de acceso", ["Sucursal", "RRHH (Global)"])

    df_temp = cargar_reporte()
    sucursales_disponibles = sorted(df_temp["SUCURSAL"].dropna().unique().tolist()) if not df_temp.empty else []

    if modo == "Sucursal":
        sucursal_sel = st.selectbox("Sucursal", sucursales_disponibles) if sucursales_disponibles else None
        password = st.text_input("Contraseña", type="password")
        es_rrhh = False
        if sucursal_sel:
            slug = (sucursal_sel.lower()
                    .replace(" ","").replace(".","")
                    .replace("ü","u").replace("é","e").replace("ó","o")
                    .replace("á","a").replace("í","i").replace("ú","u"))
            autenticado = password == f"{slug}{date.today().year}"
        else:
            autenticado = False
    else:
        sucursal_sel = None
        password = st.text_input("Contraseña RRHH", type="password")
        es_rrhh = True
        autenticado = password == PASSWORD_RRHH

    if password and not autenticado:
        st.error("❌ Contraseña incorrecta")
    st.markdown("---")
    st.caption(f"📅 {date.today().strftime('%d/%m/%Y')}")

# ====== HEADER ======
st.markdown('<h1 style="text-align:center;color:#0f1f3d;font-size:34px;font-weight:800;margin-bottom:4px;">📋 GESTIÓN DE AUSENCIAS</h1>', unsafe_allow_html=True)
st.markdown('<p style="text-align:center;color:#666;font-size:15px;margin-bottom:20px;">Tipificación y seguimiento — plazo 48hs</p>', unsafe_allow_html=True)

if not password:
    st.info("👈 Seleccioná tu sucursal e ingresá la contraseña para continuar.")
    st.stop()
if not autenticado:
    st.stop()

df_global  = cargar_reporte()
gestiones  = cargar_gestiones()

# ====== ADMIN RRHH ======
if es_rrhh:
    with st.expander("⚙️ Administración — Subir reporte / Gestión de casos", expanded=df_global.empty):
        tab_subir, tab_manual = st.tabs(["📤 Subir Excel", "✏️ Gestión de casos"])

        with tab_subir:
            st.markdown("**Subí el Excel generado por tu script.** Reemplaza todos los datos del reporte.")
            archivo = st.file_uploader("Seleccioná el archivo .xlsx", type=["xlsx"], key="up_reporte")
            if archivo is not None:
                try:
                    df_preview = excel_a_df(archivo.getvalue())
                    aus_prev = df_preview[df_preview["SITUACION"].str.contains("AUSENTE", case=False, na=False)]
                    st.success(f"✅ {len(df_preview)} filas leídas — {len(aus_prev)} ausentes detectados")
                    st.dataframe(df_preview.head(10), use_container_width=True, hide_index=True)
                    cb1, _ = st.columns([1, 3])
                    with cb1:
                        if st.button("⬆️ Confirmar y subir", key="btn_subir"):
                            with st.spinner("Subiendo..."):
                                subir_reporte(df_preview)
                            st.success("✅ Reporte actualizado.")
                            st.cache_data.clear()
                            st.rerun()
                except Exception as e:
                    st.error(f"Error al leer el archivo: {e}")

        with tab_manual:
            st.markdown("**Eliminá casos** que ya están cerrados. Borra la fila y su tipificación.")
            if df_global.empty:
                st.info("No hay datos cargados.")
            else:
                df_adm = df_global[df_global["SITUACION"].str.contains("AUSENTE", case=False, na=False)].copy()
                df_adm["_ESTADO"] = df_adm.apply(lambda r: estado_fila(r, gestiones), axis=1)
                df_adm["_KEY"] = df_adm.apply(clave, axis=1)
                suc_adm = st.selectbox("Filtrar sucursal", ["Todas"] + sucursales_disponibles, key="adm_suc")
                if suc_adm != "Todas":
                    df_adm = df_adm[df_adm["SUCURSAL"] == suc_adm]
                seleccionados = []
                for i, (orig_idx, row) in enumerate(df_adm.iterrows()):
                    k = row["_KEY"]
                    est = row["_ESTADO"]
                    g = gestiones.get(k, {})
                    est_txt = f"✅ {g.get('tipificacion','')}" if est == "RESUELTO" else ("🚨 VENCIDO" if est == "VENCIDO" else "⏳ PENDIENTE")
                    lic = str(row.get("LICENCIA","")).strip()
                    lic_txt = f" · {lic}" if lic and lic != "nan" else ""
                    cc, ci = st.columns([0.5, 9.5])
                    with cc:
                        elegido = st.checkbox("", key=f"chk_{k}_{i}", label_visibility="collapsed")
                    with ci:
                        st.markdown(f"**{row['NOMBRE'].strip()}** — Leg. {row['LEGAJO']} · {row['SUCURSAL']} · {row['FECHA']} · {est_txt}{lic_txt}")
                    if elegido:
                        seleccionados.append((orig_idx, k))
                if seleccionados:
                    st.warning(f"Eliminás {len(seleccionados)} caso(s). No se puede deshacer.")
                    cd, _ = st.columns([1, 4])
                    with cd:
                        if st.button("🗑️ Eliminar seleccionados", key="btn_del"):
                            with st.spinner("Eliminando..."):
                                eliminar_filas_reporte([s[0] for s in seleccionados], df_global)
                                for _, k in seleccionados:
                                    if k in gestiones:
                                        eliminar_gestion(k)
                            st.success("✅ Eliminados.")
                            st.cache_data.clear()
                            st.rerun()
    st.markdown("---")

if df_global.empty:
    st.warning("No hay datos. RRHH debe subir el reporte primero.")
    st.stop()

sucursales_disponibles = sorted(df_global["SUCURSAL"].dropna().unique().tolist())
df_vista   = df_global.copy() if es_rrhh else df_global[df_global["SUCURSAL"] == sucursal_sel].copy()
df_ausentes = df_vista[df_vista["SITUACION"].str.contains("AUSENTE", case=False, na=False)].copy()
df_ausentes["_ESTADO"] = df_ausentes.apply(lambda r: estado_fila(r, gestiones), axis=1)

titulo_vista = "Vista Global — Todas las Sucursales" if es_rrhh else f"Sucursal: {sucursal_sel}"
st.markdown(f"### {titulo_vista}")

total      = len(df_ausentes)
resueltos  = int((df_ausentes["_ESTADO"] == "RESUELTO").sum())
vencidos   = int((df_ausentes["_ESTADO"] == "VENCIDO").sum())
pendientes = int((df_ausentes["_ESTADO"] == "PENDIENTE").sum())

c1, c2, c3, c4 = st.columns(4)
c1.markdown(f'<div class="metric-box"><div class="metric-num" style="color:#0f1f3d">{total}</div><div class="metric-label">Total</div></div>', unsafe_allow_html=True)
c2.markdown(f'<div class="metric-box"><div class="metric-num" style="color:#f59e0b">{pendientes}</div><div class="metric-label">⏳ Pendientes</div></div>', unsafe_allow_html=True)
c3.markdown(f'<div class="metric-box"><div class="metric-num" style="color:#dc2626">{vencidos}</div><div class="metric-label">🚨 Vencidos +48hs</div></div>', unsafe_allow_html=True)
c4.markdown(f'<div class="metric-box"><div class="metric-num" style="color:#10b981">{resueltos}</div><div class="metric-label">✅ Tipificados</div></div>', unsafe_allow_html=True)
st.markdown("<br>", unsafe_allow_html=True)

# ====== FILTROS ======
cf1, cf2, cf3 = st.columns([2, 2, 2])
with cf1:
    buscar = st.text_input("🔎 Buscar nombre o legajo")
with cf2:
    filtro_estado = st.selectbox("Estado", ["Todos","Pendientes","Vencidos (+48hs)","Tipificados"])
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
if filtro_estado == "Pendientes":
    df_filtrado = df_filtrado[df_filtrado["_ESTADO"] == "PENDIENTE"]
elif filtro_estado == "Vencidos (+48hs)":
    df_filtrado = df_filtrado[df_filtrado["_ESTADO"] == "VENCIDO"]
elif filtro_estado == "Tipificados":
    df_filtrado = df_filtrado[df_filtrado["_ESTADO"] == "RESUELTO"]

orden_map = {"VENCIDO":0, "PENDIENTE":1, "RESUELTO":2}
df_filtrado = df_filtrado.copy()
df_filtrado["_ORDEN"] = df_filtrado["_ESTADO"].map(orden_map)
df_filtrado = df_filtrado.sort_values(["_ORDEN","FECHA","SUCURSAL","NOMBRE"]).reset_index(drop=True)

st.markdown(f"**{len(df_filtrado)} registros**")
st.markdown("---")

# ====== LISTA ======
if df_filtrado.empty:
    st.info("Sin registros para los filtros seleccionados.")
else:
    for _, row in df_filtrado.iterrows():
        k = clave(row)
        est_actual, gestion_actual = estado_gestion_fn(k, gestiones)
        plazo_str, dias = calcular_plazo(str(row["FECHA"]))

        if est_actual == "RESUELTO":
            border_col = "#10b981"
            bh = badge_html("✅ " + gestion_actual.get("tipificacion",""), "#d1fae5", "#065f46")
        elif plazo_str == "VENCIDO":
            border_col = "#dc2626"
            bh = badge_html("🚨 VENCIDO (+48hs)", "#fee2e2", "#991b1b")
        else:
            border_col = "#f59e0b"
            dt = f"{dias} día(s)" if dias > 0 else "vence hoy"
            bh = badge_html(f"⏳ PENDIENTE — {dt}", "#fef3c7", "#92400e")

        suc_label = row["SUCURSAL"] if es_rrhh else ""
        lic_str   = str(row.get("LICENCIA","")).strip()
        lic_clean = lic_str if lic_str and lic_str != "nan" else ""

        render_card(
            nombre=str(row["NOMBRE"]).strip(),
            legajo=row["LEGAJO"],
            fecha=row["FECHA"],
            manana=row["MAÑANA"],
            tarde=row["TARDE"],
            licencia=lic_clean,
            suc_label=suc_label,
            border_color=border_col,
            badge_html=bh
        )

        label_exp = "Ver / Editar tipificación" if est_actual == "RESUELTO" else "➕ Tipificar ausencia"
        with st.expander(label_exp):
            form_tipificacion(k, gestion_actual, gestiones, row, es_rrhh)

# ====== EXPORTAR ======
if es_rrhh and not df_ausentes.empty:
    st.markdown("---")
    st.markdown("### 📊 Exportar")
    rows_exp = []
    for _, row in df_ausentes.iterrows():
        k = clave(row)
        est, g = estado_gestion_fn(k, gestiones)
        plazo, _ = calcular_plazo(str(row["FECHA"]))
        rows_exp.append({
            "FECHA":row["FECHA"],"LEGAJO":row["LEGAJO"],"NOMBRE":row["NOMBRE"],
            "SUCURSAL":row["SUCURSAL"],"MAÑANA":row["MAÑANA"],"TARDE":row["TARDE"],
            "NAALOO_PREVIO":row.get("LICENCIA",""),
            "ESTADO": est if est=="RESUELTO" else plazo,
            "TIPIFICACION":g.get("tipificacion",""),"NAALOO_CARGADO":g.get("naaloo",""),
            "OBSERVACIONES":g.get("observaciones",""),"RESPONSABLE":g.get("responsable",""),
            "FECHA_GESTION":g.get("fecha_gestion",""),
        })
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        pd.DataFrame(rows_exp).to_excel(writer, index=False, sheet_name="Ausencias")
    ce, cr = st.columns(2)
    with ce:
        st.download_button("⬇️ Descargar Excel", data=buf.getvalue(),
            file_name=f"ausencias_{date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with cr:
        resumen = df_ausentes.groupby("SUCURSAL")["_ESTADO"].value_counts().unstack(fill_value=0)
        for cn in ["VENCIDO","PENDIENTE","RESUELTO"]:
            if cn not in resumen.columns: resumen[cn] = 0
        resumen = resumen[["VENCIDO","PENDIENTE","RESUELTO"]].reset_index()
        resumen.columns = ["Sucursal","🚨 Vencidos","⏳ Pendientes","✅ Tipificados"]
        st.dataframe(resumen, use_container_width=True, hide_index=True)
