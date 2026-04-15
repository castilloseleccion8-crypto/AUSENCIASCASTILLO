import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import json
import gspread
from google.oauth2.service_account import Credentials
import io

st.set_page_config(page_title="Gestión de Ausencias", layout="wide", page_icon="📋")

# ===================== ESTILOS =====================
st.markdown("""
<style>
html, body, [class*="css"] { font-family: 'Segoe UI', sans-serif; }
.stApp { background-color: #f0f2f5; }
section[data-testid="stSidebar"] { background-color: #0f1f3d; }
section[data-testid="stSidebar"] * { color: white !important; }
.titulo { font-size: 34px; font-weight: 800; color: #0f1f3d; text-align: center; margin-bottom: 4px; }
.subtitulo { font-size: 16px; color: #666; text-align: center; margin-bottom: 20px; }
.card { background: white; padding: 18px 20px; border-radius: 12px;
        box-shadow: 0 3px 10px rgba(0,0,0,0.07); margin-bottom: 10px;
        border-left: 4px solid #e5e7eb; }
.card.vencido  { border-left-color: #dc2626; }
.card.pendiente{ border-left-color: #f59e0b; }
.card.resuelto { border-left-color: #10b981; }
.badge-rojo     { background:#fee2e2; color:#dc2626; padding:3px 10px; border-radius:20px; font-size:12px; font-weight:700; }
.badge-amarillo { background:#fef3c7; color:#d97706; padding:3px 10px; border-radius:20px; font-size:12px; font-weight:700; }
.badge-verde    { background:#d1fae5; color:#059669; padding:3px 10px; border-radius:20px; font-size:12px; font-weight:700; }
div.stButton > button { background-color:#0f1f3d; color:white; border-radius:8px;
                        padding:8px 20px; border:none; font-weight:600; width:100%; }
div.stButton > button:hover { background-color:#1e3a6e; }
.metric-box { background:white; border-radius:10px; padding:14px;
              text-align:center; box-shadow:0 2px 8px rgba(0,0,0,0.06); }
.metric-num   { font-size:30px; font-weight:800; }
.metric-label { font-size:12px; color:#666; margin-top:3px; }
div[data-testid="stFileUploader"] { border: 2px dashed #0f1f3d33;
    border-radius: 10px; padding: 10px; background: white; }
</style>
""", unsafe_allow_html=True)

# ===================== CONSTANTES =====================
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

# Nombres de las hojas dentro de la Google Sheet
HOJA_REPORTE   = "REPORTE"    # datos del Excel que se sube
HOJA_GESTIONES = "GESTIONES"  # tipificaciones guardadas

# ===================== GOOGLE SHEETS =====================
@st.cache_resource
def get_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=scopes
    )
    return gspread.authorize(creds)


def get_sheet(nombre_hoja):
    gc = get_client()
    sh = gc.open_by_key(st.secrets["SHEET_ID"])
    try:
        return sh.worksheet(nombre_hoja)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=nombre_hoja, rows=2000, cols=20)
        return ws


# ---- REPORTE ----
COLS_REPORTE = ["FECHA", "LEGAJO", "SUCURSAL", "NOMBRE", "MAÑANA", "TARDE", "SITUACION", "LICENCIA"]

def cargar_reporte_desde_drive():
    ws = get_sheet(HOJA_REPORTE)
    data = ws.get_all_records()
    if not data:
        return pd.DataFrame(columns=COLS_REPORTE)
    df = pd.DataFrame(data)
    df = df.fillna("")
    return df


def subir_reporte_a_drive(df: pd.DataFrame):
    ws = get_sheet(HOJA_REPORTE)
    ws.clear()
    # Asegurar columnas mínimas
    for col in COLS_REPORTE:
        if col not in df.columns:
            df[col] = ""
    df = df[COLS_REPORTE].fillna("").astype(str)
    ws.update([df.columns.tolist()] + df.values.tolist())


def eliminar_filas_reporte(indices_a_eliminar: list, df_completo: pd.DataFrame):
    """Elimina filas del reporte por índice y reescribe la hoja."""
    df_nuevo = df_completo.drop(index=indices_a_eliminar).reset_index(drop=True)
    subir_reporte_a_drive(df_nuevo)
    return df_nuevo


# ---- GESTIONES ----
COLS_GESTIONES = ["KEY", "TIPIFICACION", "NAALOO", "OBSERVACIONES",
                  "RESPONSABLE", "FECHA_GESTION", "SUCURSAL_ORIGEN"]

def cargar_gestiones_desde_drive():
    ws = get_sheet(HOJA_GESTIONES)
    data = ws.get_all_records()
    if not data:
        return {}
    return {row["KEY"]: {k.lower(): v for k, v in row.items() if k != "KEY"}
            for row in data if row.get("KEY")}


def guardar_gestion_en_drive(key: str, datos: dict):
    ws = get_sheet(HOJA_GESTIONES)
    todas = ws.get_all_records()
    keys_existentes = [r.get("KEY") for r in todas]

    fila = [
        key,
        datos.get("tipificacion", ""),
        datos.get("naaloo", ""),
        datos.get("observaciones", ""),
        datos.get("responsable", ""),
        datos.get("fecha_gestion", ""),
        datos.get("sucursal_origen", ""),
    ]

    if key in keys_existentes:
        idx = keys_existentes.index(key) + 2  # +2: header en fila 1, 1-indexed
        ws.update(f"A{idx}:G{idx}", [fila])
    else:
        if not todas:  # primera vez: poner encabezado
            ws.update("A1:G1", [COLS_GESTIONES])
        ws.append_row(fila)


def eliminar_gestion_de_drive(key: str):
    ws = get_sheet(HOJA_GESTIONES)
    todas = ws.get_all_records()
    keys_existentes = [r.get("KEY") for r in todas]
    if key in keys_existentes:
        idx = keys_existentes.index(key) + 2
        ws.delete_rows(idx)


# ===================== HELPERS =====================
def calcular_estado_plazo(fecha_str):
    try:
        fecha = datetime.strptime(str(fecha_str).strip(), "%d/%m/%Y").date()
    except Exception:
        return "PENDIENTE", 99
    limite = fecha + timedelta(days=2)
    dias = (limite - date.today()).days
    return ("VENCIDO" if date.today() > limite else "PENDIENTE"), dias


def clave_registro(row):
    return f"{str(row['FECHA']).strip()}_{str(row['LEGAJO']).strip()}_{str(row['SUCURSAL']).strip()}"


def estado_gestion(key, gestiones):
    g = gestiones.get(key, {})
    if g.get("tipificacion"):
        return "RESUELTO", g
    return "PENDIENTE", g


def get_estado_fila(row, gestiones):
    key = clave_registro(row)
    est, _ = estado_gestion(key, gestiones)
    if est == "RESUELTO":
        return "RESUELTO"
    plazo, _ = calcular_estado_plazo(row["FECHA"])
    return plazo


def excel_a_df(archivo_bytes):
    """Convierte bytes de un Excel (multi-hoja) al DataFrame unificado del reporte."""
    sheets = pd.read_excel(io.BytesIO(archivo_bytes), sheet_name=None)
    frames = []
    for _, df in sheets.items():
        df.columns = [str(c).strip() for c in df.columns]
        df = df.fillna("")
        # Normalizar columnas al formato esperado
        df = df.rename(columns={"MAÑANA": "MAÑANA"})  # por si viene con tilde rara
        for col in COLS_REPORTE:
            if col not in df.columns:
                df[col] = ""
        frames.append(df[COLS_REPORTE])
    if not frames:
        return pd.DataFrame(columns=COLS_REPORTE)
    return pd.concat(frames, ignore_index=True).astype(str)


def mostrar_formulario(key, gestion_actual, gestiones_dict, row, es_rrhh):
    col1, col2 = st.columns(2)
    with col1:
        tip_actual = gestion_actual.get("tipificacion", TIPIFICACIONES[0])
        tip_idx = TIPIFICACIONES.index(tip_actual) if tip_actual in TIPIFICACIONES else 0
        tipificacion = st.selectbox("Tipo de ausencia *", TIPIFICACIONES, index=tip_idx, key=f"tip_{key}")
    with col2:
        naaloo_actual = gestion_actual.get("naaloo", NAALOO_OPTS[0])
        naaloo_idx = NAALOO_OPTS.index(naaloo_actual) if naaloo_actual in NAALOO_OPTS else 0
        naaloo = st.selectbox("¿Cargado en Naaloo? *", NAALOO_OPTS, index=naaloo_idx, key=f"naaloo_{key}")

    observaciones = st.text_area(
        "Observaciones",
        value=gestion_actual.get("observaciones", ""),
        placeholder="Ej: Avisó por WhatsApp, presentó certificado médico el...",
        key=f"obs_{key}", height=70,
    )
    responsable = st.text_input(
        "Responsable *",
        value=gestion_actual.get("responsable", ""),
        placeholder="Nombre de quien tipifica",
        key=f"resp_{key}",
    )

    col_btn, _ = st.columns([1, 3])
    with col_btn:
        if st.button("💾 Guardar", key=f"btn_{key}"):
            if not responsable.strip():
                st.error("El campo Responsable es obligatorio.")
            else:
                datos = {
                    "tipificacion": tipificacion,
                    "naaloo": naaloo,
                    "observaciones": observaciones.strip(),
                    "responsable": responsable.strip(),
                    "fecha_gestion": datetime.now().strftime("%d/%m/%Y %H:%M"),
                    "sucursal_origen": "RRHH" if es_rrhh else str(row["SUCURSAL"]),
                }
                guardar_gestion_en_drive(key, datos)
                gestiones_dict[key] = datos
                st.success("✅ Guardado en Drive")
                st.cache_data.clear()
                st.rerun()

    if gestion_actual.get("fecha_gestion"):
        st.caption(
            f"Guardado: {gestion_actual['fecha_gestion']} "
            f"— {gestion_actual.get('responsable','')} "
            f"({gestion_actual.get('sucursal_origen','')})"
        )


# ===================== CARGA CON CACHE =====================
@st.cache_data(ttl=60)
def _cargar_reporte():
    return cargar_reporte_desde_drive()

@st.cache_data(ttl=30)
def _cargar_gestiones():
    return cargar_gestiones_desde_drive()

# ===================== SIDEBAR =====================
with st.sidebar:
    st.markdown("## 🔐 Acceso")
    modo = st.radio("Tipo de acceso", ["Sucursal", "RRHH (Global)"])

    if modo == "Sucursal":
        # Cargamos lista de sucursales para el selector
        df_temp = _cargar_reporte()
        sucursales_disponibles = sorted(df_temp["SUCURSAL"].dropna().unique().tolist()) if not df_temp.empty else []
        sucursal_sel = st.selectbox("Sucursal", sucursales_disponibles) if sucursales_disponibles else None
        password = st.text_input("Contraseña", type="password")
        es_rrhh = False
        if sucursal_sel:
            slug = (sucursal_sel.lower()
                    .replace(' ','').replace('.','')
                    .replace('ü','u').replace('é','e').replace('ó','o')
                    .replace('á','a').replace('í','i').replace('ú','u'))
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

# ===================== HEADER =====================
st.markdown('<div class="titulo">📋 GESTIÓN DE AUSENCIAS</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitulo">Tipificación y seguimiento — plazo 48hs</div>', unsafe_allow_html=True)

if not password:
    st.info("👈 Seleccioná tu sucursal e ingresá la contraseña para continuar.")
    st.stop()
if not autenticado:
    st.stop()

# ===================== CARGA DE DATOS =====================
df_global   = _cargar_reporte()
gestiones   = _cargar_gestiones()

if df_global.empty:
    if es_rrhh:
        st.warning("No hay datos cargados aún. Subí el reporte desde la sección de administración abajo.")
    else:
        st.warning("No hay datos disponibles. Contactá a RRHH.")
    if not es_rrhh:
        st.stop()

sucursales_disponibles = sorted(df_global["SUCURSAL"].dropna().unique().tolist()) if not df_global.empty else []

# ===================== PANEL RRHH — ADMIN =====================
if es_rrhh:
    with st.expander("⚙️ Administración — Subir nuevo reporte / Gestión de casos", expanded=df_global.empty):
        tab_subir, tab_manual = st.tabs(["📤 Subir Excel", "✏️ Gestión manual de casos"])

        # ---- TAB SUBIR EXCEL ----
        with tab_subir:
            st.markdown("**Subí el Excel generado por tu script Python.** Reemplaza todos los datos del reporte actual.")
            st.caption("Formato esperado: múltiples hojas, una por sucursal, con columnas: FECHA, LEGAJO, SUCURSAL, NOMBRE, MAÑANA, TARDE, SITUACION, LICENCIA")

            archivo = st.file_uploader(
                "Seleccioná el archivo .xlsx",
                type=["xlsx"],
                key="uploader_reporte",
            )

            if archivo is not None:
                try:
                    df_preview = excel_a_df(archivo.getvalue())
                    ausentes_preview = df_preview[df_preview["SITUACION"].str.contains("AUSENTE", case=False, na=False)]
                    st.success(f"✅ Archivo leído: **{len(df_preview)}** filas totales — **{len(ausentes_preview)}** ausentes detectados")
                    st.dataframe(df_preview.head(10), use_container_width=True, hide_index=True)

                    col_conf1, col_conf2 = st.columns([1, 3])
                    with col_conf1:
                        if st.button("⬆️ Confirmar y subir a Drive", key="btn_subir_excel"):
                            with st.spinner("Subiendo a Google Drive..."):
                                subir_reporte_a_drive(df_preview)
                            st.success("✅ Reporte actualizado en Drive.")
                            st.cache_data.clear()
                            st.rerun()
                except Exception as e:
                    st.error(f"Error al leer el archivo: {e}")

        # ---- TAB GESTIÓN MANUAL ----
        with tab_manual:
            st.markdown("**Eliminá casos del reporte** que ya están cerrados o que no corresponden.")
            st.caption("Esto borra la fila del reporte Y su tipificación asociada. Es irreversible.")

            if df_global.empty:
                st.info("No hay datos en el reporte.")
            else:
                df_ausentes_admin = df_global[df_global["SITUACION"].str.contains("AUSENTE", case=False, na=False)].copy()
                df_ausentes_admin["_ESTADO"] = df_ausentes_admin.apply(lambda r: get_estado_fila(r, gestiones), axis=1)
                df_ausentes_admin["_KEY"] = df_ausentes_admin.apply(clave_registro, axis=1)

                # Filtro rápido
                suc_filtro = st.selectbox("Filtrar por sucursal", ["Todas"] + sucursales_disponibles, key="admin_suc")
                df_admin_view = df_ausentes_admin.copy()
                if suc_filtro != "Todas":
                    df_admin_view = df_admin_view[df_admin_view["SUCURSAL"] == suc_filtro]

                if df_admin_view.empty:
                    st.info("Sin registros para mostrar.")
                else:
                    # Mostrar tabla con checkboxes para seleccionar a eliminar
                    st.markdown(f"**{len(df_admin_view)} casos** — Seleccioná los que querés eliminar:")

                    seleccionados = []
                    for i, (orig_idx, row) in enumerate(df_admin_view.iterrows()):
                        key = row["_KEY"]
                        est = row["_ESTADO"]
                        g = gestiones.get(key, {})

                        if est == "RESUELTO":
                            estado_txt = f"✅ {g.get('tipificacion','')}"
                        elif est == "VENCIDO":
                            estado_txt = "🚨 VENCIDO"
                        else:
                            estado_txt = "⏳ PENDIENTE"

                        col_chk, col_info = st.columns([0.5, 9.5])
                        with col_chk:
                            elegido = st.checkbox("", key=f"chk_{key}_{i}", label_visibility="collapsed")
                        with col_info:
                            lic = str(row.get("LICENCIA","")).strip()
                            lic_txt = f" · Naaloo: {lic}" if lic and lic != "nan" else ""
                            st.markdown(
                                f"**{row['NOMBRE'].strip()}** — Leg. {row['LEGAJO']} · {row['SUCURSAL']} · {row['FECHA']} · {estado_txt}{lic_txt}"
                            )

                        if elegido:
                            seleccionados.append((orig_idx, key))

                    if seleccionados:
                        st.warning(f"Vas a eliminar **{len(seleccionados)}** caso(s). Esta acción no se puede deshacer.")
                        col_del, _ = st.columns([1, 4])
                        with col_del:
                            if st.button("🗑️ Eliminar seleccionados", key="btn_eliminar"):
                                with st.spinner("Eliminando..."):
                                    indices = [s[0] for s in seleccionados]
                                    keys_a_borrar = [s[1] for s in seleccionados]
                                    # Eliminar del reporte
                                    eliminar_filas_reporte(indices, df_global)
                                    # Eliminar gestiones asociadas
                                    for k in keys_a_borrar:
                                        if k in gestiones:
                                            eliminar_gestion_de_drive(k)
                                st.success("✅ Casos eliminados.")
                                st.cache_data.clear()
                                st.rerun()

    st.markdown("---")

# ===================== DATOS PARA VISTA PRINCIPAL =====================
if df_global.empty:
    st.stop()

df_vista = df_global.copy() if es_rrhh else df_global[df_global["SUCURSAL"] == sucursal_sel].copy()
df_ausentes = df_vista[df_vista["SITUACION"].str.contains("AUSENTE", case=False, na=False)].copy()
df_ausentes["_ESTADO"] = df_ausentes.apply(lambda r: get_estado_fila(r, gestiones), axis=1)

titulo_vista = "Vista Global — Todas las Sucursales" if es_rrhh else f"Sucursal: {sucursal_sel}"
st.markdown(f"### {titulo_vista}")

# ===================== MÉTRICAS =====================
total     = len(df_ausentes)
resueltos = int((df_ausentes["_ESTADO"] == "RESUELTO").sum())
vencidos  = int((df_ausentes["_ESTADO"] == "VENCIDO").sum())
pendientes= int((df_ausentes["_ESTADO"] == "PENDIENTE").sum())

c1, c2, c3, c4 = st.columns(4)
c1.markdown(f'<div class="metric-box"><div class="metric-num" style="color:#0f1f3d">{total}</div><div class="metric-label">Total</div></div>', unsafe_allow_html=True)
c2.markdown(f'<div class="metric-box"><div class="metric-num" style="color:#f59e0b">{pendientes}</div><div class="metric-label">⏳ Pendientes</div></div>', unsafe_allow_html=True)
c3.markdown(f'<div class="metric-box"><div class="metric-num" style="color:#dc2626">{vencidos}</div><div class="metric-label">🚨 Vencidos +48hs</div></div>', unsafe_allow_html=True)
c4.markdown(f'<div class="metric-box"><div class="metric-num" style="color:#10b981">{resueltos}</div><div class="metric-label">✅ Tipificados</div></div>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ===================== FILTROS =====================
cf1, cf2, cf3 = st.columns([2, 2, 2])
with cf1:
    buscar = st.text_input("🔎 Buscar nombre o legajo", placeholder="Ej: GARCIA o 5402")
with cf2:
    filtro_estado = st.selectbox("Estado", ["Todos", "Pendientes", "Vencidos (+48hs)", "Tipificados"])
with cf3:
    if es_rrhh:
        filtro_suc = st.selectbox("Sucursal", ["Todas"] + sucursales_disponibles)
    else:
        filtro_suc = sucursal_sel

df_filtrado = df_ausentes.copy()
if buscar:
    mask = (
        df_filtrado["NOMBRE"].str.contains(buscar, case=False, na=False) |
        df_filtrado["LEGAJO"].astype(str).str.contains(buscar, na=False)
    )
    df_filtrado = df_filtrado[mask]
if es_rrhh and filtro_suc != "Todas":
    df_filtrado = df_filtrado[df_filtrado["SUCURSAL"] == filtro_suc]
if filtro_estado == "Pendientes":
    df_filtrado = df_filtrado[df_filtrado["_ESTADO"] == "PENDIENTE"]
elif filtro_estado == "Vencidos (+48hs)":
    df_filtrado = df_filtrado[df_filtrado["_ESTADO"] == "VENCIDO"]
elif filtro_estado == "Tipificados":
    df_filtrado = df_filtrado[df_filtrado["_ESTADO"] == "RESUELTO"]

orden_map = {"VENCIDO": 0, "PENDIENTE": 1, "RESUELTO": 2}
df_filtrado = df_filtrado.copy()
df_filtrado["_ORDEN"] = df_filtrado["_ESTADO"].map(orden_map)
df_filtrado = df_filtrado.sort_values(["_ORDEN", "FECHA", "SUCURSAL", "NOMBRE"]).reset_index(drop=True)

st.markdown(f"**{len(df_filtrado)} registros**")
st.markdown("---")

# ===================== LISTA REGISTROS =====================
if df_filtrado.empty:
    st.info("✨ Sin registros para los filtros seleccionados.")
else:
    for _, row in df_filtrado.iterrows():
        key = clave_registro(row)
        est_actual, gestion_actual = estado_gestion(key, gestiones)
        plazo_str, dias = calcular_estado_plazo(str(row["FECHA"]))

        if est_actual == "RESUELTO":
            border_color = "#10b981"
            badge = (f'<span style="background:#d1fae5;color:#065f46;padding:3px 12px;'
                     f'border-radius:20px;font-size:12px;font-weight:600;">'
                     f'✅ {gestion_actual.get("tipificacion","")}</span>')
        elif plazo_str == "VENCIDO":
            border_color = "#dc2626"
            badge = ('<span style="background:#fee2e2;color:#991b1b;padding:3px 12px;'
                     'border-radius:20px;font-size:12px;font-weight:600;">'
                     '🚨 VENCIDO (+48hs)</span>')
        else:
            border_color = "#f59e0b"
            dias_txt = f"{dias} día(s)" if dias > 0 else "vence hoy"
            badge = (f'<span style="background:#fef3c7;color:#92400e;padding:3px 12px;'
                     f'border-radius:20px;font-size:12px;font-weight:600;">'
                     f'⏳ PENDIENTE — {dias_txt}</span>')

        lic_str = str(row.get("LICENCIA","")).strip()
        lic_html = (f'<span style="color:#888;font-size:12px;">📄 Naaloo previo: {lic_str}</span>'
                    if lic_str and lic_str not in ("nan", "") else "")
        suc_html = (f'<span style="color:#888;font-size:12px;">📍 {row["SUCURSAL"]}</span>'
                    if es_rrhh else "")

        st.markdown(f"""
        <div style="background:white;padding:16px 20px;border-radius:10px;
                    border-left:4px solid {border_color};margin-bottom:8px;
                    box-shadow:0 2px 8px rgba(0,0,0,0.06);">
            <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;">
                <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;">
                    <span style="font-size:15px;font-weight:600;color:#111;">
                        👤 {str(row['NOMBRE']).strip()}
                    </span>
                    <span style="color:#666;font-size:13px;">Leg. {row['LEGAJO']}</span>
                    {suc_html}
                </div>
                <div>{badge}</div>
            </div>
            <div style="display:flex;gap:20px;font-size:13px;color:#444;flex-wrap:wrap;align-items:center;">
                <span>📅 <b>{row['FECHA']}</b></span>
                <span>🌅 Mañana: <b>{row['MAÑANA']}</b></span>
                <span>🌆 Tarde: <b>{row['TARDE']}</b></span>
                {lic_html}
            </div>
        </div>
        """, unsafe_allow_html=True)

        label_exp = "Ver / Editar tipificación" if est_actual == "RESUELTO" else "➕ Tipificar ausencia"
        with st.expander(label_exp):
            mostrar_formulario(key, gestion_actual, gestiones, row, es_rrhh)

# ===================== EXPORTAR (solo RRHH) =====================
if es_rrhh and not df_ausentes.empty:
    st.markdown("---")
    st.markdown("### 📊 Exportar")
    col_exp, col_res = st.columns(2)

    rows_export = []
    for _, row in df_ausentes.iterrows():
        key = clave_registro(row)
        est, g = estado_gestion(key, gestiones)
        plazo, _ = calcular_estado_plazo(str(row["FECHA"]))
        rows_export.append({
            "FECHA": row["FECHA"], "LEGAJO": row["LEGAJO"],
            "NOMBRE": row["NOMBRE"], "SUCURSAL": row["SUCURSAL"],
            "MAÑANA": row["MAÑANA"], "TARDE": row["TARDE"],
            "NAALOO_PREVIO": row.get("LICENCIA",""),
            "ESTADO": est if est == "RESUELTO" else plazo,
            "TIPIFICACION": g.get("tipificacion",""),
            "NAALOO_CARGADO": g.get("naaloo",""),
            "OBSERVACIONES": g.get("observaciones",""),
            "RESPONSABLE": g.get("responsable",""),
            "FECHA_GESTION": g.get("fecha_gestion",""),
        })
    df_export = pd.DataFrame(rows_export)
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_export.to_excel(writer, index=False, sheet_name="Ausencias")

    with col_exp:
        st.download_button(
            "⬇️ Descargar Excel con gestiones",
            data=buffer.getvalue(),
            file_name=f"ausencias_{date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with col_res:
        resumen = df_ausentes.groupby("SUCURSAL")["_ESTADO"].value_counts().unstack(fill_value=0)
        for col_name in ["VENCIDO","PENDIENTE","RESUELTO"]:
            if col_name not in resumen.columns:
                resumen[col_name] = 0
        resumen = resumen[["VENCIDO","PENDIENTE","RESUELTO"]].reset_index()
        resumen.columns = ["Sucursal","🚨 Vencidos","⏳ Pendientes","✅ Tipificados"]
        st.dataframe(resumen, use_container_width=True, hide_index=True)
