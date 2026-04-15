import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import json
import os
import io

st.set_page_config(page_title="Gestión de Ausencias", layout="wide", page_icon="📋")

st.markdown("""
<style>
html, body, [class*="css"] { font-family: 'Segoe UI', sans-serif; }
.stApp { background-color: #f0f2f5; }
section[data-testid="stSidebar"] { background-color: #0f1f3d; }
section[data-testid="stSidebar"] * { color: white !important; }
.titulo { font-size: 36px; font-weight: 800; color: #0f1f3d; text-align: center; margin-bottom: 5px; }
.subtitulo { font-size: 18px; color: #555; text-align: center; margin-bottom: 20px; }
.card { background: white; padding: 20px; border-radius: 12px; box-shadow: 0 3px 12px rgba(0,0,0,0.07); margin-bottom: 12px; border-left: 4px solid #e5e7eb; }
.card.vencido { border-left-color: #dc2626; }
.card.pendiente { border-left-color: #f59e0b; }
.card.resuelto { border-left-color: #10b981; }
.badge-rojo { background:#fee2e2; color:#dc2626; padding:3px 10px; border-radius:20px; font-size:12px; font-weight:700; }
.badge-amarillo { background:#fef3c7; color:#d97706; padding:3px 10px; border-radius:20px; font-size:12px; font-weight:700; }
.badge-verde { background:#d1fae5; color:#059669; padding:3px 10px; border-radius:20px; font-size:12px; font-weight:700; }
div.stButton > button { background-color:#0f1f3d; color:white; border-radius:8px; padding:8px 24px; border:none; font-weight:600; width:100%; }
div.stButton > button:hover { background-color:#1e3a6e; }
.metric-box { background:white; border-radius:10px; padding:16px; text-align:center; box-shadow:0 2px 8px rgba(0,0,0,0.06); }
.metric-num { font-size:32px; font-weight:800; }
.metric-label { font-size:13px; color:#666; margin-top:4px; }
</style>
""", unsafe_allow_html=True)

RUTA_XLSX = "reporte_ausencias.xlsx"
RUTA_GESTIONES = "gestiones.json"
PASSWORD_RRHH = "rrhh2026"
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


def calcular_estado_plazo(fecha_str):
    try:
        fecha = datetime.strptime(str(fecha_str).strip(), "%d/%m/%Y").date()
    except Exception:
        return "PENDIENTE", 99
    limite = fecha + timedelta(days=2)
    dias = (limite - date.today()).days
    return ("VENCIDO" if date.today() > limite else "PENDIENTE"), dias


def clave_registro(row):
    return f"{str(row['FECHA']).strip()}_{row['LEGAJO']}_{str(row['SUCURSAL']).strip()}"


def cargar_datos():
    if not os.path.exists(RUTA_XLSX):
        st.error(f"No se encontró **{RUTA_XLSX}**. Colocalo en la misma carpeta que esta app.")
        st.stop()
    sheets = pd.read_excel(RUTA_XLSX, sheet_name=None)
    frames = []
    for _, df in sheets.items():
        df.columns = [str(c).strip() for c in df.columns]
        df = df.fillna("")
        frames.append(df)
    return pd.concat(frames, ignore_index=True)


def cargar_gestiones():
    if os.path.exists(RUTA_GESTIONES):
        with open(RUTA_GESTIONES, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def guardar_gestiones(g):
    with open(RUTA_GESTIONES, "w", encoding="utf-8") as f:
        json.dump(g, f, ensure_ascii=False, indent=2)


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
        "Observaciones / Detalle",
        value=gestion_actual.get("observaciones", ""),
        placeholder="Ej: Avisó por WhatsApp, presentó certificado el...",
        key=f"obs_{key}",
        height=75,
    )
    responsable = st.text_input(
        "Responsable que carga *",
        value=gestion_actual.get("responsable", ""),
        placeholder="Nombre de quien tipifica",
        key=f"resp_{key}",
    )

    col_btn, _ = st.columns([1, 3])
    with col_btn:
        if st.button("Guardar", key=f"btn_{key}"):
            if not responsable.strip():
                st.error("El campo Responsable es obligatorio.")
            else:
                gestiones_dict[key] = {
                    "tipificacion": tipificacion,
                    "naaloo": naaloo,
                    "observaciones": observaciones.strip(),
                    "responsable": responsable.strip(),
                    "fecha_gestion": datetime.now().strftime("%d/%m/%Y %H:%M"),
                    "sucursal_origen": "RRHH" if es_rrhh else str(row["SUCURSAL"]),
                }
                guardar_gestiones(gestiones_dict)
                st.success("Guardado correctamente")
                st.rerun()

    if gestion_actual.get("fecha_gestion"):
        st.caption(
            f"Última actualización: {gestion_actual['fecha_gestion']} "
            f"— {gestion_actual.get('responsable','')} "
            f"({gestion_actual.get('sucursal_origen','')})"
        )


# ====== CARGA ======
df_global = cargar_datos()
gestiones = cargar_gestiones()
sucursales_disponibles = sorted(df_global["SUCURSAL"].dropna().unique().tolist())

# ====== SIDEBAR ======
with st.sidebar:
    st.markdown("## Acceso")
    modo = st.radio("Tipo de acceso", ["Sucursal", "RRHH (Global)"])
    if modo == "Sucursal":
        sucursal_sel = st.selectbox("Sucursal", sucursales_disponibles)
        password = st.text_input("Contraseña", type="password")
        es_rrhh = False
        slug = (sucursal_sel.lower()
                .replace(' ', '').replace('.', '')
                .replace('ü','u').replace('é','e').replace('ó','o')
                .replace('á','a').replace('í','i'))
        autenticado = password == f"{slug}{date.today().year}"
    else:
        sucursal_sel = None
        password = st.text_input("Contraseña RRHH", type="password")
        es_rrhh = True
        autenticado = password == PASSWORD_RRHH

    if password and not autenticado:
        st.error("Contraseña incorrecta")
    st.markdown("---")
    st.caption(f"Hoy: {date.today().strftime('%d/%m/%Y')}")

# ====== HEADER ======
st.markdown('<div class="titulo">GESTIÓN DE AUSENCIAS</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitulo">Tipificación y seguimiento — 48hs para resolución</div>', unsafe_allow_html=True)

if not password:
    st.info("Seleccioná tu sucursal e ingresá la contraseña para continuar.")
    st.stop()
if not autenticado:
    st.stop()

# ====== DATOS ======
df_vista = df_global.copy() if es_rrhh else df_global[df_global["SUCURSAL"] == sucursal_sel].copy()
df_ausentes = df_vista[df_vista["SITUACION"].str.contains("AUSENTE", case=False, na=False)].copy()
df_ausentes["_ESTADO"] = df_ausentes.apply(lambda r: get_estado_fila(r, gestiones), axis=1)

titulo_vista = "Vista Global — Todas las Sucursales" if es_rrhh else f"Sucursal: {sucursal_sel}"
st.markdown(f"### {titulo_vista}")

# ====== METRICAS ======
total = len(df_ausentes)
resueltos = int((df_ausentes["_ESTADO"] == "RESUELTO").sum())
vencidos = int((df_ausentes["_ESTADO"] == "VENCIDO").sum())
pendientes = int((df_ausentes["_ESTADO"] == "PENDIENTE").sum())

c1, c2, c3, c4 = st.columns(4)
c1.markdown(f'<div class="metric-box"><div class="metric-num" style="color:#0f1f3d">{total}</div><div class="metric-label">Total</div></div>', unsafe_allow_html=True)
c2.markdown(f'<div class="metric-box"><div class="metric-num" style="color:#f59e0b">{pendientes}</div><div class="metric-label">Pendientes</div></div>', unsafe_allow_html=True)
c3.markdown(f'<div class="metric-box"><div class="metric-num" style="color:#dc2626">{vencidos}</div><div class="metric-label">Vencidos +48hs</div></div>', unsafe_allow_html=True)
c4.markdown(f'<div class="metric-box"><div class="metric-num" style="color:#10b981">{resueltos}</div><div class="metric-label">Tipificados</div></div>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ====== FILTROS ======
cf1, cf2, cf3 = st.columns([2, 2, 2])
with cf1:
    buscar = st.text_input("Buscar por nombre o legajo", placeholder="Ej: GARCIA o 5402")
with cf2:
    filtro_estado = st.selectbox("Estado", ["Todos", "Pendientes", "Vencidos (+48hs)", "Tipificados"])
with cf3:
    if es_rrhh:
        filtro_suc = st.selectbox("Filtrar sucursal", ["Todas"] + sucursales_disponibles)
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

# ====== REGISTROS ======
if df_filtrado.empty:
    st.info("No hay registros con los filtros aplicados.")
else:
    for _, row in df_filtrado.iterrows():
        key = clave_registro(row)
        est_actual, gestion_actual = estado_gestion(key, gestiones)
        plazo_str, dias = calcular_estado_plazo(str(row["FECHA"]))

        if est_actual == "RESUELTO":
            css_class = "resuelto"
            badge = f'<span class="badge-verde">✅ {gestion_actual.get("tipificacion","")}</span>'
        elif plazo_str == "VENCIDO":
            css_class = "vencido"
            badge = '<span class="badge-rojo">🚨 VENCIDO (+48hs)</span>'
        else:
            css_class = "pendiente"
            dias_txt = f"{dias} día(s)" if dias > 0 else "vence hoy"
            badge = f'<span class="badge-amarillo">⏳ PENDIENTE — {dias_txt}</span>'

        lic_str = str(row.get("LICENCIA", "")).strip()
        lic_html = f'&nbsp;<span style="color:#888; font-size:12px;">Naaloo previo: {lic_str}</span>' if lic_str and lic_str != "nan" else ""
        suc_html = f'&nbsp;<span style="color:#888; font-size:12px;">📍 {row["SUCURSAL"]}</span>' if es_rrhh else ""

        st.markdown(f"""
        <div class="card {css_class}">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
                <div>
                    <strong style="font-size:15px;">👤 {str(row['NOMBRE']).strip()}</strong>
                    &nbsp;<span style="color:#666; font-size:13px;">Leg. {row['LEGAJO']}</span>
                    {suc_html}
                </div>
                <div>{badge}</div>
            </div>
            <div style="display:flex; gap:16px; font-size:13px; color:#555; flex-wrap:wrap;">
                <span>📅 <strong>{row['FECHA']}</strong></span>
                <span>M: <strong>{row['MAÑANA']}</strong></span>
                <span>T: <strong>{row['TARDE']}</strong></span>
                {lic_html}
            </div>
        </div>
        """, unsafe_allow_html=True)

        label_exp = "Ver / Editar tipificación" if est_actual == "RESUELTO" else "➕ Tipificar ausencia"
        with st.expander(label_exp):
            mostrar_formulario(key, gestion_actual, gestiones, row, es_rrhh)

# ====== EXPORTAR (solo RRHH) ======
if es_rrhh:
    st.markdown("---")
    st.markdown("### Exportar reporte completo")
    col_exp, col_res = st.columns(2)

    rows_export = []
    for _, row in df_ausentes.iterrows():
        key = clave_registro(row)
        est, g = estado_gestion(key, gestiones)
        plazo, _ = calcular_estado_plazo(str(row["FECHA"]))
        rows_export.append({
            "FECHA": row["FECHA"],
            "LEGAJO": row["LEGAJO"],
            "NOMBRE": row["NOMBRE"],
            "SUCURSAL": row["SUCURSAL"],
            "MAÑANA": row["MAÑANA"],
            "TARDE": row["TARDE"],
            "NAALOO_PREVIO": row.get("LICENCIA", ""),
            "ESTADO": est if est == "RESUELTO" else plazo,
            "TIPIFICACION": g.get("tipificacion", ""),
            "NAALOO_CARGADO": g.get("naaloo", ""),
            "OBSERVACIONES": g.get("observaciones", ""),
            "RESPONSABLE": g.get("responsable", ""),
            "FECHA_GESTION": g.get("fecha_gestion", ""),
        })
    df_export = pd.DataFrame(rows_export)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_export.to_excel(writer, index=False, sheet_name="Ausencias")

    with col_exp:
        st.download_button(
            "⬇️ Descargar Excel completo",
            data=buffer.getvalue(),
            file_name=f"ausencias_{date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with col_res:
        if not df_ausentes.empty:
            resumen = df_ausentes.groupby("SUCURSAL")["_ESTADO"].value_counts().unstack(fill_value=0)
            for col_name in ["VENCIDO", "PENDIENTE", "RESUELTO"]:
                if col_name not in resumen.columns:
                    resumen[col_name] = 0
            resumen = resumen[["VENCIDO", "PENDIENTE", "RESUELTO"]].reset_index()
            resumen.columns = ["Sucursal", "Vencidos", "Pendientes", "Tipificados"]
            st.dataframe(resumen, use_container_width=True, hide_index=True)
