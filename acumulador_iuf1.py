import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="Cartera por Edades IUF1 · DISPOWER", page_icon="⚡", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@500;700&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.stApp { background: #F0F4F8; }

.header-block {
    background: linear-gradient(135deg, #0D1B2A 0%, #1B3A5C 100%);
    border-radius: 12px; padding: 26px 32px; margin-bottom: 20px;
    display: flex; align-items: center; gap: 16px;
}
.header-title { color: #FFF; font-size: 21px; font-weight: 700; margin: 0; }
.header-sub   { color: #7DB4D8; font-size: 12px; margin: 4px 0 0; }

.section-title {
    font-size: 13px; font-weight: 700; color: #0D1B2A;
    border-bottom: 2px solid #D1D5DB; padding-bottom: 7px; margin: 22px 0 12px;
    text-transform: uppercase; letter-spacing: .04em;
}

.kpi-card {
    background: #FFF; border-radius: 10px; padding: 18px 22px;
    border-left: 4px solid var(--c); box-shadow: 0 1px 4px rgba(0,0,0,.07);
}
.kpi-label { font-size: 10px; font-weight: 700; color: #6B7280; text-transform: uppercase; letter-spacing: .07em; }
.kpi-value { font-size: 24px; font-weight: 700; color: #0D1B2A; font-family: 'JetBrains Mono', monospace; margin-top: 3px; }
.kpi-sub   { font-size: 11px; color: #9CA3AF; margin-top: 2px; }

.badge-mes {
    display:inline-block; background:#E8F4FD; color:#1B6CA8;
    border-radius:6px; padding:3px 10px; font-size:12px; font-weight:600; margin:3px;
}
.warn { background:#FFF8E1; border-left:4px solid #F59E0B; border-radius:6px;
        padding:11px 15px; font-size:13px; color:#92400E; margin-bottom:8px; }

/* rangos */
.r0  { color: #16A34A; font-weight:700; }
.r90 { color: #D97706; font-weight:700; }
.r360{ color: #DC2626; font-weight:700; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="header-block">
  <span style="font-size:34px">⚡</span>
  <div>
    <p class="header-title">Cartera por Edades · IUF1</p>
    <p class="header-sub">DISPOWER S.A.S. E.S.P. · ZNI · Antigüedad de cartera por NIU (cada mes = 30 días)</p>
  </div>
</div>
""", unsafe_allow_html=True)

# ── Estado ─────────────────────────────────────────────────────────────────────
for k, v in [("df_acum", pd.DataFrame()), ("periodos_cargados", []), ("advertencias", [])]:
    if k not in st.session_state:
        st.session_state[k] = v

MESES_ES = {1:"ENE",2:"FEB",3:"MAR",4:"ABR",5:"MAY",6:"JUN",
            7:"JUL",8:"AGO",9:"SEP",10:"OCT",11:"NOV",12:"DIC"}
MESES_NUM = {v:k for k,v in MESES_ES.items()}

def mes_a_fecha(mes_anio: str):
    """'OCT-2023' → datetime(2023,10,1)"""
    try:
        m, a = mes_anio.split("-")
        return datetime(int(a), MESES_NUM[m], 1)
    except:
        return None

def extraer_mes_anio(fecha_str):
    try:
        p = str(fecha_str).strip().split("-")
        return f"{MESES_ES.get(int(p[1]), p[1])}-{p[2]}"
    except:
        return str(fecha_str)

def rango(dias):
    if dias <= 90:
        return "0 – 90 días"
    elif dias <= 360:
        return "91 – 360 días"
    else:
        return "> 360 días"

ORDEN_RANGO = {"0 – 90 días": 0, "91 – 360 días": 1, "> 360 días": 2}
COLOR_RANGO = {"0 – 90 días": "#16A34A", "91 – 360 días": "#D97706", "> 360 días": "#DC2626"}

# ══════════════════════════════════════════════════════════════════════════════
# PASO 1 · CARGA
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<p class="section-title">① Cargar archivos IUF1</p>', unsafe_allow_html=True)

archivos = st.file_uploader(
    "Selecciona uno o varios Excel IUF1 (uno por mes)",
    type=["xlsx","xls"], accept_multiple_files=True,
    help="El programa detecta el mes/año desde FECH INI PERIO."
)

col_a, col_b = st.columns([1,5])
with col_a:
    procesar = st.button("▶ Cargar", type="primary", use_container_width=True)
with col_b:
    if st.button("🗑 Limpiar todo"):
        st.session_state.df_acum = pd.DataFrame()
        st.session_state.periodos_cargados = []
        st.session_state.advertencias = []
        st.rerun()

COLS_REQ = ["NIU","COD LOCALIDAD","FECH INI PERIO","TARIFA"]

if procesar and archivos:
    adv, nuevos = [], []
    for archivo in archivos:
        nombre = archivo.name
        try:
            df = pd.read_excel(archivo, sheet_name=0, dtype=str)
            df.columns = df.columns.str.strip().str.upper()
            falt = [c for c in COLS_REQ if c not in df.columns]
            if falt:
                adv.append(f"⚠️ **{nombre}**: faltan → {', '.join(falt)}"); continue

            df["NIU"]           = df["NIU"].str.strip()
            df["COD LOCALIDAD"] = df["COD LOCALIDAD"].str.strip()
            df["MES_ANIO"]      = df["FECH INI PERIO"].apply(extraer_mes_anio)
            df["TARIFA"]        = pd.to_numeric(
                df["TARIFA"].str.replace(",","."), errors="coerce")
            nulos = df["TARIFA"].isna().sum()
            if nulos: adv.append(f"ℹ️ **{nombre}**: {nulos} filas con TARIFA no numérica omitidas.")
            df = df.dropna(subset=["NIU","TARIFA"])
            df["_ARCHIVO"] = nombre
            nuevos.append(df[["NIU","COD LOCALIDAD","MES_ANIO","TARIFA","_ARCHIVO"]])
            if nombre not in st.session_state.periodos_cargados:
                st.session_state.periodos_cargados.append(nombre)
        except Exception as e:
            adv.append(f"❌ **{nombre}**: {e}")

    if nuevos:
        nuevo = pd.concat(nuevos, ignore_index=True)
        if not st.session_state.df_acum.empty:
            ya = st.session_state.df_acum["_ARCHIVO"].unique()
            nuevo = nuevo[~nuevo["_ARCHIVO"].isin(ya)]
        if not nuevo.empty:
            st.session_state.df_acum = pd.concat(
                [st.session_state.df_acum, nuevo], ignore_index=True)
        else:
            adv.append("ℹ️ Todos los archivos ya estaban cargados.")
    st.session_state.advertencias = adv
    st.rerun()

for a in st.session_state.advertencias:
    st.markdown(f'<div class="warn">{a}</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# PASO 2 · SELECCIÓN DE MESES
# ══════════════════════════════════════════════════════════════════════════════
if not st.session_state.df_acum.empty:
    df_base = st.session_state.df_acum.copy()

    # Ordenar meses cronológicamente
    meses_disponibles = sorted(
        df_base["MES_ANIO"].unique(),
        key=lambda x: mes_a_fecha(x) or datetime.min
    )

    st.markdown('<p class="section-title">② Seleccionar meses para la cartera</p>', unsafe_allow_html=True)

    col_sel1, col_sel2 = st.columns([3,1])
    with col_sel1:
        meses_sel = st.multiselect(
            "Meses cargados (selecciona los que quieres incluir en el análisis)",
            options=meses_disponibles,
            default=meses_disponibles,
            help="El mes más reciente seleccionado es la fecha de corte para calcular antigüedad."
        )
    with col_sel2:
        st.markdown("<div style='margin-top:28px'>", unsafe_allow_html=True)
        if st.button("✅ Todos", use_container_width=True):
            meses_sel = meses_disponibles
        st.markdown("</div>", unsafe_allow_html=True)

    if not meses_sel:
        st.info("Selecciona al menos un mes para continuar.")
        st.stop()

    # Mes de corte = más reciente entre los seleccionados
    meses_sel_ord = sorted(meses_sel, key=lambda x: mes_a_fecha(x) or datetime.min)
    mes_corte = meses_sel_ord[-1]
    fecha_corte = mes_a_fecha(mes_corte)

    st.markdown(
        f"**Fecha de corte:** `{mes_corte}` &nbsp;|&nbsp; "
        f"**Meses seleccionados:** {len(meses_sel)} &nbsp;|&nbsp; "
        f"**Rango:** `{meses_sel_ord[0]}` → `{mes_corte}`",
        unsafe_allow_html=True
    )

    # ══════════════════════════════════════════════════════════════════════════
    # PASO 3 · CÁLCULO
    # ══════════════════════════════════════════════════════════════════════════
    df_fil = df_base[df_base["MES_ANIO"].isin(meses_sel)].copy()

    # Días de antigüedad: (fecha_corte − fecha_mes) en meses × 30
    def calc_dias(mes_anio):
        f = mes_a_fecha(mes_anio)
        if f is None or fecha_corte is None:
            return 0
        meses_diff = (fecha_corte.year - f.year) * 12 + (fecha_corte.month - f.month)
        return meses_diff * 30

    df_fil["DIAS_ANTIG"] = df_fil["MES_ANIO"].apply(calc_dias)
    df_fil["RANGO"]      = df_fil["DIAS_ANTIG"].apply(rango)

    # Suma de tarifa por NIU + COD LOCALIDAD + MES_ANIO
    resumen = (
        df_fil
        .groupby(["NIU","COD LOCALIDAD","MES_ANIO","DIAS_ANTIG","RANGO"])["TARIFA"]
        .sum()
        .reset_index()
        .rename(columns={"TARIFA":"TARIFA_ACUM"})
    )
    resumen["_ORDEN_RANGO"] = resumen["RANGO"].map(ORDEN_RANGO)
    resumen = resumen.sort_values(["NIU","_ORDEN_RANGO","MES_ANIO"]).drop(columns="_ORDEN_RANGO")
    resumen = resumen.reset_index(drop=True)

    # ── KPIs ──────────────────────────────────────────────────────────────────
    st.markdown('<p class="section-title">③ Resumen de cartera</p>', unsafe_allow_html=True)

    tot_gral = resumen["TARIFA_ACUM"].sum()
    for rng, color in COLOR_RANGO.items():
        sub = resumen[resumen["RANGO"]==rng]
        val = sub["TARIFA_ACUM"].sum()
        pct = (val/tot_gral*100) if tot_gral else 0

    cols_kpi = st.columns(4)
    kpi_data = [
        ("Total cartera",    f"${tot_gral:,.0f}",  "#1B6CA8", "todos los meses seleccionados"),
    ] + [
        (rng, f"${resumen[resumen['RANGO']==rng]['TARIFA_ACUM'].sum():,.0f}", color,
         f"{resumen[resumen['RANGO']==rng]['TARIFA_ACUM'].sum()/tot_gral*100:.1f}% del total" if tot_gral else "0%")
        for rng, color in COLOR_RANGO.items()
    ]
    for col, (lbl, val, color, sub) in zip(cols_kpi, kpi_data):
        with col:
            st.markdown(f"""
            <div class="kpi-card" style="--c:{color}">
                <div class="kpi-label">{lbl}</div>
                <div class="kpi-value">{val}</div>
                <div class="kpi-sub">{sub}</div>
            </div>""", unsafe_allow_html=True)

    st.markdown("<div style='margin-top:10px'></div>", unsafe_allow_html=True)

    # ── Tabla ─────────────────────────────────────────────────────────────────
    st.markdown('<p class="section-title">④ Detalle por NIU</p>', unsafe_allow_html=True)

    f1, f2, f3 = st.columns(3)
    with f1:
        f_niu = st.text_input("Filtrar NIU", placeholder="Ej. 444301441")
    with f2:
        f_loc = st.text_input("Filtrar COD LOCALIDAD", placeholder="Ej. 4443000000054")
    with f3:
        f_rng = st.selectbox("Filtrar rango antigüedad", ["Todos"] + list(COLOR_RANGO.keys()))

    tabla = resumen.copy()
    if f_niu.strip():
        tabla = tabla[tabla["NIU"].str.contains(f_niu.strip(), na=False)]
    if f_loc.strip():
        tabla = tabla[tabla["COD LOCALIDAD"].str.contains(f_loc.strip(), na=False)]
    if f_rng != "Todos":
        tabla = tabla[tabla["RANGO"] == f_rng]

    tabla_display = tabla[["NIU","COD LOCALIDAD","MES_ANIO","DIAS_ANTIG","RANGO","TARIFA_ACUM"]].copy()
    tabla_display.columns = ["NIU","COD LOCALIDAD","MES-AÑO","DÍAS ANTIG.","RANGO","TARIFA ($)"]
    tabla_display["TARIFA ($)"] = tabla_display["TARIFA ($)"].apply(lambda x: f"${x:,.2f}")
    tabla_display.index = range(1, len(tabla_display)+1)

    st.dataframe(tabla_display, use_container_width=True, height=420)
    st.caption(f"{len(tabla_display):,} registros · {tabla['NIU'].nunique():,} NIUs únicos")

    # ── Exportar ──────────────────────────────────────────────────────────────
    st.markdown('<p class="section-title">⑤ Exportar Excel</p>', unsafe_allow_html=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        # Hoja 1: detalle completo
        exp1 = resumen[["NIU","COD LOCALIDAD","MES_ANIO","DIAS_ANTIG","RANGO","TARIFA_ACUM"]].copy()
        exp1.columns = ["NIU","COD LOCALIDAD","MES-AÑO","DÍAS ANTIGÜEDAD","RANGO CARTERA","TARIFA"]
        exp1.to_excel(writer, sheet_name="Detalle_Cartera", index=False)

        # Hoja 2: resumen por rango (NIU × RANGO → suma tarifa)
        pivot_rango = resumen.pivot_table(
            index=["NIU","COD LOCALIDAD"],
            columns="RANGO", values="TARIFA_ACUM",
            aggfunc="sum", fill_value=0
        ).reset_index()
        for r in COLOR_RANGO:
            if r not in pivot_rango.columns:
                pivot_rango[r] = 0
        pivot_rango = pivot_rango[["NIU","COD LOCALIDAD"] + list(COLOR_RANGO.keys())]
        pivot_rango["TOTAL"] = pivot_rango[list(COLOR_RANGO.keys())].sum(axis=1)
        pivot_rango.to_excel(writer, sheet_name="Resumen_x_Rango", index=False)

        # Hoja 3: totales globales por rango
        tot = resumen.groupby("RANGO")["TARIFA_ACUM"].sum().reset_index()
        tot.columns = ["RANGO CARTERA","TARIFA TOTAL"]
        tot["% DEL TOTAL"] = (tot["TARIFA TOTAL"] / tot["TARIFA TOTAL"].sum() * 100).round(2)
        tot.to_excel(writer, sheet_name="Totales_x_Rango", index=False)

    label_meses = f"{meses_sel_ord[0]}_a_{mes_corte}".replace(" ","")
    nombre_xlsx = f"cartera_edades_IUF1_{label_meses}.xlsx"

    st.download_button(
        label="⬇️ Descargar Excel — Cartera por Edades",
        data=output.getvalue(),
        file_name=nombre_xlsx,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
    st.caption(
        "**Hoja 1** Detalle NIU × Mes × Rango · "
        "**Hoja 2** Pivot NIU con columna por rango + total · "
        "**Hoja 3** Totales globales por rango"
    )

else:
    st.info("Carga uno o varios archivos IUF1 con el botón de arriba para comenzar.")
