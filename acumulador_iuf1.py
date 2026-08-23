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
.warn { background:#FFF8E1; border-left:4px solid #F59E0B; border-radius:6px;
        padding:11px 15px; font-size:13px; color:#92400E; margin-bottom:8px; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="header-block">
  <span style="font-size:34px">⚡</span>
  <div>
    <p class="header-title">Cartera por Edades · IUF1</p>
    <p class="header-sub">DISPOWER S.A.S. E.S.P. · ZNI · Antigüedad de cartera por NIU y Localidad (cada mes = 30 días)</p>
  </div>
</div>
""", unsafe_allow_html=True)

# ── Constantes ─────────────────────────────────────────────────────────────────
MESES_ES  = {1:"ENE",2:"FEB",3:"MAR",4:"ABR",5:"MAY",6:"JUN",
             7:"JUL",8:"AGO",9:"SEP",10:"OCT",11:"NOV",12:"DIC"}
MESES_NUM = {v:k for k,v in MESES_ES.items()}
RANGOS       = ["0 – 90 días", "91 – 360 días", "> 360 días"]
COLOR_RANGO  = {"0 – 90 días":"#16A34A", "91 – 360 días":"#D97706", "> 360 días":"#DC2626"}
ORDEN_RANGO  = {r:i for i,r in enumerate(RANGOS)}
COLS_LOCALIDAD = ["COD LOCALIDAD", "COD_LOCALIDAD"]
COLS_FECHA     = ["FECH INI PERIO", "FEC_INICIO_PERIODO"]
COLS_TARIFA    = ["TARIFA", "VALOR_TARIFA"]

# ── Helpers ────────────────────────────────────────────────────────────────────
def mes_a_fecha(mes_anio):
    try:
        m, a = str(mes_anio).strip().split("-")
        return datetime(int(a), MESES_NUM[m], 1)
    except Exception:
        return datetime.min

def extraer_mes_anio(fecha_str):
    """Parsea cualquier formato de fecha y retorna 'MMM-YYYY', ej: 'JUL-2023'."""
    import pandas as pd
    # Caso 1: datetime o Timestamp
    if isinstance(fecha_str, (datetime, pd.Timestamp)):
        return f"{MESES_ES[fecha_str.month]}-{fecha_str.year}"
    s = str(fecha_str).strip()
    # Caso 2: formatos ISO / con hora "2023-07-03" o "2023-07-03 00:00:00"
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            dt = datetime.strptime(s[:len(fmt)], fmt)
            return f"{MESES_ES[dt.month]}-{dt.year}"
        except Exception:
            pass
    # Caso 3: ya viene como "OCT-2023"
    try:
        partes = s.split("-")
        if partes[0] in MESES_NUM:
            return s
    except Exception:
        pass
    # Caso 4: "DD-MM-YYYY"
    try:
        p = s.split("-")
        return f"{MESES_ES.get(int(p[1]), p[1])}-{p[2]}"
    except Exception:
        return s

def clasificar_rango(dias):
    if dias <= 0:      return None           # mes de corte o futuro
    elif dias <= 120:  return "0 – 90 días"
    elif dias <= 360:  return "91 – 360 días"
    else:              return "> 360 días"

def dias_entre(mes_anio, fecha_corte):
    f = mes_a_fecha(mes_anio)
    return ((fecha_corte.year - f.year) * 12 + (fecha_corte.month - f.month)) * 30

# ── Estado de sesión ───────────────────────────────────────────────────────────
if "df_acum" not in st.session_state:
    st.session_state.df_acum = pd.DataFrame()
if "periodos_cargados" not in st.session_state:
    st.session_state.periodos_cargados = []
if "advertencias" not in st.session_state:
    st.session_state.advertencias = []

# ══════════════════════════════════════════════════════════════════════════════
# PASO 1 · CARGA
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<p class="section-title">① Cargar archivos IUF1</p>', unsafe_allow_html=True)

archivos = st.file_uploader(
    "Selecciona uno o varios Excel IUF1 (uno por mes)",
    type=["xlsx", "xls"], accept_multiple_files=True
)

col_a, col_b = st.columns([1, 5])
with col_a:
    procesar = st.button("▶ Cargar", type="primary", use_container_width=True)
with col_b:
    if st.button("🗑 Limpiar todo"):
        st.session_state.df_acum = pd.DataFrame()
        st.session_state.periodos_cargados = []
        st.session_state.advertencias = []
        st.rerun()

if procesar and archivos:
    adv, nuevos = [], []
    for archivo in archivos:
        nombre = archivo.name
        try:
            df = pd.read_excel(archivo, sheet_name=0, dtype=str)
            df.columns = df.columns.str.strip().str.upper()

            if "NIU" not in df.columns:
                adv.append(f"⚠️ **{nombre}**: no se encontró columna NIU."); continue

            col_loc = next((c for c in COLS_LOCALIDAD if c in df.columns), None)
            if col_loc is None:
                adv.append(f"⚠️ **{nombre}**: no se encontró COD LOCALIDAD ni COD_LOCALIDAD."); continue

            col_fecha = next((c for c in COLS_FECHA if c in df.columns), None)
            if col_fecha is None:
                adv.append(f"⚠️ **{nombre}**: no se encontró FECH INI PERIO ni FEC_INICIO_PERIODO."); continue

            col_tarifa = next((c for c in COLS_TARIFA if c in df.columns), None)
            if col_tarifa is None:
                adv.append(f"⚠️ **{nombre}**: no se encontró TARIFA ni VALOR_TARIFA."); continue

            df["NIU"]           = df["NIU"].str.strip()
            df["COD LOCALIDAD"] = df[col_loc].str.strip()
            df["MES_ANIO"]      = df[col_fecha].apply(extraer_mes_anio)
            df["TARIFA"]        = pd.to_numeric(df[col_tarifa].str.replace(",", "."), errors="coerce")

            nulos = df["TARIFA"].isna().sum()
            if nulos:
                adv.append(f"ℹ️ **{nombre}**: {nulos} filas con TARIFA no numérica omitidas.")
            df = df.dropna(subset=["NIU", "TARIFA"])
            df["_ARCHIVO"] = nombre
            nuevos.append(df[["NIU", "COD LOCALIDAD", "MES_ANIO", "TARIFA", "_ARCHIVO"]])
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
# PASO 2 · CONFIGURAR PERÍODO
# ══════════════════════════════════════════════════════════════════════════════
if not st.session_state.df_acum.empty:
    df_base = st.session_state.df_acum.copy()

    meses_disponibles = sorted(
        df_base["MES_ANIO"].unique(),
        key=mes_a_fecha
    )

    st.markdown('<p class="section-title">② Configurar período de cartera</p>', unsafe_allow_html=True)

    col_c1, col_c2 = st.columns([1, 2])

    with col_c1:
        st.markdown("**Mes de corte** *(referencia — no entra al cálculo)*")
        mes_corte = st.selectbox(
            "Mes de corte",
            options=meses_disponibles,
            index=len(meses_disponibles) - 1,
            label_visibility="collapsed",
            help="La antigüedad de cada mes se mide hacia atrás desde este mes."
        )

    fecha_corte = mes_a_fecha(mes_corte)

    with col_c2:
        st.markdown("**Meses de cartera** *(meses cuya tarifa se acumula)*")
        opciones_cartera = [m for m in meses_disponibles if m != mes_corte]
        default_cartera  = [m for m in opciones_cartera if mes_a_fecha(m) < fecha_corte]
        meses_sel = st.multiselect(
            "Meses de cartera",
            options=opciones_cartera,
            default=default_cartera,
            label_visibility="collapsed",
            help="Cada mes se clasifica según su distancia al mes de corte."
        )

    if not meses_sel:
        st.info("Selecciona al menos un mes de cartera para continuar.")
        st.stop()

    meses_sel_ord = sorted(meses_sel, key=mes_a_fecha)

    with st.expander("👁 Ver antigüedad por mes seleccionado"):
        prev = []
        for m in meses_sel_ord:
            d = dias_entre(m, fecha_corte)
            prev.append({"Mes": m, "Días antigüedad": d, "Rango": clasificar_rango(d) or "—"})
        st.dataframe(pd.DataFrame(prev), use_container_width=True, hide_index=True)

    st.markdown(
        f"**Corte:** `{mes_corte}` &nbsp;|&nbsp; "
        f"**Meses en cartera:** {len(meses_sel)} &nbsp;|&nbsp; "
        f"**Período:** `{meses_sel_ord[0]}` → `{meses_sel_ord[-1]}`",
        unsafe_allow_html=True
    )

    # ══════════════════════════════════════════════════════════════════════════
    # CÁLCULO
    # ══════════════════════════════════════════════════════════════════════════
    df_fil = df_base[df_base["MES_ANIO"].isin(meses_sel)].copy()
    df_fil["DIAS_ANTIG"] = df_fil["MES_ANIO"].apply(lambda m: dias_entre(m, fecha_corte))
    df_fil["RANGO"]      = df_fil["DIAS_ANTIG"].apply(clasificar_rango)
    df_fil = df_fil[df_fil["RANGO"].notna()].copy()

    # Detalle NIU × mes
    det_niu = (
        df_fil
        .groupby(["NIU", "COD LOCALIDAD", "MES_ANIO", "DIAS_ANTIG", "RANGO"])["TARIFA"]
        .sum().reset_index().rename(columns={"TARIFA": "TARIFA_ACUM"})
    )
    det_niu["_ORD"] = det_niu["RANGO"].map(ORDEN_RANGO)
    det_niu = det_niu.sort_values(["NIU", "_ORD", "MES_ANIO"]).drop(columns="_ORD").reset_index(drop=True)

    # Cartera total por NIU
    cartera_niu = (
        df_fil.groupby(["NIU", "COD LOCALIDAD"])["TARIFA"]
        .sum().reset_index().rename(columns={"TARIFA": "TARIFA_TOTAL"})
        .sort_values("TARIFA_TOTAL", ascending=False).reset_index(drop=True)
    )

    # Cartera por localidad × rango
    pivot_loc = df_fil.pivot_table(
        index="COD LOCALIDAD", columns="RANGO",
        values="TARIFA", aggfunc="sum", fill_value=0
    ).reset_index()
    for r in RANGOS:
        if r not in pivot_loc.columns:
            pivot_loc[r] = 0
    pivot_loc = pivot_loc[["COD LOCALIDAD"] + RANGOS].copy()
    pivot_loc["TOTAL"] = pivot_loc[RANGOS].sum(axis=1)
    pivot_loc = pivot_loc.sort_values("TOTAL", ascending=False).reset_index(drop=True)

    # ══════════════════════════════════════════════════════════════════════════
    # PASO 3 · KPIs
    # ══════════════════════════════════════════════════════════════════════════
    st.markdown('<p class="section-title">③ Resumen de cartera</p>', unsafe_allow_html=True)

    tot_gral = det_niu["TARIFA_ACUM"].sum()
    kpi_data = [("Total cartera", f"${tot_gral:,.0f}", "#1B6CA8", f"{det_niu['NIU'].nunique():,} NIUs")] + [
        (rng,
         f"${det_niu[det_niu['RANGO'] == rng]['TARIFA_ACUM'].sum():,.0f}",
         color,
         f"{det_niu[det_niu['RANGO'] == rng]['TARIFA_ACUM'].sum() / tot_gral * 100:.1f}% del total" if tot_gral else "0%")
        for rng, color in COLOR_RANGO.items()
    ]
    for col, (lbl, val, color, sub) in zip(st.columns(4), kpi_data):
        with col:
            st.markdown(f"""
            <div class="kpi-card" style="--c:{color}">
                <div class="kpi-label">{lbl}</div>
                <div class="kpi-value">{val}</div>
                <div class="kpi-sub">{sub}</div>
            </div>""", unsafe_allow_html=True)

    st.markdown("<div style='margin-top:10px'></div>", unsafe_allow_html=True)

    # ══════════════════════════════════════════════════════════════════════════
    # PASO 4 · TABLAS
    # ══════════════════════════════════════════════════════════════════════════
    tab1, tab2, tab3 = st.tabs(["📋 Detalle por NIU", "🏘 Por Localidad", "💰 Cartera total NIU"])

    with tab1:
        st.markdown('<p class="section-title">Detalle por NIU × Mes × Rango</p>', unsafe_allow_html=True)
        f1, f2, f3 = st.columns(3)
        with f1: f_niu = st.text_input("Filtrar NIU", placeholder="Ej. 444301441")
        with f2: f_loc = st.text_input("Filtrar COD LOCALIDAD", placeholder="Ej. 4443000000054")
        with f3: f_rng = st.selectbox("Filtrar rango", ["Todos"] + RANGOS)

        t = det_niu.copy()
        if f_niu.strip(): t = t[t["NIU"].str.contains(f_niu.strip(), na=False)]
        if f_loc.strip(): t = t[t["COD LOCALIDAD"].str.contains(f_loc.strip(), na=False)]
        if f_rng != "Todos": t = t[t["RANGO"] == f_rng]

        td = t[["NIU", "COD LOCALIDAD", "MES_ANIO", "DIAS_ANTIG", "RANGO", "TARIFA_ACUM"]].copy()
        td.columns = ["NIU", "COD LOCALIDAD", "MES-AÑO", "DÍAS ANTIG.", "RANGO", "TARIFA ($)"]
        td["TARIFA ($)"] = td["TARIFA ($)"].apply(lambda x: f"${x:,.2f}")
        td.index = range(1, len(td) + 1)
        st.dataframe(td, use_container_width=True, height=420)
        st.caption(f"{len(td):,} registros · {t['NIU'].nunique():,} NIUs únicos")

    with tab2:
        st.markdown('<p class="section-title">Cartera por Localidad × Rango</p>', unsafe_allow_html=True)
        td_loc = pivot_loc.copy()
        for col in RANGOS + ["TOTAL"]:
            td_loc[col] = td_loc[col].apply(lambda x: f"${x:,.2f}")
        td_loc.index = range(1, len(td_loc) + 1)
        st.dataframe(td_loc, use_container_width=True, height=420)
        st.caption(f"{len(pivot_loc):,} localidades")

    with tab3:
        st.markdown('<p class="section-title">Cartera total por NIU</p>', unsafe_allow_html=True)
        f_niu2 = st.text_input("Filtrar NIU", placeholder="Ej. 444301441", key="f_niu2")
        t2 = cartera_niu.copy()
        if f_niu2.strip(): t2 = t2[t2["NIU"].str.contains(f_niu2.strip(), na=False)]
        td2 = t2.copy()
        td2["TARIFA_TOTAL"] = td2["TARIFA_TOTAL"].apply(lambda x: f"${x:,.2f}")
        td2.columns = ["NIU", "COD LOCALIDAD", "TARIFA TOTAL ($)"]
        td2.index = range(1, len(td2) + 1)
        st.dataframe(td2, use_container_width=True, height=420)
        st.caption(f"{len(td2):,} NIUs")

    # ══════════════════════════════════════════════════════════════════════════
    # PASO 5 · EXPORTAR
    # ══════════════════════════════════════════════════════════════════════════
    st.markdown('<p class="section-title">⑤ Exportar</p>', unsafe_allow_html=True)
    label_meses = f"{meses_sel_ord[0]}_a_{meses_sel_ord[-1]}_corte_{mes_corte}".replace(" ", "")

    col_e1, col_e2 = st.columns(2)

    with col_e1:
        st.markdown("**📥 Excel 1 — Cartera por Edades**")
        st.caption("Detalle NIU×Mes · Cartera×Localidad · Pivot NIU×Rango · Totales globales")
        out1 = io.BytesIO()
        with pd.ExcelWriter(out1, engine="openpyxl") as w:
            e1 = det_niu[["NIU", "COD LOCALIDAD", "MES_ANIO", "DIAS_ANTIG", "RANGO", "TARIFA_ACUM"]].copy()
            e1.columns = ["NIU", "COD LOCALIDAD", "MES-AÑO", "DÍAS ANTIGÜEDAD", "RANGO CARTERA", "TARIFA"]
            e1.to_excel(w, sheet_name="Detalle_NIU_Mes", index=False)

            e2 = pivot_loc.copy()
            e2.columns = ["COD LOCALIDAD"] + RANGOS + ["TOTAL"]
            e2.to_excel(w, sheet_name="Cartera_x_Localidad", index=False)

            piv = det_niu.pivot_table(
                index=["NIU", "COD LOCALIDAD"], columns="RANGO",
                values="TARIFA_ACUM", aggfunc="sum", fill_value=0
            ).reset_index()
            for r in RANGOS:
                if r not in piv.columns: piv[r] = 0
            piv = piv[["NIU", "COD LOCALIDAD"] + RANGOS]
            piv["TOTAL"] = piv[RANGOS].sum(axis=1)
            piv.to_excel(w, sheet_name="Pivot_NIU_x_Rango", index=False)

            tot_rng = det_niu.groupby("RANGO")["TARIFA_ACUM"].sum().reset_index()
            tot_rng.columns = ["RANGO CARTERA", "TARIFA TOTAL"]
            tot_rng["% DEL TOTAL"] = (tot_rng["TARIFA TOTAL"] / tot_rng["TARIFA TOTAL"].sum() * 100).round(2)
            tot_rng.to_excel(w, sheet_name="Totales_x_Rango", index=False)

        st.download_button(
            label="⬇️ Descargar Cartera por Edades",
            data=out1.getvalue(),
            file_name=f"cartera_edades_{label_meses}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary", use_container_width=True
        )

    with col_e2:
        st.markdown("**📥 Excel 2 — Cartera Total por NIU**")
        st.caption("Un renglón por NIU con la suma total · Resumen por localidad")
        out2 = io.BytesIO()
        with pd.ExcelWriter(out2, engine="openpyxl") as w:
            e3 = cartera_niu.copy()
            e3.columns = ["NIU", "COD LOCALIDAD", "TARIFA TOTAL"]
            e3.to_excel(w, sheet_name="Cartera_Total_NIU", index=False)

            res_loc = cartera_niu.groupby("COD LOCALIDAD")["TARIFA_TOTAL"].sum().reset_index()
            res_loc.columns = ["COD LOCALIDAD", "TARIFA TOTAL"]
            res_loc = res_loc.sort_values("TARIFA TOTAL", ascending=False).reset_index(drop=True)
            res_loc.to_excel(w, sheet_name="Total_x_Localidad", index=False)

        st.download_button(
            label="⬇️ Descargar Cartera Total NIU",
            data=out2.getvalue(),
            file_name=f"cartera_total_NIU_{label_meses}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary", use_container_width=True
        )

else:
    st.info("Carga uno o varios archivos IUF1 con el botón de arriba para comenzar.")
