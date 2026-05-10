import streamlit as st
import requests
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.series import DataPoint
import io

MESOS_CA = {
    'January': 'Gener', 'February': 'Febrer', 'March': 'Marc',
    'April': 'Abril', 'May': 'Maig', 'June': 'Juny',
    'July': 'Juliol', 'August': 'Agost', 'September': 'Setembre',
    'October': 'Octubre', 'November': 'Novembre', 'December': 'Desembre'
}

COLOR_PRINCIPAL = colors.HexColor('#1A3A5C')
COLOR_ACCENT = colors.HexColor('#2E86AB')
COLOR_GROC = colors.HexColor('#F4A923')
COLOR_GRIS_CLAR = colors.HexColor('#F5F7FA')
COLOR_GRIS = colors.HexColor('#CCCCCC')
COLOR_NEGRE = colors.HexColor('#1A1A1A')
COLOR_VERD = colors.HexColor('#2E7D32')

st.set_page_config(page_title="EnergyDataGP", page_icon="⚡", layout="wide")

st.markdown("""
<style>
    .header-bar {
        background: linear-gradient(90deg, #1A3A5C 0%, #2E86AB 100%);
        padding: 1.2rem 2rem; border-radius: 8px; margin-bottom: 1rem;
    }
    .header-title { color: white; font-size: 1.6rem; font-weight: bold; margin: 0; letter-spacing: 2px; }
    .header-sub { color: #F4A923; font-size: 0.85rem; margin: 0; margin-top: 4px; }
    .avis { background: #FFF8E1; border-left: 4px solid #F4A923; padding: 0.8rem 1rem;
            border-radius: 4px; font-size: 0.82rem; color: #555; margin-bottom: 1rem; }
</style>
<div class="header-bar">
    <p class="header-title">ENERGYDATAGP</p>
    <p class="header-sub">Sistema de Generacio d'Informes Tecnics de Produccio Fotovoltaica</p>
</div>
<div class="avis">
    <b>Avis important:</b> Els informes generats tenen caracter orientatiu i no substitueixen
    la signatura d'un tecnic collegiat. Les dades provenen de PVGIS (Comissio Europea).
</div>
""", unsafe_allow_html=True)

st.markdown("---")
st.markdown("### 1. DADES DEL PROJECTE")
col1, col2 = st.columns(2)
with col1:
    nom_projecte = st.text_input("Denominacio del projecte", "")
    num_expedient = st.text_input("Numero de referencia intern", "")
    nom_promotor = st.text_input("Promotor / Titular", "")
with col2:
    nom_tecnic = st.text_input("Tecnic responsable", "")
    num_collegio = st.text_input("Num. col·legiacio", "")
    adreca = st.text_input("Adreca de la instal·lacio", "")

st.markdown("---")
st.markdown("### 2. UBICACIO")
col1, col2, col3 = st.columns(3)
with col1:
    lat = st.number_input("Latitud (N)", value=41.54, format="%.4f")
    lon = st.number_input("Longitud (E)", value=2.45, format="%.4f")
with col2:
    any_ref = st.selectbox("Any de referencia meteorologic", [2020, 2019, 2018, 2017])
with col3:
    st.info("Dades meteorologiques obtingudes automaticament de PVGIS (Comissio Europea)")

st.markdown("---")
st.markdown("### 3. PARAMETRES DE LA INSTAL·LACIO")

mode_calc = st.radio(
    "Com vols definir la instal·lacio?",
    ["Per potencia total (kWp)", "Per superficie disponible (m2)", "Per nombre de panells"],
    horizontal=True
)

col1, col2 = st.columns(2)
with col2:
    st.markdown("**Parametres del panell fotovoltaic:**")
    potencia_panel_wp = st.number_input("Potencia pic del panell (Wp)", value=550, step=5)
    amplada_panel = st.number_input("Amplada del panell (m)", value=1.096, format="%.3f")
    llargada_panel = st.number_input("Llargada del panell (m)", value=2.384, format="%.3f")
    area_panel = amplada_panel * llargada_panel
    st.caption(f"Area per panell: {area_panel:.3f} m2")

with col1:
    if mode_calc == "Per potencia total (kWp)":
        potencia_kwp = st.number_input("Potencia pic total (kWp)", value=100.0)
        num_panels = int((potencia_kwp * 1000) / potencia_panel_wp)
        area_total = num_panels * area_panel
        area_coberta = area_total / 0.75
        st.metric("Nombre de panells", f"{num_panels}")
        st.metric("Area ocupada pels panells", f"{area_total:.1f} m2")
        st.metric("Area de coberta necessaria (75%)", f"{area_coberta:.1f} m2")
    elif mode_calc == "Per superficie disponible (m2)":
        area_coberta = st.number_input("Superficie de coberta disponible (m2)", value=200.0)
        area_util = area_coberta * 0.75
        num_panels = int(area_util / area_panel)
        potencia_kwp = (num_panels * potencia_panel_wp) / 1000
        st.metric("Nombre de panells resultants", f"{num_panels}")
        st.metric("Potencia instal·lada resultant", f"{potencia_kwp:.1f} kWp")
        st.metric("Area util (75% coberta)", f"{area_util:.1f} m2")
    else:
        num_panels = st.number_input("Nombre de panells", value=136, step=1)
        potencia_kwp = (num_panels * potencia_panel_wp) / 1000
        area_total = num_panels * area_panel
        area_coberta = area_total / 0.75
        st.metric("Potencia instal·lada resultant", f"{potencia_kwp:.1f} kWp")
        st.metric("Area ocupada pels panells", f"{area_total:.1f} m2")
        st.metric("Area de coberta necessaria (75%)", f"{area_coberta:.1f} m2")

perdues = st.number_input("Perdues del sistema (%)", value=14)

st.markdown("---")
st.markdown("### 4. PARAMETRES ECONOMICS")
col1, col2, col3 = st.columns(3)
with col1:
    preu_kwh = st.number_input("Preu energia (EUR/kWh)", value=0.18, format="%.3f")
    consum_anual = st.number_input("Consum anual edifici (kWh/any)", value=96378)
with col2:
    cost_instalacio = st.number_input("Cost instal·lacio sense IVA (EUR)", value=0)
    iva = st.number_input("IVA (%)", value=21)
    if cost_instalacio > 0:
        cost_amb_iva = cost_instalacio * (1 + iva/100)
        st.metric("Cost amb IVA", f"{cost_amb_iva:,.0f} EUR")
    else:
        st.info("La IA estimara el cost si es deixa en 0")
        cost_amb_iva = 0
with col3:
    factor_co2 = st.number_input("Factor CO2 (kg CO2/kWh)", value=0.25, format="%.3f")
    preu_excedent = st.number_input("Preu excedent (EUR/kWh)", value=0.05, format="%.3f")

st.markdown("---")
st.markdown("### 5. DESGLOSSAMENT DEL PRESSUPOST")
st.caption("Deixa en blanc els costos que vols que la IA estimi automaticament")

conceptes_default = [
    [f"Moduls FV {potencia_panel_wp} Wp", str(num_panels), ""],
    ["Inversors", "", ""],
    ["Estructura", "1", ""],
    ["Adequacio de sales tecniques", "", ""],
    ["Distribucio i proteccio electrica", "", ""],
    ["Mesura i punt de connexio", "", ""],
    ["Equips de control i monitoratge", "", ""],
    ["Mesures PRL", "", ""],
    ["Senyalitzacio i condicionament", "", ""],
    ["Gestio de residus", "", ""],
    ["Legalitzacio i projecte As-Built", "", ""],
    ["Seguretat i salut", "", ""],
    ["Transport i acopi de material", "", ""],
]

df_pressupost = pd.DataFrame(conceptes_default, columns=["Concepte", "Unitats", "Cost (EUR)"])
pressupost_editat = st.data_editor(
    df_pressupost, use_container_width=True, num_rows="dynamic",
    column_config={
        "Concepte": st.column_config.TextColumn("Concepte", width="large"),
        "Unitats": st.column_config.TextColumn("Unitats", width="small"),
        "Cost (EUR)": st.column_config.TextColumn("Cost (EUR)", width="medium"),
    }
)

st.markdown("---")
st.markdown("### 6. DESCRIPCIO DE L'EMPLACAMENT")
col1, col2 = st.columns(2)
with col1:
    tipus_coberta = st.selectbox("Tipus de coberta", [
        "Coberta plana", "Coberta inclinada", "Coberta coplanar orientada al Sud",
        "Coberta zenital", "Coberta mixta", "Facana", "Altres"
    ])
    orientacio = st.selectbox("Orientacio principal", ["Sud", "Sud-Est", "Sud-Oest", "Est", "Oest", "Zenital"])
    inclinacio = st.number_input("Inclinacio (graus)", value=30)
with col2:
    tipus_activitat = st.text_input("Tipus d'activitat de l'edifici", "Administratiu")
    horari_activitat = st.selectbox("Horari principal d'activitat", [
        "Diurn (8h-18h)", "Nocturn (18h-6h)", "Continu (24h)", "Cap de setmana"
    ])
    obstacles_ombres = st.selectbox("Presencia d'obstacles o ombres", [
        "No hi ha obstacles significatius",
        "Hi ha alguns obstacles menors",
        "Hi ha obstacles significatius"
    ])

st.markdown("---")

with st.expander("Avis legal i condicions d'us"):
    st.markdown("""
    Els documents generats no tenen valor juridic ni tecnic oficial i no substitueixen la memoria
    tecnica signada per un enginyer collegiat competent. Les dades provenen de PVGIS (Comissio Europea).
    EnergyDataGP no assumeix cap responsabilitat pels errors dels informes generats.
    """)

if st.button("GENERAR INFORME TECNIC", type="primary", use_container_width=True):

    if not nom_projecte or not nom_tecnic:
        st.error("Cal omplir com a minim la denominacio del projecte i el tecnic responsable.")
        st.stop()

    with st.spinner("Connectant amb PVGIS..."):
        url = "https://re.jrc.ec.europa.eu/api/v5_2/seriescalc"
        params = {
            "lat": lat, "lon": lon,
            "startyear": any_ref, "endyear": any_ref,
            "pvcalculation": 1, "peakpower": potencia_kwp,
            "loss": perdues, "outputformat": "json", "browser": 0
        }
        response = requests.get(url, params=params)
        df = pd.DataFrame(response.json()['outputs']['hourly'])
        df['time'] = pd.to_datetime(df['time'], format='%Y%m%d:%H%M')
        df.set_index('time', inplace=True)
        produccio_mensual = df['P'].resample('ME').sum() / 1000
        produccio_anual = produccio_mensual.sum()

    mes_max = MESOS_CA[produccio_mensual.idxmax().strftime('%B')]
    mes_min = MESOS_CA[produccio_mensual.idxmin().strftime('%B')]

    autoconsum_ratio = min(consum_anual / produccio_anual, 1.0) if produccio_anual > 0 else 0.5
    autoconsum_kwh = produccio_anual * autoconsum_ratio
    excedents_kwh = produccio_anual - autoconsum_kwh
    estalvi_autoconsum = autoconsum_kwh * preu_kwh
    benefici_excedent = excedents_kwh * preu_excedent
    estalvi_anual = estalvi_autoconsum + benefici_excedent
    co2_estalviat = produccio_anual * factor_co2

    # =====================
    # GENERACIÓ IA
    # =====================
    with st.spinner("Generant textos i pressupost amb IA..."):
        try:
            from groq import Groq
            client = Groq(api_key=st.secrets["GROQ_API_KEY"])

            # PRESSUPOST AMB IA si hi ha costos buits
            costos_buits = pressupost_editat['Cost (EUR)'].apply(
                lambda x: str(x).strip() == '' or str(x).strip() == 'None'
            ).any()

            if costos_buits:
                # Prepara la llista de conceptes amb i sense cost
                conceptes_str = ""
                for _, row in pressupost_editat.iterrows():
                    concepte = str(row['Concepte']) if row['Concepte'] else ''
                    unitats = str(row['Unitats']) if row['Unitats'] else ''
                    cost = str(row['Cost (EUR)']) if row['Cost (EUR)'] else ''
                    if concepte:
                        estat = f"COST CONEGUT: {cost} EUR" if cost and cost != 'None' else "ESTIMA EL COST"
                        conceptes_str += f"- {concepte} (Unitats: {unitats if unitats else 'N/A'}): {estat}\n"

                prompt_pressupost = f"""Ets un enginyer expert en instal·lacions fotovoltaiques a Catalunya.
Estima els costos dels conceptes marcats com "ESTIMA EL COST" per a una instal·lacio de {potencia_kwp:.1f} kWp amb {num_panels} panells de {potencia_panel_wp} Wp.

Preus de referencia del sector a Catalunya:
- Moduls FV 550 Wp: 120 EUR/unitat
- Inversors: 150-200 EUR/kW
- Estructura suport: 70-90 EUR/modul
- Adequacio sales tecniques: 3.000-5.000 EUR
- Distribucio i proteccio electrica: 3.500-5.500 EUR
- Mesura i punt de connexio: 1.500-2.000 EUR
- Equips control i monitoratge: 1.000-1.500 EUR
- Mesures PRL: 4.000-6.000 EUR
- Senyalitzacio i condicionament: 1.000-1.500 EUR
- Gestio de residus: 1.800-2.500 EUR
- Legalitzacio i projecte As-Built: 1.800-2.200 EUR
- Seguretat i salut: 1.500-2.200 EUR
- Transport i acopi: 1.800-2.500 EUR

Conceptes del pressupost:
{conceptes_str}

Respon NOMES en format JSON, sense cap text addicional, seguint exactament aquesta estructura:
{{"pressupost": [{{"concepte": "nom", "unitats": "valor", "cost": 1234.00}}, ...]}}

Inclou TOTS els conceptes, tant els que ja tenien cost com els nous. Per als que ja tenien cost conegut, usa el cost indicat."""

                resp_pressupost = client.chat.completions.create(
                    messages=[{"role": "user", "content": prompt_pressupost}],
                    model="llama-3.3-70b-versatile",
                )

                import json
                text_resp = resp_pressupost.choices[0].message.content
                text_resp = text_resp.strip()
                if text_resp.startswith('```'):
                    text_resp = text_resp.split('```')[1]
                    if text_resp.startswith('json'):
                        text_resp = text_resp[4:]
                pressupost_ia = json.loads(text_resp)
                pressupost_final = pressupost_ia['pressupost']
            else:
                pressupost_final = []
                for _, row in pressupost_editat.iterrows():
                    if row['Concepte']:
                        try:
                            cost_val = float(str(row['Cost (EUR)']).replace('.', '').replace(',', '.').replace('EUR', '').strip())
                        except:
                            cost_val = 0
                        pressupost_final.append({
                            'concepte': str(row['Concepte']),
                            'unitats': str(row['Unitats']) if row['Unitats'] else '',
                            'cost': cost_val
                        })

            total_sense_iva = sum(item['cost'] for item in pressupost_final)
            total_amb_iva = total_sense_iva * (1 + iva/100)
            if cost_instalacio == 0:
                cost_amb_iva = total_amb_iva

            anys_retorn = cost_amb_iva / estalvi_anual if estalvi_anual > 0 else 0

            # TEXT ESTAT ACTUAL
            prompt_estat = f"""Ets un enginyer expert en energia solar fotovoltaica redactant un informe tecnic en catala formal.
Redacta 3-4 frases d'ESTAT ACTUAL de l'edifici sense titol:
- Coberta: {tipus_coberta}, orientacio {orientacio}, inclinacio {inclinacio} graus
- Activitat: {tipus_activitat}, horari {horari_activitat}
- Obstacles: {obstacles_ombres}
Comenca directament amb la descripcio."""
            resp_estat = client.chat.completions.create(
                messages=[{"role": "user", "content": prompt_estat}],
                model="llama-3.3-70b-versatile",
            )
            text_estat = resp_estat.choices[0].message.content

            # TEXT PROPOSTA
            prompt_proposta = f"""Ets un enginyer expert en energia solar fotovoltaica redactant un informe tecnic en catala formal.
Redacta 3-4 frases de PROPOSTA DE MILLORA sense titol:
- {num_panels} moduls de {potencia_panel_wp} Wp, potencia total {potencia_kwp:.1f} kWp
- Orientacio {orientacio}, inclinacio {inclinacio} graus
- Produccio anual estimada: {produccio_anual:,.0f} kWh
- {obstacles_ombres}
Comenca directament amb la proposta."""
            resp_proposta = client.chat.completions.create(
                messages=[{"role": "user", "content": prompt_proposta}],
                model="llama-3.3-70b-versatile",
            )
            text_proposta = resp_proposta.choices[0].message.content

            # TEXT CONCLUSIONS
            prompt_conclusions = f"""Ets un enginyer expert en energia solar fotovoltaica redactant un informe tecnic en catala formal.
Redacta 4-5 frases de CONCLUSIONS sense titol:
- Potencia: {potencia_kwp:.1f} kWp, {num_panels} panells
- Produccio anual: {produccio_anual:,.0f} kWh
- Autoconsum: {autoconsum_kwh:,.0f} kWh/any, Excedents: {excedents_kwh:,.0f} kWh/any
- Cost amb IVA: {cost_amb_iva:,.0f} EUR
- Estalvi anual: {estalvi_anual:,.0f} EUR/any, Payback: {anys_retorn:.1f} anys
- CO2 estalviat: {co2_estalviat:,.0f} kg/any
Comenca directament amb les conclusions."""
            resp_conclusions = client.chat.completions.create(
                messages=[{"role": "user", "content": prompt_conclusions}],
                model="llama-3.3-70b-versatile",
            )
            text_conclusions = resp_conclusions.choices[0].message.content
            st.success("Textos i pressupost generats per IA correctament")

        except Exception as e:
            st.warning(f"IA no disponible ({e}). S'usa text estandard.")
            pressupost_final = []
            for _, row in pressupost_editat.iterrows():
                if row['Concepte']:
                    try:
                        cost_val = float(str(row['Cost (EUR)']).replace('.', '').replace(',', '.').replace('EUR', '').strip())
                    except:
                        cost_val = 0
                    pressupost_final.append({
                        'concepte': str(row['Concepte']),
                        'unitats': str(row['Unitats']) if row['Unitats'] else '',
                        'cost': cost_val
                    })
            total_sense_iva = sum(item['cost'] for item in pressupost_final)
            total_amb_iva = total_sense_iva * (1 + iva/100)
            if cost_instalacio == 0:
                cost_amb_iva = total_amb_iva
            anys_retorn = cost_amb_iva / estalvi_anual if estalvi_anual > 0 else 0
            text_estat = f"L'edifici disposa d'una {tipus_coberta.lower()} amb orientacio {orientacio} i inclinacio de {inclinacio} graus. Activitat principal: {tipus_activitat.lower()}, horari {horari_activitat.lower()}. {obstacles_ombres}."
            text_proposta = f"Es proposa instal·lar {num_panels} moduls de {potencia_panel_wp} Wp, potencia total {potencia_kwp:.1f} kWp, orientacio {orientacio} a {inclinacio} graus. Produccio anual estimada: {produccio_anual:,.0f} kWh."
            text_conclusions = f"La instal·lacio presenta una produccio anual de {produccio_anual:,.0f} kWh, estalvi de {estalvi_anual:,.0f} EUR/any i payback de {anys_retorn:.1f} anys. Reduccio CO2: {co2_estalviat:,.0f} kg/any."

    # =====================
    # GRÀFIQUES
    # =====================
    with st.spinner("Generant grafiques..."):
        CP = '#1A3A5C'
        CA = '#2E86AB'
        CV = '#2E7D32'

        fig1, ax1 = plt.subplots(figsize=(12, 4))
        ax1.bar(range(12), produccio_mensual.values, color=CP, alpha=0.85, width=0.6)
        ax1.set_xticks(range(12))
        ax1.set_xticklabels(['Gen','Feb','Mar','Abr','Mai','Jun','Jul','Ago','Set','Oct','Nov','Des'], fontsize=10)
        ax1.set_ylabel('Energia produida (kWh)', fontsize=10)
        ax1.set_title(f'Figura 1. Produccio energetica mensual - {nom_projecte}', fontsize=11, fontweight='bold', color=CP)
        ax1.axhline(produccio_mensual.mean(), color=CA, linestyle='--', linewidth=1.5, label='Mitjana mensual')
        for i, val in enumerate(produccio_mensual.values):
            ax1.text(i, val + 150, f'{val:.0f}', ha='center', fontsize=8, color=CP)
        ax1.legend(fontsize=9)
        ax1.spines['top'].set_visible(False)
        ax1.spines['right'].set_visible(False)
        ax1.grid(axis='y', alpha=0.3)
        plt.tight_layout()
        plt.savefig('g1_mensual.png', dpi=180, bbox_inches='tight')
        plt.close()

        dia_estiu = df.loc[f'{any_ref}-06-21', 'P'] / 1000
        dia_hivern = df.loc[f'{any_ref}-12-21', 'P'] / 1000
        fig2, ax2 = plt.subplots(figsize=(12, 4))
        ax2.plot(range(len(dia_estiu)), dia_estiu.values, color=CA, linewidth=2.5, label="Solstici d'estiu (21 juny)")
        ax2.plot(range(len(dia_hivern)), dia_hivern.values, color=CP, linewidth=2.5, label="Solstici d'hivern (21 desembre)")
        ax2.fill_between(range(len(dia_estiu)), dia_estiu.values, alpha=0.1, color=CA)
        ax2.fill_between(range(len(dia_hivern)), dia_hivern.values, alpha=0.1, color=CP)
        ax2.set_ylabel('Potencia generada (kW)', fontsize=10)
        ax2.set_title('Figura 2. Corba de generacio horaria - Comparativa estacional', fontsize=11, fontweight='bold', color=CP)
        ax2.legend(fontsize=9)
        ax2.spines['top'].set_visible(False)
        ax2.spines['right'].set_visible(False)
        ax2.grid(alpha=0.3)
        plt.tight_layout()
        plt.savefig('g2_corbes.png', dpi=180, bbox_inches='tight')
        plt.close()

        mesos_tipics = {'Gener': f'{any_ref}-01-15', 'Abril': f'{any_ref}-04-15', 'Juliol': f'{any_ref}-07-15', 'Octubre': f'{any_ref}-10-15'}
        colors_mesos = [CP, CA, '#F4A923', CV]
        fig3, ax3 = plt.subplots(figsize=(12, 5))
        for (mes_nom, data), color in zip(mesos_tipics.items(), colors_mesos):
            try:
                dia = df.loc[data, 'P'] / 1000
                ax3.plot(range(len(dia)), dia.values, linewidth=2, label=mes_nom, color=color)
                ax3.fill_between(range(len(dia)), dia.values, alpha=0.05, color=color)
            except:
                pass
        ax3.set_xlabel('Hora del dia', fontsize=10)
        ax3.set_ylabel('Potencia generada (kW)', fontsize=10)
        ax3.set_title('Figura 3. Corbes de generacio horaria per mesos tipus', fontsize=11, fontweight='bold', color=CP)
        ax3.set_xticks(range(0, 24, 2))
        ax3.set_xticklabels([f'{h:02d}:00' for h in range(0, 24, 2)], rotation=45, fontsize=8)
        ax3.legend(fontsize=9)
        ax3.spines['top'].set_visible(False)
        ax3.spines['right'].set_visible(False)
        ax3.grid(alpha=0.3)
        plt.tight_layout()
        plt.savefig('g3_horaria.png', dpi=180, bbox_inches='tight')
        plt.close()

        anys_list = list(range(0, 26))
        flux_acumulat = [-cost_amb_iva + estalvi_anual * a for a in anys_list]
        fig4, ax4 = plt.subplots(figsize=(12, 4))
        colors_bars = [CA if v < 0 else CV for v in flux_acumulat]
        ax4.bar(anys_list, flux_acumulat, color=colors_bars, alpha=0.8, width=0.7)
        ax4.axhline(0, color='black', linewidth=1)
        ax4.set_xlabel('Anys des de la posada en marxa', fontsize=10)
        ax4.set_ylabel('Flux de caixa acumulat (EUR)', fontsize=10)
        ax4.set_title('Figura 4. Analisi del retorn de la inversio (ROI)', fontsize=11, fontweight='bold', color=CP)
        ax4.spines['top'].set_visible(False)
        ax4.spines['right'].set_visible(False)
        ax4.grid(axis='y', alpha=0.3)
        llegenda = [mpatches.Patch(color=CA, label='Periode de retorn'), mpatches.Patch(color=CV, label='Benefici net')]
        ax4.legend(handles=llegenda, fontsize=9)
        plt.tight_layout()
        plt.savefig('g4_roi.png', dpi=180, bbox_inches='tight')
        plt.close()

        fig5, ax5 = plt.subplots(figsize=(8, 5))
        labels = ['Autoconsum', 'Excedents exportats']
        sizes = [autoconsum_kwh, excedents_kwh]
        colors_pie = [CP, CA]
        ax5.pie(sizes, labels=labels, colors=colors_pie, autopct='%1.1f%%',
               shadow=False, startangle=90, textprops={'fontsize': 11})
        ax5.set_title("Figura 5. Distribucio de l'energia generada", fontsize=11, fontweight='bold', color=CP)
        plt.tight_layout()
        plt.savefig('g5_autoconsum.png', dpi=180, bbox_inches='tight')
        plt.close()

    # Mostra resum
    st.success("Informe generat correctament")
    st.markdown("### Resum executiu")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Produccio anual", f"{produccio_anual:,.0f} kWh")
    c2.metric("Autoconsum", f"{autoconsum_kwh:,.0f} kWh")
    c3.metric("Estalvi anual", f"{estalvi_anual:,.0f} EUR")
    c4.metric("Payback", f"{anys_retorn:.1f} anys")
    c5.metric("CO2 estalviat", f"{co2_estalviat:,.0f} kg")
    st.pyplot(fig1)
    st.pyplot(fig3)

    # =====================
    # PDF
    # =====================
    with st.spinner("Generant PDF..."):
        doc_pdf = SimpleDocTemplate(
            "informe_energydatagp.pdf", pagesize=A4,
            rightMargin=2.5*cm, leftMargin=2.5*cm, topMargin=2*cm, bottomMargin=2*cm
        )
        styles = getSampleStyleSheet()
        estil_seccio = ParagraphStyle('seccio', fontSize=11, fontName='Helvetica-Bold',
            textColor=COLOR_PRINCIPAL, spaceBefore=12, spaceAfter=6)
        estil_normal = ParagraphStyle('normal', fontSize=9, fontName='Helvetica',
            textColor=COLOR_NEGRE, leading=14, spaceAfter=4)
        estil_peu = ParagraphStyle('peu', fontSize=7, fontName='Helvetica',
            textColor=colors.grey, alignment=TA_CENTER)
        estil_avis = ParagraphStyle('avis', fontSize=8, fontName='Helvetica',
            textColor=colors.grey, leading=12, spaceAfter=8)

        elements = []
        elements.append(HRFlowable(width="100%", thickness=4, color=COLOR_PRINCIPAL, spaceAfter=8))
        cap_data = [[
            Paragraph('<b><font size=14>ENERGYDATAGP</font></b><br/><font size=8 color=grey>Solucions d\'Analisi de Dades Energetiques</font>',
                     ParagraphStyle('cap', fontSize=9, fontName='Helvetica-Bold', textColor=COLOR_PRINCIPAL)),
            Paragraph('<b>INFORME TECNIC DE PRODUCCIO</b><br/>INSTAL·LACIO FOTOVOLTAICA',
                     ParagraphStyle('cap2', fontSize=12, fontName='Helvetica-Bold', textColor=COLOR_PRINCIPAL, alignment=TA_CENTER)),
            Paragraph(f'Ref.: <b>{num_expedient if num_expedient else "-"}</b><br/>Data: <b>{datetime.now().strftime("%d/%m/%Y")}</b><br/>Versio: 1.0',
                     ParagraphStyle('cap3', fontSize=9, fontName='Helvetica', textColor=COLOR_PRINCIPAL, alignment=TA_RIGHT)),
        ]]
        taula_cap = Table(cap_data, colWidths=[5.5*cm, 8*cm, 4*cm])
        taula_cap.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'MIDDLE'),('BACKGROUND',(0,0),(-1,-1),COLOR_GRIS_CLAR),('PADDING',(0,0),(-1,-1),8)]))
        elements.append(taula_cap)
        elements.append(HRFlowable(width="100%", thickness=2, color=COLOR_GROC, spaceAfter=12))

        elements.append(Paragraph("1. IDENTIFICACIO DEL PROJECTE", estil_seccio))
        elements.append(HRFlowable(width="100%", thickness=0.5, color=COLOR_GRIS, spaceAfter=6))
        dades_id = [
            ['Denominacio:', nom_projecte, 'Promotor:', nom_promotor if nom_promotor else '-'],
            ['Referencia:', num_expedient if num_expedient else '-', 'Adreca:', adreca if adreca else '-'],
            ['Coordenades:', f'{lat}N, {lon}E', 'Any ref.:', str(any_ref)],
            ['Tecnic:', nom_tecnic, 'Col·legiacio:', num_collegio if num_collegio else '-'],
        ]
        taula_id = Table(dades_id, colWidths=[3.5*cm, 7*cm, 3*cm, 4*cm])
        taula_id.setStyle(TableStyle([
            ('FONTNAME',(0,0),(0,-1),'Helvetica-Bold'),('FONTNAME',(2,0),(2,-1),'Helvetica-Bold'),
            ('FONTSIZE',(0,0),(-1,-1),9),('TEXTCOLOR',(0,0),(0,-1),COLOR_PRINCIPAL),
            ('TEXTCOLOR',(2,0),(2,-1),COLOR_PRINCIPAL),
            ('ROWBACKGROUNDS',(0,0),(-1,-1),[COLOR_GRIS_CLAR,colors.white]),
            ('GRID',(0,0),(-1,-1),0.3,COLOR_GRIS),('PADDING',(0,0),(-1,-1),5)
        ]))
        elements.append(taula_id)
        elements.append(Spacer(1, 0.4*cm))

        elements.append(Paragraph("2. ESTAT ACTUAL", estil_seccio))
        elements.append(HRFlowable(width="100%", thickness=0.5, color=COLOR_GRIS, spaceAfter=6))
        elements.append(Paragraph(text_estat, estil_normal))
        elements.append(Spacer(1, 0.3*cm))

        elements.append(Paragraph("3. PROPOSTA DE MILLORA", estil_seccio))
        elements.append(HRFlowable(width="100%", thickness=0.5, color=COLOR_GRIS, spaceAfter=6))
        elements.append(Paragraph(text_proposta, estil_normal))
        elements.append(Spacer(1, 0.3*cm))

        elements.append(Paragraph("4. PRESSUPOST DETALLAT", estil_seccio))
        elements.append(HRFlowable(width="100%", thickness=0.5, color=COLOR_GRIS, spaceAfter=6))
        pressupost_data = [['CONCEPTE', 'UNITATS', 'COST']]
        for item in pressupost_final:
            cost_str = f"{item['cost']:,.2f} EUR" if item['cost'] > 0 else '-'
            pressupost_data.append([item['concepte'], str(item['unitats']), cost_str])
        pressupost_data.append(['TOTAL (sense IVA)', '', f'{total_sense_iva:,.2f} EUR'])
        pressupost_data.append([f'TOTAL (amb IVA {iva}%)', '', f'{total_amb_iva:,.2f} EUR'])
        taula_pressupost = Table(pressupost_data, colWidths=[10*cm, 3*cm, 4.5*cm])
        taula_pressupost.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,0),COLOR_PRINCIPAL),('TEXTCOLOR',(0,0),(-1,0),colors.white),
            ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,-1),9),
            ('ROWBACKGROUNDS',(0,1),(-1,-3),[COLOR_GRIS_CLAR,colors.white]),
            ('BACKGROUND',(0,-2),(-1,-1),COLOR_GRIS_CLAR),('FONTNAME',(0,-2),(-1,-1),'Helvetica-Bold'),
            ('GRID',(0,0),(-1,-1),0.3,COLOR_GRIS),('PADDING',(0,0),(-1,-1),5),('ALIGN',(2,0),(2,-1),'RIGHT')
        ]))
        elements.append(taula_pressupost)
        elements.append(Spacer(1, 0.4*cm))

        elements.append(Paragraph("5. PARAMETRES TECNICS", estil_seccio))
        elements.append(HRFlowable(width="100%", thickness=0.5, color=COLOR_GRIS, spaceAfter=6))
        dades_tec = [
            ['Parametre', 'Valor', 'Unitat', 'Observacions'],
            ['Potencia pic instal·lada', f'{potencia_kwp:.1f}', 'kWp', 'Condicions estandard (STC)'],
            ['Nombre de moduls', f'{num_panels}', 'ut.', f'{potencia_panel_wp} Wp per modul'],
            ['Dimensions modul', f'{llargada_panel:.3f}x{amplada_panel:.3f}', 'm', f'Area: {area_panel:.3f} m2/modul'],
            ['Perdues sistema', f'{perdues}', '%', 'Cablejat, inversor, bruticia, temperatura'],
            ['Produccio anual estimada', f'{produccio_anual:,.0f}', 'kWh/any', f'Any ref.: {any_ref}'],
            ['Mes maxima produccio', mes_max, f'{produccio_mensual.max():,.0f} kWh', 'Pic estival'],
            ['Mes minima produccio', mes_min, f'{produccio_mensual.min():,.0f} kWh', 'Minim hivernal'],
            ['Hores equivalents sol', f'{produccio_anual/potencia_kwp:,.0f}', 'HES/any', 'Hores a potencia nominal'],
            ['Orientacio / Inclinacio', orientacio, f'{inclinacio} graus', tipus_coberta],
        ]
        taula_tec = Table(dades_tec, colWidths=[5.5*cm, 3*cm, 3*cm, 6*cm])
        taula_tec.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,0),COLOR_PRINCIPAL),('TEXTCOLOR',(0,0),(-1,0),colors.white),
            ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,-1),9),
            ('ROWBACKGROUNDS',(0,1),(-1,-1),[COLOR_GRIS_CLAR,colors.white]),
            ('GRID',(0,0),(-1,-1),0.3,COLOR_GRIS),('PADDING',(0,0),(-1,-1),5),('ALIGN',(1,1),(2,-1),'CENTER')
        ]))
        elements.append(taula_tec)
        elements.append(Spacer(1, 0.4*cm))

        elements.append(Paragraph("6. ANALISI DE LA PRODUCCIO ENERGETICA MENSUAL", estil_seccio))
        elements.append(HRFlowable(width="100%", thickness=0.5, color=COLOR_GRIS, spaceAfter=6))
        elements.append(Image('g1_mensual.png', width=16*cm, height=6*cm))
        elements.append(Paragraph(
            f"La instal·lacio presenta una produccio anual estimada de <b>{produccio_anual:,.0f} kWh</b>. "
            f"Maxima produccio: <b>{mes_max}</b> ({produccio_mensual.max():,.0f} kWh). "
            f"Minima produccio: <b>{mes_min}</b> ({produccio_mensual.min():,.0f} kWh).", estil_normal))
        elements.append(Spacer(1, 0.3*cm))

        elements.append(Paragraph("7. CORBES DE GENERACIO HORARIA", estil_seccio))
        elements.append(HRFlowable(width="100%", thickness=0.5, color=COLOR_GRIS, spaceAfter=6))
        elements.append(Image('g2_corbes.png', width=16*cm, height=6*cm))
        elements.append(Spacer(1, 0.2*cm))
        elements.append(Image('g3_horaria.png', width=16*cm, height=6*cm))
        elements.append(Paragraph(
            f"Potencia maxima solstici d'estiu: <b>{dia_estiu.max():.1f} kW</b>. "
            f"Potencia maxima solstici d'hivern: <b>{dia_hivern.max():.1f} kW</b>.", estil_normal))
        elements.append(Spacer(1, 0.3*cm))

        elements.append(Paragraph("8. ANALISI ECONOMICA I RETORN DE LA INVERSIO", estil_seccio))
        elements.append(HRFlowable(width="100%", thickness=0.5, color=COLOR_GRIS, spaceAfter=6))
        dades_sim = [
            ['', 'Consum xarxa (kWh/any)', 'Generacio (kWh/any)', 'Autoconsum (kWh/any)', 'Excedents (kWh/any)'],
            ['Energia', f'{consum_anual:,.0f}', f'{produccio_anual:,.0f}', f'{autoconsum_kwh:,.0f}', f'{excedents_kwh:,.0f}'],
            ['Cost/Benefici', '-', '-', f'{estalvi_autoconsum:,.0f} EUR', f'{benefici_excedent:,.0f} EUR'],
        ]
        taula_sim = Table(dades_sim, colWidths=[2.5*cm, 3.5*cm, 3.5*cm, 3.5*cm, 3.5*cm])
        taula_sim.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,0),COLOR_ACCENT),('TEXTCOLOR',(0,0),(-1,0),colors.white),
            ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,-1),8),
            ('ROWBACKGROUNDS',(0,1),(-1,-1),[COLOR_GRIS_CLAR,colors.white]),
            ('GRID',(0,0),(-1,-1),0.3,COLOR_GRIS),('PADDING',(0,0),(-1,-1),4),
            ('ALIGN',(1,0),(-1,-1),'CENTER'),('FONTNAME',(0,1),(0,-1),'Helvetica-Bold')
        ]))
        elements.append(taula_sim)
        elements.append(Spacer(1, 0.3*cm))
        dades_eco = [
            ['Consum actual (kWh/any)', 'Estalvi energia (kWh/any)', 'Estalvi economic (EUR/any)', 'Inversio (EUR)', 'Payback (anys)', 'Estalvi CO2 (kg/any)'],
            [f'{consum_anual:,.0f}', f'{autoconsum_kwh:,.0f}', f'{estalvi_anual:,.0f}', f'{cost_amb_iva:,.0f}', f'{anys_retorn:.1f}', f'{co2_estalviat:,.0f}'],
        ]
        taula_eco = Table(dades_eco, colWidths=[3*cm, 3*cm, 3*cm, 2.5*cm, 2.5*cm, 3.5*cm])
        taula_eco.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,0),COLOR_PRINCIPAL),('TEXTCOLOR',(0,0),(-1,0),colors.white),
            ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,-1),8),
            ('ROWBACKGROUNDS',(0,1),(-1,-1),[COLOR_GRIS_CLAR]),
            ('GRID',(0,0),(-1,-1),0.3,COLOR_GRIS),('PADDING',(0,0),(-1,-1),4),('ALIGN',(0,0),(-1,-1),'CENTER')
        ]))
        elements.append(taula_eco)
        elements.append(Spacer(1, 0.3*cm))
        elements.append(Image('g5_autoconsum.png', width=10*cm, height=7*cm))
        elements.append(Spacer(1, 0.3*cm))
        elements.append(Image('g4_roi.png', width=16*cm, height=5.5*cm))
        elements.append(Spacer(1, 0.3*cm))

        elements.append(Paragraph("9. CONCLUSIONS", estil_seccio))
        elements.append(HRFlowable(width="100%", thickness=0.5, color=COLOR_GRIS, spaceAfter=6))
        elements.append(Paragraph(text_conclusions, estil_normal))
        elements.append(Spacer(1, 0.8*cm))

        sig_data = [
            [Paragraph('El/La tecnic/a responsable', estil_normal), Paragraph('Conforme el/la promotor/a', estil_normal)],
            [Paragraph(f'<br/><br/><br/>Nom: {nom_tecnic}<br/>Col·legiacio: {num_collegio if num_collegio else "-"}<br/>Data: {datetime.now().strftime("%d/%m/%Y")}', estil_normal),
             Paragraph('<br/><br/><br/>Nom:<br/>DNI/NIF:<br/>Data:', estil_normal)],
        ]
        taula_sig = Table(sig_data, colWidths=[8.5*cm, 8.5*cm])
        taula_sig.setStyle(TableStyle([('BOX',(0,0),(0,-1),0.5,COLOR_GRIS),('BOX',(1,0),(1,-1),0.5,COLOR_GRIS),('PADDING',(0,0),(-1,-1),8)]))
        elements.append(taula_sig)
        elements.append(Spacer(1, 0.5*cm))
        elements.append(HRFlowable(width="100%", thickness=0.5, color=COLOR_GRIS, spaceAfter=4))
        elements.append(Paragraph(
            "<b>AVIS LEGAL:</b> Document generat amb EnergyDataGP amb finalitat orientativa. "
            "Dades: PVGIS - Comissio Europea. No substitueix la memoria tecnica oficial.", estil_avis))
        elements.append(HRFlowable(width="100%", thickness=2, color=COLOR_PRINCIPAL, spaceAfter=4))
        elements.append(Paragraph(
            f"EnergyDataGP - {datetime.now().strftime('%d/%m/%Y %H:%M')} - PVGIS - Comissio Europea - www.energydatagp.com",
            estil_peu))
        doc_pdf.build(elements)

    # =====================
    # WORD
    # =====================
    with st.spinner("Generant document Word..."):
        doc_word = Document()

        # Estils
        style = doc_word.styles['Normal']
        style.font.name = 'Calibri'
        style.font.size = Pt(10)

        # Capçalera
        header = doc_word.add_heading('ENERGYDATAGP', 0)
        header.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = header.runs[0]
        run.font.color.rgb = RGBColor(0x1A, 0x3A, 0x5C)

        sub = doc_word.add_paragraph('INFORME TECNIC DE PRODUCCIO - INSTAL·LACIO FOTOVOLTAICA')
        sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sub.runs[0].font.bold = True

        doc_word.add_paragraph(f'Referencia: {num_expedient if num_expedient else "-"} | Data: {datetime.now().strftime("%d/%m/%Y")}').alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc_word.add_paragraph()

        def add_seccio(doc, titol):
            h = doc.add_heading(titol, level=1)
            h.runs[0].font.color.rgb = RGBColor(0x1A, 0x3A, 0x5C)

        # Seccions
        add_seccio(doc_word, '1. IDENTIFICACIO DEL PROJECTE')
        taula_id_w = doc_word.add_table(rows=4, cols=4)
        taula_id_w.style = 'Table Grid'
        dades_id_w = [
            ['Denominacio:', nom_projecte, 'Promotor:', nom_promotor if nom_promotor else '-'],
            ['Referencia:', num_expedient if num_expedient else '-', 'Adreca:', adreca if adreca else '-'],
            ['Coordenades:', f'{lat}N, {lon}E', 'Any ref.:', str(any_ref)],
            ['Tecnic:', nom_tecnic, 'Col·legiacio:', num_collegio if num_collegio else '-'],
        ]
        for i, fila in enumerate(dades_id_w):
            for j, val in enumerate(fila):
                cell = taula_id_w.rows[i].cells[j]
                cell.text = val
                if j % 2 == 0:
                    cell.paragraphs[0].runs[0].font.bold = True
        doc_word.add_paragraph()

        add_seccio(doc_word, '2. ESTAT ACTUAL')
        doc_word.add_paragraph(text_estat)

        add_seccio(doc_word, '3. PROPOSTA DE MILLORA')
        doc_word.add_paragraph(text_proposta)

        add_seccio(doc_word, '4. PRESSUPOST DETALLAT')
        taula_pres_w = doc_word.add_table(rows=len(pressupost_final)+3, cols=3)
        taula_pres_w.style = 'Table Grid'
        caps = taula_pres_w.rows[0].cells
        for j, cap in enumerate(['CONCEPTE', 'UNITATS', 'COST']):
            caps[j].text = cap
            caps[j].paragraphs[0].runs[0].font.bold = True
        for i, item in enumerate(pressupost_final):
            fila = taula_pres_w.rows[i+1].cells
            fila[0].text = item['concepte']
            fila[1].text = str(item['unitats'])
            fila[2].text = f"{item['cost']:,.2f} EUR" if item['cost'] > 0 else '-'
        total_row1 = taula_pres_w.rows[-2].cells
        total_row1[0].text = 'TOTAL (sense IVA)'
        total_row1[2].text = f'{total_sense_iva:,.2f} EUR'
        total_row1[0].paragraphs[0].runs[0].font.bold = True
        total_row2 = taula_pres_w.rows[-1].cells
        total_row2[0].text = f'TOTAL (amb IVA {iva}%)'
        total_row2[2].text = f'{total_amb_iva:,.2f} EUR'
        total_row2[0].paragraphs[0].runs[0].font.bold = True
        doc_word.add_paragraph()

        add_seccio(doc_word, '5. PARAMETRES TECNICS')
        taula_tec_w = doc_word.add_table(rows=10, cols=4)
        taula_tec_w.style = 'Table Grid'
        dades_tec_w = [
            ['Parametre', 'Valor', 'Unitat', 'Observacions'],
            ['Potencia pic', f'{potencia_kwp:.1f}', 'kWp', 'STC'],
            ['Nombre moduls', f'{num_panels}', 'ut.', f'{potencia_panel_wp} Wp/modul'],
            ['Perdues', f'{perdues}', '%', 'Sistema complet'],
            ['Produccio anual', f'{produccio_anual:,.0f}', 'kWh/any', f'Any {any_ref}'],
            ['Mes maxim', mes_max, f'{produccio_mensual.max():,.0f} kWh', ''],
            ['Mes minim', mes_min, f'{produccio_mensual.min():,.0f} kWh', ''],
            ['HES/any', f'{produccio_anual/potencia_kwp:,.0f}', 'h', ''],
            ['Orientacio', orientacio, f'{inclinacio} graus', tipus_coberta],
        ]
        for i, fila in enumerate(dades_tec_w):
            for j, val in enumerate(fila):
                taula_tec_w.rows[i].cells[j].text = val
                if i == 0:
                    taula_tec_w.rows[i].cells[j].paragraphs[0].runs[0].font.bold = True
        doc_word.add_paragraph()

        add_seccio(doc_word, '6. PRODUCCIO MENSUAL')
        doc_word.add_picture('g1_mensual.png', width=Inches(6))
        doc_word.add_paragraph()

        add_seccio(doc_word, '7. CORBES DE GENERACIO HORARIA')
        doc_word.add_picture('g2_corbes.png', width=Inches(6))
        doc_word.add_paragraph()
        doc_word.add_picture('g3_horaria.png', width=Inches(6))
        doc_word.add_paragraph()

        add_seccio(doc_word, '8. ANALISI ECONOMICA')
        taula_eco_w = doc_word.add_table(rows=3, cols=6)
        taula_eco_w.style = 'Table Grid'
        caps_eco = ['Consum (kWh/any)', 'Estalvi energia (kWh/any)', 'Estalvi economic (EUR/any)', 'Inversio (EUR)', 'Payback (anys)', 'CO2 estalviat (kg/any)']
        vals_eco = [f'{consum_anual:,.0f}', f'{autoconsum_kwh:,.0f}', f'{estalvi_anual:,.0f}', f'{cost_amb_iva:,.0f}', f'{anys_retorn:.1f}', f'{co2_estalviat:,.0f}']
        for j, cap in enumerate(caps_eco):
            taula_eco_w.rows[0].cells[j].text = cap
            taula_eco_w.rows[0].cells[j].paragraphs[0].runs[0].font.bold = True
        for j, val in enumerate(vals_eco):
            taula_eco_w.rows[1].cells[j].text = val
        doc_word.add_paragraph()
        doc_word.add_picture('g4_roi.png', width=Inches(6))
        doc_word.add_paragraph()

        add_seccio(doc_word, '9. CONCLUSIONS')
        doc_word.add_paragraph(text_conclusions)
        doc_word.add_paragraph()

        # Signatura
        taula_sig_w = doc_word.add_table(rows=3, cols=2)
        taula_sig_w.style = 'Table Grid'
        taula_sig_w.rows[0].cells[0].text = 'El/La tecnic/a responsable'
        taula_sig_w.rows[0].cells[1].text = 'Conforme el/la promotor/a'
        taula_sig_w.rows[1].cells[0].text = f'Nom: {nom_tecnic}'
        taula_sig_w.rows[2].cells[0].text = f'Data: {datetime.now().strftime("%d/%m/%Y")}'

        doc_word.add_paragraph()
        peu = doc_word.add_paragraph(f'EnergyDataGP - {datetime.now().strftime("%d/%m/%Y")} - PVGIS Comissio Europea - www.energydatagp.com')
        peu.alignment = WD_ALIGN_PARAGRAPH.CENTER
        peu.runs[0].font.size = Pt(8)
        peu.runs[0].font.color.rgb = RGBColor(0x88, 0x88, 0x88)

        word_buffer = io.BytesIO()
        doc_word.save(word_buffer)
        word_buffer.seek(0)

    # =====================
    # EXCEL
    # =====================
    with st.spinner("Generant Excel..."):
        wb = openpyxl.Workbook()

        # Colors
        blau_fosc = '1A3A5C'
        blau_clar = '2E86AB'
        groc = 'F4A923'
        gris = 'F5F7FA'
        verd = '2E7D32'

        header_font = Font(bold=True, color='FFFFFF', size=11)
        header_fill_blau = PatternFill("solid", fgColor=blau_fosc)
        header_fill_accent = PatternFill("solid", fgColor=blau_clar)
        gris_fill = PatternFill("solid", fgColor='F5F7FA')
        border = Border(
            left=Side(style='thin', color='CCCCCC'),
            right=Side(style='thin', color='CCCCCC'),
            top=Side(style='thin', color='CCCCCC'),
            bottom=Side(style='thin', color='CCCCCC')
        )

        # ---- FULL 1: RESUM ----
        ws1 = wb.active
        ws1.title = "Resum"
        ws1.column_dimensions['A'].width = 35
        ws1.column_dimensions['B'].width = 20
        ws1.column_dimensions['C'].width = 15

        ws1['A1'] = 'ENERGYDATAGP - INFORME TECNIC FOTOVOLTAIC'
        ws1['A1'].font = Font(bold=True, size=14, color=blau_fosc)
        ws1.merge_cells('A1:C1')
        ws1['A1'].alignment = Alignment(horizontal='center')

        ws1['A2'] = f'Projecte: {nom_projecte}'
        ws1['A3'] = f'Data: {datetime.now().strftime("%d/%m/%Y")}'
        ws1['A4'] = f'Tecnic: {nom_tecnic}'

        ws1['A6'] = 'PARAMETRE'
        ws1['B6'] = 'VALOR'
        ws1['C6'] = 'UNITAT'
        for col in ['A6', 'B6', 'C6']:
            ws1[col].font = header_font
            ws1[col].fill = header_fill_blau
            ws1[col].alignment = Alignment(horizontal='center')
            ws1[col].border = border

        dades_resum = [
            ('Potencia instal·lada', f'{potencia_kwp:.1f}', 'kWp'),
            ('Nombre de moduls', num_panels, 'ut.'),
            ('Produccio anual estimada', f'{produccio_anual:,.0f}', 'kWh/any'),
            ('Mes de maxima produccio', f'{mes_max} ({produccio_mensual.max():,.0f} kWh)', ''),
            ('Mes de minima produccio', f'{mes_min} ({produccio_mensual.min():,.0f} kWh)', ''),
            ('Hores equivalents de sol', f'{produccio_anual/potencia_kwp:,.0f}', 'HES/any'),
            ('Autoconsum', f'{autoconsum_kwh:,.0f}', 'kWh/any'),
            ('Excedents exportats', f'{excedents_kwh:,.0f}', 'kWh/any'),
            ('Estalvi economic anual', f'{estalvi_anual:,.0f}', 'EUR/any'),
            ('Cost instal·lacio amb IVA', f'{cost_amb_iva:,.0f}', 'EUR'),
            ('Payback', f'{anys_retorn:.1f}', 'anys'),
            ('CO2 estalviat', f'{co2_estalviat:,.0f}', 'kg/any'),
        ]
        for i, (param, valor, unitat) in enumerate(dades_resum):
            row = i + 7
            ws1[f'A{row}'] = param
            ws1[f'B{row}'] = valor
            ws1[f'C{row}'] = unitat
            fill = gris_fill if i % 2 == 0 else PatternFill("solid", fgColor='FFFFFF')
            for col in [f'A{row}', f'B{row}', f'C{row}']:
                ws1[col].fill = fill
                ws1[col].border = border
            ws1[f'B{row}'].alignment = Alignment(horizontal='center')

        # ---- FULL 2: PRODUCCIO MENSUAL ----
        ws2 = wb.create_sheet("Produccio Mensual")
        ws2.column_dimensions['A'].width = 15
        ws2.column_dimensions['B'].width = 20

        ws2['A1'] = 'MES'
        ws2['B1'] = 'PRODUCCIO (kWh)'
        ws2['A1'].font = header_font
        ws2['B1'].font = header_font
        ws2['A1'].fill = header_fill_blau
        ws2['B1'].fill = header_fill_blau
        ws2['A1'].alignment = Alignment(horizontal='center')
        ws2['B1'].alignment = Alignment(horizontal='center')
        ws2['A1'].border = border
        ws2['B1'].border = border

        mesos_noms = ['Gener', 'Febrer', 'Marc', 'Abril', 'Maig', 'Juny',
                     'Juliol', 'Agost', 'Setembre', 'Octubre', 'Novembre', 'Desembre']
        for i, (mes, prod) in enumerate(zip(mesos_noms, produccio_mensual.values)):
            row = i + 2
            ws2[f'A{row}'] = mes
            ws2[f'B{row}'] = round(prod, 0)
            fill = gris_fill if i % 2 == 0 else PatternFill("solid", fgColor='FFFFFF')
            ws2[f'A{row}'].fill = fill
            ws2[f'B{row}'].fill = fill
            ws2[f'A{row}'].border = border
            ws2[f'B{row}'].border = border
            ws2[f'B{row}'].alignment = Alignment(horizontal='center')

        ws2[f'A14'] = 'TOTAL ANUAL'
        ws2[f'B14'] = round(produccio_anual, 0)
        ws2['A14'].font = Font(bold=True, color=blau_fosc)
        ws2['B14'].font = Font(bold=True, color=blau_fosc)
        ws2['A14'].fill = PatternFill("solid", fgColor=groc)
        ws2['B14'].fill = PatternFill("solid", fgColor=groc)
        ws2['A14'].border = border
        ws2['B14'].border = border

        # Graf de barres mensual
        chart_mensual = BarChart()
        chart_mensual.title = "Produccio Energetica Mensual (kWh)"
        chart_mensual.style = 10
        chart_mensual.y_axis.title = "kWh"
        chart_mensual.x_axis.title = "Mes"
        chart_mensual.shape = 4
        data_chart = Reference(ws2, min_col=2, min_row=1, max_row=13)
        cats = Reference(ws2, min_col=1, min_row=2, max_row=13)
        chart_mensual.add_data(data_chart, titles_from_data=True)
        chart_mensual.set_categories(cats)
        chart_mensual.shape = 4
        chart_mensual.width = 20
        chart_mensual.height = 12
        ws2.add_chart(chart_mensual, "D2")

        # ---- FULL 3: CORBES HORARIES ----
        ws3 = wb.create_sheet("Corbes Horaries")
        ws3.column_dimensions['A'].width = 10

        # Capçaleres
        ws3['A1'] = 'HORA'
        ws3['A1'].font = header_font
        ws3['A1'].fill = header_fill_blau
        ws3['A1'].border = border
        ws3['A1'].alignment = Alignment(horizontal='center')

        mesos_corbes = [
            ('Gener (15 gen)', f'{any_ref}-01-15'),
            ('Abril (15 abr)', f'{any_ref}-04-15'),
            ('Juliol (15 jul)', f'{any_ref}-07-15'),
            ('Octubre (15 oct)', f'{any_ref}-10-15'),
            ('Estiu (21 jun)', f'{any_ref}-06-21'),
            ('Hivern (21 des)', f'{any_ref}-12-21'),
        ]
        cols_letters = ['B', 'C', 'D', 'E', 'F', 'G']

        for col_letter, (mes_nom, _) in zip(cols_letters, mesos_corbes):
            ws3.column_dimensions[col_letter].width = 18
            ws3[f'{col_letter}1'] = mes_nom
            ws3[f'{col_letter}1'].font = header_font
            ws3[f'{col_letter}1'].fill = header_fill_blau
            ws3[f'{col_letter}1'].border = border
            ws3[f'{col_letter}1'].alignment = Alignment(horizontal='center')

        # Dades horàries
        for h in range(24):
            row = h + 2
            ws3[f'A{row}'] = f'{h:02d}:00'
            ws3[f'A{row}'].alignment = Alignment(horizontal='center')
            ws3[f'A{row}'].border = border
            fill = gris_fill if h % 2 == 0 else PatternFill("solid", fgColor='FFFFFF')
            ws3[f'A{row}'].fill = fill

            for col_letter, (_, data) in zip(cols_letters, mesos_corbes):
                try:
                    dia_data = df.loc[data, 'P'] / 1000
                    vals_hora = dia_data[dia_data.index.hour == h]
                    val = round(float(vals_hora.mean()), 2) if len(vals_hora) > 0 else 0
                except:
                    val = 0
                ws3[f'{col_letter}{row}'] = val
                ws3[f'{col_letter}{row}'].alignment = Alignment(horizontal='center')
                ws3[f'{col_letter}{row}'].border = border
                ws3[f'{col_letter}{row}'].fill = fill

        # Graf de línies horàries
        chart_horari = LineChart()
        chart_horari.title = "Corbes de Generacio Horaria (kW)"
        chart_horari.style = 10
        chart_horari.y_axis.title = "Potencia (kW)"
        chart_horari.x_axis.title = "Hora del dia"
        chart_horari.width = 25
        chart_horari.height = 15

        for i, col_letter in enumerate(cols_letters):
            col_num = ord(col_letter) - ord('A') + 1
            data_ref = Reference(ws3, min_col=col_num, min_row=1, max_row=25)
            chart_horari.add_data(data_ref, titles_from_data=True)

        cats_horari = Reference(ws3, min_col=1, min_row=2, max_row=25)
        chart_horari.set_categories(cats_horari)
        ws3.add_chart(chart_horari, "I2")

        # ---- FULL 4: PRESSUPOST ----
        ws4 = wb.create_sheet("Pressupost")
        ws4.column_dimensions['A'].width = 40
        ws4.column_dimensions['B'].width = 15
        ws4.column_dimensions['C'].width = 20

        ws4['A1'] = 'CONCEPTE'
        ws4['B1'] = 'UNITATS'
        ws4['C1'] = 'COST (EUR)'
        for col in ['A1', 'B1', 'C1']:
            ws4[col].font = header_font
            ws4[col].fill = header_fill_blau
            ws4[col].alignment = Alignment(horizontal='center')
            ws4[col].border = border

        for i, item in enumerate(pressupost_final):
            row = i + 2
            ws4[f'A{row}'] = item['concepte']
            ws4[f'B{row}'] = str(item['unitats'])
            ws4[f'C{row}'] = item['cost']
            fill = gris_fill if i % 2 == 0 else PatternFill("solid", fgColor='FFFFFF')
            for col in [f'A{row}', f'B{row}', f'C{row}']:
                ws4[col].fill = fill
                ws4[col].border = border
            ws4[f'B{row}'].alignment = Alignment(horizontal='center')
            ws4[f'C{row}'].alignment = Alignment(horizontal='right')
            ws4[f'C{row}'].number_format = '#,##0.00 €'

        total_row = len(pressupost_final) + 2
        ws4[f'A{total_row}'] = 'TOTAL (sense IVA)'
        ws4[f'C{total_row}'] = total_sense_iva
        ws4[f'A{total_row}'].font = Font(bold=True, color=blau_fosc)
        ws4[f'C{total_row}'].font = Font(bold=True, color=blau_fosc)
        ws4[f'C{total_row}'].number_format = '#,##0.00 €'
        ws4[f'C{total_row}'].alignment = Alignment(horizontal='right')
        for col in [f'A{total_row}', f'B{total_row}', f'C{total_row}']:
            ws4[col].fill = PatternFill("solid", fgColor=groc)
            ws4[col].border = border

        total_iva_row = total_row + 1
        ws4[f'A{total_iva_row}'] = f'TOTAL (amb IVA {iva}%)'
        ws4[f'C{total_iva_row}'] = total_amb_iva
        ws4[f'A{total_iva_row}'].font = Font(bold=True, color='FFFFFF')
        ws4[f'C{total_iva_row}'].font = Font(bold=True, color='FFFFFF')
        ws4[f'C{total_iva_row}'].number_format = '#,##0.00 €'
        ws4[f'C{total_iva_row}'].alignment = Alignment(horizontal='right')
        for col in [f'A{total_iva_row}', f'B{total_iva_row}', f'C{total_iva_row}']:
            ws4[col].fill = header_fill_blau
            ws4[col].border = border

        # ---- FULL 5: ROI ----
        ws5 = wb.create_sheet("ROI")
        ws5.column_dimensions['A'].width = 15
        ws5.column_dimensions['B'].width = 25
        ws5.column_dimensions['C'].width = 25

        ws5['A1'] = 'ANY'
        ws5['B1'] = 'FLUX ANUAL (EUR)'
        ws5['C1'] = 'FLUX ACUMULAT (EUR)'
        for col in ['A1', 'B1', 'C1']:
            ws5[col].font = header_font
            ws5[col].fill = header_fill_blau
            ws5[col].alignment = Alignment(horizontal='center')
            ws5[col].border = border

        for i in range(26):
            row = i + 2
            flux_anual = estalvi_anual if i > 0 else -cost_amb_iva
            flux_acum = -cost_amb_iva + estalvi_anual * i
            ws5[f'A{row}'] = i
            ws5[f'B{row}'] = round(flux_anual, 0)
            ws5[f'C{row}'] = round(flux_acum, 0)
            fill = gris_fill if i % 2 == 0 else PatternFill("solid", fgColor='FFFFFF')
            for col in [f'A{row}', f'B{row}', f'C{row}']:
                ws5[col].fill = fill
                ws5[col].border = border
                ws5[col].alignment = Alignment(horizontal='center')
            ws5[f'B{row}'].number_format = '#,##0 €'
            ws5[f'C{row}'].number_format = '#,##0 €'

        chart_roi = LineChart()
        chart_roi.title = "Retorn de la Inversio (ROI)"
        chart_roi.style = 10
        chart_roi.y_axis.title = "EUR"
        chart_roi.x_axis.title = "Anys"
        chart_roi.width = 20
        chart_roi.height = 12
        data_roi = Reference(ws5, min_col=3, min_row=1, max_row=27)
        chart_roi.add_data(data_roi, titles_from_data=True)
        cats_roi = Reference(ws5, min_col=1, min_row=2, max_row=27)
        chart_roi.set_categories(cats_roi)
        ws5.add_chart(chart_roi, "E2")

        excel_buffer = io.BytesIO()
        wb.save(excel_buffer)
        excel_buffer.seek(0)

    # =====================
    # BOTONS DESCÀRREGA
    # =====================
    st.markdown("---")
    st.markdown("### Descarrega l'informe")
    col1, col2, col3 = st.columns(3)

    nom_fitxer = num_expedient if num_expedient else nom_projecte.replace(' ', '_')

    with col1:
        with open("informe_energydatagp.pdf", "rb") as f:
            st.download_button(
                label="Descarregar PDF",
                data=f,
                file_name=f"informe_{nom_fitxer}.pdf",
                mime="application/pdf",
                use_container_width=True
            )
    with col2:
        st.download_button(
            label="Descarregar Word (.docx)",
            data=word_buffer,
            file_name=f"informe_{nom_fitxer}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
    with col3:
        st.download_button(
            label="Descarregar Excel (.xlsx)",
            data=excel_buffer,
            file_name=f"dades_{nom_fitxer}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )