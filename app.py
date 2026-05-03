import streamlit as st
import requests
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from datetime import datetime

MESOS_CA = {
    'January': 'Gener', 'February': 'Febrer', 'March': 'Març',
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

st.set_page_config(
    page_title="EnergyDataGP — Informes Tècnics FV",
    page_icon="⚡",
    layout="wide"
)

st.markdown("""
<style>
    .header-bar {
        background: linear-gradient(90deg, #1A3A5C 0%, #2E86AB 100%);
        padding: 1.2rem 2rem;
        border-radius: 8px;
        margin-bottom: 1rem;
    }
    .header-title {
        color: white;
        font-size: 1.6rem;
        font-weight: bold;
        margin: 0;
        letter-spacing: 2px;
    }
    .header-sub {
        color: #F4A923;
        font-size: 0.85rem;
        margin: 0;
        margin-top: 4px;
    }
    .avis {
        background: #FFF8E1;
        border-left: 4px solid #F4A923;
        padding: 0.8rem 1rem;
        border-radius: 4px;
        font-size: 0.82rem;
        color: #555;
        margin-bottom: 1rem;
    }
</style>
<div class="header-bar">
    <p class="header-title">⚡ ENERGYDATAGP</p>
    <p class="header-sub">Sistema de Generació d'Informes Tècnics de Producció Fotovoltaica</p>
</div>
<div class="avis">
    ⚠️ <b>Avís important:</b> Els informes generats per aquesta eina tenen caràcter orientatiu i no substitueixen
    la signatura d'un tècnic col·legiat competent. Les dades provenen de PVGIS (Comissió Europea).
    EnergyDataGP no es fa responsable de les decisions preses sense validació professional prèvia.
</div>
""", unsafe_allow_html=True)

st.markdown("---")

col1, col2, col3 = st.columns(3)

with col1:
    st.markdown("**1. DADES DEL PROJECTE**")
    nom_projecte = st.text_input("Denominació del projecte", "")
    num_expedient = st.text_input("Número de referència intern", "")
    nom_promotor = st.text_input("Promotor / Titular", "")
    nom_tecnic = st.text_input("Tècnic responsable", "")
    num_collegio = st.text_input("Número de col·legiació", "")

with col2:
    st.markdown("**2. PARÀMETRES TÈCNICS**")
    lat = st.number_input("Latitud (°N)", value=41.54, format="%.4f")
    lon = st.number_input("Longitud (°E)", value=2.45, format="%.4f")
    adreca = st.text_input("Adreça de la instal·lació", "")
    potencia = st.number_input("Potència pic instal·lada (kWp)", value=100)
    perdues = st.number_input("Pèrdues del sistema (%)", value=14)

with col3:
    st.markdown("**3. PARÀMETRES ECONÒMICS**")
    preu_kwh = st.number_input("Preu energia (€/kWh)", value=0.18, format="%.3f")
    cost_instalacio = st.number_input("Cost instal·lació (€)", value=85000)
    any_ref = st.selectbox("Any de referència meteorològic", [2020, 2019, 2018, 2017])
    factor_co2 = st.number_input("Factor d'emissió CO₂ (kg CO₂/kWh)", value=0.25, format="%.3f")

st.markdown("---")

with st.expander("📋 Avís legal i condicions d'ús"):
    st.markdown("""
    **AVÍS LEGAL — EnergyDataGP**

    **1. Naturalesa de l'eina**
    EnergyDataGP és una eina de suport tècnic per a la generació d'informes orientatius de producció
    fotovoltaica. Els documents generats no tenen valor jurídic ni tècnic oficial per si mateixos i
    no substitueixen en cap cas la memòria tècnica signada per un enginyer col·legiat competent.

    **2. Font de les dades meteorològiques**
    Les dades provenen exclusivament de PVGIS (Photovoltaic Geographical Information System),
    desenvolupada per la Comissió Europea — Joint Research Centre (JRC).

    **3. Limitació de responsabilitat**
    EnergyDataGP no assumeix cap responsabilitat pels errors o inexactituds dels informes generats.
    L'usuari és l'únic responsable de verificar la informació amb un professional habilitat.

    **4. Acceptació de les condicions**
    L'ús d'aquesta eina implica l'acceptació expressa de totes les condicions descrites.
    """)

if st.button("⚡ Generar Informe Tècnic", type="primary", use_container_width=True):

    if not nom_projecte or not nom_tecnic:
        st.error("⚠️ Cal omplir com a mínim la denominació del projecte i el tècnic responsable.")
        st.stop()

    with st.spinner("Connectant amb PVGIS — Comissió Europea..."):
        url = "https://re.jrc.ec.europa.eu/api/v5_2/seriescalc"
        params = {
            "lat": lat, "lon": lon,
            "startyear": any_ref, "endyear": any_ref,
            "pvcalculation": 1, "peakpower": potencia,
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
    estalvi_anual = produccio_anual * preu_kwh
    anys_retorn = cost_instalacio / estalvi_anual
    co2_estalviat = produccio_anual * factor_co2

    with st.spinner("Generant gràfiques tècniques..."):
        CP = '#1A3A5C'
        CA = '#2E86AB'

        fig1, ax1 = plt.subplots(figsize=(12, 4))
        ax1.bar(range(12), produccio_mensual.values, color=CP, alpha=0.85, width=0.6)
        ax1.set_xticks(range(12))
        ax1.set_xticklabels(['Gen','Feb','Mar','Abr','Mai','Jun',
                             'Jul','Ago','Set','Oct','Nov','Des'], fontsize=10)
        ax1.set_ylabel('Energia produïda (kWh)', fontsize=10)
        ax1.set_title(f'Figura 1. Producció energètica mensual — {nom_projecte}',
                     fontsize=11, fontweight='bold', color=CP)
        ax1.axhline(produccio_mensual.mean(), color=CA, linestyle='--',
                   linewidth=1.5, label='Mitjana mensual')
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
        ax2.plot(range(len(dia_estiu)), dia_estiu.values,
                color=CA, linewidth=2.5, label="Solstici d'estiu (21 juny)")
        ax2.plot(range(len(dia_hivern)), dia_hivern.values,
                color=CP, linewidth=2.5, label="Solstici d'hivern (21 desembre)")
        ax2.fill_between(range(len(dia_estiu)), dia_estiu.values, alpha=0.1, color=CA)
        ax2.fill_between(range(len(dia_hivern)), dia_hivern.values, alpha=0.1, color=CP)
        ax2.set_ylabel('Potència generada (kW)', fontsize=10)
        ax2.set_title('Figura 2. Corba de generació horària — Comparativa estacional',
                     fontsize=11, fontweight='bold', color=CP)
        ax2.legend(fontsize=9)
        ax2.spines['top'].set_visible(False)
        ax2.spines['right'].set_visible(False)
        ax2.grid(alpha=0.3)
        plt.tight_layout()
        plt.savefig('g2_corbes.png', dpi=180, bbox_inches='tight')
        plt.close()

        anys = list(range(0, 26))
        flux_acumulat = [-cost_instalacio + estalvi_anual * a for a in anys]
        fig3, ax3 = plt.subplots(figsize=(12, 4))
        colors_bars = [CA if v < 0 else '#2E7D32' for v in flux_acumulat]
        ax3.bar(anys, flux_acumulat, color=colors_bars, alpha=0.8, width=0.7)
        ax3.axhline(0, color='black', linewidth=1)
        ax3.set_xlabel('Anys des de la posada en marxa', fontsize=10)
        ax3.set_ylabel('Flux de caixa acumulat (€)', fontsize=10)
        ax3.set_title('Figura 3. Anàlisi del retorn de la inversió (ROI)',
                     fontsize=11, fontweight='bold', color=CP)
        ax3.spines['top'].set_visible(False)
        ax3.spines['right'].set_visible(False)
        ax3.grid(axis='y', alpha=0.3)
        llegenda = [mpatches.Patch(color=CA, label='Període de retorn'),
                   mpatches.Patch(color='#2E7D32', label='Benefici net')]
        ax3.legend(handles=llegenda, fontsize=9)
        plt.tight_layout()
        plt.savefig('g3_roi.png', dpi=180, bbox_inches='tight')
        plt.close()

    st.success("✅ Informe generat correctament")
    st.markdown("### Resum executiu")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Producció anual", f"{produccio_anual:,.0f} kWh")
    c2.metric("Estalvi econòmic anual", f"{estalvi_anual:,.0f} €")
    c3.metric("Retorn de la inversió", f"{anys_retorn:.1f} anys")
    c4.metric("CO₂ estalviat/any", f"{co2_estalviat:,.0f} kg")

    with st.spinner("Generant conclusions amb IA (Groq)..."):
        try:
            from groq import Groq
            client = Groq(api_key=st.secrets["GROQ_API_KEY"])
            prompt = f"""Ets un enginyer expert en energia solar fotovoltaica.
Redacta les conclusions tècniques d'un informe professional en català,
formal i precís, de màxim 200 paraules, a partir d'aquestes dades reals:
- Projecte: {nom_projecte}
- Ubicació: {lat}°N, {lon}°E
- Potència instal·lada: {potencia} kWp
- Producció anual estimada: {produccio_anual:,.0f} kWh
- Millor mes: {mes_max} ({produccio_mensual.max():,.0f} kWh)
- Pitjor mes: {mes_min} ({produccio_mensual.min():,.0f} kWh)
- Cost instal·lació: {cost_instalacio:,.0f} €
- Estalvi econòmic anual: {estalvi_anual:,.0f} €
- Retorn de la inversió: {anys_retorn:.1f} anys
- CO₂ estalviat: {co2_estalviat:,.0f} kg/any
No incloguis títols ni encapçalaments. Redacta directament el text de conclusions."""
            resposta = client.chat.completions.create(
                messages=[{"role": "user", "content": prompt}],
                model="llama3-70b-8192",
            )
            text_conclusions = resposta.choices[0].message.content
            st.success("✅ Conclusions generades per IA")
        except Exception as e:
            st.warning(f"⚠️ IA no disponible ({e}). S'usa text estàndard.")
            text_conclusions = (
                f"La instal·lació fotovoltaica {nom_projecte} presenta una producció anual estimada de "
                f"{produccio_anual:,.0f} kWh, un estalvi econòmic de {estalvi_anual:,.0f} €/any i un "
                f"retorn de la inversió en {anys_retorn:.1f} anys. La instal·lació contribuirà a la "
                f"reducció d'emissions de CO₂ en {co2_estalviat:,.0f} kg per any."
            )

    with st.spinner("Composant el document PDF..."):
        doc = SimpleDocTemplate(
            "informe_energydatagp.pdf",
            pagesize=A4,
            rightMargin=2.5*cm, leftMargin=2.5*cm,
            topMargin=2*cm, bottomMargin=2*cm
        )

        styles = getSampleStyleSheet()

        estil_seccio = ParagraphStyle('seccio',
            fontSize=11, fontName='Helvetica-Bold',
            textColor=COLOR_PRINCIPAL,
            spaceBefore=12, spaceAfter=6)

        estil_normal = ParagraphStyle('normal',
            fontSize=9, fontName='Helvetica',
            textColor=COLOR_NEGRE, leading=14, spaceAfter=4)

        estil_peu = ParagraphStyle('peu',
            fontSize=7, fontName='Helvetica',
            textColor=colors.grey, alignment=TA_CENTER)

        estil_avis = ParagraphStyle('avis',
            fontSize=8, fontName='Helvetica',
            textColor=colors.grey, leading=12,
            spaceAfter=8)

        elements = []

        elements.append(HRFlowable(width="100%", thickness=4,
                                   color=COLOR_PRINCIPAL, spaceAfter=8))
        cap_data = [[
            Paragraph('<b><font size=14>⚡ ENERGYDATAGP</font></b><br/>'
                     '<font size=8 color=grey>Solucions d\'Anàlisi de Dades Energètiques</font>',
                     ParagraphStyle('cap', fontSize=9, fontName='Helvetica-Bold',
                                   textColor=COLOR_PRINCIPAL)),
            Paragraph('<b>INFORME TÈCNIC DE PRODUCCIÓ</b><br/>INSTAL·LACIÓ FOTOVOLTAICA',
                     ParagraphStyle('cap2', fontSize=12, fontName='Helvetica-Bold',
                                   textColor=COLOR_PRINCIPAL, alignment=TA_CENTER)),
            Paragraph(f'Ref.: <b>{num_expedient if num_expedient else "—"}</b><br/>'
                     f'Data: <b>{datetime.now().strftime("%d/%m/%Y")}</b><br/>'
                     f'Versió: 1.0',
                     ParagraphStyle('cap3', fontSize=9, fontName='Helvetica',
                                   textColor=COLOR_PRINCIPAL, alignment=TA_RIGHT)),
        ]]
        taula_cap = Table(cap_data, colWidths=[5.5*cm, 8*cm, 4*cm])
        taula_cap.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('BACKGROUND', (0,0), (-1,-1), COLOR_GRIS_CLAR),
            ('PADDING', (0,0), (-1,-1), 8),
        ]))
        elements.append(taula_cap)
        elements.append(HRFlowable(width="100%", thickness=2,
                                   color=COLOR_GROC, spaceAfter=12))

        elements.append(Paragraph("1. IDENTIFICACIÓ DEL PROJECTE", estil_seccio))
        elements.append(HRFlowable(width="100%", thickness=0.5,
                                   color=COLOR_GRIS, spaceAfter=6))
        dades_id = [
            ['Denominació:', nom_projecte, 'Promotor / Titular:', nom_promotor if nom_promotor else '—'],
            ['Referència:', num_expedient if num_expedient else '—', 'Adreça:', adreca if adreca else '—'],
            ['Coordenades:', f'{lat}°N, {lon}°E', 'Any meteorològic ref.:', str(any_ref)],
            ['Tècnic responsable:', nom_tecnic, 'Núm. col·legiació:', num_collegio if num_collegio else '—'],
        ]
        taula_id = Table(dades_id, colWidths=[4*cm, 6.5*cm, 3.5*cm, 4.5*cm])
        taula_id.setStyle(TableStyle([
            ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold'),
            ('FONTNAME', (2,0), (2,-1), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,-1), 9),
            ('TEXTCOLOR', (0,0), (0,-1), COLOR_PRINCIPAL),
            ('TEXTCOLOR', (2,0), (2,-1), COLOR_PRINCIPAL),
            ('ROWBACKGROUNDS', (0,0), (-1,-1), [COLOR_GRIS_CLAR, colors.white]),
            ('GRID', (0,0), (-1,-1), 0.3, COLOR_GRIS),
            ('PADDING', (0,0), (-1,-1), 5),
        ]))
        elements.append(taula_id)
        elements.append(Spacer(1, 0.4*cm))

        elements.append(Paragraph("2. PARÀMETRES TÈCNICS DE LA INSTAL·LACIÓ", estil_seccio))
        elements.append(HRFlowable(width="100%", thickness=0.5,
                                   color=COLOR_GRIS, spaceAfter=6))
        dades_tec = [
            ['Paràmetre', 'Valor', 'Unitat', 'Observacions'],
            ['Potència pic instal·lada', f'{potencia}', 'kWp', 'Condicions estàndard (STC)'],
            ['Pèrdues del sistema', f'{perdues}', '%', 'Cablejat, inversor, brutícia, temperatura'],
            ['Producció anual estimada', f'{produccio_anual:,.0f}', 'kWh/any', f'Any de referència: {any_ref}'],
            ['Mes de màxima producció', mes_max, f'{produccio_mensual.max():,.0f} kWh', 'Pic de producció estival'],
            ['Mes de mínima producció', mes_min, f'{produccio_mensual.min():,.0f} kWh', 'Mínim de producció hivernal'],
            ['Hores equivalents de sol', f'{produccio_anual/potencia:,.0f}', 'HES/any', 'Hores a potència nominal'],
        ]
        taula_tec = Table(dades_tec, colWidths=[5.5*cm, 3*cm, 3*cm, 6*cm])
        taula_tec.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), COLOR_PRINCIPAL),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,-1), 9),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [COLOR_GRIS_CLAR, colors.white]),
            ('GRID', (0,0), (-1,-1), 0.3, COLOR_GRIS),
            ('PADDING', (0,0), (-1,-1), 5),
            ('ALIGN', (1,1), (2,-1), 'CENTER'),
        ]))
        elements.append(taula_tec)
        elements.append(Spacer(1, 0.4*cm))

        elements.append(Paragraph("3. ANÀLISI DE LA PRODUCCIÓ ENERGÈTICA MENSUAL", estil_seccio))
        elements.append(HRFlowable(width="100%", thickness=0.5,
                                   color=COLOR_GRIS, spaceAfter=6))
        elements.append(Image('g1_mensual.png', width=16*cm, height=6*cm))
        elements.append(Paragraph(
            f"La instal·lació fotovoltaica objecte d'aquest informe presenta una producció anual estimada de "
            f"<b>{produccio_anual:,.0f} kWh</b>, d'acord amb les dades meteorològiques de l'any {any_ref} "
            f"de PVGIS. El mes de màxima producció correspon a <b>{mes_max}</b> amb "
            f"<b>{produccio_mensual.max():,.0f} kWh</b>, mentre que el mes de mínima producció és "
            f"<b>{mes_min}</b> amb <b>{produccio_mensual.min():,.0f} kWh</b>.", estil_normal))
        elements.append(Spacer(1, 0.3*cm))

        elements.append(Paragraph("4. CORBA DE GENERACIÓ HORÀRIA", estil_seccio))
        elements.append(HRFlowable(width="100%", thickness=0.5,
                                   color=COLOR_GRIS, spaceAfter=6))
        elements.append(Image('g2_corbes.png', width=16*cm, height=6*cm))
        elements.append(Paragraph(
            f"La figura anterior mostra la corba de generació horària per als solsticis d'estiu i d'hivern. "
            f"La potència màxima durant el solstici d'estiu és de <b>{dia_estiu.max():.1f} kW</b>, "
            f"mentre que durant el solstici d'hivern és de <b>{dia_hivern.max():.1f} kW</b>.", estil_normal))
        elements.append(Spacer(1, 0.3*cm))

        elements.append(Paragraph("5. ANÀLISI ECONÒMICA I RETORN DE LA INVERSIÓ", estil_seccio))
        elements.append(HRFlowable(width="100%", thickness=0.5,
                                   color=COLOR_GRIS, spaceAfter=6))
        dades_eco = [
            ['Concepte', 'Valor', 'Unitat'],
            ['Cost total de la instal·lació', f'{cost_instalacio:,.0f}', '€'],
            ['Preu de l\'energia considerada', f'{preu_kwh:.3f}', '€/kWh'],
            ['Estalvi econòmic anual estimat', f'{estalvi_anual:,.0f}', '€/any'],
            ['Període de retorn de la inversió', f'{anys_retorn:.1f}', 'anys'],
            ['Emissions de CO₂ estalviades', f'{co2_estalviat:,.0f}', 'kg CO₂/any'],
        ]
        taula_eco = Table(dades_eco, colWidths=[9*cm, 4*cm, 4.5*cm])
        taula_eco.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), COLOR_ACCENT),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,-1), 9),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [COLOR_GRIS_CLAR, colors.white]),
            ('GRID', (0,0), (-1,-1), 0.3, COLOR_GRIS),
            ('PADDING', (0,0), (-1,-1), 5),
            ('ALIGN', (1,1), (2,-1), 'CENTER'),
        ]))
        elements.append(taula_eco)
        elements.append(Spacer(1, 0.3*cm))
        elements.append(Image('g3_roi.png', width=16*cm, height=5.5*cm))
        elements.append(Spacer(1, 0.3*cm))

        elements.append(Paragraph("6. CONCLUSIONS", estil_seccio))
        elements.append(HRFlowable(width="100%", thickness=0.5,
                                   color=COLOR_GRIS, spaceAfter=6))
        elements.append(Paragraph(text_conclusions, estil_normal))
        elements.append(Spacer(1, 0.8*cm))

        sig_data = [
            [Paragraph('El/La tècnic/a responsable', estil_normal),
             Paragraph('Conforme el/la promotor/a', estil_normal)],
            [Paragraph(f'<br/><br/><br/>Nom: {nom_tecnic}<br/>'
                      f'Núm. col·legiació: {num_collegio if num_collegio else "—"}<br/>'
                      f'Data: {datetime.now().strftime("%d/%m/%Y")}', estil_normal),
             Paragraph('<br/><br/><br/>Nom:<br/>DNI/NIF:<br/>Data:', estil_normal)],
        ]
        taula_sig = Table(sig_data, colWidths=[8.5*cm, 8.5*cm])
        taula_sig.setStyle(TableStyle([
            ('BOX', (0,0), (0,-1), 0.5, COLOR_GRIS),
            ('BOX', (1,0), (1,-1), 0.5, COLOR_GRIS),
            ('PADDING', (0,0), (-1,-1), 8),
        ]))
        elements.append(taula_sig)
        elements.append(Spacer(1, 0.5*cm))

        elements.append(HRFlowable(width="100%", thickness=0.5,
                                   color=COLOR_GRIS, spaceAfter=4))
        elements.append(Paragraph(
            "<b>AVÍS LEGAL:</b> Document generat amb EnergyDataGP amb finalitat orientativa. "
            "Dades de producció: PVGIS © Comissió Europea. No substitueix la memòria tècnica "
            "oficial signada per un tècnic col·legiat competent.", estil_avis))
        elements.append(HRFlowable(width="100%", thickness=2,
                                   color=COLOR_PRINCIPAL, spaceAfter=4))
        elements.append(Paragraph(
            f"EnergyDataGP · Informe generat el {datetime.now().strftime('%d/%m/%Y a les %H:%M')} · "
            f"Dades: PVGIS © Comissió Europea · www.energydatagp.com",
            estil_peu))

        doc.build(elements)

    with open("informe_energydatagp.pdf", "rb") as f:
        st.download_button(
            label="📥 Descarregar Informe Tècnic PDF",
            data=f,
            file_name=f"informe_{num_expedient if num_expedient else nom_projecte.replace(' ','_')}.pdf",
            mime="application/pdf",
            use_container_width=True
        )
