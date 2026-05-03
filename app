import streamlit as st
import requests
import pandas as pd
import matplotlib.pyplot as plt
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
import io
import os

st.set_page_config(page_title="Generador d'Informes Solars", page_icon="☀️")

st.title("☀️ Generador d'Informes de Producció Solar")
st.markdown("Introdueix les dades de la planta i genera l'informe automàticament.")

# FORMULARI
col1, col2 = st.columns(2)
with col1:
    nom_projecte = st.text_input("Nom del projecte", "Planta FV Mataró")
    nom_client = st.text_input("Client", "Client Example SL")
    lat = st.number_input("Latitud", value=41.54, format="%.4f")
    lon = st.number_input("Longitud", value=2.45, format="%.4f")

with col2:
    potencia = st.number_input("Potència (kWp)", value=100)
    perdues = st.number_input("Pèrdues sistema (%)", value=14)
    any_ref = st.selectbox("Any de referència", [2020, 2019, 2018, 2017])

if st.button("🔄 Generar Informe", type="primary"):
    with st.spinner("Descarregant dades de PVGIS..."):
        
        # Descarrega dades
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

    with st.spinner("Generant gràfiques..."):

        # Gràfica mensual
        fig1, ax1 = plt.subplots(figsize=(10, 4))
        ax1.bar(range(12), produccio_mensual.values, color='#378ADD', alpha=0.8)
        ax1.set_xticks(range(12))
        ax1.set_xticklabels(['Gen','Feb','Mar','Abr','Mai','Jun','Jul','Ago','Set','Oct','Nov','Des'])
        ax1.set_title('Producció mensual (kWh)')
        ax1.set_ylabel('kWh')
        for i, val in enumerate(produccio_mensual.values):
            ax1.text(i, val + 200, f'{val:.0f}', ha='center', fontsize=8)
        plt.tight_layout()
        plt.savefig('grafica_mensual.png', dpi=150, bbox_inches='tight')
        plt.close()

        # Gràfica corbes
        dia_estiu = df.loc[f'{any_ref}-06-21', 'P'] / 1000
        dia_hivern = df.loc[f'{any_ref}-12-21', 'P'] / 1000
        fig2, ax2 = plt.subplots(figsize=(10, 4))
        ax2.plot(range(len(dia_estiu)), dia_estiu.values, color='#F59F27', linewidth=2, label='Estiu')
        ax2.plot(range(len(dia_hivern)), dia_hivern.values, color='#378ADD', linewidth=2, label='Hivern')
        ax2.set_title('Corba de producció diària')
        ax2.set_ylabel('Potència (kW)')
        ax2.legend()
        ax2.grid(True, alpha=0.3)
        plt.tight_layout()
        plt.savefig('grafica_corbes.png', dpi=150, bbox_inches='tight')
        plt.close()

    # Mostra resultats a la web
    st.success("✅ Dades carregades!")
    col1, col2, col3 = st.columns(3)
    col1.metric("Producció anual", f"{produccio_mensual.sum():.0f} kWh")
    col2.metric("Millor mes", f"{produccio_mensual.max():.0f} kWh")
    col3.metric("Pitjor mes", f"{produccio_mensual.min():.0f} kWh")
    st.pyplot(fig1)

    with st.spinner("Generant PDF..."):

        # Genera PDF
        doc = SimpleDocTemplate(
            "informe.pdf", pagesize=A4,
            rightMargin=2*cm, leftMargin=2*cm,
            topMargin=2*cm, bottomMargin=2*cm
        )
        styles = getSampleStyleSheet()
        elements = []

        elements.append(Paragraph(f"INFORME DE PRODUCCIÓ SOLAR", styles['Title']))
        elements.append(Paragraph(f"{nom_projecte}", styles['Heading2']))
        elements.append(Paragraph(f"Client: {nom_client} · Any: {any_ref} · Dades: PVGIS", styles['Normal']))
        elements.append(Spacer(1, 0.5*cm))

        data_taula = [
            ['Paràmetre', 'Valor'],
            ['Ubicació', f'{lat}°N, {lon}°E'],
            ['Potència instal·lada', f'{potencia} kWp'],
            ['Pèrdues sistema', f'{perdues}%'],
            ['Producció anual total', f'{produccio_mensual.sum():.0f} kWh'],
            ['Millor mes', f'{produccio_mensual.idxmax().strftime("%B")} ({produccio_mensual.max():.0f} kWh)'],
            ['Pitjor mes', f'{produccio_mensual.idxmin().strftime("%B")} ({produccio_mensual.min():.0f} kWh)'],
            ['Hores de sol anuals', f'{(df["P"] > 0).sum()} h'],
        ]
        taula = Table(data_taula, colWidths=[8*cm, 8*cm])
        taula.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#378ADD')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,-1), 10),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.HexColor('#F5F5F5'), colors.white]),
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ('PADDING', (0,0), (-1,-1), 6),
        ]))
        elements.append(taula)
        elements.append(Spacer(1, 0.5*cm))
        elements.append(Paragraph("Producció mensual", styles['Heading3']))
        elements.append(Image('grafica_mensual.png', width=16*cm, height=6.5*cm))
        elements.append(Spacer(1, 0.3*cm))
        elements.append(Paragraph("Corba de producció diària", styles['Heading3']))
        elements.append(Image('grafica_corbes.png', width=16*cm, height=6.5*cm))
        elements.append(Spacer(1, 0.5*cm))
        elements.append(Paragraph(f"Informe generat automàticament amb Python · Dades PVGIS © Comissió Europea", styles['Normal']))
        doc.build(elements)

    # Botó de descàrrega
    with open("informe.pdf", "rb") as f:
        st.download_button(
            label="📥 Descarregar PDF",
            data=f,
            file_name=f"informe_{nom_projecte.replace(' ','_')}.pdf",
            mime="application/pdf"
        )
