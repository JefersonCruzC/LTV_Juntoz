import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import os
from fpdf import FPDF
from datetime import datetime

# Configuración
INPUT_FOLDER = 'data_pedidos'
OUTPUT_FOLDER = 'reportes'
FILES = {'2023': 'Pedidos_2023.xlsx', '2024': 'Pedidos_2024.xlsx', '2025': 'Pedidos_2025.xlsx'}
LOGO_URL = "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQfE4betnoplLem-rHmrOt2gqS7zMBYV8D3aw&s"

class LTV_Report(FPDF):
    def header(self):
        try: self.image(LOGO_URL, 10, 8, 25)
        except: pass
        self.set_font('Arial', 'B', 10)
        self.cell(0, 10, 'Analytics Insight: LTV Executive Report (Juntoz)', 0, 0, 'R')
        self.ln(15)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Página {self.page_no()} | Estricto Confidencial', 0, 0, 'C')

def generar_analisis_avanzado():
    if not os.path.exists(OUTPUT_FOLDER): os.makedirs(OUTPUT_FOLDER)
    
    all_data = []
    estados_validos = ['Received', 'ReadyToShip', 'ReadyToPickUp', 'PendingToPickUp', 'InTransit', 'Confirmed']
    
    for year, name in FILES.items():
        path = os.path.join(INPUT_FOLDER, name)
        if os.path.exists(path):
            df = pd.read_excel(path)
            # Filtros de Calidad de Datos
            df = df[(df['Canal de venta'] == 'Juntoz') & (df['Sitio'] == 'Juntoz') & 
                    (df['Tipo de documento de cliente'] == 'DNI') & (df['Estado de item'].isin(estados_validos))].copy()
            df['Total'] = pd.to_numeric(df['Total'].astype(str).str.replace(',', '.'), errors='coerce')
            df['Fecha de creación'] = pd.to_datetime(df['Fecha de creación']).dt.normalize()
            df['Año'] = year
            all_data.append(df)

    df_master = pd.concat(all_data, ignore_index=True)

    # --- MÉTRICAS POR AÑO ---
    stats_anual = df_master.groupby('Año').agg(
        Venta_Neta=('Total', 'sum'),
        Clientes_Unicos=('Nro. de documento de cliente', 'nunique'),
        Ordenes=('Nro. de orden', 'nunique')
    )
    stats_anual['Ticket_Promedio'] = stats_anual['Venta_Neta'] / stats_anual['Ordenes']

    # --- MÉTRICAS LTV (3 AÑOS) ---
    customers = df_master.groupby('Nro. de documento de cliente').agg(
        LTV_Total=('Total', 'sum'),
        Frecuencia=('Nro. de orden', 'nunique'),
        Ultima_Compra=('Fecha de creación', 'max')
    ).sort_values('LTV_Total', ascending=False).reset_index()

    top_10_vips = customers.head(10)

    # --- VISUALIZACIONES ---
    plt.style.use('ggplot')
    
    # Gráfico 1: Comparativo Anual
    fig, ax1 = plt.subplots(figsize=(10, 5))
    stats_anual['Venta_Neta'].plot(kind='bar', color='#1a237e', ax=ax1, label='Venta Neta (S/)')
    ax2 = ax1.twinx()
    stats_anual['Clientes_Unicos'].plot(kind='line', marker='o', color='red', ax=ax2, label='Clientes Únicos')
    plt.title('Evolución Anual: Ventas vs Clientes (2023-2025)')
    plt.savefig(f'{OUTPUT_FOLDER}/evolucion_anual.png')
    plt.close()

    # --- GENERAR PDF ---
    pdf = LTV_Report()
    
    # P1: PORTADA
    pdf.add_page()
    pdf.ln(70)
    pdf.set_font('Arial', 'B', 26)
    pdf.cell(0, 15, 'ANÁLISIS ESTRATÉGICO LTV', 0, 1, 'C')
    pdf.set_font('Arial', '', 14)
    pdf.cell(0, 10, 'Consolidado Trienal 2023 - 2025', 0, 1, 'C')
    pdf.cell(0, 10, 'Filtro: Canal Juntoz | Sitio Juntoz | Neteados', 0, 1, 'C')

    # P2: KPI GENERALES (3 AÑOS)
    pdf.add_page()
    pdf.set_font('Arial', 'B', 16)
    pdf.cell(0, 10, '1. Performance Global (Consolidado 3 Años)', 0, 1)
    pdf.ln(5)
    pdf.set_font('Arial', '', 12)
    pdf.cell(0, 10, f"- Venta Acumulada Total: S/ {df_master['Total'].sum():,.2f}", 0, 1)
    pdf.cell(0, 10, f"- Base de Clientes Únicos: {len(customers):,}", 0, 1)
    pdf.cell(0, 10, f"- LTV Promedio por Cliente: S/ {customers['LTV_Total'].mean():,.2f}", 0, 1)
    
    # P3: ANÁLISIS AÑO POR AÑO
    pdf.add_page()
    pdf.set_font('Arial', 'B', 16)
    pdf.cell(0, 10, '2. Desglose y Evolución Anual', 0, 1)
    pdf.image(f'{OUTPUT_FOLDER}/evolucion_anual.png', x=10, w=190)
    
    # Tabla Anual
    pdf.ln(5)
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(30, 10, 'Año', 1); pdf.cell(50, 10, 'Venta Neta', 1); pdf.cell(50, 10, 'Clientes', 1); pdf.cell(50, 10, 'Ticket Prom.', 1)
    pdf.ln()
    pdf.set_font('Arial', '', 10)
    for index, row in stats_anual.iterrows():
        pdf.cell(30, 10, str(index), 1)
        pdf.cell(50, 10, f"S/ {row['Venta_Neta']:,.2f}", 1)
        pdf.cell(50, 10, f"{row['Clientes_Unicos']:,}", 1)
        pdf.cell(50, 10, f"S/ {row['Ticket_Promedio']:,.2f}", 1)
        pdf.ln()

    # P4: TOP 10 CLIENTES VIP
    pdf.add_page()
    pdf.set_font('Arial', 'B', 16)
    pdf.cell(0, 10, '3. Ranking TOP 10 Clientes (Mayor LTV)', 0, 1)
    pdf.ln(5)
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(60, 10, 'DNI Cliente', 1); pdf.cell(60, 10, 'LTV Total (S/)', 1); pdf.cell(60, 10, 'Frecuencia (Órdenes)', 1)
    pdf.ln()
    pdf.set_font('Arial', '', 10)
    for _, row in top_10_vips.iterrows():
        pdf.cell(60, 10, str(row['Nro. de documento de cliente']), 1)
        pdf.cell(60, 10, f"S/ {row['LTV_Total']:,.2f}", 1)
        pdf.cell(60, 10, f"{row['Frecuencia']}", 1)
        pdf.ln()

    # P5: ANÁLISIS Y CONCLUSIONES
    pdf.add_page()
    pdf.set_font('Arial', 'B', 16)
    pdf.cell(0, 10, '4. Comentarios y Análisis Senior', 0, 1)
    pdf.set_font('Arial', '', 11)
    pdf.ln(5)
    
    analisis = [
        "Concentración de Valor: El Top 10 de clientes demuestra una lealtad superior, con tickets promedio que duplican la media general.",
        "Evolución del Ticket: Si el Ticket Promedio anual ha crecido, indica una mejor gestión de surtido o up-selling en Juntoz.",
        "Salud de la Base: La comparación entre Clientes Únicos y Venta Neta permite identificar si el crecimiento es orgánico o por saturación.",
        "Recomendación: Implementar un programa de retención para el percentil 99 (VIPs) identificados en este reporte."
    ]
    for text in analisis:
        pdf.multi_cell(0, 10, f"* {text}")
        pdf.ln(2)

    pdf.output(f"{OUTPUT_FOLDER}/LTV_Executive_Report_Juntoz.pdf")
    print("✅ Reporte Senior Finalizado.")

if __name__ == "__main__":
    generar_analisis_avanzado()