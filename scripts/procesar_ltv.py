import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import os
from fpdf import FPDF
from datetime import datetime

# Configuración de Entorno
INPUT_FOLDER = 'data_pedidos'
OUTPUT_FOLDER = 'reportes'
FILES = {'2023': 'Pedidos_2023.xlsx', '2024': 'Pedidos_2024.xlsx', '2025': 'Pedidos_2025.xlsx'}
LOGO_URL = "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQfE4betnoplLem-rHmrOt2gqS7zMBYV8D3aw&s"

class LTV_Report(FPDF):
    def header(self):
        self.image(LOGO_URL, 10, 8, 25)
        self.set_font('Arial', 'B', 10)
        self.cell(0, 10, 'Reporte Ejecutivo de Customer Lifetime Value (LTV)', 0, 0, 'R')
        self.ln(15)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Página {self.page_no()} | Canal: Juntoz | Generado: {datetime.now().strftime("%Y-%m-%d")}', 0, 0, 'C')

def generar_analisis_senior():
    if not os.path.exists(OUTPUT_FOLDER): os.makedirs(OUTPUT_FOLDER)
    
    # 1. EXTRACCIÓN Y LIMPIEZA RIGUROSA
    all_data = []
    estados_validos = ['Received', 'ReadyToShip', 'ReadyToPickUp', 'PendingToPickUp', 'InTransit', 'Confirmed']
    
    for year, name in FILES.items():
        path = os.path.join(INPUT_FOLDER, name)
        if os.path.exists(path):
            df = pd.read_excel(path)
            
            # APLICACIÓN DE TODOS LOS FILTROS SOLICITADOS (Incluyendo SITIO)
            df = df[
                (df['Canal de venta'] == 'Juntoz') & 
                (df['Sitio'] == 'Juntoz') & 
                (df['Tipo de documento de cliente'] == 'DNI') &
                (df['Estado de item'].isin(estados_validos))
            ].copy()
            
            # Normalización de Fechas y Montos
            df['Fecha de creación'] = pd.to_datetime(df['Fecha de creación']).dt.normalize()
            df['Total'] = pd.to_numeric(df['Total'].astype(str).str.replace(',', '.'), errors='coerce')
            df['Año'] = year
            all_data.append(df)

    if not all_data: return print("Sin datos para procesar.")
    df_master = pd.concat(all_data, ignore_index=True)

    # 2. MÉTRICAS POR CLIENTE (CUSTOMER LEVEL)
    customers = df_master.groupby('Nro. de documento de cliente').agg(
        ltv_acumulado=('Total', 'sum'),
        total_ordenes=('Nro. de orden', 'nunique'),
        fecha_primera=('Fecha de creación', 'min'),
        fecha_ultima=('Fecha de creación', 'max')
    ).reset_index()

    # Cálculo de Antigüedad y Recurrencia
    customers['antiguedad_dias'] = (customers['fecha_ultima'] - customers['fecha_primera']).dt.days
    customers['es_recurrente'] = customers['total_ordenes'] > 1
    customers['año_cohorte'] = customers['fecha_primera'].dt.year

    # 3. SEGMENTACIÓN Y PERCENTILES
    percentiles = customers['ltv_acumulado'].quantile([0.25, 0.5, 0.75, 0.9, 0.99])
    
    # Pareto 80/20
    customers = customers.sort_values('ltv_acumulado', ascending=False)
    customers['venta_acum_perc'] = 100 * customers['ltv_acumulado'].cumsum() / customers['ltv_acumulado'].sum()
    customers['cliente_rank_perc'] = 100 * np.arange(1, len(customers) + 1) / len(customers)

    # 4. GENERACIÓN DE VISUALIZACIONES (ESTILO GERENCIAL)
    plt.style.use('seaborn-v0_8-whitegrid')
    
    # Gráfico A: Evolución Anual
    plt.figure(figsize=(8, 4))
    df_master.groupby('Año')['Total'].sum().plot(kind='bar', color='#1a237e')
    plt.title('Ingresos Netos por Año (Soles)')
    plt.savefig(f'{OUTPUT_FOLDER}/plot_ventas_año.png')

    # Gráfico B: Pareto
    plt.figure(figsize=(8, 4))
    plt.plot(customers['cliente_rank_perc'], customers['venta_acum_perc'], color='red', lw=2)
    plt.fill_between(customers['cliente_rank_perc'], customers['venta_acum_perc'], color='red', alpha=0.1)
    plt.title('Curva de Pareto: Concentración de LTV')
    plt.xlabel('% Clientes')
    plt.ylabel('% Venta Acumulada')
    plt.savefig(f'{OUTPUT_FOLDER}/plot_pareto.png')

    # 5. GENERACIÓN DEL PDF MULTIPÁGINA
    pdf = LTV_Report()
    
    # Portada
    pdf.add_page()
    pdf.set_font('Arial', 'B', 28)
    pdf.ln(60)
    pdf.cell(0, 20, 'Reporte de Customer', 0, 1, 'C')
    pdf.cell(0, 20, 'Lifetime Value (LTV)', 0, 1, 'C')
    pdf.set_font('Arial', '', 14)
    pdf.cell(0, 10, 'Canal: Juntoz | Periodo: 2023 - 2025', 0, 1, 'C')
    
    # Resumen Ejecutivo
    pdf.add_page()
    pdf.set_font('Arial', 'B', 16)
    pdf.cell(0, 10, '1. Resumen Ejecutivo (KPIs Globales)', 0, 1)
    pdf.set_font('Arial', '', 12)
    pdf.ln(5)
    
    metrics = [
        f"Clientes Únicos Totales: {len(customers):,}",
        f"Ventas Netas Totales: S/ {df_master['Total'].sum():,.2f}",
        f"LTV Promedio General: S/ {customers['ltv_acumulado'].mean():,.2f}",
        f"LTV Mediano (Percentil 50): S/ {percentiles[0.5]:,.2f}",
        f"Ticket Promedio (AOV): S/ {df_master['Total'].mean():,.2f}",
        f"Tasa de Recurrencia: {100*customers['es_recurrente'].mean():.1f}%"
    ]
    for m in metrics: pdf.cell(0, 10, f"- {m}", 0, 1)

    # Análisis Temporal
    pdf.add_page()
    pdf.set_font('Arial', 'B', 16)
    pdf.cell(0, 10, '2. Análisis Temporal y Evolución', 0, 1)
    pdf.image(f'{OUTPUT_FOLDER}/plot_ventas_año.png', x=10, w=180)
    pdf.ln(5)
    
    # Análisis de Clientes
    pdf.add_page()
    pdf.set_font('Arial', 'B', 16)
    pdf.cell(0, 10, '3. Segmentación y Pareto', 0, 1)
    pdf.image(f'{OUTPUT_FOLDER}/plot_pareto.png', x=10, w=180)
    
    top_10_perc = customers[customers['cliente_rank_perc'] <= 10]['venta_acum_perc'].max()
    pdf.ln(5)
    pdf.set_font('Arial', 'I', 11)
    pdf.multi_cell(0, 10, f"Interpretación: El Top 10% de clientes concentra el {top_10_perc:.1f}% del valor total. Es imperativo desarrollar estrategias de fidelización para este segmento VIP.")

    pdf.output(f"{OUTPUT_FOLDER}/LTV_Report_2023_2025.pdf")
    print("✅ Reporte LTV Generado con éxito.")

if __name__ == "__main__":
    generar_analisis_senior()