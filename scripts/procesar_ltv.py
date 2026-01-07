import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import os
from fpdf import FPDF
from datetime import datetime

# Configuración de Rutas
INPUT_FOLDER = 'data_pedidos'
OUTPUT_FOLDER = 'reportes'
LOGO_URL = "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQfE4betnoplLem-rHmrOt2gqS7zMBYV8D3aw&s"

class LTV_Report(FPDF):
    def header(self):
        if self.page_no() > 1:
            self.image(LOGO_URL, 10, 8, 33)
            self.set_font('Arial', 'I', 8)
            self.cell(0, 10, f'LTV Report Juntoz 2023-2025 - Página {self.page_no()}', 0, 0, 'R')
            self.ln(20)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Generado el {datetime.now().strftime("%Y-%m-%d %H:%M")}', 0, 0, 'C')

def generar_analisis():
    if not os.path.exists(OUTPUT_FOLDER): os.makedirs(OUTPUT_FOLDER)
    
    # 1. CARGA Y FILTRADO
    files = {'2023': 'Pedidos_2023.xlsx', '2024': 'Pedidos_2024.xlsx', '2025': 'Pedidos_2025.xlsx'}
    all_dfs = []
    
    for year, name in files.items():
        path = os.path.join(INPUT_FOLDER, name)
        if os.path.exists(path):
            df = pd.read_excel(path)
            # Filtros solicitados
            df = df[(df['Canal de venta'] == 'Juntoz') & (df['Tipo de documento de cliente'] == 'DNI')]
            estados_validos = ['Received', 'ReadyToShip', 'ReadyToPickUp', 'PendingToPickUp', 'InTransit', 'Confirmed']
            df = df[df['Estado de item'].isin(estados_validos)]
            
            # Normalización de Fechas
            df['Fecha de creación'] = pd.to_datetime(df['Fecha de creación']).dt.normalize()
            # Limpieza de Montos
            df['Total'] = pd.to_numeric(df['Total'].astype(str).str.replace(',', '.'), errors='coerce')
            all_dfs.append(df)

    df_master = pd.concat(all_dfs, ignore_index=True)

    # 2. MÉTRICAS POR CLIENTE
    customers = df_master.groupby('Nro. de documento de cliente').agg(
        ventas_totales=('Total', 'sum'),
        num_ordenes=('Nro. de orden', 'nunique'),
        fecha_primera=('Fecha de creación', 'min'),
        fecha_ultima=('Fecha de creación', 'max')
    ).reset_index()

    customers['año_cohorte'] = customers['fecha_primera'].dt.year
    
    # Segmentación Pareto
    customers = customers.sort_values('ventas_totales', ascending=False)
    customers['cum_sum'] = customers['ventas_totales'].cumsum()
    customers['cum_perc'] = 100 * customers['cum_sum'] / customers['ventas_totales'].sum()
    customers['rank_perc'] = range(1, len(customers) + 1)
    customers['rank_perc'] = 100 * customers['rank_perc'] / len(customers)

    # 3. GENERACIÓN DE VISUALIZACIONES
    plt.style.use('ggplot')
    
    # Gráfico 1: Evolución Mensual
    plt.figure(figsize=(10, 4))
    df_master.set_index('Fecha de creación').resample('M')['Total'].sum().plot(color='#1a237e', lw=2)
    plt.title('Evolución Mensual de Ventas Netas (Juntoz)')
    plt.savefig(f'{OUTPUT_FOLDER}/ventas_mensuales.png')
    
    # Gráfico 2: Pareto LTV
    plt.figure(figsize=(10, 4))
    plt.plot(customers['rank_perc'].values, customers['cum_perc'].values, color='red')
    plt.axvline(10, color='grey', linestyle='--')
    plt.title('Curva de Concentración LTV (Pareto)')
    plt.xlabel('% Clientes')
    plt.ylabel('% Venta Acumulada')
    plt.savefig(f'{OUTPUT_FOLDER}/pareto_ltv.png')

    # 4. CONSTRUCCIÓN DEL PDF
    pdf = LTV_Report()
    
    # PÁGINA 1: PORTADA
    pdf.add_page()
    pdf.image(LOGO_URL, 80, 50, 50)
    pdf.ln(100)
    pdf.set_font('Arial', 'B', 24)
    pdf.cell(0, 20, 'Customer Lifetime Value Report', 0, 1, 'C')
    pdf.set_font('Arial', '', 16)
    pdf.cell(0, 10, 'Periodo: 2023 - 2025 | Canal: Juntoz', 0, 1, 'C')
    
    # PÁGINA 2: RESUMEN EJECUTIVO
    pdf.add_page()
    pdf.set_font('Arial', 'B', 16)
    pdf.cell(0, 10, '2. Resumen Ejecutivo', 0, 1)
    pdf.set_font('Arial', '', 12)
    
    kpis = [
        f"Clientes Únicos (DNI): {len(customers):,}",
        f"Ventas Netas Totales: S/ {df_master['Total'].sum():,.2f}",
        f"LTV Promedio: S/ {customers['ventas_totales'].mean():,.2f}",
        f"Ticket Promedio: S/ {df_master['Total'].mean():,.2f}",
        f"Frecuencia Promedio: {customers['num_ordenes'].mean():,.2f} órdenes/cliente"
    ]
    for kpi in kpis:
        pdf.cell(0, 10, f"- {kpi}", 0, 1)
    
    # PÁGINA 3: ANÁLISIS TEMPORAL
    pdf.add_page()
    pdf.set_font('Arial', 'B', 16)
    pdf.cell(0, 10, '3. Análisis Temporal y Evolución', 0, 1)
    pdf.image(f'{OUTPUT_FOLDER}/ventas_mensuales.png', x=10, w=190)
    pdf.ln(10)
    pdf.set_font('Arial', '', 11)
    pdf.multi_cell(0, 10, "El gráfico superior muestra la tendencia de ingresos netos. Se observa la estacionalidad y el comportamiento del canal Juntoz bajo los estados de neteo aplicados.")

    # PÁGINA 4: SEGMENTACIÓN PARETO
    pdf.add_page()
    pdf.set_font('Arial', 'B', 16)
    pdf.cell(0, 10, '4. Distribución de Valor (Pareto)', 0, 1)
    pdf.image(f'{OUTPUT_FOLDER}/pareto_ltv.png', x=10, w=190)
    top_10_val = customers[customers['rank_perc'] <= 10]['ventas_totales'].sum() / df_master['Total'].sum() * 100
    pdf.ln(5)
    pdf.multi_cell(0, 10, f"Insight Clave: El top 10% de los clientes genera el {top_10_val:.1f}% de la venta total de Juntoz. Este segmento es crítico para la estabilidad del canal.")

    # GUARDAR
    pdf.output(f"{OUTPUT_FOLDER}/LTV_Report_2023_2025.pdf")
    print("✅ PDF Generado con éxito.")

if __name__ == "__main__":
    generar_analisis()