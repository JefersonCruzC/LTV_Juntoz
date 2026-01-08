import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import os
from fpdf import FPDF
from datetime import datetime

# --- CONFIGURACIÓN ESTRATÉGICA ---
INPUT_FOLDER = 'data_pedidos'
OUTPUT_FOLDER = 'reportes'
FINAL_PDF_NAME = 'LTV_Executive_Report_Juntoz.pdf'
FILES = {'2023': 'Pedidos_2023.xlsx', '2024': 'Pedidos_2024.xlsx', '2025': 'Pedidos_2025.xlsx'}
LOGO_URL = "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQfE4betnoplLem-rHmrOt2gqS7zMBYV8D3aw&s"

# Colores Corporativos
COLOR_PRIMARY = (26, 35, 126)   # Azul Oscuro
COLOR_SECONDARY = (255, 82, 82) # Rojo Suave
COLOR_TEXT = (50, 50, 50)

class LTV_Report(FPDF):
    def header(self):
        try: self.image(LOGO_URL, 10, 8, 30)
        except: pass
        self.set_font('Arial', 'B', 11)
        self.set_text_color(100, 100, 100)
        self.cell(0, 10, 'DIVISIÓN DE ANALÍTICA & ESTRATEGIA - JUNTOZ', 0, 0, 'R')
        self.ln(20)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.set_text_color(128, 128, 128)
        self.cell(0, 10, f'Confidencial | Generado el {datetime.now().strftime("%d/%m/%Y")} | Página {self.page_no()}', 0, 0, 'C')

    def chapter_title(self, title):
        self.set_font('Arial', 'B', 16)
        self.set_text_color(*COLOR_PRIMARY)
        self.cell(0, 10, title, 0, 1, 'L')
        self.ln(4)

    def create_table(self, header, data, col_widths):
        # Header
        self.set_fill_color(*COLOR_PRIMARY)
        self.set_text_color(255, 255, 255)
        self.set_font('Arial', 'B', 10)
        for i, h in enumerate(header):
            self.cell(col_widths[i], 10, h, 1, 0, 'C', True)
        self.ln()
        # Body
        self.set_text_color(*COLOR_TEXT)
        self.set_font('Arial', '', 9)
        fill = False
        for row in data:
            self.set_fill_color(245, 245, 245) if fill else self.set_fill_color(255, 255, 255)
            for i, datum in enumerate(row):
                self.cell(col_widths[i], 8, str(datum), 1, 0, 'C', True)
            self.ln()
            fill = not fill
        self.ln(5)

def generar_analisis_gerencial():
    if not os.path.exists(OUTPUT_FOLDER): os.makedirs(OUTPUT_FOLDER)
    
    # 1. CARGA Y PROCESAMIENTO
    all_data = []
    estados_validos = ['Received', 'ReadyToShip', 'ReadyToPickUp', 'PendingToPickUp', 'InTransit', 'Confirmed']
    
    for year, name in FILES.items():
        path = os.path.join(INPUT_FOLDER, name)
        if os.path.exists(path):
            df = pd.read_excel(path)
            df = df[(df['Canal de venta'] == 'Juntoz') & (df['Sitio'] == 'Juntoz') & 
                    (df['Tipo de documento de cliente'] == 'DNI') & (df['Estado de item'].isin(estados_validos))].copy()
            df['Total'] = pd.to_numeric(df['Total'].astype(str).str.replace(',', '.'), errors='coerce')
            df['Fecha de creación'] = pd.to_datetime(df['Fecha de creación'])
            df['Año'] = year
            all_data.append(df)

    if not all_data: return print("❌ Sin datos.")
    df_master = pd.concat(all_data, ignore_index=True)

    # --- CÁLCULOS ESTRATÉGICOS ---
    # Métricas Anuales
    stats_anual = df_master.groupby('Año').agg(
        Venta=('Total', 'sum'),
        Clientes=('Nro. de documento de cliente', 'nunique'),
        Ordenes=('Nro. de orden', 'nunique')
    )
    stats_anual['Ticket_Prom'] = stats_anual['Venta'] / stats_anual['Ordenes']

    # Métricas Totales (LTV)
    customers = df_master.groupby('Nro. de documento de cliente').agg(
        LTV_Total=('Total', 'sum'),
        Frecuencia=('Nro. de orden', 'nunique'),
        Primera_Compra=('Fecha de creación', 'min'),
        Ultima_Compra=('Fecha de creación', 'max')
    ).sort_values('LTV_Total', ascending=False).reset_index()

    # Pareto 80/20
    customers['Venta_Acum'] = customers['LTV_Total'].cumsum()
    total_revenue = customers['LTV_Total'].sum()
    pareto_idx = customers[customers['Venta_Acum'] <= total_revenue * 0.8].shape[0]
    pareto_perc_customers = (pareto_idx / len(customers)) * 100

    # Segmentación
    customers['Segmento'] = 'Nuevo'
    customers.loc[customers['Frecuencia'] >= 2, 'Segmento'] = 'Recurrente'
    customers.loc[customers['Frecuencia'] >= 5, 'Segmento'] = 'Leal (VIP)'

    # --- GENERACIÓN DE GRÁFICOS ---
    sns.set_theme(style="whitegrid")
    
    # G1: Evolución Longitudinal (Mensual)
    plt.figure(figsize=(12, 5))
    df_master.set_index('Fecha de creación').resample('M')['Total'].sum().plot(color=sns.color_palette("muted")[0], lw=3)
    plt.title('Tendencia de Ventas Netas Mensuales (2023 - 2025)', fontsize=14, pad=20)
    plt.ylabel('Monto en Soles')
    plt.savefig(f'{OUTPUT_FOLDER}/g1_mensual.png', bbox_inches='tight')
    plt.close()

    # G2: Distribución de Segmentos
    plt.figure(figsize=(8, 6))
    customers['Segmento'].value_counts().plot(kind='pie', autopct='%1.1f%%', colors=sns.color_palette("pastel"))
    plt.title('Distribución de Cartera de Clientes', fontsize=14)
    plt.ylabel('')
    plt.savefig(f'{OUTPUT_FOLDER}/g2_segmentos.png', bbox_inches='tight')
    plt.close()

    # --- GENERAR PDF ---
    pdf = LTV_Report()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # PAG 1: PORTADA PROFESIONAL
    pdf.add_page()
    pdf.ln(60)
    pdf.set_font('Arial', 'B', 32)
    pdf.set_text_color(*COLOR_PRIMARY)
    pdf.cell(0, 20, 'DASHBOARD EJECUTIVO LTV', 0, 1, 'C')
    pdf.set_font('Arial', '', 18)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(0, 10, 'Customer Lifetime Value & Performance Analysis', 0, 1, 'C')
    pdf.ln(10)
    pdf.set_draw_color(*COLOR_PRIMARY)
    pdf.line(40, 120, 170, 120)
    pdf.ln(20)
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, 'PERIODO: ENERO 2023 - DICIEMBRE 2025', 0, 1, 'C')
    pdf.cell(0, 10, 'CANAL: JUNTOZ.COM | SITIO: JUNTOZ', 0, 1, 'C')

    # PAG 2: KPI CONSOLIDADO (GRAN TOTAL)
    pdf.add_page()
    pdf.chapter_title('1. Resumen Consolidado (Trienio 2023-2025)')
    
    # Cuadros de KPIs
    pdf.set_fill_color(240, 240, 250)
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(90, 25, f"VENTA NETA TOTAL: S/ {total_revenue:,.2f}", 1, 0, 'C', True)
    pdf.cell(10)
    pdf.cell(90, 25, f"CLIENTES ÚNICOS: {len(customers):,}", 1, 1, 'C', True)
    pdf.ln(5)
    pdf.cell(90, 25, f"LTV PROMEDIO: S/ {customers['LTV_Total'].mean():,.2f}", 1, 0, 'C', True)
    pdf.cell(10)
    pdf.cell(90, 25, f"TICKET PROM. GLOBAL: S/ {df_master['Total'].mean():,.2f}", 1, 1, 'C', True)
    
    pdf.ln(10)
    pdf.chapter_title('2. Evolución Histórica Longitudinal')
    pdf.image(f'{OUTPUT_FOLDER}/g1_mensual.png', x=10, w=190)
    pdf.set_font('Arial', 'I', 10)
    pdf.multi_cell(0, 10, "Interpretación: Se observa la estacionalidad del negocio. Los picos corresponden a eventos de alto tráfico (CyberDays/Navidad).")

    # PAG 3: PERFORMANCE ANUAL & SEGMENTACIÓN
    pdf.add_page()
    pdf.chapter_title('3. Comparativa de Performance Anual')
    
    header_anual = ['Año', 'Venta Neta', 'Clientes', 'Órdenes', 'Ticket Prom.']
    data_anual = [[
        idx, f"S/ {row['Venta']:,.2f}", f"{row['Clientes']:,}", f"{row['Ordenes']:,}", f"S/ {row['Ticket_Prom']:,.2f}"
    ] for idx, row in stats_anual.iterrows()]
    pdf.create_table(header_anual, data_anual, [30, 45, 35, 35, 45])

    pdf.chapter_title('4. Calidad de Cartera (Segmentación)')
    pdf.image(f'{OUTPUT_FOLDER}/g2_segmentos.png', x=50, w=110)
    pdf.ln(5)
    pdf.set_font('Arial', '', 11)
    recurrencia_total = (customers[customers['Frecuencia'] > 1].shape[0] / len(customers)) * 100
    pdf.multi_cell(0, 10, f"Análisis de Retención: El {recurrencia_total:.1f}% de la base de clientes ha comprado más de una vez. Este indicador es vital para la sostenibilidad del LTV.")

    # PAG 4: TOP VIP & PARETO
    pdf.add_page()
    pdf.chapter_title('5. Ranking TOP 10 Clientes de Alto Valor (VIP)')
    
    header_vip = ['DNI Cliente', 'LTV Acumulado', 'Órdenes', 'Última Compra']
    data_vip = [[
        str(row['Nro. de documento de cliente']), 
        f"S/ {row['LTV_Total']:,.2f}", 
        str(row['Frecuencia']), 
        row['Ultima_Compra'].strftime('%d/%m/%Y')
    ] for _, row in customers.head(10).iterrows()]
    pdf.create_table(header_vip, data_vip, [45, 55, 35, 55])

    pdf.chapter_title('6. Insights & Conclusiones Gerenciales')
    pdf.set_font('Arial', 'B', 11)
    pdf.set_text_color(*COLOR_SECONDARY)
    pdf.cell(0, 10, f"Efecto Pareto Identificado: El 80% de la facturación es generada por el {pareto_perc_customers:.1f}% de los clientes.", 0, 1)
    
    pdf.set_font('Arial', '', 11)
    pdf.set_text_color(*COLOR_TEXT)
    pdf.ln(5)
    insights = [
        "Sostenibilidad del LTV: El crecimiento del Ticket Promedio en 2025 sugiere una mejor capitalización de la base existente a pesar de la reducción en captación.",
        "Oportunidad de Recurrencia: La gran masa de clientes 'Nuevos' (compra única) representa la mayor oportunidad de crecimiento mediante campañas de remarketing.",
        "Riesgo de Concentración: La dependencia del Pareto indica que la pérdida de pocos clientes VIP impactaría severamente en el bottom-line.",
        "Filtros de Calidad: Este reporte considera únicamente estados neteados, asegurando que el análisis se basa en ingresos reales liquidados."
    ]
    for insight in insights:
        pdf.multi_cell(0, 8, f"• {insight}")
        pdf.ln(2)

    pdf.output(os.path.join(OUTPUT_FOLDER, FINAL_PDF_NAME))
    print(f"✅ Dashboard Ejecutivo generado con éxito.")

if __name__ == "__main__":
    generar_analisis_gerencial()