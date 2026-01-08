import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import os
from fpdf import FPDF
from datetime import datetime

# --- CONFIGURACIÓN ---
INPUT_FOLDER = 'data_pedidos'
OUTPUT_FOLDER = 'reportes'
FINAL_PDF_NAME = 'LTV_Executive_Report_Juntoz.pdf'
FILES = {'2023': 'Pedidos_2023.xlsx', '2024': 'Pedidos_2024.xlsx', '2025': 'Pedidos_2025.xlsx'}

COLS_TO_USE = [
    'Canal de venta', 'Sitio', 'Tipo de documento de cliente', 
    'Nro. de documento de cliente', 'Estado de item', 
    'Total', 'Fecha de creación', 'Nro. de orden', 'Cantidad'
]

COLOR_PRIMARY = (26, 35, 126) 
COLOR_TEXT = (50, 50, 50)

class LTV_Report(FPDF):
    def header(self):
        self.set_font('Helvetica', 'B', 11)
        self.set_text_color(100, 100, 100)
        self.cell(0, 10, 'DIVISIÓN DE ANALÍTICA & ESTRATEGIA - JUNTOZ', 0, 0, 'L')
        self.set_font('Helvetica', 'B', 8)
        self.cell(0, 10, 'REPORTE ESTRATÉGICO INTEGRAL 2023-2025', 0, 0, 'R')
        self.ln(15)

    def footer(self):
        self.set_y(-15)
        self.set_font('Helvetica', 'I', 8)
        self.set_text_color(128, 128, 128)
        self.cell(0, 10, f'Confidencial | Generado: {datetime.now().strftime("%d/%m/%Y")} | Página {self.page_no()}', 0, 0, 'C')

    def chapter_title(self, title):
        self.set_font('Helvetica', 'B', 16)
        self.set_text_color(*COLOR_PRIMARY)
        self.cell(0, 10, title, 0, 1, 'L')
        self.ln(4)

    def create_table(self, header, data, col_widths):
        self.set_fill_color(*COLOR_PRIMARY)
        self.set_text_color(255, 255, 255)
        self.set_font('Helvetica', 'B', 10)
        for i, h in enumerate(header):
            self.cell(col_widths[i], 10, h, 1, 0, 'C', True)
        self.ln()
        self.set_text_color(*COLOR_TEXT)
        self.set_font('Helvetica', '', 9)
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
    
    all_years_data = []
    estados_validos = ['Received', 'ReadyToShip', 'ReadyToPickUp', 'PendingToPickUp', 'InTransit', 'Confirmed']

    print("--- Procesando Pipeline Senior ---")
    for year, name in FILES.items():
        path = os.path.join(INPUT_FOLDER, name)
        if os.path.exists(path):
            df = pd.read_excel(path, engine='calamine', usecols=COLS_TO_USE)
            df = df[(df['Sitio'] == 'Juntoz') & (df['Estado de item'].isin(estados_validos))].copy()
            df['Total'] = pd.to_numeric(df['Total'].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
            df['Cantidad'] = pd.to_numeric(df['Cantidad'], errors='coerce').fillna(0)
            df['Fecha de creación'] = pd.to_datetime(df['Fecha de creación'], errors='coerce')
            df['Año'] = year
            all_years_data.append(df)

    if not all_years_data: return print("❌ Sin datos.")
    
    df_master = pd.concat(all_years_data, ignore_index=True)
    fecha_max = df_master['Fecha de creación'].max()

    # 1. LÓGICA MAYORISTA VS MINORISTA
    order_type = df_master.groupby('Nro. de orden')['Cantidad'].sum().reset_index()
    order_type['Tipo_Venta'] = order_type['Cantidad'].apply(lambda x: 'Mayorista' if x > 2 else 'Minorista')
    df_master = df_master.merge(order_type[['Nro. de orden', 'Tipo_Venta']], on='Nro. de orden', how='left')

    # 2. ANÁLISIS DE RETENCIÓN (2023 vs 2025)
    c_2023 = set(df_master[df_master['Año'] == '2023']['Nro. de documento de cliente'])
    c_2025 = set(df_master[df_master['Año'] == '2025']['Nro. de documento de cliente'])
    tasa_retencion = (len(c_2023 & c_2025) / len(c_2023) * 100) if c_2023 else 0

    # 3. MÉTRICAS POR CLIENTE (LTV + RFM)
    customers = df_master.groupby('Nro. de documento de cliente').agg(
        LTV_Total=('Total', 'sum'),
        Frecuencia=('Nro. de orden', 'nunique'),
        Ultima_Compra=('Fecha de creación', 'max'),
        Tipo_Doc=('Tipo de documento de cliente', 'first')
    ).reset_index()

    def get_status(fecha):
        days = (fecha_max - fecha).days
        return "Activo" if days < 180 else ("En Riesgo" if days < 365 else "Dormido")
    
    customers['Status'] = customers['Ultima_Compra'].apply(get_status)
    customers = customers.sort_values('LTV_Total', ascending=False)
    
    total_revenue = df_master['Total'].sum()
    customers['Venta_Acum'] = customers['LTV_Total'].cumsum()
    pareto_perc = (customers[customers['Venta_Acum'] <= total_revenue * 0.8].shape[0] / len(customers)) * 100

    # --- GRÁFICOS ---
    sns.set_theme(style="whitegrid")
    # G1: Mayorista vs Minorista
    plt.figure(figsize=(8, 6))
    df_master.groupby('Tipo_Venta')['Total'].sum().plot(kind='pie', autopct='%1.1f%%', colors=['#1A237E', '#FF5252'])
    plt.title('Distribución de Venta: Mayorista vs Minorista')
    plt.savefig(f'{OUTPUT_FOLDER}/g_tipo.png', bbox_inches='tight')
    plt.close()
    
    # G2: Canales
    plt.figure(figsize=(10, 5))
    df_master.groupby('Canal de venta')['Total'].sum().sort_values().plot(kind='barh', color='#1A237E')
    plt.title('Ingresos por Canal de Venta (Total 3 Años)')
    plt.savefig(f'{OUTPUT_FOLDER}/g_canales.png', bbox_inches='tight')
    plt.close()

    # --- PDF GENERATION ---
    pdf = LTV_Report()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # P1: PORTADA
    pdf.add_page()
    pdf.ln(70)
    pdf.set_font('Helvetica', 'B', 30); pdf.set_text_color(*COLOR_PRIMARY)
    pdf.cell(0, 20, 'DASHBOARD ESTRATÉGICO LTV 360', 0, 1, 'C')
    pdf.set_font('Helvetica', '', 16); pdf.set_text_color(80, 80, 80)
    pdf.cell(0, 10, 'Multicanal | Segmentación Mayorista | Análisis de Retención', 0, 1, 'C')

    # P2: RESUMEN EJECUTIVO
    pdf.add_page()
    pdf.chapter_title('1. Resumen Ejecutivo y Salud de Cartera')
    pdf.set_fill_color(240, 240, 250); pdf.set_font('Helvetica', 'B', 12)
    pdf.cell(90, 20, f"VENTA TOTAL: S/ {total_revenue:,.2f}", 1, 0, 'C', True)
    pdf.cell(10)
    pdf.cell(90, 20, f"TASA RETENCIÓN (23-25): {tasa_retencion:.1f}%", 1, 1, 'C', True)
    pdf.ln(5)
    pdf.cell(90, 20, f"LTV PROMEDIO: S/ {customers['LTV_Total'].mean():,.2f}", 1, 0, 'C', True)
    pdf.cell(10)
    pdf.cell(90, 20, f"TICKET PROM. GLOBAL: S/ {df_master['Total'].mean():,.2f}", 1, 1, 'C', True)
    pdf.ln(10)
    pdf.image(f'{OUTPUT_FOLDER}/g_tipo.png', x=50, w=110)

    # P3: PERFORMANCE ANUAL
    pdf.add_page()
    pdf.chapter_title('2. Desglose de Performance por Año (2023-2025)')
    stats_anual = df_master.groupby('Año').agg(Venta=('Total', 'sum'), Clientes=('Nro. de documento de cliente', 'nunique'), Ordenes=('Nro. de orden', 'nunique'))
    stats_anual['Ticket_Prom'] = stats_anual['Venta'] / stats_anual['Ordenes']
    header_anual = ['Año', 'Venta Neta', 'Clientes Únicos', 'Ticket Prom.']
    data_anual = [[idx, f"S/ {row['Venta']:,.2f}", f"{row['Clientes']:,}", f"S/ {row['Ticket_Prom']:,.2f}"] for idx, row in stats_anual.iterrows()]
    pdf.create_table(header_anual, data_anual, [30, 60, 50, 50])
    pdf.image(f'{OUTPUT_FOLDER}/g_canales.png', x=10, w=190)

    # P4: TOP 10 VIP
    pdf.add_page()
    pdf.chapter_title('3. Top 10 Clientes de Mayor Impacto (VIP)')
    header_vip = ['ID / Documento', 'LTV Total (S/)', 'Órdenes', 'Status Actual']
    data_vip = [[str(row['Nro. de documento de cliente']), f"S/ {row['LTV_Total']:,.2f}", str(row['Frecuencia']), row['Status']] for _, row in customers.head(10).iterrows()]
    pdf.create_table(header_vip, data_vip, [50, 50, 40, 50])

    # P5: INSIGHTS
    pdf.add_page()
    pdf.chapter_title('4. Insights y Comentarios Senior')
    pdf.set_font('Helvetica', '', 11); pdf.set_text_color(*COLOR_TEXT)
    
    # Cálculo dinámico para insight
    perc_may = (df_master[df_master['Tipo_Venta'] == 'Mayorista']['Total'].sum() / total_revenue) * 100
    
    insights = [
        f"Segmentación de Volumen: El canal Mayorista representa el {perc_may:.1f}% de la facturación total.",
        f"Salud de Retención: La tasa de supervivencia interanual es del {tasa_retencion:.1f}%. Hay oportunidad de reactivación.",
        f"Concentración (Pareto): El 80% de los ingresos es generado por el {pareto_perc:.1f}% de la base de clientes.",
        "Multicanalidad: La apertura de canales y documentos permite identificar nichos de mercado fuera del DNI tradicional."
    ]
    for ins in insights:
        pdf.multi_cell(0, 10, f"* {ins}")
        pdf.ln(2)

    pdf.output(os.path.join(OUTPUT_FOLDER, FINAL_PDF_NAME))
    print("✅ Reporte Maestro Finalizado.")

if __name__ == "__main__":
    generar_analisis_gerencial()