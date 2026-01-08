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
LOGO_URL = "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQfE4betnoplLem-rHmrOt2gqS7zMBYV8D3aw&s"

# Definimos solo las columnas que vamos a usar para no saturar la RAM
COLS_TO_USE = [
    'Canal de venta', 'Sitio', 'Tipo de documento de cliente', 
    'Nro. de documento de cliente', 'Estado de item', 
    'Total', 'Fecha de creación', 'Nro. de orden'
]

COLOR_PRIMARY = (26, 35, 126) 
COLOR_TEXT = (50, 50, 50)

class LTV_Report(FPDF):
    def header(self):
        try: self.image(LOGO_URL, 10, 8, 30)
        except: pass
        self.set_font('Helvetica', 'B', 11)
        self.set_text_color(100, 100, 100)
        self.cell(0, 10, 'DIVISIÓN DE ANALÍTICA & ESTRATEGIA - JUNTOZ', 0, 0, 'R')
        self.ln(20)

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

    print("--- Cargando y Filtrando Data (Optimizado) ---")
    for year, name in FILES.items():
        path = os.path.join(INPUT_FOLDER, name)
        if os.path.exists(path):
            # Leemos solo lo necesario con el motor más rápido disponible
            df = pd.read_excel(path, engine='calamine', usecols=COLS_TO_USE)
            
            # Filtro inmediato: Canal, Sitio, DNI y Estados Neteados
            df = df[
                (df['Canal de venta'] == 'Juntoz') & 
                (df['Sitio'] == 'Juntoz') & 
                (df['Tipo de documento de cliente'] == 'DNI') & 
                (df['Estado de item'].isin(estados_validos))
            ].copy()
            
            # Limpieza de datos
            df['Total'] = pd.to_numeric(df['Total'].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
            df['Fecha de creación'] = pd.to_datetime(df['Fecha de creación'], errors='coerce')
            df['Año'] = year
            
            all_years_data.append(df)
            print(f"✅ Año {year} procesado.")

    if not all_years_data: return print("❌ Error: No hay datos para procesar.")
    
    df_master = pd.concat(all_years_data, ignore_index=True)

    # Cálculos Consolidados
    stats_anual = df_master.groupby('Año').agg(
        Venta=('Total', 'sum'),
        Clientes=('Nro. de documento de cliente', 'nunique'),
        Ordenes=('Nro. de orden', 'nunique')
    )
    stats_anual['Ticket_Prom'] = stats_anual['Venta'] / stats_anual['Ordenes']

    customers = df_master.groupby('Nro. de documento de cliente').agg(
        LTV_Total=('Total', 'sum'),
        Frecuencia=('Nro. de orden', 'nunique')
    ).sort_values('LTV_Total', ascending=False).reset_index()

    total_revenue = customers['LTV_Total'].sum()
    customers['Venta_Acum'] = customers['LTV_Total'].cumsum()
    pareto_perc = (customers[customers['Venta_Acum'] <= total_revenue * 0.8].shape[0] / len(customers)) * 100 if len(customers) > 0 else 0

    # Gráfico de Tendencia
    sns.set_theme(style="whitegrid")
    plt.figure(figsize=(12, 5))
    df_master.set_index('Fecha de creación').resample('M')['Total'].sum().plot(color='#1A237E', lw=3)
    plt.title('Tendencia Longitudinal de Ventas Mensuales', fontsize=14)
    plt.ylabel('Soles (S/)')
    plt.savefig(f'{OUTPUT_FOLDER}/g1_trend.png', bbox_inches='tight')
    plt.close()

    # --- PDF GENERATION ---
    pdf = LTV_Report()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # Portada
    pdf.add_page()
    pdf.ln(60)
    pdf.set_font('Helvetica', 'B', 32); pdf.set_text_color(*COLOR_PRIMARY)
    pdf.cell(0, 20, 'DASHBOARD ESTRATÉGICO LTV', 0, 1, 'C')
    pdf.set_font('Helvetica', '', 18); pdf.set_text_color(80, 80, 80)
    pdf.cell(0, 10, 'Consolidado Trienal | Canal y Sitio Juntoz', 0, 1, 'C')
    pdf.ln(20); pdf.set_draw_color(*COLOR_PRIMARY); pdf.line(40, 125, 170, 125)

    # Página 1: KPIs
    pdf.add_page()
    pdf.chapter_title('1. Resumen Ejecutivo (3 Años)')
    pdf.set_fill_color(240, 240, 250); pdf.set_font('Helvetica', 'B', 12)
    pdf.cell(90, 20, f"VENTA TOTAL: S/ {total_revenue:,.2f}", 1, 0, 'C', True)
    pdf.cell(10)
    pdf.cell(90, 20, f"CLIENTES DNI: {len(customers):,}", 1, 1, 'C', True)
    pdf.ln(5)
    pdf.cell(90, 20, f"LTV PROMEDIO: S/ {customers['LTV_Total'].mean():,.2f}", 1, 0, 'C', True)
    pdf.cell(10)
    pdf.cell(90, 20, f"TICKET PROM. GLOBAL: S/ {df_master['Total'].mean():,.2f}", 1, 1, 'C', True)
    pdf.ln(10)
    pdf.image(f'{OUTPUT_FOLDER}/g1_trend.png', x=10, w=190)

    # Página 2: Desglose Anual
    pdf.add_page()
    pdf.chapter_title('2. Performance por Año')
    header_anual = ['Año', 'Venta Neta', 'Clientes', 'Ticket Prom.']
    data_anual = [[idx, f"S/ {row['Venta']:,.2f}", f"{row['Clientes']:,}", f"S/ {row['Ticket_Prom']:,.2f}"] for idx, row in stats_anual.iterrows()]
    pdf.create_table(header_anual, data_anual, [30, 60, 50, 50])

    # Página 3: Top VIP
    pdf.add_page()
    pdf.chapter_title('3. Top 10 Clientes VIP (Valor LTV)')
    header_vip = ['DNI Cliente', 'LTV Acumulado', 'Órdenes', 'Frecuencia']
    data_vip = [[str(row['Nro. de documento de cliente']), f"S/ {row['LTV_Total']:,.2f}", str(row['Frecuencia']), "Alta" if row['Frecuencia'] >= 3 else "Regular"] for _, row in customers.head(10).iterrows()]
    pdf.create_table(header_vip, data_vip, [45, 55, 45, 45])

    pdf.chapter_title('4. Insights Clave')
    pdf.set_font('Helvetica', '', 11); pdf.set_text_color(*COLOR_TEXT)
    pdf.multi_cell(0, 10, f"* Concentración de Ingresos: El 80% del valor es generado por solo el {pareto_perc:.1f}% de los clientes.\n* Nota Metodológica: Análisis basado exclusivamente en transacciones DNI con estados neteados confirmados.")

    pdf.output(os.path.join(OUTPUT_FOLDER, FINAL_PDF_NAME))
    print("✅ Proceso completado.")

if __name__ == "__main__":
    generar_analisis_gerencial()