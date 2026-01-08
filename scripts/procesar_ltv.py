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

# ELIMINAMOS LOGO_URL EXTERNO PARA EVITAR ERRORES DE LIBRERÍA
COLS_TO_USE = [
    'Canal de venta', 'Sitio', 'Tipo de documento de cliente', 
    'Nro. de documento de cliente', 'Estado de item', 
    'Total', 'Fecha de creación', 'Nro. de orden', 'Cantidad'
]

COLOR_PRIMARY = (26, 35, 126) 
COLOR_TEXT = (50, 50, 50)

class LTV_Report(FPDF):
    def header(self):
        # HEADER SIN IMAGEN PARA EVITAR ERROR DE "UNSUPPORTED TYPE"
        self.set_font('Helvetica', 'B', 11)
        self.set_text_color(100, 100, 100)
        self.cell(0, 10, 'DIVISIÓN DE ANALÍTICA & ESTRATEGIA - JUNTOZ', 0, 0, 'L')
        self.set_font('Helvetica', 'B', 8)
        self.cell(0, 10, 'REPORTE MULTICANAL 2023-2025', 0, 0, 'R')
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

    print("--- Cargando y Segmentando Data (Multicanal) ---")
    for year, name in FILES.items():
        path = os.path.join(INPUT_FOLDER, name)
        if os.path.exists(path):
            df = pd.read_excel(path, engine='calamine', usecols=COLS_TO_USE)
            
            # FILTRO: Solo Sitio Juntoz (Abrimos Canal y Tipo Doc)
            df = df[(df['Sitio'] == 'Juntoz') & (df['Estado de item'].isin(estados_validos))].copy()
            
            df['Total'] = pd.to_numeric(df['Total'].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
            df['Cantidad'] = pd.to_numeric(df['Cantidad'], errors='coerce').fillna(0)
            df['Año'] = year
            all_years_data.append(df)

    if not all_years_data: return print("❌ Error: No hay datos.")
    
    df_master = pd.concat(all_years_data, ignore_index=True)

    # LÓGICA MAYORISTA VS MINORISTA POR ORDEN
    order_type = df_master.groupby('Nro. de orden')['Cantidad'].sum().reset_index()
    order_type['Tipo_Venta'] = order_type['Cantidad'].apply(lambda x: 'Mayorista' if x > 2 else 'Minorista')
    df_master = df_master.merge(order_type[['Nro. de orden', 'Tipo_Venta']], on='Nro. de orden', how='left')

    # ESTADÍSTICAS POR TIPO
    stats_tipo = df_master.groupby('Tipo_Venta').agg(Venta=('Total', 'sum'), Clientes=('Nro. de documento de cliente', 'nunique'), Ordenes=('Nro. de orden', 'nunique'))
    stats_tipo['Ticket_Prom'] = stats_tipo['Venta'] / stats_tipo['Ordenes']

    # GRÁFICOS
    sns.set_theme(style="whitegrid")
    plt.figure(figsize=(8, 6))
    stats_tipo['Venta'].plot(kind='pie', autopct='%1.1f%%', colors=['#1A237E', '#FF5252'])
    plt.title('Distribución de Ingresos: Mayorista vs Minorista')
    plt.savefig(f'{OUTPUT_FOLDER}/g_tipo_venta.png', bbox_inches='tight')
    plt.close()

    plt.figure(figsize=(10, 5))
    df_master.groupby('Canal de venta')['Total'].sum().sort_values().plot(kind='barh', color='#1A237E')
    plt.title('Ingresos Totales por Canal de Venta')
    plt.savefig(f'{OUTPUT_FOLDER}/g_canales.png', bbox_inches='tight')
    plt.close()

    # --- GENERAR PDF ---
    pdf = LTV_Report()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # Portada
    pdf.add_page()
    pdf.ln(70)
    pdf.set_font('Helvetica', 'B', 30); pdf.set_text_color(*COLOR_PRIMARY)
    pdf.cell(0, 20, 'REPORTE ESTRATÉGICO MULTICANAL', 0, 1, 'C')
    pdf.set_font('Helvetica', '', 18); pdf.set_text_color(80, 80, 80)
    pdf.cell(0, 10, 'Segmentación por Volumen y Origen de Cliente', 0, 1, 'C')

    # Página 1: Mayorista vs Minorista
    pdf.add_page()
    pdf.chapter_title('1. Análisis de Segmentos: Mayorista vs Minorista')
    pdf.image(f'{OUTPUT_FOLDER}/g_tipo_venta.png', x=50, w=110)
    header_tipo = ['Segmento', 'Venta Total', 'Clientes', 'Ticket Prom.']
    data_tipo = [[idx, f"S/ {row['Venta']:,.2f}", f"{row['Clientes']:,}", f"S/ {row['Ticket_Prom']:,.2f}"] for idx, row in stats_tipo.iterrows()]
    pdf.create_table(header_tipo, data_tipo, [45, 50, 45, 50])

    # Página 2: Canales
    pdf.add_page()
    pdf.chapter_title('2. Distribución de Ingresos por Canal')
    pdf.image(f'{OUTPUT_FOLDER}/g_canales.png', x=10, w=190)
    pdf.ln(5)
    pdf.set_font('Helvetica', 'B', 12)
    pdf.cell(0, 10, 'Top Documentos de Cliente por Ingreso:', 0, 1)
    doc_stats = df_master.groupby('Tipo de documento de cliente')['Total'].sum().sort_values(ascending=False).head(5)
    for doc, val in doc_stats.items():
        pdf.set_font('Helvetica', '', 10)
        pdf.cell(0, 8, f"- {doc}: S/ {val:,.2f}", 0, 1)

    # Página 3: Conclusiones
    pdf.add_page()
    pdf.chapter_title('3. Conclusiones Ejecutivas')
    pdf.set_font('Helvetica', '', 11); pdf.set_text_color(*COLOR_TEXT)
    pdf.multi_cell(0, 10, f"* Estrategia de Volumen: El segmento Mayorista (>2 unidades) ya es visible en la distribución de ingresos.\n* Origen del Cliente: La apertura de canales permite identificar fuentes de tráfico prioritarias.\n* Diversificación: Se identifican documentos de identidad no convencionales con alto potencial de LTV.")

    pdf.output(os.path.join(OUTPUT_FOLDER, FINAL_PDF_NAME))
    print("✅ Dashboard Multicanal Generado.")

if __name__ == "__main__":
    generar_analisis_gerencial()