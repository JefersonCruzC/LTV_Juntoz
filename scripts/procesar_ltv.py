import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
from fpdf import FPDF
from datetime import datetime

# Configuración
INPUT_FOLDER = 'data_pedidos'
OUTPUT_FOLDER = 'reportes'
# Nombre unificado para evitar errores de ruta en GitHub
FINAL_PDF_NAME = 'LTV_Executive_Report_Juntoz.pdf'
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
            print(f"Leyendo: {name}")
            df = pd.read_excel(path)
            # Filtros solicitados
            df = df[(df['Canal de venta'] == 'Juntoz') & (df['Sitio'] == 'Juntoz') & 
                    (df['Tipo de documento de cliente'] == 'DNI') & (df['Estado de item'].isin(estados_validos))].copy()
            df['Total'] = pd.to_numeric(df['Total'].astype(str).str.replace(',', '.'), errors='coerce')
            df['Fecha de creación'] = pd.to_datetime(df['Fecha de creación']).dt.normalize()
            df['Año'] = year
            all_data.append(df)

    if not all_data:
        print("❌ ERROR: No se encontraron datos para procesar.")
        return

    df_master = pd.concat(all_data, ignore_index=True)

    # Métricas Anuales
    stats_anual = df_master.groupby('Año').agg(
        Venta_Neta=('Total', 'sum'),
        Clientes_Unicos=('Nro. de documento de cliente', 'nunique'),
        Ordenes=('Nro. de orden', 'nunique')
    )
    stats_anual['Ticket_Promedio'] = stats_anual['Venta_Neta'] / stats_anual['Ordenes']

    # Clientes VIP
    customers = df_master.groupby('Nro. de documento de cliente').agg(
        LTV_Total=('Total', 'sum'),
        Frecuencia=('Nro. de orden', 'nunique'),
        Ultima_Compra=('Fecha de creación', 'max')
    ).sort_values('LTV_Total', ascending=False).reset_index()

    top_10_vips = customers.head(10)

    # Gráfico
    plt.style.use('ggplot')
    fig, ax1 = plt.subplots(figsize=(10, 5))
    stats_anual['Venta_Neta'].plot(kind='bar', color='#1a237e', ax=ax1)
    plt.title('Evolución Anual: Ventas Juntoz (2023-2025)')
    plt.savefig(f'{OUTPUT_FOLDER}/evolucion_anual.png')
    plt.close()

    # GENERAR PDF
    pdf = LTV_Report()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # Portada
    pdf.add_page()
    pdf.ln(70)
    pdf.set_font('Arial', 'B', 26)
    pdf.cell(0, 15, 'ANÁLISIS ESTRATÉGICO LTV', 0, 1, 'C')
    pdf.set_font('Arial', '', 14)
    pdf.cell(0, 10, 'Consolidado Trienal 2023 - 2025', 0, 1, 'C')
    pdf.cell(0, 10, 'Canal y Sitio: Juntoz', 0, 1, 'C')

    # Desglose Anual
    pdf.add_page()
    pdf.set_font('Arial', 'B', 16)
    pdf.cell(0, 10, '1. Performance Anual Detallada', 0, 1)
    pdf.image(f'{OUTPUT_FOLDER}/evolucion_anual.png', x=10, w=190)
    pdf.ln(5)
    
    # Tabla Anual
    pdf.set_font('Arial', 'B', 10)
    for h in ['Año', 'Venta Neta', 'Clientes', 'Ticket Prom.']: pdf.cell(47, 10, h, 1)
    pdf.ln()
    pdf.set_font('Arial', '', 10)
    for index, row in stats_anual.iterrows():
        pdf.cell(47, 10, str(index), 1)
        pdf.cell(47, 10, f"S/ {row['Venta_Neta']:,.2f}", 1)
        pdf.cell(47, 10, f"{row['Clientes_Unicos']:,}", 1)
        pdf.cell(47, 10, f"S/ {row['Ticket_Promedio']:,.2f}", 1)
        pdf.ln()

    # TOP 10 VIP
    pdf.add_page()
    pdf.set_font('Arial', 'B', 16)
    pdf.cell(0, 10, '2. Ranking TOP 10 Clientes VIP', 0, 1)
    pdf.ln(5)
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(60, 10, 'DNI Cliente', 1); pdf.cell(60, 10, 'LTV Total (S/)', 1); pdf.cell(60, 10, 'Frecuencia', 1)
    pdf.ln()
    pdf.set_font('Arial', '', 10)
    for _, row in top_10_vips.iterrows():
        pdf.cell(60, 10, str(row['Nro. de documento de cliente']), 1)
        pdf.cell(60, 10, f"S/ {row['LTV_Total']:,.2f}", 1)
        pdf.cell(60, 10, str(row['Frecuencia']), 1)
        pdf.ln()

    # Análisis
    pdf.add_page()
    pdf.set_font('Arial', 'B', 16)
    pdf.cell(0, 10, '3. Conclusiones y Análisis Senior', 0, 1)
    pdf.set_font('Arial', '', 11)
    pdf.ln(5)
    comentarios = [
        "Fidelización: El Top 10 concentra un valor estratégico alto; se recomienda un plan de lealtad.",
        "Calidad de Datos: Análisis basado en estados neteados (dinero real ingresado).",
        "Evolución: El Ticket Promedio anual es el principal indicador de salud de ventas en el sitio Juntoz."
    ]
    for c in comentarios:
        pdf.multi_cell(0, 10, f"* {c}")
        pdf.ln(2)

    # GUARDAR PDF
    pdf_path = os.path.join(OUTPUT_FOLDER, FINAL_PDF_NAME)
    pdf.output(pdf_path)
    print(f"✅ Reporte generado exitosamente en: {pdf_path}")

if __name__ == "__main__":
    generar_analisis_avanzado()