import pandas as pd
import os
import matplotlib.pyplot as plt

# Configuración de carpetas
INPUT_FOLDER = 'data_pedidos'
OUTPUT_FOLDER = 'reportes'
FILES = {'2023': 'Pedidos_2023.xlsx', '2024': 'Pedidos_2024.xlsx', '2025': 'Pedidos_2025.xlsx'}

def generar_reporte_gerencial():
    if not os.path.exists(OUTPUT_FOLDER):
        os.makedirs(OUTPUT_FOLDER)

    all_data = []
    
    print("--- Iniciando Procesamiento de Datos ---")
    for year, name in FILES.items():
        path = os.path.join(INPUT_FOLDER, name)
        if os.path.exists(path):
            print(f"Leyendo año {year}...")
            df = pd.read_excel(path)
            
            # 1. FILTROS ESTRICTOS
            # Solo Canal Juntoz y Tipo Documento DNI
            df = df[(df['Canal de venta'] == 'Juntoz') & (df['Tipo de documento de cliente'] == 'DNI')]
            
            # Solo Estados Neteados
            estados_netos = ['Received', 'ReadyToShip', 'ReadyToPickUp', 'PendingToPickUp', 'InTransit', 'Confirmed']
            df = df[df['Estado de item'].isin(estados_netos)]
            
            # Limpieza de montos (Total)
            df['Total'] = pd.to_numeric(df['Total'].astype(str).str.replace(',', '.'), errors='coerce')
            df['Año'] = year
            all_data.append(df)

    if not all_data:
        print("❌ No se encontraron datos válidos.")
        return

    df_master = pd.concat(all_data, ignore_index=True)

    # --- CÁLCULOS ESTRATÉGICOS ---
    
    # Agrupación por Cliente (DNI)
    ltv_reporte = df_master.groupby('Nro. de documento de cliente').agg({
        'Total': 'sum',
        'Nro. de orden': 'nunique',
        'Fecha de creación': ['min', 'max']
    }).reset_index()
    
    ltv_reporte.columns = ['DNI', 'LTV_Acumulado', 'Total_Pedidos', 'Primera_Compra', 'Ultima_Compra']

    # Segmentación TOP 1% (VIP)
    umbral_vip = ltv_reporte['LTV_Acumulado'].quantile(0.99)
    ltv_reporte['Segmento'] = ltv_reporte['LTV_Acumulado'].apply(lambda x: 'TOP 1% VIP' if x >= umbral_vip else 'Regular')

    # Análisis de Deserción (Churn) - Clientes que no compran hace más de 6 meses
    ultima_fecha_global = pd.to_datetime(df_master['Fecha de creación']).max()
    ltv_reporte['Status'] = ltv_reporte['Ultima_Compra'].apply(
        lambda x: 'Activo' if (ultima_fecha_global - pd.to_datetime(x)).days < 180 else 'Fugado (Churn)'
    )

    # --- GENERACIÓN DE GRÁFICOS PARA GERENCIA ---
    plt.figure(figsize=(12, 10))

    # Gráfico 1: Evolución de Ventas por Año
    plt.subplot(2, 1, 1)
    ventas_anuales = df_master.groupby('Año')['Total'].sum()
    ventas_anuales.plot(kind='bar', color='#004c99', edgecolor='black')
    plt.title('INGRESOS TOTALES POR AÑO (JUNTOZ - NETEADOS)', fontsize=14)
    plt.ylabel('Soles (S/)')
    plt.xticks(rotation=0)

    # Gráfico 2: Participación del TOP 1% VIP
    plt.subplot(2, 1, 2)
    concentracion = ltv_reporte.groupby('Segmento')['LTV_Acumulado'].sum()
    concentracion.plot(kind='pie', autopct='%1.1f%%', colors=['#66b3ff', '#ff9999'], startangle=140)
    plt.title('IMPACTO DEL SEGMENTO VIP EN LA VENTA TOTAL', fontsize=14)
    plt.ylabel('')

    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_FOLDER, 'dashboard_gerencial.png'))
    
    # Guardar CSV Detallado
    ltv_reporte.to_csv(os.path.join(OUTPUT_FOLDER, 'resultado_ltv_detallado.csv'), index=False, sep=';', encoding='utf-8-sig')
    
    print("✅ Proceso completado con éxito.")

if __name__ == "__main__":
    generar_reporte_gerencial()