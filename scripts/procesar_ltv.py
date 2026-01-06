import pandas as pd
import os

# Rutas relativas para que funcionen tanto en tu PC como en GitHub
INPUT_FOLDER = 'data_pedidos'
OUTPUT_FOLDER = 'reportes'
FILES = ['Pedidos_2023.xlsx', 'Pedidos_2024.xlsx', 'Pedidos_2025.xlsx']

def generar_ltv():
    # Crear carpeta de reportes si no existe
    if not os.path.exists(OUTPUT_FOLDER):
        os.makedirs(OUTPUT_FOLDER)

    all_data = []
    
    for file in FILES:
        path = os.path.join(INPUT_FOLDER, file)
        if os.path.exists(path):
            print(f"Procesando: {file}")
            # Cargamos el Excel
            df = pd.read_excel(path)
            
            # Filtros solicitados por el negocio
            df = df[(df['Canal de venta'] == 'Juntoz') & (df['Tipo de documento de cliente'] == 'DNI')]
            estados_netos = ['Received', 'ReadyToShip', 'ReadyToPickUp', 'PendingToPickUp', 'InTransit', 'Confirmed']
            df = df[df['Estado de item'].isin(estados_netos)]
            
            # Limpieza de montos
            df['Total'] = pd.to_numeric(df['Total'].astype(str).str.replace(',', '.'), errors='coerce')
            all_data.append(df)

    if all_data:
        df_master = pd.concat(all_data, ignore_index=True)
        
        # --- CÁLCULO LTV ---
        # Agrupamos por DNI para obtener la foto completa del cliente
        ltv_reporte = df_master.groupby('Nro. de documento de cliente').agg({
            'Total': 'sum',               # Cuánto dinero ha dejado (LTV)
            'Nro. de orden': 'nunique',   # Cuántas veces ha comprado
            'Fecha de creación': ['min', 'max'] # Antigüedad y Recencia
        }).reset_index()
        
        # Renombrar columnas para claridad
        ltv_reporte.columns = ['DNI', 'LTV_Acumulado', 'Total_Pedidos', 'Primera_Compra', 'Ultima_Compra']
        
        # Guardar resultado
        ruta_salida = os.path.join(OUTPUT_FOLDER, 'resultado_ltv.csv')
        ltv_reporte.to_csv(ruta_salida, index=False, sep=';', encoding='utf-8-sig')
        print(f"✅ Reporte generado en: {ruta_salida}")
    else:
        print("❌ No se encontraron archivos en 'data_pedidos'")

if __name__ == "__main__":
    generar_ltv()