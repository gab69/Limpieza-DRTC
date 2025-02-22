from flask import Flask, render_template, request, send_file
import pandas as pd
from datetime import datetime
import os

app = Flask(__name__)

# Ruta principal
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Verificar si se ha subido un archivo
        if 'file' not in request.files:
            return "No se ha subido ningún archivo."

        file = request.files['file']

        # Verificar si el archivo tiene un nombre
        if file.filename == '':
            return "No se ha seleccionado ningún archivo."

        # Guardar el archivo temporalmente
        file_path = os.path.join('uploads', file.filename)
        file.save(file_path)

        # Procesar el archivo
        output_path = process_file(file_path)

        # Eliminar el archivo temporal
        os.remove(file_path)

        # Enviar el archivo procesado al usuario
        return send_file(output_path, as_attachment=True)

    return render_template('index.html')

def process_file(file_path):
    # Leer el archivo de Excel
    df = pd.read_excel(file_path)

    # Procesar los datos (tu código original)
    bd = df[['Unnamed: 1', 'Unnamed: 6', 'Unnamed: 7', 'Unnamed: 8']]
    b = bd.drop(0)
    b = b.dropna()

    # Salida tarde
    b['Unnamed: 1'] = pd.to_datetime(b['Unnamed: 1'])
    hora_objetivo = datetime.strptime('17:00:00', '%H:%M:%S')
    resultados_Starde = b[b['Unnamed: 1'].dt.time > hora_objetivo.time()]
    resultados_sin_duplicados_salida_tarde = resultados_Starde.drop_duplicates(subset=['Unnamed: 6'])

    # Entrada Tarde
    hora_inicial = datetime.strptime('13:30:00', '%H:%M:%S')
    hora_final = datetime.strptime('16:00:00', '%H:%M:%S')
    resultados_Etarde = b[(b['Unnamed: 1'].dt.time >= hora_inicial.time()) & (b['Unnamed: 1'].dt.time <= hora_final.time())]
    resultados_sin_duplicados_entrada_tarde = resultados_Etarde.drop_duplicates(subset=['Unnamed: 6'])

    # Salida de Mañana
    hora_inicial = datetime.strptime('12:30:00', '%H:%M:%S')
    hora_final = datetime.strptime('13:30:00', '%H:%M:%S')
    resultados_Smañana = b[(b['Unnamed: 1'].dt.time >= hora_inicial.time()) & (b['Unnamed: 1'].dt.time <= hora_final.time())]
    resultados_sin_duplicados_salida_mañana = resultados_Smañana.drop_duplicates(subset=['Unnamed: 6'])

    # Entrada Mañana
    hora_inicial = datetime.strptime('9:00:00', '%H:%M:%S')
    resultados_Emañana = b[(b['Unnamed: 1'].dt.time <= hora_inicial.time())]
    resultados_sin_duplicados_entrada_mañana = resultados_Emañana.drop_duplicates(subset=['Unnamed: 6'])

    # Renombra las columnas
    column_mapping_entrada_mañana = {
        'Unnamed: 6': 'ID',
        'Unnamed: 7': 'Nombre',
        'Unnamed: 8': 'Apellidos',
        'Unnamed: 1': 'Entrada Mañana'
    }
    column_mapping_salida_mañana = {'Unnamed: 6': 'ID', 'Unnamed: 1': 'Salida Mañana'}
    column_mapping_entrada_tarde = {'Unnamed: 6': 'ID', 'Unnamed: 1': 'Entrada Tarde'}
    column_mapping_salida_tarde = {'Unnamed: 6': 'ID', 'Unnamed: 1': 'Salida Tarde'}

    resultados_sin_duplicados_entrada_mañana = resultados_sin_duplicados_entrada_mañana.rename(columns=column_mapping_entrada_mañana)
    resultados_sin_duplicados_salida_mañana = resultados_sin_duplicados_salida_mañana.rename(columns=column_mapping_salida_mañana)
    resultados_sin_duplicados_entrada_tarde = resultados_sin_duplicados_entrada_tarde.rename(columns=column_mapping_entrada_tarde)
    resultados_sin_duplicados_salida_tarde = resultados_sin_duplicados_salida_tarde.rename(columns=column_mapping_salida_tarde)

    # Fusionar DataFrames
    nueva_tabla = resultados_sin_duplicados_entrada_mañana.merge(resultados_sin_duplicados_salida_mañana, on='ID', how='left')
    nueva_tabla = nueva_tabla.merge(resultados_sin_duplicados_entrada_tarde, on='ID', how='left')
    nueva_tabla = nueva_tabla.merge(resultados_sin_duplicados_salida_tarde, on='ID', how='left')

    # Limpiar y formatear datos
    nueva_tabla['Entrada Tarde'] = nueva_tabla['Entrada Tarde'].fillna('')
    nueva_tabla['Salida Tarde'] = nueva_tabla['Salida Tarde'].fillna('')
    nueva_tabla = nueva_tabla[['ID', 'Nombre', 'Apellidos', 'Entrada Mañana', 'Salida Mañana', 'Entrada Tarde', 'Salida Tarde']]

    columnas_de_tiempo = ['Entrada Mañana', 'Salida Mañana', 'Entrada Tarde', 'Salida Tarde']
    for columna in columnas_de_tiempo:
        nueva_tabla[columna] = nueva_tabla[columna].astype(str)

    nueva_tabla['Fecha'] = nueva_tabla['Entrada Mañana'].str[:10]
    for columna in columnas_de_tiempo:
        nueva_tabla[columna] = nueva_tabla[columna].str[-8:]

    nueva_tabla = nueva_tabla[['ID', 'Nombre', 'Apellidos', 'Entrada Mañana', 'Salida Mañana', 'Entrada Tarde', 'Salida Tarde', 'Fecha']]

    # Exportar a Excel con bordes negros y gruesos
    output_path = os.path.join('output', 'resultado.xlsx')
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        nueva_tabla.to_excel(writer, sheet_name='Hoja1', index=False)
        worksheet = writer.sheets['Hoja1']

        # Definir el estilo de borde negro y grueso
        border_format = writer.book.add_format({
            'border': 2,  # 2 = Borde grueso
            'border_color': '#000000'  # Color negro
        })

        # Aplicar bordes a todas las celdas
        for row in range(0, len(nueva_tabla) + 1):  # +1 para incluir el encabezado
            for col in range(len(nueva_tabla.columns)):
                worksheet.write(row, col, nueva_tabla.iat[row - 1, col] if row > 0 else nueva_tabla.columns[col], border_format)

        # Ajustar el ancho de las columnas
        for i, col in enumerate(nueva_tabla.columns):
            column_len = max(nueva_tabla[col].astype(str).str.len().max(), len(col)) + 2  # Ajusta el ancho mínimo
            worksheet.set_column(i, i, column_len)

    return output_path

if __name__ == '__main__':
    # Crear carpetas necesarias
    os.makedirs('uploads', exist_ok=True)
    os.makedirs('output', exist_ok=True)
    # Obtener el puerto de la variable de entorno o usar 5000 por defecto
    port = int(os.environ.get('PORT', 5000))
    # Escuchar en todas las interfaces de red (0.0.0.0) para desplegar
    app.run(host='0.0.0.0', port=port, debug=False)
    # Ejecutar la aplicación en local
    #app.run(debug=True)
    
    
    
    