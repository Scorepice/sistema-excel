from flask import Flask, render_template, request, send_file, redirect, flash
import pandas as pd
import sqlite3
import os
from fpdf import FPDF

app = Flask(__name__)
app.secret_key = 'super_secreto_maritime'
DB_NAME = 'database.db'

def get_db_connection():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/subir', methods=['POST'])
def subir_excel():
    archivo = request.files['archivo']
    if archivo:
        df = pd.read_excel(archivo)
        if 'Desde' in df.columns:
            df['Desde'] = pd.to_datetime(df['Desde'], errors='coerce').dt.strftime('%Y-%m-%d')
        
        conn = sqlite3.connect(DB_NAME)
        df.to_sql('registros', conn, if_exists='append', index=False)
        conn.close()
        
        flash('¡Archivo Excel procesado y guardado con éxito!', 'success')
        return redirect('/datos')
    
    flash('Error al procesar el archivo.', 'danger')
    return redirect('/')

@app.route('/datos')
def ver_datos():
    try:
        conn = get_db_connection()
        filas = conn.execute("SELECT rowid, * FROM registros").fetchall()
        cursor = conn.cursor()
        cursor.execute("PRAGMA table_info(registros)")
        columnas = [col[1] for col in cursor.fetchall()]
        conn.close()
        return render_template('datos.html', filas=filas, columnas=columnas)
    except:
        return "La base de datos está vacía. Sube un Excel primero."

@app.route('/dashboard')
def dashboard():
    try:
        conn = get_db_connection()
        df = pd.read_sql_query("SELECT * FROM registros", conn)
        conn.close()
        
        columna_estado = 'Estado'
        columna_fecha = 'Descripción'
        
        if columna_estado in df.columns and columna_fecha in df.columns:
            # Extraemos Año y Mes
            df['Mes'] = pd.to_datetime(df[columna_fecha], errors='coerce').dt.strftime('%Y-%m')
            
            # 1. EJECUTADAS
            df_ejecutadas = df[df[columna_estado].astype(str).str.upper().str.contains('EJECUTADA', na=False)]
            df_ejec = df_ejecutadas['Mes'].value_counts().reset_index()
            df_ejec.columns = ['Mes', 'Total']
            df_ejec = df_ejec.sort_values('Mes')
            
            labels_ejec = df_ejec['Mes'].astype(str).tolist()
            valores_ejec = df_ejec['Total'].tolist()

            # 2. POR EJECUTAR
            df_por_ejecutar = df[~df[columna_estado].astype(str).str.upper().str.contains('EJECUTADA', na=False)]
            df_por = df_por_ejecutar['Mes'].value_counts().reset_index()
            df_por.columns = ['Mes', 'Total']
            df_por = df_por.sort_values('Mes')
            
            labels_por = df_por['Mes'].astype(str).tolist()
            valores_por = df_por['Total'].tolist()
        else:
            labels_ejec, valores_ejec, labels_por, valores_por = [], [], [], []

        return render_template('dashboard.html', 
                               labels_ejec=labels_ejec, valores_ejec=valores_ejec,
                               labels_por=labels_por, valores_por=valores_por)
    except Exception as e:
        return f"Aún no hay datos suficientes para graficar. Error: {e}"

@app.route('/editar/<int:id>', methods=['GET', 'POST'])
def editar(id):
    conn = get_db_connection()
    if request.method == 'POST':
        datos = dict(request.form)
        set_clause = ", ".join([f'"{k}" = ?' for k in datos.keys()])
        conn.execute(f"UPDATE registros SET {set_clause} WHERE rowid=?", list(datos.values()) + [id])
        conn.commit()
        conn.close()
        flash('Registro actualizado correctamente.', 'success')
        return redirect('/datos')
    
    fila = conn.execute("SELECT rowid, * FROM registros WHERE rowid=?", (id,)).fetchone()
    cursor = conn.cursor()
    cursor.execute("PRAGMA table_info(registros)")
    columnas = [col[1] for col in cursor.fetchall()]
    conn.close()
    return render_template('editar.html', fila=fila, columnas=columnas)

@app.route('/eliminar/<int:id>')
def eliminar(id):
    conn = get_db_connection()
    conn.execute("DELETE FROM registros WHERE rowid=?", (id,))
    conn.commit()
    conn.close()
    flash('Registro eliminado.', 'danger')
    return redirect('/datos')

@app.route('/limpiar_base')
def limpiar_base():
    try:
        conn = sqlite3.connect(DB_NAME)
        conn.execute("DROP TABLE IF EXISTS registros")
        conn.commit()
        conn.close()
        flash('Base de datos limpiada.', 'success')
    except Exception as e:
        pass
    return redirect('/')

@app.route('/reporte/<tipo>')
def generar_reporte(tipo):
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql_query("SELECT * FROM registros", conn)
    conn.close()

    busqueda = request.args.get('q', '').strip()
    if busqueda:
        mask = df.astype(str).apply(lambda x: x.str.contains(busqueda, case=False, na=False)).any(axis=1)
        df = df[mask]

    if df.empty:
        flash('No hay datos que coincidan con la búsqueda para exportar.', 'warning')
        return redirect('/datos')

    if tipo == 'excel':
        ruta = 'reporte_filtrado.xlsx'
        df.to_excel(ruta, index=False)
        return send_file(ruta, as_attachment=True)
    
    elif tipo == 'pdf':
        # Tamaño de hoja A3 Horizontal
        pdf = FPDF(orientation="L", format="A3")
        pdf.add_page()
        
        pdf.set_font("Arial", style="B", size=14)
        titulo = f"Reporte de RDM´S : '{busqueda}'" if busqueda else "Reporte de RDM´S General"
        pdf.cell(0, 10, txt=titulo, ln=True, align='C')
        pdf.ln(5)

        # 1. DISTRIBUCIÓN MILIMÉTRICA EXACTA
        anchos = []
        for col in df.columns:
            nombre = str(col).lower()
            if nombre in ['rg', 'um']:
                anchos.append(10)
            elif nombre in ['usuario']:
                anchos.append(14)
            elif nombre in ['cant. orig.', 'cant', 'rdm #', 'prioridad']: # Quitamos 'estado' de aquí
                anchos.append(18)
            elif nombre in ['desde', 'hasta', 'fecha', '# parte', 'descripción', 'descripcion']:
                anchos.append(20)
            elif nombre in ['estado', 'esatado']:
                # AQUÍ ESTÁ EL CAMBIO: Le damos 26 milímetros exclusivos a Estado
                anchos.append(26)
            elif 'codigo sistemas' in nombre:
                anchos.append(26)
            elif nombre in ['departamento']:
                anchos.append(28)
            elif nombre == 'código' or nombre == 'codigo':
                anchos.append(115) 
            else:
                anchos.append(20) 

        # Encabezados de la tabla
        pdf.set_font("Arial", style="B", size=8)
        for i, col in enumerate(df.columns):
            pdf.cell(anchos[i], 8, str(col)[:int(anchos[i]*0.6)], border=1, align='C')
        pdf.ln()

        # Rellenar los datos
        pdf.set_font("Arial", size=7)
        for _, row in df.iterrows():
            for i, val in enumerate(row):
                texto = str(val).replace(' 00:00:00', '')
                if texto == 'nan' or texto == 'None': 
                    texto = '' 
                
                # Freno anti-desbordamiento (0.55)
                limite = int(anchos[i] * 0.55) 
                pdf.cell(anchos[i], 7, texto[:limite], border=1)
            pdf.ln()

        pdf.output("reporte.pdf")
        return send_file("reporte.pdf", as_attachment=True)
# ¡ESTAS DOS LÍNEAS SON VITALES PARA QUE ARRANQUE!
if __name__ == '__main__':
    app.run(debug=True)