from flask import Flask, render_template, request, send_file, redirect
import pandas as pd
import sqlite3
import os
from fpdf import FPDF

app = Flask(__name__)
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
        # Cambia 'Fecha' por el nombre real de tu columna de fechas en el Excel
        if 'Fecha' in df.columns:
            df['Fecha'] = pd.to_datetime(df['Fecha']).dt.strftime('%Y-%m-%d')
        
        conn = sqlite3.connect(DB_NAME)
        
        # ¡AQUÍ ESTÁ LA MAGIA! Cambiamos 'replace' por 'append'
        # Ahora los datos nuevos se sumarán a los viejos sin borrar nada
        df.to_sql('registros', conn, if_exists='append', index=False)
        
        conn.close()
        return redirect('/datos')
    return "Error al subir", 400

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
        
        # Cambia 'Fecha' y 'Rendimiento' por los nombres reales de tus columnas
        if 'Fecha' in df.columns and 'Rendimiento' in df.columns:
            df['Fecha'] = pd.to_datetime(df['Fecha'])
            df_mensual = df.groupby(df['Fecha'].dt.strftime('%Y-%m')).sum(numeric_only=True).reset_index()
            labels = df_mensual['Fecha'].tolist()
            valores = df_mensual['Rendimiento'].tolist()
        else:
            labels, valores = [], []
        return render_template('dashboard.html', labels=labels, valores=valores)
    except:
        return "Aún no hay datos suficientes para graficar."

@app.route('/editar/<int:id>', methods=['GET', 'POST'])
def editar(id):
    conn = get_db_connection()
    if request.method == 'POST':
        datos = dict(request.form)
        set_clause = ", ".join([f'"{k}" = ?' for k in datos.keys()])
        conn.execute(f"UPDATE registros SET {set_clause} WHERE rowid=?", list(datos.values()) + [id])
        conn.commit()
        conn.close()
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
    return redirect('/datos')

@app.route('/reporte/<tipo>')
def generar_reporte(tipo):
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql_query("SELECT * FROM registros", conn)
    conn.close()
    if tipo == 'excel':
        ruta = 'reporte_final.xlsx'
        df.to_excel(ruta, index=False)
        return send_file(ruta, as_attachment=True)
    else:
        pdf = FPDF(orientation="L")
        pdf.add_page()
        pdf.set_font("Arial", size=10)
        pdf.cell(200, 10, txt="Reporte General", ln=True, align='C')
        pdf.ln(10)
        for col in df.columns: pdf.cell(35, 10, str(col)[:10], 1)
        pdf.ln()
        for _, row in df.iterrows():
            for val in row: pdf.cell(35, 10, str(val)[:10], 1)
            pdf.ln()
        pdf.output("reporte.pdf")
        return send_file("reporte.pdf", as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)