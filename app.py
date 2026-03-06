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
        
        # Filtro Supremo: Destruye saltos de línea y espacios ocultos
        df.columns = [' '.join(str(c).split()) for c in df.columns]
        
        # =========================================================
        # CORRECCIÓN: Ahora sí le decimos cuáles son las verdaderas fechas
        # =========================================================
        columnas_fecha = ['# parte', 'Descripción', 'Descripcion']
        for col in columnas_fecha:
            if col in df.columns:
                # Convierte a fecha y deja en blanco si hay error
                df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%Y-%m-%d')
        
        # "Desde" y "Hasta" se quedan tal cual (como números)
        
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
        
        col_estado = next((c for c in df.columns if 'estad' in c.lower() or 'esatad' in c.lower()), None)
        col_fecha = next((c for c in df.columns if 'descripci' in c.lower() or 'desde' in c.lower()), None)
        col_depto = next((c for c in df.columns if 'departamento' in c.lower()), None)
        
        if not col_estado or not col_fecha or not col_depto:
            return "<h3>⚠️ Faltan datos</h3><p>Asegúrate de tener las columnas: Estado, Descripción (Fecha) y Departamento.</p>"
            
        # Convertimos la columna a fechas
        df['Fecha_Real'] = pd.to_datetime(df[col_fecha], errors='coerce')
        
        # Filtro de rango de fechas
        inicio = request.args.get('inicio', '')
        fin = request.args.get('fin', '')
        
        if inicio:
            df = df[df['Fecha_Real'] >= pd.to_datetime(inicio)]
        if fin:
            df = df[df['Fecha_Real'] <= pd.to_datetime(fin)]
            
        if df.empty:
            return render_template('dashboard.html', data_estatus={'labels':[],'ejecutadas':[],'por_ejecutar':[]}, 
                                   data_depto={'labels':[],'valores':[],'colores':[]}, 
                                   data_cruce={'labels':[],'datasets':[]}, inicio=inicio, fin=fin)

        df['Mes'] = df['Fecha_Real'].dt.strftime('%Y-%m').fillna('Sin Fecha')
        
        # ==========================================
        # EL FILTRO ANTI-DUPLICADOS (Limpieza de texto)
        # ==========================================
        # 1. Convierte a texto (por si acaso).
        # 2. .str.strip() -> Quita espacios accidentales al inicio y al final.
        # 3. .str.upper() -> Convierte todo a MAYÚSCULAS para que coincidan siempre.
        df[col_depto] = df[col_depto].fillna('SIN DEPTO').astype(str).str.strip().str.upper()
        df[col_estado] = df[col_estado].fillna('PENDIENTE').astype(str).str.strip().str.upper()

        # Asignación de colores
        deptos_unicos = df[col_depto].unique().tolist()
        paleta = ['#17659d', '#fd7e14', '#6f42c1', '#20c997', '#e83e8c', '#dc3545', '#0dcaf0', '#ffc107', '#28a745', '#6610f2']
        mapa_colores = {depto: paleta[i % len(paleta)] for i, depto in enumerate(deptos_unicos)}

        # 1. ESTATUS POR FECHA
        meses_unicos = sorted(df['Mes'].unique().tolist())
        mask_ejec = df[col_estado].str.contains('EJECUTADA')
        data_estatus = {'labels': meses_unicos, 'ejecutadas': [], 'por_ejecutar': []}
        for mes in meses_unicos:
            subset = df[df['Mes'] == mes]
            data_estatus['ejecutadas'].append(int(subset[mask_ejec].shape[0]))
            data_estatus['por_ejecutar'].append(int(subset[~mask_ejec].shape[0]))

        # 2. TOTAL POR DEPARTAMENTO
        depto_counts = df[col_depto].value_counts()
        colores_depto = [mapa_colores.get(d, '#cccccc') for d in depto_counts.index]
        data_depto = {'labels': depto_counts.index.astype(str).tolist(), 'valores': depto_counts.values.tolist(), 'colores': colores_depto}

        # 3. DEPARTAMENTO POR FECHA
        pivot = pd.crosstab(df['Mes'], df[col_depto])
        data_cruce = {'labels': pivot.index.tolist(), 'datasets': []}
        for depto in pivot.columns:
            color_asignado = mapa_colores.get(depto, '#cccccc')
            data_cruce['datasets'].append({
                'label': str(depto),
                'data': pivot[depto].tolist(),
                'backgroundColor': color_asignado,
                'borderColor': color_asignado,
                'borderWidth': 1
            })

        return render_template('dashboard.html', 
                               data_estatus=data_estatus, data_depto=data_depto, data_cruce=data_cruce,
                               inicio=inicio, fin=fin)
    except Exception as e:
        return f"<h3>⚠️ Error interno</h3><p>{str(e)}</p>"
    
      
@app.route('/editar/<int:id>', methods=['GET', 'POST'])
def editar(id):
    conn = get_db_connection()
    conn.row_factory = sqlite3.Row 
    cursor = conn.cursor()
    
    # Obtenemos las columnas dinámicamente
    cursor.execute("PRAGMA table_info(registros)")
    columnas = [col[1] for col in cursor.fetchall()]

    if request.method == 'POST':
        datos = dict(request.form)
        set_clause = ", ".join([f'"{k}" = ?' for k in datos.keys()])
        valores = list(datos.values())
        valores.append(id)
        
        try:
            cursor.execute(f"UPDATE registros SET {set_clause} WHERE rowid = ?", valores)
            conn.commit()
            flash('RDM actualizada correctamente.', 'success')
        except Exception as e:
            flash(f'Error al actualizar: {e}', 'danger')
        finally:
            conn.close()
            
        return redirect('/datos')

    # Buscamos los datos actuales para llenar el formulario
    cursor.execute("SELECT rowid, * FROM registros WHERE rowid = ?", (id,))
    registro = cursor.fetchone()
    conn.close()
    
    if registro is None:
        flash('El registro no existe.', 'warning')
        return redirect('/datos')

    # ¡ESTA ES LA LÍNEA CLAVE QUE ARREGLA EL ERROR!
    # Aquí le enviamos el 'registro' y las 'columnas' al HTML
    return render_template('editar.html', registro=registro, columnas=columnas)
@app.route('/agregar', methods=['GET', 'POST'])
def agregar():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("PRAGMA table_info(registros)")
    columnas_info = cursor.fetchall()
    
    if not columnas_info:
        conn.close()
        flash('La base de datos está vacía. Sube un Excel primero para crear la estructura.', 'warning')
        return redirect('/')
        
    columnas = [col[1] for col in columnas_info]

    if request.method == 'POST':
        datos = dict(request.form)
        columnas_str = ", ".join([f'"{k}"' for k in datos.keys()])
        placeholders = ", ".join(["?" for _ in datos])
        valores = list(datos.values())
        
        try:
            conn.execute(f"INSERT INTO registros ({columnas_str}) VALUES ({placeholders})", valores)
            conn.commit()
            flash('RDM registrada correctamente.', 'success')
        except Exception as e:
            flash(f'Error al guardar: {e}', 'danger')
        finally:
            conn.close()
            
        return redirect('/datos')
    
    conn.close()
    return render_template('agregar.html', columnas=columnas)

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