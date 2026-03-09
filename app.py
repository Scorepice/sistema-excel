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
        
        columnas_fecha = ['# parte', 'Descripción', 'Descripcion']
        for col in columnas_fecha:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%Y-%m-%d')
        
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
        # 🚀 MEJORA DE VELOCIDAD: Carga los últimos 500 para evitar que el buscador y la tabla colapsen
        filas = conn.execute("SELECT rowid, * FROM registros ORDER BY rowid DESC LIMIT 500").fetchall()
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
        # Búsqueda más inteligente para la fecha
        col_fecha = next((c for c in df.columns if 'descripci' in c.lower() or 'desde' in c.lower() or 'parte' in c.lower()), None)
        col_depto = next((c for c in df.columns if 'departamento' in c.lower()), None)
        
        if not col_estado or not col_fecha or not col_depto:
            return "<h3>⚠️ Faltan datos</h3><p>Asegúrate de tener las columnas: Estado, Descripción (Fecha) y Departamento.</p>"
            
        df['Fecha_Real'] = pd.to_datetime(df[col_fecha], errors='coerce')
        
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
        
        df[col_depto] = df[col_depto].fillna('SIN DEPTO').astype(str).str.strip().str.upper()
        df[col_estado] = df[col_estado].fillna('PENDIENTE').astype(str).str.strip().str.upper()

        deptos_unicos = df[col_depto].unique().tolist()
        paleta = ['#17659d', '#fd7e14', '#6f42c1', '#20c997', '#e83e8c', '#dc3545', '#0dcaf0', '#ffc107', '#28a745', '#6610f2']
        mapa_colores = {depto: paleta[i % len(paleta)] for i, depto in enumerate(deptos_unicos)}

        meses_unicos = sorted(df['Mes'].unique().tolist())
        mask_ejec = df[col_estado].str.contains('EJECUTADA')
        data_estatus = {'labels': meses_unicos, 'ejecutadas': [], 'por_ejecutar': []}
        for mes in meses_unicos:
            subset = df[df['Mes'] == mes]
            data_estatus['ejecutadas'].append(int(subset[mask_ejec].shape[0]))
            data_estatus['por_ejecutar'].append(int(subset[~mask_ejec].shape[0]))

        depto_counts = df[col_depto].value_counts()
        colores_depto = [mapa_colores.get(d, '#cccccc') for d in depto_counts.index]
        data_depto = {'labels': depto_counts.index.astype(str).tolist(), 'valores': depto_counts.values.tolist(), 'colores': colores_depto}

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
    
    # 1. Obtenemos las columnas exactas de la base de datos
    cursor.execute("PRAGMA table_info(registros)")
    columnas_db = [col[1] for col in cursor.fetchall()]

    if request.method == 'POST':
        datos_html = dict(request.form)
        datos_a_actualizar = {}
        
        # ==========================================
        # TRADUCTOR INTELIGENTE DE COLUMNAS (Para Editar)
        # ==========================================
        # Comparamos ignorando mayúsculas, puntos y espacios
        for col_db in columnas_db:
            col_db_limpia = col_db.lower().replace('.', '').replace(' ', '').strip()
            
            for key_html, valor in datos_html.items():
                key_html_limpia = key_html.lower().replace('.', '').replace(' ', '').strip()
                
                # Si hacen match, preparamos el dato para actualizar
                if col_db_limpia == key_html_limpia:
                    if valor.strip() == '':
                        datos_a_actualizar[col_db] = None 
                    else:
                        datos_a_actualizar[col_db] = valor
                    break 

        # Si el traductor logró emparejar datos, armamos la actualización
        if datos_a_actualizar:
            set_clause = ", ".join([f'"{k}" = ?' for k in datos_a_actualizar.keys()])
            valores = list(datos_a_actualizar.values())
            valores.append(id) # Agregamos el ID al final para el WHERE
            
            try:
                cursor.execute(f"UPDATE registros SET {set_clause} WHERE rowid = ?", valores)
                conn.commit()
                flash('¡RDM actualizada correctamente!', 'success')
            except Exception as e:
                print("\n" + "="*50)
                print(f"⚠️ ERROR AL ACTUALIZAR MANUALMENTE: {e}")
                print("="*50 + "\n")
                flash('Error al actualizar en la base de datos.', 'danger')
        else:
            flash('No se encontraron datos válidos para actualizar.', 'warning')

        conn.close()
        return redirect('/datos')

    # Si es GET (solo entrar a la vista)
    cursor.execute("SELECT rowid, * FROM registros WHERE rowid = ?", (id,))
    registro = cursor.fetchone()
    conn.close()
    
    if registro is None:
        flash('El registro no existe.', 'warning')
        return redirect('/datos')

    return render_template('editar.html', registro=registro, columnas=columnas_db)

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
        
    columnas_db = [col[1] for col in columnas_info]

    if request.method == 'POST':
        datos_html = dict(request.form)
        datos_a_guardar = {}
        
        # ==========================================
        # TRADUCTOR INTELIGENTE DE COLUMNAS (Nivel Dios)
        # ==========================================
        # Comparamos el HTML y la BD ignorando mayúsculas, puntos Y ESPACIOS
        for col_db in columnas_db:
            # Le agregamos .replace(' ', '') para destruir los espacios
            col_db_limpia = col_db.lower().replace('.', '').replace(' ', '').strip()
            
            for key_html, valor in datos_html.items():
                # Le agregamos .replace(' ', '') aquí también
                key_html_limpia = key_html.lower().replace('.', '').replace(' ', '').strip()
                
                # Si hacen "match" (ej: "cantorig" con "cantorig"), lo preparamos para guardar
                if col_db_limpia == key_html_limpia:
                    if valor.strip() == '':
                        datos_a_guardar[col_db] = None 
                    else:
                        datos_a_guardar[col_db] = valor
                    break

        # Construimos la orden SQL solo con los datos que sí existen en la BD
        columnas_str = ", ".join([f'"{k}"' for k in datos_a_guardar.keys()])
        placeholders = ", ".join(["?" for _ in datos_a_guardar])
        valores = list(datos_a_guardar.values())
        
        try:
            cursor.execute(f"INSERT INTO registros ({columnas_str}) VALUES ({placeholders})", valores)
            conn.commit()
            flash('¡RDM registrada y guardada correctamente!', 'success')
        except Exception as e:
            print("\n" + "="*50)
            print(f"⚠️ ERROR AL GUARDAR MANUALMENTE: {e}")
            print("="*50 + "\n")
            flash(f'Error al guardar en la base de datos.', 'danger')
        finally:
            conn.close()
            
        return redirect('/datos')
    
    conn.close()
    return render_template('agregar.html', columnas=columnas_db)
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
        pdf = FPDF(orientation="L", format="A3")
        pdf.add_page()
        
        # ==========================================
        # NUEVO: AGREGAR EL LOGO A LA ESQUINA DEL PDF
        # ==========================================
        # Verificamos que la imagen exista para evitar errores
        if os.path.exists('static/maritime_foot.png'):
            # x=15 (distancia desde la izquierda)
            # y=8 (distancia desde arriba)
            # w=25 (ancho de la imagen en milímetros)
            pdf.image('static/maritime_foot.png', x=15, y=3, w=40)
        
        # Mover el título un poco hacia abajo para que se alinee bonito con el logo
        pdf.set_font("Arial", style="B", size=16)
        titulo = f"Reporte de RDM´S : '{busqueda}'" if busqueda else "Reporte de RDM´S General"
        
        # Imprimimos el título centrado
        pdf.cell(0, 15, txt=titulo, ln=True, align='C')
        pdf.ln(5) # Espacio en blanco antes de que empiece la tabla

        anchos = []
        # 1. DISTRIBUCIÓN MILIMÉTRICA EXACTA (Inteligente y Ampliada)
        anchos = []
        # 1. DISTRIBUCIÓN MILIMÉTRICA EXACTA (Balanceada para MVP)
        anchos = []
        for col in df.columns:
            nombre = str(col).lower().replace(' ', '').replace('.', '')
            
            if nombre == 'um': anchos.append(10)
            elif nombre == 'rg': anchos.append(55)  # ⬆️ SÚPER GIGANTE: Subió de 45 a 60 mm
            elif nombre == 'usuario': anchos.append(12)
            elif 'cant' in nombre: anchos.append(25) # ⬇️ REDUCIDO: Bajó de 38 a 25 mm
            elif 'rdm' in nombre or 'prioridad' in nombre: anchos.append(16)
            elif nombre in ['desde', 'hasta']: anchos.append(14)
            elif nombre in ['fecha', 'parte', 'descripción', 'descripcion']: anchos.append(18)
            elif 'estado' in nombre or 'esatado' in nombre: anchos.append(22)
            elif 'codigosistemas' in nombre: anchos.append(22)
            elif 'departamento' in nombre: anchos.append(26)
            elif 'código' in nombre or 'codigo' in nombre: anchos.append(68) # ⬇️ Ajustado levemente
            else: anchos.append(20)

        # 2. ENCABEZADOS DE LA TABLA 
        pdf.set_font("Arial", style="B", size=8)
        for i, col in enumerate(df.columns):
            texto_col = str(col).replace('“', '"').replace('”', '"').replace('‘', "'").replace('’', "'").replace('–', '-').replace('—', '-')
            texto_col = texto_col.encode('latin-1', errors='ignore').decode('latin-1')
            pdf.cell(anchos[i], 8, texto_col[:int(anchos[i]*0.75)], border=1, align='C')
        pdf.ln()

        # 3. RELLENAR LOS DATOS (Con tijera calibrada anti-choques)
        pdf.set_font("Arial", size=7)
        for _, row in df.iterrows():
            for i, val in enumerate(row):
                texto = str(val).replace(' 00:00:00', '')
                if texto == 'nan' or texto == 'None': texto = '' 
                
                if texto.endswith('.0'):
                    texto = texto[:-2]
                
                texto_limpio = texto.replace('“', '"').replace('”', '"').replace('‘', "'").replace('’', "'").replace('–', '-').replace('—', '-')
                texto_limpio = texto_limpio.encode('latin-1', errors='ignore').decode('latin-1')
                
                # ✂️ TIJERA CALIBRADA A 0.65: Evita que las letras anchas invadan la siguiente columna
                limite = int(anchos[i] * 0.65) 
                pdf.cell(anchos[i], 7, texto_limpio[:limite], border=1)
            pdf.ln()

        # 4. EXPORTAR Y RETORNAR EL ARCHIVO AL USUARIO
        pdf.output("reporte.pdf")
        return send_file("reporte.pdf", as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)