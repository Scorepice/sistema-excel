from flask import Flask, render_template, request, send_file, redirect, flash, url_for, abort
import pandas as pd
import sqlite3
import os
import re
import sys
from fpdf import FPDF


def app_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.abspath(os.path.dirname(__file__))


def resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
    return os.path.join(base_path, relative_path)


BASE_DIR = app_base_dir()

app = Flask(
    __name__,
    template_folder=resource_path('templates'),
    static_folder=resource_path('static'),
)
app.secret_key = 'super_secreto_maritime'
DB_NAME = os.path.join(BASE_DIR, 'database.db')
DEFAULT_MODULE = 'rdm_abiertas'

MODULES = {
    'rdm_abiertas': {'nombre': 'RDM ABIERTAS'},
    'rdms': {'nombre': 'RDMs'},
    'estatus_importacion': {'nombre': 'ESTATUS IMPORTACION'},
    'oc_comprometidas': {'nombre': 'OC COMPROMETIDAS'},
    'ordenes_pendientes': {'nombre': 'ORDENES PENDIENTES'},
    'rdms_valoradas': {'nombre': 'RDMS valoradas'},
    'rdms_x_valorar': {'nombre': 'RDMS x VALORAR'},
    'consulta_oc': {'nombre': 'CONSULTA O C'},
}


def get_db_connection():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn


def normalizar_texto(valor):
    return str(valor).lower().replace('.', '').replace(' ', '').strip()


def table_name_for_module(modulo):
    if modulo == 'rdm_abiertas':
        return 'registros'
    return f"registros_{modulo}"


def get_module_or_404(modulo):
    if modulo not in MODULES:
        abort(404)
    return MODULES[modulo]


def table_exists(conn, table_name):
    cursor = conn.cursor()
    cursor.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name=?",
        (table_name,),
    )
    return cursor.fetchone() is not None


def detectar_tipo_campo(columna):
    nombre = normalizar_texto(columna)

    if 'estado' in nombre or 'esatad' in nombre:
        return 'estado_select'
    if 'prioridad' in nombre:
        return 'prioridad_select'
    if (
        'fecha' in nombre
        or 'descripci' in nombre
        or nombre == '#parte'
        or nombre == 'parte'
    ):
        return 'date'
    if any(token in nombre for token in ['cant', 'monto', 'valor', 'precio', 'total', 'desde', 'hasta']):
        return 'number'
    return 'text'


def construir_campos_formulario(columnas):
    return [{'name': col, 'kind': detectar_tipo_campo(col)} for col in columnas]


@app.context_processor
def inject_globals():
    return {'modulos': MODULES}


@app.route('/')
def selector_modulo():
    return render_template('module_selector.html', modules=MODULES)


@app.route('/modulo/<modulo>')
def index(modulo):
    modulo_info = get_module_or_404(modulo)
    return render_template('index.html', modulo=modulo, module_name=modulo_info['nombre'])


@app.route('/modulo/<modulo>/subir', methods=['POST'])
def subir_excel(modulo):
    modulo_info = get_module_or_404(modulo)
    archivo = request.files.get('archivo')

    if not archivo:
        flash('Selecciona un archivo para procesar.', 'warning')
        return redirect(url_for('index', modulo=modulo))

    try:
        df = pd.read_excel(archivo)
        df.columns = [' '.join(str(c).split()) for c in df.columns]

        for col in df.columns:
            col_limpia = normalizar_texto(col)
            if 'fecha' in col_limpia or 'descripci' in col_limpia or 'parte' in col_limpia:
                serie_fecha = pd.to_datetime(df[col], errors='coerce')
                if serie_fecha.notna().sum() > 0:
                    df[col] = serie_fecha.dt.strftime('%Y-%m-%d')

        conn = sqlite3.connect(DB_NAME)
        df.to_sql(table_name_for_module(modulo), conn, if_exists='append', index=False)
        conn.close()

        flash(f"Excel del modulo '{modulo_info['nombre']}' guardado con exito.", 'success')
        return redirect(url_for('ver_datos', modulo=modulo))
    except Exception as e:
        flash(f'Error al procesar el archivo: {str(e)}', 'danger')
        return redirect(url_for('index', modulo=modulo))


@app.route('/modulo/<modulo>/datos')
def ver_datos(modulo):
    modulo_info = get_module_or_404(modulo)
    table_name = table_name_for_module(modulo)

    conn = get_db_connection()
    try:
        if not table_exists(conn, table_name):
            flash(f"El modulo '{modulo_info['nombre']}' aun no tiene datos. Sube un Excel primero.", 'warning')
            return render_template(
                'datos.html',
                modulo=modulo,
                module_name=modulo_info['nombre'],
                filas=[],
                columnas=[],
            )

        filas = conn.execute(
            f'SELECT rowid, * FROM "{table_name}" ORDER BY rowid DESC LIMIT 500'
        ).fetchall()
        cursor = conn.cursor()
        cursor.execute(f'PRAGMA table_info("{table_name}")')
        columnas = [col[1] for col in cursor.fetchall()]

        return render_template(
            'datos.html',
            modulo=modulo,
            module_name=modulo_info['nombre'],
            filas=filas,
            columnas=columnas,
        )
    finally:
        conn.close()


@app.route('/modulo/<modulo>/dashboard')
def dashboard(modulo):
    modulo_info = get_module_or_404(modulo)
    table_name = table_name_for_module(modulo)

    data_vacia = {
        'labels': [],
        'ejecutadas': [],
        'por_ejecutar': [],
        'por_cotizar': [],
        'cotizadas': [],
    }

    inicio = request.args.get('inicio', '')
    fin = request.args.get('fin', '')

    conn = get_db_connection()
    try:
        if not table_exists(conn, table_name):
            flash(f"El modulo '{modulo_info['nombre']}' aun no tiene datos para graficas.", 'warning')
            return render_template(
                'dashboard.html',
                modulo=modulo,
                module_name=modulo_info['nombre'],
                data_estatus=data_vacia,
                data_depto={'labels': [], 'valores': [], 'colores': []},
                data_cruce={'labels': [], 'datasets': []},
                inicio=inicio,
                fin=fin,
            )

        df = pd.read_sql_query(f'SELECT * FROM "{table_name}"', conn)
    finally:
        conn.close()

    if df.empty:
        return render_template(
            'dashboard.html',
            modulo=modulo,
            module_name=modulo_info['nombre'],
            data_estatus=data_vacia,
            data_depto={'labels': [], 'valores': [], 'colores': []},
            data_cruce={'labels': [], 'datasets': []},
            inicio=inicio,
            fin=fin,
        )

    col_estado = next((c for c in df.columns if 'estad' in c.lower() or 'esatad' in c.lower()), None)
    col_fecha = next(
        (
            c
            for c in df.columns
            if 'descripci' in c.lower() or 'desde' in c.lower() or 'parte' in c.lower() or 'fecha' in c.lower()
        ),
        None,
    )
    col_depto = next((c for c in df.columns if 'departamento' in c.lower() or 'depto' in c.lower()), None)

    if not col_estado or not col_fecha or not col_depto:
        flash('No se detectaron columnas necesarias para graficas (Estado, Fecha y Departamento).', 'danger')
        return render_template(
            'dashboard.html',
            modulo=modulo,
            module_name=modulo_info['nombre'],
            data_estatus=data_vacia,
            data_depto={'labels': [], 'valores': [], 'colores': []},
            data_cruce={'labels': [], 'datasets': []},
            inicio=inicio,
            fin=fin,
        )

    df['Fecha_Real'] = pd.to_datetime(df[col_fecha], errors='coerce')

    if inicio:
        df = df[df['Fecha_Real'] >= pd.to_datetime(inicio)]
    if fin:
        df = df[df['Fecha_Real'] <= pd.to_datetime(fin)]

    if df.empty:
        return render_template(
            'dashboard.html',
            modulo=modulo,
            module_name=modulo_info['nombre'],
            data_estatus=data_vacia,
            data_depto={'labels': [], 'valores': [], 'colores': []},
            data_cruce={'labels': [], 'datasets': []},
            inicio=inicio,
            fin=fin,
        )

    df['Mes'] = df['Fecha_Real'].dt.strftime('%Y-%m').fillna('Sin Fecha')
    df[col_depto] = df[col_depto].fillna('SIN DEPTO').astype(str).str.strip().str.upper()
    df[col_estado] = df[col_estado].fillna('PENDIENTE').astype(str).str.strip().str.upper()

    deptos_unicos = df[col_depto].unique().tolist()
    paleta = ['#17659d', '#fd7e14', '#6f42c1', '#20c997', '#e83e8c', '#dc3545', '#0dcaf0', '#ffc107', '#28a745', '#6610f2']
    mapa_colores = {depto: paleta[i % len(paleta)] for i, depto in enumerate(deptos_unicos)}

    meses_unicos = sorted(df['Mes'].unique().tolist())
    data_estatus = {
        'labels': meses_unicos,
        'ejecutadas': [],
        'por_ejecutar': [],
        'por_cotizar': [],
        'cotizadas': [],
    }

    for mes in meses_unicos:
        subset = df[df['Mes'] == mes]
        estado_mes = subset[col_estado].fillna('').astype(str).str.upper()
        conteo_ejecutadas = int(estado_mes.str.contains('EJECUTADA').sum())
        conteo_por_cotizar = int(estado_mes.str.contains('POR COTIZAR').sum())
        conteo_cotizadas = int(estado_mes.str.contains('COTIZADA').sum())
        conteo_total = int(subset.shape[0])
        conteo_por_ejecutar = max(conteo_total - conteo_ejecutadas - conteo_por_cotizar - conteo_cotizadas, 0)

        data_estatus['ejecutadas'].append(conteo_ejecutadas)
        data_estatus['por_ejecutar'].append(conteo_por_ejecutar)
        data_estatus['por_cotizar'].append(conteo_por_cotizar)
        data_estatus['cotizadas'].append(conteo_cotizadas)

    depto_counts = df[col_depto].value_counts()
    colores_depto = [mapa_colores.get(d, '#cccccc') for d in depto_counts.index]
    data_depto = {
        'labels': depto_counts.index.astype(str).tolist(),
        'valores': depto_counts.values.tolist(),
        'colores': colores_depto,
    }

    pivot = pd.crosstab(df['Mes'], df[col_depto])
    data_cruce = {'labels': pivot.index.tolist(), 'datasets': []}
    for depto in pivot.columns:
        color_asignado = mapa_colores.get(depto, '#cccccc')
        data_cruce['datasets'].append(
            {
                'label': str(depto),
                'data': pivot[depto].tolist(),
                'backgroundColor': color_asignado,
                'borderColor': color_asignado,
                'borderWidth': 1,
            }
        )

    return render_template(
        'dashboard.html',
        modulo=modulo,
        module_name=modulo_info['nombre'],
        data_estatus=data_estatus,
        data_depto=data_depto,
        data_cruce=data_cruce,
        inicio=inicio,
        fin=fin,
    )


@app.route('/modulo/<modulo>/editar/<int:id>', methods=['GET', 'POST'])
def editar(modulo, id):
    modulo_info = get_module_or_404(modulo)
    table_name = table_name_for_module(modulo)

    conn = get_db_connection()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    if not table_exists(conn, table_name):
        conn.close()
        flash(f"El modulo '{modulo_info['nombre']}' aun no tiene tabla de datos.", 'warning')
        return redirect(url_for('index', modulo=modulo))

    cursor.execute(f'PRAGMA table_info("{table_name}")')
    columnas_db = [col[1] for col in cursor.fetchall()]

    if request.method == 'POST':
        datos_html = dict(request.form)
        datos_a_actualizar = {}

        for col_db in columnas_db:
            col_db_limpia = normalizar_texto(col_db)
            for key_html, valor in datos_html.items():
                key_html_limpia = normalizar_texto(key_html)
                if col_db_limpia == key_html_limpia:
                    datos_a_actualizar[col_db] = None if valor.strip() == '' else valor
                    break

        if datos_a_actualizar:
            set_clause = ', '.join([f'"{k}" = ?' for k in datos_a_actualizar.keys()])
            valores = list(datos_a_actualizar.values())
            valores.append(id)
            try:
                cursor.execute(f'UPDATE "{table_name}" SET {set_clause} WHERE rowid = ?', valores)
                conn.commit()
                flash('Registro actualizado correctamente.', 'success')
            except Exception as e:
                flash(f'Error al actualizar: {str(e)}', 'danger')
        else:
            flash('No se encontraron datos validos para actualizar.', 'warning')

        conn.close()
        return redirect(url_for('ver_datos', modulo=modulo))

    cursor.execute(f'SELECT rowid, * FROM "{table_name}" WHERE rowid = ?', (id,))
    registro = cursor.fetchone()
    conn.close()

    if registro is None:
        flash('El registro no existe.', 'warning')
        return redirect(url_for('ver_datos', modulo=modulo))

    campos = construir_campos_formulario(columnas_db)
    return render_template(
        'editar.html',
        modulo=modulo,
        module_name=modulo_info['nombre'],
        registro=registro,
        campos=campos,
    )


@app.route('/modulo/<modulo>/agregar', methods=['GET', 'POST'])
def agregar(modulo):
    modulo_info = get_module_or_404(modulo)
    table_name = table_name_for_module(modulo)

    conn = get_db_connection()
    cursor = conn.cursor()

    if not table_exists(conn, table_name):
        conn.close()
        flash(
            f"El modulo '{modulo_info['nombre']}' no tiene estructura aun. Sube un Excel primero.",
            'warning',
        )
        return redirect(url_for('index', modulo=modulo))

    cursor.execute(f'PRAGMA table_info("{table_name}")')
    columnas_info = cursor.fetchall()
    columnas_db = [col[1] for col in columnas_info]

    if request.method == 'POST':
        datos_html = dict(request.form)
        datos_a_guardar = {}

        for col_db in columnas_db:
            col_db_limpia = normalizar_texto(col_db)
            for key_html, valor in datos_html.items():
                key_html_limpia = normalizar_texto(key_html)
                if col_db_limpia == key_html_limpia:
                    datos_a_guardar[col_db] = None if valor.strip() == '' else valor
                    break

        if not datos_a_guardar:
            conn.close()
            flash('No se detectaron campos validos para guardar.', 'warning')
            return redirect(url_for('agregar', modulo=modulo))

        columnas_str = ', '.join([f'"{k}"' for k in datos_a_guardar.keys()])
        placeholders = ', '.join(['?' for _ in datos_a_guardar])
        valores = list(datos_a_guardar.values())

        try:
            cursor.execute(
                f'INSERT INTO "{table_name}" ({columnas_str}) VALUES ({placeholders})',
                valores,
            )
            conn.commit()
            flash('Registro guardado correctamente.', 'success')
        except Exception as e:
            flash(f'Error al guardar en la base de datos: {str(e)}', 'danger')
        finally:
            conn.close()

        return redirect(url_for('ver_datos', modulo=modulo))

    conn.close()
    campos = construir_campos_formulario(columnas_db)
    return render_template(
        'agregar.html',
        modulo=modulo,
        module_name=modulo_info['nombre'],
        campos=campos,
    )


@app.route('/modulo/<modulo>/eliminar/<int:id>')
def eliminar(modulo, id):
    get_module_or_404(modulo)
    table_name = table_name_for_module(modulo)

    conn = get_db_connection()
    try:
        if not table_exists(conn, table_name):
            flash('No hay tabla para eliminar registros en este modulo.', 'warning')
            return redirect(url_for('index', modulo=modulo))

        conn.execute(f'DELETE FROM "{table_name}" WHERE rowid = ?', (id,))
        conn.commit()
        flash('Registro eliminado.', 'danger')
    finally:
        conn.close()

    return redirect(url_for('ver_datos', modulo=modulo))


@app.route('/modulo/<modulo>/limpiar_base')
def limpiar_base(modulo):
    modulo_info = get_module_or_404(modulo)
    table_name = table_name_for_module(modulo)

    conn = get_db_connection()
    try:
        conn.execute(f'DROP TABLE IF EXISTS "{table_name}"')
        conn.commit()
        flash(f"Base del modulo '{modulo_info['nombre']}' limpiada.", 'success')
    except Exception:
        flash('No se pudo limpiar la base de este modulo.', 'danger')
    finally:
        conn.close()

    return redirect(url_for('index', modulo=modulo))


@app.route('/modulo/<modulo>/reporte/<tipo>')
def generar_reporte(modulo, tipo):
    modulo_info = get_module_or_404(modulo)
    table_name = table_name_for_module(modulo)

    conn = get_db_connection()
    try:
        if not table_exists(conn, table_name):
            flash(f"El modulo '{modulo_info['nombre']}' aun no tiene datos para exportar.", 'warning')
            return redirect(url_for('ver_datos', modulo=modulo))

        df = pd.read_sql_query(f'SELECT * FROM "{table_name}"', conn)
    finally:
        conn.close()

    busqueda = request.args.get('q', '').strip()
    if busqueda:
        mask = df.astype(str).apply(lambda x: x.str.contains(busqueda, case=False, na=False)).any(axis=1)
        df = df[mask]

    if df.empty:
        flash('No hay datos que coincidan con la busqueda para exportar.', 'warning')
        return redirect(url_for('ver_datos', modulo=modulo))

    if tipo == 'excel':
        ruta = os.path.join(BASE_DIR, f"reporte_{modulo}.xlsx")
        df.to_excel(ruta, index=False)
        return send_file(ruta, as_attachment=True)

    if tipo == 'pdf':
        pdf = FPDF(orientation='L', format='A3')
        pdf.add_page()

        ruta_logo = os.path.join('static', 'maritime_foot.png')
        if os.path.exists(ruta_logo):
            pdf.image(ruta_logo, x=15, y=3, w=40)

        pdf.set_font('Arial', style='B', size=16)
        titulo = (
            f"Reporte {modulo_info['nombre']}: '{busqueda}'"
            if busqueda
            else f"Reporte General {modulo_info['nombre']}"
        )
        pdf.cell(0, 15, txt=titulo, ln=True, align='C')
        pdf.ln(5)

        anchos = []
        for col in df.columns:
            nombre = re.sub(r'[\s\.]', '', str(col).lower())
            if nombre == 'um':
                anchos.append(10)
            elif nombre == 'rg':
                anchos.append(55)
            elif nombre == 'usuario':
                anchos.append(12)
            elif 'cant' in nombre:
                anchos.append(25)
            elif 'rdm' in nombre or 'prioridad' in nombre:
                anchos.append(16)
            elif nombre in ['desde', 'hasta']:
                anchos.append(14)
            elif nombre in ['fecha', 'parte', 'descripción', 'descripcion']:
                anchos.append(18)
            elif 'estado' in nombre or 'esatado' in nombre:
                anchos.append(22)
            elif 'codigosistemas' in nombre:
                anchos.append(22)
            elif 'departamento' in nombre:
                anchos.append(26)
            elif 'código' in nombre or 'codigo' in nombre:
                anchos.append(68)
            else:
                anchos.append(20)

        pdf.set_font('Arial', style='B', size=8)
        for i, col in enumerate(df.columns):
            texto_col = (
                str(col)
                .replace('“', '"')
                .replace('”', '"')
                .replace('‘', "'")
                .replace('’', "'")
                .replace('–', '-')
                .replace('—', '-')
            )
            texto_col = texto_col.encode('latin-1', errors='ignore').decode('latin-1')
            pdf.cell(anchos[i], 8, texto_col[: int(anchos[i] * 0.75)], border=1, align='C')
        pdf.ln()

        pdf.set_font('Arial', size=7)
        for _, row in df.iterrows():
            for i, val in enumerate(row):
                texto = str(val).replace(' 00:00:00', '')
                if texto in ['nan', 'None']:
                    texto = ''
                if texto.endswith('.0'):
                    texto = texto[:-2]

                texto_limpio = (
                    texto.replace('“', '"')
                    .replace('”', '"')
                    .replace('‘', "'")
                    .replace('’', "'")
                    .replace('–', '-')
                    .replace('—', '-')
                )
                texto_limpio = texto_limpio.encode('latin-1', errors='ignore').decode('latin-1')
                limite = int(anchos[i] * 0.65)
                pdf.cell(anchos[i], 7, texto_limpio[:limite], border=1)
            pdf.ln()

        ruta_pdf = os.path.join(BASE_DIR, f"reporte_{modulo}.pdf")
        pdf.output(ruta_pdf)
        return send_file(ruta_pdf, as_attachment=True)

    flash('Tipo de reporte no valido. Usa excel o pdf.', 'warning')
    return redirect(url_for('ver_datos', modulo=modulo))


# Redirecciones de compatibilidad para enlaces antiguos
@app.route('/datos')
def legacy_datos():
    return redirect(url_for('ver_datos', modulo=DEFAULT_MODULE))


@app.route('/dashboard')
def legacy_dashboard():
    return redirect(url_for('dashboard', modulo=DEFAULT_MODULE))


@app.route('/agregar')
def legacy_agregar():
    return redirect(url_for('agregar', modulo=DEFAULT_MODULE))


@app.route('/editar/<int:id>', methods=['GET', 'POST'])
def legacy_editar(id):
    return redirect(url_for('editar', modulo=DEFAULT_MODULE, id=id))


@app.route('/eliminar/<int:id>')
def legacy_eliminar(id):
    return redirect(url_for('eliminar', modulo=DEFAULT_MODULE, id=id))


@app.route('/limpiar_base')
def legacy_limpiar_base():
    return redirect(url_for('limpiar_base', modulo=DEFAULT_MODULE))


@app.route('/reporte/<tipo>')
def legacy_reporte(tipo):
    return redirect(url_for('generar_reporte', modulo=DEFAULT_MODULE, tipo=tipo, q=request.args.get('q', '')))


@app.route('/subir', methods=['POST'])
def legacy_subir():
    return redirect(url_for('index', modulo=DEFAULT_MODULE))


if __name__ == '__main__':
    debug_mode = os.environ.get('SISTEMA_EXCEL_DEBUG', '0') == '1'
    port = int(os.environ.get('SISTEMA_EXCEL_PORT', '5000'))
    app.run(host='127.0.0.1', port=port, debug=debug_mode)
