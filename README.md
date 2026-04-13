# Sistema Excel Multi-Modulo

Aplicacion web en Flask para cargar archivos Excel por modulo, almacenar datos en SQLite, gestionar registros (CRUD), visualizar graficas y exportar reportes en Excel/PDF.

## 1) Caracteristicas principales

- Seleccion de modulo desde una pantalla inicial.
- Carga de Excel por modulo.
- Almacenamiento en SQLite con tablas separadas por modulo.
- Visualizacion tabular con DataTables (busqueda, scroll horizontal, ordenamiento).
- Alta manual, edicion y eliminacion de registros.
- Dashboard con filtros por fecha y varias vistas graficas.
- Exportacion de reportes a Excel y PDF (con filtro por texto opcional).
- Rutas legacy para compatibilidad con enlaces antiguos.

## 2) Estructura del proyecto

- app.py: aplicacion principal Flask y logica de negocio.
- requerimientos.txt: dependencias Python.
- templates/: vistas HTML (Jinja2).
- static/: estilos e imagenes.

## 3) Requisitos

- Python 3.10+ recomendado.
- pip.

Dependencias usadas:

- Flask==3.0.0
- pandas==2.1.1
- openpyxl==3.1.2
- fpdf2==2.7.5

## 4) Instalacion y ejecucion

### Windows (PowerShell)

1. Crear entorno virtual:

```powershell
python -m venv .venv
```

2. Activar entorno virtual:

```powershell
.\.venv\Scripts\Activate.ps1
```

3. Instalar dependencias:

```powershell
pip install -r requerimientos.txt
```

4. Ejecutar la app:

```powershell
python app.py
```

5. Abrir en navegador:

- http://127.0.0.1:5000

## 5) Como funciona la base de datos

La app usa un solo archivo SQLite llamado database.db.

No crea una base diferente por modulo. Lo que si hace es crear/usar una tabla distinta por modulo:

- rdm_abiertas -> tabla registros
- cualquier otro modulo -> tabla registros_<modulo>

Ejemplos:

- rdms -> registros_rdms
- estatus_importacion -> registros_estatus_importacion

Cuando se sube un Excel, pandas hace append sobre la tabla del modulo. Si la tabla no existe, SQLite/pandas la crea con base en las columnas del archivo.

## 6) Modulos disponibles

Definidos actualmente:

- rdm_abiertas (RDM ABIERTAS)
- rdms (RDMs)
- estatus_importacion (ESTATUS IMPORTACION)
- oc_comprometidas (OC COMPROMETIDAS)
- ordenes_pendientes (ORDENES PENDIENTES)
- rdms_valoradas (RDMS valoradas)
- rdms_x_valorar (RDMS x VALORAR)
- consulta_oc (CONSULTA O C)

## 7) Flujo funcional

1. Entrar al modulo.
2. Subir Excel o registrar manualmente.
3. Ver datos en tabla.
4. Editar o eliminar registros.
5. Ver dashboard con filtros de fecha.
6. Exportar reporte en Excel o PDF.

## 8) Logica de deteccion de campos

Para formularios y dashboard, la app detecta columnas por coincidencia de texto normalizado (minusculas, sin puntos ni espacios).

Tipos de campo del formulario:

- Estado: si el nombre contiene estado o esatad.
- Prioridad: si contiene prioridad.
- Fecha: si contiene fecha, descripci, parte o #parte.
- Numero: si contiene cant, monto, valor, precio, total, desde o hasta.
- Texto: cualquier otro.

Columnas minimas para graficas:

- Estado (contiene estad/esatad)
- Fecha (contiene descripci/desde/parte/fecha)
- Departamento (contiene departamento/depto)

Si faltan esas columnas, el dashboard muestra aviso y no grafica.

## 9) Graficas del dashboard

Se generan 3 estructuras de datos:

- data_estatus:
  - Conteo mensual por estado: Ejecutada, Por Ejecutar, Por Cotizar, Cotizada.
- data_depto:
  - Conteo total por departamento.
- data_cruce:
  - Tabla cruzada Mes vs Departamento.

Filtros:

- inicio (fecha desde)
- fin (fecha hasta)

## 10) Reportes

Ruta base por modulo:

- /modulo/<modulo>/reporte/excel
- /modulo/<modulo>/reporte/pdf

Parametro opcional de filtro:

- q: busca texto en cualquier columna (case-insensitive).

Salida:

- Excel: reporte_<modulo>.xlsx
- PDF: reporte_<modulo>.pdf

## 11) Rutas principales

- GET / : selector de modulo.
- GET /modulo/<modulo> : pantalla de carga del modulo.
- POST /modulo/<modulo>/subir : procesa Excel.
- GET /modulo/<modulo>/datos : tabla de datos.
- GET /modulo/<modulo>/dashboard : graficas.
- GET/POST /modulo/<modulo>/agregar : alta manual.
- GET/POST /modulo/<modulo>/editar/<id> : edicion.
- GET /modulo/<modulo>/eliminar/<id> : eliminacion.
- GET /modulo/<modulo>/limpiar_base : elimina la tabla del modulo.
- GET /modulo/<modulo>/reporte/<tipo> : exportacion.

Rutas legacy redirigen al modulo por defecto rdm_abiertas.

## 12) Notas de mantenimiento

- Clave secreta hardcodeada: conviene mover a variable de entorno (FLASK_SECRET_KEY).
- Modo debug activo en app.run(debug=True): desactivar en produccion.
- Si cambia el layout de columnas del Excel de un modulo, validar que coincidan los nombres esperados para dashboard.
- Limpiar base borra toda la tabla del modulo actual (operacion destructiva).

## 13) Solucion de problemas

- Error al leer Excel:
  - Verificar extension .xlsx/.xls y que openpyxl este instalado.
- Dashboard vacio:
  - Revisar columnas detectadas para Estado, Fecha y Departamento.
  - Revisar filtro de fechas aplicado.
- Error en PDF:
  - Revisar caracteres especiales en columnas/datos; el codigo ya aplica limpieza latin-1 basica.

## 14) Mejoras recomendadas

- Agregar autenticacion de usuarios.
- Migrar a SQLAlchemy + migraciones.
- Configuracion por entorno (.env).
- Pruebas unitarias e integracion.
- Paginar datos en servidor para tablas grandes.
