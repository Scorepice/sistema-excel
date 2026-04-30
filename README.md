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

Al subir un Excel, la app lee todas las hojas del libro y las concatena para no dejar filas fuera.

## 8) Logica de deteccion de campos

Para formularios y dashboard, la app detecta columnas por coincidencia de texto normalizado (minusculas, sin puntos ni espacios).

Tipos de campo del formulario:

- Estado: si el nombre contiene estado o esatad.
- Prioridad: si contiene prioridad.
- Fecha: si contiene fecha, emision, desde o hasta.
- Numero: si contiene cant, monto, valor, precio, total, desde o hasta.
- Texto: cualquier otro.

Columnas minimas para graficas:

- Estado (contiene estad/esatad)
- Fecha (contiene fecha/emision/desde/hasta)
- Departamento (contiene departamento/depto)

Si faltan esas columnas, el dashboard muestra aviso y no grafica.

Importante en carga de Excel:

- Si la tabla del modulo ya existe, la app valida compatibilidad minima de columnas para evitar corrimientos de datos.
- Si el Excel trae columnas nuevas y es compatible, la app amplia la tabla automaticamente y conserva esas columnas.
- Si la compatibilidad es baja (archivo de otro modulo/layout), la carga se rechaza con mensaje de error.

## 9) Graficas del dashboard

Se generan 3 estructuras de datos:

- data_estatus:
  - Conteo mensual por todos los valores unicos de la columna Estado (sin agrupar categorias fijas).
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

## 12.2) Ejecucion portable en otra PC o USB

El proyecto ya incluye una version portable en `release/SistemaExcelPortable/` para ejecutarse en otra computadora sin instalar Python ni Visual Studio.

Flujo de uso:

1. Copiar la carpeta completa `release/SistemaExcelPortable/` a una memoria USB o a la PC destino.
2. Abrir la carpeta `SistemaExcelPortable` en la otra PC.
3. Ejecutar `Iniciar Sistema Excel.cmd`.
4. El sistema levanta un servidor local y abre `http://127.0.0.1:5000` en el navegador.
5. Para detenerlo, ejecutar `Detener Sistema Excel.cmd`.

La carpeta portable ya incluye:

- Python embebido.
- `app.py`.
- `templates/`.
- `static/`.
- `database.db`.
- Scripts de inicio y detencion.

## 12.3) Crear el EXE del instalador

Si quieres distribuir un instalador en formato `.exe`, el proyecto incluye el instalador grafico en `release/installer/`.

Orden recomendado de construccion:

1. Generar el paquete portable comprimido:

```powershell
Compress-Archive -Path .\release\SistemaExcelPortable\* -DestinationPath .\release\installer\SistemaExcelPortable_payload_v2.zip -Force
```

2. Compilar el instalador grafico con PyInstaller:

```powershell
pyinstaller --noconfirm --clean --onefile --name SistemaExcel_Instalador --add-data "release\installer\SistemaExcelPortable_payload_v2.zip;." release\installer\instalador_gui.py
```

3. El ejecutable resultante queda en `dist\SistemaExcel_Instalador.exe`.

4. Ese instalador copia la version portable a la carpeta que elijas, por ejemplo en una USB o en `%LOCALAPPDATA%\SistemaExcel`.

Notas importantes:

- El instalador no requiere que la PC destino tenga Python ni Visual Studio.
- Para una distribucion totalmente portable, tambien puedes copiar directamente `release\SistemaExcelPortable\` a la USB sin crear el instalador.
- Si cambias el nombre del ZIP del payload, actualiza el archivo `release/installer/SistemaExcel_Instalador.spec`.

## 12.4) Limites de carga y volumen (configurable)

La app ahora separa limite de almacenamiento, paginacion de tabla y tamano de archivo:

- MAX_UPLOAD_MB: tamano maximo por archivo en MB. Default: 100.
- MAX_STORED_ROWS_PER_MODULE: maximo de filas guardadas por modulo (se eliminan las mas antiguas si se supera). Default: 100000.
- MAX_TABLE_PAGE_SIZE: maximo de filas por pagina en /datos_json. Default: 5000.

Ejemplo en PowerShell antes de ejecutar:

```powershell
$env:MAX_UPLOAD_MB = "250"
$env:MAX_STORED_ROWS_PER_MODULE = "300000"
$env:MAX_TABLE_PAGE_SIZE = "5000"
python app.py
```

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

## 15) Crear ejecutable principal (.exe) de la aplicacion

Si quieres generar un `SistemaExcel.exe` que arranque la app Flask directamente en la PC donde se ejecute, puedes usar PyInstaller.

1. Instalar PyInstaller en el entorno virtual:

```powershell
pip install pyinstaller
```

2. Compilar desde la carpeta del proyecto:

```powershell
pyinstaller --noconfirm --clean --onefile --name SistemaExcel --add-data "templates;templates" --add-data "static;static" app.py
```

3. El archivo final queda en `dist\SistemaExcel.exe`.

4. Al ejecutarlo, la app levanta el servidor local y puedes abrir el navegador en `http://127.0.0.1:5000`.

Notas importantes:

- La app ya detecta `templates/` y `static/` tanto en modo normal como empaquetado.
- La base de datos `database.db` se crea junto al `.exe`.
- Los reportes Excel/PDF también se guardan junto al `.exe`.
- Si Windows Defender bloquea el `.exe`, agrega una excepción o firma el ejecutable.
