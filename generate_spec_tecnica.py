#!/usr/bin/env python3
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)


def add_table_row(table, col_widths, cells_data, bg_color, bold_first=False):
    row = table.add_row()
    for i, (cell, text, w) in enumerate(zip(row.cells, cells_data, col_widths)):
        set_cell_bg(cell, bg_color)
        cell.width = w
        para = cell.paragraphs[0]
        para.clear()
        para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r = para.add_run(text)
        r.bold = bold_first and i == 0
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(0x20, 0x20, 0x20)


def add_header_row(table, col_widths, headers, bg_color):
    row = table.add_row()
    for i, (cell, text, w) in enumerate(zip(row.cells, headers, col_widths)):
        set_cell_bg(cell, bg_color)
        cell.width = w
        para = cell.paragraphs[0]
        para.clear()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = para.add_run(text)
        r.bold = True
        r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)


# ── Colores ──────────────────────────────────────────────────────────────────
BG_HEADER   = '1F4E79'
BG_CAT      = 'D6E4F0'
BG_ROW      = 'FFFFFF'
BG_ROW_ALT  = 'F5F8FC'
BG_WARN     = 'FFF2CC'
BG_IN       = 'E2EFDA'
BG_OUT      = 'FCE4D6'
TXT_HDR     = (0x1F, 0x4E, 0x79)
TXT_CAT     = (0x1F, 0x4E, 0x79)
TXT_BLACK   = (0x20, 0x20, 0x20)
TXT_GREEN   = (0x37, 0x58, 0x23)
TXT_RED     = (0x9C, 0x00, 0x06)

# ── Construir documento ───────────────────────────────────────────────────────
doc = Document()

for section in doc.sections:
    section.top_margin    = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin   = Cm(2.5)
    section.right_margin  = Cm(2.5)

# ─── Portada ─────────────────────────────────────────────────────────────────
doc.add_paragraph()
doc.add_paragraph()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('ESPECIFICACION TECNICA')
r.bold = True; r.font.size = Pt(24)
r.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('Sistema de Prediccion de Consumo de Materiales')
r.bold = True; r.font.size = Pt(15)
r.font.color.rgb = RGBColor(0x2E, 0x75, 0xB6)

doc.add_paragraph()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('Cliente: Grupo Ramos — La Sirena')
r.font.size = Pt(13); r.font.color.rgb = RGBColor(0x40, 0x40, 0x40)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('Octubre 2026  |  Version 1.0')
r.font.size = Pt(11); r.font.color.rgb = RGBColor(0x70, 0x70, 0x70)

doc.add_page_break()

# ─── 1. Objetivo ─────────────────────────────────────────────────────────────
doc.add_heading('1. Objetivo del Proyecto', level=1)
doc.add_paragraph(
    'Implementar un sistema de prediccion de consumo de materiales para la cadena La Sirena '
    '(Grupo Ramos, Republica Dominicana) que permita anticipar necesidades de reabastecimiento '
    'con suficiente antelacion para actuar dentro del lead time del proveedor, incorporando '
    'las senales de negocio propias del mercado dominicano que los modelos estandar de SAP '
    'no pueden capturar.'
)
doc.add_paragraph(
    'El sistema se desplegara sobre SAP AI Core (BTP) con un backend SAP CAP (OData V4) y '
    'una interfaz de consulta basada en SAP Fiori Elements.'
)

# ─── 2. Alcance ──────────────────────────────────────────────────────────────
doc.add_heading('2. Alcance', level=1)

doc.add_heading('2.1  Dentro del Alcance', level=2)

scope_in = [
    ('Tienda',         'La Sirena — cadena unica cubierta en esta fase.'),
    ('Productos',      'Set especifico de SKUs a definir y entregar por Grupo Ramos '
                       '(ver Seccion 4 — Requerimientos del Cliente). El modelo se '
                       'entrenara y desplegara unicamente sobre estos productos.'),
    ('Horizonte de prediccion',
                       'Prediccion diaria para los proximos 7 dias calendario '
                       '(configurable antes del inicio del desarrollo).'),
    ('Senales de negocio RD',
                       'Regalia Pascual, quincenas (dia 15 y ultimo de cada mes), '
                       'feriados nacionales, temporada turistica, alertas de huracan '
                       '(jun–nov), Carnaval (febrero), regreso a clases (agosto), '
                       'Dia de las Madres, Navidad / Aguinaldo.'),
    ('Pipeline de entrenamiento',
                       'Proceso automatizable en SAP AI Core que puede re-ejecutarse '
                       'cuando se disponga de nuevos datos historicos.'),
    ('API de inferencia',
                       'Servicio de prediccion accesible via OData V4 desde el '
                       'backend SAP CAP.'),
    ('Interfaz de consulta',
                       'Lista de predicciones por material y fecha, generada con '
                       'SAP Fiori Elements (List Report).'),
]

t_in = doc.add_table(rows=0, cols=2)
t_in.style = 'Table Grid'
W_IN = [Cm(5.0), Cm(11.5)]
add_header_row(t_in, W_IN, ['Elemento', 'Descripcion'], BG_HEADER)
for idx, (elem, desc) in enumerate(scope_in):
    bg = BG_ROW if idx % 2 == 0 else BG_ROW_ALT
    add_table_row(t_in, W_IN, [elem, desc], bg, bold_first=True)

doc.add_paragraph()
doc.add_heading('2.2  Fuera del Alcance', level=2)

scope_out = [
    'Supermercados Nacional (excluida en esta fase).',
    'Productos fuera del set entregado por el cliente.',
    'Integracion con SAP MRP o IBP (no es parte de esta fase).',
    'Override manual de predicciones (requiere desarrollo adicional fuera de fase 1).',
    'Datos de ventas del POS como fuente primaria — el modelo opera sobre consumo e inventario.',
    'Soporte a demanda intermitente (productos con muy baja rotacion) — evaluable como extension.',
]
for item in scope_out:
    doc.add_paragraph(item, style='List Bullet')

# ─── 3. Arquitectura ──────────────────────────────────────────────────────────
doc.add_heading('3. Arquitectura de la Solucion', level=1)
doc.add_paragraph(
    'La solucion se compone de tres capas principales:'
)

arch_rows = [
    ('Capa de Entrenamiento',
     'SAP AI Core — Argo Workflows',
     'Ejecuta el pipeline de entrenamiento del modelo sobre los datos historicos. '
     'Produce como artefacto el modelo serializado y sus metadatos de precision '
     '(MAPE, MAE). Puede re-ejecutarse bajo demanda.'),
    ('Capa de Inferencia / Serving',
     'SAP AI Core — KServe',
     'Despliega el modelo entrenado como API REST (protocolo KServe V2). '
     'Recibe un payload con las features calculadas para cada SKU y devuelve '
     'la prediccion de consumo para el horizonte solicitado.'),
    ('Capa de Aplicacion',
     'SAP CAP (Node.js) — OData V4',
     'Backend que orquesta la llamada al modelo: lee el historico de consumo de '
     'SQLite/HANA, calcula las features de lag y calendario, llama al endpoint de '
     'inferencia y persiste las predicciones en la entidad MaterialPredictions.'),
    ('Capa de Presentacion',
     'SAP Fiori Elements',
     'Interfaz de usuario generada automaticamente desde el servicio OData. '
     'Muestra predicciones por material, fecha y cantidad estimada. '
     'La accion refreshPredictions permite actualizar el calculo bajo demanda.'),
]

t_arch = doc.add_table(rows=0, cols=3)
t_arch.style = 'Table Grid'
W_ARCH = [Cm(4.0), Cm(4.5), Cm(8.0)]
add_header_row(t_arch, W_ARCH, ['Capa', 'Tecnologia', 'Responsabilidad'], BG_HEADER)
for idx, (capa, tech, resp) in enumerate(arch_rows):
    bg = BG_ROW if idx % 2 == 0 else BG_ROW_ALT
    add_table_row(t_arch, W_ARCH, [capa, tech, resp], bg, bold_first=True)

doc.add_paragraph()
doc.add_paragraph(
    'Flujo end-to-end: Los datos historicos de consumo del cliente se cargan en un bucket '
    'S3-compatible registrado en SAP AI Core como artifact de entrada. El pipeline de '
    'entrenamiento lee ese artifact, genera el modelo y lo almacena como artifact de salida. '
    'El serving template monta el modelo en el endpoint de inferencia. El backend CAP calcula '
    'las features en tiempo real y llama al endpoint para obtener las predicciones, '
    'que se almacenan en la base de datos y se exponen via OData.'
)

# ─── 4. Requerimientos del Cliente ────────────────────────────────────────────
doc.add_heading('4. Requerimientos del Cliente (Prerrequisitos)', level=1)
doc.add_paragraph(
    'Antes del inicio del desarrollo, Grupo Ramos debe entregar los siguientes elementos. '
    'La ausencia de cualquiera de ellos bloquea el inicio de la fase de entrenamiento.'
)

prereqs = [
    ('OBLIGATORIO',
     'Lista de productos (SKUs)',
     'Listado definitivo de los materiales incluidos en el alcance, con: '
     'codigo interno, descripcion, unidad de medida y categoria. '
     'No se iniciara el entrenamiento sin este listado.',
     BG_OUT),
    ('OBLIGATORIO',
     'Historico de consumo',
     'Minimo 24 meses de consumo diario por SKU para la tienda La Sirena. '
     'Formato: fecha (YYYY-MM-DD), codigo_material, cantidad_consumida. '
     'Puede entregarse como CSV o mediante acceso a la fuente de datos.',
     BG_OUT),
    ('OBLIGATORIO',
     'Historico de inventario',
     'Inventario diario o semanal por SKU (stock_inicial, entradas, salidas, stock_final). '
     'Utilizado para calcular features de lag y detectar quiebres de stock pasados.',
     BG_OUT),
    ('RECOMENDADO',
     'Historico de promociones',
     'Fechas de inicio y fin de promociones por SKU. Mejora significativamente '
     'la precision del modelo al separar demanda organica de demanda promovida.',
     BG_WARN),
    ('RECOMENDADO',
     'Calendario de fechas especiales internas',
     'Si el cliente maneja fechas internas propias (ej. eventos de tienda, liquidaciones, '
     'cierres parciales), deben incluirse para que el modelo no las trate como ruido.',
     BG_WARN),
    ('INFORMATIVO',
     'Acceso a SAP AI Core (BTP)',
     'Instance de SAP AI Core activa con resource group configurado. '
     'El equipo de implementacion puede proveer guia de configuracion inicial.',
     BG_ROW_ALT),
]

t_pre = doc.add_table(rows=0, cols=4)
t_pre.style = 'Table Grid'
W_PRE = [Cm(3.0), Cm(3.5), Cm(10.0), Cm(0.0)]
add_header_row(t_pre, [Cm(2.5), Cm(3.5), Cm(10.5)],
               ['Prioridad', 'Entregable', 'Descripcion'], BG_HEADER)

for prio, entregable, desc, bg in prereqs:
    row = t_pre.add_row()
    cells = row.cells
    cols = [Cm(2.5), Cm(3.5), Cm(10.5)]
    data = [prio, entregable, desc]
    for i, (cell, text, w) in enumerate(zip(cells, data, cols)):
        set_cell_bg(cell, bg)
        cell.width = w
        para = cell.paragraphs[0]; para.clear()
        para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r = para.add_run(text)
        r.bold = (i == 0)
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(0x20, 0x20, 0x20)

# ─── 5. Datos y Features ──────────────────────────────────────────────────────
doc.add_heading('5. Datos y Diseno de Features', level=1)

doc.add_heading('5.1  Features de Lag (Serie Temporal)', level=2)
doc.add_paragraph(
    'Se calculan directamente del historico de consumo del cliente. '
    'Representan el consumo pasado del producto como predictor del consumo futuro.'
)
lag_features = [
    ('lag_7d',         'Consumo del mismo dia hace 7 dias'),
    ('lag_14d',        'Consumo del mismo dia hace 14 dias'),
    ('lag_28d',        'Consumo del mismo dia hace 28 dias'),
    ('rolling_4w_avg', 'Promedio movil de consumo de las ultimas 4 semanas'),
    ('rolling_4w_std', 'Desviacion estandar de consumo de las ultimas 4 semanas (opcional)'),
]
for col, desc in lag_features:
    p = doc.add_paragraph(style='List Bullet')
    r = p.add_run(col + ':  '); r.bold = True; r.font.size = Pt(9)
    r2 = p.add_run(desc); r2.font.size = Pt(9)

doc.add_paragraph()
doc.add_heading('5.2  Features de Calendario y Negocio RD', level=2)
doc.add_paragraph(
    'Calculadas automaticamente a partir de la fecha. Codifican los eventos '
    'que generan picos o caidas de consumo en el mercado dominicano.'
)
cal_features = [
    ('is_regalia_pascual', 'Semana de Semana Santa — bono obligatorio RD, mayor pico del ano'),
    ('is_quincena',        'Dia 15 o ultimo dia del mes — dia de pago'),
    ('is_holiday',         'Feriado nacional dominicano'),
    ('is_weekend',         'Sabado o domingo'),
    ('day_of_week',        'Dia de la semana (0=lunes ... 6=domingo)'),
    ('month',              'Mes del ano (1-12) — captura estacionalidad anual'),
    ('is_navidad',         'Diciembre — temporada navidena y aguinaldo'),
    ('is_dia_madres',      'Segunda semana de mayo — segundo pico comercial mas alto'),
    ('is_temporada_turistica', 'Diciembre a abril — alta afluencia turistica RD'),
    ('is_huracan_season',  'Junio a noviembre — temporada de ciclones tropicales'),
    ('is_carnaval',        'Febrero — Carnaval de Santiago y La Vega'),
    ('is_regreso_clases',  'Agosto — inicio de ano escolar'),
]
for col, desc in cal_features:
    p = doc.add_paragraph(style='List Bullet')
    r = p.add_run(col + ':  '); r.bold = True; r.font.size = Pt(9)
    r2 = p.add_run(desc); r2.font.size = Pt(9)

# ─── 6. Modelo Predictivo ──────────────────────────────────────────────────────
doc.add_heading('6. Modelo Predictivo', level=1)
doc.add_paragraph(
    'El tipo de modelo final se definira durante la fase de experimentacion, '
    'una vez se cuente con el historico de consumo del cliente. La arquitectura '
    'de la solucion (SAP AI Core + SAP CAP) es agnóstica al algoritmo utilizado '
    'y permite desplegar cualquier modelo serializable en Python (scikit-learn, '
    'XGBoost, LightGBM, redes neuronales, etc.).'
)

doc.add_heading('6.1  Criterios de Seleccion del Modelo', level=2)
criterios = [
    'Precision sobre el conjunto de test: MAPE (Mean Absolute Percentage Error) '
    'y MAE (Mean Absolute Error) sobre los ultimos 60 dias del historico.',
    'Capacidad de capturar interacciones entre features de calendario y lag.',
    'Tiempo de inferencia dentro del SLA del endpoint (< 2 segundos por batch).',
    'Interpretabilidad: se priorizaran modelos que ofrezcan feature importance '
    'para comunicar al cliente que variables impulsan el consumo.',
]
for item in criterios:
    doc.add_paragraph(item, style='List Bullet')

doc.add_paragraph()
doc.add_heading('6.2  Metricas de Evaluacion', level=2)

metricas = [
    ('MAPE',  'Mean Absolute Percentage Error',
     'Porcentaje promedio de error relativo. Umbral de aceptacion a definir '
     'con el cliente antes del inicio del proyecto (sugerido: < 15%).'),
    ('MAE',   'Mean Absolute Error',
     'Error absoluto promedio en unidades de consumo. '
     'Util para productos con alta varianza donde MAPE puede ser inestable.'),
    ('RMSE',  'Root Mean Squared Error',
     'Penaliza errores grandes. Indicador complementario.'),
    ('Bias',  'Error sistematico',
     'Detecta si el modelo sobreestima o subestima consistentemente, '
     'lo cual podria generar exceso o faltante de stock de forma sistematica.'),
]

t_met = doc.add_table(rows=0, cols=3)
t_met.style = 'Table Grid'
W_MET = [Cm(2.0), Cm(4.0), Cm(10.5)]
add_header_row(t_met, W_MET, ['Metrica', 'Nombre completo', 'Descripcion'], BG_HEADER)
for idx, (m, nombre, desc) in enumerate(metricas):
    bg = BG_ROW if idx % 2 == 0 else BG_ROW_ALT
    add_table_row(t_met, W_MET, [m, nombre, desc], bg, bold_first=True)

doc.add_paragraph()
doc.add_heading('6.3  Reentrenamiento', level=2)
doc.add_paragraph(
    'El pipeline de entrenamiento puede re-ejecutarse bajo demanda en SAP AI Core. '
    'Se recomienda reentrenar el modelo:'
)
retrain = [
    'Al incorporar nuevos productos al set de prediccion.',
    'Trimestralmente para incorporar el historico reciente.',
    'Ante cambios estructurales en el patron de consumo (nueva tienda abierta proxima, '
    'cambio en politica de precios, nueva competencia).',
]
for item in retrain:
    doc.add_paragraph(item, style='List Bullet')

# ─── 7. Entregables ───────────────────────────────────────────────────────────
doc.add_heading('7. Entregables del Proyecto', level=1)

entregables = [
    ('E1', 'Modelo entrenado y desplegado',
     'Modelo serializado registrado como artifact en SAP AI Core, con endpoint de '
     'inferencia activo y documentado. Incluye metadata de precision (MAPE, MAE) '
     'sobre el set de test.'),
    ('E2', 'Pipeline de entrenamiento',
     'Argo WorkflowTemplate en SAP AI Core, ejecutable bajo demanda. '
     'Incluye script de preprocesamiento, ingenieria de features y entrenamiento.'),
    ('E3', 'Backend SAP CAP',
     'Servicio OData V4 con entidades Materials, ConsumptionHistory y '
     'MaterialPredictions. Accion refreshPredictions que orquesta el calculo '
     'de features y la llamada al modelo.'),
    ('E4', 'Interfaz Fiori Elements',
     'Dos List Reports: historico de consumo por material y predicciones '
     'por material/fecha. Accion de refresco de predicciones disponible desde la UI.'),
    ('E5', 'Documentacion tecnica',
     'Descripcion de la arquitectura, variables de entorno requeridas, '
     'instrucciones de despliegue, guia de reentrenamiento y manual de operacion.'),
    ('E6', 'Analisis de feature importance',
     'Informe con las variables que mas explican el consumo por producto, '
     'incluyendo el peso relativo de las senales de negocio RD sobre los resultados.'),
]

t_ent = doc.add_table(rows=0, cols=3)
t_ent.style = 'Table Grid'
W_ENT = [Cm(1.5), Cm(4.0), Cm(11.0)]
add_header_row(t_ent, W_ENT, ['ID', 'Entregable', 'Descripcion'], BG_HEADER)
for idx, (eid, nombre, desc) in enumerate(entregables):
    bg = BG_ROW if idx % 2 == 0 else BG_ROW_ALT
    add_table_row(t_ent, W_ENT, [eid, nombre, desc], bg, bold_first=True)

# ─── 8. Supuestos y Restricciones ─────────────────────────────────────────────
doc.add_heading('8. Supuestos y Restricciones', level=1)

doc.add_heading('8.1  Supuestos', level=2)
supuestos = [
    'Grupo Ramos entregara la lista definitiva de SKUs y el historico de consumo '
    'antes del inicio de la fase de entrenamiento.',
    'El historico de consumo contiene al menos 24 meses de datos diarios '
    'por SKU para producir un modelo con precision aceptable.',
    'El cliente dispone de una instancia activa de SAP AI Core en BTP.',
    'El calendario de fechas especiales RD provisto en este documento es '
    'representativo. El cliente puede solicitar ajustes durante la fase de validacion.',
    'Los datos de consumo entregados corresponden a consumo real, no a ventas '
    'del POS (son conceptos distintos para productos de uso interno vs. venta al publico).',
]
for item in supuestos:
    doc.add_paragraph(item, style='List Bullet')

doc.add_paragraph()
doc.add_heading('8.2  Restricciones', level=2)
restricciones = [
    'El modelo se entrenara y validara unicamente sobre los SKUs incluidos en el '
    'set entregado por el cliente. Agregar productos nuevos requiere reentrenamiento.',
    'La tienda cubierta en esta fase es exclusivamente La Sirena. '
    'Extender la solucion a Supermercados Nacional es un alcance separado.',
    'El sistema predice consumo futuro; no genera automaticamente ordenes de compra. '
    'La integracion con los procesos de reabastecimiento es responsabilidad del cliente.',
    'La precision del modelo depende directamente de la calidad y completitud '
    'de los datos historicos entregados.',
]
for item in restricciones:
    doc.add_paragraph(item, style='List Bullet')

# ─── 9. Criterios de Aceptacion ───────────────────────────────────────────────
doc.add_heading('9. Criterios de Aceptacion', level=1)
doc.add_paragraph(
    'El proyecto se considera exitoso cuando se cumplen todos los criterios '
    'marcados como OBLIGATORIO. Los criterios DESEABLES son mejoras adicionales.'
)

doc.add_paragraph()

aceptacion = [
    ('OBLIGATORIO',
     'MAPE del modelo sobre set de test dentro del umbral acordado con el cliente '
     'antes del inicio del proyecto.',
     BG_OUT),
    ('OBLIGATORIO',
     'Endpoint de inferencia activo en SAP AI Core respondiendo predicciones '
     'en menos de 2 segundos para un batch de hasta 50 productos.',
     BG_OUT),
    ('OBLIGATORIO',
     'Servicio OData V4 disponible con entidades y accion refreshPredictions '
     'funcionales en el ambiente del cliente.',
     BG_OUT),
    ('OBLIGATORIO',
     'Interfaz Fiori Elements mostrando predicciones por material y fecha, '
     'con datos actualizables desde la UI.',
     BG_OUT),
    ('OBLIGATORIO',
     'Pipeline de reentrenamiento documentado y ejecutable de forma autonoma '
     'por el equipo del cliente.',
     BG_OUT),
    ('DESEABLE',
     'Informe de feature importance mostrando el peso de las senales de negocio '
     'RD (especialmente Regalia Pascual) en el modelo.',
     BG_WARN),
    ('DESEABLE',
     'Dashboard de metricas de precision (MAPE, MAE, bias) por producto, '
     'actualizable tras cada reentrenamiento.',
     BG_WARN),
]

t_ac = doc.add_table(rows=0, cols=2)
t_ac.style = 'Table Grid'
W_AC = [Cm(3.0), Cm(13.5)]
add_header_row(t_ac, W_AC, ['Tipo', 'Criterio'], BG_HEADER)
for tipo, criterio, bg in aceptacion:
    row = t_ac.add_row()
    for cell, text, w in zip(row.cells, [tipo, criterio], W_AC):
        set_cell_bg(cell, bg)
        cell.width = w
        para = cell.paragraphs[0]; para.clear()
        para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r = para.add_run(text)
        r.bold = False
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(0x20, 0x20, 0x20)

# ─── Guardar ──────────────────────────────────────────────────────────────────
output = '/home/user/projects/material-consumption-ai/EspecificacionTecnica_GrupoRamos_LaSirena.docx'
doc.save(output)
print(f'Documento guardado: {output}')
