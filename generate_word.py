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


def set_cell_content(cell, label, note, txt_color, align=WD_ALIGN_PARAGRAPH.CENTER):
    para = cell.paragraphs[0]
    para.clear()
    para.alignment = align
    r = para.add_run(label)
    r.bold = True
    r.font.size = Pt(9)
    r.font.color.rgb = RGBColor(*txt_color)
    if note:
        r2 = para.add_run('\n' + note)
        r2.bold = False
        r2.italic = True
        r2.font.size = Pt(7)
        r2.font.color.rgb = RGBColor(*txt_color)


# ── Colores ──────────────────────────────────────────────────────────────────
BG_GREEN  = 'C6EFCE'
BG_RED    = 'FFC7CE'
BG_YELLOW = 'FFEB9C'
BG_HEADER = '1F4E79'
BG_CAT    = 'D6E4F0'
BG_WHITE  = 'FFFFFF'

TXT_GREEN  = (0x37, 0x58, 0x23)
TXT_RED    = (0x9C, 0x00, 0x06)
TXT_YELLOW = (0x74, 0x4B, 0x00)
TXT_CAT    = (0x1F, 0x4E, 0x79)
TXT_WHITE  = (0xFF, 0xFF, 0xFF)
TXT_BLACK  = (0x20, 0x20, 0x20)

# ── Datos de la tabla ─────────────────────────────────────────────────────────
# (tipo, capacidad, udf_status, udf_nota, custom_status, custom_nota)
# tipo: 'CAT' | 'ROW'
# status: 'SI' | 'NO' | 'PARCIAL'
TABLE_DATA = [
    ('CAT', 'CALENDARIO Y FECHAS ESPECIALES'),
    ('ROW', 'Feriados nacionales estandar',
        'SI',      'Via SAP Factory Calendar',
        'SI',      'Feature en entrenamiento'),
    ('ROW', 'Dias de pago / quincenas (15 y ultimo del mes)',
        'NO',      'El concepto no existe en UDF',
        'SI',      'Feature directo de entrenamiento'),
    ('ROW', 'Regalia Pascual (bono obligatorio — exclusivo RD)',
        'NO',      'No existe en ningun calendario SAP',
        'SI',      'Feature personalizado para RD'),
    ('ROW', 'Dia de las Madres (mayor pico comercial en RD)',
        'NO',      'No configurable como evento de demanda',
        'SI',      'Alta importancia esperada en el modelo'),
    ('ROW', 'Carnaval (Santiago, La Vega — febrero)',
        'NO',      'Sin soporte regional/cultural',
        'SI',      'Feature regional por ciudad/tienda'),
    ('ROW', 'Semana Santa',
        'PARCIAL', 'Solo como feriado, efecto de demanda parcial',
        'SI',      'Captura duracion + interaccion con Regalia Pascual'),
    ('ROW', 'Regreso a clases (agosto — inicio ano escolar RD)',
        'NO',      'No configurable',
        'SI',      'Feature de temporada escolar'),
    ('ROW', 'Aguinaldo / Sueldo 13 (diciembre)',
        'NO',      'El concepto no existe en UDF',
        'SI',      'Feature de ciclo de pago anual'),
    ('ROW', 'Independencia (27 feb) y Restauracion (16 ago)',
        'PARCIAL', 'Como feriado, sin efecto comercial modelado',
        'SI',      'Efecto de demanda aprendido del historico'),
    ('ROW', 'Black Friday / Cyber Monday',
        'NO',      'No configurable',
        'SI',      'Feature de temporada comercial'),
    ('ROW', 'Dia de Reyes (6 enero)',
        'NO',      'No existe como evento de demanda',
        'SI',      'Feature de temporada navidena extendida'),
    ('ROW', 'Amor y Amistad / San Valentin',
        'NO',      'No configurable',
        'SI',      'Feature de temporada comercial'),

    ('CAT', 'SENALES ECONOMICAS Y EXTERNAS'),
    ('ROW', 'Remesas diaspora (picos en Navidad y Dia de Madres)',
        'NO',      'Fuera del universo de datos SAP',
        'SI',      'Integrable via API Banco Central RD'),
    ('ROW', 'Temporada turistica alta (diciembre–abril)',
        'NO',      'Sin integracion con datos de turismo',
        'SI',      'Feature de temporada con efecto en empleo y consumo'),
    ('ROW', 'Alertas de huracan — compra masiva (critico RD)',
        'NO',      'Sin integracion meteorologica',
        'SI',      'Feature via API climatica, aplica jun–nov'),
    ('ROW', 'Datos de clima historico',
        'NO',      'No soportado',
        'SI',      'API de clima como feature de entrenamiento'),
    ('ROW', 'Indicadores economicos externos (Banco Central RD)',
        'NO',      'Solo consume datos SAP internos',
        'SI',      'Cualquier fuente externa integrable'),
    ('ROW', 'Cosecha azucarera (poder adquisitivo zonas caneras)',
        'NO',      'Sin integracion con ciclos agricolas',
        'SI',      'Feature de ciclo agricola regional'),

    ('CAT', 'PROMOCIONES'),
    ('ROW', 'Promociones con SAP Trade Promotion Mgmt (TPM)',
        'SI',      'Integracion nativa SAP CAR + TPM',
        'SI',      'Feature de promocion en entrenamiento'),
    ('ROW', 'Promociones sin SAP TPM',
        'NO',      'Depende del flujo de datos TPM',
        'SI',      'Cualquier fuente de datos de promocion'),
    ('ROW', 'Interaccion: promocion + fecha especial simultanea',
        'NO',      'Modelos estadisticos no capturan interacciones',
        'SI',      'RandomForest aprende la combinacion automaticamente'),
    ('ROW', 'Efecto de promocion por categoria de producto',
        'PARCIAL', 'Modelos lineales, efecto plano',
        'SI',      'Aprendido por categoria en entrenamiento'),

    ('CAT', 'CAPACIDADES DEL MODELO PREDICTIVO'),
    ('ROW', 'Interacciones no-lineales entre variables',
        'NO',      'Modelos estadisticos son fundamentalmente lineales',
        'SI',      'Fortaleza central de RandomForest'),
    ('ROW', 'Multiples estacionalidades simultaneas',
        'PARCIAL', 'Holt-Winters: solo una estacionalidad a la vez',
        'SI',      'Features de lag capturan multiples ciclos'),
    ('ROW', 'Demanda intermitente / esporadica',
        'SI',      'Modelo Croston disponible en UDF',
        'SI',      'Soportado con features adicionales'),
    ('ROW', 'Cold start — productos nuevos sin historico',
        'NO',      'Requiere historico minimo para ajustar el modelo',
        'SI',      'Generaliza desde productos similares via features'),
    ('ROW', 'Diferenciacion por tienda / region',
        'PARCIAL', 'Forecasta por SKU, no por SKU+tienda facilmente',
        'SI',      'Feature de tienda/region en entrenamiento'),
    ('ROW', 'Explicabilidad — feature importance para el negocio',
        'NO',      'Caja negra estadistica, sin desglose por variable',
        'SI',      '"Regalia Pascual explica 18% del error de abril"'),
    ('ROW', 'Aprendizaje cruzado entre categorias de producto',
        'NO',      'Cada serie modelada independientemente',
        'SI',      'El modelo aprende patrones transversales'),
    ('ROW', 'Diferenciacion del efecto por tipo de producto',
        'NO',      'Mismo modelo para todos los SKUs',
        'SI',      'Perecederos, electronica y ropa tienen patrones distintos'),

    ('CAT', 'ASPECTOS OPERACIONALES'),
    ('ROW', 'Soporte SAP estandar (OSS, upgrades, parches)',
        'SI',      'Soporte oficial SAP incluido en licencia',
        'NO',      'Responsabilidad del equipo de desarrollo'),
    ('ROW', 'Integracion nativa con MM / MRP / IBP',
        'SI',      'Flujos estandar sin desarrollo adicional',
        'PARCIAL', 'Integracion custom via OData/API — ya implementada'),
    ('ROW', 'Reentrenamiento bajo demanda con nuevos datos',
        'NO',      'Ciclos fijos de actualizacion de modelos SAP',
        'SI',      'Pipeline automatizable en SAP AI Core'),
    ('ROW', 'Ingesta de datos externos (APIs, CSV, web)',
        'NO',      'Solo datos del ecosistema SAP',
        'SI',      'Cualquier fuente de datos integrable'),
    ('ROW', 'Governance / override manual por planificador',
        'SI',      'Workflow estandar SAP de aprobacion',
        'PARCIAL', 'Implementable, requiere desarrollo adicional'),
    ('ROW', 'KPIs de accuracy incorporados (MAPE, MAE, etc.)',
        'SI',      'Dashboard estandar SAP CAR',
        'SI',      'Calculados en pipeline + dashboard CAP'),
    ('ROW', 'Escalabilidad enterprise (miles de SKUs)',
        'SI',      'Disenado para escala retail enterprise',
        'SI',      'Escala con infraestructura SAP AI Core'),

    ('CAT', 'COSTO Y LICENCIAMIENTO'),
    ('ROW', 'Sin costo adicional si el cliente ya tiene SAP CAR',
        'SI',      'Funcionalidad incluida en licencia CAR',
        'NO',      'Requiere licencia SAP AI Core adicional'),
    ('ROW', 'Aplicable si el cliente NO tiene SAP CAR',
        'NO',      'Requiere adquirir licencia SAP CAR',
        'SI',      'Solo SAP AI Core — menor inversion inicial'),
    ('ROW', 'Control total sobre logica y features del modelo',
        'NO',      'Modelo fijo de SAP, no customizable',
        'SI',      'Total flexibilidad en ingenieria de features'),
]

# ── Construir documento ───────────────────────────────────────────────────────
doc = Document()

for section in doc.sections:
    section.top_margin    = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin   = Cm(2.5)
    section.right_margin  = Cm(2.5)

# ─── Portada ─────────────────────────────────────────────────────────────────
doc.add_paragraph()
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('ANALISIS COMPARATIVO')
r.bold = True; r.font.size = Pt(22)
r.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('SAP CAR UDF  vs  Modelo de Prediccion Custom en SAP AI Core')
r.bold = True; r.font.size = Pt(15)
r.font.color.rgb = RGBColor(0x2E, 0x75, 0xB6)

doc.add_paragraph()
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('Cliente: Grupo Ramos — Republica Dominicana')
r.font.size = Pt(12); r.font.color.rgb = RGBColor(0x40, 0x40, 0x40)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('Proyecto: Sistema de Prediccion de Consumo de Materiales')
r.font.size = Pt(11); r.font.color.rgb = RGBColor(0x60, 0x60, 0x60)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('Octubre 2026')
r.font.size = Pt(11); r.font.color.rgb = RGBColor(0x60, 0x60, 0x60)

doc.add_page_break()

# ─── 1. Contexto ─────────────────────────────────────────────────────────────
doc.add_heading('1. Contexto', level=1)
doc.add_paragraph(
    'Grupo Ramos opera La Sirena y Supermercados Nacional, dos de las cadenas de gran superficie '
    'mas importantes de Republica Dominicana. El presente documento evalua dos alternativas '
    'tecnologicas para implementar un sistema de prediccion de consumo y demanda de materiales: '
    'la funcionalidad estandar SAP CAR UDF (Unified Demand Forecast) y un modelo de Machine '
    'Learning custom desplegado sobre SAP AI Core con backend SAP CAP (Node.js).'
)

# ─── 2. SAP CAR UDF ───────────────────────────────────────────────────────────
doc.add_heading('2. SAP CAR UDF — Unified Demand Forecast', level=1)
doc.add_paragraph(
    'UDF es la funcionalidad de forecasting estadistico de SAP Customer Activity Repository (CAR). '
    'Utiliza modelos de series de tiempo como Holt-Winters, Croston (demanda esporadica) y regresiones '
    'lineales. Se integra nativamente con el ecosistema SAP (MM, MRP, IBP) y opera sobre los datos '
    'que fluyen a traves de los canales estandar de SAP.'
)
p = doc.add_paragraph(); p.add_run('Fortalezas principales:').bold = True
for item in [
    'Soporte oficial SAP: mantenimiento, upgrades y notas OSS incluidas.',
    'Integracion nativa con MM, MRP e IBP sin desarrollo adicional.',
    'Madurez operacional: manejo de outliers, model selection, workflows de override.',
    'KPIs de accuracy y dashboards incorporados.',
    'Disenado para escala enterprise con miles de SKUs.',
]:
    doc.add_paragraph(item, style='List Bullet')

# ─── 3. Modelo Custom ─────────────────────────────────────────────────────────
doc.add_heading('3. Modelo Custom — RandomForest en SAP AI Core', level=1)
doc.add_paragraph(
    'El modelo custom utiliza RandomForest (scikit-learn) entrenado y desplegado sobre SAP AI Core '
    '(Argo Workflows para entrenamiento, KServe para serving). El backend es SAP CAP con servicio '
    'OData V4. La arquitectura permite incorporar cualquier senal de negocio como feature de '
    'entrenamiento, capturando interacciones no-lineales imposibles de modelar con estadistica clasica.'
)
p = doc.add_paragraph(); p.add_run('Fortalezas principales:').bold = True
for item in [
    'Features de negocio propios: quincenas, Regalia Pascual, alertas de huracan, remesas, temporada turistica, etc.',
    'Modelo no-lineal: captura interacciones entre variables que UDF no puede detectar.',
    'Explicabilidad (feature importance): cuantifica el impacto de cada senal de negocio.',
    'Reentrenamiento bajo demanda con nuevos datos o nuevas features.',
    'Integracion de datos externos de cualquier fuente: APIs de clima, indicadores economicos, promociones.',
    'Diferenciacion geografica por tienda o region dentro del mismo modelo.',
    'Aprendizaje de comportamientos distintos por categoria de producto.',
]:
    doc.add_paragraph(item, style='List Bullet')

# ─── 4. Diferencia fundamental de orientacion ─────────────────────────────────
doc.add_heading('4. Diferencia de Orientacion', level=1)
doc.add_paragraph(
    'Ambas soluciones generan forecasts hacia el futuro y pueden alimentar procesos proactivos '
    'de reabastecimiento. Sin embargo, difieren en su fuente de datos, arquitectura y el tipo '
    'de pregunta de negocio que responden.'
)
p = doc.add_paragraph()
p.add_run('UDF — Forecast basado en ventas del POS: ').bold = True
p.add_run(
    'UDF se alimenta de los datos de venta que salen en caja. Genera un forecast de demanda futura '
    'que, conectado con MRP e IBP, puede disparar ordenes de compra preventivas. '
    'Adicionalmente genera alertas cuando las ventas reales se desvian del forecast — '
    'ese mecanismo especifico si es reactivo (el evento ya ocurrio en caja), pero el forecast '
    'en si mismo es prospectivo. Su limitacion no es la direccion temporal sino la calidad '
    'de las senales que puede incorporar.'
)
p = doc.add_paragraph()
p.add_run('Modelo Custom — Forecast basado en consumo e inventario proyectado: ').bold = True
p.add_run(
    'El modelo predice consumo futuro incorporando senales de negocio propias (Regalia Pascual, '
    'quincenas, temporada turistica, alertas de huracan). Genera directamente una proyeccion de '
    'insuficiencia de inventario antes de que ocurra el desabastecimiento, con suficiente '
    'anticipacion para activar una orden de compra dentro del lead time del proveedor. '
    'La pregunta que responde es: "en cuantos dias nos quedaremos sin stock de este producto '
    'si no compramos hoy" — y lo hace con mejor precision porque conoce los eventos que '
    'mueven la demanda en el mercado dominicano.'
)
p = doc.add_paragraph()
p.add_run('Nota sobre la cadena SAP: ').bold = True
p.add_run(
    'Para que UDF sea proactivo requiere la cadena completa: UDF conectado con MRP e IBP. '
    'El modelo custom opera de forma mas directa integrando CAP con el sistema de compras. '
    'Ninguno de estos enfoques es superior en si mismo — son arquitecturas distintas con '
    'diferentes requisitos de infraestructura SAP.'
)
doc.add_paragraph()

# ─── 5. Argumento clave ───────────────────────────────────────────────────────
doc.add_heading('5. El Argumento Clave — La Regalia Pascual', level=1)
doc.add_paragraph(
    'La Regalia Pascual es un bono salarial obligatorio establecido por ley dominicana, '
    'pagado durante la Semana Santa. Genera un pico de consumo predecible y recurrente en todas '
    'las tiendas de gran superficie de la Republica Dominicana.'
)
p = doc.add_paragraph()
r = p.add_run('Este evento no existe en ningun calendario SAP estandar. ')
r.bold = True
p.add_run(
    'UDF lo veria como ruido o como un outlier historico sin explicacion. '
    'El modelo custom, con una unica columna is_regalia_pascual en el dataset de entrenamiento, '
    'aprende y anticipa este pico con exactitud. '
    'Este solo argumento representa una limitacion estructural de UDF que no puede resolverse '
    'con configuracion estandar. Lo mismo aplica para las quincenas, las remesas de la diaspora '
    'y las alertas de huracan — senales completamente fuera del universo de UDF.'
)

doc.add_paragraph()

# ─── 5. Tabla comparativa ─────────────────────────────────────────────────────
doc.add_heading('6. Tabla Comparativa de Capacidades', level=1)
p = doc.add_paragraph()
p.add_run('La tabla compara ambas soluciones en seis dimensiones. ')
for label, color in [('Verde = Soportado nativo. ', TXT_GREEN),
                     ('Rojo = No soportado. ',      TXT_RED),
                     ('Amarillo = Parcial / con esfuerzo adicional.', TXT_YELLOW)]:
    r = p.add_run(label); r.bold = True
    r.font.color.rgb = RGBColor(*color)

doc.add_paragraph()

table = doc.add_table(rows=0, cols=3)
table.style = 'Table Grid'
COL_W = [Cm(8.0), Cm(3.8), Cm(4.7)]

# Cabecera
hdr_row = table.add_row()
for i, (cell, text) in enumerate(zip(hdr_row.cells,
        ['Capacidad / Caracteristica', 'SAP CAR UDF', 'Modelo Custom\n(SAP AI Core)'])):
    set_cell_bg(cell, BG_HEADER)
    para = cell.paragraphs[0]; para.clear()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = para.add_run(text); r.bold = True
    r.font.size = Pt(10); r.font.color.rgb = RGBColor(*TXT_WHITE)
    cell.width = COL_W[i]

STATUS_MAP = {
    'SI':      (BG_GREEN,  TXT_GREEN,  '✓'),
    'NO':      (BG_RED,    TXT_RED,    '✗'),
    'PARCIAL': (BG_YELLOW, TXT_YELLOW, '~'),
}

for item in TABLE_DATA:
    if item[0] == 'CAT':
        row = table.add_row()
        cells = row.cells
        cells[0].merge(cells[2])
        set_cell_bg(cells[0], BG_CAT)
        para = cells[0].paragraphs[0]; para.clear()
        para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r = para.add_run(item[1]); r.bold = True
        r.font.size = Pt(9); r.font.color.rgb = RGBColor(*TXT_CAT)
    else:
        _, cap, udf_s, udf_n, cst_s, cst_n = item
        row = table.add_row()
        cells = row.cells
        for i, w in enumerate(COL_W):
            cells[i].width = w

        # Columna capacidad
        set_cell_bg(cells[0], BG_WHITE)
        para = cells[0].paragraphs[0]; para.clear()
        para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r = para.add_run(cap); r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(*TXT_BLACK)

        # Columna UDF
        bg, tc, lbl = STATUS_MAP[udf_s]
        set_cell_bg(cells[1], bg)
        set_cell_content(cells[1], lbl, udf_n, tc)

        # Columna modelo custom
        bg, tc, lbl = STATUS_MAP[cst_s]
        set_cell_bg(cells[2], bg)
        set_cell_content(cells[2], lbl, cst_n, tc)

# ─── 6. Conclusiones ──────────────────────────────────────────────────────────
doc.add_paragraph()
doc.add_paragraph()
doc.add_heading('7. Lo que UDF Nunca Podra Hacer — Limitaciones Estructurales', level=1)
doc.add_paragraph(
    'Las siguientes capacidades son estructuralmente imposibles para UDF independientemente '
    'del nivel de configuracion, de contar con SAP TPM, MRP, IBP u otros modulos adicionales. '
    'No son brechas de implementacion — son limitaciones inherentes a la arquitectura y '
    'los algoritmos estadisticos que UDF utiliza por diseno.'
)
doc.add_paragraph()

struct_table = doc.add_table(rows=0, cols=3)
struct_table.style = 'Table Grid'

BG_STRUCT_HDR = '7B2D00'
BG_STRUCT_UDF = 'FFC7CE'
BG_STRUCT_CST = 'C6EFCE'
TXT_STRUCT_HDR = (0xFF, 0xFF, 0xFF)
TXT_STRUCT_UDF = (0x9C, 0x00, 0x06)
TXT_STRUCT_CST = (0x37, 0x58, 0x23)
TXT_STRUCT_CAP = (0x20, 0x20, 0x20)

S_COL_W = [Cm(4.5), Cm(7.0), Cm(5.0)]

# Cabecera
sh = struct_table.add_row()
for i, (cell, text) in enumerate(zip(sh.cells, [
        'Capacidad',
        'UDF — configuracion maxima\n(con TPM y todos los modulos SAP)',
        'Modelo Custom (SAP AI Core)'])):
    set_cell_bg(cell, BG_STRUCT_HDR)
    para = cell.paragraphs[0]; para.clear()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = para.add_run(text); r.bold = True
    r.font.size = Pt(9); r.font.color.rgb = RGBColor(*TXT_STRUCT_HDR)
    cell.width = S_COL_W[i]

STRUCT_ROWS = [
    ('Interacciones no-lineales entre variables',
     'Imposible. Holt-Winters y Croston son modelos lineales/aditivos por diseno matematico. '
     'Ninguna configuracion cambia su naturaleza algoritmica.',
     'Fortaleza central de RandomForest — captura combinaciones como "quincena + fin de mes + lunes" como un efecto unico.'),
    ('Ingesta de datos externos via API\n(clima, Banco Central RD, remesas, turismo)',
     'Imposible. Arquitectura cerrada — UDF solo ingiere lo que fluye por pipelines SAP estandar. '
     'No existe mecanismo para consumir fuentes externas.',
     'Cualquier fuente externa integrable: API meteorologica, indicadores del Banco Central RD, datos de turismo.'),
    ('Feature importance / explicabilidad de negocio',
     'Imposible. Los modelos estadisticos no producen importancia de variables. '
     'Nunca sabras cuanto pesa la Regalia Pascual en el forecast.',
     '"La Regalia Pascual explica el 22% del error de abril en la tienda de Santiago."'),
    ('Magnitud aprendida de eventos unicos\n(ej. Regalia Pascual)',
     'Solo puede marcarla como feriado con efecto plano y fijo. '
     'Nunca aprende del historico cuanto sube exactamente cada producto en ese evento.',
     'Aprende la magnitud real por producto y por tienda directamente desde el historico.'),
    ('Aprendizaje cruzado entre SKUs',
     'Imposible. Cada serie se modela de forma completamente independiente. '
     'No detecta que cuando los licores suben en diciembre, los vasos desechables siguen el mismo patron.',
     'El modelo aprende correlaciones entre productos dentro de la misma sesion de entrenamiento.'),
]

for cap, udf_txt, cst_txt in STRUCT_ROWS:
    row = struct_table.add_row()
    cells = row.cells
    for i, w in enumerate(S_COL_W):
        cells[i].width = w

    # Capacidad
    set_cell_bg(cells[0], 'F2F2F2')
    para = cells[0].paragraphs[0]; para.clear()
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = para.add_run(cap); r.bold = True
    r.font.size = Pt(9); r.font.color.rgb = RGBColor(*TXT_STRUCT_CAP)

    # UDF
    set_cell_bg(cells[1], BG_STRUCT_UDF)
    para = cells[1].paragraphs[0]; para.clear()
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = para.add_run('✗  '); r.bold = True
    r.font.size = Pt(10); r.font.color.rgb = RGBColor(*TXT_STRUCT_UDF)
    r2 = para.add_run(udf_txt); r2.bold = False
    r2.font.size = Pt(8); r2.font.color.rgb = RGBColor(*TXT_STRUCT_UDF)

    # Custom
    set_cell_bg(cells[2], BG_STRUCT_CST)
    para = cells[2].paragraphs[0]; para.clear()
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = para.add_run('✓  '); r.bold = True
    r.font.size = Pt(10); r.font.color.rgb = RGBColor(*TXT_STRUCT_CST)
    r2 = para.add_run(cst_txt); r2.bold = False
    r2.font.size = Pt(8); r2.font.color.rgb = RGBColor(*TXT_STRUCT_CST)

# ─── 8. Diferencias de Diseno ─────────────────────────────────────────────────
doc.add_paragraph()
doc.add_heading('8. Diferencias de Diseno — Enfoques Distintos', level=1)
doc.add_paragraph(
    'Las siguientes caracteristicas no representan ventajas ni desventajas de ninguna solucion — '
    'son decisiones de arquitectura y enfoque. Se presentan separadas para evitar confusiones '
    'en la interpretacion de la tabla comparativa anterior.'
)
doc.add_paragraph()

design_table = doc.add_table(rows=0, cols=3)
design_table.style = 'Table Grid'
BG_DESIGN_HDR  = '2E4057'
BG_DESIGN_ROW  = 'F5F5F5'
BG_DESIGN_ALT  = 'EAEAEA'
TXT_DESIGN_HDR = (0xFF, 0xFF, 0xFF)
TXT_DESIGN     = (0x20, 0x20, 0x20)

D_COL_W = [Cm(5.0), Cm(5.5), Cm(6.0)]

# Cabecera
dh = design_table.add_row()
for i, (cell, text) in enumerate(zip(dh.cells,
        ['Dimension', 'SAP CAR UDF', 'Modelo Custom (SAP AI Core)'])):
    set_cell_bg(cell, BG_DESIGN_HDR)
    para = cell.paragraphs[0]; para.clear()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = para.add_run(text); r.bold = True
    r.font.size = Pt(10); r.font.color.rgb = RGBColor(*TXT_DESIGN_HDR)
    cell.width = D_COL_W[i]

DESIGN_ROWS = [
    ('Fuente principal de datos',
     'Ventas del POS (caja registradora)',
     'Consumo historico e inventario proyectado'),
    ('Pregunta de negocio que responde',
     '"Cuanto venderemos el proximo periodo?"',
     '"Cuantos dias de stock nos quedan antes de desabastecernos?"'),
    ('Arquitectura para reabastecimiento proactivo',
     'UDF conectado con MRP e IBP (cadena SAP completa)',
     'Modelo integrado directamente con CAP y sistema de compras'),
    ('Tipo de alerta adicional disponible',
     'Desviacion de ventas vs. forecast (post-evento)',
     'Proyeccion de quiebre de stock por fecha futura'),
    ('Soporte y mantenimiento',
     'SAP Standard incluido en licencia CAR',
     'Responsabilidad del equipo de desarrollo del cliente'),
    ('Modelo de licenciamiento',
     'Incluido en licencia SAP CAR',
     'SAP AI Core (licencia independiente)'),
]

for idx, (dim, udf_txt, cst_txt) in enumerate(DESIGN_ROWS):
    row = design_table.add_row()
    bg = BG_DESIGN_ROW if idx % 2 == 0 else BG_DESIGN_ALT
    for i, (cell, text, w) in enumerate(zip(
            row.cells, [dim, udf_txt, cst_txt], D_COL_W)):
        set_cell_bg(cell, bg)
        cell.width = w
        para = cell.paragraphs[0]; para.clear()
        para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r = para.add_run(text)
        r.bold = (i == 0)
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(*TXT_DESIGN)

# ─── 8. Conclusiones ──────────────────────────────────────────────────────────
doc.add_paragraph()
doc.add_heading('9. Conclusiones y Recomendacion', level=1)

p = doc.add_paragraph(); p.add_run('UDF es la solucion adecuada si:').bold = True
for item in [
    'Grupo Ramos ya cuenta con licencia SAP CAR activa.',
    'El foco es la integracion estandar con MRP/IBP sin desarrollo.',
    'No se requiere incorporar senales de negocio propias del mercado dominicano.',
]:
    doc.add_paragraph(item, style='List Bullet')

doc.add_paragraph()
p = doc.add_paragraph(); p.add_run('El Modelo Custom es la solucion superior si:').bold = True
for item in [
    'Se quiere incorporar Regalia Pascual, quincenas, remesas, alertas de huracan o cualquier senal propia.',
    'Se requiere diferenciacion de prediccion por tienda o region (Santo Domingo, Santiago, Punta Cana, etc.).',
    'Se quiere entender QUE variables impulsan el consumo (feature importance = inteligencia de negocio).',
    'Grupo Ramos no tiene licencia SAP CAR (evita esa inversion).',
    'Se proyecta crecimiento del modelo con nuevas fuentes de datos externas.',
    'Se necesita reentrenar el modelo frecuentemente a medida que el negocio evoluciona.',
]:
    doc.add_paragraph(item, style='List Bullet')

doc.add_paragraph()
p = doc.add_paragraph()
r = p.add_run('Recomendacion final: ')
r.bold = True
p.add_run(
    'Dado el perfil de Grupo Ramos como operador de gran superficie en Republica Dominicana, '
    'con eventos comerciales y ciclos de pago unicos del mercado dominicano — especialmente la '
    'Regalia Pascual, las quincenas y el efecto de las remesas en Navidad y Dia de las Madres — '
    'el Modelo Custom en SAP AI Core representa la opcion con mayor retorno de inversion en '
    'precision de prediccion. UDF estandar no fue disenado para capturar la riqueza de senales '
    'de negocio disponibles en este contexto especifico.'
)

# ─── Guardar ──────────────────────────────────────────────────────────────────
output = '/home/user/projects/material-consumption-ai/Comparativa_UDF_vs_ModeloCustom_GrupoRamos.docx'
doc.save(output)
print(f'Documento guardado: {output}')
