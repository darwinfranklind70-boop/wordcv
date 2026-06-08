#!/usr/bin/env python3
"""Genera un CV honesto y completable (.docx) para Técnico Electrónico junior sin experiencia."""

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

AZUL = RGBColor(0x1F, 0x4E, 0x79)
GRIS = RGBColor(0x59, 0x59, 0x59)
NEGRO = RGBColor(0x00, 0x00, 0x00)


def set_cell_bg(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)


def add_heading(doc, text):
    p = doc.add_paragraph()
    p.space_before = Pt(10)
    run = p.add_run(text.upper())
    run.bold = True
    run.font.size = Pt(12)
    run.font.color.rgb = AZUL
    run.font.name = 'Calibri'
    # línea inferior
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '2')
    bottom.set(qn('w:color'), '1F4E79')
    pbdr.append(bottom)
    pPr.append(pbdr)
    return p


def add_bullet(doc, text, placeholder=False):
    p = doc.add_paragraph(style='List Bullet')
    run = p.add_run(text)
    run.font.size = Pt(10.5)
    run.font.name = 'Calibri'
    if placeholder:
        run.font.color.rgb = GRIS
        run.italic = True
    return p


def add_field(doc, label, hint):
    p = doc.add_paragraph()
    r1 = p.add_run(f"{label}: ")
    r1.bold = True
    r1.font.size = Pt(10.5)
    r1.font.name = 'Calibri'
    r2 = p.add_run(hint)
    r2.font.size = Pt(10.5)
    r2.font.name = 'Calibri'
    r2.font.color.rgb = GRIS
    r2.italic = True
    return p


doc = Document()

# Márgenes
for section in doc.sections:
    section.top_margin = Inches(0.6)
    section.bottom_margin = Inches(0.6)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)

# Estilo base
style = doc.styles['Normal']
style.font.name = 'Calibri'
style.font.size = Pt(10.5)

# ---------------- ENCABEZADO ----------------
nombre = doc.add_paragraph()
nombre.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = nombre.add_run("[NOMBRE Y APELLIDO]")
r.bold = True
r.font.size = Pt(22)
r.font.color.rgb = AZUL

subt = doc.add_paragraph()
subt.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = subt.add_run("Técnico Electrónico  |  Perfil Junior")
r.font.size = Pt(12)
r.font.color.rgb = GRIS

contacto = doc.add_paragraph()
contacto.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = contacto.add_run("[Teléfono]  •  [Email]  •  [Ciudad/Barrio, CABA]  •  [LinkedIn opcional]")
r.font.size = Pt(10)
r.font.color.rgb = GRIS

# ---------------- PERFIL ----------------
add_heading(doc, "Perfil profesional")
p = doc.add_paragraph()
r = p.add_run(
    "Técnico electrónico con formación en la E.T. N.º 1 \u201cOtto Krause\u201d, con sólida base "
    "en electrónica, redes y reparación de equipos. Persona responsable, proactiva y con "
    "fuerte vocación de aprendizaje continuo. Busco mi primera experiencia formal como "
    "Técnico de Campo en videovigilancia, control de acceso e informática, donde poder "
    "aplicar y seguir desarrollando mis conocimientos. "
    "[Ajustá este texto a tu estilo.]"
)
r.font.size = Pt(10.5)
r.font.color.rgb = NEGRO

# ---------------- FORMACIÓN ----------------
add_heading(doc, "Formación académica")
add_field(doc, "Título", "Técnico Electrónico \u2014 E.T. N.º 1 \u201cOtto Krause\u201d (indicá AÑO de egreso / o \u201cTítulo en trámite\u201d)")
add_field(doc, "Estado", "Egresado \u2014 plan de estudios completo, todas las materias aprobadas. Título en trámite.")

# ---------------- CURSOS (roles/skills que se adquieren rápido) ----------------
add_heading(doc, "Cursos y capacitaciones (sugeridos para sumar rápido)")
p = doc.add_paragraph()
r = p.add_run("Cursos cortos y gratuitos/accesibles que te hacen más competitivo en pocas semanas. "
              "Hacelos y agregalos acá con la fecha real:")
r.font.size = Pt(10)
r.italic = True
r.font.color.rgb = GRIS
add_bullet(doc, "Redes / Cableado estructurado \u2014 Cisco \u201cNetworking Basics\u201d (Networking Academy, online y gratis)", placeholder=True)
add_bullet(doc, "Introducción a la Ciberseguridad \u2014 Cisco / Fundación Telefónica", placeholder=True)
add_bullet(doc, "Reparación y armado de PC \u2014 cursos en YouTube/EducaciónIT/CoderHouse", placeholder=True)
add_bullet(doc, "CCTV / Videovigilancia IP \u2014 capacitaciones de fabricantes (Hikvision, Dahua \u2014 academias online gratuitas)", placeholder=True)
add_bullet(doc, "Control de acceso y sistemas de alarmas \u2014 cursos básicos del rubro", placeholder=True)
add_bullet(doc, "Excel / Ofimática básica \u2014 Fundación Telefónica (gratis)", placeholder=True)

# ---------------- CONOCIMIENTOS TÉCNICOS ----------------
add_heading(doc, "Conocimientos técnicos")
add_bullet(doc, "Electrónica analógica y digital, uso de instrumental (multímetro, osciloscopio).")
add_bullet(doc, "Soldadura, armado y diagnóstico de circuitos.")
add_bullet(doc, "Redes básicas (TCP/IP, cableado, routers/switches).")
add_bullet(doc, "Instalación y mantenimiento de PC, software y periféricos.")
add_bullet(doc, "[Agregá lo que realmente manejes: Arduino, fuentes, etc.]", placeholder=True)

# ---------------- EXPERIENCIA (HONESTA) ----------------
add_heading(doc, "Experiencia y proyectos")
p = doc.add_paragraph()
r = p.add_run("Si nunca trabajaste en relación de dependencia, NO inventes empresas. "
              "Listá experiencia real aunque sea informal: changas, ayuda a familiares/vecinos, "
              "proyectos de la escuela. Esto cuenta y es honesto:")
r.font.size = Pt(10)
r.italic = True
r.font.color.rgb = GRIS
add_bullet(doc, "Reparación / armado de PCs para conocidos y vecinos (changas). [Detallá qué hacías]", placeholder=True)
add_bullet(doc, "Proyecto final / trabajos prácticos del Otto Krause: [describí un proyecto de electrónica]", placeholder=True)
add_bullet(doc, "Ayudante en [negocio familiar / tareas técnicas]: [qué tareas realizabas]", placeholder=True)
add_bullet(doc, "Instalaciones o reparaciones hechas por tu cuenta (cámaras, redes hogareñas, etc.)", placeholder=True)

# ---------------- HABILIDADES BLANDAS ----------------
add_heading(doc, "Habilidades personales")
add_bullet(doc, "Gran disposición y entusiasmo por aprender cosas nuevas.")
add_bullet(doc, "Responsabilidad y compromiso.")
add_bullet(doc, "Trabajo en equipo y buen trato con clientes.")
add_bullet(doc, "Resolución de problemas y atención al detalle.")
add_bullet(doc, "Disponibilidad para tareas de campo y movilidad en CABA.")

# ---------------- INFORMÁTICA / IDIOMAS / EXTRA ----------------
add_heading(doc, "Otros datos")
add_field(doc, "Idiomas", "Español (nativo) \u2014 Inglés [nivel: básico/intermedio]")
add_field(doc, "Informática", "Windows, paquete Office, [otros]")
add_field(doc, "Movilidad", "[Licencia de conducir / SUBE / disponibilidad para viajar a clientes en CABA]")
add_field(doc, "Disponibilidad", "[Horaria \u2014 el puesto es de 8:30 a 14:30 h]")
add_field(doc, "Edad", "24 a\u00f1os")

# ---------------- NOTA AL PIE (se borra antes de enviar) ----------------
doc.add_paragraph()
nota = doc.add_paragraph()
r = nota.add_run(
    "\u2014\u2014\u2014 NOTA (BORRAR ANTES DE ENVIAR): Reemplazá todo lo que está entre [corchetes] "
    "y en gris/itálica con tus datos reales. No agregues experiencia falsa. Una vez completo, "
    "guardalo como PDF para enviarlo."
)
r.font.size = Pt(9)
r.italic = True
r.font.color.rgb = RGBColor(0xB0, 0x00, 0x00)

doc.save("/projects/sandbox/CV_Tecnico_Electronico.docx")
print("CV generado: CV_Tecnico_Electronico.docx")
