from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import io
import requests

app = Flask(__name__)
CORS(app, origins='*')

@app.after_request
def after_request(response):
    response.headers.add('Access-Control-Allow-Origin', '*')
    response.headers.add('Access-Control-Allow-Headers', 'Content-Type')
    response.headers.add('Access-Control-Allow-Methods', 'GET,POST,OPTIONS')
    return response

# ── COLORES ───────────────────────────────────────────────────
OR  = RGBColor(0xFF, 0x5B, 0x00)
DK  = RGBColor(0x21, 0x27, 0x2A)
WH  = RGBColor(0xFF, 0xFF, 0xFF)
GR  = RGBColor(0xF0, 0xF0, 0xF0)
BK  = RGBColor(0x00, 0x00, 0x00)
GY2 = RGBColor(0xEA, 0xE8, 0xE9)
DK2 = RGBColor(0x2A, 0x2A, 0x2A)
LGY = RGBColor(0x55, 0x55, 0x55)
MGY = RGBColor(0xBB, 0xBB, 0xBB)
OR2 = RGBColor(0xFB, 0x79, 0x01)

LOGO_W = 'https://res.cloudinary.com/dpcojkrta/image/upload/v1780945082/Logo_Kruger_naranja_blanco_negro_y_plomo-01_mmdzkk.png'
LOGO_B = 'https://res.cloudinary.com/dpcojkrta/image/upload/v1780944928/Logo_Kruger_naranja_blanco_negro_y_plomo-02_ynrqui.png'

W = Inches(13.33)
H = Inches(7.5)

def i(v): return Inches(v)
def p(v): return Pt(v)
def hc(h): 
    h = h.lstrip('#')
    return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

def fetch_image(url):
    r = requests.get(url, timeout=10)
    return io.BytesIO(r.content)

def set_bg(sl, rgb):
    bg = sl.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = rgb

def add_rect(sl, x, y, w, h, fill_rgb, line_rgb=None):
    shape = sl.shapes.add_shape(1, i(x), i(y), i(w), i(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_rgb
    if line_rgb:
        shape.line.color.rgb = line_rgb
        shape.line.width = Pt(0.5)
    else:
        shape.line.fill.background()
    return shape

def add_text(sl, x, y, w, h, text, size, bold=False, color=None,
             align=PP_ALIGN.LEFT, italic=False, wrap=True, spacing=0):
    txb = sl.shapes.add_textbox(i(x), i(y), i(w), i(h))
    tf = txb.text_frame
    tf.word_wrap = wrap
    para = tf.paragraphs[0]
    para.alignment = align
    run = para.add_run()
    run.text = text
    run.font.size = p(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = 'Inter'
    if color:
        run.font.color.rgb = color
    return txb

def add_text_runs(sl, x, y, w, h, runs, align=PP_ALIGN.LEFT, wrap=True):
    """runs = [(text, size, bold, color), ...]"""
    txb = sl.shapes.add_textbox(i(x), i(y), i(w), i(h))
    tf = txb.text_frame
    tf.word_wrap = wrap
    para = tf.paragraphs[0]
    para.alignment = align
    for txt, sz, bold, clr in runs:
        run = para.add_run()
        run.text = txt
        run.font.size = p(sz)
        run.font.bold = bold
        run.font.name = 'Inter'
        if clr:
            run.font.color.rgb = clr
    return txb

def add_logo(sl, url, x, y, w, h):
    try:
        img_io = fetch_image(url)
        sl.shapes.add_picture(img_io, i(x), i(y), i(w), i(h))
    except Exception as e:
        print(f"Logo error: {e}")

def add_placeholder(sl, x, y, w, h, label):
    shape = sl.shapes.add_shape(1, i(x), i(y), i(w), i(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = hc('F5F5F5')
    shape.line.color.rgb = OR
    shape.line.width = Pt(1)
    add_text(sl, x, y, w, h, label, 9, bold=True,
             color=hc('FF5B0088'), align=PP_ALIGN.CENTER)

# ── SLIDE 1: PORTADA ─────────────────────────────────────────
def slide_portada(prs, sl, cfg):
    set_bg(sl, OR)
    # Placeholder imagen derecha (55% del ancho)
    add_placeholder(sl, 5.98, 0, 7.35, 7.5, 'Imagen principal')
    # Logo blanco arriba izquierda
    add_logo(sl, LOGO_W, 0.4, 0.25, 2.5, 0.62)
    # Nombre y cargo arriba derecha
    add_text(sl, 8.5, 0.22, 4.6, 0.4, cfg.get('name','Nombre Apellido'),
             13, bold=True, color=WH, align=PP_ALIGN.RIGHT)
    add_text(sl, 8.5, 0.6, 4.6, 0.35, cfg.get('title','Cargo'),
             11, color=hc('FFD0B0'), align=PP_ALIGN.RIGHT)
    # Título grande (escalado de 79pt en 20" → 53pt en 13.33")
    add_text(sl, 0.4, 2.8, 5.4, 2.8,
             'WEB & MOBILE APP\nDEVELOPMENT\nCASE STUDIES',
             53, bold=True, color=WH)
    # Subtítulo
    add_text_runs(sl, 0.4, 5.9, 5.4, 0.9, [
        ('Diseñando experiencias digitales escalables para ', 17, False, WH),
        ('enterprise, telecom, fintech y ecosistemas de cliente.', 17, True, WH)
    ], wrap=True)

# ── SLIDE 2: DISCLAIMER ──────────────────────────────────────
def slide_disclaimer(prs, sl):
    set_bg(sl, GY2)
    # Decoración derecha abstracta
    add_rect(sl, 8.0, 1.0, 4.8, 5.5, hc('C8C4C6'))
    add_rect(sl, 9.5, 0.5, 3.5, 6.5, hc('D4D0D2'))
    # Logo negro
    add_logo(sl, LOGO_B, 0.75, 0.5, 2.2, 0.55)
    # Título Disclaimer (69.81pt → 47pt)
    add_text(sl, 0.75, 1.25, 6.5, 1.0, 'Disclaimer', 47, bold=True, color=OR)
    # Línea
    add_rect(sl, 0.75, 2.3, 0.85, 0.05, DK)
    # Texto con bold parcial (25.85pt → 17pt)
    add_text_runs(sl, 0.75, 2.5, 6.2, 3.8, [
        ('La información contenida en este documento es ', 17, False, DK),
        ('confidencial, privilegiada y solo para uso del destinatario previsto', 17, True, DK),
        (', y no puede ser usada, publicada ni redistribuida ', 17, False, DK),
        ('sin el consentimiento previo por escrito ', 17, True, DK),
        ('de Kruger Corporation.', 17, True, DK),
    ], wrap=True)

# ── SLIDE 3: QUIÉNES SOMOS ───────────────────────────────────
def slide_quienes(prs, sl):
    set_bg(sl, GR)
    add_rect(sl, 12.33, 0, 1.0, 1.0, OR)
    # Logo negro
    add_logo(sl, LOGO_B, 0.6, 0.3, 2.2, 0.55)
    # Línea naranja
    add_rect(sl, 0.6, 1.0, 1.5, 0.05, OR)
    # Tagline (20.95pt → 14pt)
    add_text_runs(sl, 0.6, 1.15, 6.5, 0.4, [
        ('BUILDING SOLUTIONS. ', 14, True, DK),
        ('DELIVERING SUCCESS.', 14, True, OR),
    ])
    # Descripción (29.17pt → 19pt)
    add_text_runs(sl, 0.6, 1.65, 5.8, 1.4, [
        ('Kruger crea aplicaciones web y móviles que ayudan a las empresas a convertir ', 19, False, DK),
        ('productos digitales complejos', 19, True, DK),
        (' en experiencias simples y centradas en el usuario.', 19, False, DK),
    ], wrap=True)
    # Stats
    stats = [('30+','AÑOS DE INNOVACIÓN'),('1500+','PROYECTOS ENTREGADOS'),
             ('450+','EXPERTOS CERTIFICADOS'),('14','PAÍSES DE PRESENCIA GLOBAL')]
    for i_s,(v,l) in enumerate(stats):
        y = 3.2 + i_s * 0.95
        add_rect(sl, 0.6, y, 0.38, 0.38, WH, OR)
        add_text(sl, 1.15, y, 5.5, 0.38,
                 f'{v}  {l}', 14, bold=True, color=OR)
    # Placeholder imagen derecha
    add_placeholder(sl, 7.1, 0.3, 5.9, 6.9, 'Foto persona · traje naranja')

# ── SLIDE 4: WHY KRUGER ──────────────────────────────────────
def slide_porque(prs, sl):
    set_bg(sl, OR)
    # Número 01 watermark (339.97pt → 227pt)
    add_text(sl, 7.2, -0.3, 5.5, 4.5, '01', 227, bold=True, color=hc('FFFFFF'))
    # WHY (99.99pt → 67pt)
    add_text(sl, 6.9, 4.3, 6.0, 1.3, 'WHY', 67, color=WH)
    # KRUGER (126.92pt → 85pt)
    add_text(sl, 6.9, 5.3, 6.3, 1.6, 'KRUGER', 85, bold=True, color=WH)
    # Bullets
    puntos = [
        'Transformamos negocios con tecnología centrada en personas',
        '14 países · 30 años · 1500+ proyectos entregados',
        'Expertos en sectores regulados y ecosistemas enterprise'
    ]
    for idx,(pt) in enumerate(puntos):
        add_text_runs(sl, 7.2, 3.1 + idx * 0.72, 5.9, 0.65, [
            (f'{idx+1}.  ', 13, True, hc('FFE0CC')),
            (pt, 13, False, hc('FFE0CC')),
        ], wrap=True)
    # Logo blanco abajo derecha
    add_logo(sl, LOGO_W, 11.5, 6.85, 1.6, 0.45)
    # Placeholder imagen izquierda
    add_placeholder(sl, 0.3, 0.3, 6.3, 6.9, 'Persona editorial · traje naranja')

# ── SLIDE 5: WHAT WE BUILD ───────────────────────────────────
def slide_como(prs, sl):
    set_bg(sl, GR)
    add_rect(sl, 12.33, 0, 1.0, 1.0, OR)
    # Título (95.06pt → 63pt)
    add_text_runs(sl, 0.6, 0.2, 7.5, 1.3, [
        ('What ', 63, True, DK),
        ('We Build', 63, True, OR),
    ])
    # Subtítulo (25.79pt → 17pt)
    add_text_runs(sl, 0.6, 1.6, 7.0, 0.55, [
        ('We build ', 17, False, LGY),
        ('scalable digital experiences', 17, True, DK),
        (' that connect ', 17, False, LGY),
        ('technology, operations and customer engagement.', 17, True, DK),
    ], wrap=True)
    # Línea vertical naranja
    add_rect(sl, 0.6, 2.3, 0.08, 4.8, OR)
    # Items (20.59pt → 14pt)
    items = [
        ('Mobile & Self-Service Apps', 'Journeys digitales para clientes.'),
        ('Payment & Transactional UX', 'Flujos seguros para conversión.'),
        ('Enterprise Portals', 'Procesos complejos simplificados.'),
        ('B2C Digital Experiences', 'Centrados en usabilidad.'),
        ('B2B Platforms', 'Herramientas operativas.'),
        ('UX/UI Systems', 'Interfaces multi-pantalla.'),
    ]
    for idx,(t,d) in enumerate(items):
        y = 2.35 + idx * 0.75
        add_text_runs(sl, 0.85, y, 6.5, 0.65, [
            (f'• {t}', 14, True, DK),
            (f'  —  {d}', 14, False, LGY),
        ], wrap=True)
    # Footer HUMAN-CENTERED (50.28pt → 34pt / 36.23pt → 24pt)
    add_text_runs(sl, 0.6, 7.1, 6.5, 0.3, [
        ('HUMAN-CENTERED ', 14, True, OR),
        ('digital products powered by INFINITE DIGITAL TECHNOLOGY.', 14, False, DK),
    ])
    # Placeholder
    add_placeholder(sl, 7.1, 0.3, 5.9, 6.9, 'Mockups multi-dispositivo')

# ── SLIDE 6: DOLORES ─────────────────────────────────────────
def slide_dolores(prs, sl, vertical_nm, pain):
    set_bg(sl, DK)
    add_text_runs(sl, 0.4, 0.2, 12.0, 0.9, [
        ('Dolores en ', 32, True, OR),
        (vertical_nm, 32, True, WH),
    ])
    add_rect(sl, 0.4, 1.15, 12.5, 0.05, hc('FF5B0066'))
    add_placeholder(sl, 0.4, 1.3, 2.6, 5.9, 'Ejecutivo analítico')
    for idx, pt in enumerate(pain):
        col = 0 if idx < 3 else 1
        row = idx % 3
        x = 3.25 + col * 4.8
        y = 1.35 + row * 2.0
        add_rect(sl, x, y, 4.6, 1.85, hc('2E2E2E'), hc('FF5B0055'))
        add_text(sl, x+0.2, y+0.12, 4.2, 0.45,
                 ' '.join(pt.split()[:4]), 13, bold=True, color=OR, wrap=True)
        add_text(sl, x+0.2, y+0.6, 4.2, 1.1, pt, 11, color=MGY, wrap=True)

# ── CASE INTRO ───────────────────────────────────────────────
def slide_case_intro(prs, sl, case, dark):
    bg = DK2 if dark else GR
    tc = WH if dark else DK
    logo_url = LOGO_W if dark else LOGO_B
    set_bg(sl, bg)

    words = case['n'].split()
    mid = (len(words)+1)//2
    line1 = ' '.join(words[:mid])
    line2 = ' '.join(words[mid:]) if len(words) > 1 else ''

    # Nombre cliente (99.99pt → 67pt)
    add_text(sl, 0.6, 1.0, 6.5, 1.4, line1, 67, bold=True, color=tc)
    if line2:
        add_text(sl, 0.6, 2.3, 6.5, 1.4, line2, 67, bold=True, color=OR)
    # Subtítulo (50.01pt → 33pt)
    add_text(sl, 0.6, 3.85, 6.3, 1.0, case['sub'], 33, bold=True,
             color=hc('EEEEEE') if dark else DK, wrap=True)
    # Línea
    add_rect(sl, 0.6, 4.95, 1.2, 0.07, OR)
    # Meta (19.49pt → 13pt)
    add_text(sl, 0.6, 5.15, 6.0, 0.35,
             'CASO DE ÉXITO · KRUGER WORLDWIDE', 13, bold=True,
             color=hc('999999') if dark else hc('777777'))
    # Logo
    add_logo(sl, logo_url, 0.6, 5.65, 2.0, 0.5)
    add_placeholder(sl, 7.1, 0.3, 5.9, 6.9, 'Foto profesional · traje naranja')

# ── CASE DETAIL ──────────────────────────────────────────────
def slide_case_detail(prs, sl, case):
    set_bg(sl, BK)
    # Header gris (nombre cliente 85.98pt → 57pt)
    add_rect(sl, 0, 0, 13.33, 1.25, GR)
    add_text(sl, 0.4, 0.1, 10.5, 1.0, case['n'], 57, bold=True, color=DK)
    add_logo(sl, LOGO_B, 11.5, 0.3, 1.6, 0.55)
    # Sub-header (35.6pt → 24pt)
    add_rect(sl, 0, 1.25, 13.33, 0.6, hc('111111'))
    add_text(sl, 0.4, 1.32, 12.0, 0.45, case['sub'].upper(), 16, bold=True, color=WH)
    # Placeholder imagen izquierda
    add_placeholder(sl, 0.3, 1.95, 4.8, 5.3, 'Screenshots · Persona')
    # Cards CHALLENGE/APPROACH/OUTCOME
    cards = [('CHALLENGE', case['ch']), ('APPROACH', case['ap']), ('OUTCOME', case['out'])]
    for idx,(lbl,txt) in enumerate(cards):
        y = 2.0 + idx * 1.82
        add_rect(sl, 5.4, y, 7.6, 1.65, hc('1A1A1A'), hc('333333'))
        # Label (32.19pt → 21pt)
        add_text(sl, 5.65, y+0.15, 4.0, 0.45, lbl, 21, bold=True, color=OR2)
        # Body (19.03pt → 13pt)
        add_text(sl, 5.65, y+0.65, 7.0, 0.9, txt, 13, color=MGY, wrap=True)

# ── PRODUCTOS ────────────────────────────────────────────────
def slide_productos(prs, sl, vertical_nm, prods):
    set_bg(sl, GR)
    add_text_runs(sl, 0.4, 0.1, 11.5, 0.75, [
        ('Productos para ', 26, True, DK),
        (vertical_nm, 26, True, OR),
    ])
    add_logo(sl, LOGO_B, 11.5, 0.15, 1.6, 0.48)
    for idx, prod in enumerate(prods[:4]):
        x = 0.4 if idx < 2 else 6.85
        y = 1.0 + (idx % 2) * 2.95
        add_rect(sl, x, y, 6.2, 2.75, WH, hc('DDDDDD'))
        add_rect(sl, x, y, 6.2, 0.14, OR)
        add_text(sl, x+0.25, y+0.28, 4.5, 0.5, prod['n'], 18, bold=True, color=DK)
        # Badge categoría
        badge = sl.shapes.add_shape(1, i(x+4.85), i(y+0.28), i(1.1), i(0.38))
        badge.fill.solid(); badge.fill.fore_color.rgb = OR
        badge.line.fill.background()
        add_text(sl, x+4.85, y+0.29, 1.1, 0.36, prod['cat'], 10,
                 bold=True, color=WH, align=PP_ALIGN.CENTER)
        add_text(sl, x+0.25, y+0.85, 5.8, 1.75, prod['d'], 13, color=LGY, wrap=True)

# ── SERVICIOS ────────────────────────────────────────────────
def slide_servicios(prs, sl, vertical_nm, servs):
    set_bg(sl, DK)
    add_text_runs(sl, 0.4, 0.2, 11.5, 0.85, [
        ('Servicios para ', 28, True, WH),
        (vertical_nm, 28, True, OR),
    ])
    add_logo(sl, LOGO_W, 11.5, 0.28, 1.6, 0.48)
    add_rect(sl, 0.4, 1.15, 12.5, 0.05, hc('FF5B0066'))
    for idx, s in enumerate(servs):
        x = 0.5 + idx * 3.2
        # Círculo
        circ = sl.shapes.add_shape(9, i(x+1.15), i(1.75), i(0.65), i(0.65))
        circ.fill.solid(); circ.fill.fore_color.rgb = OR
        circ.line.fill.background()
        add_text(sl, x+1.15, 1.78, 0.65, 0.6, s['n'], 14,
                 bold=True, color=WH, align=PP_ALIGN.CENTER)
        # Card
        add_rect(sl, x, 2.55, 3.1, 4.65, hc('2E2E2E'), hc('FF5B0044'))
        add_text(sl, x+0.15, 2.78, 2.8, 0.55, s['t'], 16,
                 bold=True, color=OR, align=PP_ALIGN.CENTER)
        add_text(sl, x+0.15, 3.45, 2.8, 3.5, s['d'], 13,
                 color=MGY, wrap=True, align=PP_ALIGN.CENTER)

# ── PARTNERS ─────────────────────────────────────────────────
def slide_partners(prs, sl):
    set_bg(sl, WH)
    add_logo(sl, LOGO_B, 5.4, 0.2, 2.5, 0.65)
    add_text(sl, 0, 0.85, 13.33, 1.0, 'PARTNERS', 42,
             bold=True, color=OR, align=PP_ALIGN.CENTER)
    clientes = ['Amdocs','Verizon','Movistar','Claro','Coca-Cola',
                'Banco del Pacífico','IESS','SRI','Holcim','Pronaca',
                'CNT','Zurich','Banco Nación','Tottus','BCP',
                'Santander','El Ordeño','Rimac','Difare','CNEL']
    for idx, l in enumerate(clientes):
        col = idx % 5
        row = idx // 5
        x = 0.4 + col * 2.5
        y = 1.98 + row * 1.15
        add_rect(sl, x, y, 2.35, 0.95, GR, hc('DDDDDD'))
        add_text(sl, x, y, 2.35, 0.95, l, 11,
                 bold=True, color=LGY, align=PP_ALIGN.CENTER)
    add_text(sl, 0, 7.2, 13.33, 0.25,
             '14 países · 1500+ proyectos entregados',
             10, color=hc('AAAAAA'), align=PP_ALIGN.CENTER)

# ── GANCHO ───────────────────────────────────────────────────
def slide_gancho(prs, sl):
    set_bg(sl, WH)
    add_rect(sl, 0, 0, 0.9, 7.5, OR)
    add_text(sl, 1.1, 0.1, 1.0, 1.5, '"', 80,
             color=hc('FF5B0033'), italic=True)
    add_text(sl, 1.2, 1.5, 7.2, 3.0,
             '"El futuro de los negocios es digital. La diferencia entre liderar y quedarse atrás está en la experiencia que ofreces a tus clientes."',
             17, bold=True, italic=True, color=DK, wrap=True)
    add_rect(sl, 1.2, 4.65, 1.5, 0.08, OR)
    add_text(sl, 1.2, 4.85, 7.5, 0.9,
             'Kruger ha transformado más de 1500 productos digitales en 14 países.',
             13, color=LGY, wrap=True)
    add_placeholder(sl, 9.1, 0.3, 3.9, 6.9, 'Imagen gancho')

# ── CIERRE ───────────────────────────────────────────────────
def slide_cierre(prs, sl, cfg):
    set_bg(sl, OR)
    # Nombre y cargo (18.99pt → 13pt)
    add_text(sl, 0.4, 0.28, 7.0, 0.42,
             cfg.get('name','Nombre Apellido'), 13, bold=True, color=WH)
    add_text(sl, 0.4, 0.7, 7.0, 0.35,
             cfg.get('title',''), 11, color=hc('FFD0B0'))
    add_text(sl, 7.0, 0.28, 5.9, 0.42,
             'KRUGER CORPORATION', 13, bold=True, color=WH, align=PP_ALIGN.RIGHT)
    add_text(sl, 7.0, 0.7, 5.9, 0.35,
             'krugerworldwide.com', 11, color=hc('FFD0B0'), align=PP_ALIGN.RIGHT)
    # Logo blanco grande centrado
    add_logo(sl, LOGO_W, 4.8, 1.5, 3.7, 0.92)
    # Línea separadora
    add_rect(sl, 3.0, 2.65, 7.3, 0.06, hc('FFFFFF88'))
    # Placeholder camaleón
    add_placeholder(sl, 5.65, 2.85, 2.0, 2.0, 'Camaleón 3D')
    # Contacto (21.99pt → 15pt)
    email = cfg.get('email','worldwide@krugercorp.com')
    phone = cfg.get('phone','')
    contact = f"{email}  ·  {phone}" if phone else email
    add_text(sl, 0, 5.1, 13.33, 0.45, contact,
             15, bold=True, color=WH, align=PP_ALIGN.CENTER)
    if cfg.get('client'):
        add_text(sl, 0, 5.65, 13.33, 0.38,
                 f"Preparado especialmente para {cfg['client']}",
                 12, color=hc('FFE8D6'), align=PP_ALIGN.CENTER, italic=True)
    add_text(sl, 0.4, 7.2, 5.0, 0.22,
             '© 2026 KRUGER CORP.', 8, color=WH)
    add_text(sl, 8.0, 7.2, 5.0, 0.22,
             'Confidencial', 8, color=WH, align=PP_ALIGN.RIGHT)

# ── ENDPOINT ─────────────────────────────────────────────────
@app.route('/generate', methods=['POST', 'OPTIONS'])
def generate():
    if request.method == 'OPTIONS':
        return '', 200
    data = request.json
    vertical = data.get('vertical', {})
    cfg = data.get('cfg', {})

    prs = Presentation()
    prs.slide_width  = W
    prs.slide_height = H
    blank = prs.slide_layouts[6]
    def ns(): return prs.slides.add_slide(blank)

    slide_portada(prs, ns(), cfg)
    slide_disclaimer(prs, ns())
    slide_quienes(prs, ns())
    slide_porque(prs, ns())
    slide_como(prs, ns())
    slide_dolores(prs, ns(), vertical.get('nm',''), vertical.get('pain',[]))
    for idx, case in enumerate(vertical.get('cases', [])):
        slide_case_intro(prs, ns(), case, idx % 2 == 0)
        slide_case_detail(prs, ns(), case)
    slide_productos(prs, ns(), vertical.get('nm',''), vertical.get('prods',[]))
    slide_servicios(prs, ns(), vertical.get('nm',''), vertical.get('servs',[]))
    slide_partners(prs, ns())
    slide_gancho(prs, ns())
    slide_cierre(prs, ns(), cfg)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)

    nm = vertical.get('nm','Vertical').replace(' ','_')
    cl = cfg.get('client','Cliente').replace(' ','_')
    filename = f'Kruger_{nm}_{cl}_2026.pptx'

    return send_file(buf,
        mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
        as_attachment=True,
        download_name=filename)

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=False)
