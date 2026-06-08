from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import io
import requests
from PIL import Image

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

LOGO_W = 'https://res.cloudinary.com/dpcojkrta/image/upload/v1780945082/Logo_Kruger_naranja_blanco_negro_y_plomo-01_mmdzkk.png'
LOGO_B = 'https://res.cloudinary.com/dpcojkrta/image/upload/v1780944928/Logo_Kruger_naranja_blanco_negro_y_plomo-02_ynrqui.png'

# Slide dimensions: LAYOUT_WIDE = 13.33" x 7.5"
W = Inches(13.33)
H = Inches(7.5)


def in_(v): return Inches(v)
def pt_(v): return Pt(v)


def hex_color(h):
    h = h.lstrip('#')
    return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))


def fetch_image(url):
    """Fetch image from URL and return as BytesIO"""
    r = requests.get(url, timeout=10)
    return io.BytesIO(r.content)


def add_rect(sl, prs, x, y, w, h, fill_rgb, line_rgb=None, line_width=0):
    from pptx.util import Pt
    shape = sl.shapes.add_shape(1, in_(x), in_(y), in_(w), in_(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_rgb
    if line_rgb:
        shape.line.color.rgb = line_rgb
        shape.line.width = Pt(line_width) if line_width else Pt(0)
    else:
        shape.line.fill.background()
    return shape


def add_textbox(sl, x, y, w, h, text, font_size, bold=False, color=None,
                align=PP_ALIGN.LEFT, font_name='Inter', italic=False,
                char_spacing=0, wrap=True):
    txb = sl.shapes.add_textbox(in_(x), in_(y), in_(w), in_(h))
    tf = txb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = pt_(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = font_name
    if color:
        run.font.color.rgb = color
    return txb


def add_placeholder_box(sl, x, y, w, h, label):
    """Dashed orange placeholder box"""
    shape = sl.shapes.add_shape(1, in_(x), in_(y), in_(w), in_(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = hex_color('F5F5F5')
    shape.line.color.rgb = OR
    shape.line.width = Pt(1)
    add_textbox(sl, x, y, w, h, label, 10, bold=True,
                color=hex_color('FF5B0088'), align=PP_ALIGN.CENTER)


def add_logo(sl, url, x, y, w, h):
    """Add logo image from URL"""
    try:
        img_io = fetch_image(url)
        sl.shapes.add_picture(img_io, in_(x), in_(y), in_(w), in_(h))
    except Exception:
        pass


def set_bg(sl, rgb):
    bg = sl.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = rgb


# ── SLIDES ────────────────────────────────────────────────────

def slide_portada(prs, sl, cfg):
    set_bg(sl, OR)
    # Placeholder imagen derecha
    add_placeholder_box(sl, 6.0, 0, 7.33, 7.5, 'Imagen principal')
    # Logo blanco
    add_logo(sl, LOGO_W, 0.4, 0.28, 2.2, 0.55)
    # Nombre y cargo arriba derecha
    add_textbox(sl, 8.0, 0.28, 5.0, 0.38, cfg.get('name','Nombre Apellido'),
                11, bold=True, color=WH, align=PP_ALIGN.RIGHT)
    add_textbox(sl, 8.0, 0.65, 5.0, 0.32, cfg.get('title','Cargo'),
                9, color=hex_color('FFD0B0'), align=PP_ALIGN.RIGHT)
    # Texto principal
    add_textbox(sl, 0.4, 2.5, 5.4, 1.4,
                'WEB & MOBILE APP DEVELOPMENT\nCASE STUDIES',
                32, bold=True, color=WH)
    add_textbox(sl, 0.4, 4.1, 5.4, 1.0,
                'Diseñando experiencias digitales escalables para enterprise, telecom, fintech y ecosistemas de cliente.',
                11, color=hex_color('FFE8D6'))


def slide_disclaimer(prs, sl):
    set_bg(sl, GY2)
    # Decoración derecha (rectángulos redondeados simulando freeforms)
    add_rect(sl, prs, 7.5, 0.5, 5.5, 6.5, hex_color('C8C4C6'))
    # Logo negro
    add_logo(sl, LOGO_B, 0.7, 0.5, 2.0, 0.5)
    # Título
    add_textbox(sl, 0.7, 1.2, 6.0, 0.9, 'Disclaimer',
                36, bold=True, color=OR)
    # Línea decorativa
    add_rect(sl, prs, 0.7, 2.2, 0.8, 0.05, DK)
    # Texto
    add_textbox(sl, 0.7, 2.4, 6.0, 3.5,
                'La información contenida en este documento es confidencial, privilegiada y solo para uso del destinatario previsto. No puede ser usada, publicada ni redistribuida sin el consentimiento previo por escrito de Kruger Corporation.',
                13, color=DK, wrap=True)


def slide_quienes(prs, sl):
    set_bg(sl, GR)
    # Esquina naranja
    add_rect(sl, prs, 12.33, 0, 1.0, 1.0, OR)
    # Logo negro
    add_logo(sl, LOGO_B, 0.4, 0.3, 2.0, 0.5)
    # Línea naranja
    add_rect(sl, prs, 0.4, 0.95, 1.4, 0.05, OR)
    # Tagline
    add_textbox(sl, 0.4, 1.1, 6.0, 0.28,
                'BUILDING SOLUTIONS. DELIVERING SUCCESS.',
                8, bold=True, color=DK)
    # Descripción
    add_textbox(sl, 0.4, 1.5, 5.8, 1.1,
                'Kruger crea aplicaciones web y móviles que ayudan a las empresas a convertir productos digitales complejos en experiencias simples y centradas en el usuario.',
                9, color=hex_color('555555'), wrap=True)
    # Stats
    stats = [('30+','AÑOS DE INNOVACIÓN'),('1500+','PROYECTOS ENTREGADOS'),
             ('450+','EXPERTOS CERTIFICADOS'),('14','PAÍSES PRESENCIA GLOBAL')]
    for i,(v,l) in enumerate(stats):
        y = 3.0 + i * 0.85
        add_textbox(sl, 0.95, y, 5.5, 0.35, f'{v}  {l}', 12, bold=True, color=OR)
    # Placeholder imagen
    add_placeholder_box(sl, 7.0, 0.3, 5.9, 6.9, 'Foto persona · traje naranja')


def slide_porque(prs, sl):
    set_bg(sl, OR)
    # Número watermark
    add_textbox(sl, 7.5, 0.1, 4.0, 3.2, '01', 120, bold=True, color=WH)
    # WHY KRUGER
    add_textbox(sl, 7.0, 3.1, 5.0, 0.8, 'WHY', 28, color=WH)
    add_textbox(sl, 7.0, 3.85, 5.9, 1.1, 'KRUGER', 46, bold=True, color=WH)
    # Bullets
    puntos = [
        'Transformamos negocios con tecnología centrada en personas',
        '14 países · 30 años · 1500+ proyectos entregados',
        'Expertos en sectores regulados y ecosistemas enterprise'
    ]
    for i, p in enumerate(puntos):
        add_textbox(sl, 7.4, 5.2 + i*0.65, 5.5, 0.55, f'{i+1}.  {p}',
                    10, color=hex_color('FFE0CC'), wrap=True)
    # Logo blanco abajo
    add_logo(sl, LOGO_W, 11.5, 6.9, 1.6, 0.4)
    # Placeholder imagen
    add_placeholder_box(sl, 0.3, 0.3, 6.3, 6.9, 'Persona editorial · traje naranja')


def slide_como(prs, sl):
    set_bg(sl, GR)
    # Esquina naranja
    add_rect(sl, prs, 12.33, 0, 1.0, 1.0, OR)
    # Título
    add_textbox(sl, 0.4, 0.25, 2.6, 0.8, 'What ', 30, bold=True, color=DK)
    add_textbox(sl, 2.4, 0.25, 3.5, 0.8, 'We Build', 30, bold=True, color=OR)
    # Subtítulo
    add_textbox(sl, 0.4, 1.15, 6.0, 0.65,
                'Construimos experiencias digitales escalables que conectan tecnología, operaciones y engagement del cliente.',
                9, color=hex_color('555555'), wrap=True)
    # Línea vertical naranja
    add_rect(sl, prs, 0.4, 1.95, 0.08, 4.4, OR)
    # Items
    items = [
        ('Mobile & Self-Service Apps', 'Journeys digitales para clientes.'),
        ('Payment & Transactional UX', 'Flujos seguros para conversión.'),
        ('Enterprise Portals', 'Procesos complejos simplificados.'),
        ('B2C Digital Experiences', 'Centrados en usabilidad.'),
        ('B2B Platforms', 'Herramientas operativas.'),
        ('UX/UI Systems', 'Interfaces multi-pantalla.'),
    ]
    for i,(t,d) in enumerate(items):
        add_textbox(sl, 0.65, 2.05 + i*0.66, 3.5, 0.38, f'• {t}', 9, bold=True, color=DK)
        add_textbox(sl, 3.5,  2.05 + i*0.66, 2.9, 0.38, f'— {d}', 9, color=LGY)
    # Footer
    add_textbox(sl, 0.4, 7.1, 6.2, 0.28,
                'HUMAN-CENTERED digital products powered by INFINITE DIGITAL TECHNOLOGY.',
                8, bold=True, color=OR)
    # Placeholder
    add_placeholder_box(sl, 7.0, 0.3, 5.9, 6.9, 'Mockups multi-dispositivo')


def slide_dolores(prs, sl, vertical_nm, pain):
    set_bg(sl, DK)
    add_textbox(sl, 0.4, 0.25, 4.5, 0.72, 'Dolores en ', 24, bold=True, color=OR)
    add_textbox(sl, 3.35, 0.25, 9.0, 0.72, vertical_nm, 24, bold=True, color=WH)
    add_rect(sl, prs, 0.4, 1.08, 12.5, 0.04, OR)
    add_placeholder_box(sl, 0.4, 1.25, 2.6, 5.9, 'Ejecutivo analítico')
    for i, p in enumerate(pain):
        col = 0 if i < 3 else 1
        row = i % 3
        x = 3.25 + col * 4.55
        y = 1.3 + row * 2.0
        shape = sl.shapes.add_shape(5, in_(x), in_(y), in_(4.3), in_(1.85))  # roundRect
        shape.fill.solid()
        shape.fill.fore_color.rgb = hex_color('2E2E2E')
        from pptx.util import Pt as PtU
        shape.fill.fore_color.rgb = hex_color('2E2E2E')
        shape.line.color.rgb = OR
        shape.line.width = PtU(0.5)
        add_textbox(sl, x+0.15, y+0.1, 4.0, 0.38,
                    ' '.join(p.split()[:4]), 10, bold=True, color=OR, wrap=True)
        add_textbox(sl, x+0.15, y+0.55, 4.0, 1.1, p, 8, color=MGY, wrap=True)


def slide_case_intro(prs, sl, case, dark):
    bg = DK2 if dark else GR
    tc = WH if dark else DK
    logo_url = LOGO_W if dark else LOGO_B
    set_bg(sl, bg)

    words = case['n'].split()
    mid = (len(words)+1)//2
    line1 = ' '.join(words[:mid])
    line2 = ' '.join(words[mid:]) if len(words) > 1 else ''

    add_textbox(sl, 0.4, 1.2, 6.5, 1.2, line1, 38, bold=True, color=tc, wrap=True)
    if line2:
        add_textbox(sl, 0.4, 2.35, 6.5, 1.2, line2, 38, bold=True, color=OR, wrap=True)
    add_textbox(sl, 0.4, 3.7, 6.3, 0.65, case['sub'], 14, bold=True,
                color=hex_color('EEEEEE') if dark else DK, wrap=True)
    add_rect(sl, prs, 0.4, 4.5, 1.0, 0.07, OR)
    add_textbox(sl, 0.4, 4.7, 6.0, 0.3,
                'CASO DE ÉXITO · KRUGER WORLDWIDE', 8, bold=True,
                color=hex_color('999999') if dark else hex_color('777777'))
    add_logo(sl, logo_url, 0.4, 5.2, 1.8, 0.45)
    add_placeholder_box(sl, 7.1, 0.3, 5.8, 6.9, 'Foto profesional · traje naranja')


def slide_case_detail(prs, sl, case):
    set_bg(sl, BK)
    # Header gris
    add_rect(sl, prs, 0, 0, 13.33, 1.1, GR)
    add_textbox(sl, 0.4, 0.1, 10.0, 0.88, case['n'], 26, bold=True, color=DK)
    add_logo(sl, LOGO_B, 11.5, 0.28, 1.6, 0.55)
    # Sub-header negro
    add_rect(sl, prs, 0, 1.1, 13.33, 0.52, hex_color('111111'))
    add_textbox(sl, 0.4, 1.18, 12.0, 0.38, case['sub'].upper(), 9, bold=True, color=WH)
    # Placeholder imagen
    add_placeholder_box(sl, 0.3, 1.75, 4.5, 5.5, 'Screenshots · Persona')
    # Cards
    cards = [('CHALLENGE', case['ch']), ('APPROACH', case['ap']), ('OUTCOME', case['out'])]
    for i,(lbl,txt) in enumerate(cards):
        y = 1.8 + i * 1.88
        shape = sl.shapes.add_shape(1, in_(5.15), in_(y), in_(7.8), in_(1.75))
        shape.fill.solid()
        shape.fill.fore_color.rgb = hex_color('1A1A1A')
        shape.line.color.rgb = hex_color('333333')
        shape.line.width = Pt(0.5)
        add_textbox(sl, 5.45, y+0.15, 4.0, 0.38, lbl, 11, bold=True, color=OR)
        add_textbox(sl, 5.45, y+0.58, 7.2, 1.08, txt, 9, color=MGY, wrap=True)


def slide_productos(prs, sl, vertical_nm, prods):
    set_bg(sl, GR)
    add_textbox(sl, 0.4, 0.12, 5.5, 0.65, 'Productos para ', 22, bold=True, color=DK)
    add_textbox(sl, 3.85, 0.12, 8.0, 0.65, vertical_nm, 22, bold=True, color=OR)
    add_logo(sl, LOGO_B, 11.5, 0.18, 1.6, 0.45)
    for i, p in enumerate(prods[:4]):
        x = 0.4 if i < 2 else 6.85
        y = 1.0 + (i % 2) * 2.9
        add_rect(sl, prs, x, y, 6.2, 2.72, WH)
        add_rect(sl, prs, x, y, 6.2, 0.12, OR)
        add_textbox(sl, x+0.25, y+0.25, 4.2, 0.38, p['n'], 12, bold=True, color=DK)
        shape = sl.shapes.add_shape(1, in_(x+4.85), in_(y+0.25), in_(0.95), in_(0.32))
        shape.fill.solid(); shape.fill.fore_color.rgb = OR
        shape.line.fill.background()
        add_textbox(sl, x+4.85, y+0.26, 0.95, 0.3, p['cat'], 7, bold=True,
                    color=WH, align=PP_ALIGN.CENTER)
        add_textbox(sl, x+0.25, y+0.78, 5.8, 1.75, p['d'], 9, color=LGY, wrap=True)


def slide_servicios(prs, sl, vertical_nm, servs):
    set_bg(sl, DK)
    add_textbox(sl, 0.4, 0.22, 5.5, 0.72, 'Servicios para ', 24, bold=True, color=WH)
    add_textbox(sl, 3.7, 0.22, 8.2, 0.72, vertical_nm, 24, bold=True, color=OR)
    add_logo(sl, LOGO_W, 11.5, 0.28, 1.6, 0.45)
    add_rect(sl, prs, 0.4, 1.08, 12.5, 0.04, OR)
    for i, s in enumerate(servs):
        x = 0.5 + i * 3.2
        # Círculo naranja
        shape = sl.shapes.add_shape(9, in_(x+1.12), in_(1.75), in_(0.6), in_(0.6))
        shape.fill.solid(); shape.fill.fore_color.rgb = OR
        shape.line.fill.background()
        add_textbox(sl, x+1.12, 1.78, 0.6, 0.55, s['n'], 10, bold=True,
                    color=WH, align=PP_ALIGN.CENTER)
        # Card
        add_rect(sl, prs, x, 2.52, 3.1, 4.7, hex_color('2E2E2E'))
        add_textbox(sl, x+0.1, 2.72, 2.9, 0.45, s['t'], 12, bold=True,
                    color=OR, align=PP_ALIGN.CENTER)
        add_textbox(sl, x+0.1, 3.28, 2.9, 3.6, s['d'], 9,
                    color=hex_color('AAAAAA'), wrap=True, align=PP_ALIGN.CENTER)


def slide_partners(prs, sl):
    set_bg(sl, WH)
    add_logo(sl, LOGO_B, 5.4, 0.22, 2.5, 0.62)
    add_textbox(sl, 0, 0.85, 13.33, 0.85, 'PARTNERS', 38, bold=True,
                color=OR, align=PP_ALIGN.CENTER)
    clientes = ['Amdocs','Verizon','Movistar','Claro','Coca-Cola',
                'Banco del Pacífico','IESS','SRI','Holcim','Pronaca',
                'CNT','Zurich','Banco Nación','Tottus','BCP',
                'Santander','El Ordeño','Rimac','Difare','CNEL']
    for i, l in enumerate(clientes):
        col = i % 5
        row = i // 5
        x = 0.4 + col * 2.5
        y = 1.88 + row * 1.15
        add_rect(sl, prs, x, y, 2.35, 0.95, GR)
        add_textbox(sl, x, y, 2.35, 0.95, l, 9, bold=True,
                    color=LGY, align=PP_ALIGN.CENTER)
    add_textbox(sl, 0, 7.2, 13.33, 0.25,
                '14 países · 1500+ proyectos entregados', 9,
                color=hex_color('AAAAAA'), align=PP_ALIGN.CENTER)


def slide_gancho(prs, sl):
    set_bg(sl, WH)
    add_rect(sl, prs, 0, 0, 0.9, 7.5, OR)
    add_textbox(sl, 1.1, 0.2, 1.0, 1.4, '"', 72, color=OR,
                font_name='Georgia', italic=True)
    add_textbox(sl, 1.2, 1.5, 7.2, 2.8,
                '"El futuro de los negocios es digital. La diferencia entre liderar y quedarse atrás está en la experiencia que ofreces a tus clientes."',
                15, bold=True, italic=True, color=DK, wrap=True)
    add_rect(sl, prs, 1.2, 4.5, 1.5, 0.08, OR)
    add_textbox(sl, 1.2, 4.7, 7.5, 0.8,
                'Kruger ha transformado más de 1500 productos digitales en 14 países.',
                10, color=LGY, wrap=True)
    add_placeholder_box(sl, 9.1, 0.3, 3.9, 6.9, 'Imagen gancho naranja')


def slide_cierre(prs, sl, cfg):
    set_bg(sl, OR)
    add_textbox(sl, 0.4, 0.3, 7.0, 0.38, cfg.get('name','Nombre Apellido'),
                11, bold=True, color=WH)
    add_textbox(sl, 0.4, 0.68, 7.0, 0.3, cfg.get('title',''), 9,
                color=hex_color('FFD0B0'))
    add_textbox(sl, 7.0, 0.3, 5.9, 0.38, 'KRUGER CORPORATION', 11,
                bold=True, color=WH, align=PP_ALIGN.RIGHT)
    add_textbox(sl, 7.0, 0.68, 5.9, 0.3, 'krugerworldwide.com', 9,
                color=hex_color('FFD0B0'), align=PP_ALIGN.RIGHT)
    add_logo(sl, LOGO_W, 5.2, 1.6, 2.9, 0.72)
    add_rect(sl, prs, 3.0, 2.55, 7.3, 0.06, hex_color('FFFFFF'))
    add_placeholder_box(sl, 5.65, 2.8, 2.0, 2.0, 'Camaleón 3D')
    add_textbox(sl, 0, 5.05, 13.33, 0.4,
                cfg.get('email','worldwide@krugercorp.com') + (f" · {cfg['phone']}" if cfg.get('phone') else ''),
                11, bold=True, color=WH, align=PP_ALIGN.CENTER)
    if cfg.get('client'):
        add_textbox(sl, 0, 5.6, 13.33, 0.35,
                    f"Preparado especialmente para {cfg['client']}",
                    10, color=hex_color('FFE8D6'), align=PP_ALIGN.CENTER,
                    italic=True)
    add_textbox(sl, 0.4, 7.15, 5.0, 0.22, '© 2026 KRUGER CORP.', 7, color=WH)
    add_textbox(sl, 8.0, 7.15, 5.0, 0.22, 'Confidencial', 7,
                color=WH, align=PP_ALIGN.RIGHT)


# ── ENDPOINT ──────────────────────────────────────────────────

@app.route('/generate', methods=['POST'])
def generate():
    data = request.json
    vertical = data.get('vertical', {})
    cfg = data.get('cfg', {})

    prs = Presentation()
    prs.slide_width  = W
    prs.slide_height = H
    blank = prs.slide_layouts[6]  # completamente en blanco

    def new_slide():
        return prs.slides.add_slide(blank)

    # Portada
    slide_portada(prs, new_slide(), cfg)
    # Disclaimer
    slide_disclaimer(prs, new_slide())
    # Quiénes somos
    slide_quienes(prs, new_slide())
    # Por qué Kruger
    slide_porque(prs, new_slide())
    # Cómo lo hacemos
    slide_como(prs, new_slide())
    # Dolores
    slide_dolores(prs, new_slide(), vertical.get('nm',''), vertical.get('pain',[]))
    # Cases
    for i, case in enumerate(vertical.get('cases', [])):
        slide_case_intro(prs, new_slide(), case, i % 2 == 0)
        slide_case_detail(prs, new_slide(), case)
    # Productos
    slide_productos(prs, new_slide(), vertical.get('nm',''), vertical.get('prods',[]))
    # Servicios
    slide_servicios(prs, new_slide(), vertical.get('nm',''), vertical.get('servs',[]))
    # Partners
    slide_partners(prs, new_slide())
    # Gancho
    slide_gancho(prs, new_slide())
    # Cierre
    slide_cierre(prs, new_slide(), cfg)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)

    nm = vertical.get('nm','Vertical').replace(' ','_')
    cl = cfg.get('client','Cliente').replace(' ','_')
    filename = f'Kruger_{nm}_{cl}_2026.pptx'

    return send_file(
        buf,
        mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
        as_attachment=True,
        download_name=filename
    )


@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=False)
