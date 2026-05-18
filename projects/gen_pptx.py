#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Навигатор специалиста: Пептидные комплексы
Структура: PDF Peptide_Specialist_Playbook | Стиль: Geneline brand (тёмный)
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree
import os

# ── DARK GENELINE PALETTE ─────────────────────────────────────────────────────
BG    = RGBColor(0x28, 0x12, 0x08)   # slide background (very dark brown)
CARD  = RGBColor(0x3C, 0x1E, 0x0C)   # card fill
CARD2 = RGBColor(0x4E, 0x28, 0x12)   # card hover / alt
TERRA = RGBColor(0xBF, 0x78, 0x50)   # terracotta — main accent
GOLD  = RGBColor(0xC8, 0x9A, 0x50)   # gold — secondary accent
MED   = RGBColor(0x9E, 0x6A, 0x3A)   # medium brown
DARK  = RGBColor(0x7A, 0x40, 0x20)   # dark brown
CREAM = RGBColor(0xF0, 0xE8, 0xD8)   # cream (headlines)
BODY  = RGBColor(0xD5, 0xC5, 0xAA)   # body text (warm grey)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
RED_W = RGBColor(0xC0, 0x38, 0x18)   # red warning
GRN_W = RGBColor(0x5A, 0x8A, 0x40)   # green ok

# ── LAYOUT ────────────────────────────────────────────────────────────────────
W = Inches(13.33); H = Inches(7.5); M = Inches(0.5)
IMG = '/tmp/pptx_images'
prs = Presentation(); prs.slide_width = W; prs.slide_height = H
BL  = prs.slide_layouts[6]

# ── HELPERS ───────────────────────────────────────────────────────────────────
def ns(): return prs.slides.add_slide(BL)

def bg(s, c=BG):
    f = s.background.fill; f.solid(); f.fore_color.rgb = c

def box(s, l, t, w, h, fc, lc=None, rnd=False, lw=1.0):
    sh = s.shapes.add_shape(5 if rnd else 1, l, t, w, h)
    sh.fill.solid(); sh.fill.fore_color.rgb = fc
    if lc: sh.line.color.rgb = lc; sh.line.width = Pt(lw)
    else:   sh.line.fill.background()
    return sh

def tx(s, text, l, t, w, h, sz=13, bold=False, c=BODY,
       al=PP_ALIGN.LEFT, it=False, wrap=True):
    tb = s.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame; tf.word_wrap = wrap
    p = tf.paragraphs[0]; p.alignment = al
    r = p.add_run()
    r.text = text; r.font.size = Pt(sz); r.font.bold = bold
    r.font.italic = it; r.font.color.rgb = c; r.font.name = 'Calibri'
    return tb

def multi_para(s, paras, l, t, w, h, sz=12, c=BODY, bc=TERRA, sp=6):
    """Bullet list"""
    tb = s.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame; tf.word_wrap = True
    for i, (bullet, text) in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        if bullet:
            rb = p.add_run(); rb.text = bullet+'  '
            rb.font.color.rgb = bc; rb.font.size = Pt(sz-1)
            rb.font.name = 'Calibri'; rb.font.bold = True
        r = p.add_run(); r.text = text
        r.font.size = Pt(sz); r.font.color.rgb = c; r.font.name = 'Calibri'
        p.space_after = Pt(sp)
    return tb

def buls(s, items, l, t, w, h, sz=12, c=BODY, bc=TERRA, sp=6, bullet='●'):
    paras = [(bullet, item) for item in items]
    return multi_para(s, paras, l, t, w, h, sz=sz, c=c, bc=bc, sp=sp)

def logo(s):
    tx(s, 'GENELINE', W-Inches(2.1), Inches(0.15), Inches(1.9),
       Inches(0.38), sz=12, bold=True, c=TERRA, al=PP_ALIGN.RIGHT)

def pic(s, path, l, t, w, h):
    if os.path.exists(path): return s.shapes.add_picture(path, l, t, w, h)

def set_alpha(shape, pct):
    sp = shape._element
    sf = sp.find('.//' + qn('a:solidFill'))
    if sf is None: return
    clr = sf.find('*')
    if clr is None: return
    a = etree.SubElement(clr, qn('a:alpha'))
    a.set('val', str(int((100-pct)*1000)))

def connector(s, x1, y1, x2, y2, col=TERRA, w=2):
    from pptx.enum.shapes import MSO_CONNECTOR_TYPE
    c = s.shapes.add_connector(MSO_CONNECTOR_TYPE.STRAIGHT, x1, y1, x2, y2)
    c.line.color.rgb = col; c.line.width = Pt(w)
    return c

def icon_circle(s, sym, l, t, sz_in, font_sz, bg_col, fg_col=WHITE):
    sh = s.shapes.add_shape(9, l, t, sz_in, sz_in)
    sh.fill.solid(); sh.fill.fore_color.rgb = bg_col
    sh.line.fill.background()
    tx(s, sym, l, t, sz_in, sz_in, sz=font_sz, bold=True,
       c=fg_col, al=PP_ALIGN.CENTER)
    return sh

def card(s, l, t, w, h, title, title_col=TERRA, body_lines=None,
         icon=None, icon_col=TERRA, fc=CARD, lc=None, title_sz=13):
    box(s, l, t, w, h, fc, lc=lc, rnd=True, lw=1.5)
    if icon:
        icon_circle(s, icon, l+Inches(0.18), t+Inches(0.15),
                    Inches(0.42), 14, icon_col)
        tx(s, title, l+Inches(0.7), t+Inches(0.18),
           w-Inches(0.85), Inches(0.45), sz=title_sz, bold=True, c=title_col)
    else:
        tx(s, title, l+Inches(0.22), t+Inches(0.18),
           w-Inches(0.44), Inches(0.45), sz=title_sz, bold=True, c=title_col)
    if body_lines:
        top_offset = Inches(0.72)
        tx(s, '\n'.join(body_lines), l+Inches(0.22), t+top_offset,
           w-Inches(0.44), h-top_offset-Inches(0.1), sz=11.5, c=BODY)

def slide_header(s, title, subtitle=None, col=CREAM):
    logo(s)
    tx(s, title, M, Inches(0.18), W-2*M, Inches(0.72),
       sz=30, bold=True, c=col)
    if subtitle:
        tx(s, subtitle, M, Inches(0.92), W-2*M, Inches(0.38),
           sz=13, c=MED, it=True)
    box(s, M, Inches(1.35), W-2*M, Pt(1.5), TERRA)

def cluster_card(s, l, t, w, h, name, subtitle, items, fc=CARD, lc=TERRA):
    box(s, l, t, w, h, fc, lc=lc, rnd=True, lw=1.5)
    # header
    tx(s, name, l+Inches(0.2), t+Inches(0.15), w-Inches(0.4),
       Inches(0.5), sz=16, bold=True, c=CREAM)
    tx(s, subtitle, l+Inches(0.2), t+Inches(0.65), w-Inches(0.4),
       Inches(0.35), sz=11, c=TERRA, it=True)
    box(s, l+Inches(0.2), t+Inches(1.05), w-Inches(0.4), Pt(1), TERRA)
    # items: list of (label, text) tuples
    y_off = Inches(1.18)
    for label, text in items:
        icon_circle(s, '◎', l+Inches(0.2), t+y_off, Inches(0.38), 11,
                    DARK, TERRA)
        tb2 = s.shapes.add_textbox(l+Inches(0.68), t+y_off,
                                   w-Inches(0.88), Inches(0.72))
        tf2 = tb2.text_frame; tf2.word_wrap = True
        p2 = tf2.paragraphs[0]
        r1 = p2.add_run(); r1.text = label+': '
        r1.font.bold = True; r1.font.size = Pt(11.5)
        r1.font.color.rgb = GOLD; r1.font.name = 'Calibri'
        r2 = p2.add_run(); r2.text = text
        r2.font.size = Pt(11.5); r2.font.color.rgb = BODY
        r2.font.name = 'Calibri'
        y_off += Inches(0.82)
        # tag pill
        tag_text = text[:25]+'…' if len(text) > 25 else text
        # mini separator line
        box(s, l+Inches(0.68), t+y_off-Inches(0.12),
            w-Inches(0.88), Pt(0.5), CARD2)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 1: TITLE
# ═══════════════════════════════════════════════════════════════════════════════
s = ns(); bg(s)
# watermark G
tx(s,'G',Inches(8.5),Inches(-0.3),Inches(5.5),Inches(8.5),
   sz=420,c=RGBColor(0x38,0x1C,0x0C),bold=True)
# top accent stripe
box(s, 0, 0, W, Inches(0.08), TERRA)
# left brand stripe
box(s, 0, 0, Inches(0.3), H, TERRA)
# main title block
box(s, Inches(0.5), Inches(0.8), Inches(8.5), Inches(2.4), CARD, rnd=True)
box(s, Inches(0.5), Inches(0.8), Inches(0.12), Inches(2.4), TERRA)
tx(s,'Навигатор специалиста:',Inches(0.8),Inches(0.95),Inches(8.0),
   Inches(0.75),sz=32,bold=True,c=CREAM)
tx(s,'Пептидные комплексы',Inches(0.8),Inches(1.68),Inches(8.0),
   Inches(0.85),sz=38,bold=True,c=WHITE)
tx(s,'Инструмент точной регуляции здоровья',
   Inches(0.8),Inches(2.5),Inches(8.0),Inches(0.45),sz=15,c=TERRA,it=True)
# 3 nav badges
badges = [('◉  10 комплексов', TERRA),
          ('⬡  Алгоритмы подбора', MED),
          ('◈  Сценарии продаж', GOLD)]
bw = Inches(3.7); bh = Inches(0.72)
for i,(txt,col) in enumerate(badges):
    bx = Inches(0.5) + i*(bw+Inches(0.2))
    box(s, bx, Inches(3.55), bw, bh, CARD2, lc=col, rnd=True, lw=2)
    tx(s, txt, bx+Inches(0.25), Inches(3.65), bw-Inches(0.5), bh-Inches(0.1),
       sz=15, bold=True, c=col, al=PP_ALIGN.CENTER)
# 3 bottom feature cards
features = [('Клинические данные','Подбор по симптомам\nи жалобам'),
            ('Модули воздействия','Механизм каждого\nкомплекса'),
            ('Безопасность','Красные флаги\nи противопоказания')]
fw = Inches(3.7); fh = Inches(2.4)
for i,(title,desc) in enumerate(features):
    fx = Inches(0.5) + i*(fw+Inches(0.2))
    box(s, fx, Inches(4.55), fw, fh, CARD, rnd=True, lc=DARK, lw=1)
    tx(s, title, fx+Inches(0.25), Inches(4.68), fw-Inches(0.5), Inches(0.5),
       sz=14, bold=True, c=CREAM)
    box(s, fx+Inches(0.25), Inches(5.22), fw-Inches(0.5), Pt(1), TERRA)
    tx(s, desc, fx+Inches(0.25), Inches(5.35), fw-Inches(0.5), Inches(1.5),
       sz=12, c=BODY)
box(s,M,Inches(7.12),W-2*M,Inches(0.27),RGBColor(0x38,0x1C,0x0C),rnd=True)
tx(s,'Маркетинговая версия с акцентом на безопасность, этику и работу специалиста.',
   M+Inches(0.3),Inches(7.16),W-2*M-Inches(0.6),Inches(0.25),
   sz=11,c=BODY,al=PP_ALIGN.CENTER)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 2: ЧТО ТАКОЕ ПЕПТИДЫ / ПРИНЦИП КЛЮЧ-ЗАМОК
# ═══════════════════════════════════════════════════════════════════════════════
s = ns(); bg(s); logo(s)
slide_header(s,'Основы: Что такое пептиды',
             'Понять механизм — значит уметь объяснить клиенту')
# Left card: ЧТО ЭТО
lw = Inches(6.1)
box(s, M, Inches(1.52), lw, Inches(5.68), CARD, rnd=True, lc=TERRA, lw=1.5)
box(s, M, Inches(1.52), lw, Inches(0.1), TERRA)
tx(s,'ЧТО ЭТО?', M+Inches(0.25), Inches(1.62), lw-Inches(0.5),
   Inches(0.5), sz=16, bold=True, c=TERRA)
# molecule image
pic(s, f'{IMG}/molecule.png', M+Inches(0.15), Inches(2.2),
    lw-Inches(0.3), Inches(2.2))
# stat badges
for i, (num, desc) in enumerate([('< 5000 Да','молекулярная масса'),
                                   ('до 50','аминокислот в цепочке')]):
    bx2 = M+Inches(0.25)+i*Inches(2.8)
    box(s, bx2, Inches(4.52), Inches(2.55), Inches(0.65), CARD2, rnd=True, lc=TERRA, lw=1)
    tx(s, num, bx2+Inches(0.15), Inches(4.58), Inches(1.2), Inches(0.35),
       sz=14, bold=True, c=TERRA)
    tx(s, desc, bx2+Inches(1.35), Inches(4.64), Inches(1.0), Inches(0.35),
       sz=10, c=BODY)
tx(s,'Наш организм построен из белков. Если белок «разрезать» на короткие фрагменты, получатся пептиды. Это не строительный материал, а сигнальные молекулы.',
   M+Inches(0.25),Inches(5.28),lw-Inches(0.5),Inches(1.5),sz=12,c=BODY)
box(s,M+Inches(0.25),Inches(6.58),lw-Inches(0.5),Inches(0.45),
    RGBColor(0x55,0x28,0x10),rnd=True)
tx(s,'ПЕПТИДЫ — ЭТО ДИРИЖЁРЫ ОРГАНИЗМА.',
   M+Inches(0.4),Inches(6.65),lw-Inches(0.8),Inches(0.35),
   sz=12,bold=True,c=TERRA,al=PP_ALIGN.CENTER)

# Right card: ПРИНЦИП КЛЮЧ-ЗАМОК
rl = M+lw+Inches(0.35); rw = W-M-rl
box(s, rl, Inches(1.52), rw, Inches(5.68), CARD, rnd=True, lc=GOLD, lw=1.5)
box(s, rl, Inches(1.52), rw, Inches(0.1), GOLD)
tx(s,'ПРИНЦИП КЛЮЧ–ЗАМОК', rl+Inches(0.25), Inches(1.62),
   rw-Inches(0.5), Inches(0.5), sz=16, bold=True, c=GOLD)
pic(s, f'{IMG}/key_lock.png', rl+Inches(0.15), Inches(2.2),
    rw-Inches(0.3), Inches(2.2))
tx(s,'Пептид действует на специфические рецепторы клеток. Когда пептид связывается с рецептором, клетка получает чёткую команду: усилить или замедлить процесс, синтезировать гормон или фермент.',
   rl+Inches(0.25),Inches(4.55),rw-Inches(0.5),Inches(1.7),sz=12,c=BODY)
box(s,rl+Inches(0.25),Inches(6.38),rw-Inches(0.5),Inches(0.7),
    RGBColor(0x55,0x30,0x08),rnd=True)
tx(s,'Они не подменяют систему,\nа возвращают ей правильный ритм.',
   rl+Inches(0.4),Inches(6.45),rw-Inches(0.8),Inches(0.6),
   sz=12,bold=True,c=GOLD,al=PP_ALIGN.CENTER)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 3: СМЕНА ПАРАДИГМЫ — ПЕПТИДЫ VS КЛАССИКА
# ═══════════════════════════════════════════════════════════════════════════════
s = ns(); bg(s); logo(s)
slide_header(s,'Смена парадигмы: Пептиды vs. Привычные препараты')
# Two column headers
col_w = (W-2*M-Inches(0.5))/2
lc_x = M; rc_x = M+col_w+Inches(0.5)
box(s, lc_x, Inches(1.52), col_w, Inches(0.6), CARD2, rnd=True, lc=DARK, lw=1.5)
tx(s,'⊗  Классические препараты', lc_x+Inches(0.2), Inches(1.6),
   col_w-Inches(0.4), Inches(0.5), sz=14, bold=True, c=DARK)
box(s, rc_x, Inches(1.52), col_w, Inches(0.6), CARD2, rnd=True, lc=TERRA, lw=1.5)
tx(s,'◉  Пептидные комплексы', rc_x+Inches(0.2), Inches(1.6),
   col_w-Inches(0.4), Inches(0.5), sz=14, bold=True, c=TERRA)
# comparison rows
rows = [
    ('⚙  Механизм действия',
     'Замещение или подавление функций (работают «за» организм).',
     'Мягкая регуляция и настройка (помогают системе работать естественно).'),
    ('◎  Точность',
     'Широкое системное воздействие.',
     'Действуют точечно — «ключ под замок» на свои рецепторы.'),
    ('⛓  Риски привыкания',
     'Возможен синдром отмены и зависимость.',
     'Не вызывают привыкания и резкого синдрома отмены.'),
    ('◈  Зона применения',
     'Купирование острых симптомов.',
     'Работа там, где симптом сохраняется, а классические препараты дают слабый эффект.'),
]
row_h = Inches(1.2)
for i, (label, left, right) in enumerate(rows):
    y = Inches(2.22) + i*row_h
    # separator label
    box(s, M, y, W-2*M, Inches(0.3), RGBColor(0x35,0x18,0x08))
    tx(s, label, M+Inches(0.2), y+Inches(0.04), W-2*M-Inches(0.4),
       Inches(0.28), sz=11, bold=True, c=GOLD)
    # left cell
    box(s, lc_x, y+Inches(0.32), col_w, row_h-Inches(0.4), CARD, rnd=True)
    tx(s, left, lc_x+Inches(0.18), y+Inches(0.42),
       col_w-Inches(0.36), row_h-Inches(0.6), sz=12, c=BODY)
    # right cell
    box(s, rc_x, y+Inches(0.32), col_w, row_h-Inches(0.4), CARD, rnd=True,
        lc=TERRA, lw=0.75)
    tx(s, right, rc_x+Inches(0.18), y+Inches(0.42),
       col_w-Inches(0.36), row_h-Inches(0.6), sz=12, c=CREAM)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 4: МАТРИЦА ПОДБОРА
# ═══════════════════════════════════════════════════════════════════════════════
s = ns(); bg(s); logo(s)
slide_header(s,'Матрица подбора: Главный дашборд специалиста',
             'Быстрый ориентир на старте разговора с клиентом')
matrix = [
    ('🧠','«Голова не варит», рассеянность','ACTIVEBRAIN',TERRA),
    ('☁','Хронический стресс, выгорание','NO STRESS + STRESSRELIEF',MED),
    ('🌙','Не могу заснуть, не восстанавливаюсь','RECOVERY / NO STRESS',DARK),
    ('🛡','Часто болею, долго восстанавливаюсь','IMMUNACTIV',TERRA),
    ('◎','Частые проблемы со слизистыми','STOPBACTERIA',MED),
    ('〰','Хроническая боль, гиперчувствительность','RELIEF',DARK),
    ('◉','Социальная тревога, «холодность»','OXYGEN',TERRA),
    ('♡','«Нет желания», переутомление','TESTOBOOSTER + AMORE',GOLD),
    ('⚡','Снижение либидо (гормоны в норме)','TESTOBOOSTER',GOLD),
    ('☽','Нерегулярный цикл, аменорея','AMORE (только после врача)',MED),
]
row_h = Inches(0.55); col1 = Inches(4.8); col2 = Inches(4.5)
tt = Inches(1.52)
# headers
box(s, M, tt, col1, row_h, CARD2)
box(s, M+col1+Inches(0.08), tt, col2+Inches(3.0), row_h, CARD2)
tx(s,'ЖАЛОБА КЛИЕНТА', M+Inches(0.15), tt+Inches(0.1),
   col1-Inches(0.3), row_h-Inches(0.15), sz=11, bold=True, c=CREAM, al=PP_ALIGN.CENTER)
tx(s,'ПРИОРИТЕТНЫЙ КОМПЛЕКС', M+col1+Inches(0.25), tt+Inches(0.1),
   col2+Inches(2.7), row_h-Inches(0.15), sz=11, bold=True, c=CREAM, al=PP_ALIGN.CENTER)
for i,(ico, complaint, complex_name, accent) in enumerate(matrix):
    y = tt + row_h + i*row_h
    fc2 = RGBColor(0x38,0x1C,0x0C) if i%2==0 else RGBColor(0x30,0x16,0x08)
    box(s, M, y, col1, row_h, fc2)
    box(s, M+col1+Inches(0.08), y, col2+Inches(3.0), row_h, fc2)
    # separator
    box(s, M+col1, y, Inches(0.08), row_h, BG)
    # complaint
    tx(s, ico+'  '+complaint, M+Inches(0.15), y+Inches(0.1),
       col1-Inches(0.3), row_h-Inches(0.15), sz=11.5, c=BODY)
    # complex badge
    bw3 = Inches(len(complex_name)*0.085+0.5)
    box(s, M+col1+Inches(0.25), y+Inches(0.1), bw3, row_h-Inches(0.2),
        accent, rnd=True)
    tx(s, complex_name, M+col1+Inches(0.3), y+Inches(0.12),
       bw3-Inches(0.1), row_h-Inches(0.25), sz=11.5, bold=True,
       c=WHITE, al=PP_ALIGN.CENTER)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 5: КЛАСТЕР I — НЕРВНАЯ СИСТЕМА И МОЗГ
# ═══════════════════════════════════════════════════════════════════════════════
s = ns(); bg(s); logo(s)
slide_header(s,'Кластер I: Нервная система и мозг')
cw5 = (W-2*M-Inches(0.4))/3; ch5 = Inches(5.62)
clusters_brain = [
    ('ACTIVEBRAIN','(Ясность мышления)',
     [('Для кого','Высокая умственная нагрузка, жалобы на «туман в голове».'),
      ('Механизм','Оптимизирует работу мозга под нагрузкой. Защищает нейроны. Не стимулятор.'),
      ('Скрипт','«Это не таблетка концентрации. Он помогает мозгу выйти из перегрузки и работать спокойно.»')],
     TERRA, ['Умственная нагрузка','Защита нейронов']),
    ('NO STRESS','(Переключатель)',
     [('Для кого','Хронический стресс, «нервная система не выключается», тревожная бессонница.'),
      ('Механизм','Возвращает работу «педали тормоза», помогает естественно переключиться из режима тревоги.'),
      ('Скрипт','«Это не снотворное. Это переключатель из боевого режима в режим восстановления.»')],
     MED, ['Хронический стресс','Бессонница']),
    ('STRESSRELIEF','(Восстановление ресурсов)',
     [('Для кого','Эмоциональное истощение, выгорание, снижение памяти на фоне перегрузки.'),
      ('Механизм','Поддержка нейропластичности. Помогает нейронам выживать и восстанавливать связи.'),
      ('Скрипт','«Это не антидепрессант. Он помогает мозгу заново обрести устойчивость к стрессу.»')],
     GOLD, ['Выгорание','Снижение памяти']),
]
for i, (name, subtitle, items, accent, tags) in enumerate(clusters_brain):
    l5 = M + i*(cw5+Inches(0.2)); t5 = Inches(1.52)
    box(s, l5, t5, cw5, ch5, CARD, rnd=True, lc=accent, lw=1.5)
    # header
    tx(s, name, l5+Inches(0.2), t5+Inches(0.15), cw5-Inches(0.4),
       Inches(0.52), sz=17, bold=True, c=accent)
    tx(s, subtitle, l5+Inches(0.2), t5+Inches(0.68), cw5-Inches(0.4),
       Inches(0.32), sz=11, c=BODY, it=True)
    box(s, l5+Inches(0.2), t5+Inches(1.04), cw5-Inches(0.4), Pt(1), accent)
    y_off = Inches(1.18)
    for label, text in items:
        icon_circle(s, '◎', l5+Inches(0.2), t5+y_off, Inches(0.38), 11, DARK, accent)
        tb2 = s.shapes.add_textbox(l5+Inches(0.68), t5+y_off, cw5-Inches(0.88), Inches(0.85))
        tf2 = tb2.text_frame; tf2.word_wrap = True
        p2 = tf2.paragraphs[0]
        r1 = p2.add_run(); r1.text = label+': '
        r1.font.bold = True; r1.font.size = Pt(11); r1.font.color.rgb = GOLD; r1.font.name='Calibri'
        r2 = p2.add_run(); r2.text = text
        r2.font.size = Pt(11); r2.font.color.rgb = BODY; r2.font.name='Calibri'
        y_off += Inches(1.0)
    # tags at bottom
    tx_y = t5 + Inches(4.35)
    box(s, l5+Inches(0.2), tx_y, cw5-Inches(0.4), Pt(1), DARK)
    tx_x = l5+Inches(0.2)
    for tag in tags:
        tw = Inches(len(tag)*0.085+0.35)
        box(s, tx_x, tx_y+Inches(0.12), tw, Inches(0.32), DARK, rnd=True)
        tx(s, tag, tx_x+Inches(0.08), tx_y+Inches(0.15), tw-Inches(0.16),
           Inches(0.25), sz=9, c=BODY, al=PP_ALIGN.CENTER)
        tx_x += tw + Inches(0.1)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 6: КЛАСТЕР II — ВОССТАНОВЛЕНИЕ И СНЯТИЕ БОЛИ
# ═══════════════════════════════════════════════════════════════════════════════
s = ns(); bg(s); logo(s)
slide_header(s,'Кластер II: Восстановление и снятие боли')
hw = (W-2*M-Inches(0.4))/2; ht6 = Inches(5.62)
clusters_recover = [
    ('RECOVERY','Глубокое восстановление',
     [('Для кого','Поверхностный сон, недосып, джетлаг, отсутствие утренней бодрости.'),
      ('Механизм','Индуктор медленноволнового сна. Заново выстраивает нормальную архитектуру сна для починки организма.'),
      ('Скрипт','«Это не "уснуть любой ценой". Это возвращение той глубины сна, в которой мозг и тело действительно отдыхают.»')],
     TERRA, 'sleep_chart.png'),
    ('RELIEF','Снижение гиперчувствительности',
     [('Для кого','Хроническая и нейропатическая боль, постнагрузочные состояния, воспалительный фон.'),
      ('Механизм','Снижает гипервозбудимость ЦНС и «расшумляет» систему. Боль снижается естественно.'),
      ('Скрипт','«Это не обезболивающее в привычном смысле. Он работает с тем, насколько резко нервная система реагирует на сигнал.»')],
     MED, 'pain_wave.png'),
]
for i, (name, subtitle, items, accent, chart_file) in enumerate(clusters_recover):
    l6 = M + i*(hw+Inches(0.4))
    box(s, l6, Inches(1.52), hw, ht6, CARD, rnd=True, lc=accent, lw=1.5)
    tx(s, name, l6+Inches(0.2), Inches(1.65), hw-Inches(0.4),
       Inches(0.55), sz=18, bold=True, c=accent)
    tx(s, '('+subtitle+')', l6+Inches(0.2), Inches(2.2), hw-Inches(0.4),
       Inches(0.32), sz=11.5, c=BODY, it=True)
    # chart image
    pic(s, f'{IMG}/{chart_file}', l6+Inches(0.15), Inches(2.58),
        hw-Inches(0.3), Inches(1.3))
    # items
    y_off = Inches(4.0)
    for label, text in items:
        icon_circle(s,'◎', l6+Inches(0.2), Inches(1.52)+y_off,
                    Inches(0.35),10,DARK,accent)
        tb3 = s.shapes.add_textbox(l6+Inches(0.65),Inches(1.52)+y_off,
                                   hw-Inches(0.85),Inches(0.82))
        tf3=tb3.text_frame; tf3.word_wrap=True; p3=tf3.paragraphs[0]
        r1=p3.add_run(); r1.text=label+': '
        r1.font.bold=True; r1.font.size=Pt(10.5); r1.font.color.rgb=GOLD; r1.font.name='Calibri'
        r2=p3.add_run(); r2.text=text
        r2.font.size=Pt(10.5); r2.font.color.rgb=BODY; r2.font.name='Calibri'
        y_off+=Inches(0.88)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 7: КЛАСТЕР III — ИММУНИТЕТ И ЗАЩИТА
# ═══════════════════════════════════════════════════════════════════════════════
s = ns(); bg(s); logo(s)
slide_header(s,'Кластер III: Иммунитет и защита слизистых')
hw7 = (W-2*M-Inches(0.4))/2; ht7 = Inches(5.62)
clusters_immune = [
    ('IMMUNACTIV','Нейроиммунная регуляция',
     [('Для кого','Часто болеющие, долго восстанавливающиеся, состояния «тревожусь — болею».'),
      ('Механизм','Помогает клеткам врождённого иммунитета выйти из хронического напряжения. Успокаивает нервную систему.'),
      ('Скрипт','«Это не подстёгивание иммунитета. Это аккуратная настройка, чтобы клетки реагировали вовремя, а голова успокаивалась.»')],
     TERRA, 'IMMUNACTIV.png'),
    ('STOPBACTERIA','Передовая защита',
     [('Для кого','Частые инфекции верхних дыхательных путей, медленно заживающие повреждения.'),
      ('Механизм','Встраивается в мембраны микробов и разрушает их, координируя местное заживление.'),
      ('Скрипт','«Это не антибиотик. Это пептид собственной защитной системы, к которому сложнее формируется устойчивость бактерий.»')],
     MED, 'STOPBACTERIA.png'),
]
for i, (name, subtitle, items, accent, img_file) in enumerate(clusters_immune):
    l7 = M + i*(hw7+Inches(0.4))
    box(s, l7, Inches(1.52), hw7, ht7, CARD, rnd=True, lc=accent, lw=1.5)
    tx(s, name, l7+Inches(0.2), Inches(1.65), hw7-Inches(0.4),
       Inches(0.55), sz=18, bold=True, c=accent)
    tx(s, '('+subtitle+')', l7+Inches(0.2), Inches(2.2), hw7-Inches(0.4),
       Inches(0.32), sz=11.5, c=BODY, it=True)
    pic(s, f'{IMG}/{img_file}', l7+Inches(0.15), Inches(2.58),
        hw7-Inches(0.3), Inches(1.3))
    y_off = Inches(4.0)
    for label, text in items:
        icon_circle(s,'◎',l7+Inches(0.2),Inches(1.52)+y_off,Inches(0.35),10,DARK,accent)
        tb4=s.shapes.add_textbox(l7+Inches(0.65),Inches(1.52)+y_off,hw7-Inches(0.85),Inches(0.82))
        tf4=tb4.text_frame; tf4.word_wrap=True; p4=tf4.paragraphs[0]
        r1=p4.add_run(); r1.text=label+': '
        r1.font.bold=True; r1.font.size=Pt(10.5); r1.font.color.rgb=GOLD; r1.font.name='Calibri'
        r2=p4.add_run(); r2.text=text
        r2.font.size=Pt(10.5); r2.font.color.rgb=BODY; r2.font.name='Calibri'
        y_off+=Inches(0.88)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 8: КЛАСТЕР IV — ЭМОЦИИ, КОНТАКТ И ЛИБИДО
# ═══════════════════════════════════════════════════════════════════════════════
s = ns(); bg(s); logo(s)
slide_header(s,'Кластер IV: Эмоции, контакт и либидо')
cw8 = (W-2*M-Inches(0.4))/3; ch8 = Inches(5.62)
clusters_emo = [
    ('OXYGEN','Спокойствие и доверие',
     [('Для кого','Социальная настороженность, эмоциональное выгорание, «барьер с людьми».'),
      ('Механизм','Перенастраивает датчики безопасности мозга. Снижает гиперактивацию тревожных контуров.'),
      ('Скрипт','«Это не эйфория. Это мягкая регуляция контуров мозга, отвечающих за ощущение безопасности и контакта.»')],
     TERRA, 'OXYGEN.png'),
    ('TESTOBOOSTER','Центральные механизмы',
     [('Для кого','Функциональное снижение либидо на фоне стресса и истощения (без гормональных сбоев).'),
      ('Механизм','Аккуратно «будит» центральные нейронные сети желания и мотивации в мозге.'),
      ('Скрипт','«Это не тестостерон и не виагра. Он работает с механизмами желания в мозге, когда причина в усталости.»')],
     GOLD, 'TESTOBOOSTER.png'),
    ('AMORE','Регулятор оси',
     [('Для кого','Сбои цикла, функциональное бесплодие, снижение либидо (строго после оценки врачом).'),
      ('Механизм','Подаёт сигнал гипоталамусу для естественного запуска синтеза собственных гормонов.'),
      ('Скрипт','«Это не гормоны извне. Это сигнал телу включить собственную систему так, как задумано природой.»')],
     MED, 'AMORE.png'),
]
for i,(name,subtitle,items,accent,img_file) in enumerate(clusters_emo):
    l8=M+i*(cw8+Inches(0.2)); t8=Inches(1.52)
    box(s,l8,t8,cw8,ch8,CARD,rnd=True,lc=accent,lw=1.5)
    pic(s,f'{IMG}/{img_file}',l8+Inches(0.1),t8+Inches(0.1),cw8-Inches(0.2),Inches(1.6))
    ov=box(s,l8+Inches(0.1),t8+Inches(0.1),cw8-Inches(0.2),Inches(1.6),CARD,rnd=True)
    set_alpha(ov,60)
    tx(s,name,l8+Inches(0.2),t8+Inches(0.18),cw8-Inches(0.4),Inches(0.52),
       sz=16,bold=True,c=accent)
    tx(s,'('+subtitle+')',l8+Inches(0.2),t8+Inches(0.7),cw8-Inches(0.4),
       Inches(0.32),sz=10.5,c=BODY,it=True)
    box(s,l8+Inches(0.2),t8+Inches(1.07),cw8-Inches(0.4),Pt(1),accent)
    y_off=Inches(1.2)
    for label,text in items:
        icon_circle(s,'◎',l8+Inches(0.2),t8+y_off,Inches(0.35),10,DARK,accent)
        tb5=s.shapes.add_textbox(l8+Inches(0.65),t8+y_off,cw8-Inches(0.85),Inches(0.82))
        tf5=tb5.text_frame; tf5.word_wrap=True; p5=tf5.paragraphs[0]
        r1=p5.add_run(); r1.text=label+': '
        r1.font.bold=True; r1.font.size=Pt(10); r1.font.color.rgb=GOLD; r1.font.name='Calibri'
        r2=p5.add_run(); r2.text=text
        r2.font.size=Pt(10); r2.font.color.rgb=BODY; r2.font.name='Calibri'
        y_off+=Inches(0.92)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 9: ПУТЬ КЛИЕНТА — АЛГОРИТМ ВИЗИТА
# ═══════════════════════════════════════════════════════════════════════════════
s = ns(); bg(s); logo(s)
slide_header(s,'Путь клиента: Алгоритм визита специалиста')
# 8 steps in 2 rows of 4, connected by arrows
steps9 = [
    ('⚕','СБОР ЖАЛОБ','Что беспокоит, динамика.',['Боль','Динамика'],TERRA),
    ('☰','АНАМНЕЗ','Текущие препараты, заболевания.',['Лекарства','Хр. заболевания'],MED),
    ('⊕','СКРИНИНГ\nПРОТИВОПОКАЗАНИЙ','Проверка абсолютных стопов.',['Противопоказания'],TERRA),
    ('⚠','СКРИНИНГ\n«КРАСНЫХ ФЛАГОВ»','При наличии → направление к врачу.',['Красные флаги'],RED_W),
    ('◉','ФОРМИРОВАНИЕ\nОЖИДАНИЙ','Никаких 100% гарантий.',['Реалистичные цели'],MED),
    ('⬡','ПОДБОР\nКОМПЛЕКСА','По матрице приоритетов.',['Инд. план'],TERRA),
    ('☑','СОГЛАСОВАНИЕ\nСХЕМЫ','Строго по инструкции.',['Инструкция'],MED),
    ('▲','КОНТРОЛЬНЫЕ\nТОЧКИ','Фиксация динамики.',['Мониторинг'],GOLD),
]
sw9=Inches(2.72); sh9=Inches(2.25); gap9=Inches(0.12)
row1_y=Inches(1.5); row2_y=Inches(4.05)
for i,(icon,title,desc,tags,accent) in enumerate(steps9):
    col=i%4; row=i//4
    l9=M+col*(sw9+gap9); t9=row1_y if row==0 else row2_y
    fc9=RGBColor(0x48,0x28,0x10) if accent==RED_W else CARD
    lc9=RED_W if accent==RED_W else accent
    box(s,l9,t9,sw9,sh9,fc9,rnd=True,lc=lc9,lw=1.5)
    icon_circle(s,icon,l9+Inches(0.15),t9+Inches(0.15),Inches(0.48),14,accent,WHITE)
    tx(s,title,l9+Inches(0.72),t9+Inches(0.18),sw9-Inches(0.87),Inches(0.65),
       sz=11.5,bold=True,c=CREAM)
    tx(s,desc,l9+Inches(0.15),t9+Inches(0.88),sw9-Inches(0.3),Inches(0.85),
       sz=10.5,c=BODY)
    # tags
    tx2=l9+Inches(0.15)
    for tag in tags:
        tw=Inches(len(tag)*0.09+0.3)
        box(s,tx2,t9+sh9-Inches(0.4),tw,Inches(0.28),CARD2,rnd=True)
        tx(s,tag,tx2+Inches(0.08),t9+sh9-Inches(0.37),tw-Inches(0.16),
           Inches(0.22),sz=8,c=BODY,al=PP_ALIGN.CENTER)
        tx2+=tw+Inches(0.08)
    # arrow to next in row
    if col<3:
        connector(s,l9+sw9,t9+sh9/2,l9+sw9+gap9,t9+sh9/2,accent,2)
# red detour arrow from step 4 down
box(s,M+3*(sw9+gap9)+Inches(0.15),Inches(3.75),sw9-Inches(0.3),Inches(0.28),RED_W,rnd=True)
tx(s,'Если ДА → Остановка. Направление к врачу.',
   M+3*(sw9+gap9)+Inches(0.2),Inches(3.78),sw9-Inches(0.4),Inches(0.25),
   sz=9,bold=True,c=WHITE,al=PP_ALIGN.CENTER)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 10: КРАСНЫЕ ФЛАГИ
# ═══════════════════════════════════════════════════════════════════════════════
s = ns(); bg(s,RGBColor(0x30,0x10,0x06)); logo(s)
box(s,0,0,W,Inches(0.08),RED_W)
box(s,0,0,Inches(0.3),H,RED_W)
tx(s,'Зона риска: Абсолютные «Красные флаги»',
   Inches(0.5),Inches(0.18),W-Inches(0.7),Inches(0.72),sz=30,bold=True,c=CREAM)
# warning banner
box(s,Inches(0.5),Inches(0.98),W-Inches(0.7),Inches(0.62),
    RGBColor(0x50,0x18,0x08),rnd=True,lc=RED_W,lw=1.5)
tx(s,'⚠  При наличии данных состояний самостоятельный подбор ЗАПРЕЩЁН. Обязательное направление к профильному врачу до старта курса.',
   Inches(0.7),Inches(1.05),W-Inches(1.1),Inches(0.52),sz=12,bold=True,c=CREAM)
# 7 flag cards (4+3)
flags10 = [
    ('⛉','Беременность\nи лактация.'),
    ('☣','Активная или\nподозреваемая онкология.'),
    ('⚡','Тяжёлые психиатрические\nсостояния (психозы,\nнестабильное БАР, тяжёлая депрессия).'),
    ('🌡','Острые инфекции,\nлихорадка, аутоиммунные\nобострения.'),
    ('⚠','Эндокринная патология\nнеясного генеза.'),
    ('♡','Необъяснимая потеря\nвеса, кровотечения,\nнарастающая боль.'),
    ('⏲','Возраст до 18 лет.'),
]
fw10=(W-Inches(1.1))/4; fh10=Inches(1.95); ft10=Inches(1.72)
for i,(ico,text) in enumerate(flags10):
    col=i%4; row=i//4
    l10=Inches(0.5)+col*(fw10+Inches(0.08))
    t10=ft10+row*(fh10+Inches(0.12))
    if row==1 and col==3: break
    if row==1:
        fw10b=(W-Inches(1.1))/3
        l10=Inches(0.5)+col*(fw10b+Inches(0.1))
    box(s,l10,t10,fw10 if row==0 else fw10b,fh10,
        RGBColor(0x40,0x14,0x08),rnd=True,lc=RED_W,lw=1.5)
    tx(s,ico,l10+Inches(0.15),t10+Inches(0.12),Inches(0.42),Inches(0.42),
       sz=20,c=RED_W,al=PP_ALIGN.CENTER)
    tx(s,text,l10+Inches(0.15),t10+Inches(0.58),
       (fw10 if row==0 else fw10b)-Inches(0.3),fh10-Inches(0.7),
       sz=11.5,c=BODY)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 11: АРХИТЕКТУРА КУРСА — ПИРАМИДА + ПРАВИЛА
# ═══════════════════════════════════════════════════════════════════════════════
s = ns(); bg(s); logo(s)
slide_header(s,'Архитектура курса: Приоритеты и комбинирование')
# LEFT: pyramid (3 bars of decreasing width = pyramid)
py_l=M; py_t=Inches(1.55)
tx(s,'Пирамида приоритетов',py_l,py_t,Inches(5.5),Inches(0.45),
   sz=16,bold=True,c=CREAM)
levels = [
    (TERRA,   Inches(1.8),  'Уровень 3','Либидо, Репродуктивная ось, Эмоции','(TESTOBOOSTER, AMORE, OXYGEN)','Решаем после снятия истощения'),
    (MED,     Inches(3.5),  'Уровень 2','Когнитивная нагрузка, Выгорание','(ACTIVEBRAIN, STRESSRELIEF)',''),
    (RGBColor(0x5A,0x7A,0x4A), Inches(5.2), 'Уровень 1','Сон, Боль, Острый стресс, ОРВИ','(RECOVERY, RELIEF, NO STRESS, IMMUNACTIV)','Сначала стабилизируем базу'),
]
bar_h=Inches(1.0); max_w=Inches(5.3); py_content_t=Inches(2.08)
for i,(col,_,lvl_name,desc,comps,note) in enumerate(levels):
    bar_w=max_w*(0.45+0.27*(2-i))
    bar_l=M+(max_w-bar_w)/2
    t_bar=py_content_t+i*bar_h
    box(s,bar_l,t_bar,bar_w,bar_h-Inches(0.08),col,rnd=True)
    tx(s,lvl_name,bar_l+Inches(0.18),t_bar+Inches(0.08),
       bar_w-Inches(0.36),Inches(0.4),sz=12,bold=True,c=WHITE)
    tx(s,desc,bar_l+Inches(0.18),t_bar+Inches(0.48),
       bar_w-Inches(0.36),Inches(0.25),sz=9.5,c=WHITE)
    tx(s,comps,bar_l+Inches(0.18),t_bar+Inches(0.72),
       bar_w-Inches(0.36),Inches(0.2),sz=8.5,c=RGBColor(0xFF,0xEE,0xD8))
    if note:
        box(s,bar_l+bar_w+Inches(0.12),t_bar+bar_h/3,
            Inches(2.0),Inches(0.38),CARD2,rnd=True)
        tx(s,note,bar_l+bar_w+Inches(0.2),t_bar+bar_h/3+Inches(0.05),
           Inches(1.8),Inches(0.32),sz=9.5,c=BODY)
# RIGHT: golden rules
rx=M+Inches(5.7); rw=W-M-rx
tx(s,'Золотые правила',rx,Inches(1.55),rw,Inches(0.45),
   sz=16,bold=True,c=CREAM)
# Rule 1
box(s,rx,Inches(2.08),rw,Inches(0.72),CARD2,rnd=True)
tx(s,'Правило №1: Старт всегда с одного комплекса.',
   rx+Inches(0.2),Inches(2.16),rw-Inches(0.4),Inches(0.55),
   sz=13,bold=True,c=CREAM)
# CAN combine
box(s,rx,Inches(2.92),rw,Inches(1.42),CARD,rnd=True,lc=GRN_W,lw=1.5)
icon_circle(s,'✓',rx+Inches(0.18),Inches(3.0),Inches(0.42),14,GRN_W,WHITE)
tx(s,'МОЖНО комбинировать:',rx+Inches(0.7),Inches(3.02),rw-Inches(0.85),
   Inches(0.38),sz=12,bold=True,c=GRN_W)
tx(s,'Если запросы разные, но связаны (напр. недосып + выгорание), и нет пересекающихся противопоказаний.',
   rx+Inches(0.7),Inches(3.42),rw-Inches(0.85),Inches(0.82),sz=11,c=BODY)
# CANNOT combine
box(s,rx,Inches(4.46),rw,Inches(1.42),CARD,rnd=True,lc=RED_W,lw=1.5)
icon_circle(s,'✕',rx+Inches(0.18),Inches(4.54),Inches(0.42),14,RED_W,WHITE)
tx(s,'НЕЛЬЗЯ комбинировать:',rx+Inches(0.7),Inches(4.56),rw-Inches(0.85),
   Inches(0.38),sz=12,bold=True,c=RED_W)
tx(s,'Стимулирующие и седативные профили одновременно без чёткой цели; назначать «всё сразу».',
   rx+Inches(0.7),Inches(4.96),rw-Inches(0.85),Inches(0.82),sz=11,c=BODY)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 12: КОНТРОЛЬНАЯ ТОЧКА — ДЕРЕВО РЕШЕНИЙ
# ═══════════════════════════════════════════════════════════════════════════════
s = ns(); bg(s); logo(s)
slide_header(s,'Оценка результатов: Контрольная точка (2–4 недели)')
# Root node
rw12=Inches(4.0); rh12=Inches(0.85)
rl12=(W-rw12)/2; rt12=Inches(1.52)
box(s,rl12,rt12,rw12,rh12,CARD2,rnd=True,lc=GOLD,lw=2)
icon_circle(s,'☑',rl12+Inches(0.18),rt12+Inches(0.2),Inches(0.45),14,GOLD,WHITE)
tx(s,'Оценка самочувствия\nи соблюдения схемы.',rl12+Inches(0.73),rt12+Inches(0.08),
   rw12-Inches(0.88),rh12-Inches(0.15),sz=12,bold=True,c=CREAM)
# 3 branches
branch_y=Inches(2.62); bh12=Inches(1.3); bw12=Inches(3.5); gap12=Inches(0.25)
bl12=M+Inches(0.3)
branches=[
    ('✓','Положительная\nдинамика',GRN_W,
     'Продолжить курс по инструкции. Фиксация результата.'),
    ('⏸','Отсутствие\nэффекта',GOLD,
     'Проверить дозу, регулярность, приоритет жалобы.'),
    ('⚠','Нежелательная\nреакция\n(Ухудшение)',RED_W,
     'Немедленная остановка курса. Очная консультация врача. Фиксация симптомов.'),
]
bx12_positions=[M+Inches(0.3), M+Inches(0.3)+bw12+gap12, M+Inches(0.3)+2*(bw12+gap12)]
for i,(ico,title,col,action) in enumerate(branches):
    bx12=bx12_positions[i]
    # connector from root
    connector(s,rl12+rw12/2,rt12+rh12,bx12+bw12/2,branch_y,col,1)
    box(s,bx12,branch_y,bw12,bh12,CARD,rnd=True,lc=col,lw=2)
    icon_circle(s,ico,bx12+Inches(0.18),branch_y+Inches(0.18),Inches(0.45),14,col,WHITE)
    tx(s,title,bx12+Inches(0.73),branch_y+Inches(0.15),bw12-Inches(0.88),
       Inches(0.75),sz=12,bold=True,c=col)
    # action box
    ab_y=branch_y+bh12+Inches(0.15)
    box(s,bx12,ab_y,bw12,Inches(1.6),CARD2,rnd=True,lc=col,lw=1)
    tx(s,'Действие:',bx12+Inches(0.18),ab_y+Inches(0.1),bw12-Inches(0.36),
       Inches(0.35),sz=11,bold=True,c=col)
    tx(s,action,bx12+Inches(0.18),ab_y+Inches(0.45),bw12-Inches(0.36),
       Inches(1.05),sz=11,c=BODY)
    connector(s,bx12+bw12/2,branch_y,bx12+bw12/2,ab_y,col,1)
# Note for branch 2
note_y=branch_y+bh12+Inches(1.9)
box(s,bx12_positions[1],note_y,bw12,Inches(1.05),RGBColor(0x48,0x28,0x08),rnd=True,lc=GOLD,lw=1)
tx(s,'НЕ увеличивать дозу!\nНЕ добавлять второй комплекс!\nПри соблюдении схемы — к врачу.',
   bx12_positions[1]+Inches(0.18),note_y+Inches(0.1),bw12-Inches(0.36),
   Inches(0.9),sz=10.5,bold=True,c=GOLD)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 13: АНАТОМИЯ ВОЗРАЖЕНИЙ
# ═══════════════════════════════════════════════════════════════════════════════
s = ns(); bg(s); logo(s)
slide_header(s,'Анатомия возражений: Аргументация специалиста')
objs13 = [
    ('?','«Это гормоны?»',
     'Нет. Это короткие пептиды. Они не подменяют систему, а просят ваше тело самостоятельно выработать нужные вещества в физиологичной норме.',
     'Самостоятельная выработка', TERRA),
    ('?','«Это просто очередной БАД?»',
     'Это точечный регулятор. Пептиды работают по принципу «ключ-замок» строго со своими рецепторами, а не просто «укрепляют всё понемногу».',
     'Принцип ключ-замок', MED),
    ('⛓','«Вызовет ли это привыкание?»',
     'Не вызывает зависимости и синдрома отмены. После курса организм продолжает использовать восстановленные ресурсы самостоятельно.',
     'Курс → самостоятельная работа', GOLD),
    ('?','«А вдруг не поможет?»',
     'Мы не обещаем чудес и 100% гарантий. Эффект развивается постепенно. Для этого мы ставим контрольные точки и сверяем динамику.',
     'Контрольная точка', TERRA),
]
ow13=(W-2*M-Inches(0.4))/2; oh13=Inches(2.52)
for i,(ico,obj,ans,tag,accent) in enumerate(objs13):
    row=i//2; col=i%2
    l13=M+col*(ow13+Inches(0.4)); t13=Inches(1.52)+row*(oh13+Inches(0.2))
    box(s,l13,t13,ow13,oh13,CARD,rnd=True,lc=accent,lw=1.5)
    # left colored strip
    box(s,l13,t13,Inches(0.12),oh13,accent,rnd=True)
    # icon + title
    icon_circle(s,ico,l13+Inches(0.22),t13+Inches(0.18),Inches(0.42),13,accent,WHITE)
    tx(s,'Возражение:',l13+Inches(0.74),t13+Inches(0.12),
       ow13-Inches(0.9),Inches(0.3),sz=10,bold=False,c=accent)
    tx(s,obj,l13+Inches(0.74),t13+Inches(0.38),ow13-Inches(0.9),
       Inches(0.45),sz=13,bold=True,c=CREAM)
    box(s,l13+Inches(0.22),t13+Inches(0.9),ow13-Inches(0.37),Pt(1),accent)
    tx(s,'Ответ:',l13+Inches(0.22),t13+Inches(0.98),
       ow13-Inches(0.37),Inches(0.28),sz=10,bold=True,c=accent)
    tx(s,ans,l13+Inches(0.22),t13+Inches(1.25),ow13-Inches(0.37),
       Inches(1.0),sz=11,c=BODY)
    # tag
    tw13=Inches(len(tag)*0.088+0.4)
    box(s,l13+ow13-tw13-Inches(0.18),t13+oh13-Inches(0.45),tw13,Inches(0.32),
        DARK,rnd=True)
    tx(s,tag,l13+ow13-tw13-Inches(0.13),t13+oh13-Inches(0.43),
       tw13-Inches(0.1),Inches(0.28),sz=9,c=BODY,al=PP_ALIGN.CENTER)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 14: ОШИБКИ И ПОЗИЦИОНИРОВАНИЕ
# ═══════════════════════════════════════════════════════════════════════════════
s = ns(); bg(s); logo(s)
slide_header(s,'Профиль специалиста: Ошибки и позиционирование')
hw14=(W-2*M-Inches(0.5))/2
# LEFT: Ошибки (ловушки)
tx(s,'Ошибки (Ловушки)',M,Inches(1.52),hw14,Inches(0.45),
   sz=16,bold=True,c=RED_W)
errs14=[
    'Обещать 100% результат и «исцеление» от хронических болезней.',
    'Назначать пептиды без сбора анамнеза и проверки противопоказаний.',
    'Отменять или заменять пептидами препараты, назначенные врачом.',
    'Назначать «все баночки сразу» при нескольких жалобах.',
    'Продолжать курс, если у клиента появилась нежелательная реакция.',
]
for i,err in enumerate(errs14):
    ey=Inches(2.05)+i*Inches(0.88)
    box(s,M,ey,hw14,Inches(0.78),CARD,rnd=True,lc=RED_W,lw=1)
    icon_circle(s,'✕',M+Inches(0.15),ey+Inches(0.18),Inches(0.38),12,RED_W,WHITE)
    tx(s,err,M+Inches(0.63),ey+Inches(0.12),hw14-Inches(0.78),Inches(0.62),
       sz=11.5,c=BODY)
# RIGHT: Верное позиционирование
rx14=M+hw14+Inches(0.5)
tx(s,'Верное позиционирование',rx14,Inches(1.52),hw14,Inches(0.45),
   sz=16,bold=True,c=GRN_W)
rights14=[
    'Использовать термины: «помогает поддержать», «направлен на регуляцию».',
    'Работать строго по приоритетам: 1 запрос = 1 комплекс на старте.',
    'Чётко проговаривать: пептиды — часть комплексного подхода, а не замена врача.',
    'При любых сомнениях или «красных флагах» — делегировать врачу.',
]
for i,right in enumerate(rights14):
    ry=Inches(2.05)+i*Inches(1.05)
    box(s,rx14,ry,hw14,Inches(0.92),CARD,rnd=True,lc=GRN_W,lw=1)
    icon_circle(s,'✓',rx14+Inches(0.15),ry+Inches(0.22),Inches(0.38),12,GRN_W,WHITE)
    tx(s,right,rx14+Inches(0.63),ry+Inches(0.15),hw14-Inches(0.78),Inches(0.72),
       sz=12,c=BODY)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 15: ЭТИЧЕСКИЙ КОДЕКС
# ═══════════════════════════════════════════════════════════════════════════════
s = ns(); bg(s); logo(s)
slide_header(s,'Этический кодекс и Медицинский дисклеймер')
ethics=[
    ('☰','Настоящий материал носит информационно-образовательный характер.'),
    ('⊕','Материал не заменяет очной консультации врача и не является инструкцией по самолечению.'),
    ('⚕','При наличии заболеваний, во время беременности и лактации, при онкологических, психиатрических и эндокринных состояниях применение возможно только после очной оценки врачом.'),
    ('☑','Конкретная схема применения определяется инструкцией производителя и/или назначением врача.'),
    ('◎','Мы не даём гарантий полного излечения или замены стандартной медицинской терапии.'),
]
ew15=W-2*M; eh15_base=Inches(0.82)
for i,(ico,text) in enumerate(ethics):
    ey15=Inches(1.52)+i*(eh15_base+Inches(0.12))
    is_long = len(text) > 120
    eh15=Inches(1.15) if is_long else eh15_base
    box(s,M,ey15,ew15,eh15,CARD,rnd=True,lc=TERRA,lw=1)
    icon_circle(s,ico,M+Inches(0.2),ey15+Inches(0.18),Inches(0.42),13,TERRA,WHITE)
    tx(s,text,M+Inches(0.75),ey15+Inches(0.12),ew15-Inches(0.9),
       eh15-Inches(0.22),sz=12,c=BODY)
# footer
box(s,M,Inches(6.82),W-2*M,Inches(0.52),RGBColor(0x48,0x24,0x08),rnd=True,lc=GOLD,lw=1.5)
tx(s,'Главный принцип: Безопасность клиента превыше всего.',
   M+Inches(0.3),Inches(6.9),W-2*M-Inches(0.6),Inches(0.38),
   sz=14,bold=True,c=GOLD,al=PP_ALIGN.CENTER)

# ═══════════════════════════════════════════════════════════════════════════════
# SAVE
# ═══════════════════════════════════════════════════════════════════════════════
OUT = '/Users/larisa.vlasova/Documents/GENELINE_OFFICE/projects/Пептидные_комплексы_дистрибьюторы_2026.pptx'
prs.save(OUT)
print('OK:', OUT)
