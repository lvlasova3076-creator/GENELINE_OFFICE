#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Генерирует 10 тематических иллюстраций для пептидных слайдов."""
from PIL import Image, ImageDraw, ImageFilter
import os, math, random

OUT = '/tmp/pptx_images'
os.makedirs(OUT, exist_ok=True)
W, H = 1200, 900

def grad_v(img, top, bot):
    draw = ImageDraw.Draw(img)
    for y in range(H):
        t = y / H
        r = int(top[0] + (bot[0]-top[0])*t)
        g = int(top[1] + (bot[1]-top[1])*t)
        b = int(top[2] + (bot[2]-top[2])*t)
        draw.line([(0,y),(W,y)], fill=(r,g,b))

def grad_h(img, left, right):
    draw = ImageDraw.Draw(img)
    for x in range(W):
        t = x / W
        r = int(left[0] + (right[0]-left[0])*t)
        g = int(left[1] + (right[1]-left[1])*t)
        b = int(left[2] + (right[2]-left[2])*t)
        draw.line([(x,0),(x,H)], fill=(r,g,b))

def blur(img, r=8):
    return img.filter(ImageFilter.GaussianBlur(r))

# ── AMORE — лепестки/цветок, тёмно-розовый/терракот ─────────────────────────
def make_amore():
    img = Image.new('RGB', (W, H))
    grad_v(img, (55, 12, 18), (95, 35, 25))
    draw = ImageDraw.Draw(img)
    cx, cy = W//2 + 50, H//2
    random.seed(42)
    for layer in range(4):
        n_petals = 6 + layer * 2
        r_base = 80 + layer * 75
        for i in range(n_petals):
            angle = 2*math.pi * i / n_petals + layer * 0.3
            px = cx + r_base * math.cos(angle)
            py = cy + r_base * math.sin(angle) * 0.85
            pr = max(12, 60 - layer * 10)
            alpha = 200 - layer * 35
            col = (min(255,180+layer*15), min(255,55+layer*20), min(255,40+layer*12))
            draw.ellipse([px-pr, py-pr, px+pr, py+pr], fill=col)
    # inner glow
    for r in range(70, 5, -5):
        c = min(255, 220 + (70-r)*0)
        draw.ellipse([cx-r, cy-r, cx+r, cy+r],
                     fill=(min(255,235), min(255,130+r), min(255,80+r//2)))
    img = blur(img, 6)
    # fine detail layer
    draw2 = ImageDraw.Draw(img)
    for i in range(8):
        angle = 2*math.pi*i/8
        for rr in range(10, 280, 22):
            px = cx + rr * math.cos(angle)
            py = cy + rr * math.sin(angle) * 0.9
            draw2.ellipse([px-4,py-4,px+4,py+4], fill=(240,160,100,120))
    img = blur(img, 2)
    img.save(f'{OUT}/AMORE.png')
    print('  AMORE done')

# ── ACTIVEBRAIN — нейронная сеть, тёмно-синий/золотой ───────────────────────
def make_activebrain():
    img = Image.new('RGB', (W, H))
    grad_v(img, (10, 8, 35), (20, 15, 60))
    draw = ImageDraw.Draw(img)
    random.seed(7)
    # node points
    nodes = [(random.randint(80, W-80), random.randint(60, H-60)) for _ in range(55)]
    # connections
    for i, (x1,y1) in enumerate(nodes):
        for x2,y2 in nodes[i+1:]:
            dist = math.hypot(x2-x1, y2-y1)
            if dist < 220:
                alpha = max(20, int(200*(1-dist/220)))
                draw.line([(x1,y1),(x2,y2)], fill=(alpha, int(alpha*0.8), 20), width=1)
    # nodes
    for x,y in nodes:
        r = random.randint(4,11)
        draw.ellipse([x-r,y-r,x+r,y+r], fill=(200,165,60))
    # highlight nodes
    for x,y in random.sample(nodes, 12):
        draw.ellipse([x-14,y-14,x+14,y+14], fill=(240,200,80))
        draw.ellipse([x-7,y-7,x+7,y+7], fill=(255,240,180))
    img = blur(img, 3)
    img.save(f'{OUT}/ACTIVEBRAIN.png')
    print('  ACTIVEBRAIN done')

# ── IMMUNACTIV — сотовая/кристаллическая структура, зелёный/изумруд ─────────
def make_immunactiv():
    img = Image.new('RGB', (W, H))
    grad_v(img, (8, 28, 20), (15, 50, 35))
    draw = ImageDraw.Draw(img)
    # hexagonal grid
    hex_r = 55
    dx = hex_r * math.sqrt(3)
    dy = hex_r * 1.5
    for row in range(-1, int(H/dy)+2):
        for col in range(-1, int(W/dx)+2):
            cx = col * dx + (hex_r * math.sqrt(3)/2 if row%2 else 0)
            cy = row * dy
            pts = []
            for i in range(6):
                angle = math.pi/180 * (60*i - 30)
                pts.append((cx + hex_r*math.cos(angle)*0.88,
                             cy + hex_r*math.sin(angle)*0.88))
            draw.polygon(pts, outline=(30, 120, 70), fill=None)
    # bright hexes
    random.seed(13)
    centers = [(random.randint(100,W-100), random.randint(60,H-60)) for _ in range(8)]
    for cx,cy in centers:
        pts = []
        for i in range(6):
            angle = math.pi/180 * (60*i-30)
            pts.append((cx+hex_r*0.7*math.cos(angle), cy+hex_r*0.7*math.sin(angle)))
        draw.polygon(pts, fill=(25, 95, 55))
    img = blur(img, 4)
    img.save(f'{OUT}/IMMUNACTIV.png')
    print('  IMMUNACTIV done')

# ── NO STRESS — концентрические волны, глубокий синий ───────────────────────
def make_nostress():
    img = Image.new('RGB', (W, H))
    grad_v(img, (8, 12, 45), (5, 8, 30))
    draw = ImageDraw.Draw(img)
    cx, cy = W//2, H//2 + 40
    for r in range(20, 550, 28):
        alpha = max(15, 140 - r//4)
        col = (alpha//4, alpha//3, min(255, alpha+60))
        draw.ellipse([cx-r, cy-r*0.7, cx+r, cy+r*0.7], outline=col, width=2)
    # horizontal calm lines
    for y in range(0, H, 18):
        offset = 15 * math.sin(y * 0.04)
        draw.line([(0, y+offset),(W, y+offset)], fill=(20,25,80), width=1)
    img = blur(img, 5)
    # bright center point
    draw2 = ImageDraw.Draw(img)
    for r in range(35, 0, -3):
        c = 255 - r*4
        draw2.ellipse([cx-r,cy-r,cx+r,cy+r], fill=(max(0,c-60), max(0,c-20), min(255,c+40)))
    img = blur(img, 2)
    img.save(f'{OUT}/NOSTRESS.png')
    print('  NO STRESS done')

# ── OXYGEN — лучи/тепло/открытость, золото+персик ───────────────────────────
def make_oxygen():
    img = Image.new('RGB', (W, H))
    grad_v(img, (45, 22, 8), (80, 45, 15))
    draw = ImageDraw.Draw(img)
    cx, cy = W//2 - 30, H//2 + 20
    # rays
    for angle_deg in range(0, 360, 12):
        angle = math.radians(angle_deg)
        x2 = cx + 700 * math.cos(angle)
        y2 = cy + 700 * math.sin(angle)
        brightness = int(60 + 30*math.sin(math.radians(angle_deg*3)))
        draw.line([(cx,cy),(x2,y2)], fill=(brightness, brightness//2, 10), width=3)
    img = blur(img, 12)
    draw2 = ImageDraw.Draw(img)
    # concentric warm circles
    for r in range(300, 10, -25):
        c = min(255, 100 + r//3)
        draw2.ellipse([cx-r, cy-r*0.85, cx+r, cy+r*0.85],
                      outline=(min(255,c+50), min(255,c//2), 15), width=2)
    for r in range(90, 5, -8):
        draw2.ellipse([cx-r,cy-r,cx+r,cy+r],
                      fill=(min(255,200+r//2), min(255,140+r//3), min(255,60+r//4)))
    img = blur(img, 3)
    img.save(f'{OUT}/OXYGEN.png')
    print('  OXYGEN done')

# ── RECOVERY — рассвет/горизонт, оранжево-золотой ───────────────────────────
def make_recovery():
    img = Image.new('RGB', (W, H))
    grad_v(img, (10, 6, 18), (65, 25, 5))
    draw = ImageDraw.Draw(img)
    # horizon glow
    horizon = H * 2 // 3
    for y in range(H):
        if y > horizon - 80:
            t = min(1.0, (y - (horizon-80)) / 160)
            r = int(20 + 180*t)
            g = int(8 + 80*t)
            b = int(15 - 10*t)
            draw.line([(0,y),(W,y)], fill=(min(255,r),min(255,g),max(0,b)))
    img = blur(img, 10)
    draw2 = ImageDraw.Draw(img)
    # sun/sunrise
    cx = W//2
    cy = horizon
    for r in range(200, 8, -10):
        bright = min(255, 80 + (200-r))
        draw2.ellipse([cx-r, cy-r*0.6, cx+r, cy+r*0.6],
                      fill=(min(255,bright+50), min(255,bright//2+20), 5))
    # rising lines
    for i in range(12):
        angle = math.radians(-90 + (-30 + i*5))
        length = random.randint(150, 400)
        x2 = cx + length * math.cos(angle)
        y2 = cy + length * math.sin(angle)
        draw2.line([(cx,cy),(x2,y2)], fill=(240,180,30), width=max(1,3-i//5))
    img = blur(img, 3)
    img.save(f'{OUT}/RECOVERY.png')
    print('  RECOVERY done')

# ── RELIEF — разрыв/освобождение, тёмно-терракот → свет ────────────────────
def make_relief():
    img = Image.new('RGB', (W, H))
    grad_h(img, (30, 12, 8), (18, 8, 5))
    draw = ImageDraw.Draw(img)
    cx, cy = W//2, H//2
    # chains/rings — broken
    for i in range(7):
        y_pos = cy - 150 + i * 50
        r = 30
        # closed rings left
        for x_pos in range(80, cx-40, 70):
            draw.ellipse([x_pos-r, y_pos-r//2, x_pos+r, y_pos+r//2],
                         outline=(100, 50, 30), width=4)
        # open/breaking in center
        draw.arc([cx-40, y_pos-20, cx+40, y_pos+20], start=30, end=330,
                 fill=(200, 120, 70), width=5)
        # light bursting through break
        for j in range(5):
            angle = math.radians(-30 + j*15)
            lx = cx + 50 * math.cos(angle)
            ly = y_pos + 50 * math.sin(angle)
            draw.line([(cx,y_pos),(lx,ly)], fill=(240,195,120), width=2)
    img = blur(img, 5)
    draw2 = ImageDraw.Draw(img)
    # bright break center
    for r in range(60, 3, -5):
        c = min(255, 150 + (60-r)*2)
        draw2.ellipse([cx-r, cy-r//2, cx+r, cy+r//2],
                      fill=(min(255,c+30), min(255,c//2+40), min(255,c//4)))
    img = blur(img, 2)
    img.save(f'{OUT}/RELIEF.png')
    print('  RELIEF done')

# ── STOPBACTERIA — кристалл/броня, тёмно-фиолетовый/серебро ─────────────────
def make_stopbacteria():
    img = Image.new('RGB', (W, H))
    grad_v(img, (18, 8, 35), (8, 4, 22))
    draw = ImageDraw.Draw(img)
    cx, cy = W//2, H//2
    # diamond/crystal facets
    random.seed(99)
    for layer in range(5):
        n = 4 + layer * 2
        r1 = 60 + layer * 65
        r2 = r1 * 0.65
        pts_outer = []
        for i in range(n):
            angle = 2*math.pi*i/n + layer*0.2
            pts_outer.append((cx + r1*math.cos(angle), cy + r1*math.sin(angle)*0.85))
        pts_inner = []
        for i in range(n):
            angle = 2*math.pi*(i+0.5)/n + layer*0.2
            pts_inner.append((cx + r2*math.cos(angle), cy + r2*math.sin(angle)*0.85))
        # draw facet edges
        for i in range(n):
            draw.line([pts_outer[i], pts_inner[i]], fill=(90, 60, 140), width=2)
            draw.line([pts_outer[i], pts_outer[(i+1)%n]], fill=(70, 45, 120), width=1)
        # fill some facets
        for i in range(0, n, 2):
            tri = [pts_outer[i], pts_inner[i], pts_outer[(i+1)%n]]
            draw.polygon(tri, fill=(30, 18, 60))
    img = blur(img, 4)
    draw2 = ImageDraw.Draw(img)
    for r in range(50, 3, -4):
        c = min(255, 100 + (50-r)*3)
        draw2.ellipse([cx-r,cy-r,cx+r,cy+r], fill=(min(255,c+20),c//2,min(255,c+40)))
    img = blur(img, 2)
    img.save(f'{OUT}/STOPBACTERIA.png')
    print('  STOPBACTERIA done')

# ── STRESSRELIEF — нейронная тишина, глубокий индиго/синий ──────────────────
def make_stressrelief():
    img = Image.new('RGB', (W, H))
    grad_v(img, (12, 10, 50), (6, 5, 30))
    draw = ImageDraw.Draw(img)
    cx, cy = W//2, H//2
    # soft concentric rings
    for r in range(350, 15, -20):
        alpha = max(10, 100 - r//5)
        col = (alpha//6, alpha//4, min(255, alpha+100))
        draw.ellipse([cx-r, cy-r*0.75, cx+r, cy+r*0.75], outline=col, width=1)
    img = blur(img, 8)
    draw2 = ImageDraw.Draw(img)
    # scattered soft dots (neurons at rest)
    random.seed(55)
    for _ in range(40):
        x = random.randint(50, W-50)
        y = random.randint(40, H-40)
        r = random.randint(2, 8)
        c = random.randint(80, 160)
        draw2.ellipse([x-r,y-r,x+r,y+r], fill=(c//3, c//2, min(255,c+80)))
    for r in range(55, 5, -6):
        c = min(255, 80 + (55-r)*3)
        draw2.ellipse([cx-r,cy-r,cx+r,cy+r], fill=(c//5, c//4, min(255,c+80)))
    img = blur(img, 2)
    img.save(f'{OUT}/STRESSRELIEF.png')
    print('  STRESSRELIEF done')

# ── TESTOBOOSTER — пламя/притяжение, тёмный/золотой ─────────────────────────
def make_testobooster():
    img = Image.new('RGB', (W, H))
    grad_v(img, (8, 5, 5), (25, 10, 5))
    draw = ImageDraw.Draw(img)
    cx = W//2
    base_y = H - 60
    # flame shapes
    random.seed(17)
    for layer in range(3):
        for i in range(12):
            flame_cx = cx + random.randint(-80, 80)
            flame_base = base_y + random.randint(-20, 20)
            flame_h = random.randint(180, 420)
            flame_w = random.randint(35, 90)
            # draw flame as narrow ellipse
            col_r = min(255, 180 + layer*25)
            col_g = min(255, 80 + layer*30 + i*4)
            draw.ellipse([flame_cx - flame_w,
                           flame_base - flame_h,
                           flame_cx + flame_w,
                           flame_base],
                          fill=(col_r, col_g, 10))
    img = blur(img, 14)
    draw2 = ImageDraw.Draw(img)
    # bright core
    for r in range(120, 5, -8):
        c = min(255, 100 + (120-r)*1.2)
        draw2.ellipse([cx-r, base_y-r*2.5, cx+r, base_y],
                      fill=(min(255,int(c)+50), min(255,int(c)//2), 5))
    # sparks
    for _ in range(25):
        sx = cx + random.randint(-180, 180)
        sy = base_y - random.randint(80, 500)
        sr = random.randint(2,7)
        draw2.ellipse([sx-sr,sy-sr,sx+sr,sy+sr], fill=(255, 200, 60))
    img = blur(img, 2)
    img.save(f'{OUT}/TESTOBOOSTER.png')
    print('  TESTOBOOSTER done')

if __name__ == '__main__':
    print('Generating images...')
    make_amore()
    make_activebrain()
    make_immunactiv()
    make_nostress()
    make_oxygen()
    make_recovery()
    make_relief()
    make_stopbacteria()
    make_stressrelief()
    make_testobooster()
    print(f'All images saved to {OUT}')
