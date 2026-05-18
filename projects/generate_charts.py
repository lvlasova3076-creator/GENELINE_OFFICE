#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Специфические графики для слайдов: сон, волны боли, пирамида."""
from PIL import Image, ImageDraw, ImageFilter, ImageFont
import math, os

OUT = '/tmp/pptx_images'
os.makedirs(OUT, exist_ok=True)

# Geneline dark palette
BG    = (40,  18,  8)
CARD  = (58,  30, 12)
TERRA = (191,120, 80)
GOLD  = (200,155, 80)
MED   = (158,106, 58)
CREAM = (240,232,216)
BODY  = (210,195,175)
RED   = (192, 56, 24)
GREEN = (90, 138, 64)
WHITE = (255,255,255)

W, H = 1200, 600

def line_on(draw, pts, col, w=3):
    for i in range(len(pts)-1):
        draw.line([pts[i], pts[i+1]], fill=col, width=w)

# ── SLEEP ARCHITECTURE CHART ──────────────────────────────────────────────────
def make_sleep():
    img = Image.new('RGB', (W, H), BG)
    draw = ImageDraw.Draw(img)
    # grid
    PAD_L, PAD_R, PAD_T, PAD_B = 120, 40, 60, 80
    cw = W - PAD_L - PAD_R
    ch = H - PAD_T - PAD_B
    # Y levels: Awake=0, Light=1, Deep=2, REM=3
    levels = {'Awake':0, 'Light':1, 'Deep':2, 'REM':3}
    n_levels = 4
    def gy(lv): return PAD_T + int(ch * lv / (n_levels-1))
    # horizontal grid lines
    for lv, name in enumerate(['Awake','Light','Deep','REM']):
        y = gy(lv)
        draw.line([(PAD_L, y),(W-PAD_R, y)], fill=CARD, width=1)
        draw.text((PAD_L-110, y-10), name, fill=BODY)
    # time axis
    times = ['22:00','0:00','2:00','4:00','6:00','8:00']
    for i, t in enumerate(times):
        x = PAD_L + int(cw * i / (len(times)-1))
        draw.line([(x, PAD_T),(x, H-PAD_B)], fill=CARD, width=1)
        draw.text((x-20, H-PAD_B+8), t, fill=BODY)
    # sleep curve — step function
    # sequence: Awake → Light → Deep → REM → Light → Deep → REM → Light → Awake
    raw = [
        (0.00, 0),(0.05, 1),(0.12, 2),(0.22, 3),(0.32, 2),
        (0.38, 2),(0.43, 3),(0.53, 2),(0.58, 1),(0.63, 2),
        (0.70, 3),(0.78, 2),(0.83, 1),(0.88, 0),(1.00, 0)
    ]
    pts = []
    for t, lv in raw:
        x = PAD_L + int(cw * t)
        y = gy(lv)
        pts.append((x, y))
    # draw step line
    step_pts = []
    for i in range(len(pts)-1):
        x1, y1 = pts[i]
        x2, y2 = pts[i+1]
        step_pts.append((x1, y1))
        step_pts.append((x2, y1))
    step_pts.append(pts[-1])
    # fill area under curve
    fill_pts = [(PAD_L, H-PAD_B)] + step_pts + [(W-PAD_R, H-PAD_B)]
    draw.polygon(fill_pts, fill=(80, 45, 20))
    # draw curve on top
    line_on(draw, step_pts, TERRA, 4)
    # highlight deep sleep phases
    deep_ranges = [(0.12,0.22),(0.38,0.43),(0.63,0.70)]
    for s, e in deep_ranges:
        x1 = PAD_L + int(cw*s); x2 = PAD_L + int(cw*e)
        y_deep = gy(2)
        draw.rectangle([x1, y_deep, x2, H-PAD_B], fill=(100, 55, 22))
    # re-draw curve
    line_on(draw, step_pts, TERRA, 4)
    img = img.filter(ImageFilter.GaussianBlur(0.5))
    img.save(f'{OUT}/sleep_chart.png')
    print('  sleep_chart done')

# ── PAIN WAVE CHART ───────────────────────────────────────────────────────────
def make_pain_wave():
    img = Image.new('RGB', (W, H), BG)
    draw = ImageDraw.Draw(img)
    PAD_L, PAD_R, PAD_T, PAD_B = 60, 60, 80, 80
    cw = W - PAD_L - PAD_R
    ch = H - PAD_T - PAD_B
    mid_y = PAD_T + ch // 2
    # before: high amplitude chaotic wave
    before_pts = []
    n = 300
    for i in range(n):
        t = i / n
        x = PAD_L + int(cw * t * 0.48)
        amp = ch * 0.38 * (1 - t*0.3)
        freq = 12 + t * 5
        y = mid_y + int(amp * math.sin(freq * math.pi * t + 0.5*math.sin(t*20)))
        before_pts.append((x, y))
    line_on(draw, before_pts, TERRA, 3)
    # divider
    div_x = PAD_L + int(cw * 0.5)
    draw.line([(div_x, PAD_T),(div_x, H-PAD_B)], fill=GOLD, width=2)
    draw.text((div_x-80, PAD_T-30), 'HYPERSENSITIVITY', fill=TERRA)
    # after: low amplitude calm wave
    after_pts = []
    for i in range(n):
        t = i / n
        x = div_x + 20 + int(cw * t * 0.48)
        amp = ch * 0.08 + ch * 0.04 * math.exp(-t*3)
        freq = 8
        y = mid_y + int(amp * math.sin(freq * math.pi * t))
        after_pts.append((x, y))
    line_on(draw, after_pts, MED, 3)
    draw.text((div_x+60, PAD_T-30), 'CALM / REDUCED REACTIVITY', fill=MED)
    # center line
    draw.line([(PAD_L, mid_y),(W-PAD_R, mid_y)], fill=CARD, width=1)
    img.save(f'{OUT}/pain_wave.png')
    print('  pain_wave done')

# ── MOLECULE ABSTRACT (for slide 2) ───────────────────────────────────────────
def make_molecule():
    img = Image.new('RGB', (600, 500), BG)
    draw = ImageDraw.Draw(img)
    cx, cy = 300, 250
    # chain of amino acids
    chain_pts = []
    for i in range(12):
        angle = math.pi * 0.15 * i - math.pi*0.3
        x = cx - 200 + i * 38 + int(15*math.sin(i*0.8))
        y = cy + int(40*math.sin(i*0.6))
        chain_pts.append((x, y))
        # draw bead
        r = 14 if i % 3 == 0 else 10
        col = TERRA if i % 3 == 0 else MED
        draw.ellipse([x-r,y-r,x+r,y+r], fill=col)
    # connect beads
    for i in range(len(chain_pts)-1):
        draw.line([chain_pts[i], chain_pts[i+1]], fill=CARD, width=3)
    # re-draw beads on top
    for i, (x, y) in enumerate(chain_pts):
        r = 14 if i % 3 == 0 else 10
        col = TERRA if i % 3 == 0 else MED
        draw.ellipse([x-r,y-r,x+r,y+r], fill=col)
    img = img.filter(ImageFilter.GaussianBlur(1))
    img.save(f'{OUT}/molecule.png')
    print('  molecule done')

# ── KEY-LOCK ABSTRACT ─────────────────────────────────────────────────────────
def make_key_lock():
    img = Image.new('RGB', (600, 500), BG)
    draw = ImageDraw.Draw(img)
    # key body
    draw.ellipse([80,180,200,300], outline=GOLD, width=5)
    draw.rectangle([160,225,380,255], fill=GOLD)
    draw.rectangle([260,255,290,300], fill=GOLD)
    draw.rectangle([320,255,350,285], fill=GOLD)
    # receptor (striped surface)
    for i in range(10):
        y = 120 + i * 28
        col = TERRA if i % 2 == 0 else MED
        draw.rectangle([420, y, 540, y+24], fill=col)
    # lock hole
    draw.ellipse([455,215,495,255], fill=BG, outline=CREAM, width=2)
    # connection glow
    for r in range(30, 5, -4):
        c = min(255, 60 + (30-r)*6)
        draw.ellipse([390-r,230-r,430+r,270+r], fill=(c//2,c//3,c//6))
    img = img.filter(ImageFilter.GaussianBlur(1.5))
    img.save(f'{OUT}/key_lock.png')
    print('  key_lock done')

if __name__ == '__main__':
    print('Generating charts...')
    make_sleep()
    make_pain_wave()
    make_molecule()
    make_key_lock()
    print(f'Done → {OUT}')
