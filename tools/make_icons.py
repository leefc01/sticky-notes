"""Draws the ribbon icons at 320 px and downsamples to 16/32/80 px PNGs."""
import sys, os
from PIL import Image, ImageDraw
OUT = sys.argv[1]; os.makedirs(OUT, exist_ok=True)
S = 320; INK = (51, 51, 51, 255); NOTE = (255, 233, 122, 255); EDGE = (224, 176, 0, 255); RED = (198, 40, 40, 255)
W = 26  # stroke width at 320 px

def note(d, box=(40, 40, 280, 280)):
    x0, y0, x1, y1 = box; f = 70
    d.polygon([(x0, y0), (x1, y0), (x1, y1 - f), (x1 - f, y1), (x0, y1)], fill=NOTE, outline=EDGE, width=14)
    d.polygon([(x1, y1 - f), (x1 - f, y1 - f), (x1 - f, y1)], fill=EDGE)

def chevron(d, left):
    pts = [(200, 80), (120, 160), (200, 240)] if left else [(120, 80), (200, 160), (120, 240)]
    d.line(pts, fill=INK, width=W * 2, joint="curve")

def eye(d, slash):
    d.ellipse((40, 100, 280, 220), outline=INK, width=W)
    d.ellipse((120, 120, 200, 200), fill=INK)
    if slash: d.line((60, 270, 260, 50), fill=RED, width=W + 6)

def make(name, fn):
    im = Image.new("RGBA", (S, S), (0, 0, 0, 0)); d = ImageDraw.Draw(im); fn(d)
    for px in (16, 32, 80):
        im.resize((px, px), Image.LANCZOS).save(os.path.join(OUT, f"{name}-{px}.png"))

make("group", lambda d: note(d))
def add(d):
    note(d, (20, 20, 240, 240))
    d.ellipse((170, 170, 310, 310), fill=(46, 125, 50, 255))
    d.line((240, 195, 240, 285), fill="white", width=W); d.line((195, 240, 285, 240), fill="white", width=W)
make("add", add)
make("prev", lambda d: chevron(d, True))
make("next", lambda d: chevron(d, False))
def delete(d):
    note(d, (20, 20, 240, 240))
    d.ellipse((170, 170, 310, 310), fill=RED)
    d.line((205, 205, 275, 275), fill="white", width=W); d.line((275, 205, 205, 275), fill="white", width=W)
make("delete", delete)
def trash(d):
    d.rectangle((60, 60, 260, 90), fill=INK); d.rectangle((130, 30, 190, 60), outline=INK, width=W // 2)
    d.polygon([(80, 100), (240, 100), (220, 300), (100, 300)], outline=INK, width=W, fill=NOTE)
    for x in (130, 160, 190): d.line((x, 130, x, 270), fill=INK, width=W // 2)
make("deleteall", trash)
make("hide", lambda d: eye(d, True))
make("show", lambda d: eye(d, False))
def pane(d):
    d.rectangle((30, 50, 290, 270), outline=INK, width=W // 2 + 4)
    d.rectangle((190, 50, 290, 270), fill=NOTE, outline=INK, width=W // 2 + 4)
    for y in (100, 140, 180): d.line((210, y, 270, y), fill=INK, width=12)
make("pane", pane)
print(sorted(os.listdir(OUT)))
