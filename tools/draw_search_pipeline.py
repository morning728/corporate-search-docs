"""Схема поиска в редактируемом SVG и PNG для включения в Word."""

from html import escape
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parents[1]
WIDTH, HEIGHT = 1320, 470
BOX_WIDTH, BOX_HEIGHT = 280, 110
NODES = [
    (30, 30, ["Документы"]),
    (350, 30, ["Обработка", "документов"]),
    (670, 30, ["Обратный", "индекс"]),
    (30, 320, ["Запрос", "пользователя"]),
    (350, 320, ["Обработка", "запроса"]),
    (670, 320, ["Поиск", "по индексу"]),
    (990, 320, ["Ранжирование", "и результаты"]),
]
ARROWS = [
    (310, 85, 350, 85), (630, 85, 670, 85),
    (810, 140, 810, 320),
    (310, 375, 350, 375), (630, 375, 670, 375), (950, 375, 990, 375),
]


def font():
    for path in ("C:/Windows/Fonts/arial.ttf",
                 "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
                 "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf"):
        if Path(path).is_file():
            return ImageFont.truetype(path, 32)
    raise ValueError("Для схемы нужен шрифт Arial, DejaVu Sans или Liberation Sans с кириллицей.")


def generate():
    directory = ROOT / "images"
    directory.mkdir(exist_ok=True)
    picture = Image.new("RGB", (WIDTH, HEIGHT), "white")
    draw = ImageDraw.Draw(picture)
    label_font = font()
    svg = [f'<svg xmlns="http://www.w3.org/2000/svg" width="{WIDTH}" height="{HEIGHT}" viewBox="0 0 {WIDTH} {HEIGHT}">',
           '<title>Упрощённая схема индексирования и поиска</title>',
           '<rect width="100%" height="100%" fill="white"/>']
    for x1, y1, x2, y2 in ARROWS:
        points = [(x2, y2), (x2-12, y2-7), (x2-12, y2+7)] if y1 == y2 else [(x2, y2), (x2-7, y2-12), (x2+7, y2-12)]
        draw.line((x1, y1, x2, y2), fill="#555555", width=3)
        draw.polygon(points, fill="#555555")
        svg.append(f'<path d="M{x1},{y1} L{x2},{y2}" fill="none" stroke="#555" stroke-width="3"/>')
        svg.append('<polygon points="'+' '.join(f'{x},{y}' for x,y in points)+'" fill="#555"/>')
    for x, y, lines in NODES:
        draw.rectangle((x,y,x+BOX_WIDTH,y+BOX_HEIGHT),fill="white",outline="#555555",width=2)
        svg.append(f'<rect x="{x}" y="{y}" width="{BOX_WIDTH}" height="{BOX_HEIGHT}" fill="white" stroke="#555" stroke-width="2"/>')
        first_y=y+BOX_HEIGHT/2-(len(lines)-1)*21
        for i, line in enumerate(lines):
            cy=first_y+i*42
            draw.text((x+BOX_WIDTH/2,cy),line,font=label_font,fill="black",anchor="mm")
            svg.append(f'<text x="{x+BOX_WIDTH/2}" y="{cy}" text-anchor="middle" dominant-baseline="central" font-family="Arial, DejaVu Sans, sans-serif" font-size="32">{escape(line)}</text>')
    svg.append('</svg>')
    (directory/"search-pipeline.svg").write_text('\n'.join(svg)+'\n',encoding='utf-8')
    picture.save(directory/"search-pipeline.png",dpi=(240,240))


if __name__ == "__main__":
    generate()
    print("Созданы images/search-pipeline.svg и images/search-pipeline.png")
