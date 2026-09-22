"""Derive Windows ICO sizes from the PNG rendered directly from the SVG master."""
from pathlib import Path
from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parent.parent
assets = ROOT / 'assets'
sizes = [16, 20, 24, 32, 40, 48, 64, 96, 128, 256]
with Image.open(assets / 'OfficePDF-master.png') as original:
    original = original.convert('RGBA')
    original.resize((256, 256), Image.Resampling.LANCZOS).save(assets / 'OfficePDF.png')
    original.save(assets / 'OfficePDF.ico', format='ICO', sizes=[(s, s) for s in sizes])
    preview = Image.new('RGB', (720, 320), '#F3F5F8')
    draw = ImageDraw.Draw(preview)
    draw.rectangle((360, 0, 720, 320), fill='#202733')
    for background in [0, 360]:
        preview.paste(original.resize((176, 176), Image.Resampling.LANCZOS), (background+92, 20), original.resize((176, 176), Image.Resampling.LANCZOS))
        x = background + 30
        for size in [16, 20, 24, 32, 48]:
            icon = original.resize((size, size), Image.Resampling.LANCZOS)
            preview.paste(icon, (x, 244-size//2), icon)
            draw.text((x, 279), str(size), fill='#657083' if background == 0 else '#C9D5E4')
            x += size + 24
    preview.save(assets / 'icon-preview.png')
print('SVG-derived icon sizes:', sizes)
