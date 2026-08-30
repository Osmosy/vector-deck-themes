#!/bin/bash
# Render all 15 decks + build contact sheets
set -e
OUTDIR=~/projects/vector-legal-decks15
RENDER=$OUTDIR/render
mkdir -p "$RENDER"
PY=~/.venvs/pptx/bin/python

for f in "$OUTDIR"/vector-legal-*.pptx; do
  name=$(basename "$f" .pptx)
  if [ ! -d "$RENDER/$name" ]; then
    $PY /home/lenovo/.hermes/skills/productivity/powerpoint/scripts/pptx_render.py "$f" --outdir "$RENDER/$name" >/dev/null 2>&1 && echo "rendered $name"
  fi
done

# contact sheet: one row per theme (12 thumbs), 4 sheets of themes for review + full grid
$PY - <<'EOF'
from PIL import Image
import glob, os
OUTDIR = os.path.expanduser('~/projects/vector-legal-decks15')
RENDER = f'{OUTDIR}/render'
themes = sorted(os.path.basename(p) for p in glob.glob(f'{RENDER}/vector-legal-*'))
TH_W, TH_H = 480, 270  # thumb size
# sheet1: titles of all 15 (4 cols x 4 rows)
cols = 4
def make_sheet(paths, cols, cell_w, cell_h, out, labels=True):
    rows = (len(paths) + cols - 1) // cols
    from PIL import ImageDraw
    label_h = 28 if labels else 0
    sheet = Image.new('RGB', (cols * cell_w, rows * (cell_h + label_h)), (20, 20, 24))
    d = ImageDraw.Draw(sheet)
    for i, p in enumerate(paths):
        im = Image.open(p).resize((cell_w, cell_h))
        x = (i % cols) * cell_w
        y = (i // cols) * (cell_h + label_h)
        sheet.paste(im, (x, y + label_h))
        if labels:
            d.text((x + 6, y + 6), os.path.basename(os.path.dirname(p)), fill=(255, 220, 120))
    sheet.save(out)
    print('sheet', out)
# titles sheet
titles = [f'{RENDER}/{t}/slide-01.png' for t in themes]
make_sheet(titles, 4, TH_W, TH_H, f'{OUTDIR}/sheet_titles.png')
# content sample sheet: slides 2 and 7 of each theme
contents = []
for t in themes:
    contents += [f'{RENDER}/{t}/slide-02.png', f'{RENDER}/{t}/slide-07.png']
make_sheet(contents, 6, TH_W, TH_H, f'{OUTDIR}/sheet_content.png')
EOF
