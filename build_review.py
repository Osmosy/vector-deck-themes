#!/usr/bin/env python3
"""Build HTML review page for 15 deck variants."""
import glob, os, base64
OUTDIR = os.path.expanduser(
    os.environ.get('VECTOR_DECK_OUTDIR', '~/projects/vector-legal-decks15'))
RENDER = f'{OUTDIR}/render'
themes = sorted(os.path.basename(p) for p in glob.glob(f'{RENDER}/vector-legal-*'))

META = {
 '01-obsidian-neon': ('Тёмная', 'Текущий вайб v2: неон-синий на тёмно-синем, Inter Display'),
 '02-academic-light': ('Светлая', 'Мягкий свет, стеклянные весы, деликатный синий'),
 '03-editorial-cream': ('Светлая', 'Кремовая бумага, Playfair Display serif, шалфейный акцент'),
 '04-mono-swiss': ('Светлая', 'Швейцарский моно: белое/чёрное, красный акцент, нулевые радиусы'),
 '05-midnight-terminal': ('Тёмная', 'Терминальная эстетика: тёмный графит, тил-зелёный, mono'),
 '06-paper-minimal': ('Светлая', 'Тёплая бумага, serif-заголовки, терракотовый акцент'),
 '07-corporate-navy': ('Тёмная', 'Классика юрфирмы: полуночный синий + золото, Playfair'),
 '08-vanta-glass': ('Тёмная', 'Vantablack + изумрудное стекло (Ethereal Glass из high-end навыка)'),
 '09-light-blueprint': ('Светлая', 'Чертёжный светло-серый, жёсткая сетка, малые радиусы'),
 '10-violet-grad': ('Тёмная', 'Фиолетовый неон, Unbounded display-шрифт — самый смелый'),
 '11-sepia-legal': ('Светлая', 'Сепия «старого юридического архива», бордовый акцент'),
 '12-arctic-frost': ('Светлая', 'Арктический голубой, спокойный корпоративный'),
 '13-carbon-orange': ('Тёмная', 'Карбон + оранжевый неон, Unbounded, индустриальный'),
 '14-forest-institute': ('Тёмная', 'Тёмный изумруд + латунь, «институтский» Playfair'),
 '15-typographic-bw': ('Светлая', 'Чистая типографика: ч/б + один синий, без правил-линеек'),
}

def img64(p):
    with open(p, 'rb') as f:
        return 'data:image/png;base64,' + base64.b64encode(f.read()).decode()

html = ['''<!DOCTYPE html><html lang="ru"><head><meta charset="utf-8">
<title>Vector Legal — 15 тем</title>
<style>
body{background:#0d1117;color:#e6edf3;font-family:Inter,sans-serif;margin:0;padding:32px}
h1{font-size:26px} h2{font-size:15px;color:#8b949e;font-weight:400;margin-top:4px}
.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(430px,1fr));gap:28px}
.card{background:#161b22;border:1px solid #30363d;border-radius:12px;overflow:hidden}
.card h3{margin:0;padding:12px 16px 2px;font-size:16px}
.card p{margin:0;padding:0 16px 10px;color:#8b949e;font-size:12.5px}
.tabs{display:flex;gap:6px;padding:0 16px 12px;flex-wrap:wrap}
.tabs button{background:#21262d;color:#c9d1d9;border:1px solid #30363d;border-radius:6px;
 padding:3px 10px;font-size:11.5px;cursor:pointer}
.tabs button.on{background:#1f6feb;border-color:#1f6feb;color:#fff}
.imgs img{display:none;width:100%}
.imgs img.on{display:block}
.note{color:#8b949e;font-size:12.5px;margin:18px 0}
</style></head><body>
<h1>Vector Legal — 15 вариантов темы</h1>
<h2>Выбери номер — победитель оформим как навык. Кликай по кнопкам слайдов внутри карточки.</h2>
<div class="grid">''']

for t in themes:
    num, slug = t.split('-', 1)
    mode, desc = META.get(t, ('', ''))
    badge = '🌙' if mode == 'Тёмная' else '☀️'
    html.append(f'<div class="card" id="{t}"><h3>{badge} {num} · {slug}</h3><p>{desc}</p>')
    html.append('<div class="tabs">')
    slides = sorted(glob.glob(f'{RENDER}/{t}/slide-*.png'))
    for i, s in enumerate(slides):
        on = ' class="on"' if i == 0 else ''
        html.append(f'<button data-t="{t}" data-i="{i}"{on}>{i+1}</button>')
    html.append('</div><div class="imgs">')
    for i, s in enumerate(slides):
        on = ' class="on"' if i == 0 else ''
        html.append(f'<img src="{img64(s)}" data-t="{t}" data-i="{i}"{on}>')
    html.append('</div></div>')

html.append('''</div>
<p class="note">12 слайдов в каждой теме: титул, введение, 4 шага, домены 1/2, домены 2/2, архитектура,
агенты, MCP, коннекторы, руководство, roadmap, финал. Контент идентичен, отличается только дизайн-слой.</p>
<script>
document.querySelectorAll('.tabs button').forEach(b=>{
 b.onclick=()=>{
  const t=b.dataset.t, i=b.dataset.i;
  document.querySelectorAll(`.tabs button[data-t="${t}"]`).forEach(x=>x.classList.remove('on'));
  document.querySelectorAll(`.imgs img[data-t="${t}"]`).forEach(x=>x.classList.remove('on'));
  b.classList.add('on');
  document.querySelector(`.imgs img[data-t="${t}"][data-i="${i}"]`).classList.add('on');
 };
});
</script></body></html>''')

out = f'{OUTDIR}/review.html'
open(out, 'w').write('\n'.join(html))
print(out, os.path.getsize(out) // 1024, 'KB')
