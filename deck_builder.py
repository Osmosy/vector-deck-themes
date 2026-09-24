#!/usr/bin/env python3
"""Универсальный дек-билдер: контент в DATA, дизайн — темы из engine.py.
Использование: ~/.venvs/pptx/bin/python deck_builder.py <deck.py> [тема]
Пример: deck_builder.py deck-vector-music.py 07-corporate-navy

Колонтитулы «N / M · бренд»: M = len(SLIDES) деки, бренд — DATA['footer_brand'],
иначе «<DATA['title']> · Hermes Agent · Osmosy».
Движок — engine.py рядом с билдером; другой путь: VECTOR_DECK_ENGINE=<файл>.
VECTOR_DECK_OUTDIR=<каталог> — положить деку туда (имя файла — из OUT_FMT).
В OUT_FMT доступны {theme} (01-obsidian-neon) и {theme_short} (obsidian-neon)."""
import sys, os, importlib.util

ENGINE = os.environ.get('VECTOR_DECK_ENGINE',
                        os.path.join(os.path.dirname(os.path.abspath(__file__)), 'engine.py'))
spec = importlib.util.spec_from_file_location('engine', ENGINE)
eng = importlib.util.module_from_spec(spec)
sys.modules['engine'] = eng
# не запускать __main__ engine — он под importlib и так не выполнится
spec.loader.exec_module(eng)

def build(deck_path, theme_name=None):
    d = importlib.util.spec_from_file_location('deck', deck_path)
    deck = importlib.util.module_from_spec(d)
    sys.modules['deck'] = deck
    d.loader.exec_module(deck)
    themes = [t for t in eng.THEMES if not theme_name or t['name'] in theme_name]
    assert themes, f'тема {theme_name} не найдена'
    brand = deck.DATA.get('footer_brand') or f"{deck.DATA['title']} · Hermes Agent · Osmosy"
    eng.set_footer(len(deck.SLIDES), brand)
    outs = []
    for th in themes:
        prs = eng.Presentation()
        prs.slide_width, prs.slide_height = eng.SW, eng.SH
        blank = prs.slide_layouts[6]
        for fn in deck.SLIDES:          # список функций-слайдов из deck-файла
            fn(prs.slides.add_slide(blank), th, deck.DATA)
        # {theme} — «01-obsidian-neon», {theme_short} — «obsidian-neon»
        out = deck.OUT_FMT.format(theme=th['name'], theme_short=th['name'].split('-', 1)[-1])
        if os.environ.get('VECTOR_DECK_OUTDIR'):
            out = os.path.join(os.path.expanduser(os.environ['VECTOR_DECK_OUTDIR']),
                               os.path.basename(out))
        os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
        prs.save(out)
        outs.append(out)
    return outs

if __name__ == '__main__':
    for p in build(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None):
        print('built', p)
