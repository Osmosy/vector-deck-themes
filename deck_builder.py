#!/usr/bin/env python3
"""Универсальный дек-билдер: контент в DATA, дизайн — темы из engine.py.
Использование: ~/.venvs/pptx/bin/python deck_builder.py <deck.py> [тема]
Пример: deck_builder.py deck-vector-music.py 07-corporate-navy"""
import sys, os, importlib.util

ENGINE = os.path.expanduser('~/projects/vector-legal-decks15/engine.py')
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
    outs = []
    for th in themes:
        prs = eng.Presentation()
        prs.slide_width, prs.slide_height = eng.SW, eng.SH
        blank = prs.slide_layouts[6]
        for fn in deck.SLIDES:          # список функций-слайдов из deck-файла
            fn(prs.slides.add_slide(blank), th, deck.DATA)
        out = deck.OUT_FMT.format(theme=th['name'])
        prs.save(out)
        outs.append(out)
    return outs

if __name__ == '__main__':
    for p in build(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None):
        print('built', p)
