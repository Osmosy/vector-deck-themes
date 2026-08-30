# Тест универсальности: дека «Vector Music» (не legal), темы из engine.py
import re
OUT_FMT = os_fmt = '~/projects/vector-legal-decks15/vector-music-{theme}.pptx'
import os
OUT_FMT = os.path.expanduser('~/projects/vector-legal-decks15/vector-music-{theme}.pptx')

DATA = dict(
    kicker='MUSIC AI PLATFORM',
    title='Vector Music',
    subtitle='Локальная генерация музыки для Hermes Agent',
    chips=[('2', 'сервиса'), ('4', 'модели'), ('GPU', 'AMD')],
    github='github.com/Osmosy/vector-music',
    footer_tag='Osmosy · Hermes Agent · 2026',
    intro_head='Что такое Vector Music',
    intro_lead='Два сервиса: API на 8001 и Studio на 3001. Генерация треков '
               'по текстовому описанию, полностью локально.',
    intro_cards=[
        ('API (8001)', ['REST-эндпоинт генерации', 'Очередь задач', 'Webhook статусов']),
        ('Studio (3001)', ['Веб-интерфейс', 'Библиотека треков', 'Экспорт WAV/MP3']),
        ('Модели', ['4 локальные модели', 'Стиль/жанр/настроение', 'До 3 минут трека']),
    ],
    wf_steps=[('Запрос', 'Текстовое описание трека и параметры'),
              ('Генерация', 'Модель синтезирует аудио на GPU'),
              ('Библиотека', 'Трек попадает в Studio, доступен экспорт')],
    arch_items=[('Frontend', 'Studio UI — React, порт 3001'),
                ('Backend', 'API — FastAPI, порт 8001'),
                ('Inference', 'GPU-инференс через AMD ROCm')],
    fact_big='100%', fact_small='локальная генерация — без облака и подписок',
    final_msg='Музыка без облачных сервисов',
    final_sub='github.com/Osmosy/vector-music',
)

from engine import kicker, title_block, footer, bg_fill, card, chip, wide_panel, add_text, add_rect
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN


def sl_title(slide, th, D):
    if th.get('art'):
        eng_style = th['art']
        from engine import full_art, half_art
        if eng_style[0] == 'full':
            full_art(slide, th, eng_style[1], overlay_pct=26)
        else:
            half_art(slide, th, eng_style[1])
    ta = th['title_align']
    if ta == 'center' and not (th.get('art') and th['art'][0] == 'half'):
        add_rect(slide, (13.333-4.2)/2, 1.52, 4.2, 0.42, fill=th['surface'],
                 line=th['card_line'], radius=0.21, alpha=70)
        add_text(slide, (13.333-4.2)/2, 1.60, 4.2, 0.3, D['kicker'], size=10.5,
                 bold=True, font=th['f_mono'], color=th['text'],
                 align=PP_ALIGN.CENTER, spacing=180)
        add_text(slide, 1.67, 2.02, 10.0, 1.2, D['title'], size=64, bold=True,
                 font=th['f_title'], color=th['accent'], align=PP_ALIGN.CENTER)
        add_text(slide, 3.17, 3.30, 7.0, 0.4, D['subtitle'], size=17,
                 font=th['f_body'], color=th['text'], align=PP_ALIGN.CENTER)
        for i, (big, small) in enumerate(D['chips']):
            chip(slide, th, 3.47 + i*2.25, 4.12, 1.95, big, small)
        add_rect(slide, 4.87, 5.62, 3.6, 0.52, fill=th['accent'], radius=0.26)
        add_text(slide, 4.87, 5.74, 3.6, 0.3, D['github'], size=12.5, bold=True,
                 font=th['f_mono'], color='FFFFFF', align=PP_ALIGN.CENTER)
        add_text(slide, 4.87, 6.90, 3.6, 0.26, D['footer_tag'], size=9.5,
                 font=th['f_mono'], color=th['muted'], align=PP_ALIGN.CENTER)
    else:
        add_rect(slide, 0.62, 1.30, 0.05, 2.2, fill=th['accent'])
        add_text(slide, 0.92, 1.30, 6.5, 0.3, D['kicker'], size=11, bold=True,
                 font=th['f_mono'], color=th['kicker_accent'], spacing=200)
        add_text(slide, 0.88, 1.66, 6.2, 1.15, D['title'], size=58, bold=True,
                 font=th['f_title'], color=th['accent'])
        add_text(slide, 0.92, 2.86, 5.9, 0.7, D['subtitle'], size=16.5,
                 font=th['f_body'], color=th['text'])
        for i, (big, small) in enumerate(D['chips']):
            x = 0.92 + i*1.95
            add_text(slide, x, 3.85, 1.8, 0.5, big, size=26, bold=True,
                     font=th['f_title'], color=th['accent'])
            add_text(slide, x, 4.36, 1.8, 0.3, small, size=10, font=th['f_mono'],
                     color=th['muted'], spacing=100)
        add_text(slide, 0.92, 5.15, 6.0, 0.3, D['github'], size=13, bold=True,
                 font=th['f_mono'], color=th['accent'])
        add_text(slide, 0.92, 6.0, 6.0, 0.26, D['footer_tag'], size=9.5,
                 font=th['f_mono'], color=th['muted'])


def sl_intro(slide, th, D):
    bg_fill(slide, th)
    kicker(slide, th, '01 · Введение')
    title_block(slide, th, D['intro_head'])
    add_text(slide, 0.62, 1.75, 12.1, 0.5, D['intro_lead'], size=13.5,
             font=th['f_body'], color=th['muted'])
    cw, gap, m = 3.95, 0.25, 0.62
    for i, (head, lines) in enumerate(D['intro_cards']):
        card(slide, th, m + i*(cw+gap), 2.55, cw, 3.4, head, lines)


def sl_arch(slide, th, D):
    bg_fill(slide, th)
    kicker(slide, th, '02 · Архитектура')
    title_block(slide, th, 'Три слоя')
    y = 2.5
    for i, (head, lines) in enumerate(D['arch_items']):
        wide_panel(slide, th, 0.62, y + i*1.5, 12.1, 1.3, head, [lines])


def sl_final(slide, th, D):
    bg_fill(slide, th)
    add_text(slide, 1.67, 2.6, 10.0, 1.0, D['fact_big'], size=72, bold=True,
             font=th['f_title'], color=th['accent'], align=PP_ALIGN.CENTER)
    add_text(slide, 2.67, 3.9, 8.0, 0.5, D['fact_small'], size=16,
             font=th['f_body'], color=th['text'], align=PP_ALIGN.CENTER)
    add_text(slide, 2.67, 5.1, 8.0, 0.6, D['final_msg'], size=22, bold=True,
             font=th['f_title'], color=th['text'], align=PP_ALIGN.CENTER)
    add_text(slide, 2.67, 5.8, 8.0, 0.3, D['final_sub'], size=12,
             font=th['f_mono'], color=th['accent'], align=PP_ALIGN.CENTER)
    footer(slide, th, 4)


SLIDES = [sl_title, sl_intro, sl_arch, sl_final]
