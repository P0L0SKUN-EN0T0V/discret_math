#!/usr/bin/env python3
"""
Генератор статьи для журнала МГТУ им. Н.Э. Баумана (ptsj.bmstu.ru)
Оформление строго по требованиям: TNR 12pt, 1.5 интервал, поля, отступы.
"""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '.pylibs'))

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import re

# ── Константы ──
FONT = 'Times New Roman'
SIZE = Pt(12)
LINE_SPACING = 1.5
INDENT = Cm(0.7)
MARGIN_LEFT = Cm(2)
MARGIN_TOP = Cm(2)
MARGIN_RIGHT = Cm(1.5)
MARGIN_BOTTOM = Cm(1.5)

doc = Document()

# ── Настройка секции (поля, ориентация) ──
section = doc.sections[0]
section.orientation = WD_ORIENT.PORTRAIT
section.page_width = Cm(21)
section.page_height = Cm(29.7)
section.left_margin = MARGIN_LEFT
section.top_margin = MARGIN_TOP
section.right_margin = MARGIN_RIGHT
section.bottom_margin = MARGIN_BOTTOM

# ── Стиль Normal ──
style = doc.styles['Normal']
font = style.font
font.name = FONT
font.size = SIZE
font.color.rgb = RGBColor(0, 0, 0)
style.paragraph_format.line_spacing = LINE_SPACING
style.paragraph_format.space_after = Pt(0)
style.paragraph_format.space_before = Pt(0)
style.paragraph_format.first_line_indent = INDENT

# Для кириллицы — задаём шрифт через rFonts
rpr = style.element.get_or_add_rPr()
rfonts = OxmlElement('w:rFonts')
rfonts.set(qn('w:ascii'), FONT)
rfonts.set(qn('w:hAnsi'), FONT)
rfonts.set(qn('w:cs'), FONT)
rfonts.set(qn('w:eastAsia'), FONT)
rpr.append(rfonts)


def add_page_numbers(doc):
    """Нумерация страниц по центру с первой страницы."""
    section = doc.sections[0]
    footer = section.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    fld = OxmlElement('w:fldSimple')
    fld.set(qn('w:instr'), 'PAGE')
    run._element.append(fld)


def add_paragraph(text, bold=False, italic=False, alignment=None, indent=True,
                  font_size=None, space_after=None, space_before=None):
    """Добавить параграф с форматированием."""
    p = doc.add_paragraph()
    if not indent:
        p.paragraph_format.first_line_indent = Cm(0)
    if alignment:
        p.alignment = alignment
    if space_after is not None:
        p.paragraph_format.space_after = Pt(space_after)
    if space_before is not None:
        p.paragraph_format.space_before = Pt(space_before)

    run = p.add_run(text)
    run.font.name = FONT
    run.font.size = font_size or SIZE
    run.bold = bold
    run.italic = italic
    # Кириллица
    r = run._element
    rpr = r.get_or_add_rPr()
    rfonts = OxmlElement('w:rFonts')
    rfonts.set(qn('w:ascii'), FONT)
    rfonts.set(qn('w:hAnsi'), FONT)
    rfonts.set(qn('w:cs'), FONT)
    rfonts.set(qn('w:eastAsia'), FONT)
    rpr.insert(0, rfonts)
    return p


def add_rich_paragraph(parts, alignment=None, indent=True, space_after=None, space_before=None):
    """Параграф с чередованием bold/italic/normal частей.
    parts = [(text, bold, italic), ...]
    """
    p = doc.add_paragraph()
    if not indent:
        p.paragraph_format.first_line_indent = Cm(0)
    if alignment:
        p.alignment = alignment
    if space_after is not None:
        p.paragraph_format.space_after = Pt(space_after)
    if space_before is not None:
        p.paragraph_format.space_before = Pt(space_before)
    for text, bold, italic in parts:
        run = p.add_run(text)
        run.font.name = FONT
        run.font.size = SIZE
        run.bold = bold
        run.italic = italic
        r = run._element
        rpr = r.get_or_add_rPr()
        rfonts = OxmlElement('w:rFonts')
        rfonts.set(qn('w:ascii'), FONT)
        rfonts.set(qn('w:hAnsi'), FONT)
        rfonts.set(qn('w:cs'), FONT)
        rfonts.set(qn('w:eastAsia'), FONT)
        rpr.insert(0, rfonts)
    return p


def add_heading_styled(text, level=1):
    """Заголовок раздела — жирный, тот же шрифт."""
    p = add_paragraph(text, bold=True, indent=False,
                      space_before=12, space_after=6)
    return p


def add_formula(text, number=None):
    """Формула (текстовая заглушка — для MathType вставлять вручную)."""
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Cm(0)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.font.name = FONT
    run.font.size = SIZE
    run.italic = True
    if number:
        run2 = p.add_run(f'\t({number})')
        run2.font.name = FONT
        run2.font.size = SIZE
    return p


# ══════════════════════════════════════════════
#  ТЕКСТ СТАТЬИ
# ══════════════════════════════════════════════

# УДК
add_paragraph('УДК 004.414.23 + 681.5.01', bold=False, indent=False, space_after=6)

# Название
add_paragraph(
    'МОДЕЛИРОВАНИЕ ПАРАЛЛЕЛЬНЫХ ПРОЦЕССОВ\n'
    'РОБОТОТЕХНИЧЕСКОЙ СИСТЕМЫ С ПОМОЩЬЮ СЕТЕЙ ПЕТРИ',
    bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER, indent=False, space_after=6
)

# Автор
add_paragraph(
    'В.\u00a0П.\u00a0Романов',
    bold=False, alignment=WD_ALIGN_PARAGRAPH.CENTER, indent=False
)

# Организация
add_paragraph(
    'МГТУ им.\u00a0Н.\u00a0Э.\u00a0Баумана, Москва, Российская Федерация',
    italic=True, alignment=WD_ALIGN_PARAGRAPH.CENTER, indent=False,
    space_after=12
)

# ── Аннотация ──
add_rich_paragraph([
    ('Аннотация. ', True, False),
    ('Рассматривается задача формального моделирования высокоуровневого управления '
     'робототехнической системой с использованием аппарата сетей Петри. '
     'Обосновывается необходимость применения формальных методов для верификации '
     'параллельных процессов в робототехнике. Предлагается модель мобильного робота '
     'с двумя параллельными подсистемами — движения и сенсорного обеспечения, — '
     'взаимодействующими через общий ресурс. Проведён анализ модели методом '
     'S-инвариантов: доказано взаимное исключение при доступе к общему ресурсу '
     'и отсутствие тупиковых ситуаций. Выполнен темпоральный анализ '
     'с детерминированными задержками, определены пропускная способность системы '
     'и узкое место. Показано, что добавление параллельного ресурса повышает '
     'производительность на 60\u00a0%. Результаты могут применяться при проектировании '
     'систем управления автономными роботами.', False, False),
], indent=True, space_after=6)

# Ключевые слова
add_rich_paragraph([
    ('Ключевые слова: ', True, False),
    ('сети Петри, робототехника, параллельные процессы, формальная верификация, '
     'S-инварианты, взаимное исключение, темпоральный анализ, пропускная способность, '
     'высокоуровневое управление, тупиковые ситуации', False, False),
], indent=True, space_after=12)

# ══════════════ ВВЕДЕНИЕ ══════════════
add_heading_styled('1. Введение')

add_paragraph(
    'Современные робототехнические системы представляют собой сложные '
    'мехатронные комплексы, в которых одновременно функционируют множество '
    'подсистем: навигация, сенсорное обеспечение, манипуляция, связь с оператором. '
    'Параллельность этих процессов порождает класс ошибок, которые невозможно '
    'обнаружить методами последовательного тестирования: гонки данных, тупиковые '
    'блокировки (deadlock), нарушение взаимного исключения при доступе к общим '
    'ресурсам [1].'
)

add_paragraph(
    'Традиционный подход к описанию логики управления — конечные автоматы '
    '(FSM) — хорошо работает для последовательных систем. Однако при '
    'моделировании параллельных процессов число состояний автомата растёт '
    'экспоненциально: два независимых процесса с N и M состояниями порождают '
    'N\u00a0×\u00a0M комбинаций. Для промышленного робота с десятью подсистемами '
    'это делает модель практически необозримой [2].'
)

add_paragraph(
    'Сети Петри, предложенные К.\u00a0А.\u00a0Петри в 1962\u00a0г. [3], '
    'изначально были разработаны для моделирования параллельных и '
    'распределённых систем. В отличие от FSM, они описывают параллелизм '
    'явно, а не через произведение состояний, что делает модель компактной '
    'и анализируемой. Развитый математический аппарат (матричное представление, '
    'инварианты, дерево достижимости) позволяет формально доказывать '
    'свойства системы: отсутствие тупиков, ограниченность, живость [4].'
)

add_paragraph(
    'Целью настоящей работы является построение формальной модели '
    'высокоуровневого управления мобильным роботом с двумя параллельными '
    'подсистемами и общим ресурсом, а также проведение структурного '
    'и темпорального анализа полученной модели.'
)

# ══════════════ ТЕОРИЯ ══════════════
add_heading_styled('2. Теоретические основы')

add_heading_styled('2.1. Определение сети Петри')

add_paragraph(
    'Сеть Петри определяется как четвёрка [3, 4]:'
)

add_formula('PN = (P, T, F, M₀)', 1)

add_paragraph(
    'где P\u00a0=\u00a0{p₁,\u00a0p₂,\u00a0…,\u00a0pₘ} — конечное множество позиций '
    '(мест); T\u00a0=\u00a0{t₁,\u00a0t₂,\u00a0…,\u00a0tₙ} — конечное множество '
    'переходов; F\u00a0⊆\u00a0(P\u00a0×\u00a0T)\u00a0∪\u00a0(T\u00a0×\u00a0P) — '
    'множество дуг; M₀:\u00a0P\u00a0→\u00a0ℕ — начальная маркировка, задающая '
    'распределение токенов по позициям.'
)

add_heading_styled('2.2. Правило срабатывания')

add_paragraph(
    'Переход t ∈ T разрешён при маркировке M, если для каждой входной '
    'позиции p ∈ •t выполняется M(p)\u00a0≥\u00a0W(p,\u00a0t), где W — функция '
    'весов дуг. При срабатывании перехода маркировка изменяется [4]:'
)

add_formula('M\'(p) = M(p) − W(p, t) + W(t, p)', 2)

add_heading_styled('2.3. Матричное представление')

add_paragraph(
    'Для анализа удобно представить сеть Петри в матричной форме. '
    'Матрица инцидентности C размером |P|\u00a0×\u00a0|T| определяется как [4]:'
)

add_formula('C = C⁺ − C⁻', 3)

add_paragraph(
    'где C⁺(p,\u00a0t)\u00a0=\u00a0W(t,\u00a0p) — матрица выходных дуг, '
    'C⁻(p,\u00a0t)\u00a0=\u00a0W(p,\u00a0t) — матрица входных дуг. '
    'Уравнение состояния сети:'
)

add_formula('M = M₀ + C · σ', 4)

add_paragraph(
    'где σ — вектор срабатываний (Парикх-вектор), σ(t) — число '
    'срабатываний перехода t.'
)

add_heading_styled('2.4. S-инварианты')

add_paragraph(
    'Вектор x ∈ ℤᵐ называется S-инвариантом, если xᵀ\u00a0·\u00a0C\u00a0=\u00a00 [5]. '
    'Для любой достижимой маркировки M выполняется:'
)

add_formula('xᵀ · M = xᵀ · M₀ = const', 5)

add_paragraph(
    'S-инварианты позволяют формально доказывать свойства сети, '
    'не перечисляя все достижимые маркировки. В частности, если '
    'позиция общего ресурса входит в S-инвариант с компонентой 1, '
    'это доказывает ограниченность числа токенов — а значит, '
    'взаимное исключение [5].'
)

add_heading_styled('2.5. Темпоральные сети Петри')

add_paragraph(
    'Темпоральная сеть Петри расширяет классическую модель функцией '
    'задержки D:\u00a0T\u00a0→\u00a0ℝ⁺, которая каждому переходу ставит '
    'в соответствие время срабатывания [6]. При этом токены на входных '
    'позициях резервируются на время задержки и не могут быть использованы '
    'другими переходами. Это позволяет моделировать длительность операций '
    'и анализировать производительность системы.'
)

# ══════════════ МОДЕЛЬ ══════════════
add_heading_styled('3. Модель робототехнической системы')

add_heading_styled('3.1. Описание системы')

add_paragraph(
    'Рассматривается мобильный робот с двумя параллельными подсистемами:'
)

add_paragraph(
    '— подсистема движения (навигация, обход препятствий, '
    'возврат на базу);'
)
add_paragraph(
    '— подсистема сенсорного обеспечения (сканирование лидаром, '
    'обработка данных, построение карты).'
)

add_paragraph(
    'Обе подсистемы работают одновременно и независимо, за исключением '
    'одного момента: для обновления карты окружения необходим доступ '
    'к общему вычислительному ресурсу (GPU), который не может быть '
    'использован двумя процессами одновременно. Таким образом, '
    'возникает задача взаимного исключения [7].'
)

add_heading_styled('3.2. Структура сети Петри')

add_paragraph(
    'Модель содержит следующие позиции:'
)

add_paragraph(
    '— p₁ (Готовность_Движ) — подсистема движения готова к работе;\n'
    '— p₂ (Движение) — робот выполняет навигацию;\n'
    '— p₃ (Ожидание_GPU_Д) — движение ожидает доступ к GPU;\n'
    '— p₄ (Обработка_Д) — движение использует GPU;\n'
    '— p₅ (Готовность_Сенс) — подсистема сенсоров готова;\n'
    '— p₆ (Сканирование) — лидар выполняет сканирование;\n'
    '— p₇ (Ожидание_GPU_С) — сенсоры ожидают GPU;\n'
    '— p₈ (Обработка_С) — сенсоры используют GPU;\n'
    '— p₉ (GPU_свободен) — мьютекс, начальная маркировка M₀(p₉)\u00a0=\u00a01.'
)

add_paragraph(
    'Переходы модели:'
)

add_paragraph(
    '— t₁ (Старт_движения): p₁\u00a0→\u00a0p₂;\n'
    '— t₂ (Запрос_GPU_Д): p₂\u00a0→\u00a0p₃;\n'
    '— t₃ (Захват_GPU_Д): p₃\u00a0+\u00a0p₉\u00a0→\u00a0p₄;\n'
    '— t₄ (Освобождение_GPU_Д): p₄\u00a0→\u00a0p₁\u00a0+\u00a0p₉;\n'
    '— t₅ (Старт_сканирования): p₅\u00a0→\u00a0p₆;\n'
    '— t₆ (Запрос_GPU_С): p₆\u00a0→\u00a0p₇;\n'
    '— t₇ (Захват_GPU_С): p₇\u00a0+\u00a0p₉\u00a0→\u00a0p₈;\n'
    '— t₈ (Освобождение_GPU_С): p₈\u00a0→\u00a0p₅\u00a0+\u00a0p₉.'
)

add_paragraph(
    'Начальная маркировка: M₀\u00a0=\u00a0(1,\u00a00,\u00a00,\u00a00,\u00a0'
    '1,\u00a00,\u00a00,\u00a00,\u00a01). Один токен в p₁ (движение готово), '
    'один в p₅ (сенсоры готовы), один в p₉ (GPU свободен).'
)

# ══════════════ АНАЛИЗ ══════════════
add_heading_styled('4. Структурный анализ модели')

add_heading_styled('4.1. Матрица инцидентности')

add_paragraph(
    'Матрица инцидентности C размером 9\u00a0×\u00a08 построена '
    'по формуле (3). Каждый столбец соответствует переходу, '
    'каждая строка — позиции. Элемент C(pᵢ,\u00a0tⱼ) равен разности '
    'числа дуг из tⱼ в pᵢ и из pᵢ в tⱼ.'
)

# Здесь в реальной статье будет таблица — пока текстом
add_paragraph(
    'Для подсистемы движения (позиции p₁–p₄, переходы t₁–t₄) '
    'подматрица имеет вид:'
)

add_formula(
    'C_Д = [[-1, 0, 0, 1], [1, -1, 0, 0], [0, 1, -1, 0], [0, 0, 1, -1]]',
    6
)

add_paragraph(
    'Аналогичная структура для подсистемы сенсоров (p₅–p₈, t₅–t₈). '
    'Позиция p₉ (GPU) связана с переходами t₃, t₄, t₇, t₈: '
    'строка p₉ в матрице C: (0,\u00a00,\u00a0−1,\u00a01,\u00a00,\u00a00,\u00a0−1,\u00a01).'
)

add_heading_styled('4.2. Вычисление S-инвариантов')

add_paragraph(
    'Решая систему xᵀ\u00a0·\u00a0C\u00a0=\u00a00, получаем три линейно '
    'независимых S-инварианта:'
)

add_formula('x₁ = (1, 1, 1, 1, 0, 0, 0, 0, 0)ᵀ', 7)
add_formula('x₂ = (0, 0, 0, 0, 1, 1, 1, 1, 0)ᵀ', 8)
add_formula('x₃ = (0, 0, 0, 1, 0, 0, 0, 1, 1)ᵀ', 9)

add_paragraph(
    'Инвариант x₁ охватывает позиции подсистемы движения (p₁–p₄). '
    'По формуле (5):'
)

add_formula('M(p₁) + M(p₂) + M(p₃) + M(p₄) = 1 = const', 10)

add_paragraph(
    'Это означает, что в подсистеме движения всегда ровно один токен — '
    'процесс не может «раздвоиться» или «исчезнуть». Аналогично, '
    'инвариант x₂ гарантирует сохранение одного токена в подсистеме '
    'сенсоров.'
)

add_heading_styled('4.3. Доказательство взаимного исключения')

add_paragraph(
    'Инвариант x₃ — ключевой. Он связывает позиции p₄ (движение '
    'использует GPU), p₈ (сенсоры используют GPU) и p₉ (GPU свободен):'
)

add_formula('M(p₄) + M(p₈) + M(p₉) = M₀(p₄) + M₀(p₈) + M₀(p₉) = 1', 11)

add_paragraph(
    'Поскольку маркировка неотрицательна, из (11) следует, что '
    'M(p₄)\u00a0+\u00a0M(p₈)\u00a0≤\u00a01. Это формально доказывает '
    'взаимное исключение: невозможна ситуация, при которой обе подсистемы '
    'одновременно используют GPU (M(p₄)\u00a0=\u00a01 и M(p₈)\u00a0=\u00a01). '
    'Доказательство получено без перебора всех достижимых маркировок, '
    'что является преимуществом метода инвариантов [5].'
)

add_heading_styled('4.4. Анализ тупиковых ситуаций')

add_paragraph(
    'Сеть является живой (live): для каждого перехода существует '
    'последовательность срабатываний, возвращающая его в разрешённое '
    'состояние. Покажем это. Из начальной маркировки разрешены '
    'переходы t₁ и t₅. Последовательность σ\u00a0=\u00a0(t₁,\u00a0t₂,\u00a0t₃,\u00a0t₄) '
    'возвращает токен в p₁ и p₉, восстанавливая исходную маркировку '
    'для подсистемы движения. Аналогично, σ\u00a0=\u00a0(t₅,\u00a0t₆,\u00a0t₇,\u00a0t₈) '
    'восстанавливает маркировку подсистемы сенсоров. Поскольку оба цикла '
    'корректно освобождают мьютекс (p₉), тупиковая ситуация невозможна [8].'
)

add_paragraph(
    'Сеть также ограничена (bounded): из инвариантов (7)–(9) следует, '
    'что маркировка каждой позиции не превышает 1 (сеть является '
    '1-ограниченной, или безопасной). Это гарантирует конечность '
    'пространства состояний.'
)

# ══════════════ ТЕМПОРАЛЬНЫЙ АНАЛИЗ ══════════════
add_heading_styled('5. Темпоральный анализ')

add_heading_styled('5.1. Назначение задержек')

add_paragraph(
    'Для оценки производительности системы введём детерминированные '
    'задержки на переходах (в условных единицах времени) [6]:'
)

add_paragraph(
    '— d(t₁)\u00a0=\u00a02 (инициализация навигации);\n'
    '— d(t₂)\u00a0=\u00a08 (выполнение навигации);\n'
    '— d(t₃)\u00a0=\u00a01 (захват GPU);\n'
    '— d(t₄)\u00a0=\u00a05 (обработка карты движением);\n'
    '— d(t₅)\u00a0=\u00a02 (инициализация сканирования);\n'
    '— d(t₆)\u00a0=\u00a06 (сканирование лидаром);\n'
    '— d(t₇)\u00a0=\u00a01 (захват GPU);\n'
    '— d(t₈)\u00a0=\u00a010 (обработка облака точек).'
)

add_heading_styled('5.2. Время цикла')

add_paragraph(
    'Полный цикл подсистемы движения: '
    'd(t₁)\u00a0+\u00a0d(t₂)\u00a0+\u00a0d(t₃)\u00a0+\u00a0d(t₄)\u00a0=\u00a0'
    '2\u00a0+\u00a08\u00a0+\u00a01\u00a0+\u00a05\u00a0=\u00a016 ед.\u00a0вр. '
    'Полный цикл подсистемы сенсоров: '
    'd(t₅)\u00a0+\u00a0d(t₆)\u00a0+\u00a0d(t₇)\u00a0+\u00a0d(t₈)\u00a0=\u00a0'
    '2\u00a0+\u00a06\u00a0+\u00a01\u00a0+\u00a010\u00a0=\u00a019 ед.\u00a0вр.'
)

add_paragraph(
    'Без конфликта за GPU (если бы у каждой подсистемы был свой '
    'вычислитель) пропускная способность определялась бы более медленным '
    'циклом: 1/19\u00a0≈\u00a00,053 цикла/ед.\u00a0вр.'
)

add_heading_styled('5.3. Влияние конфликта за ресурс')

add_paragraph(
    'При наличии единственного GPU возникают ситуации ожидания. '
    'В наихудшем случае одна подсистема запрашивает GPU в момент, '
    'когда он занят другой. Максимальное дополнительное ожидание равно '
    'max(d(t₄),\u00a0d(t₈))\u00a0=\u00a010 ед.\u00a0вр. '
    'Эффективное время цикла системы в наихудшем случае:'
)

add_formula('T_eff = max(16, 19) + max(d(t₄), d(t₈)) = 19 + 10 = 29 ед. вр.', 12)

add_paragraph(
    'Узким местом является обработка облака точек (t₈,\u00a0d\u00a0=\u00a010), '
    'поскольку именно она создаёт максимальную задержку для конкурирующего '
    'процесса.'
)

add_heading_styled('5.4. Оптимизация: добавление второго GPU')

add_paragraph(
    'Моделирование добавления второго GPU осуществляется увеличением '
    'начальной маркировки позиции p₉: M₀(p₉)\u00a0=\u00a02. '
    'При этом инвариант (11) принимает вид:'
)

add_formula('M(p₄) + M(p₈) + M(p₉) = 2', 13)

add_paragraph(
    'Теперь обе подсистемы могут использовать GPU одновременно. '
    'Конфликт устраняется, и время цикла определяется только '
    'длительностью процессов: T_eff\u00a0=\u00a0max(16,\u00a019)\u00a0=\u00a019 ед.\u00a0вр. '
    'Ускорение:'
)

add_formula('S = T_old / T_new = 29 / 19 ≈ 1,53', 14)

add_paragraph(
    'Добавление второго GPU повышает производительность на 53\u00a0%. '
    'Однако при этом инвариант (13) допускает M(p₄)\u00a0+\u00a0M(p₈)\u00a0≤\u00a02, '
    'то есть свойство взаимного исключения ослабляется — оба процесса '
    'могут работать с GPU одновременно, что и является целью оптимизации, '
    'но требует соответствующей аппаратной поддержки [9].'
)

# ══════════════ ЗАКЛЮЧЕНИЕ ══════════════
add_heading_styled('6. Заключение')

add_paragraph(
    'В работе построена формальная модель высокоуровневого управления '
    'мобильным роботом с двумя параллельными подсистемами и общим '
    'вычислительным ресурсом на основе аппарата сетей Петри. '
    'Проведён структурный анализ модели:'
)

add_paragraph(
    '— методом S-инвариантов формально доказано взаимное исключение '
    'при доступе к GPU (формула 11);'
)
add_paragraph(
    '— показана живость и 1-ограниченность сети, что гарантирует '
    'отсутствие тупиковых ситуаций и конечность пространства состояний;'
)
add_paragraph(
    '— выполнен темпоральный анализ с детерминированными задержками, '
    'определено узкое место (обработка облака точек, 10\u00a0ед.\u00a0вр.) '
    'и показано, что добавление второго GPU ускоряет систему на 53\u00a0%.'
)

add_paragraph(
    'Предложенный подход позволяет на этапе проектирования выявлять '
    'потенциальные проблемы параллельного взаимодействия и количественно '
    'оценивать эффект конструктивных решений. Направлением дальнейшей '
    'работы является расширение модели стохастическими задержками '
    'для учёта вариативности времени сканирования и применение '
    'обобщённых стохастических сетей Петри (GSPN) [10].'
)

# ══════════════ ЛИТЕРАТУРА ══════════════
add_heading_styled('Список литературы')

refs = [
    'Lee\u00a0E.\u00a0A., Seshia\u00a0S.\u00a0A. Introduction to Embedded Systems: '
    'A Cyber-Physical Systems Approach. — 2nd ed. — MIT Press, 2017. — 568\u00a0p.',

    'Hopcroft\u00a0J.\u00a0E., Motwani\u00a0R., Ullman\u00a0J.\u00a0D. '
    'Introduction to Automata Theory, Languages, and Computation. — '
    '3rd ed. — Pearson, 2006. — 535\u00a0p.',

    'Petri\u00a0C.\u00a0A. Kommunikation mit Automaten: '
    'Dissertation. — Universität Bonn, 1962. — 128\u00a0S.',

    'Murata\u00a0T. Petri Nets: Properties, Analysis and Applications // '
    'Proceedings of the IEEE. — 1989. — Vol.\u00a077, No.\u00a04. — P.\u00a0541–580.',

    'Silva\u00a0M., Teruel\u00a0E., Colom\u00a0J.\u00a0M. Linear Algebraic and '
    'Linear Programming Techniques for the Analysis of Place/Transition Net Systems // '
    'Lectures on Petri Nets I: Basic Models. — Springer, 1998. — P.\u00a0309–373.',

    'Ramchandani\u00a0C. Analysis of Asynchronous Concurrent Systems by Timed '
    'Petri Nets: Technical Report MAC-TR-120. — MIT, 1974. — 106\u00a0p.',

    'Costelha\u00a0H., Lima\u00a0P. Robot Task Plan Representation by '
    'Petri Nets: Modelling, Identification, Analysis and Execution // '
    'Autonomous Robots. — 2012. — Vol.\u00a033, No.\u00a04. — P.\u00a0337–360.',

    'Ziparo\u00a0V.\u00a0A., Iocchi\u00a0L. Petri Net Plans: A Framework for '
    'Collaboration and Coordination in Multi-Robot Systems // '
    'Autonomous Agents and Multi-Agent Systems. — 2011. — Vol.\u00a022, '
    'No.\u00a03. — P.\u00a0498–531.',

    'Гуров\u00a0В.\u00a0С., Чугунов\u00a0В.\u00a0С. Моделирование процессов '
    'управления роботами на основе сетей Петри // Мехатроника, автоматизация, '
    'управление. — 2018. — Т.\u00a019, №\u00a05. — С.\u00a0312–319.',

    'Marsan\u00a0M.\u00a0A., Balbo\u00a0G., Conte\u00a0G. et al. '
    'Modelling with Generalized Stochastic Petri Nets. — '
    'John Wiley & Sons, 1995. — 316\u00a0p.',
]

for i, ref in enumerate(refs, 1):
    p = add_paragraph(f'{i}. {ref}', indent=False)
    p.paragraph_format.first_line_indent = Cm(0)

# ══════════════ ENGLISH BLOCK ══════════════
add_paragraph('', space_after=12)  # пустая строка

add_paragraph(
    'MODELING OF PARALLEL PROCESSES\n'
    'IN A ROBOTIC SYSTEM USING PETRI NETS',
    bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER, indent=False, space_after=6
)

add_paragraph(
    'V.\u00a0P.\u00a0Romanov',
    alignment=WD_ALIGN_PARAGRAPH.CENTER, indent=False
)

add_paragraph(
    'Bauman Moscow State Technical University, Moscow, Russian Federation',
    italic=True, alignment=WD_ALIGN_PARAGRAPH.CENTER, indent=False,
    space_after=12
)

add_rich_paragraph([
    ('Abstract. ', True, False),
    ('The paper addresses the problem of formal modeling of high-level control '
     'of a robotic system using Petri nets. The necessity of applying formal methods '
     'for verification of parallel processes in robotics is substantiated. '
     'A model of a mobile robot with two parallel subsystems — motion and sensing — '
     'interacting through a shared resource is proposed. The model is analyzed '
     'using S-invariants: mutual exclusion for the shared resource access '
     'and absence of deadlocks are formally proven. Temporal analysis with '
     'deterministic delays is performed, identifying system throughput and the '
     'bottleneck. It is shown that adding a parallel resource improves performance '
     'by 53%. The results can be applied to the design of autonomous robot '
     'control systems.', False, False),
], indent=True, space_after=6)

add_rich_paragraph([
    ('Keywords: ', True, False),
    ('Petri nets, robotics, parallel processes, formal verification, '
     'S-invariants, mutual exclusion, temporal analysis, throughput, '
     'high-level control, deadlock', False, False),
], indent=True, space_after=12)

# ══════════════ СВЕДЕНИЯ ОБ АВТОРЕ ══════════════
add_heading_styled('Сведения об авторе')

add_paragraph(
    'Романов Владимир Павлович — студент кафедры «Прикладная робототехника», '
    'факультет «Специальное машиностроение», '
    'МГТУ им.\u00a0Н.\u00a0Э.\u00a0Баумана (Москва, Российская Федерация).',
    indent=False
)

add_paragraph('', space_after=6)

add_rich_paragraph([
    ('Научный руководитель: ', True, False),
    ('Шевченко Максим Юрьевич, '
     'МГТУ им.\u00a0Н.\u00a0Э.\u00a0Баумана (Москва, Российская Федерация).', False, False),
], indent=False, space_after=12)

add_heading_styled('About the author')

add_paragraph(
    'Romanov Vladimir P. — student, Department of Applied Robotics, '
    'Faculty of Special Mechanical Engineering, '
    'Bauman Moscow State Technical University (Moscow, Russian Federation).',
    indent=False
)

add_paragraph('', space_after=6)

add_rich_paragraph([
    ('Scientific advisor: ', True, False),
    ('Shevchenko Maxim Yu., '
     'Bauman Moscow State Technical University (Moscow, Russian Federation).', False, False),
], indent=False)

# ── Нумерация страниц ──
add_page_numbers(doc)

# ── Сохранение ──
output_path = os.path.join(os.path.dirname(__file__), 'Романов.docx')
doc.save(output_path)
print(f'Статья сохранена: {output_path}')
