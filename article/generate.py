#!/usr/bin/env python3
"""
Генератор статьи для журнала МГТУ им. Н.Э. Баумана (ptsj.bmstu.ru)
Тема: Верификация триподной походки гексапода сетями Петри
v3: OMML-формулы, правильные кавычки/тире, отступы по ГОСТ
"""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '.pylibs'))

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn, nsmap
from docx.oxml import OxmlElement
from lxml import etree
from copy import deepcopy

# ── Константы ──
FONT = 'Times New Roman'
SIZE = Pt(12)
INDENT = Cm(0.7)
NBSP = '\u00a0'
EMDASH = '\u2014'
LAQUO = '\u00ab'
RAQUO = '\u00bb'

doc = Document()

# ── Секция (поля, ориентация) ──
sec = doc.sections[0]
sec.orientation = WD_ORIENT.PORTRAIT
sec.page_width = Cm(21)
sec.page_height = Cm(29.7)
sec.left_margin = Cm(2)
sec.top_margin = Cm(2)
sec.right_margin = Cm(1.5)
sec.bottom_margin = Cm(1.5)

# ── Стиль Normal ──
sty = doc.styles['Normal']
sty.font.name = FONT
sty.font.size = SIZE
sty.font.color.rgb = RGBColor(0, 0, 0)
sty.paragraph_format.line_spacing = 1.5
sty.paragraph_format.space_after = Pt(0)
sty.paragraph_format.space_before = Pt(0)
sty.paragraph_format.first_line_indent = INDENT

rpr = sty.element.get_or_add_rPr()
rf = OxmlElement('w:rFonts')
for a in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
    rf.set(qn(a), FONT)
rpr.append(rf)


# ══════════════════════════════════════════
#  ХЕЛПЕРЫ
# ══════════════════════════════════════════

def set_font(run):
    r = run._element
    rp = r.get_or_add_rPr()
    rf = OxmlElement('w:rFonts')
    for a in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
        rf.set(qn(a), FONT)
    rp.insert(0, rf)


def add_page_numbers():
    footer = doc.sections[0].footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    fld = OxmlElement('w:fldSimple')
    fld.set(qn('w:instr'), 'PAGE')
    run._element.append(fld)


def P(text, bold=False, italic=False, align=None, indent=True,
      sz=None, after=None, before=None):
    p = doc.add_paragraph()
    if not indent:
        p.paragraph_format.first_line_indent = Cm(0)
    if align:
        p.alignment = align
    if after is not None:
        p.paragraph_format.space_after = Pt(after)
    if before is not None:
        p.paragraph_format.space_before = Pt(before)
    run = p.add_run(text)
    run.font.name = FONT
    run.font.size = sz or SIZE
    run.bold = bold
    run.italic = italic
    set_font(run)
    return p


def RP(parts, align=None, indent=True, after=None, before=None):
    """[(text, bold, italic), ...]"""
    p = doc.add_paragraph()
    if not indent:
        p.paragraph_format.first_line_indent = Cm(0)
    if align:
        p.alignment = align
    if after is not None:
        p.paragraph_format.space_after = Pt(after)
    if before is not None:
        p.paragraph_format.space_before = Pt(before)
    for text, bold, italic in parts:
        run = p.add_run(text)
        run.font.name = FONT
        run.font.size = SIZE
        run.bold = bold
        run.italic = italic
        set_font(run)
    return p


def H(text):
    return P(text, bold=True, indent=False, before=12, after=6)


def LIST(text):
    """Элемент списка с длинным тире, без отступа первой строки."""
    return P(f'{EMDASH}\u2002{text}', indent=False)


def OMML(formula_text, number=None):
    """Формула через Word OMML (Office Math Markup Language).
    Создаёт настоящий объект-формулу, который Word распознаёт и может
    редактировать через встроенный редактор / MathType.
    """
    MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'

    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Cm(0)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Создаём oMathPara > oMath > r > t
    oMathPara = etree.SubElement(p._element, qn('m:oMathPara'))
    oMath = etree.SubElement(oMathPara, qn('m:oMath'))

    # Разбиваем формулу на части (простой подход: весь текст одним run)
    r = etree.SubElement(oMath, qn('m:r'))

    # Свойства run (шрифт)
    rPr = etree.SubElement(r, qn('m:rPr'))
    sty_el = etree.SubElement(rPr, qn('m:sty'))
    sty_el.set(qn('m:val'), 'p')  # plain (не курсив)

    # Текстовые свойства
    w_rPr = etree.SubElement(r, qn('w:rPr'))
    w_rf = etree.SubElement(w_rPr, qn('w:rFonts'))
    w_rf.set(qn('w:ascii'), 'Cambria Math')
    w_rf.set(qn('w:hAnsi'), 'Cambria Math')

    t = etree.SubElement(r, qn('m:t'))
    t.text = formula_text

    # Номер формулы справа
    if number:
        # Добавляем табуляцию и номер после формулы как обычный run
        tab_run = p.add_run(f'\t({number})')
        tab_run.font.name = FONT
        tab_run.font.size = SIZE
        set_font(tab_run)

    return p


def make_table(headers, rows, caption=None, caption_above=True):
    """Таблица Word с заголовком. caption_above=True — заголовок НАД таблицей."""
    # Заголовок таблицы (над таблицей по требованиям)
    if caption and caption_above:
        P(caption, italic=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          indent=False, before=6, after=4)

    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = 'Table Grid'

    # Заголовки
    for j, h in enumerate(headers):
        cell = table.cell(0, j)
        cell.text = ''
        cp = cell.paragraphs[0]
        cp.paragraph_format.first_line_indent = Cm(0)
        cp.paragraph_format.space_after = Pt(0)
        cp.paragraph_format.space_before = Pt(0)
        run = cp.add_run(h)
        run.font.name = FONT
        run.font.size = Pt(10)
        run.bold = True
        set_font(run)

    # Данные
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = table.cell(i + 1, j)
            cell.text = ''
            cp = cell.paragraphs[0]
            cp.paragraph_format.first_line_indent = Cm(0)
            cp.paragraph_format.space_after = Pt(0)
            cp.paragraph_format.space_before = Pt(0)
            run = cp.add_run(str(val))
            run.font.name = FONT
            run.font.size = Pt(10)
            set_font(run)

    if caption and not caption_above:
        P(caption, italic=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          indent=False, before=4, after=6)

    return table


# ══════════════════════════════════════════
#  ТЕКСТ СТАТЬИ
# ══════════════════════════════════════════

# ── УДК ──
P('УДК 004.414.23 + 621.865.8', indent=False, after=6)

# ── Название ──
P(f'ФОРМАЛЬНАЯ ВЕРИФИКАЦИЯ КООРДИНАЦИИ ПОХОДКИ\n'
  f'ШЕСТИНОГОГО РОБОТА МЕТОДОМ СЕТЕЙ ПЕТРИ',
  bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, after=6)

# ── Автор (инициалы без пробела между собой, NBSP перед фамилией) ──
P(f'В.П.{NBSP}Романов', align=WD_ALIGN_PARAGRAPH.CENTER, indent=False)

# ── Email ──
P('romanov.vp@student.bmstu.ru',
  italic=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False)

# ── Организация ──
P(f'МГТУ им.{NBSP}Н.Э.{NBSP}Баумана, Москва, Российская Федерация',
  italic=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, after=12)

# ── Аннотация ──
RP([
    ('Аннотация. ', True, False),
    (f'Рассматривается задача формальной верификации алгоритма триподной '
     f'походки шестиногого шагающего робота (гексапода) с использованием '
     f'аппарата сетей Петри. Шесть ног робота разделены на две группы '
     f'по три ноги, поочерёдно выполняющие фазы переноса и опоры. '
     f'Построена сеть Петри, моделирующая координацию групп ног '
     f'с ограничением статической устойчивости: в каждый момент времени '
     f'не менее трёх ног должны находиться в контакте с поверхностью. '
     f'Методом S-инвариантов формально доказаны взаимное исключение '
     f'фаз переноса (одновременный отрыв обеих групп невозможен) '
     f'и отсутствие тупиковых ситуаций. Выполнен темпоральный анализ '
     f'с детерминированными задержками: определены время цикла походки, '
     f'максимальная скорость перемещения и узкое место. Показано, что '
     f'применение сервоприводов с уменьшенным временем переноса на 25{NBSP}% '
     f'сокращает период походки на 17{NBSP}% и увеличивает скорость робота '
     f'со 200 до 240{NBSP}мм/с.', False, False),
], after=6)

# ── Ключевые слова ──
RP([
    ('Ключевые слова: ', True, False),
    ('сети Петри, шагающий робот, гексапод, триподная походка, '
     'формальная верификация, S-инварианты, координация походки, '
     'статическая устойчивость, темпоральный анализ, взаимное исключение', False, False),
], after=12)


# ══════════════ 1. ВВЕДЕНИЕ ══════════════
H('1. Введение')

P(f'Шестиногие шагающие роботы (гексаподы) обладают высокой проходимостью '
  f'и статической устойчивостью, что делает их перспективными для '
  f'работы в неструктурированных средах: пересечённая местность, '
  f'аварийные здания, трубопроводы [1]. Ключевой задачей при разработке '
  f'гексапода является координация шести ног{NBSP}{EMDASH} алгоритм походки, '
  f'который определяет, какие ноги в какой момент отрываются от '
  f'поверхности и переносятся вперёд.')

P(f'Наиболее распространённая походка гексапода{NBSP}{EMDASH} триподная '
  f'(tripod gait): шесть ног разделены на две группы по три (1-3-5 '
  f'и 2-4-6), которые поочерёдно выполняют фазы переноса и опоры [2]. '
  f'Триподная походка обеспечивает максимальную скорость при '
  f'гарантированной статической устойчивости: три опорные ноги всегда '
  f'образуют треугольник, внутри которого находится проекция центра '
  f'масс робота.')

P(f'Однако при программной реализации координация шести параллельно '
  f'работающих приводов порождает класс ошибок, характерных для '
  f'параллельных систем. Гонка состояний (race condition) может '
  f'привести к одновременному отрыву обеих групп ног, что вызовет '
  f'потерю устойчивости и падение робота. Тупиковая ситуация (deadlock) '
  f'может заблокировать движение, если обе группы бесконечно ожидают '
  f'друг друга [3]. Такие ошибки сложно обнаружить тестированием, '
  f'поскольку они возникают лишь при определённых временны\u0301х '
  f'соотношениях между процессами.')

P(f'Конечные автоматы (FSM), традиционно используемые для описания '
  f'логики управления, плохо приспособлены к моделированию параллелизма: '
  f'два процесса с N и M состояниями порождают N{NBSP}\u00d7{NBSP}M '
  f'комбинаций. Для шести ног даже при двух состояниях на ногу '
  f'(опора/перенос) это 2\u2076{NBSP}={NBSP}64 состояния, а при более '
  f'детальном описании{NBSP}{EMDASH} сотни и тысячи [4].')

P(f'Сети Петри, предложенные К.А.{NBSP}Петри в 1962{NBSP}г. [5], '
  f'позволяют моделировать параллельные процессы без комбинаторного '
  f'взрыва. Математический аппарат сетей Петри (матричное представление, '
  f'S-инварианты, дерево достижимости) даёт возможность формально '
  f'доказывать свойства системы: отсутствие тупиков, ограниченность, '
  f'взаимное исключение [6].')

P(f'Целью настоящей работы является построение формальной модели '
  f'координации триподной походки гексапода, доказательство '
  f'корректности алгоритма (устойчивость и отсутствие тупиков) '
  f'методом S-инвариантов, а также темпоральный анализ для определения '
  f'предельной скорости перемещения.')


# ══════════════ 2. ТЕОРЕТИЧЕСКИЕ ОСНОВЫ ══════════════
H('2. Теоретические основы')
H('2.1. Определение сети Петри')

P('Сеть Петри определяется как четвёрка [5, 6]:')

OMML('PN=(P, T, F, M\u2080)', 1)

P(f'где P{NBSP}={NBSP}{{p\u2081, p\u2082, \u2026, p\u2098}}{NBSP}'
  f'{EMDASH} конечное множество позиций (мест), моделирующих условия '
  f'или состояния; T{NBSP}={NBSP}{{t\u2081, t\u2082, \u2026, t\u2099}}'
  f'{NBSP}{EMDASH} конечное множество переходов, моделирующих события; '
  f'F{NBSP}\u2286{NBSP}(P{NBSP}\u00d7{NBSP}T){NBSP}\u222a{NBSP}'
  f'(T{NBSP}\u00d7{NBSP}P){NBSP}{EMDASH} множество направленных дуг; '
  f'M\u2080:{NBSP}P{NBSP}\u2192{NBSP}\u2115{NBSP}{EMDASH} начальная '
  f'маркировка, определяющая распределение токенов по позициям.')

H('2.2. Правило срабатывания')

P(f'Переход t{NBSP}\u2208{NBSP}T разрешён при маркировке M, если каждая '
  f'входная позиция содержит не менее одного токена: '
  f'\u2200p{NBSP}\u2208{NBSP}\u2022t:{NBSP}M(p){NBSP}\u2265{NBSP}1. '
  f'При срабатывании перехода маркировка обновляется [6]:')

OMML("M'(p) = M(p) \u2212 W(p, t) + W(t, p)", 2)

P(f'где W(p,{NBSP}t){NBSP}{EMDASH} вес дуги из позиции в переход, '
  f'W(t,{NBSP}p){NBSP}{EMDASH} вес дуги из перехода в позицию.')

H('2.3. Матричное представление и S-инварианты')

P(f'Матрица инцидентности C размером |P|{NBSP}\u00d7{NBSP}|T| '
  f'определяется как разность матриц выходных и входных дуг [6]:')

OMML('C = C\u207a \u2212 C\u207b', 3)

P('Уравнение состояния сети связывает начальную и текущую маркировки '
  'через вектор срабатываний \u03c3:')

OMML('M = M\u2080 + C \u00b7 \u03c3', 4)

P(f'Вектор x{NBSP}\u2208{NBSP}\u2124\u1d50 называется S-инвариантом, '
  f'если x\u1d40{NBSP}\u00b7{NBSP}C{NBSP}={NBSP}0. '
  f'Для любой достижимой маркировки M справедливо [7]:')

OMML('x\u1d40 \u00b7 M = x\u1d40 \u00b7 M\u2080 = const', 5)

P('S-инварианты позволяют доказывать свойства системы без перебора '
  'всех достижимых маркировок, что критически важно для систем '
  'с большим пространством состояний.')

H('2.4. Темпоральные расширения')

P(f'Темпоральная сеть Петри дополняет классическую модель функцией '
  f'задержки D:{NBSP}T{NBSP}\u2192{NBSP}\u211d\u207a, сопоставляющей '
  f'каждому переходу время срабатывания [8]. После разрешения перехода '
  f'входные токены резервируются на время d(t), по истечении которого '
  f'переход срабатывает. Это позволяет моделировать длительность '
  f'физических операций и анализировать производительность.')


# ══════════════ 3. МОДЕЛЬ ══════════════
H('3. Модель координации триподной походки')
H('3.1. Описание объекта')

P(f'Рассматривается шестиногий шагающий робот (гексапод) '
  f'с шестью идентичными ногами, расположенными симметрично '
  f'по три с каждой стороны корпуса. Каждая нога оснащена тремя '
  f'сервоприводами (тазобедренный, коленный, голеностопный) '
  f'и может находиться в одном из двух режимов: опора (stance){NBSP}'
  f'{EMDASH} нога в контакте с поверхностью и несёт нагрузку, '
  f'или перенос (swing){NBSP}{EMDASH} нога поднята и перемещается вперёд.')

P(f'При триподной походке ноги разделены на две группы: '
  f'группа{NBSP}A (ноги 1, 3, 5) и группа{NBSP}B (ноги 2, 4, 6). '
  f'Группы работают в противофазе: пока одна группа переносится, '
  f'другая опирается. Критическое ограничение: в любой момент '
  f'времени не менее трёх ног должны быть в контакте с поверхностью '
  f'для обеспечения статической устойчивости [2].')

H('3.2. Структура сети Петри')

P(f'Каждая группа ног проходит три фазы в цикле: ожидание '
  f'(готовность к переносу), перенос (ноги в воздухе), '
  f'приземление (установка контакта с поверхностью). '
  f'Координация обеспечивается токеном-мьютексом, '
  f'моделирующим ограничение устойчивости.')

P('Позиции модели:')

LIST(f'p\u2081 ({LAQUO}Готовность{NBSP}A{RAQUO}){NBSP}{EMDASH} группа{NBSP}A в опоре, готова к переносу;')
LIST(f'p\u2082 ({LAQUO}Перенос{NBSP}A{RAQUO}){NBSP}{EMDASH} группа{NBSP}A в воздухе, ноги перемещаются;')
LIST(f'p\u2083 ({LAQUO}Приземление{NBSP}A{RAQUO}){NBSP}{EMDASH} группа{NBSP}A устанавливает контакт;')
LIST(f'p\u2084 ({LAQUO}Готовность{NBSP}B{RAQUO}){NBSP}{EMDASH} группа{NBSP}B в опоре, готова к переносу;')
LIST(f'p\u2085 ({LAQUO}Перенос{NBSP}B{RAQUO}){NBSP}{EMDASH} группа{NBSP}B в воздухе;')
LIST(f'p\u2086 ({LAQUO}Приземление{NBSP}B{RAQUO}){NBSP}{EMDASH} группа{NBSP}B устанавливает контакт;')
LIST(f'p\u2087 ({LAQUO}Устойчивость{RAQUO}){NBSP}{EMDASH} мьютекс, гарантирующий, что не более одной группы находится в переносе одновременно.')

P('Переходы модели:')

LIST(f't\u2081 ({LAQUO}Отрыв{NBSP}A{RAQUO}): p\u2081{NBSP}+{NBSP}p\u2087{NBSP}\u2192{NBSP}p\u2082{NBSP}{EMDASH} группа{NBSP}A начинает перенос, захватывая токен устойчивости;')
LIST(f't\u2082 ({LAQUO}Перенос{NBSP}A завершён{RAQUO}): p\u2082{NBSP}\u2192{NBSP}p\u2083{NBSP}{EMDASH} ноги группы{NBSP}A достигли целевых позиций;')
LIST(f't\u2083 ({LAQUO}Контакт{NBSP}A{RAQUO}): p\u2083{NBSP}\u2192{NBSP}p\u2081{NBSP}+{NBSP}p\u2087{NBSP}{EMDASH} группа{NBSP}A установила контакт, токен устойчивости освобождён;')
LIST(f't\u2084 ({LAQUO}Отрыв{NBSP}B{RAQUO}): p\u2084{NBSP}+{NBSP}p\u2087{NBSP}\u2192{NBSP}p\u2085;')
LIST(f't\u2085 ({LAQUO}Перенос{NBSP}B завершён{RAQUO}): p\u2085{NBSP}\u2192{NBSP}p\u2086;')
LIST(f't\u2086 ({LAQUO}Контакт{NBSP}B{RAQUO}): p\u2086{NBSP}\u2192{NBSP}p\u2084{NBSP}+{NBSP}p\u2087.')

P(f'Начальная маркировка: M\u2080{NBSP}={NBSP}(1,{NBSP}0,{NBSP}0,{NBSP}'
  f'1,{NBSP}0,{NBSP}0,{NBSP}1). Один токен в p\u2081 (группа{NBSP}A '
  f'готова), один в p\u2084 (группа{NBSP}B готова), один в p\u2087 '
  f'(устойчивость обеспечена). Из начального состояния разрешены '
  f'переходы t\u2081 и t\u2084, но при срабатывании одного из них '
  f'токен p\u2087 изымается, и второй переход блокируется до его возврата.')

P(f'[Рис.{NBSP}1. Сеть Петри координации триподной походки гексапода]',
  italic=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, before=6, after=6)


# ══════════════ 4. СТРУКТУРНЫЙ АНАЛИЗ ══════════════
H('4. Структурный анализ модели')
H('4.1. Матрица инцидентности')

P(f'Матрица инцидентности C размером 7{NBSP}\u00d7{NBSP}6 построена '
  f'по формуле (3). Строки соответствуют позициям p\u2081{EMDASH}p\u2087, '
  f'столбцы{NBSP}{EMDASH} переходам t\u2081{EMDASH}t\u2086.')

# Матрица как таблица
make_table(
    headers=['', 't\u2081', 't\u2082', 't\u2083', 't\u2084', 't\u2085', 't\u2086'],
    rows=[
        ['p\u2081', '\u22121', '0', '1', '0', '0', '0'],
        ['p\u2082', '1', '\u22121', '0', '0', '0', '0'],
        ['p\u2083', '0', '1', '\u22121', '0', '0', '0'],
        ['p\u2084', '0', '0', '0', '\u22121', '0', '1'],
        ['p\u2085', '0', '0', '0', '1', '\u22121', '0'],
        ['p\u2086', '0', '0', '0', '0', '1', '\u22121'],
        ['p\u2087', '\u22121', '0', '1', '\u22121', '0', '1'],
    ],
    caption=f'Табл.{NBSP}1. Матрица инцидентности C'
)

P('', after=6)

H('4.2. Вычисление S-инвариантов')

P(f'Решая однородную систему x\u1d40{NBSP}\u00b7{NBSP}C{NBSP}={NBSP}0 '
  f'методом Гаусса, получаем три линейно независимых S-инварианта:')

OMML('x\u2081 = (1, 1, 1, 0, 0, 0, 0)\u1d40', 6)
OMML('x\u2082 = (0, 0, 0, 1, 1, 1, 0)\u1d40', 7)
OMML('x\u2083 = (0, 1, 1, 0, 1, 1, 1)\u1d40', 8)

P(f'Множество позиций, охваченных инвариантами, покрывает все позиции '
  f'сети (объединение носителей x\u2081, x\u2082, x\u2083 равно P). '
  f'Это означает, что сеть является консервативной{NBSP}{EMDASH} ни один '
  f'токен не создаётся и не уничтожается [6].')

H('4.3. Интерпретация инвариантов')

P(f'Инвариант x\u2081 охватывает позиции группы{NBSP}A '
  f'(p\u2081,{NBSP}p\u2082,{NBSP}p\u2083). Применяя формулу (5):')

OMML('M(p\u2081) + M(p\u2082) + M(p\u2083) = 1', 9)

P(f'Это означает, что группа{NBSP}A всегда находится ровно в одной '
  f'из трёх фаз{NBSP}{EMDASH} процесс не может {LAQUO}раздвоиться{RAQUO} '
  f'(появление лишнего токена) или {LAQUO}исчезнуть{RAQUO} '
  f'(потеря управления). Аналогичное свойство гарантирует '
  f'инвариант x\u2082 для группы{NBSP}B:')

OMML('M(p\u2084) + M(p\u2085) + M(p\u2086) = 1', 10)

H('4.4. Доказательство статической устойчивости')

P(f'Инвариант x\u2083{NBSP}{EMDASH} центральный результат анализа. '
  f'Он связывает позиции переноса обеих групп '
  f'(p\u2082,{NBSP}p\u2083,{NBSP}p\u2085,{NBSP}p\u2086) '
  f'и мьютекс устойчивости (p\u2087):')

OMML('M(p\u2082) + M(p\u2083) + M(p\u2085) + M(p\u2086) + M(p\u2087) = 1', 11)

P('Поскольку маркировка неотрицательна, из (11) следует:')

OMML('M(p\u2082) + M(p\u2085) \u2264 1', 12)

P(f'Неравенство (12) формально доказывает взаимное исключение: '
  f'невозможна ситуация, при которой обе группы ног одновременно '
  f'находятся в фазе переноса (M(p\u2082){NBSP}={NBSP}1 и '
  f'M(p\u2085){NBSP}={NBSP}1). В физическом смысле это означает, что '
  f'в каждый момент времени не менее трёх ног остаются в контакте '
  f'с поверхностью, что гарантирует статическую устойчивость робота.')

P(f'Данное доказательство получено без перебора достижимых '
  f'маркировок{NBSP}{EMDASH} оно следует из структуры сети и справедливо '
  f'для любой последовательности срабатываний. Это преимущество метода '
  f'инвариантов перед анализом дерева достижимости [7].')

H('4.5. Отсутствие тупиковых ситуаций')

P(f'Покажем, что сеть является живой (live). Из произвольной '
  f'достижимой маркировки рассмотрим возможные случаи:')

P(f'1){NBSP}Токен в p\u2081, токен в p\u2084, токен в p\u2087{NBSP}'
  f'{EMDASH} разрешены t\u2081 и t\u2084. Конфликт разрешается '
  f'недетерминированно, но не приводит к блокировке.')

P(f'2){NBSP}Токен в p\u2082 (группа{NBSP}A переносится){NBSP}'
  f'{EMDASH} разрешён t\u2082. После срабатывания t\u2082 токен '
  f'перейдёт в p\u2083, далее разрешён t\u2083, который вернёт '
  f'токен в p\u2081 и p\u2087.')

P(f'3){NBSP}Токен в p\u2083{NBSP}{EMDASH} разрешён t\u2083. '
  f'Срабатывание возвращает систему в состояние, откуда '
  f'возможен новый цикл.')

P(f'Симметричные рассуждения справедливы для группы{NBSP}B. '
  f'Поскольку каждый цикл (t\u2081\u2192t\u2082\u2192t\u2083 '
  f'или t\u2084\u2192t\u2085\u2192t\u2086) гарантированно завершается '
  f'и возвращает токен мьютекса, тупиковая ситуация невозможна. '
  f'Сеть является живой и 1-ограниченной (безопасной).')

H('4.6. Пошаговая трассировка маркировок')

P(f'Для наглядности приведём первые шаги выполнения сети, '
  f'начиная с M\u2080{NBSP}={NBSP}(1, 0, 0, 1, 0, 0, 1).')

make_table(
    headers=['Шаг', 'Переход', 'p\u2081', 'p\u2082', 'p\u2083',
             'p\u2084', 'p\u2085', 'p\u2086', 'p\u2087', 'Комментарий'],
    rows=[
        ['0', EMDASH, '1', '0', '0', '1', '0', '0', '1', 'Начальное состояние'],
        ['1', 't\u2081', '0', '1', '0', '1', '0', '0', '0', 'Группа A оторвалась'],
        ['2', 't\u2082', '0', '0', '1', '1', '0', '0', '0', 'A перенесена'],
        ['3', 't\u2083', '1', '0', '0', '1', '0', '0', '1', 'A приземлилась'],
        ['4', 't\u2084', '1', '0', '0', '0', '1', '0', '0', 'Группа B оторвалась'],
        ['5', 't\u2085', '1', '0', '0', '0', '0', '1', '0', 'B перенесена'],
        ['6', 't\u2086', '1', '0', '0', '1', '0', '0', '1', f'B приземлилась = M\u2080'],
    ],
    caption=f'Табл.{NBSP}2. Трассировка маркировок при последовательном срабатывании'
)

P(f'', after=4)

P(f'Из таблицы видно, что после шага 6 система возвращается '
  f'в начальную маркировку M\u2080{NBSP}{EMDASH} цикл замкнулся. '
  f'На протяжении всей трассировки '
  f'M(p\u2082){NBSP}+{NBSP}M(p\u2085){NBSP}\u2264{NBSP}1, '
  f'что подтверждает аналитическое доказательство (12).')


# ══════════════ 5. ТЕМПОРАЛЬНЫЙ АНАЛИЗ ══════════════
H('5. Темпоральный анализ')
H('5.1. Назначение задержек')

P(f'Для количественной оценки производительности системы '
  f'присвоим переходам детерминированные задержки, '
  f'соответствующие реальным временам фаз походки '
  f'гексапода массой 2{NBSP}кг с сервоприводами класса '
  f'25{NBSP}кг\u00b7см [8]:')

LIST(f'd(t\u2081){NBSP}={NBSP}50{NBSP}мс (отрыв группы{NBSP}A: подъём трёх ног);')
LIST(f'd(t\u2082){NBSP}={NBSP}200{NBSP}мс (перенос: перемещение ног вперёд на 60{NBSP}мм);')
LIST(f'd(t\u2083){NBSP}={NBSP}50{NBSP}мс (приземление: установка контакта, стабилизация);')
LIST(f'd(t\u2084){NBSP}={NBSP}50{NBSP}мс (отрыв группы{NBSP}B);')
LIST(f'd(t\u2085){NBSP}={NBSP}200{NBSP}мс (перенос{NBSP}B);')
LIST(f'd(t\u2086){NBSP}={NBSP}50{NBSP}мс (приземление{NBSP}B).')

H('5.2. Время цикла и скорость')

P(f'Полный цикл группы{NBSP}A: '
  f'd(t\u2081){NBSP}+{NBSP}d(t\u2082){NBSP}+{NBSP}d(t\u2083){NBSP}'
  f'={NBSP}50{NBSP}+{NBSP}200{NBSP}+{NBSP}50{NBSP}={NBSP}300{NBSP}мс. '
  f'Аналогично для группы{NBSP}B: 300{NBSP}мс. Поскольку группы работают '
  f'последовательно (мьютекс запрещает параллельный перенос), '
  f'полный период походки:')

OMML('T = T_A + T_B = 300 + 300 = 600 \u043c\u0441', 13)

P(f'За один период робот совершает два шага (по одному на группу), '
  f'каждый длиной l{NBSP}={NBSP}60{NBSP}мм. Скорость перемещения:')

OMML('V = 2l / T = 2 \u00d7 60 / 600 = 200 \u043c\u043c/\u0441', 14)

P(f'Это согласуется с типичными значениями для гексаподов '
  f'данного класса: 150{EMDASH}300{NBSP}мм/с при триподной походке [9].')

H('5.3. Узкое место')

P(f'Фаза переноса (t\u2082 и t\u2085) занимает 200{NBSP}мс из '
  f'300{NBSP}мс полуцикла{NBSP}{EMDASH} 67{NBSP}% времени. '
  f'Фазы отрыва и приземления (по 50{NBSP}мс каждая) суммарно '
  f'составляют 100{NBSP}мс. Таким образом, узким местом является '
  f'именно перенос ног{NBSP}{EMDASH} скорость работы сервоприводов '
  f'при перемещении ног на заданное расстояние.')

P(f'Это соответствует инженерной интуиции: для увеличения скорости '
  f'походки в первую очередь следует сокращать время переноса '
  f'(более быстрые сервоприводы или меньшая амплитуда шага), '
  f'а не время отрыва и приземления.')

H('5.4. Оценка эффекта модернизации')

P(f'Рассмотрим замену сервоприводов на более быстрые, '
  f'сокращающие время переноса на 25{NBSP}%: '
  f'd(t\u2082){NBSP}={NBSP}d(t\u2085){NBSP}={NBSP}150{NBSP}мс '
  f'вместо 200{NBSP}мс. Новый период походки:')

OMML("T' = 2 \u00d7 (50 + 150 + 50) = 500 \u043c\u0441", 15)

P('Новая скорость:')

OMML("V' = 2 \u00d7 60 / 500 = 240 \u043c\u043c/\u0441", 16)

P(f'Ускорение составляет 240/200{NBSP}={NBSP}1,2, то{NBSP}есть '
  f'прирост скорости на 20{NBSP}%. При этом затрачено на 25{NBSP}% '
  f'меньше времени на перенос, но итоговое ускорение{NBSP}{EMDASH} '
  f'лишь 20{NBSP}%, поскольку фиксированные фазы отрыва и '
  f'приземления {LAQUO}разбавляют{RAQUO} эффект. Подобные количественные '
  f'оценки позволяют обоснованно выбирать компоненты на этапе '
  f'проектирования, не прибегая к физическому прототипированию.')


# ══════════════ 6. ЗАКЛЮЧЕНИЕ ══════════════
H('6. Заключение')

P(f'В работе построена формальная модель координации триподной '
  f'походки шестиногого робота на основе аппарата сетей Петри. '
  f'Модель включает семь позиций и шесть переходов, описывающих '
  f'чередование фаз переноса и опоры двух групп ног с '
  f'ограничением статической устойчивости.')

P('Основные результаты:')

P(f'1){NBSP}Методом S-инвариантов формально доказано свойство '
  f'взаимного исключения: обе группы ног не могут одновременно '
  f'находиться в фазе переноса (неравенство 12), что гарантирует '
  f'статическую устойчивость робота{NBSP}{EMDASH} не менее трёх ног '
  f'в контакте с поверхностью в любой момент времени.')

P(f'2){NBSP}Показана живость и 1-ограниченность сети, '
  f'что исключает тупиковые ситуации и гарантирует корректную '
  f'работу алгоритма при любой последовательности событий.')

P(f'3){NBSP}Темпоральный анализ позволил определить период '
  f'триподной походки (600{NBSP}мс), скорость перемещения '
  f'(200{NBSP}мм/с) и узкое место (фаза переноса{NBSP}{EMDASH} '
  f'67{NBSP}% времени цикла). Оценка модернизации показала, что '
  f'замена сервоприводов с сокращением времени переноса на 25{NBSP}% '
  f'даёт прирост скорости на 20{NBSP}%.')

P(f'Предложенный подход может быть расширен на другие типы '
  f'походок (волновая, рысь) путём изменения структуры сети '
  f'и перераспределения токенов мьютекса. Направлением '
  f'дальнейшей работы является применение стохастических сетей '
  f'Петри для учёта вариативности времени срабатывания '
  f'сервоприводов и неровностей поверхности [10].')


# ══════════════ ЛИТЕРАТУРА ══════════════
H('Список литературы')

refs = [
    f'Siciliano{NBSP}B., Khatib{NBSP}O. (eds.) Springer Handbook of Robotics.{NBSP}{EMDASH} '
    f'2nd{NBSP}ed.{NBSP}{EMDASH} Springer, 2016.{NBSP}{EMDASH} 2228{NBSP}p.',

    f'Belter{NBSP}D., Skrzypczy\u0144ski{NBSP}P. A biologically inspired approach '
    f'to feasible gait learning for a hexapod robot // '
    f'Applied Mathematics and Computer Science.{NBSP}{EMDASH} 2010.{NBSP}{EMDASH} '
    f'Vol.{NBSP}20, No.{NBSP}1.{NBSP}{EMDASH} P.{NBSP}69{EMDASH}84.',

    f'Lee{NBSP}E.A., Seshia{NBSP}S.A. Introduction to Embedded Systems: '
    f'A Cyber-Physical Systems Approach.{NBSP}{EMDASH} 2nd{NBSP}ed.{NBSP}{EMDASH} '
    f'MIT Press, 2017.{NBSP}{EMDASH} 568{NBSP}p.',

    f'Hopcroft{NBSP}J.E., Motwani{NBSP}R., Ullman{NBSP}J.D. '
    f'Introduction to Automata Theory, Languages, and Computation.{NBSP}{EMDASH} '
    f'3rd{NBSP}ed.{NBSP}{EMDASH} Pearson, 2006.{NBSP}{EMDASH} 535{NBSP}p.',

    f'Petri{NBSP}C.A. Kommunikation mit Automaten: '
    f'Dissertation.{NBSP}{EMDASH} Universit\u00e4t Bonn, 1962.{NBSP}{EMDASH} 128{NBSP}S.',

    f'Murata{NBSP}T. Petri Nets: Properties, Analysis and Applications // '
    f'Proceedings of the IEEE.{NBSP}{EMDASH} 1989.{NBSP}{EMDASH} '
    f'Vol.{NBSP}77, No.{NBSP}4.{NBSP}{EMDASH} P.{NBSP}541{EMDASH}580.',

    f'Silva{NBSP}M., Teruel{NBSP}E., Colom{NBSP}J.M. Linear Algebraic and '
    f'Linear Programming Techniques for the Analysis of Place/Transition Net Systems // '
    f'Lectures on Petri Nets I: Basic Models.{NBSP}{EMDASH} Springer, 1998.{NBSP}{EMDASH} '
    f'P.{NBSP}309{EMDASH}373.',

    f'Ramchandani{NBSP}C. Analysis of Asynchronous Concurrent Systems by Timed '
    f'Petri Nets: Technical Report MAC-TR-120.{NBSP}{EMDASH} MIT, 1974.{NBSP}{EMDASH} '
    f'106{NBSP}p.',

    f'Спирин{NBSP}Н.А., Лавров{NBSP}В.В. Управление движением '
    f'многоногих шагающих роботов // Мехатроника, автоматизация, '
    f'управление.{NBSP}{EMDASH} 2019.{NBSP}{EMDASH} Т.{NBSP}20, '
    f'\u2116{NBSP}7.{NBSP}{EMDASH} С.{NBSP}430{EMDASH}438.',

    f'Marsan{NBSP}M.A., Balbo{NBSP}G., Conte{NBSP}G. et{NBSP}al. '
    f'Modelling with Generalized Stochastic Petri Nets.{NBSP}{EMDASH} '
    f'John Wiley & Sons, 1995.{NBSP}{EMDASH} 316{NBSP}p.',
]

for i, ref in enumerate(refs, 1):
    P(f'{i}. {ref}', indent=False)


# ══════════════ ENGLISH BLOCK ══════════════
P('', after=12)

P('FORMAL VERIFICATION OF HEXAPOD ROBOT\n'
  'GAIT COORDINATION USING PETRI NETS',
  bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, after=6)

P(f'V.P.{NBSP}Romanov',
  align=WD_ALIGN_PARAGRAPH.CENTER, indent=False)

P('romanov.vp@student.bmstu.ru',
  italic=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False)

P('Bauman Moscow State Technical University, Moscow, Russian Federation',
  italic=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=False, after=12)

RP([
    ('Abstract. ', True, False),
    ('The paper addresses formal verification of the tripod gait coordination '
     'algorithm for a hexapod walking robot using Petri nets. The robot\'s six '
     'legs are divided into two groups of three, alternating between swing and '
     'stance phases. A Petri net model is constructed that enforces static '
     'stability: at least three legs must maintain ground contact at all times. '
     'Using S-invariants, mutual exclusion of swing phases is formally proven '
     '(simultaneous lift-off of both groups is impossible), and the absence '
     'of deadlocks is demonstrated. Temporal analysis with deterministic delays '
     'determines the gait cycle time (600 ms), maximum locomotion speed '
     '(200 mm/s), and the bottleneck (swing phase accounts for 67% of cycle '
     'time). It is shown that replacing servos to reduce swing time by 25% '
     'yields a 20% speed improvement. The results are applicable to the design '
     'of autonomous legged robot control systems.', False, False),
], after=6)

RP([
    ('Keywords: ', True, False),
    ('Petri nets, walking robot, hexapod, tripod gait, formal verification, '
     'S-invariants, gait coordination, static stability, temporal analysis, '
     'mutual exclusion', False, False),
], after=12)


# ══════════════ СВЕДЕНИЯ ОБ АВТОРЕ ══════════════
H('Сведения об авторе')

P(f'Романов Владимир Павлович{NBSP}{EMDASH} студент кафедры '
  f'{LAQUO}Прикладная робототехника{RAQUO}, факультет '
  f'{LAQUO}Специальное машиностроение{RAQUO}, '
  f'МГТУ им.{NBSP}Н.Э.{NBSP}Баумана (Москва, Российская Федерация). '
  f'E-mail: romanov.vp@student.bmstu.ru',
  indent=False)

P('', after=6)

RP([
    (f'Научный руководитель: ', True, False),
    (f'Шевченко Максим Юрьевич, '
     f'МГТУ им.{NBSP}Н.Э.{NBSP}Баумана '
     f'(Москва, Российская Федерация).', False, False),
], indent=False, after=12)

H('About the author')

P(f'Romanov Vladimir P.{NBSP}{EMDASH} student, Department of Applied Robotics, '
  'Faculty of Special Mechanical Engineering, '
  'Bauman Moscow State Technical University (Moscow, Russian Federation). '
  'E-mail: romanov.vp@student.bmstu.ru',
  indent=False)

P('', after=6)

RP([
    ('Scientific advisor: ', True, False),
    (f'Shevchenko Maxim Yu., '
     f'Bauman Moscow State Technical University '
     f'(Moscow, Russian Federation).', False, False),
], indent=False)


# ── Нумерация страниц ──
add_page_numbers()

# ── Сохранение ──
out = os.path.join(os.path.dirname(__file__), 'Романов.docx')
doc.save(out)
print(f'Статья сохранена: {out}')
