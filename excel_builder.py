"""
excel_builder.py
────────────────
Генерация Excel-файла в фирменном стиле Alfa-Service через openpyxl.

ВАЖНО: пишет ЖИВЫЕ ФОРМУЛЫ Excel (=B5*C5), а не готовые числа. Ставка,
кол-во человек, смены — константы; всё производное (итого смен, ЗП на всех,
итоги по месяцу, итог за период) — формулы. Так пользователь может открыть
готовый файл, поправить ставку вручную — всё пересчитается автоматически.

Работает ТОЛЬКО на сервере (см. README.md §1 — почему клиентский JS
(SheetJS/ExcelJS) ненадёжен для стилизованного Excel).
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

VIOLET, LAVEND, OFFWHITE = '260A5B', 'A897FC', 'F7F5FF'
WHITE, LAVEND_L, BORDER_C = 'FFFFFF', 'EFEAFE', 'D9CFF7'
RUB_FMT, NUM_FMT = '#,##0" ₽"', '#,##0'

MNAME = ['', 'Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль',
         'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь']
MNAME_UP = [m.upper() for m in MNAME]

BUDGET_LABELS = {
    'chem': 'Расходные материалы (химия и РМ)', 'inv': 'Инвентарь',
    'uniform': 'Спецодежда', 'gsm': 'ГСМ',
    'gsm_summer': 'ГСМ — летний период', 'gsm_winter': 'ГСМ — зимний период',
    'reagents': 'Реагенты (ПГМ)',
    'maint': 'Эксплуатация и ремонт оборудования', 'sout': 'Расходы на СОУТ',
}


def _fl(h): return PatternFill('solid', fgColor=h)
def _fn(bold=False, color=None, size=10, italic=False):
    return Font(name='Arial', bold=bold, color=color or VIOLET, size=size, italic=italic)
def _al(h='left', v='center', wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)
def _bdr(color=BORDER_C):
    s = Side(style='thin', color=color)
    return Border(left=s, right=s, top=s, bottom=s)
def _bdr_bold(color=LAVEND):
    thick, thin = Side(style='medium', color=color), Side(style='thin', color=BORDER_C)
    return Border(left=thin, right=thin, top=thick, bottom=thick)


def _write(ws, r, c, val='', font=None, fill=None, align=None, border=None, nf=None):
    """ВАЖНО: для объединяемых ячеек сначала _write(), потом merge_cells() —
    запись в уже объединённую (не top-left) ячейку кидает AttributeError."""
    cell = ws.cell(row=r, column=c, value=val)
    if font: cell.font = font
    if fill: cell.fill = fill
    if align: cell.alignment = align
    if border: cell.border = border
    if nf: cell.number_format = nf
    return cell


def _merge(ws, r1, c1, r2, c2):
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)


def _fill_row(ws, r, c1, c2, fill_obj, border_obj=None):
    for c in range(c1, c2 + 1):
        ws.cell(row=r, column=c).fill = fill_obj
        if border_obj:
            ws.cell(row=r, column=c).border = border_obj


def shift_sum(staff):
    return sum(s['tot_shifts'] for s in staff if not s['is_oklad'])


# ═══════════════════════════════════════════════════════════
# ЛИСТ 1 — БЮДЖЕТ ФОТ
# ═══════════════════════════════════════════════════════════
def _build_fot_sheet(wb, obj_name, obj_addr, period, months):
    ws = wb.active
    ws.title = 'Бюджет ФОТ'
    NC = 10
    for i, w in enumerate([3, 32, 9, 16, 11, 14, 16, 11, 12, 18], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Переключатель Фикс/Смена: {должность: номер_строки_первого_появления}.
    # Выпадающий список стоит только там; остальные месяцы этой же должности
    # ссылаются формулой на эту ячейку — так тип меняется один раз и
    # пересчитывает все месяцы контракта для данной должности автоматически.
    first_type_cell: dict[str, int] = {}
    type_dv = DataValidation(type="list", formula1='"Фикс,Смена"', allow_blank=False)
    ws.add_data_validation(type_dv)

    r = 1
    ws.row_dimensions[r].height = 52
    _write(ws, r, 1, 'alfa', font=Font(name='Arial', bold=True, color=LAVEND, size=30),
           fill=_fl(VIOLET), align=_al('left', 'center'))
    _merge(ws, r, 1, r, NC); _fill_row(ws, r, 2, NC, _fl(VIOLET))

    r += 1; ws.row_dimensions[r].height = 20
    _write(ws, r, 1, 'БЮДЖЕТ НА ФОНД ОПЛАТЫ ТРУДА', font=_fn(False, WHITE, 11), fill=_fl(VIOLET), align=_al())
    _merge(ws, r, 1, r, NC); _fill_row(ws, r, 2, NC, _fl(VIOLET))

    r += 1; ws.row_dimensions[r].height = 16
    _write(ws, r, 1, f"{obj_name}   |   {period}", font=_fn(False, LAVEND, 9), fill=_fl(VIOLET), align=_al())
    _merge(ws, r, 1, r, NC); _fill_row(ws, r, 2, NC, _fl(VIOLET))

    if obj_addr:
        r += 1; ws.row_dimensions[r].height = 14
        _write(ws, r, 1, obj_addr, font=_fn(False, LAVEND, 9), fill=_fl(VIOLET), align=_al())
        _merge(ws, r, 1, r, NC); _fill_row(ws, r, 2, NC, _fl(VIOLET))

    r += 1; ws.row_dimensions[r].height = 6
    _fill_row(ws, r, 1, NC, _fl(LAVEND)); _merge(ws, r, 1, r, NC)
    r += 1; ws.row_dimensions[r].height = 10
    _fill_row(ws, r, 1, NC, _fl(OFFWHITE)); _merge(ws, r, 1, r, NC)

    # ── Сводная таблица ──
    r += 1; ws.row_dimensions[r].height = 26
    _write(ws, r, 1, 'СВОДНЫЙ БЮДЖЕТ ФОТ ПО МЕСЯЦАМ', font=_fn(True, VIOLET, 11),
           fill=_fl(LAVEND), align=_al(), border=_bdr())
    _merge(ws, r, 1, r, NC); _fill_row(ws, r, 2, NC, _fl(LAVEND), _bdr())

    r += 1; ws.row_dimensions[r].height = 30
    for ci, h in enumerate(['№', 'Месяц', 'Дней', 'Должн.', 'Итого смен', '', '', '', 'ФОТ в месяц', ''], 1):
        a = 'left' if ci == 2 else 'center'
        _write(ws, r, ci, h, font=_fn(True, VIOLET, 9), fill=_fl(LAVEND_L),
               align=_al(a, 'center', True), border=_bdr())
    _merge(ws, r, 9, r, NC)

    summary_start = r + 1
    summary_refs = []
    for i, m in enumerate(months):
        r += 1; ws.row_dimensions[r].height = 20
        bg = OFFWHITE if i % 2 == 0 else WHITE
        _write(ws, r, 1, i + 1, font=_fn(False, VIOLET, 8), fill=_fl(bg), align=_al('center'), border=_bdr())
        _write(ws, r, 2, f"{MNAME[m['mn']]} {m['yr']}", font=_fn(True, VIOLET, 10), fill=_fl(bg), align=_al(), border=_bdr())
        _write(ws, r, 3, m['days'], font=_fn(False, VIOLET, 10), fill=_fl(bg), align=_al('center'), border=_bdr())
        _write(ws, r, 4, len(m['staff']), font=_fn(False, VIOLET, 10), fill=_fl(bg), align=_al('center'), border=_bdr())
        _write(ws, r, 5, 0, font=_fn(False, VIOLET, 10), fill=_fl(bg), align=_al('center'), border=_bdr(), nf=NUM_FMT)
        for c in range(6, 9):
            _write(ws, r, c, '', fill=_fl(bg), border=_bdr())
        _write(ws, r, 9, 0, font=_fn(True, VIOLET, 12), fill=_fl(bg), align=_al('center'), border=_bdr(), nf=RUB_FMT)
        _write(ws, r, 10, '', fill=_fl(bg), border=_bdr()); _merge(ws, r, 9, r, NC)
        summary_refs.append(r)
    summary_end = r

    r += 1; ws.row_dimensions[r].height = 26
    _write(ws, r, 1, '', fill=_fl(LAVEND), border=_bdr())
    _write(ws, r, 2, 'ИТОГО ЗА ПЕРИОД', font=_fn(True, VIOLET, 10), fill=_fl(LAVEND), align=_al(), border=_bdr())
    _merge(ws, r, 2, r, 4); _fill_row(ws, r, 3, 4, _fl(LAVEND), _bdr())
    _write(ws, r, 5, f'=SUM(E{summary_start}:E{summary_end})', font=_fn(True, VIOLET, 10), fill=_fl(LAVEND), align=_al('center'), border=_bdr(), nf=NUM_FMT)
    for c in range(6, 9):
        _write(ws, r, c, '', fill=_fl(LAVEND), border=_bdr())
    _write(ws, r, 9, f'=SUM(I{summary_start}:I{summary_end})', font=_fn(True, VIOLET, 14), fill=_fl(LAVEND), align=_al('center'), border=_bdr(), nf=RUB_FMT)
    _write(ws, r, 10, '', fill=_fl(LAVEND), border=_bdr()); _merge(ws, r, 9, r, NC)
    grand_total_row = r

    r += 1; ws.row_dimensions[r].height = 10
    _fill_row(ws, r, 1, NC, _fl(OFFWHITE)); _merge(ws, r, 1, r, NC)

    # ── Помесячная детализация ──
    for mi, m in enumerate(months):
        r += 1; ws.row_dimensions[r].height = 26
        if m.get('full_days') and m['days'] < m['full_days']:
            date_note = f"с {m['d1']} по {m['d2']} число  ·  {m['days']} дн."
        else:
            dw = 'дней' if m['days'] >= 5 else ('день' if m['days'] == 1 else 'дня')
            date_note = f"{m['days']} {dw}"
        # Первый месяц — дополнительно явно указываем дату начала оказания
        # услуг (не только диапазон дней внутри месяца, а полную дату).
        if mi == 0 and 'd1' in m:
            start_date_label = f"  ·  дата начала услуг: {m['d1']:02d}.{m['mn']:02d}.{m['yr']}"
        else:
            start_date_label = ""
        _write(ws, r, 1, f"{MNAME_UP[m['mn']]} {m['yr']}  ·  {date_note}{start_date_label}",
               font=_fn(True, VIOLET, 13), fill=_fl(LAVEND), align=_al())
        _merge(ws, r, 1, r, NC); _fill_row(ws, r, 2, NC, _fl(LAVEND))

        r += 1; ws.row_dimensions[r].height = 36
        headers = [('№', 'c'), ('Должность', 'l'), ('Кол-во', 'c'), ('График', 'c'), ('Тип', 'c'),
                   ('Ставка 1 смена', 'r'), ('ЗП 1 сотр. на руки', 'r'), ('Смен 1 сотр.', 'c'),
                   ('Итого смен', 'c'), ('ЗП на всех', 'r')]
        for ci, (h, a) in enumerate(headers, 1):
            align_h = 'right' if a == 'r' else ('left' if a == 'l' else 'center')
            _write(ws, r, ci, h, font=_fn(True, VIOLET, 9), fill=_fl(LAVEND_L),
                   align=_al(align_h, 'center', True), border=_bdr())

        first_data = r + 1
        if not m['staff']:
            r += 1; ws.row_dimensions[r].height = 18
            _write(ws, r, 1, '—', font=_fn(False, '94a3b8', 9), fill=_fl(OFFWHITE), align=_al('center'), border=_bdr())
            _merge(ws, r, 1, r, 10); _fill_row(ws, r, 2, 10, _fl(OFFWHITE), _bdr())

        for si, s in enumerate(m['staff']):
            r += 1; ws.row_dimensions[r].height = 20
            bg = OFFWHITE if si % 2 == 0 else WHITE
            is_oklad = s['is_oklad']

            _write(ws, r, 1, si + 1, font=_fn(False, VIOLET, 8), fill=_fl(bg), align=_al('center'), border=_bdr())
            _write(ws, r, 2, s['pos'], font=_fn(True, VIOLET, 10), fill=_fl(bg), align=_al(), border=_bdr())
            _write(ws, r, 3, s['qty'], font=_fn(False, VIOLET, 10), fill=_fl(bg), align=_al('center'), border=_bdr())
            _write(ws, r, 4, s['sched'], font=_fn(True, VIOLET, 10, italic=True), fill=_fl(bg), align=_al('center'), border=_bdr())

            typ = 'Фикс' if is_oklad else 'Смена'
            typ_bg, typ_fg = ('F0DDB8', '7A4A00') if is_oklad else ('D6EBD9', '1F5C2E')

            # ── Переключатель Фикс/Смена: живой выпадающий список только
            # в ПЕРВОМ месяце, где встретилась эта должность; во всех
            # последующих месяцах — формула-ссылка на первую ячейку.
            # Меняя тип в первом месяце, пользователь пересчитывает все
            # остальные месяцы этой же должности автоматически.
            if s['pos'] not in first_type_cell:
                _write(ws, r, 5, typ, font=_fn(True, typ_fg, 10), fill=_fl(typ_bg), align=_al('center'), border=_bdr())
                type_dv.add(ws.cell(row=r, column=5))
                first_type_cell[s['pos']] = r
            else:
                ref_r = first_type_cell[s['pos']]
                _write(ws, r, 5, f'=E{ref_r}', font=_fn(True, typ_fg, 10, italic=True), fill=_fl(typ_bg), align=_al('center'), border=_bdr())

            if is_oklad:
                shifts_ref = s.get('shifts_ref', '—')
                _write(ws, r, 6, '—', font=_fn(False, '94a3b8', 9), fill=_fl(bg), align=_al('center'), border=_bdr())
                _write(ws, r, 7, round(s['sal']), font=_fn(True, VIOLET, 10), fill=_fl(bg), align=_al('right'), border=_bdr(), nf=RUB_FMT)
                _write(ws, r, 8, shifts_ref, font=_fn(False, '64748b' if shifts_ref != '—' else '94a3b8', 9, italic=(shifts_ref != '—')), fill=_fl(bg), align=_al('center'), border=_bdr())
                _write(ws, r, 9, '—', font=_fn(False, '94a3b8', 9), fill=_fl(bg), align=_al('center'), border=_bdr())
                # ЗП на всех — через IF: если тип переключат на "Смена",
                # формула сама перейдёт на посменный расчёт (Ставка×Итого смен).
                _write(ws, r, 10, f'=IF(E{r}="Фикс",G{r}*C{r},F{r}*I{r})', font=_fn(True, VIOLET, 11), fill=_fl(bg), align=_al('right'), border=_bdr(), nf=RUB_FMT)
            else:
                _write(ws, r, 6, round(s['rate']), font=_fn(False, VIOLET, 10), fill=_fl(bg), align=_al('right'), border=_bdr(), nf=NUM_FMT)
                _write(ws, r, 7, f'=F{r}*H{r}', font=_fn(True, VIOLET, 10), fill=_fl(bg), align=_al('right'), border=_bdr(), nf=RUB_FMT)
                _write(ws, r, 8, s['shifts'], font=_fn(False, VIOLET, 10), fill=_fl(bg), align=_al('center'), border=_bdr())
                _write(ws, r, 9, f'=C{r}*H{r}', font=_fn(True, VIOLET, 10), fill=_fl(bg), align=_al('center'), border=_bdr(), nf=NUM_FMT)
                _write(ws, r, 10, f'=IF(E{r}="Фикс",G{r}*C{r},F{r}*I{r})', font=_fn(True, VIOLET, 11), fill=_fl(bg), align=_al('right'), border=_bdr(), nf=RUB_FMT)

        last_data = r
        r += 1; ws.row_dimensions[r].height = 24
        _write(ws, r, 1, '', fill=_fl(LAVEND), border=_bdr_bold())
        _write(ws, r, 2, f"ИТОГО  {MNAME[m['mn']]} {m['yr']}", font=_fn(True, VIOLET, 10), fill=_fl(LAVEND), align=_al(), border=_bdr_bold())
        _merge(ws, r, 2, r, 8); _fill_row(ws, r, 3, 8, _fl(LAVEND), _bdr_bold())
        if m['staff']:
            _write(ws, r, 9, f'=SUM(I{first_data}:I{last_data})', font=_fn(True, VIOLET, 10), fill=_fl(LAVEND), align=_al('center'), border=_bdr_bold(), nf=NUM_FMT)
            _write(ws, r, 10, f'=SUM(J{first_data}:J{last_data})', font=_fn(True, VIOLET, 13), fill=_fl(LAVEND), align=_al('right'), border=_bdr_bold(), nf=RUB_FMT)
        else:
            _write(ws, r, 9, 0, font=_fn(True, VIOLET, 10), fill=_fl(LAVEND), align=_al('center'), border=_bdr_bold(), nf=NUM_FMT)
            _write(ws, r, 10, 0, font=_fn(True, VIOLET, 13), fill=_fl(LAVEND), align=_al('right'), border=_bdr_bold(), nf=RUB_FMT)
        month_total_row = r

        r += 1; ws.row_dimensions[r].height = 8
        _fill_row(ws, r, 1, NC, _fl(OFFWHITE)); _merge(ws, r, 1, r, NC)

        sref = summary_refs[mi]
        ws.cell(row=sref, column=5).value = f'=I{month_total_row}'
        ws.cell(row=sref, column=9).value = f'=J{month_total_row}'

    r += 1; ws.row_dimensions[r].height = 6
    _fill_row(ws, r, 1, NC, _fl(LAVEND)); _merge(ws, r, 1, r, NC)
    r += 1; ws.row_dimensions[r].height = 36
    _write(ws, r, 1, '', fill=_fl(VIOLET))
    _write(ws, r, 2, 'ИТОГО БЮДЖЕТ ФОТ ЗА ПЕРИОД', font=_fn(True, LAVEND, 11), fill=_fl(VIOLET), align=_al())
    _merge(ws, r, 2, r, 8); _fill_row(ws, r, 3, 8, _fl(VIOLET))
    _write(ws, r, 9, f'=E{grand_total_row}', font=_fn(True, LAVEND, 11), fill=_fl(VIOLET), align=_al('center'), nf=NUM_FMT)
    _write(ws, r, 10, f'=I{grand_total_row}', font=_fn(True, LAVEND, 16), fill=_fl(VIOLET), align=_al('right'), nf=RUB_FMT)
    r += 1; ws.row_dimensions[r].height = 6
    _fill_row(ws, r, 1, NC, _fl(LAVEND)); _merge(ws, r, 1, r, NC)

    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = 9
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitPage = True


# ═══════════════════════════════════════════════════════════
# ЛИСТ 2 — РАСХОДЫ И МАТЕРИАЛЫ
# ═══════════════════════════════════════════════════════════
def _build_budget_sheet(wb, obj_name, budget, equipment, n_months, discount=None, budget_multipliers=None):
    """
    discount: опционально float 0..1 (например 0.8 для скидки −20%) —
    применяется ТОЛЬКО к категориям расходников (chem/inv/uniform/gsm/
    maint/sout), НЕ применяется к перечню техники/оборудования.
    """
    ws = wb.create_sheet('Расходы и материалы')
    NC = 4
    for i, w in enumerate([38, 14, 22, 22], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    r = 1
    ws.row_dimensions[r].height = 52
    _write(ws, r, 1, 'alfa', font=Font(name='Arial', bold=True, color=LAVEND, size=30),
           fill=_fl(VIOLET), align=_al())
    _merge(ws, r, 1, r, NC); _fill_row(ws, r, 2, NC, _fl(VIOLET))

    r += 1; ws.row_dimensions[r].height = 20
    _write(ws, r, 1, 'БЮДЖЕТ РАСХОДОВ И МАТЕРИАЛОВ', font=_fn(False, WHITE, 11), fill=_fl(VIOLET), align=_al())
    _merge(ws, r, 1, r, NC); _fill_row(ws, r, 2, NC, _fl(VIOLET))

    r += 1; ws.row_dimensions[r].height = 20
    note = 'БЕЗ НДС' + (f' · СО СКИДКОЙ {int((1 - discount) * 100)}%' if discount else '')
    _write(ws, r, 1, obj_name, font=_fn(False, LAVEND, 9), fill=_fl(VIOLET), align=_al())
    _write(ws, r, 2, '', fill=_fl(VIOLET))
    _write(ws, r, 3, note, font=_fn(True, LAVEND, 10), fill=_fl(VIOLET), align=_al('right'))
    _write(ws, r, 4, '', fill=_fl(VIOLET)); _merge(ws, r, 3, r, NC)

    r += 1; ws.row_dimensions[r].height = 6
    _fill_row(ws, r, 1, NC, _fl(LAVEND)); _merge(ws, r, 1, r, NC)
    r += 1; ws.row_dimensions[r].height = 10
    _fill_row(ws, r, 1, NC, _fl(OFFWHITE)); _merge(ws, r, 1, r, NC)

    r += 1; ws.row_dimensions[r].height = 26
    title = 'РАСХОДНЫЕ МАТЕРИАЛЫ И ПРОЧИЕ ЗАТРАТЫ' + (' (СО СКИДКОЙ)' if discount else '')
    _write(ws, r, 1, title, font=_fn(True, VIOLET, 11), fill=_fl(LAVEND), align=_al(), border=_bdr())
    _merge(ws, r, 1, r, NC); _fill_row(ws, r, 2, NC, _fl(LAVEND), _bdr())

    r += 1; ws.row_dimensions[r].height = 36
    for ci, (h, a) in enumerate([
        ('Наименование статьи затрат', 'l'), ('', 'c'),
        ('Бюджет в месяц без НДС, руб.', 'c'),
        (f'Бюджет за период ({n_months} мес.) без НДС, руб.', 'c'),
    ], 1):
        _write(ws, r, ci, h, font=_fn(True, VIOLET, 9), fill=_fl(LAVEND_L),
               align=_al('left' if a == 'l' else 'center', 'center', True), border=_bdr())
    _merge(ws, r, 1, r, 2)

    budget_multipliers = budget_multipliers or {}
    category_keys = ['chem', 'inv', 'uniform', 'gsm', 'gsm_summer', 'gsm_winter', 'reagents', 'maint', 'sout']
    for idx, key in enumerate(category_keys):
        val_orig = budget.get(key, 0)
        if val_orig <= 0:
            continue
        val = val_orig * discount if discount else val_orig
        mult = budget_multipliers.get(key, n_months)  # своё кол-во месяцев для сезонных статей
        r += 1; ws.row_dimensions[r].height = 24
        idx_row = r
        bg = OFFWHITE if (r % 2 == 0) else WHITE
        label = BUDGET_LABELS.get(key, key)
        if key in budget_multipliers:
            label += f' ({mult} мес.)'
        _write(ws, r, 1, label, font=_fn(True, VIOLET, 11), fill=_fl(bg), align=_al(), border=_bdr())
        _write(ws, r, 2, '', fill=_fl(bg), border=_bdr()); _merge(ws, r, 1, r, 2)
        _write(ws, r, 3, round(val), font=_fn(True, VIOLET, 13), fill=_fl(bg), align=_al('right'), border=_bdr(), nf=RUB_FMT)
        _write(ws, r, 4, f'=C{idx_row}*{mult}', font=_fn(True, VIOLET, 13), fill=_fl(bg), align=_al('right'), border=_bdr(), nf=RUB_FMT)

    r += 1; ws.row_dimensions[r].height = 10
    _fill_row(ws, r, 1, NC, _fl(OFFWHITE)); _merge(ws, r, 1, r, NC)
    r += 1; ws.row_dimensions[r].height = 6
    _fill_row(ws, r, 1, NC, _fl(LAVEND)); _merge(ws, r, 1, r, NC)

    if equipment:
        r += 1; ws.row_dimensions[r].height = 10
        _fill_row(ws, r, 1, NC, _fl(OFFWHITE)); _merge(ws, r, 1, r, NC)

        r += 1; ws.row_dimensions[r].height = 24
        _write(ws, r, 1, 'ПЕРЕЧЕНЬ ТЕХНИКИ И ОБОРУДОВАНИЯ', font=_fn(True, VIOLET, 11), fill=_fl(LAVEND), align=_al(), border=_bdr())
        _merge(ws, r, 1, r, NC); _fill_row(ws, r, 2, NC, _fl(LAVEND), _bdr())

        r += 1; ws.row_dimensions[r].height = 32
        for ci, h in enumerate(['Наименование', 'Кол-во', 'Аморт./мес, руб.', f'За период ({n_months} мес.), руб.'], 1):
            a = 'left' if ci == 1 else 'center'
            _write(ws, r, ci, h, font=_fn(True, VIOLET, 9), fill=_fl(LAVEND_L), align=_al(a, 'center', True), border=_bdr())

        eq_rows = []
        for ei, e in enumerate(equipment):
            bg = OFFWHITE if ei % 2 == 0 else WHITE
            val_month = round(e['val'])
            name_display = e['name'].replace(', аренда', '').replace(',аренда', '')
            r += 1; ws.row_dimensions[r].height = 20
            _write(ws, r, 1, name_display, font=_fn(True, VIOLET, 10), fill=_fl(bg), align=_al(), border=_bdr())
            _write(ws, r, 2, int(e['qty']) if e['qty'] else 1, font=_fn(False, VIOLET, 10), fill=_fl(bg), align=_al('center'), border=_bdr())
            _write(ws, r, 3, val_month, font=_fn(True, VIOLET, 10), fill=_fl(bg), align=_al('right'), border=_bdr(), nf=RUB_FMT)
            _write(ws, r, 4, f'=C{r}*{n_months}', font=_fn(False, VIOLET, 10), fill=_fl(bg), align=_al('right'), border=_bdr(), nf=RUB_FMT)
            eq_rows.append(r)

            if e.get('is_rent'):
                r += 1; ws.row_dimensions[r].height = 15
                _write(ws, r, 1, '   ● АРЕНДА', font=_fn(True, 'B45309', 8), fill=_fl(bg), border=_bdr())
                for c in range(2, 5):
                    _write(ws, r, c, '', fill=_fl(bg), border=_bdr())

        r += 1; ws.row_dimensions[r].height = 26
        _write(ws, r, 1, 'ИТОГО АМОРТИЗАЦИЯ / АРЕНДА ОБОРУДОВАНИЯ', font=_fn(True, VIOLET, 11), fill=_fl(LAVEND), align=_al(), border=_bdr_bold())
        _write(ws, r, 2, '', fill=_fl(LAVEND), border=_bdr_bold()); _merge(ws, r, 1, r, 2)
        eq_sum_c = '+'.join(f'C{er}' for er in eq_rows)
        eq_sum_d = '+'.join(f'D{er}' for er in eq_rows)
        _write(ws, r, 3, f'={eq_sum_c}', font=_fn(True, VIOLET, 13), fill=_fl(LAVEND), align=_al('right'), border=_bdr_bold(), nf=RUB_FMT)
        _write(ws, r, 4, f'={eq_sum_d}', font=_fn(True, VIOLET, 13), fill=_fl(LAVEND), align=_al('right'), border=_bdr_bold(), nf=RUB_FMT)

        r += 1; ws.row_dimensions[r].height = 6
        _fill_row(ws, r, 1, NC, _fl(LAVEND)); _merge(ws, r, 1, r, NC)

    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = 9
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitPage = True


# ═══════════════════════════════════════════════════════════
# ТОЧКА ВХОДА
# ═══════════════════════════════════════════════════════════
def build_workbook(obj_name, obj_addr, period, months, budget, equipment=None, discount=None, budget_multipliers=None):
    """
    discount: если указан (например 0.8 для скидки −20%), применяется
    к расходным материалам на Листе 2. НЕ применяется к ФОТ (Лист 1)
    и НЕ применяется к перечню техники/аренды.

    budget_multipliers: опционально {'gsm_summer': 7, 'gsm_winter': 5, ...} —
    для сезонных категорий, у которых «итого за период» считается не по
    общему n_months контракта, а по собственному числу месяцев сезона.
    """
    wb = Workbook()
    n_months = len(months)
    _build_fot_sheet(wb, obj_name, obj_addr, period, months)
    _build_budget_sheet(wb, obj_name, budget, equipment or [], n_months, discount=discount, budget_multipliers=budget_multipliers)
    return wb
