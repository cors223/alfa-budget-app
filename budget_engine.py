"""
budget_engine.py
────────────────
Парсинг штатных файлов Альфа-Сервис (.xlsx) и расчёт помесячного бюджета ФОТ
+ расходных материалов. Вся бизнес-логика — здесь, без какого-либо UI.

Использование:
    from budget_engine import parse_workbook, build_months

    parsed = parse_workbook('расчет.xlsx')
    # parsed = {
    #   'obj_name': str, 'obj_addr': str,
    #   'staff': [...],       # список должностей (уже объединены 6/1+1/6 пары)
    #   'budget': {...},      # расходные материалы по категориям
    #   'contract_months': int,
    #   'start_month': int, 'start_year': int,
    # }

    months = build_months(parsed, start_month=7, start_year=2026, n_months=18)
    # months = [{'yr':.., 'mn':.., 'days':.., 'staff':[...], 'total': ..}, ...]
"""

import re
import calendar
from openpyxl import load_workbook
from production_calendar import workdays_5x2, workdays_5x2_range


# ═══════════════════════════════════════════════════════════
# КАЛЕНДАРНЫЕ ФУНКЦИИ
# ═══════════════════════════════════════════════════════════

def days_in(year: int, month: int) -> int:
    return calendar.monthrange(year, month)[1]


def workdays_in(year: int, month: int) -> int:
    """
    Рабочие дни по ОФИЦИАЛЬНОМУ производственному календарю РФ (график 5/2),
    с учётом праздников — не просто «пн–пт по календарю». См.
    production_calendar.py: для года без подтверждённых данных используется
    приближённый подсчёт (пн–пт), с занижением точности в праздничных
    месяцах (январь, февраль, май, июнь, ноябрь).
    """
    n, _is_official = workdays_5x2(year, month)
    return n


def fri_sat_sun_in(year: int, month: int) -> int:
    """Пятница + суббота + воскресенье в месяце (график 3/4)."""
    return sum(
        1 for d in range(1, days_in(year, month) + 1)
        if calendar.weekday(year, month, d) >= 4
    )


def calc_shifts(schedule: str, year: int, month: int) -> int | None:
    """
    Количество смен/дней на 1 сотрудника в месяц по графику работы.
    Возвращает None для оклада (график не влияет на расчёт).

    ВАЖНО: 6/1 и 1/6 сами по себе (без пары) считаются ПРОПОРЦИОНАЛЬНО,
    а не как "все дни" — "все дни" получается только при объединении
    пары 6/1+1/6 в единый график 7х0 (см. merge_shift_pairs).
    """
    if not schedule:
        return None

    # убираем время из строки вида "6х1 08:00-20:00"
    s = re.sub(r'\s*\d{1,2}[:.]\d{2}[-–]\d{1,2}[:.]\d{2}', '', schedule).strip().lower()
    if not s or 'оклад' in s:
        return None

    d = days_in(year, month)

    if re.match(r'^7[/хx]0', s):
        return d
    if re.match(r'^5[/хx]2', s):
        return workdays_in(year, month)
    if re.match(r'^6[/хx]1', s):
        return round(d * 6 / 7)
    if re.match(r'^1[/хx]6', s):
        return round(d * 1 / 7)
    if re.match(r'^3[/хx]4', s):
        return fri_sat_sun_in(year, month)
    if re.match(r'^2[/хx]2', s):
        return round(d / 2)
    if re.search(r'разъ|выезд', s):
        return workdays_in(year, month)

    # общая формула N/M
    m = re.match(r'(\d+)[/хx](\d+)', s)
    if m:
        on, off = int(m.group(1)), int(m.group(2))
        return round(d * on / (on + off))

    return d  # неопознанный график — считаем как "каждый день"


# ═══════════════════════════════════════════════════════════
# СЕЗОННАЯ АКТИВНОСТЬ ДОЛЖНОСТЕЙ
# ═══════════════════════════════════════════════════════════

SUMMER_8 = {3, 4, 5, 6, 7, 8, 9, 10}      # март–октябрь
SUMMER_7 = {4, 5, 6, 7, 8, 9, 10}          # апрель–октябрь
WINTER_4 = {11, 12, 1, 2}                  # ноябрь–февраль
WINTER_5 = {11, 12, 1, 2, 3}               # ноябрь–март


def is_position_active(pos_name: str, contract_months: int, month: int, section_season: str | None = None) -> bool:
    """
    Активна ли должность в данном календарном месяце (сезонность).

    section_season: если задан ('лето'/'зима') — берётся из явного
    секционного маркера в исходном файле (см. _parse_raw_staff) и имеет
    ПРИОРИТЕТ над определением по названию должности. Нужен для файлов
    типа ЦСКА Автозаводская, где одна и та же должность (например,
    «Старший менеджер Объекта») может встречаться в обеих секциях без
    сезонных слов в самом названии.
    """
    if contract_months == 12:
        return True
    if section_season == 'лето':
        return month in (SUMMER_8 if contract_months >= 8 else SUMMER_7)
    if section_season == 'зима':
        return month in (WINTER_4 if contract_months <= 4 else WINTER_5)
    name = pos_name.lower()
    is_summer = 'лет' in name
    is_winter = any(k in name for k in ('зим', 'трактор', 'снег'))
    if is_summer:
        return month in (SUMMER_8 if contract_months >= 8 else SUMMER_7)
    if is_winter:
        return month in (WINTER_4 if contract_months <= 4 else WINTER_5)
    return True  # нет признака сезонности — считаем активной весь срок


def is_oklad_by_name(pos_name: str) -> bool:
    """Автоправило: менеджер/директор/руководитель по умолчанию на окладе."""
    return bool(re.search(r'менеджер|управля|директор|руковод', pos_name, re.I))


# ═══════════════════════════════════════════════════════════
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ЧТЕНИЯ ЯЧЕЕК
# ═══════════════════════════════════════════════════════════

def _cn(ws, r: int, c: int) -> float:
    """Число из ячейки (r,c), с запасным парсингом строк типа '1 234,5'."""
    v = ws.cell(row=r, column=c).value
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v or '').replace(' ', '').replace(',', '.'))
    except ValueError:
        return 0.0


def _cs(ws, r: int, c: int) -> str:
    v = ws.cell(row=r, column=c).value
    return str(v).strip() if v is not None else ''


# ═══════════════════════════════════════════════════════════
# ПОИСК СТРОКИ ЗАГОЛОВКА И КОЛОНОК (динамически, не хардкод!)
# ═══════════════════════════════════════════════════════════

def detect_columns(ws) -> dict:
    """
    Ищет строку с заголовком «Кол-во, чел.» и определяет по соседним
    заголовкам номера нужных столбцов. Работает для файлов, где колонки
    смещены относительно "типичного" расположения.
    """
    cols = {
        'hdr_row': 4, 'col_pos': 2, 'col_qty': 3, 'col_sched': 4,
        'col_shifts_month': 7, 'col_months': 8, 'col_sal_month': 10,
        'col_rate_shift': 11,
    }
    for r in range(1, 13):
        for c in range(1, 9):
            v = _cs(ws, r, c).lower().replace('-', '').replace(' ', '')
            if 'кол' in v and 'чел' in v:
                cols['hdr_row'] = r
                cols['col_qty'] = c
                cols['col_pos'] = max(1, c - 1)
                for c2 in range(c + 1, c + 15):
                    h = _cs(ws, r, c2).lower().strip()
                    if re.search(r'график|время работы', h):
                        cols['col_sched'] = c2
                    if re.search(r'кол.{0,8}смен.{0,8}мес|рабочих смен в мес', h):
                        cols['col_shifts_month'] = c2
                    if re.search(r'кол.{0,8}месяц', h) and not re.search(r'смен|час|нетто', h):
                        cols['col_months'] = c2
                    if h == 'в месяц':
                        cols['col_sal_month'] = c2
                    if re.search(r'в смену|ставка', h) and 'час' not in h:
                        cols['col_rate_shift'] = c2
                return cols
    return cols


def get_meta(wb, filename: str) -> dict:
    """Название заказчика и адрес из листа Калькуляция + фолбэк на имя файла."""
    obj_name, obj_addr = '', ''
    kn = next((n for n in wb.sheetnames if re.search(r'калькул', n, re.I)), None)
    if kn:
        ws = wb[kn]
        for r in range(1, 11):
            label = _cs(ws, r, 2).lower()
            val = _cs(ws, r, 3) or _cs(ws, r, 4) or _cs(ws, r, 5)
            # ВАЖНО: строго "заказчик", не "объект" — иначе поймает
            # строку "Вид объекта" (тип недвижимости) по ошибке.
            if re.search(r'заказчик', label) and val:
                obj_name = obj_name or val
            if re.search(r'адрес', label) and val:
                obj_addr = obj_addr or val

    if not obj_name and filename:
        obj_name = re.sub(r'\.xlsx?$', '', filename, flags=re.I)
        obj_name = re.sub(r'[_\-]+', ' ', obj_name)  # сначала разбить на слова
        obj_name = re.sub(r'\bрасчет\b|\bкалькуляция\b|\bиндексация\b|\bпериод\b|\bна\b|\bс\b|\bдо\b', ' ', obj_name, flags=re.I)
        obj_name = re.sub(r'\d{1,2}[_.]\d{1,2}[_.]\d{4}', ' ', obj_name)
        obj_name = re.sub(r'\b\d+\b', ' ', obj_name)
        obj_name = re.sub(r'\s+', ' ', obj_name).strip()

    return {'obj_name': obj_name, 'obj_addr': obj_addr}


# ═══════════════════════════════════════════════════════════
# ПАРСИНГ ШТАТА (оба формата: "Штат" и "Расчет ФОТ")
# ═══════════════════════════════════════════════════════════

SKIP_POS_PREFIXES = [
    'расходы на оплату', 'зарплата', 'налоги', 'ндфл', 'взнос',
    'табелир', 'должность', 'наименование',
]
SKIP_POS_CONTAINS = ['итого кол', 'итого смен', 'кол.-во']


def _parse_raw_staff(ws, cols: dict) -> list[dict]:
    """
    Читает сырые строки должностей (ещё БЕЗ объединения пар 6/1+1/6).

    Дополнительно отслеживает секционные маркеры «ЛЕТО» / «ЗИМА» —
    строки без количества человек, которые в некоторых файлах (например,
    ЦСКА Автозаводская) явно размечают, что все последующие строки до
    следующего маркера относятся к летнему/зимнему периоду. Если такие
    маркеры есть — сезон берётся из ближайшего маркера (`section_season`),
    а не только из названия должности (см. is_position_active).
    """
    raw = []
    max_row = ws.max_row
    current_section_season = None

    for r in range(cols['hdr_row'] + 1, max_row + 1):
        pos = _cs(ws, r, cols['col_pos']) or _cs(ws, r, 2)
        low = pos.lower().strip()
        if not pos:
            continue

        qty = _cn(ws, r, cols['col_qty'])

        # Секционный маркер: строка "ЛЕТО"/"ЗИМА" без кол-ва человек
        if qty <= 0:
            if low == 'лето':
                current_section_season = 'лето'
            elif low == 'зима':
                current_section_season = 'зима'
            continue

        if len(low) < 3:
            continue
        if re.match(r'^\d+(\.\d+)?$', low):
            continue
        if any(low.startswith(p) for p in SKIP_POS_PREFIXES):
            continue
        if any(p in low for p in SKIP_POS_CONTAINS):
            continue
        if qty > 300:
            continue

        months_cnt_raw = _cn(ws, r, cols['col_months'])
        months_cnt = int(months_cnt_raw) if 1 <= months_cnt_raw <= 60 else 12

        sal_month = _cn(ws, r, cols['col_sal_month'])
        if sal_month < 1000:
            for c in range(cols['col_sal_month'] - 2, cols['col_sal_month'] + 6):
                v = _cn(ws, r, c)
                if 5000 <= v <= 3_000_000:
                    sal_month = v
                    break
        if sal_month < 1000:
            continue

        sched_raw = _cs(ws, r, cols['col_sched'])
        sched = re.sub(r'\s*\d{1,2}[:.]\d{2}[-–]\d{1,2}[:.]\d{2}', '', sched_raw).strip()

        rate_fixed = _cn(ws, r, cols['col_rate_shift'])

        raw.append({
            'pos': pos.strip(), 'qty': int(qty), 'sched': sched,
            'months_cnt': months_cnt, 'sal_month': sal_month,
            'rate_fixed': rate_fixed,
            'section_season': current_section_season,
        })
    return raw


def merge_shift_pairs(raw_staff: list[dict]) -> list[dict]:
    """
    Объединяет пары строк "6/1" + "1/6" одной должности в единый график "7х0".
    Ставка (rate_fixed) переносится как есть (она одинакова у обеих строк),
    зарплаты складываются.
    """
    merged, used = [], set()
    for i, a in enumerate(raw_staff):
        if i in used:
            continue
        sched_a = a['sched'].lower()
        is_61 = re.match(r'^6[/хx]1', sched_a)
        is_16 = re.match(r'^1[/хx]6', sched_a)

        if is_61 or is_16:
            partner_idx = None
            for j in range(i + 1, len(raw_staff)):
                if j in used:
                    continue
                b = raw_staff[j]
                sched_b = b['sched'].lower()
                matches = (
                    (is_61 and re.match(r'^1[/хx]6', sched_b)) or
                    (is_16 and re.match(r'^6[/хx]1', sched_b))
                )
                # ВАЖНО: сверяем и ставку (rate_fixed) — при разных ставках
                # у одноимённой должности (напр. ВИ.ру: ОПУ/ОПМ день по 3700
                # и по 3500 — это РАЗНЫЕ группы, объединять их нельзя).
                same_position = (
                    b['pos'] == a['pos'] and
                    b['qty'] == a['qty'] and
                    b['months_cnt'] == a['months_cnt'] and
                    abs(b['rate_fixed'] - a['rate_fixed']) < 0.01 and
                    b.get('section_season') == a.get('section_season')
                )
                if matches and same_position:
                    partner_idx = j
                    break

            if partner_idx is not None:
                b = raw_staff[partner_idx]
                used.add(i)
                used.add(partner_idx)
                merged.append({
                    'pos': a['pos'], 'qty': a['qty'], 'sched': '7х0',
                    'months_cnt': a['months_cnt'],
                    'sal_month': a['sal_month'] + b['sal_month'],
                    'rate_fixed': a['rate_fixed'],
                    'section_season': a.get('section_season'),
                })
                continue

        used.add(i)
        merged.append(a)

    return merged


def parse_workbook(filepath: str) -> dict:
    """Точка входа: парсит .xlsx любого из двух форматов."""
    wb = load_workbook(filepath, data_only=True)
    filename = filepath.split('/')[-1]
    meta = get_meta(wb, filename)

    staff_sheet_name = None
    for n in wb.sheetnames:
        if re.match(r'^штат$', n.strip(), re.I) or re.search(r'расчет.{0,5}фот', n, re.I):
            staff_sheet_name = n
            break
    if not staff_sheet_name:
        staff_sheet_name = wb.sheetnames[0]

    ws = wb[staff_sheet_name]
    cols = detect_columns(ws)
    raw_staff = _parse_raw_staff(ws, cols)
    merged_staff = merge_shift_pairs(raw_staff)

    contract_months = max((s['months_cnt'] for s in merged_staff), default=12)

    budget = parse_budget_sheet(wb)

    return {
        'obj_name': meta['obj_name'],
        'obj_addr': meta['obj_addr'],
        'staff': merged_staff,
        'budget': budget,
        'contract_months': contract_months,
    }


# ═══════════════════════════════════════════════════════════
# РАСХОДЫ И МАТЕРИАЛЫ (лист «Калькуляция»)
# ═══════════════════════════════════════════════════════════

def parse_budget_sheet(wb) -> dict:
    """
    Читает фиксированные строки листа «Калькуляция» (стабильны между файлами
    одного шаблона): 16=расходные материалы(ИТОГО раздела — не путать со
    строкой 17 «химия», которая лишь одна из подкатегорий и может быть
    пустой при ненулевом общем итоге раздела), 21=инвентарь, 30=спецодежда,
    39=амортизация(итого), 48=ГСМ, 55=ремонт, 57=СОУТ.
    Плюс детализация оборудования (41-52) с пометкой аренды.
    """
    kn = next((n for n in wb.sheetnames if re.search(r'калькул', n, re.I)), None)
    if not kn:
        return {}
    ws = wb[kn]

    budget = {
        'chem':    {'label': 'Расходные материалы (химия и РМ)', 'total': _cn(ws, 16, 6)},
        'inv':     {'label': 'Инвентарь',                          'total': _cn(ws, 21, 6)},
        'uniform': {'label': 'Спецодежда',                         'total': _cn(ws, 30, 6)},
        'gsm':     {'label': 'ГСМ',                                'total': _cn(ws, 48, 6)},
        'maint':   {'label': 'Эксплуатация и ремонт оборудования', 'total': _cn(ws, 55, 6)},
        'sout':    {'label': 'Расходы на СОУТ',                    'total': _cn(ws, 57, 6)},
    }

    equipment = []
    for r in range(41, 53):
        name = _cs(ws, r, 2)
        qty = _cn(ws, r, 3)
        price = _cn(ws, r, 4)
        life = _cn(ws, r, 5)
        val = _cn(ws, r, 6)
        if name and val > 0:
            equipment.append({
                'name': name, 'qty': qty, 'price': price, 'life': life,
                'val': val,
                'is_rent': 'аренда' in name.lower(),  # признак аренды
            })
    budget['equipment'] = equipment

    return budget


# ═══════════════════════════════════════════════════════════
# ПОСТРОЕНИЕ ПОМЕСЯЧНОГО БЮДЖЕТА
# ═══════════════════════════════════════════════════════════

def build_months(
    parsed: dict,
    start_month: int,
    start_year: int,
    n_months: int,
    partial_bounds: dict | None = None,
    oklad_overrides: dict | None = None,
) -> list[dict]:
    """
    Строит список месяцев с посчитанным ФОТ по каждой должности.

    partial_bounds: опционально {0: (d1,d2), n_months-1: (d1,d2)} —
    для частичных первого/последнего месяца (сотрудники вышли не с 1 числа).

    oklad_overrides: опционально {'название должности (lowercase, часть строки)': True/False, ...}
    Позволяет вручную переопределить автоправило "менеджер/директор → Фикс".
    Пример: должность "Ночной Администратор" по умолчанию не подходит под
    автоправило (нет ключевых слов), но по факту у неё фиксированный оклад:
        oklad_overrides = {'ночной администратор': True}
    Также можно принудительно СНЯТЬ Фикс с реального менеджера:
        oklad_overrides = {'менеджер по клинингу': False}
    Сравнение — по вхождению подстроки в название должности (без регистра).
    """
    staff = parsed['staff']
    months = []
    oklad_overrides = oklad_overrides or {}

    def resolve_oklad(pos_name: str) -> bool:
        name_low = pos_name.lower()
        for key, forced in oklad_overrides.items():
            if key.lower() in name_low:
                return forced
        return is_oklad_by_name(pos_name)

    for i in range(n_months):
        mn = ((start_month - 1 + i) % 12) + 1
        yr = start_year + (start_month - 1 + i) // 12
        full_days = days_in(yr, mn)

        if partial_bounds and i in partial_bounds:
            d1, d2 = partial_bounds[i]
            avail_days = d2 - d1 + 1
        else:
            d1, d2 = 1, full_days
            avail_days = full_days

        month_staff = []
        for s in staff:
            if not is_position_active(s['pos'], s['months_cnt'], mn, s.get('section_season')):
                continue

            is_oklad = resolve_oklad(s['pos'])

            if is_oklad:
                # Фикс-оклад НЕ пропорционируется по частичному месяцу — платится
                # полностью независимо от даты выхода на объект в этом месяце.
                salary_all = s['sal_month'] * s['qty']
                month_staff.append({
                    'pos': s['pos'], 'qty': s['qty'], 'sched': s['sched'],
                    'rate': 0, 'sal': s['sal_month'], 'shifts': 0,
                    'tot_shifts': 0, 'salary_all': salary_all, 'is_oklad': True,
                })
            else:
                if avail_days < full_days:
                    # Частичный месяц: считаем РЕАЛЬНЫЕ рабочие смены внутри
                    # диапазона по графику должности (не просто доступные дни)
                    sched_low = (s['sched'] or '').lower()
                    if re.match(r'^5[/хx]2', sched_low):
                        shifts = workdays_5x2_range(yr, mn, d1, d2)
                    elif re.match(r'^7[/хx]0', sched_low):
                        shifts = avail_days
                    else:
                        # прочие графики — пропорция от полной месячной нормы
                        full_shifts = calc_shifts(s['sched'], yr, mn) or full_days
                        shifts = round(full_shifts * avail_days / full_days)
                else:
                    shifts = calc_shifts(s['sched'], yr, mn) or full_days
                rate = s['rate_fixed']  # ФИКСИРОВАННАЯ ставка — не пересчитывать!
                tot_shifts = shifts * s['qty']
                salary_all = rate * tot_shifts
                month_staff.append({
                    'pos': s['pos'], 'qty': s['qty'], 'sched': s['sched'],
                    'rate': rate, 'sal': s['sal_month'], 'shifts': shifts,
                    'tot_shifts': tot_shifts, 'salary_all': salary_all,
                    'is_oklad': False,
                })

        total = sum(x['salary_all'] for x in month_staff)
        months.append({
            'yr': yr, 'mn': mn, 'days': avail_days, 'full_days': full_days,
            'd1': d1, 'd2': d2, 'staff': month_staff, 'total': total,
        })

    return months


def build_months_from_date(
    parsed: dict,
    start_date,
    n_years: float,
    oklad_overrides: dict | None = None,
) -> list[dict]:
    """
    Удобная обёртка над build_months() для мультигодового периода от
    произвольной календарной даты начала оказания услуг (не обязательно
    1-е число месяца — тогда первый и последний месяцы станут частичными
    автоматически).

    start_date: datetime.date — дата начала оказания услуг.
    n_years: количество лет контракта (может быть дробным, например 1.5).

    Возвращает тот же формат, что и build_months().
    """
    total_months = round(n_years * 12)
    start_month, start_year = start_date.month, start_date.year

    # Считаем реальное число календарных МЕСЯЧНЫХ БЛОКОВ (на 1 больше, если
    # старт не с 1-го числа — тогда добавляется частичный "хвост"). Сама
    # последовательность месяцев для каждого блока строит build_months().
    n_blocks = total_months
    partial_bounds = {}
    if start_date.day != 1:
        first_full_days = days_in(start_year, start_month)
        partial_bounds[0] = (start_date.day, first_full_days)
        n_blocks += 1
        last_day = start_date.day - 1
        partial_bounds[n_blocks - 1] = (1, last_day)

    return build_months(
        parsed, start_month=start_month, start_year=start_year,
        n_months=n_blocks, partial_bounds=partial_bounds,
        oklad_overrides=oklad_overrides,
    )
