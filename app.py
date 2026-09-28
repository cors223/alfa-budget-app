"""
app.py
──────
Flask-сервер: принимает загруженный .xlsx, дату начала оказания услуг и
количество лет — строит бюджет (с живыми Excel-формулами, включая
переключатель Фикс/Смена прямо в файле) через budget_engine + excel_builder
и отдаёт готовый файл на скачивание.

Запуск:
    pip install -r requirements.txt
    python app.py
    # открыть http://localhost:5000
"""

import io
import os
import re
import tempfile
from datetime import date, datetime

from flask import Flask, request, render_template, send_file, jsonify

from budget_engine import parse_workbook, build_months_from_date
from excel_builder import build_workbook, MNAME

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024  # 20 МБ лимит на файл


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/parse-preview', methods=['POST'])
def api_parse_preview():
    """
    Первый шаг: только парсит файл и возвращает список найденных должностей
    с автоматически определённым типом (Фикс/Смена) — чтобы показать в форме
    чекбоксы для ручной корректировки ПЕРЕД генерацией итогового файла.
    """
    file = request.files.get('file')
    if not file or not file.filename.lower().endswith(('.xlsx', '.xls')):
        return jsonify({'error': 'Загрузите файл .xlsx'}), 400

    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        file.save(tmp.name)
        tmp_path = tmp.name
    try:
        parsed = parse_workbook(tmp_path)
    finally:
        os.unlink(tmp_path)

    if not parsed['staff']:
        return jsonify({'error': 'Не удалось найти должности в файле. Проверьте структуру.'}), 422

    from budget_engine import is_oklad_by_name
    positions = []
    seen = set()
    for s in parsed['staff']:
        if s['pos'] in seen:
            continue
        seen.add(s['pos'])
        positions.append({'pos': s['pos'], 'auto_fix': is_oklad_by_name(s['pos'])})

    return jsonify({
        'obj_name': parsed['obj_name'],
        'obj_addr': parsed['obj_addr'],
        'positions': positions,
        'contract_months_hint': parsed['contract_months'],
    })


@app.route('/api/budget', methods=['POST'])
def api_budget():
    file = request.files.get('file')
    if not file or not file.filename.lower().endswith(('.xlsx', '.xls')):
        return jsonify({'error': 'Загрузите файл .xlsx'}), 400

    # ── Дата начала оказания услуг + срок в годах ──
    start_date_raw = request.form.get('start_date', '').strip()
    n_years_raw = request.form.get('n_years', '1').strip()
    try:
        start_date = datetime.strptime(start_date_raw, '%Y-%m-%d').date() if start_date_raw else date.today()
    except ValueError:
        return jsonify({'error': 'Некорректная дата начала услуг'}), 400
    try:
        n_years = float(n_years_raw) if n_years_raw else 1.0
    except ValueError:
        n_years = 1.0


    # Ручные переопределения Фикс/Смена — приходят из формы как
    # oklad_overrides[Должность]=1 или 0 (чекбоксы на превью-шаге)
    oklad_overrides = {}
    for key in request.form:
        if key.startswith('oklad_overrides[') and key.endswith(']'):
            pos_name = key[len('oklad_overrides['):-1]
            val = request.form.get(key)
            oklad_overrides[pos_name.strip()] = val in ('1', 'true', 'on', 'да')

    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        file.save(tmp.name)
        tmp_path = tmp.name
    try:
        parsed = parse_workbook(tmp_path)
    finally:
        os.unlink(tmp_path)

    if not parsed['staff']:
        return jsonify({'error': 'Не удалось найти должности в файле. Проверьте структуру.'}), 422

    months = build_months_from_date(
        parsed, start_date=start_date, n_years=n_years,
        oklad_overrides=oklad_overrides or None,
    )

    period = (
        f"{months[0]['d1']:02d}.{months[0]['mn']:02d}.{months[0]['yr']} — "
        f"{months[-1]['d2']:02d}.{months[-1]['mn']:02d}.{months[-1]['yr']}"
    )


    # budget_engine.parse_budget_sheet() возвращает вложенные словари
    # {'chem': {'label':.., 'total':..}, ...} + отдельный ключ 'equipment'.
    # excel_builder.build_workbook() ожидает плоские числа и equipment
    # отдельным параметром — приводим формат здесь.
    raw_budget = parsed.get('budget') or {}
    equipment = raw_budget.get('equipment', [])
    flat_budget = {
        k: v['total'] for k, v in raw_budget.items()
        if k != 'equipment' and isinstance(v, dict)
    }

    wb = build_workbook(
        obj_name=parsed['obj_name'] or 'Объект',
        obj_addr=parsed['obj_addr'],
        period=period,
        months=months,
        budget=flat_budget,
        equipment=equipment,
    )

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    safe_name = re.sub(r'[^\w\-]+', '_', parsed['obj_name'] or 'Объект').strip('_')
    out_filename = f"{safe_name}_{start_date.strftime('%d%m%Y')}_{n_years}г.xlsx"

    return send_file(
        buf,
        as_attachment=True,
        download_name=out_filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


if __name__ == '__main__':
    app.run(debug=True, port=5000)
