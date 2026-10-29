"""
Парсер PDF "Резерв+" (військово-обліковий документ) - версія 4.

Колонки в межах одного рядка визначаються порядком груп зліва направо
(групи розділені розривом > GAP_THRESHOLD) - бо кожен рядок документа
має свою власну кількість і позицію колонок (Адреса/Телефон - лише 2,
на інших x-координатах, ніж Звання/Реєстр/ВОС), і намагання звести все
до єдиного набору "еталонних" координат ламається на цих розбіжностях.

Єдиний виняток - перенесене на другий рядок значення "Відстрочка до"
(напр. "завершення" / "мобілізації"): рядок з одним словом не має з чим
порівнювати розрив, тому для нього окремо обчислюється поріг x0 із самого
рядка міток "Тип відстрочки: / Відстрочка до".
"""
import re
from collections import defaultdict
import pdfplumber

GAP_THRESHOLD = 20
DASH_VALUES = {'-', '—', ''}


def _group_lines(words):
    lines = defaultdict(list)
    for w in words:
        key = round(w['top'] / 3) * 3
        lines[key].append(w)
    return [sorted(lines[top], key=lambda w: w['x0']) for top in sorted(lines.keys())]


def _split_columns(line_words, max_cols=3):
    """Групи зліва направо = колонки 0,1,2 за порядком появи."""
    cols = [''] * max_cols
    if not line_words:
        return cols
    groups = [[line_words[0]]]
    prev_x1 = line_words[0]['x1']
    for w in line_words[1:]:
        if w['x0'] - prev_x1 > GAP_THRESHOLD:
            groups.append([])
        groups[-1].append(w)
        prev_x1 = w['x1']
    for idx, g in enumerate(groups[:max_cols]):
        cols[idx] = ' '.join(w['text'] for w in g)
    return cols


def parse_reserve_plus_pdf(path):
    data = {}
    with pdfplumber.open(path) as pdf:
        page = pdf.pages[0]
        words = page.extract_words(use_text_flow=False)
        full_text = page.extract_text() or ''
        raw_lines = _group_lines(words)

    mm = re.search(r'Сформовано:\s*([\d.]+),?\s*([\d:]+)', full_text)
    data['formed_at'] = f"{mm.group(1)} {mm.group(2)}" if mm else ''

    cols = [_split_columns(l) for l in raw_lines]

    anchor_idx = None
    for i, (l, m, r) in enumerate(cols):
        if l.startswith('Дата народження'):
            anchor_idx = i
            break
    if anchor_idx is None:
        raise ValueError("Не вдалося розпізнати документ - не знайдено 'Дата народження'")

    header_lines = []
    for l, m, r in cols[:anchor_idx]:
        if l.startswith('Військово-обліковий') or l.startswith('Сформовано'):
            continue
        if l:
            header_lines.append(l)
    KNOWN_BADGES = ['Військовозобовʼязаний', "Військовозобов'язаний", 'Призовник', 'Резервіст', 'Невійськовозобовʼязаний']
    if header_lines and header_lines[0] in KNOWN_BADGES:
        data['status_badge'] = header_lines[0]
        name_lines = header_lines[1:]
    else:
        data['status_badge'] = ''
        name_lines = header_lines
    data['full_name'] = ' '.join(name_lines).strip()

    vl, vm, vr = cols[anchor_idx + 1]
    data['birth_date'] = vl
    data['valid_until'] = vm
    data['tax_id'] = vr

    i = anchor_idx + 2
    assert cols[i][0].startswith('Категорія обліку'), (i, cols[i])
    vl, vm, vr = cols[i + 1]
    data['registration_category'] = vl
    data['removal_basis'] = (vm or vr) if (vm or vr) not in DASH_VALUES else ''
    i += 2

    assert cols[i][0].startswith('ТЦК'), (i, cols[i])
    i += 1
    tck_lines = []
    while i < len(cols) and not cols[i][0].startswith('Звання'):
        tck_lines.append(cols[i][0])
        i += 1
    data['tck_name'] = ' '.join(tck_lines).strip()

    assert cols[i][0].startswith('Звання'), (i, cols[i])
    vl, vm, vr = cols[i + 1]
    data['military_rank'] = vl if vl not in DASH_VALUES else ''
    data['registry_number'] = vm if vm not in DASH_VALUES else ''
    data['vos_code'] = vr if vr not in DASH_VALUES else ''
    i += 2

    vos_desc_lines = []
    while i < len(cols) and not cols[i][0].startswith('Тип відстрочки'):
        full_line = ' '.join(filter(None, cols[i])).strip()
        if full_line and full_line not in DASH_VALUES:
            vos_desc_lines.append(full_line)
        i += 1
    data['vos_description'] = ' '.join(vos_desc_lines).strip()

    # --- Тип відстрочки / Відстрочка до ---
    assert cols[i][0].startswith('Тип відстрочки'), (i, cols[i])
    # Поріг для continuation-рядків з одним словом: середина між x0
    # мітки "Тип відстрочки:" і міткою "Відстрочка до" на ЦЬОМУ рядку.
    label_words = raw_lines[i]
    type_label_x0 = label_words[0]['x0']
    until_label_x0 = next((w['x0'] for w in label_words if w['text'] == 'Відстрочка'), None)
    mid_threshold = (type_label_x0 + until_label_x0) / 2 if until_label_x0 else 300
    i += 1
    deferral_type_parts, deferral_until_parts = [], []
    while i < len(cols) and not cols[i][0].startswith('Причина звернення'):
        l, m, r = cols[i]
        second = m or r
        if l and second:
            deferral_type_parts.append(l)
            deferral_until_parts.append(second)
        elif l and not second:
            # Один виявлений фрагмент на рядку - визначаємо, якій
            # колонці він належить, за власною x0-позицією слова.
            first_word_x0 = raw_lines[i][0]['x0']
            if first_word_x0 >= mid_threshold:
                deferral_until_parts.append(l)
            else:
                deferral_type_parts.append(l)
        i += 1
    dt = ' '.join(deferral_type_parts).strip()
    du = ' '.join(deferral_until_parts).strip()
    data['deferral_type'] = dt if dt not in DASH_VALUES else ''
    data['deferral_until'] = du if du not in DASH_VALUES else ''

    assert cols[i][0].startswith('Причина звернення'), (i, cols[i])
    vl, vm, vr = cols[i + 1]
    data['police_reason'] = vl if vl not in DASH_VALUES else ''
    data['police_date'] = (vm or vr) if (vm or vr) not in DASH_VALUES else ''
    i += 2

    assert cols[i][0].startswith('Постанова ВЛК'), (i, cols[i])
    vl, vm, vr = cols[i + 1]
    data['vlk_decision'] = vl if vl not in DASH_VALUES else ''
    data['vlk_date'] = (vm or vr) if (vm or vr) not in DASH_VALUES else ''
    i += 2

    assert cols[i][0].startswith('Група інвалідності'), (i, cols[i])
    vl, vm, vr = cols[i + 1]
    data['disability_group'] = vl if vl not in DASH_VALUES else ''
    data['disability_valid_until'] = (vm or vr) if (vm or vr) not in DASH_VALUES else ''
    i += 2

    assert cols[i][0].startswith('Причина інвалідності'), (i, cols[i])
    val = cols[i + 1][0]
    data['disability_reason'] = val if val not in DASH_VALUES else ''
    i += 2

    assert cols[i][0].startswith('Адреса проживання'), (i, cols[i])
    i += 1
    addr_parts, phone = [], ''
    while i < len(cols) and not cols[i][0].startswith('Email'):
        l, m, r = cols[i]
        if l:
            addr_parts.append(l)
        # Телефон - перша непорожня "друга колонка" на цих рядках
        # (не обов'язково cols[2] - бо на цьому рядку лише 2 колонки).
        second_val = m or r
        if second_val and not phone:
            phone = second_val
        i += 1
    data['address'] = ' '.join(addr_parts).strip()
    data['phone'] = phone

    assert cols[i][0].startswith('Email'), (i, cols[i])
    val = cols[i + 1][0]
    data['email'] = val if val not in DASH_VALUES else ''
    i += 2

    assert any(c.startswith('Дата уточнення') for c in cols[i]), (i, cols[i])
    i += 1
    data['data_verified_on_time'] = False
    data['data_update_date'] = ''
    while i < len(cols):
        line_text = ' '.join(filter(None, cols[i])).strip()
        if line_text.startswith('Дані уточнено вчасно'):
            data['data_verified_on_time'] = True
        date_match = next((c for c in cols[i] if re.match(r'^\d{2}\.\d{2}\.\d{4}$', c)), None)
        if date_match:
            data['data_update_date'] = date_match
            break
        if line_text.startswith('*') or line_text.startswith('Документ дійсний'):
            break
        i += 1

    return data


if __name__ == '__main__':
    import glob, json
    all_ok = True
    for path in sorted(glob.glob('/mnt/user-data/uploads/*.pdf')):
        print("="*70)
        print(path.split('/')[-1])
        try:
            data = parse_reserve_plus_pdf(path)
            print(json.dumps(data, ensure_ascii=False, indent=2))
        except Exception:
            import traceback
            traceback.print_exc()
            all_ok = False
    print("\nALL OK" if all_ok else "\nSOME FAILED")
