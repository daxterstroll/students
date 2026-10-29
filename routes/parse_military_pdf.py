"""
Парсер PDF "Резерв+" (військово-обліковий документ) - версія 7.

Колонки в межах одного рядка визначаються порядком груп зліва направо
(групи розділені розривом > GAP_THRESHOLD) - бо кожен рядок документа
має свою власну кількість і позицію колонок.

ВЕРСІЯ 7 — стійкість до багаторядкових значень:
  - Довгі значення (напр. "Не військовозобов'язаний", "Личаківсько-
    Залізничний об'єднаний районний...") можуть переноситися на
    кілька рядків. Парсер більше не покладається на assert для
    "Категорія обліку" - шукає ключове слово через пошук рядка, що
    ПОЧИНАЄТЬСЯ з нього, і склеює продовження.
  - Кількість колонок у рядку може бути 1..3 - об'єднуємо порожні
    так, щоб ключове слово завжди було в cols[i][0].
"""
import re
from collections import defaultdict
import pdfplumber
from routes.utils import title_case_ukrainian

GAP_THRESHOLD = 20
DASH_VALUES = {'-', '—', ''}

# --------------------------------------------------------------------
# Координати рамки фото (у ВІДСОТКАХ від ширини/висоти сторінки PDF).
# Підібрані під ваш PDF — змініть, якщо рамка зсунута.
# --------------------------------------------------------------------
PHOTO_ZONE = {
    'x0_pct': 0.0681,
    'y0_pct': 0.0735,
    'x1_pct': 0.2340,
    'y1_pct': 0.2246,
}


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


def _find_line_starting_with(cols, prefix, start_idx=0):
    """
    Повертає індекс першого рядка (починаючи зі start_idx), у якому
    cols[i][0] починається з `prefix`. Якщо не знайдено — None.
    """
    for i in range(start_idx, len(cols)):
        if cols[i][0].startswith(prefix):
            return i
    return None


def _join_continuation(cols, key_line_idx, next_key_startswith,
                        key_col=0, join_with=' '):
    """
    Після key_line_idx збирає всі рядки, які НЕ починаються з
    next_key_startswith, як продовження значення, і повертає
    (об'єднане_значення, індекс_наступного_рядка).

    Використовується для багаторядкових значень на кшталт
    "Категорія обліку" → "Не військовозобов'язаний" (де "Не" і
    "військовозобов'язаний" рознесені на два рядки).
    """
    parts = []
    i = key_line_idx + 1
    while i < len(cols):
        first = cols[i][0] if cols[i][0] else (cols[i][1] if len(cols[i]) > 1 else '')
        if first.startswith(next_key_startswith):
            break
        # Продовження — беремо значення з відповідної колонки
        val = cols[i][key_col] if key_col < len(cols[i]) else ''
        if not val:
            # Іноді значення "переїжджає" в наступну колонку
            val = next((c for c in cols[i] if c), '')
        if val:
            parts.append(val)
        i += 1
    return join_with.join(parts).strip(), i

def _normalize_apostrophes(s):
    """Приводить будь-який апостроф до прямого ' для порівняння."""
    if not s:
        return ''
    return (s.replace('’', "'")
             .replace('ʼ', "'")
             .replace('‘', "'")
             .replace('`', "'"))

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

    # ── Якір: "Дата народження" ──
    anchor_idx = _find_line_starting_with(cols, 'Дата народження')
    if anchor_idx is None:
        raise ValueError("Не вдалося розпізнати документ - не знайдено 'Дата народження'")

    # ── ПІБ і бейдж ──
    header_lines = []
    for l, m, r in cols[:anchor_idx]:
        if l.startswith('Військово-обліковий') or l.startswith('Сформовано'):
            continue
        if l:
            header_lines.append(l)
    KNOWN_BADGES = [
        "Військовозобов'язаний",
        "Призовник",
        "Резервіст",
        "Невійськовозобов'язаний",
        "Знято з обліку",
    ]
    # Бейдж може мати будь-який апостроф (', ’, ʼ, ‘) — нормалізуємо
    # для порівняння, але зберігаємо ОРИГІНАЛЬНИЙ текст, як він у PDF.
    known_badges_norm = {_normalize_apostrophes(b): b for b in KNOWN_BADGES}
    badge_match = None
    if header_lines:
        first_norm = _normalize_apostrophes(header_lines[0])
        badge_match = known_badges_norm.get(first_norm)
    if badge_match is not None:
        # Зберігаємо канонічний варіант із KNOWN_BADGES (з прямим '),
        # щоб потім у core.py / military.py порівняння було однаковим.
        data['status_badge'] = badge_match
        name_lines = header_lines[1:]
    else:
        data['status_badge'] = ''
        name_lines = header_lines

    data['full_name'] = title_case_ukrainian(' '.join(name_lines).strip())

    # ── Дата народження / Дійсний до / РНОКПП ──
    vl, vm, vr = cols[anchor_idx + 1]
    data['birth_date'] = vl
    data['valid_until'] = vm
    data['tax_id'] = vr

    # ── Категорія обліку ──
    # УВАГА: у деяких PDF значення "Не військовозобов'язаний"
    # переноситься на два рядки: "військовозобов'язаний" (хвіст)
    # іде ПЕРЕД рядком-ключем "Категорія обліку" в наступному рядку
    # після "Дата народження". Тому шукаємо ключ, а не покладаємось
    # на позицію.
    i = anchor_idx + 2

    # Шукаємо рядок-ключ "Категорія обліку"
    key_i = _find_line_starting_with(cols, 'Категорія обліку', start_idx=i)
    if key_i is None:
        raise ValueError("Не знайдено рядок 'Категорія обліку'")

    # Значення — у наступному рядку після ключа, але може бути
    # розбите на 2-3 рядки ("Не" + "військовозобов'язаний" тощо).
    # Також може мати 2 колонки: категорія + підстава.
    val_parts = []
    val2_parts = []
    j = key_i + 1
    while j < len(cols):
        l = cols[j][0]
        if l.startswith('Підстава') or l.startswith('ТЦК'):
            break
        val_parts.append(l)
        # Друга колонка — підстава (якщо є)
        m = cols[j][1] if len(cols[j]) > 1 else ''
        if m:
            val2_parts.append(m)
        j += 1
    data['registration_category'] = ' '.join(val_parts).strip()
    data['removal_basis'] = ' '.join(val2_parts).strip()
    if data['removal_basis'] in DASH_VALUES:
        data['removal_basis'] = ''

    # ── ТЦК ──
    tck_i = _find_line_starting_with(cols, 'ТЦК', start_idx=j)
    if tck_i is None:
        raise ValueError("Не знайдено рядок 'ТЦК'")
    tck_parts = []
    j = tck_i + 1
    while j < len(cols):
        l = cols[j][0]
        if l.startswith('Звання'):
            break
        if l:
            tck_parts.append(l)
        j += 1
    data['tck_name'] = ' '.join(tck_parts).strip()

    # ── Звання / Реєстр / ВОС ──
    rank_i = _find_line_starting_with(cols, 'Звання', start_idx=j)
    if rank_i is None:
        raise ValueError("Не знайдено рядок 'Звання'")
    vl, vm, vr = cols[rank_i + 1]
    data['military_rank'] = vl if vl not in DASH_VALUES else ''
    data['registry_number'] = vm if vm not in DASH_VALUES else ''
    data['vos_code'] = vr if vr not in DASH_VALUES else ''

    # ── Опис ВОС ──
    j = rank_i + 2
    vos_desc_lines = []
    while j < len(cols):
        l = cols[j][0]
        if l.startswith('Тип відстрочки'):
            break
        # Опис може бути в першій або в наступних колонках
        full_line = ' '.join(filter(None, cols[j])).strip()
        if full_line and full_line not in DASH_VALUES:
            vos_desc_lines.append(full_line)
        j += 1
    data['vos_description'] = ' '.join(vos_desc_lines).strip()

    # ── Тип відстрочки / Відстрочка до ──
    deferral_i = _find_line_starting_with(cols, 'Тип відстрочки', start_idx=j)
    if deferral_i is None:
        raise ValueError("Не знайдено рядок 'Тип відстрочки'")

    label_words = raw_lines[deferral_i]
    type_label_x0 = label_words[0]['x0']
    until_label_x0 = next((w['x0'] for w in label_words if w['text'] == 'Відстрочка'), None)
    mid_threshold = (type_label_x0 + until_label_x0) / 2 if until_label_x0 else 300

    deferral_type_parts, deferral_until_parts = [], []
    j = deferral_i + 1
    while j < len(cols):
        l = cols[j][0]
        if l.startswith('Причина звернення'):
            break
        m = cols[j][1] if len(cols[j]) > 1 else ''
        r = cols[j][2] if len(cols[j]) > 2 else ''
        second = m or r
        if l and second:
            deferral_type_parts.append(l)
            deferral_until_parts.append(second)
        elif l and not second:
            first_word_x0 = raw_lines[j][0]['x0']
            if first_word_x0 >= mid_threshold:
                deferral_until_parts.append(l)
            else:
                deferral_type_parts.append(l)
        j += 1
    dt = ' '.join(deferral_type_parts).strip()
    du = ' '.join(deferral_until_parts).strip()
    data['deferral_type'] = dt if dt not in DASH_VALUES else ''
    data['deferral_until'] = du if du not in DASH_VALUES else ''

    # ── Поліція ──
    police_i = _find_line_starting_with(cols, 'Причина звернення', start_idx=j)
    if police_i is not None:
        vl, vm, vr = cols[police_i + 1]
        data['police_reason'] = vl if vl not in DASH_VALUES else ''
        data['police_date'] = (vm or vr) if (vm or vr) not in DASH_VALUES else ''
        j = police_i + 2
    else:
        data['police_reason'] = ''
        data['police_date'] = ''

    # ── ВЛК ──
    vlk_i = _find_line_starting_with(cols, 'Постанова ВЛК', start_idx=j)
    if vlk_i is not None:
        vl, vm, vr = cols[vlk_i + 1]
        data['vlk_decision'] = vl if vl not in DASH_VALUES else ''
        data['vlk_date'] = (vm or vr) if (vm or vr) not in DASH_VALUES else ''
        j = vlk_i + 2
    else:
        data['vlk_decision'] = ''
        data['vlk_date'] = ''

    # ── Інвалідність ──
    dis_i = _find_line_starting_with(cols, 'Група інвалідності', start_idx=j)
    if dis_i is not None:
        vl, vm, vr = cols[dis_i + 1]
        data['disability_group'] = vl if vl not in DASH_VALUES else ''
        data['disability_valid_until'] = (vm or vr) if (vm or vr) not in DASH_VALUES else ''
        j = dis_i + 2
    else:
        data['disability_group'] = ''
        data['disability_valid_until'] = ''

    reason_i = _find_line_starting_with(cols, 'Причина інвалідності', start_idx=j)
    if reason_i is not None:
        val = cols[reason_i + 1][0]
        data['disability_reason'] = val if val not in DASH_VALUES else ''
        j = reason_i + 2
    else:
        data['disability_reason'] = ''

    # ── Адреса / Телефон ──
    addr_i = _find_line_starting_with(cols, 'Адреса проживання', start_idx=j)
    if addr_i is None:
        raise ValueError("Не знайдено рядок 'Адреса проживання'")
    addr_parts, phone = [], ''
    j = addr_i + 1
    while j < len(cols):
        l = cols[j][0]
        if l.startswith('Email'):
            break
        if l:
            addr_parts.append(l)
        m = cols[j][1] if len(cols[j]) > 1 else ''
        r = cols[j][2] if len(cols[j]) > 2 else ''
        second_val = m or r
        if second_val and not phone:
            phone = second_val
        j += 1
    data['address'] = ' '.join(addr_parts).strip()
    data['phone'] = phone

    # ── Email ──
    email_i = _find_line_starting_with(cols, 'Email', start_idx=j)
    if email_i is not None:
        val = cols[email_i + 1][0]
        data['email'] = val if val not in DASH_VALUES else ''
        j = email_i + 2
    else:
        data['email'] = ''

    # ── Дата уточнення ──
    data['data_verified_on_time'] = False
    data['data_update_date'] = ''
    date_i = _find_line_starting_with(cols, 'Дата уточнення', start_idx=j)
    if date_i is not None:
        j = date_i + 1
        while j < len(cols):
            line_text = ' '.join(filter(None, cols[j])).strip()
            if 'Дані уточнено вчасно' in line_text:
                data['data_verified_on_time'] = True
            date_match = next(
                (c for c in cols[j] if re.match(r'^\d{2}\.\d{2}\.\d{4}$', c)), None
            )
            if date_match:
                data['data_update_date'] = date_match
                break
            if line_text.startswith('*') or line_text.startswith('Документ дійсний'):
                break
            j += 1
    else:
        # fallback: пошук "Дані уточнено вчасно" в усьому тексті
        if 'Дані уточнено вчасно' in full_text:
            data['data_verified_on_time'] = True
        m = re.search(r'Дата уточнення даних\s+(\d{2}\.\d{2}\.\d{4})', full_text)
        if m:
            data['data_update_date'] = m.group(1)

    return data


def extract_photo_from_pdf(path, dpi=300, inner_margin=2):
    """Витягує фото 3×4 за PHOTO_ZONE. Див. попередні версії."""
    try:
        import fitz
    except ImportError:
        return None

    try:
        doc = fitz.open(path)
    except Exception:
        return None

    try:
        page = doc.load_page(0)
        page_rect = page.rect

        clip = fitz.Rect(
            page_rect.width * PHOTO_ZONE['x0_pct'],
            page_rect.height * PHOTO_ZONE['y0_pct'],
            page_rect.width * PHOTO_ZONE['x1_pct'],
            page_rect.height * PHOTO_ZONE['y1_pct'],
        )

        if inner_margin > 0:
            clip = fitz.Rect(
                clip.x0 + inner_margin,
                clip.y0 + inner_margin,
                clip.x1 - inner_margin,
                clip.y1 - inner_margin,
            )

        if clip.width < 20 or clip.height < 20:
            return None

        zoom = dpi / 72.0
        matrix = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=matrix, clip=clip, alpha=False)
        data = pix.tobytes("png")

        if not _has_face_pixels(data):
            return None
        return data
    except Exception:
        return None
    finally:
        doc.close()


def _has_face_pixels(png_bytes, min_ratio=0.02):
    try:
        import io
        import numpy as np
        from PIL import Image
    except ImportError:
        return True
    try:
        img = Image.open(io.BytesIO(png_bytes)).convert("RGB")
        img.thumbnail((300, 300))
        arr = np.array(img).astype(int)
        r, g, b = arr[..., 0], arr[..., 1], arr[..., 2]
        skin = (
            (r > 120) & (r < 255) &
            (g > 90) & (g < 220) &
            (b > 60) & (b < 200) &
            (r > g) & (g > b) &
            ((r - b) > 15)
        )
        return skin.mean() >= min_ratio
    except Exception:
        return True