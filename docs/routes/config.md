# `routes/config.py`

## Призначення
Уся конфігурація застосунку в одному місці: `SECRET_KEY` (підпис сесій),
налаштування ONLYOFFICE (JWT-секрет, публічна адреса, перевірка SSL),
`PREFERRED_URL_SCHEME`, `PUBLIC_APPLY_BASE_URL` (публічна адреса для
посилань `/apply` і `/update-info`, які бачить студент з інтернету, а
не внутрішня адреса адмінки). Кожне значення можна перевизначити
змінною середовища або файлом `.env` поруч з `app.py` - якщо нічого не
задано, використовується безпечне значення за замовчуванням і
попередження в лог.

## Взаємодія з app.py
`app.py` імпортує звідси напряму: `from routes.config import SECRET_KEY, PREFERRED_URL_SCHEME`
- обидва значення одразу застосовуються до `app.secret_key` і
`app.config['PREFERRED_URL_SCHEME']`. `PUBLIC_APPLY_BASE_URL` в `app.py`
не використовується - його імпортує `routes/students/core.py` для
`generate_update_link()`, а `ONLYOFFICE_*` - `routes/office_editor.py`.

## Ключові значення
- `SECRET_KEY` - якщо не задано в середовищі, генерується випадковий
  ключ "для цього запуску" (з попередженням у лог) - усі сесії
  скидаються при кожному перезапуску сервера, доки не задати постійний.
- `ONLYOFFICE_JWT_SECRET`, `ONLYOFFICE_PUBLIC_URL`, `ONLYOFFICE_VERIFY_SSL`
  - інтеграція з Document Server (`routes/office_editor.py`).
- `PUBLIC_APPLY_BASE_URL` - навмисно НЕ обчислюється з поточного
  запиту (`url_for(..., _external=True)`), бо адмін, який генерує
  посилання `/update-info`, сам перебуває на внутрішньому домені -
  посилання для студента має вести на публічну адресу завжди, незалежно
  від того, звідки адмін його створив.

## Залежить від
`os`, `secrets` (стандартна бібліотека), опційно `python-dotenv`.
