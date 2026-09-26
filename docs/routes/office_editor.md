# `routes/office_editor.py`

## Призначення
Інтеграція з ONLYOFFICE Document Server: перегляд і редагування щойно
згенерованих `.docx`-файлів прямо в браузері, замість "наосліп"
завантажувати файл і відкривати його локально. Підтримує і одиночний
документ (один студент), і пакетний перегляд/завантаження (масова
генерація по групі, ZIP).

## Взаємодія з app.py
Прямо: `from routes.office_editor import office_bp` +
`app.register_blueprint(office_bp)`. Викликається також опосередковано
з `routes/students/core.py` (`generate()`) і
`routes/admin/templates_and_export.py` (`generate_group_docs()`) - обидві
після `gen_doc(...)` реєструють щойно створений файл через
`create_editing_session(...)`, щоб дати користувачу посилання на
перегляд/редагування замість прямого завантаження.

## Маршрути
- `GET /edit/<doc_id>` - `edit()`: сторінка з вбудованим редактором
  ONLYOFFICE для одного документа.
- `GET /file/<doc_id>` - `serve_file()`: віддає сам файл Document
  Server'у (за підписаним запитом).
- `POST /callback/<doc_id>` - `callback()`: приймає збережений файл
  від Document Server після редагування (JWT-перевірений, якщо задано
  `ONLYOFFICE_JWT_SECRET`).
- `GET /finalize/<doc_id>` - `finalize()`: примусове збереження
  (`forcesave`) перед завантаженням, з очікуванням через
  `threading.Event`.
- `GET /batch/<batch_id>` - `batch_view()`: перегляд групи документів
  одразу (масова генерація).
- `GET /batch/<batch_id>/zip` - `download_batch_zip()`: пакетне
  завантаження всіх файлів групи одним ZIP.

## Залежить від
`routes.config` (`ONLYOFFICE_*`), `requests`/`jwt` (підпис запитів до
Document Server). Файли-джерела приходять із `routes/gen_docx.py`.
