# `routes/admin/documents.py`

## Призначення
Документи про освіту студентів (атестат/диплом попереднього рівня,
+ підтвердження визнання іноземного документа) і паспортні дані
(книжка/ID-картка) - обидва зі сканами-вкладеннями. Два імпорти з
Excel: документи про освіту - **двоетапний** (попередній перегляд із
перевіркою відповідності → підтвердження), паспортні дані -
**одноетапний**. Містить набір текстових допоміжних функцій для
розпізнавання нестандартно оформлених Excel-комірок при імпорті
документів (переклад, визначення країни, парсинг складених рядків).

## Взаємодія з app.py
Непряма (через `admin_bp`, див. `routes/admin/__init__.py`).

## Маршрути
`/admin/manage_education_documents`, `/admin/manage_passport_documents`,
`/admin/import_passport_documents`,
`/admin/import_education_docs_preview`, `/admin/import_docs_commit`.

## Ключові допоміжні функції
- `fuzzy_find_student(cursor, full_name)` - нечіткий пошук студента за
  ПІБ серед усіх активних (використовується й імпортом паспортних
  даних, і імпортом документів про освіту).
- `translate_to_en(text)` - переклад через `GoogleTranslator`, з
  кешуванням (`translation_cache`) щоб не робити повторні запити для
  однакового тексту.
- `parse_document`, `parse_reference_cell_ua`, `parse_recognition_cell_ua`
  - розбір складених текстових комірок Excel (де кілька полів записані
  в одну клітинку через роздільник) на окремі значення.
- `save_preview_to_file` / `load_preview_from_file` - стан
  попереднього перегляду двоетапного імпорту зберігається в
  тимчасовому файлі (`TEMP_PREVIEW_FOLDER`), не в cookie-сесії.

## Залежить від
`routes.utils` (`get_attachments`, `save_multiple_attachments`),
`openpyxl`, `pandas`, `rapidfuzz`, `deep_translator`.
