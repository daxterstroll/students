# `routes/students/import_export.py`

## Призначення
Масовий імпорт студентів з `.xlsx`-файлу: зіставлення з наявними
групами (нечітке й точне), автоматична генерація англійського
написання ПІБ для тих, хто не вказаний у файлі.

## Взаємодія з app.py
Непряма (через `students_bp`, див. `routes/students/__init__.py`).

## Маршрути
`/import_from_excel`.

## Залежить від
`routes.utils` (`generate_english_name`), `openpyxl`,
`werkzeug.utils.secure_filename`.
