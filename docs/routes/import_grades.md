# `routes/import_grades.py`

## Призначення
Майстер імпорту оцінок з Excel ("зведена відомість") у 4 кроки:
завантаження файлу → вибір аркуша (якщо декілька) → зіставлення
колонок-предметів і рядків-студентів (з нечітким автопошуком) →
попередній перегляд і підтвердження. Підтримує обидва реальні формати
відомостей (звичайна і "розширена" з 3 колонками на предмет).

## Взаємодія з app.py
Прямо: `from routes.import_grades import import_grades_bp` +
`app.register_blueprint(import_grades_bp)`.

## Маршрути
- `GET, POST /admin/import_grades` - `upload()`: крок 1, завантаження
  файлу.
- `GET, POST /admin/import_grades/<token>/sheet` - `choose_sheet()`:
  крок 2.
- `GET, POST /admin/import_grades/<token>/mapping` - `mapping()`: крок
  3, зіставлення.
- `GET, POST /admin/import_grades/<token>/preview` - `preview()`: крок
  4, підтвердження запису в базу.

Стан майстра між кроками зберігається в тимчасових файлах на диску (за
`token`), а не в cookie-сесії - той самий підхід, що й у
`routes/photo_bulk.py`.

## Залежить від
`openpyxl`, `rapidfuzz` (нечітке зіставлення), `routes.utils`
(`is_student_on_reduced_program` для фільтрації видимих предметів).
