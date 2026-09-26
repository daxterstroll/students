# `routes/init_db.py`

## Призначення
Одноразовий скрипт (не Flask-модуль, не блюпринт) для створення
`students.db` "з нуля" з усіма таблицями і користувачем
`admin`/`admin123` за замовчуванням. Використовує
`CREATE TABLE IF NOT EXISTS` - повторний запуск на вже існуючій базі
не стирає дані, лише додає відсутні таблиці. **Не додає нові колонки
до вже існуючих таблиць** - для цього по всьому проєкту є окремі
`migrate_*.py`-скрипти в корені репозиторію.

## Взаємодія з app.py
**Ніякої.** `app.py` цей файл не імпортує і не запускає. Запускається
вручну одноразово: `python routes/init_db.py` (або через docker-
entrypoint/скрипт розгортання) - до першого старту `app.py` на новому
середовищі.

## Що створює
Усі таблиці застосунку: `students`, `groups`, `military`,
`education_documents`, `foreign_education_docs`, `passport_documents`,
`pending_students`, `update_requests`, `attachments`, `users`,
`specialties`, `accreditations`, `diplomas`, `student_study_periods`
та інші довідники - повний список схеми БД.

## Залежить від
`sqlite3`, `werkzeug.security` (для хешу пароля admin123). Нічого зі
`routes/*`.
