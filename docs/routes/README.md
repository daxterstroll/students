# Документація `routes/`

Один `.md`-файл на кожен `.py`-файл у `routes/` та його підпапках -
що робить файл і як він пов'язаний з `app.py` (прямо, через
`register_blueprint`, чи опосередковано, через імпорт іншими файлами).

## Як читати цю документацію

- **"Взаємодія з app.py: Прямо"** - файл містить Blueprint, який
  `app.py` імпортує і реєструє явним рядком
  (`app.register_blueprint(...)`), або `app.py` імпортує з нього
  конкретне ім'я напряму (`SECRET_KEY`, `logger`, `format_grade` тощо).
- **"Взаємодія з app.py: Непряма"** - `app.py` про існування цього
  конкретного файлу нічого не знає; він реєструється як частина
  пакета (`routes/admin/`, `routes/students/`), сам пакет `app.py`
  бачить як єдиний Blueprint через `__init__.py`.
- **"Взаємодія з app.py: Ніякої"** - самостійний скрипт
  (`init_db.py`, `update_groups.py`), який `app.py` не запускає і не
  імпортує; виконується вручну або за розкладом окремо від сайту.

## Мапа файлів

### Реєструються в `app.py` напряму (Blueprint'и)

| Файл | Blueprint | URL-префікс |
|---|---|---|
| [`auth.md`](auth.md) | `auth_bp` | `/`, `/login`, `/logout` |
| [`students/`](students/__init__.md) | `students_bp` | `/students/...`, `/grades/...`, `/faq` |
| [`admin/`](admin/__init__.md) | `admin_bp` | `/admin/...` |
| [`office_editor.md`](office_editor.md) | `office_bp` | `/edit/...`, `/batch/...` |
| [`import_grades.md`](import_grades.md) | `import_grades_bp` | `/admin/import_grades/...` |
| [`analytics.md`](analytics.md) | `analytics_bp` | `/admin/analytics` |
| [`photo_bulk.md`](photo_bulk.md) | `photo_bulk_bp` | `/admin/photos/bulk/...` |
| [`public_apply.md`](public_apply.md) | `public_apply_bp` | `/apply` (публічний, без входу) |
| [`public_update.md`](public_update.md) | `public_update_bp` | `/update-info/<token>` (публічний) |

### Пакет `routes/admin/` (усі маршрути на спільному `admin_bp`)

[`__init__.md`](admin/__init__.md) · [`catalogs.md`](admin/catalogs.md) ·
[`groups_courses.md`](admin/groups_courses.md) ·
[`subjects.md`](admin/subjects.md) ·
[`documents.md`](admin/documents.md) ·
[`applications.md`](admin/applications.md) ·
[`users.md`](admin/users.md) ·
[`templates_and_export.md`](admin/templates_and_export.md)

### Пакет `routes/students/` (усі маршрути на спільному `students_bp`)

[`__init__.md`](students/__init__.md) · [`core.md`](students/core.md) ·
[`military.md`](students/military.md) · [`grades.md`](students/grades.md) ·
[`import_export.md`](students/import_export.md)

### Спільні модулі (не Blueprint'и, використовуються всіма іншими)

| Файл | Для чого |
|---|---|
| [`db.md`](db.md) | `get_db()` - єдина точка підключення до SQLite |
| [`utils.md`](utils.md) | Логування, права доступу, вкладення, скорочена програма |
| [`helpers.md`](helpers.md) | `current_username()`, `sort_ukrainian()` |
| [`config.md`](config.md) | `SECRET_KEY`, ONLYOFFICE-налаштування, публічні URL |
| [`photo.md`](photo.md) | Обрізка/валідація фото 3х4 |
| [`gen_docx.md`](gen_docx.md) | Генерація `.docx` з шаблону |

### Самостійні скрипти (`app.py` їх не запускає)

| Файл | Коли запускається |
|---|---|
| [`init_db.md`](init_db.md) | Один раз, при розгортанні на новому сервері |
| [`update_groups.md`](update_groups.md) | Щодня, за розкладом (Task Scheduler/cron) |

## Загальна схема залежностей

```
app.py
 ├── routes/config.py          (SECRET_KEY, PREFERRED_URL_SCHEME)
 ├── routes/auth.py            (Blueprint: auth_bp)
 ├── routes/students/          (Blueprint: students_bp, пакет із 4 файлів)
 ├── routes/admin/             (Blueprint: admin_bp, пакет із 7 файлів)
 ├── routes/office_editor.py   (Blueprint: office_bp)
 ├── routes/import_grades.py   (Blueprint: import_grades_bp)
 ├── routes/analytics.py       (Blueprint: analytics_bp)
 ├── routes/photo_bulk.py      (Blueprint: photo_bulk_bp)
 ├── routes/public_apply.py    (Blueprint: public_apply_bp - ізольований)
 ├── routes/public_update.py   (Blueprint: public_update_bp - ізольований)
 └── routes/gen_docx.py        (лише format_grade, для Jinja-фільтра)

Усі Blueprint'и (крім public_apply.py/public_update.py) спираються на:
 ├── routes/db.py       (get_db)
 ├── routes/utils.py    (log_action, permission_required, ...)
 ├── routes/helpers.py  (current_username, sort_ukrainian)
 └── routes/photo.py    (обрізка фото - де є завантаження фото)

routes/init_db.py та routes/update_groups.py - поза цією схемою,
самостійні скрипти без зв'язку з app.py.
```
