# `routes/admin/groups_courses.py`

## Призначення
Групи студентів (створення, редагування, kanban-дошка курсів),
переведення групи на наступний курс, архів груп, оформлення випуску,
відрахування, заморожені студенти (тимчасово поза активним навчанням).
Містить дві важливі спільні функції обчислення курсу групи, якими
більше ніхто в проєкті не користується (окрім самого цього файлу).

## Взаємодія з app.py
Непряма (через `admin_bp`, див. `routes/admin/__init__.py`).

## Ключові функції
- `compute_program_total_years(degree_level, credits)` - тривалість
  програми в роках (4 для бакалавра на 240 кредитів, 2 для магістра
  на 90/120 тощо).
- `compute_current_course(start_year, degree_level, credits)` -
  обчислює поточний курс групи з року вступу, щоб не вводити його
  вручну при створенні групи (та сама логіка тривалості, що й
  `routes.analytics._compute_end_year`, навмисно синхронізовано).

## Маршрути
`/admin/manage_groups`, `/admin/courses`, `/admin/course_transfer`
(+`/confirm`), `/admin/frozen_students` (+ `/resolve`, `/resolve_confirm`),
`/admin/graduation` (+`/confirm`), `/admin/expulsion` (+`/confirm`),
`/admin/archive/<id>`, `/admin/unarchive_group/<id>`, `/admin/archive`.

## Залежить від
`routes.db`, `routes.utils` (`get_attachments`,
`save_multiple_attachments` - скани наказів про переведення/
відрахування/заморозку), `difflib` (нечітке зіставлення оцінок при
переведенні студента між групами з різним переліком предметів).
