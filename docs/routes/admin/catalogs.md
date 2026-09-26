# `routes/admin/catalogs.py`

## Призначення
Довідникові сторінки, спільна риса - прості CRUD-таблиці, на які
посилаються групи студентів: спеціальності, ступені, освітні програми,
назви кваліфікацій, акредитації, ліцензії (+ переведення студентів між
ліцензіями, звіт по ліцензіях з експортом у Word), дипломи, масове
призначення періодів навчання.

## Взаємодія з app.py
Непряма - маршрути реєструються на `admin_bp` (визначеному в
`routes/admin/__init__.py`), який `app.py` реєструє як єдиний
блюпринт. Цей конкретний файл `app.py` не імпортує і не знає про нього.

## Маршрути
`/admin/study_periods/bulk_assign`, `/admin/manage_diplomas`,
`/admin/manage_accreditations`, `/admin/manage_specialties`,
`/admin/manage_degree_levels`, `/admin/manage_licenses`,
`/admin/license_transfer` (+`/confirm`), `/admin/license_report`
(+`/export_word`), `/admin/manage_educational_programs`,
`/admin/manage_qualification_names`.

## Залежить від
`routes.db`, `routes.utils` (`log_action`, `permission_required`),
`routes.helpers` (`sort_ukrainian`).
