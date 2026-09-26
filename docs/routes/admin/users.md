# `routes/admin/users.py`

## Призначення
Користувачі системи та їхні права доступу (кожне право - окремий
чекбокс, список усіх можливих прав - константа `PERMISSIONS` у цьому
ж файлі), журнал дій (`/admin/view_logs` - розбирає `app.log` у
зручний для перегляду вигляд з фільтрацією на фронтенді).

## Взаємодія з app.py
Непряма (через `admin_bp`, див. `routes/admin/__init__.py`). Список
`PERMISSIONS`, визначений тут, - джерело істини для того, які рядки
взагалі можуть з'явитись у `session['permissions']` після входу
(перевіряється в `routes/auth.py` опосередковано, через
`routes.utils.permission_required`, який звіряється з тим, що записано
в сесію, а не напряму з цим списком).

## Маршрути
`/admin/view_logs`, `/admin/users` (список), `/admin/users/add`,
`/admin/users/<id>/edit`, `/admin/users/<id>/change-password`,
`/admin/users/<id>/delete`.

## Залежить від
`werkzeug.security` (`generate_password_hash`), `routes.db`.
