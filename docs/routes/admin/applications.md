# `routes/admin/applications.py`

## Призначення
Модерація двох типів заявок, що приходять з публічних, неавторизованих
сторінок: заявки на реєстрацію (`routes/public_apply.py`, `/apply`) і
заявки на оновлення даних наявного студента (`routes/public_update.py`,
`/update-info/<token>`). В обох випадках - вибіркове застосування
полів (адмін вирішує, яке саме поле з заявки перенести в реальні
таблиці), ручна переобрізка фото (Cropper.js, той самий підхід для
обох типів заявок), виявлення ймовірних дублікатів (той самий студент
уже є в базі).

## Взаємодія з app.py
Непряма (через `admin_bp`, див. `routes/admin/__init__.py`). Дані, що
сюди потрапляють, ніколи не приходять напряму з `app.py` - лише через
таблиці `pending_students`/`update_requests`, заповнені окремими
ізольованими блюпринтами `routes/public_apply.py` і
`routes/public_update.py`.

## Маршрути
- `/admin/pending_students` (список) + `/admin/pending_students/<id>`
  (перегляд/підтвердження/відхилення/видалення заявки на реєстрацію) +
  `/admin/pending_students/<id>/photo` (переобрізка).
- `/admin/update_requests` (список) + `/admin/update_requests/<id>`
  (те саме для заявок на оновлення даних) +
  `/admin/update_requests/<id>/photo` (переобрізка).

## Залежить від
`routes.db`, `routes.utils` (`get_attachments`,
`save_multiple_attachments`), `routes.photo`
(`process_and_save_pending_photo_with_crop` - спільна з
`routes/public_apply.py`/`routes/public_update.py` функція обрізки).
