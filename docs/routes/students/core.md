# `routes/students/core.py`

## Призначення
Найбільший і найважливіший підмодуль - основний робочий простір
користувача: список студентів (пошук/фільтри/пагінація), картка
студента, додавання/редагування/видалення, фото 3х4 (з інтерактивною
обрізкою Cropper.js), заморозка студента, періоди навчання (філії),
генерація токен-посилання `/update-info` (кнопка на картці студента),
генерація документа для одного студента, сторінка FAQ.

## Взаємодія з app.py
Непряма (через `students_bp`, див. `routes/students/__init__.py`).

## Ключова константа
`UPDATE_REQUEST_FIELD_LABELS` - список полів, які можна дозволити
редагувати студенту через `/update-info` (використовується у формі
вибору полів на `generate_update_link()` і показується на сторінці
модерації в `routes/admin/applications.py`).

## Маршрути
`/faq`, `/students`, `/students/<id>`, `/students/<id>/study_periods`,
`/students/add`, `/students/<id>/photo` (+ `/delete`),
`/students/<id>/freeze`, `/students/<id>/edit`, `/students/<id>/delete`,
`/students/<id>/generate_update_link`, `/students/<id>/generate`.

## Залежить від
`routes.utils` (`generate_english_name`, `is_student_on_reduced_program`,
`save_multiple_attachments`, `program_track_condition`,
`apply_track_overrides`, `get_attachments`, `get_templates_with_metadata`),
`routes.gen_docx` (`gen_doc`), `routes.office_editor`, `routes.config`
(`PUBLIC_APPLY_BASE_URL` - для посилання `/update-info`, яке має вести
на публічну адресу, а не на внутрішній домен адмінки).
