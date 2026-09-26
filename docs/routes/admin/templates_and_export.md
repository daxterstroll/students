# `routes/admin/templates_and_export.py`

## Призначення
Керування `.docx`-шаблонами документів (завантаження нового шаблону,
приховати/показати, видалити), масова генерація документів по групі
(з фоновою задачею й індикатором прогресу - `_JOBS`), масове
вивантаження фото студентів групи одним ZIP-архівом.

## Взаємодія з app.py
Непряма (через `admin_bp`, див. `routes/admin/__init__.py`).

## Ключове - фонові задачі
`_JOBS` (словник активних задач), `_JOBS_LOCK` (`threading.Lock`),
`_JOB_TTL_SECONDS` - масова генерація документів по великій групі може
тривати довго, тому запускається у фоновому потоці
(`_run_generation_job`), а фронтенд періодично опитує прогрес через
`/admin/generate_group_docs/status/<job_id>`. `_cleanup_old_jobs()`
прибирає застарілі завершені задачі зі словника, щоб він не ріс
безмежно при активному використанні сторінки.

## Маршрути
`/admin/templates` (+ `/toggle_visibility`, `/download`, `/delete`),
`/admin/export_photos`, `/admin/group_export`,
`/admin/generate_group_docs` (+ `/status/<job_id>`, `/result/<job_id>`).

## Залежить від
`routes.gen_docx` (`gen_doc`), `routes.office_editor` (реєстрація
щойно згенерованих файлів для перегляду в браузері), `routes.utils`
(`get_templates_with_metadata`, `TEMPLATE_FOLDER`), `zipfile`,
`threading`.
