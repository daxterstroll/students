# `routes/public_update.py`

## Призначення
Публічна сторінка "оновити мої дані" за одноразовим токен-посиланням
(`/update-info/<token>`) - без входу в систему. На відміну від
`/apply` (повністю анонімна форма), тут посилання прив'язане до
КОНКРЕТНОГО, вже наявного студента - тому головна межа безпеки не
honeypot, а токен (непередбачуваний, `secrets.token_urlsafe`) + суворий
allowlist полів на сервері (`update_requests.allowed_fields`): навіть
якщо хтось руками допише в POST-запит зайве поле, буде прийнято лише
те, що адмін дозволив явно при створенні посилання. Токен дійсний 1
добу або до першого успішного заповнення - що настане раніше.

## Взаємодія з app.py
Прямо: `from routes.public_update import public_update_bp` +
`app.register_blueprint(public_update_bp)`. Посилання генерується в
`routes/students/core.py` (`generate_update_link()`, кнопка на картці
студента), а заповнене студентом лягає на модерацію в
`routes/admin/applications.py` (`update_request_review()`).

## Маршрути
- `GET, POST /update-info/<token>` - `update_info()`: показує лише
  дозволені (`allowed_fields`) поля, приймає відправку, позначає
  `status='submitted'` (не застосовує зміни одразу - чекає на ручне
  підтвердження адміном).

## Залежить від
`routes.db`, `routes.utils` (`save_multiple_attachments`),
`routes.photo` (`process_and_save_pending_photo`).
