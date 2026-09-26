# `routes/public_apply.py`

## Призначення
Публічна анкета самореєстрації абітурієнта (`/apply`) - без входу в
систему. Свідомо ІЗОЛЬОВАНА від решти застосунку: жодного імпорту з
`routes/admin/*` чи `routes/students/*`, пише виключно в "карантинну"
таблицю `pending_students` (без зовнішніх ключів на `groups`/`students`,
крім `resulting_student_id`, який проставляється вже ПІСЛЯ ручного
підтвердження адміном на `/admin/pending_students`). Захист від спаму -
honeypot-поле, без капчі й токенів (посилання публічне й незмінне).

## Взаємодія з app.py
Прямо: `from routes.public_apply import public_apply_bp` +
`app.register_blueprint(public_apply_bp)`. Це один з двох маршрутів
(другий - `/update-info/<token>`), які на сервері віддаються назовні в
інтернет через окремий, обмежений `location` в `nginx.conf` (порт 80,
без TLS-редіректу) - решта сайту звідти недосяжна.

## Маршрути
- `GET, POST /apply` - `apply()`: єдиний маршрут, і форма, і обробка
  відправки. Особисті дані, контакти, документ про освіту, паспортні
  дані, військовий облік - кожен розділ з опційними сканами (до 5
  файлів, окремий `entity_type` в `attachments` для кожного розділу,
  щоб адмін при підтвердженні знав, куди що переприв'язати).

## Залежить від
`routes.db`, `routes.utils` (`generate_english_name`,
`save_multiple_attachments`), `routes.photo`
(`process_and_save_pending_photo`).
