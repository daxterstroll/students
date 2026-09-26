# `routes/auth.py`

## Призначення
Вхід/вихід користувачів. Кожна спроба входу (успішна чи ні) і кожен
вихід записується через `log_action()` в `app.log` - зокрема, для
розслідування підозрілих спроб входу за IP-адресою (див. коментар про
`ProxyFix` в `app.py` - без нього тут завжди був би `127.0.0.1`).

## Взаємодія з app.py
Прямо: `from routes.auth import auth_bp` + `app.register_blueprint(auth_bp)`.
Це єдиний блюпринт, що обробляє `/` (головна сторінка - редірект на
логін або на список студентів) і `/login`/`/logout`.

## Маршрути
- `GET /` - `index()`: якщо є сесія - редірект на `/students`, інакше
  на `/login`.
- `GET, POST /login` - `login()`: форма входу; перевіряє пароль через
  `werkzeug.security.check_password_hash`, записує успіх/невдачу в лог.
- `GET /logout` - `logout()`: очищує сесію, записує вихід у лог.

## Залежить від
`routes.db` (пошук користувача), `routes.utils` (`log_action`, `logger`),
`routes.helpers` (`current_username`), `werkzeug.security`.
