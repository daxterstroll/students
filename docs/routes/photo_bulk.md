# `routes/photo_bulk.py`

## Призначення
Масове завантаження фото студентів: багато файлів одразу, кожен файл
зіставляється зі студентом за іменем у назві файлу
("Прізвище Ім'я.jpg", нечітке порівняння через `rapidfuzz`), обрізається
автоматично по центру до 3х4, з можливістю вручну підправити обрізку
конкретного фото на кроці попереднього перегляду перед остаточним
підтвердженням.

## Взаємодія з app.py
Прямо: `from routes.photo_bulk import photo_bulk_bp` +
`app.register_blueprint(photo_bulk_bp)`.

## Маршрути
- `GET, POST /admin/photos/bulk` - `bulk_upload()`: вибір групи +
  завантаження файлів.
- `GET, POST /admin/photos/bulk/<token>/match` - `match()`: перевірка/
  виправлення автоматичного зіставлення файл→студент.
- `GET /admin/photos/bulk/<token>/preview` - `preview()`: попередній
  перегляд усіх обрізаних фото.
- `GET, POST /admin/photos/bulk/<token>/recrop/<file_key>` - `recrop()`:
  ручна переобрізка одного конкретного фото.
- `GET /admin/photos/bulk/<token>/preview_image/<file_key>` -
  `preview_image()`: віддає саме зображення для попереднього перегляду.
- `POST /admin/photos/bulk/<token>/confirm` - `confirm()`: остаточний
  запис усіх фото на картки студентів.

Стан майстра - у тимчасових файлах на диску за `token` (той самий
підхід, що й у `routes/import_grades.py`), не в cookie-сесії.

## Залежить від
`routes.photo` (обрізка/валідація), `rapidfuzz` (зіставлення імен).
