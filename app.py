import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="docxcompose")
from flask import Flask
from config import SECRET_KEY
from routes.auth import auth_bp
from routes.students import students_bp
from routes.admin import admin_bp
import argparse
from utils import log_action
from gen_docx import format_grade

# Инициализация приложения Flask
app = Flask(__name__)
app.secret_key = SECRET_KEY

# Регистрация blueprint'ов
app.register_blueprint(auth_bp)
app.register_blueprint(admin_bp)
app.register_blueprint(students_bp)
app.jinja_env.filters['format_grade'] = format_grade

# --- Запуск приложения ---
if __name__ == '__main__':
    """
    Запускает Flask-приложение в зависимости от выбранного режима.
    
    Цель:
        Позволяет выбрать режим запуска (debug или production) через аргумент командной строки --mode.
    
    Аргументы командной строки:
        --mode (str): Режим запуска ('debug' для локального сервера, 'production' для waitress).
        --host (str): Хост для production-режима (по умолчанию '192.168.0.219').
        --port (int): Порт для production-режима (по умолчанию 8080).
    
    Логика:
        1. Парсит аргументы командной строки.
        2. Если mode='debug', запускает Flask в режиме отладки.
        3. Если mode='production', использует waitress для продакшн-среды.
        4. Логирует выбранный режим запуска.
        5. По умолчанию использует debug-режим, если аргумент не указан.
    
    Пример использования:
        python app.py --mode debug
        python app.py --mode production --host 0.0.0.0 --port 5000
    """
    parser = argparse.ArgumentParser(description='Запуск Flask-приложения в указанном режиме.')
    parser.add_argument('--mode', choices=['debug', 'production'], default='debug',
                        help='Режим запуска: debug (локальный сервер) или production (waitress)')
    parser.add_argument('--host', default='localhost',
                        help='Хост для production-режима (по умолчанию: 192.168.0.219)')
    parser.add_argument('--port', type=int, default=5000,
                        help='Порт для production-режима (по умолчанию: 5000)')
    args = parser.parse_args()

    if args.mode == 'debug':
        app.run(debug=True)
    else:  # args.mode == 'production'
        from waitress import serve
        serve(app, host=args.host, port=args.port)