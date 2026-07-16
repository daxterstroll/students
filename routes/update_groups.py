import sqlite3
from datetime import datetime
import logging
import re  # Для проверки, состоит ли строка только из букв

# Настройка логирования
logging.basicConfig(filename='group_update.log', level=logging.INFO, format='%(asctime)s - %(message)s')

# Подключение к БД (замените на путь к вашей БД)
DB_PATH = 'students.db'  # Укажите путь к SQLite-файлу

def update_groups():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()


    # today = datetime(2025, 9, 1)
    today = datetime.now()
    current_year = today.year
    if today.month == 9 and today.day == 1:  # Проверка на 1 сентября
        logging.info(f"Запуск обновления групп на 1 сентября {current_year}.")

        # Получаем все активные группы
        cursor.execute("""
            SELECT id, name, course, start_year, program_credits
            FROM groups
            WHERE archived = FALSE
        """)
        groups = cursor.fetchall()

        for group in groups:
            group_id, name, course, start_year, program_credits = group
            max_courses = 4 if program_credits == 240 else 3  # 4 года для 240, 3 года для 180
            current_academic_year = current_year - start_year + 1  # Текущий учебный год относительно start_year

            if current_academic_year == course + 1:  # Если текущий год соответствует следующему курсу
                new_course = course + 1

                if new_course > max_courses:
                    # Архивируем группу
                    cursor.execute("UPDATE groups SET archived = TRUE WHERE id = ?", (group_id,))
                    logging.info(f"Группа {name} (start_year={start_year}, credits={program_credits}) архивирована (курс: {new_course}).")
                else:
                    # Проверка, состоит ли имя только из букв
                    if re.match(r'^[А-Яа-яЄєІіЇїҐґA-Za-z]+$', name) and '-' not in name:
                        prefix = name + '-'  # Если только буквы, используем как префикс
                        current_course_digit = 1  # Начинаем с 1-го курса
                    else:
                        prefix = name.rsplit('-', 1)[0] + '-'  # Извлекаем префикс, например "КН-"
                        current_course_digit = int(name.split('-')[-1][0])  # Текущая цифра курса (например, "1" из "КН-11")

                    new_course_digit = current_course_digit + 1  # Увеличиваем курс (1 → 2)
                    new_name = f"{prefix}{new_course_digit}1"  # Формируем новое имя, например "КН-21" или "ЕП-21"

                    cursor.execute("""
                        UPDATE groups SET name = ?, course = ? WHERE id = ?
                    """, (new_name, new_course, group_id))
                    logging.info(f"Группа {name} обновлена на {new_name} (курс: {new_course}, start_year={start_year}, credits={program_credits}).")

        conn.commit()
    else:
        logging.info(f"Не 1 сентября {current_year}, обновление не требуется.")

    conn.close()

if __name__ == '__main__':
    update_groups()