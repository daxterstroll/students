import sqlite3
from datetime import datetime
import logging
import re
import os


# ============================================================
# Пути проекта
# ============================================================

# Корневая папка students
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


# База данных
DB_PATH = os.path.join(BASE_DIR, "students.db")


# Файл логирования
LOG_PATH = os.path.join(BASE_DIR, "group_update.log")


# ============================================================
# Настройка логирования
# ============================================================

logging.basicConfig(
    filename=LOG_PATH,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    encoding="utf-8"
)


# ============================================================
# Обновление групп
# ============================================================

def update_groups():
    conn = None
    try:
        logging.info("========================================")
        logging.info("Запуск проверки обновления групп")
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        today = datetime.now()
        current_year = today.year
        # Обновление только 1 сентября
        if today.month != 9 or today.day != 1:
            logging.info(
                f"Сегодня {today.strftime('%d.%m.%Y')}. "
                f"Обновление групп не требуется."
            )
            return
        logging.info(
            f"Начато обновление групп на 1 сентября {current_year}"
        )
        # Получаем активные группы
        cursor.execute("""
            SELECT 
                id,
                name,
                course,
                start_year,
                program_credits
            FROM groups
            WHERE archived = 0
        """)
        groups = cursor.fetchall()
        logging.info(
            f"Найдено активных групп: {len(groups)}"
        )
        for group in groups:
            group_id, name, course, start_year, program_credits = group
            # Определяем максимальный курс
            max_courses = 4 if program_credits == 240 else 3
            # Текущий курс по году набора
            current_academic_year = (
                current_year - start_year + 1
            )
            # Если группа должна перейти на следующий курс
            if current_academic_year == course + 1:
                new_course = course + 1
                # Если обучение завершено
                if new_course > max_courses:
                    cursor.execute("""
                        UPDATE groups
                        SET archived = 1
                        WHERE id = ?
                    """, (group_id,))
                    logging.info(
                        f"Группа {name} архивирована. "
                        f"Курс: {new_course}, "
                        f"start_year={start_year}, "
                        f"credits={program_credits}"
                    )
                else:
                    # ----------------------------------------
                    # Формирование нового имени группы
                    # ----------------------------------------
                    # Например:
                    # КН-11 -> КН-21
                    # ЕП -> ЕП-11
                    if re.match(
                        r'^[А-Яа-яЄєІіЇїҐґA-Za-z]+$',
                        name
                    ) and "-" not in name:
                        prefix = name + "-"
                        current_course_digit = 1
                    else:
                        prefix = (
                            name.rsplit("-", 1)[0]
                            + "-"
                        )
                        try:
                            current_course_digit = int(
                                name.split("-")[-1][0]
                            )
                        except Exception:
                            logging.warning(
                                f"Не удалось определить курс "
                                f"для группы {name}. Пропуск."
                            )
                            continue
                    new_course_digit = current_course_digit + 1
                    new_name = (
                        f"{prefix}"
                        f"{new_course_digit}1"
                    )
                    cursor.execute("""
                        UPDATE groups
                        SET 
                            name = ?,
                            course = ?
                        WHERE id = ?
                    """,
                    (
                        new_name,
                        new_course,
                        group_id
                    ))
                    logging.info(
                        f"Группа {name} обновлена -> {new_name}. "
                        f"Курс: {new_course}, "
                        f"start_year={start_year}, "
                        f"credits={program_credits}"
                    )
        conn.commit()
        logging.info(
            "Обновление групп успешно завершено"
        )
    except Exception as e:

        logging.exception(
            f"Ошибка при обновлении групп: {e}"
        )
    finally:
        if conn:
            conn.close()

# ============================================================
# Запуск напрямую
# ============================================================

if __name__ == "__main__":
    update_groups()