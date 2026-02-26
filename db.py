import sqlite3
import locale

def ukrainian_collation(str1, str2):
    """
    Пользовательская функция колляции для сортировки украинских строк.
    """
    locale.setlocale(locale.LC_COLLATE, 'uk_UA.UTF-8')
    return locale.strcoll(str1, str2)

def get_db():
    """
    Создает соединение с базой данных SQLite.
    """
    conn = sqlite3.connect('students.db')
    conn.row_factory = sqlite3.Row
    conn.create_collation("UKRAINIAN", ukrainian_collation)
    conn.execute("PRAGMA foreign_keys = ON")
    return conn