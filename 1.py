# -*- coding: utf-8 -*-
"""
Оновлена міграція каталогу спеціальностей (v2 - з англійськими
назвами name_en для спеціальностей і галузей знань). Безпечна для
запуску на вже наявній базі - не чіпає students/groups/grades тощо,
крім додавання нової колонки groups.specialty_code.

Запуск:  python migrate_specialties.py
"""
import sqlite3
import os

db_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'students.db')
conn = sqlite3.connect(db_path)
cur = conn.cursor()

cur.executescript("""
CREATE TABLE IF NOT EXISTS "knowledge_fields" (
    code TEXT PRIMARY KEY,      -- 'A', 'B', 'C' ... (літерні коди, чинні з 01.11.2024)
    name_ua TEXT NOT NULL,
    name_en TEXT                -- офіційного англ. відповідника немає в постанові - переклад можна відредагувати на сторінці "Спеціальності"
);

CREATE TABLE IF NOT EXISTS "specialties" (
    code TEXT PRIMARY KEY,      -- 'A1', 'D2', 'D3', 'F3' ... (літера галузі + номер)
    name_ua TEXT NOT NULL,
    name_en TEXT,               -- код і назва відповідної деталізованої галузі ISCED-F 2013 з постанови, напр. "0413 Management and administration"
    knowledge_field_code TEXT NOT NULL,
    is_active INTEGER NOT NULL DEFAULT 1,   -- чи актуальна для цього закладу (керується на сторінці "Спеціальності")
    is_custom INTEGER NOT NULL DEFAULT 0,   -- 1 = додано вручну адміністратором, не з офіційного переліку
    FOREIGN KEY (knowledge_field_code) REFERENCES knowledge_fields(code)
);
""")

# Якщо таблиці вже існували з попередньої версії міграції (без
# name_en) - додаємо колонки окремо, executescript вище створює їх
# лише "з нуля" (CREATE TABLE IF NOT EXISTS не додає колонки в наявну
# таблицю).
for table, column in [('knowledge_fields', 'name_en'), ('specialties', 'name_en')]:
    try:
        cur.execute(f"ALTER TABLE {table} ADD COLUMN {column} TEXT")
        print(f"Додано колонку {table}.{column}")
    except sqlite3.OperationalError as e:
        if "duplicate column" in str(e):
            print(f"Колонка {table}.{column} вже існує - пропускаю")
        else:
            raise

cur.executescript("""
INSERT OR IGNORE INTO knowledge_fields (code, name_ua, name_en) VALUES ('A', 'Освіта', 'Education');
INSERT OR IGNORE INTO knowledge_fields (code, name_ua, name_en) VALUES ('B', 'Культура, мистецтво та гуманітарні науки', 'Culture, Arts and Humanities');
INSERT OR IGNORE INTO knowledge_fields (code, name_ua, name_en) VALUES ('C', 'Соціальні науки, журналістика та інформація', 'Social Sciences, Journalism and Information');
INSERT OR IGNORE INTO knowledge_fields (code, name_ua, name_en) VALUES ('D', 'Бізнес, адміністрування та право', 'Business, Administration and Law');
INSERT OR IGNORE INTO knowledge_fields (code, name_ua, name_en) VALUES ('E', 'Природничі науки, математика та статистика', 'Natural Sciences, Mathematics and Statistics');
INSERT OR IGNORE INTO knowledge_fields (code, name_ua, name_en) VALUES ('F', 'Інформаційні технології', 'Information Technology');
INSERT OR IGNORE INTO knowledge_fields (code, name_ua, name_en) VALUES ('G', 'Інженерія, виробництво та будівництво', 'Engineering, Manufacturing and Construction');
INSERT OR IGNORE INTO knowledge_fields (code, name_ua, name_en) VALUES ('H', 'Сільське, лісове, рибне господарство та ветеринарна медицина', 'Agriculture, Forestry, Fisheries and Veterinary Medicine');
INSERT OR IGNORE INTO knowledge_fields (code, name_ua, name_en) VALUES ('I', 'Охорона здоров’я та соціальне забезпечення', 'Health and Welfare');
INSERT OR IGNORE INTO knowledge_fields (code, name_ua, name_en) VALUES ('J', 'Транспорт та послуги', 'Transport and Services');
INSERT OR IGNORE INTO knowledge_fields (code, name_ua, name_en) VALUES ('K', 'Безпека та оборона', 'Security and Defence');
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('A1', 'Освітні науки', '0111 Education science', 'A', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('A2', 'Дошкільна освіта', '0112 Training for pre-school teachers', 'A', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('A3', 'Початкова освіта', '0113 Teacher training without subject specialisation', 'A', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('A4', 'Середня освіта (за предметними спеціальностями)', '0114 Teacher training with subject specialisation', 'A', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('A5', 'Професійна освіта (за спеціалізаціями)', '0114 Teacher training with subject specialisation', 'A', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('A6', 'Спеціальна освіта (за спеціалізаціями)', '0113 Teacher training without subject specialisation', 'A', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('A7', 'Фізична культура і спорт', '1014 Sports', 'A', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('B1', 'Аудіовізуальне мистецтво та медіавиробництво', '0211 Audio-visual techniques and media production', 'B', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('B2', 'Дизайн', '0212 Fashion, interior and industrial design', 'B', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('B3', 'Декоративне мистецтво та ремесла', '0214 Handicrafts', 'B', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('В4', 'Образотворче мистецтво та реставрація', '0213 Fine arts', 'B', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('В5', 'Музичне мистецтво', '0215 Music and performing arts', 'B', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('В6', 'Перформативні мистецтва', '0215 Music and performing arts', 'B', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('B7', 'Релігієзнавство', '0221 Religion and theology', 'B', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('B8', 'Богослов’я', '0221 Religion and theology', 'B', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('B9', 'Історія та археологія', '0222 History and archaeology', 'B', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('B10', 'Філософія', '0223 Philosophy and ethics', 'B', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('B11', 'Філологія (за спеціалізаціями)', '0231 Language acquisition', 'B', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('B12', 'Культурологія та музеєзнавство', '0314 Sociology and cultural studies', 'B', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('B13', 'Бібліотечна, інформаційна та архівна справа', '0322 Library, information and archival studies', 'B', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('B14', 'Організація соціокультурної діяльності', '0413 Management and administration', 'B', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('C1', 'Економіка', '0311 Economics', 'C', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('C2', 'Політологія', '0312 Political sciences and civics', 'C', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('C3', 'Міжнародні відносини', '0312 Political sciences and civics', 'C', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('C4', 'Психологія', '0313 Psychology', 'C', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('C5', 'Соціологія', '0314 Sociology and cultural studies', 'C', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('C6', 'Географія та регіональні студії', '0314 Sociology and cultural studies', 'C', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('C7', 'Журналістика', '0321 Journalism and reporting', 'C', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('D1', 'Облік і оподаткування', '0411 Accounting and taxation', 'D', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('D2', 'Фінанси, банківська справа, страхування та фондовий ринок', '0412 Finance, banking and insurance', 'D', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('D3', 'Менеджмент', '0413 Management and administration', 'D', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('D4', 'Публічне управління та адміністрування', '0413 Management and administration', 'D', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('D5', 'Маркетинг', '0414 Marketing and advertising', 'D', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('D6', 'Секретарська та офісна справа', '0415 Secretarial and office work', 'D', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('D7', 'Торгівля', '0416 Wholesale and retail sales', 'D', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('D8', 'Право', '0421 Law', 'D', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('D9', 'Міжнародне право', '0421 Law', 'D', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('E1', 'Біологія та біохімія', '0511 Biology', 'E', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('E2', 'Екологія', '0521 Environmental sciences', 'E', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('E3', 'Хімія', '0531 Chemistry', 'E', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('E4', 'Науки про Землю', '0532 Earth sciences', 'E', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('E5', 'Фізика та астрономія', '0533 Physics', 'E', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('E6', 'Прикладна фізика та наноматеріали', '0533 Physics', 'E', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('E7', 'Математика', '0541 Mathematics', 'E', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('E8', 'Статистика', '0542 Statistics', 'E', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('F1', 'Прикладна математика', '0541 Mathematics', 'F', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('F2', 'Інженерія програмного забезпечення', '0613 Software and applications development and analysis', 'F', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('F3', 'Комп’ютерні науки', '0613 Software and applications development and analysis', 'F', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('F4', 'Системний аналіз та наука про дані', '0688 Inter-disciplinary programmes and qualifications involving Information and Communication Technologies', 'F', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('F5', 'Кібербезпека та захист інформації', '0612 Database and network design and administration', 'F', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('F6', 'Інформаційні системи і технології', '0612 Database and network design and administration', 'F', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('F7', 'Комп’ютерна інженерія', '0612 Database and network design and administration', 'F', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G1', 'Хімічні технології та інженерія', '0711 Chemical engineering and processes', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G2', 'Технології захисту навколишнього середовища', '0712 Environmental protection technology', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G3', 'Електрична інженерія', '0713 Electricity and energy', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G4', 'Енерговиробництво (за спеціалізацією)', '0713 Electricity and energy', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G5', 'Електроніка, електронні комунікації, приладобудування та радіотехніка', '0714 Electronics and automation', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G6', 'Інформаційно-вимірювальні технології', '0714 Electronics and automation', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G7', 'Автоматизація, комп’ютерно-інтегровані технології та робототехніка', '0714 Electronics and automation', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G8', 'Матеріалознавство', '0788 Inter-disciplinary programmes and qualifications involving engineering, manufacturing and construction', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G9', 'Прикладна механіка', '0715 Mechanics and metal trades', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G10', 'Металургія', '0715 Mechanics and metal trades', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G11', 'Машинобудування (за спеціалізаціями)', '0715 Mechanics and metal trades', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G12', 'Авіаційна та ракетно-космічна техніка', '0716 Motor vehicles, ships and aircraft', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G13', 'Харчові технології', '0721 Food processing', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G14', 'Деревообробні та меблеві технології', '0722 Materials (glass, paper, plastic and wood)', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G15', 'Технології легкої промисловості', '0723 Textiles (clothes, footwear and leather)', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G16', 'Гірництво та нафтогазові технології', '0724 Mining and extraction', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G17', 'Архітектура та містобудування', '0731 Architecture and town planning', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G18', 'Геодезія та землеустрій', '0532 Earth sciences', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G19', 'Будівництво та цивільна інженерія', '0732 Building and civil engineering', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G20', 'Видавництво та поліграфія', '0211 Audio-visual techniques and media production', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G21', 'Біотехнології та біоінженерія', '0588 Inter-disciplinary programmes and qualifications involving natural sciences, mathematics and statistics', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('G22', 'Біомедична інженерія', '0588 Inter-disciplinary programmes and qualifications involving natural sciences, mathematics and statistics', 'G', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('H1', 'Агрономія', '0811 Crop and livestock production', 'H', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('H2', 'Тваринництво', '0811 Crop and livestock production', 'H', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('H3', 'Садово-паркове господарство', '0812 Horticulture', 'H', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('H4', 'Лісове господарство', '0821 Forestry', 'H', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('H5', 'Водні біоресурси та аквакультура', '0831 Fisheries', 'H', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('H6', 'Ветеринарна медицина', '0841 Veterinary', 'H', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('H7', 'Агроінженерія', '0788 Inter-disciplinary programmes and qualifications involving engineering, manufacturing and construction', 'H', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('I1', 'Стоматологія', '0911 Dental studies', 'I', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('I2', 'Медицина', '0912 Medicine', 'I', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('I3', 'Педіатрія', '0912 Medicine', 'I', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('I4', 'Медична психологія', '0912 Medicine', 'I', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('I5', 'Медсестринство 
(за спеціалізаціями)', '0913 Nursing and midwifery', 'I', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('I6', 'Технології медичної діагностики та лікування (за спеціалізаціями)', '0914 Medical diagnostic and treatment technology', 'I', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('I7', 'Терапія та реабілітація (за спеціалізаціями)', '0915 Therapy and rehabilitation', 'I', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('I8', 'Фармація (за спеціалізаціями)', '0711 Chemical engineering and processes', 'I', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('I9', 'Громадське здоров’я', '0988 Inter-disciplinary programmes and qualifications involving health and welfare', 'I', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('I10', 'Соціальна робота та консультування', '0921 Care of the elderly and of disabled adults', 'I', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('I11', 'Дитячі та молодіжні служби', '0922 Child care and youth services', 'I', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('J1', 'Послуги краси', '1012 Hair and beauty services', 'J', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('J2', 'Готельно-ресторанна справа та кейтеринг', '1013 Hotel, restaurants and catering', 'J', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('J3', 'Туризм та рекреація', '1015 Travel, tourism and leisure', 'J', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('J4', 'Охорона праці', '1022 Occupational health and safety', 'J', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('J5', 'Морський та внутрішній водний транспорт', '1041 Transport services', 'J', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('J6', 'Авіаційний транспорт', '1041 Transport services', 'J', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('J7', 'Залізничний транспорт', '1041 Transport services', 'J', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('J8', 'Автомобільний транспорт', '1041 Transport services', 'J', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('K1', 'Державна безпека', '1031 Military and defence', 'K', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('K2', 'Безпека державного кордону', '1031 Military and defence', 'K', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('K3', 'Національна безпека (за окремими сферами забезпечення і видами діяльності)***', '1031 Military and defence', 'K', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('K4', 'Управління інформаційною безпекою', '1031 Military and defence', 'K', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('K5', 'Військове управління (за видами збройних сил)', '1031 Military and defence', 'K', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('K6', 'Забезпечення військ (сил)', '1031 Military and defence', 'K', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('K7', 'Озброєння та військова техніка', '1031 Military and defence', 'K', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('K8', 'Пожежна безпека', '1032 Protection of persons and property', 'K', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('K9', 'Правоохоронна діяльність', '1032 Protection of persons and property', 'K', 1);
INSERT OR IGNORE INTO specialties (code, name_ua, name_en, knowledge_field_code, is_active) VALUES ('K10', 'Цивільна безпека', '1032 Protection of persons and property', 'K', 1);
""")

try:
    cur.execute("ALTER TABLE groups ADD COLUMN specialty_code TEXT")
    print("Додано колонку groups.specialty_code")
except sqlite3.OperationalError as e:
    if "duplicate column" in str(e):
        print("Колонка groups.specialty_code вже існує - пропускаю")
    else:
        raise

conn.commit()

n_fields = cur.execute("SELECT COUNT(*) FROM knowledge_fields").fetchone()[0]
n_specs = cur.execute("SELECT COUNT(*) FROM specialties").fetchone()[0]
n_with_en = cur.execute("SELECT COUNT(*) FROM specialties WHERE name_en IS NOT NULL").fetchone()[0]
print(f"Готово: {n_fields} галузей знань, {n_specs} спеціальностей (з них {n_with_en} з англ. назвою) у каталозі.")

conn.close()