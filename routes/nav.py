"""
routes/nav.py
=============
Єдине джерело істини для структури головного меню (навбар).

Раніше меню було продубльоване в layout.html двічі (десктоп + мобільна
версія) - ~180 рядків, які легко розсинхронізувати при додаванні нового
пункту. Тепер меню описане ОДИН раз тут, а layout.html рендерить його
циклом в обох варіантах - додати пункт означає додати один dict у список.

Кожна секція - або 'dropdown' (кілька пунктів під однією кнопкою),
або 'link' (пряме посилання на навбарі, без випадаючого меню).

Пункт у dropdown може мати:
  - 'perm'      - назва дозволу з routes/admin/users.py::PERMISSIONS;
                  без нього пункт не показується не-адмінам
  - 'badge_key' - ключ у nav_badges (див. inject_nav_context у app.py) -
                  числове значення показується бейджем (напр. кількість
                  нових заявок)
  - 'divider'   - якщо True - горизонтальний роздільник замість пункту
"""

NAV_SECTIONS = [
    # ─── Імпорт ─────────────────────────────────────────────────
    {
        'type': 'dropdown',
        'key': 'import',
        'label': 'Імпорт',
        'icon': 'bi-cloud-arrow-up-fill',
        'items': [
            {'perm': 'import_from_excel', 'url': 'students.import_from_excel',
             'icon': 'bi-person-add', 'label': 'Студенти + військові дані'},
            {'perm': 'manage_students', 'url': 'students.military_bulk_import_pdf',
             'icon': 'bi-file-earmark-pdf', 'label': 'PDF «Резерв+» (пакетно)'},
            {'perm': 'manage_students', 'url': 'photo_bulk.bulk_upload',
             'icon': 'bi-images', 'label': 'Фото студентів (масово)'},
            {'perm': 'import_subjects', 'url': 'admin.import_subjects',
             'icon': 'bi-journal-plus', 'label': 'Предмети'},
            {'perm': 'import_grades', 'url': 'import_grades.upload',
             'icon': 'bi-card-checklist', 'label': 'Оцінки (зведена відомість)'},
            {'perm': 'import_education_docs', 'url': 'admin.import_docs_preview',
             'icon': 'bi-file-earmark-arrow-up', 'label': 'Документи студентів'},
            {'perm': 'import_passport_documents', 'url': 'admin.import_passport_documents',
             'icon': 'bi-person-vcard', 'label': 'Паспортні дані'},
        ],
    },

    # ─── Навчальний процес ──────────────────────────────────────
    {
        'type': 'dropdown',
        'key': 'academic',
        'label': 'Навчальний процес',
        'icon': 'bi-mortarboard-fill',
        'items': [
            {'perm': 'manage_pending_students', 'url': 'admin.pending_students',
             'icon': 'bi-person-plus', 'label': 'Заявки на реєстрацію',
             'badge_key': 'nav_pending_applications'},
            {'perm': 'manage_students', 'url': 'admin.update_requests',
             'icon': 'bi-link-45deg', 'label': 'Заявки на оновлення даних',
             'badge_key': 'nav_pending_updates'},
            {'perm': 'manage_groups', 'url': 'admin.manage_groups',
             'icon': 'bi-person-rolodex', 'label': 'Групи'},
            {'perm': 'manage_courses', 'url': 'admin.courses',
             'icon': 'bi-kanban', 'label': 'Курси'},
            {'perm': 'manage_subjects', 'url': 'admin.manage_subjects',
             'icon': 'bi-journal', 'label': 'Предмети'},
            {'perm': 'manage_activities', 'url': 'admin.manage_activities',
             'icon': 'bi-journal-bookmark-fill', 'label': 'Практика / атестація / курсові'},
            {'perm': 'manage_education_documents', 'url': 'admin.manage_education_documents',
             'icon': 'bi-file-earmark-text', 'label': 'Документи освіти'},
            {'perm': 'manage_passport_documents', 'url': 'admin.manage_passport_documents',
             'icon': 'bi-person-vcard', 'label': 'Паспортні дані'},
            {'perm': 'manage_diplomas', 'url': 'admin.manage_diplomas',
             'icon': 'bi-journal-bookmark', 'label': 'Номери диплому / додатку'},
            {'perm': 'study_periods', 'url': 'admin.manage_study_periods_bulk_assign',
             'icon': 'bi-calendar2-date', 'label': 'Періоди навчання'},
            {'divider': True},
            {'perm': 'manage_frozen_students', 'url': 'admin.frozen_students',
             'icon': 'bi-snow', 'label': 'Заморожені студенти'},
            {'perm': 'manage_expulsion', 'url': 'admin.expulsion',
             'icon': 'bi-person-x', 'label': 'Наказ про відрахування'},
            {'perm': 'archive', 'url': 'admin.archive',
             'icon': 'bi-archive', 'label': 'Архів груп'},
        ],
    },

    # ─── Довідники та система ───────────────────────────────────
    {
        'type': 'dropdown',
        'key': 'catalogs',
        'label': 'Довідники та система',
        'icon': 'bi-diagram-3-fill',
        'items': [
            {'perm': 'manage_specialties', 'url': 'admin.manage_specialties',
             'icon': 'bi-diagram-3', 'label': 'Спеціальності'},
            {'perm': 'manage_degree_levels', 'url': 'admin.manage_degree_levels',
             'icon': 'bi-mortarboard', 'label': 'Ступені'},
            {'perm': 'manage_educational_programs', 'url': 'admin.manage_educational_programs',
             'icon': 'bi-journal-bookmark', 'label': 'Освітня програма'},
            {'perm': 'manage_qualification_names', 'url': 'admin.manage_qualification_names',
             'icon': 'bi-mortarboard-fill', 'label': 'Назва кваліфікації'},
            {'perm': 'manage_licenses', 'url': 'admin.manage_licenses',
             'icon': 'bi-award', 'label': 'Ліцензії'},
            {'perm': 'manage_accreditations', 'url': 'admin.manage_accreditations',
             'icon': 'bi-award', 'label': 'Акредитації'},
            {'divider': True},
            {'perm': 'manage_users', 'url': 'admin.manage_users',
             'icon': 'bi-people', 'label': 'Користувачі та права'},
            {'perm': 'manage_templates', 'url': 'admin.manage_templates',
             'icon': 'bi-file-earmark-arrow-up', 'label': 'Шаблони документів'},
            {'perm': 'license_report', 'url': 'admin.license_report',
             'icon': 'bi-list-check', 'label': 'Звіт по ліцензіях'},
            {'perm': 'analytics', 'url': 'analytics.dashboard',
             'icon': 'bi-graph-up', 'label': 'Аналітика'},
            {'perm': 'view_logs', 'url': 'admin.view_logs',
             'icon': 'bi-file-earmark-binary', 'label': 'Логування'},
        ],
    },

    # ─── Прямі посилання (без dropdown) ─────────────────────────
    {
        'type': 'link',
        'key': 'group_export',
        'label': 'Масова генерація документів',
        'icon': 'bi-file-zip',
        'url': 'admin.group_export',
        'perm': 'group_export',
    },
    {
        'type': 'link',
        'key': 'export_photos',
        'label': 'Фото студентів',
        'icon': 'bi-images',
        'url': 'admin.export_photos',
        'perm': 'group_export',
    },
]