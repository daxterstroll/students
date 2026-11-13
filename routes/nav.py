"""
routes/nav.py
=============
Єдине джерело істини для структури головного меню (навбар).

4 секції + логічні розділювачі всередині кожної:
  Студенти           — особисті справи, заявки, масові операції з картками
  Навчання           — групи, курси, дисципліни, сесії, періоди
  Документи та накази — освіта, паспорт, військовий облік, дипломи, накази
  Система            — довідники, користувачі, звіти, аналітика, імпорти, логи

Навбар на десктопі вміщує 4 пункти при будь-якій ширині >= 992px.
Меню рендериться циклом у layout.html. Якщо після фільтрації за правами
в dropdown лишається рівно 1 пункт — layout.html рендерить його як
пряму посилання (див. патч нижче), тому dropdown «з одним пунктом» не
виникає.
"""

NAV_SECTIONS = [
    # ═══════════════════════════════════════════════════════════
    #  1. СТУДЕНТИ
    # ═══════════════════════════════════════════════════════════
    {
        'type': 'dropdown',
        'key': 'students',
        'label': 'Студенти',
        'icon': 'bi-people-fill',
        'items': [
            {'perm': 'manage_students',          'url': 'students.student_list',
             'icon': 'bi-person-lines-fill',     'label': 'Список студентів'},
            {'perm': 'manage_pending_students',  'url': 'admin.pending_students',
             'icon': 'bi-person-plus',           'label': 'Заявки на реєстрацію',
             'badge_key': 'nav_pending_applications'},
            {'perm': 'manage_update_requests',   'url': 'admin.update_requests',
             'icon': 'bi-link-45deg',            'label': 'Заявки на оновлення даних',
             'badge_key': 'nav_pending_updates'},
            {'divider': True},
            {'perm': 'manage_frozen_students',   'url': 'admin.frozen_students',
             'icon': 'bi-snow',                  'label': 'Заморожені студенти'},
            {'perm': 'archive',                  'url': 'admin.archive',
             'icon': 'bi-archive',               'label': 'Архів груп'},
            {'divider': True},
            {'perm': 'group_export',             'url': 'admin.group_export',
             'icon': 'bi-file-zip',              'label': 'Масова генерація документів'},
            {'perm': 'group_export',             'url': 'admin.export_photos',
             'icon': 'bi-images',                'label': 'Вивантажити фото (ZIP)'},
            {'divider': True},
            {'perm': 'import_from_excel',        'url': 'students.import_from_excel',
             'icon': 'bi-file-earmark-excel',    'label': 'Імпорт студентів (Excel)'},
            {'perm': 'bulk_photos',              'url': 'photo_bulk.bulk_upload',
             'icon': 'bi-cloud-arrow-up',        'label': 'Імпорт фото (масово)'},
        ],
    },

    # ═══════════════════════════════════════════════════════════
    #  2. НАВЧАННЯ
    # ═══════════════════════════════════════════════════════════
    {
        'type': 'dropdown',
        'key': 'academic',
        'label': 'Навчання',
        'icon': 'bi-mortarboard-fill',
        'items': [
            {'perm': 'manage_groups',            'url': 'admin.manage_groups',
             'icon': 'bi-person-rolodex',        'label': 'Групи'},
            {'perm': 'manage_courses',           'url': 'admin.courses',
             'icon': 'bi-kanban',                'label': 'Курси'},
            {'divider': True},
            {'perm': 'manage_subjects',          'url': 'admin.manage_subjects',
             'icon': 'bi-journal',               'label': 'Предмети'},
            {'perm': 'manage_activities',        'url': 'admin.manage_activities',
             'icon': 'bi-journal-bookmark-fill', 'label': 'Практика / атестація / курсові'},
            {'perm': 'study_periods',            'url': 'admin.manage_study_periods_bulk_assign',
             'icon': 'bi-calendar2-date',        'label': 'Періоди навчання'},
            {'divider': True},
            {'perm': 'manage_exam_sessions',     'url': 'exam_sessions.sessions_list',
             'icon': 'bi-clipboard2-check',      'label': 'Сесії та відомості'},
            {'perm': 'transfer_exam_grades',     'url': 'exam_sessions.transfer_queue',
             'icon': 'bi-arrow-left-right',      'label': 'Перенос оцінок із сесій'},
             {'perm': 'view_exam_session_history', 'url': 'exam_sessions.history_index',
             'icon': 'bi-journal-text',          'label': 'Журнал сесій та відомостей'},
            {'divider': True},
            {'perm': 'import_subjects',          'url': 'admin.import_subjects',
             'icon': 'bi-journal-plus',          'label': 'Імпорт предметів (Excel)'},
            {'perm': 'import_grades',            'url': 'import_grades.upload',
             'icon': 'bi-card-checklist',        'label': 'Імпорт оцінок (зведена відомість)'},
        ],
    },

    # ═══════════════════════════════════════════════════════════
    #  3. ДОКУМЕНТИ ТА НАКАЗИ
    # ═══════════════════════════════════════════════════════════
    {
        'type': 'dropdown',
        'key': 'documents',
        'label': 'Документи та накази',
        'icon': 'bi-folder2-open',
        'items': [
            {'perm': 'manage_education_documents', 'url': 'admin.manage_education_documents',
             'icon': 'bi-file-earmark-text',     'label': 'Документи освіти'},
            {'perm': 'manage_passport_documents', 'url': 'admin.manage_passport_documents',
             'icon': 'bi-person-vcard',          'label': 'Паспортні дані'},
            {'perm': 'manage_diplomas',          'url': 'admin.manage_diplomas',
             'icon': 'bi-journal-bookmark',      'label': 'Номери диплому / додатку'},
            {'perm': 'manage_military_reports',  'url': 'military_reports.dashboard',
             'icon': 'bi-shield-fill-check',     'label': 'Військовий облік (Додаток 5)'},
            {'divider': True},
            {'perm': 'manage_orders',            'url': 'admin.orders_list',
             'icon': 'bi-journal-text',          'label': 'Реєстр наказів'},
            {'perm': 'create_enrollment_order',  'url': 'admin.enrollment_order',
             'icon': 'bi-file-earmark-plus',     'label': 'Наказ про зарахування'},
            {'perm': 'manage_expulsion',         'url': 'admin.expulsion',
             'icon': 'bi-person-x',              'label': 'Наказ про відрахування'},
            {'divider': True},
            {'perm': 'import_education_docs',     'url': 'admin.import_docs_preview',
             'icon': 'bi-file-earmark-arrow-up',  'label': 'Імпорт документів освіти (Excel)'},
            {'perm': 'import_passport_documents', 'url': 'admin.import_passport_documents',
             'icon': 'bi-person-vcard',           'label': 'Імпорт паспортів (Excel)'},
            {'perm': 'import_military_pdf',       'url': 'students.military_bulk_import_pdf',
             'icon': 'bi-file-earmark-pdf',       'label': 'Імпорт PDF «Резерв+» (пакетно)'},
        ],
    },

    # ═══════════════════════════════════════════════════════════
    #  4. СИСТЕМА (довідники + адміністрування + звіти)
    # ═══════════════════════════════════════════════════════════
    {
        'type': 'dropdown',
        'key': 'system',
        'label': 'Система',
        'icon': 'bi-gear-fill',
        'items': [
            {'perm': 'manage_specialties',           'url': 'admin.manage_specialties',
             'icon': 'bi-diagram-3',                 'label': 'Спеціальності'},
            {'perm': 'manage_degree_levels',         'url': 'admin.manage_degree_levels',
             'icon': 'bi-mortarboard',               'label': 'Ступені'},
            {'perm': 'manage_educational_programs',  'url': 'admin.manage_educational_programs',
             'icon': 'bi-journal-bookmark',          'label': 'Освітня програма'},
            {'perm': 'manage_qualification_names',   'url': 'admin.manage_qualification_names',
             'icon': 'bi-mortarboard-fill',          'label': 'Назва кваліфікації'},
            {'perm': 'manage_licenses',              'url': 'admin.manage_licenses',
             'icon': 'bi-award',                     'label': 'Ліцензії'},
            {'perm': 'manage_accreditations',        'url': 'admin.manage_accreditations',
             'icon': 'bi-patch-check',               'label': 'Акредитації'},
            {'divider': True},
            {'perm': 'manage_users',                 'url': 'admin.manage_users',
             'icon': 'bi-people',                    'label': 'Користувачі та права'},
            {'perm': 'manage_templates',             'url': 'admin.manage_templates',
             'icon': 'bi-file-earmark-arrow-up',     'label': 'Шаблони документів'},
            {'divider': True},
            {'perm': 'license_report',               'url': 'admin.license_report',
             'icon': 'bi-list-check',                'label': 'Звіт по ліцензіях'},
            {'perm': 'analytics',                    'url': 'analytics.dashboard',
             'icon': 'bi-graph-up',                  'label': 'Аналітика'},
            {'perm': 'custom_export',                'url': 'admin.custom_export',
             'icon': 'bi-file-earmark-excel',        'label': 'Кастомний експорт даних'},
            {'perm': 'view_logs',                    'url': 'admin.view_logs',
             'icon': 'bi-file-earmark-binary',       'label': 'Логування'},
        ],
    },
]