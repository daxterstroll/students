# Сесії та відомості — інтеграція

Модуль додано **поряд** з існуючими предметами/оцінками. Нічого в старому контурі не ламає.

## Файли

### Нові
```
routes/exam_sessions/
  __init__.py      # Blueprint exam_sessions_bp
  schema.py        # ensure_exam_session_tables
  service.py       # шкала ECTS, студенти по ліцензії, токени
  admin.py         # UI відповідального
  public.py        # /session-grade/<token>
  transfer.py      # перенос у grades / activity_grades
  excel_export.py  # Excel-відомість

static/templates/exam_sessions/
  sessions_list.html, session_form.html, session_detail.html
  sheet_form.html, sheet_detail.html
  teacher_form.html, teacher_invalid.html
  transfer_queue.html, transfer_sheet.html
```

### Змінені
- `routes/user_scope.py` — `user_knowledge_areas` + урахування галузі в `get_allowed_group_ids`
- `routes/admin/users.py` — права `manage_exam_sessions`, `transfer_exam_grades`; скоуп галузі знань
- `static/templates/add_user.html`, `edit_user.html` — мультиселект «Галузі знань»
- `routes/nav.py` — пункти меню
- `app.py` — `register_blueprint(exam_sessions_bp)`

## Права

| Код | Призначення |
|-----|-------------|
| `manage_exam_sessions` | Сесії, відомості, токен викладачу, затвердження, Excel |
| `transfer_exam_grades` | Перенос затверджених оцінок у план |

Адмін має обидва автоматично. Скоуп: групи + спеціальності + ступені + **галузі знань**.

## Потік

1. Відповідальний створює **сесію** для групи (рік, семестр 1–3).
2. Додає **відомості**: назва дисципліни (вільна), тип для майбутнього переносу, контроль, викладач, **ліцензія**.
3. Система фіксує snapshot студентів групи з цією ліцензією.
4. **Посилання викладачу** → `/session-grade/<token>` (без логіну), бали → авто нац./ECTS.
5. Відповідальний **затверджує**.
6. Користувач з `transfer_exam_grades` (часто адмін) прив’язує відомість до subject/practice/coursework/attestation і пише в `grades` / `activity_grades`.

## Таблиці (створюються при першому заході)

- `exam_sessions`
- `session_sheets` (token, license_id, linked_*)
- `session_sheet_students` (snapshot + points/national/ects)
- `user_knowledge_areas`

Можна також викликати `ensure_exam_session_tables(conn)` з `init_db.py`.

## URL

| Маршрут | Хто |
|---------|-----|
| `/admin/exam-sessions` | manage_exam_sessions |
| `/admin/exam-sessions/<id>` | … |
| `/admin/exam-sessions/sheets/<id>` | … |
| `/admin/exam-sessions/sheets/<id>/excel` | … |
| `/admin/exam-sessions/transfer` | transfer_exam_grades |
| `/session-grade/<token>` | публічно |

База посилання викладачу: `PUBLIC_APPLY_BASE_URL` (як для update-info).
