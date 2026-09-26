# `routes/gen_docx.py`

## Призначення
Генерація `.docx`-документів студента (додаток до диплома, довідки
тощо) з `.docx`-шаблонів (`template_word/`) через `python-docx` і
`docxtpl`-подібну підстановку значень. Збирає дані з кількох таблиць
(оцінки, практики/курсові/атестації, документи про освіту, військовий
облік) в один словник контексту для шаблону, з урахуванням скороченої
програми (`apply_reduced_program_texts`, `program_track_condition`).

## Взаємодія з app.py
`app.py` імпортує лише один фільтр Jinja звідси:
`from routes.gen_docx import format_grade` →
`app.jinja_env.filters['format_grade'] = format_grade` (щоб оцінки
однаково форматувались і в шаблонах Word, і в HTML-сторінках). Основна
функція `gen_doc(...)` викликається не з `app.py`, а з
`routes/students/core.py` (`generate()` - один студент) і
`routes/admin/templates_and_export.py` (`generate_group_docs()` -
масово, по групі).

## Ключові функції
- `gen_doc(student, military, template, out, ...)` - головна точка
  входу: генерує один `.docx`-файл за шаблоном для одного студента.
- `get_subjects_grades`, `get_practice_data`, `get_coursework_data`,
  `get_attestation_data` - витягують і форматують дані для таблиць
  документа, кожна враховує `full_program_only`/`reduced_only`
  (видимість для скороченої програми).
- `format_grade(value)` - форматування оцінки (і в Word, і в HTML,
  через Jinja-фільтр).

## Залежить від
`python-docx`, `routes.utils` (логіка скороченої програми). Викликається
з `routes/students/core.py`, `routes/admin/templates_and_export.py`,
`routes/office_editor.py` (побічно, як крок перед відкриттям у
редакторі).
