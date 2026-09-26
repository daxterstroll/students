# `routes/students/grades.py`

## Призначення
Виставлення оцінок студенту: звичайні предмети (`edit_grades`) і
активності - практики/курсові/атестації (`edit_activities_grades`).
Обидва враховують скорочену програму (студент бачить/може отримати
оцінку лише за ті предмети/активності, які видимі для його треку).

## Взаємодія з app.py
Непряма (через `students_bp`, див. `routes/students/__init__.py`).

## Маршрути
`/activities_grades/<id>`, `/grades/<id>`.

## Залежить від
`routes.utils` (`is_student_on_reduced_program`,
`program_track_condition`, `apply_track_overrides` - та сама спільна
логіка скороченої програми, що й у `routes/admin/subjects.py` та
`routes/gen_docx.py`).
