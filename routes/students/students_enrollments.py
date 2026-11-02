"""
routes/students/enrollments.py

У __init__.py підключати ТІЛЬКИ після створення students_bp, наприклад:

    students_bp = Blueprint(...)
    from . import core
    from . import enrollments  # noqa: F401

Імпорт blueprint — відносний (from . import students_bp), інакше circular import.
"""
from flask import request, redirect, url_for, flash

from . import students_bp
from routes.utils import permission_required, log_action
from routes.helpers import current_username
from routes.enrollments_service import (
    add_enrollment,
    set_primary_enrollment,
    end_enrollment,
)


@students_bp.route('/students/<int:student_id>/enrollments/add', methods=['POST'])
@permission_required('manage_students')
def enrollment_add(student_id):
    group_id = request.form.get('group_id', type=int)
    is_primary = request.form.get('is_primary') == '1'
    if not group_id:
        flash('Оберіть групу.', 'error')
        return redirect(url_for('students.student_details', student_id=student_id))

    ok, msg = add_enrollment(student_id, group_id, is_primary=is_primary)
    flash(msg, 'success' if ok else 'error')
    if ok:
        log_action(
            current_username(),
            f"додав навчання student={student_id} → group={group_id}",
            details=f"primary={is_primary}",
        )
    return redirect(url_for('students.student_details', student_id=student_id))


@students_bp.route('/students/<int:student_id>/enrollments/set_primary', methods=['POST'])
@permission_required('manage_students')
def enrollment_set_primary(student_id):
    group_id = request.form.get('group_id', type=int)
    if not group_id:
        flash('Не вказано групу.', 'error')
        return redirect(url_for('students.student_details', student_id=student_id))

    ok, msg = set_primary_enrollment(student_id, group_id)
    flash(msg, 'success' if ok else 'error')
    if ok:
        log_action(
            current_username(),
            f"змінив основне навчання student={student_id} → group={group_id}",
        )
    return redirect(url_for('students.student_details', student_id=student_id))


@students_bp.route('/students/<int:student_id>/enrollments/end', methods=['POST'])
@permission_required('manage_students')
def enrollment_end(student_id):
    group_id = request.form.get('group_id', type=int)
    status = request.form.get('status') or 'completed'
    if status not in ('completed', 'expelled'):
        status = 'completed'
    if not group_id:
        flash('Не вказано групу.', 'error')
        return redirect(url_for('students.student_details', student_id=student_id))

    ok, msg = end_enrollment(student_id, group_id, status=status)
    flash(msg, 'success' if ok else 'error')
    if ok:
        log_action(
            current_username(),
            f"завершив навчання student={student_id} group={group_id} status={status}",
        )
    return redirect(url_for('students.student_details', student_id=student_id))
