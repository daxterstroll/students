"""Allow cancelling any new public registration application."""

from flask import abort, flash, redirect, request, url_for

from routes.admin import admin_bp
from routes.db import get_db
from routes.helpers import current_username
from routes.utils import log_action, permission_required


@admin_bp.route('/admin/pending_students/<int:pending_id>/cancel', methods=['POST'])
@permission_required('manage_pending_students')
def cancel_pending_student(pending_id):
    """Cancel a new application and require the administrator's reason."""
    conn = get_db()
    conn.row_factory = __import__('sqlite3').Row
    row = conn.execute(
        "SELECT last_name_UA, first_name_UA, status FROM pending_students WHERE id = ?",
        (pending_id,),
    ).fetchone()
    if not row:
        conn.close()
        abort(404)
    if row['status'] != 'new':
        conn.close()
        flash('Скасувати можна лише нову заявку.', 'warning')
        return redirect(url_for('admin.pending_student_review', pending_id=pending_id))

    note = (request.form.get('review_note') or '').strip()
    if not note:
        conn.close()
        flash('Вкажіть причину скасування заявки.', 'error')
        return redirect(url_for('admin.pending_student_review', pending_id=pending_id))

    conn.execute(
        "UPDATE pending_students SET status='rejected', reviewed_by=?, "
        "reviewed_at=datetime('now','localtime'), review_note=? "
        "WHERE id=? AND status='new'",
        (current_username(), note, pending_id),
    )
    conn.commit()
    conn.close()
    log_action(
        current_username(),
        f"скасував заявку на реєстрацію: {row['last_name_UA']} {row['first_name_UA']} (заявка ID {pending_id})",
        details=note,
    )
    flash('Заявку скасовано', 'success')
    return redirect(url_for('admin.pending_students', status='new'))
