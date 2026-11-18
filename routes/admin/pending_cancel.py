"""Allow cancelling any new public registration application."""
import re

from flask import abort, flash, redirect, request, url_for

from routes.admin import admin_bp
from routes.db import get_db
from routes.helpers import current_username
from routes.utils import log_action, permission_required


@admin_bp.route('/admin/pending_students/<int:pending_id>/cancel', methods=['POST'])
@permission_required('manage_pending_students')
def cancel_pending_student(pending_id):
    """Reject/cancel a pending application whether or not a student exists yet."""
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
        flash('Відхилити можна лише нову заявку.', 'warning')
        return redirect(url_for('admin.pending_student_review', pending_id=pending_id))

    note = (request.form.get('review_note') or '').strip() or 'Заявку скасовано адміністратором'
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


@admin_bp.after_app_request
def add_cancel_button_to_pending_application(response):
    """Render a general cancellation action even when duplicate UI is absent."""
    if request.endpoint != 'admin.pending_student_review' or request.method != 'GET':
        return response
    if not response.mimetype or response.mimetype != 'text/html':
        return response
    pending_id = (request.view_args or {}).get('pending_id')
    if pending_id is None:
        return response

    conn = get_db()
    try:
        row = conn.execute(
            "SELECT status FROM pending_students WHERE id = ?", (pending_id,)
        ).fetchone()
    finally:
        conn.close()
    if not row or row['status'] != 'new':
        return response

    body = response.get_data(as_text=True)
    # The duplicate-specific rejection card is sufficient when a matching
    # student exists. For all other new applications, add a separate cancel
    # button just below the page's top navigation row.
    if 'value="reject"' in body:
        return response
    button = (
        '<form method="post" action="/admin/pending_students/'
        f'{pending_id}/cancel" class="mb-3" '
        'onsubmit="return confirm(\'Скасувати цю заявку?\');">'
        '<input type="hidden" name="review_note" value="Заявку скасовано адміністратором">'
        '<button type="submit" class="btn btn-sm btn-outline-danger">'
        '<i class="bi bi-x-circle me-1"></i>Скасувати заявку</button></form>'
    )
    marker = '</div>\n\n\n\n\n{% if row.status != \'new\' %}'
    # The Jinja marker is not present in the rendered response. Insert before
    # the application details by matching the stable page-header wrapper.
    marker = '<div class="row g-3 mb-4">'
    if marker in body:
        body = body.replace(marker, button + '\n' + marker, 1)
    response.set_data(body)
    return response
