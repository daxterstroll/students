vidomist_sheet.docx

НЕ використовуйте {%tr for %} / {%tr endfor %} — Word їх ламає (помилка endfor).

У шаблоні:
  • {{ exam_date }}, {{ discipline_name }}, {% if license_name %}...{% endif %}
  • таблиця: 2 рядки шапки + 1 порожній рядок-зразок (шрифти колонок)
  • студентів підставляє КОД після render

Жирний підпис+значення в одному рядку:
  label звичайний run, {{ value }} — окремий run (можна жирний).
