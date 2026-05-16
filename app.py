from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
import os

import pandas as pd
from flask import Flask, redirect, render_template_string, request, session, url_for

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "dev-secret-key")

PREVIEW_ROWS = int(os.getenv("PREVIEW_ROWS", "20"))


@dataclass
class SheetInfo:
    name: str
    rows: int
    cols: int
    preview_html: str


def build_preview(file_name: str, payload: bytes) -> tuple[dict[str, str | int], list[SheetInfo]]:
    excel = pd.ExcelFile(BytesIO(payload))

    metadata: dict[str, str | int] = {
        "file_name": file_name,
        "size_kb": round(len(payload) / 1024, 1),
        "sheets_count": len(excel.sheet_names),
    }

    sheets: list[SheetInfo] = []
    for sheet_name in excel.sheet_names:
        df = pd.read_excel(excel, sheet_name=sheet_name)
        sheets.append(
            SheetInfo(
                name=sheet_name,
                rows=len(df),
                cols=len(df.columns),
                preview_html=df.head(PREVIEW_ROWS).to_html(classes="preview", index=False, border=0),
            )
        )

    return metadata, sheets


UPLOAD_TEMPLATE = """
<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <title>Выбор Excel файла</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 2rem; background: #f6f7fb; color: #222; }
    .card { max-width: 760px; margin: 0 auto; background: white; border-radius: 14px; padding: 1.4rem 1.6rem; box-shadow: 0 2px 10px rgba(0,0,0,.06); }
    h1 { margin-top: 0; }
    .hint { color: #666; margin-top: .3rem; }
    .field { margin-top: 1rem; }
    .btn { margin-top: 1rem; background: #1f63ff; color: white; border: none; border-radius: 10px; padding: .65rem 1rem; cursor: pointer; font-size: 1rem; }
    .btn:hover { background: #1755df; }
    .error { background: #ffe9e9; border: 1px solid #ffc9c9; color: #9f2929; border-radius: 10px; padding: .7rem .9rem; margin-bottom: .8rem; }
  </style>
</head>
<body>
  <div class="card">
    <h1>Страница 1: выбор Excel файла</h1>
    <p class="hint">Нажмите кнопку выбора и укажите файл в папке на вашем устройстве.</p>

    {% if error_message %}
      <div class="error">{{ error_message }}</div>
    {% endif %}

    <form action="/upload" method="post" enctype="multipart/form-data">
      <div class="field">
        <label for="excel_file"><b>Excel файл (.xls, .xlsx)</b></label><br>
        <input id="excel_file" type="file" name="excel_file" accept=".xls,.xlsx" required>
      </div>
      <button class="btn" type="submit">Открыть предпросмотр</button>
    </form>
  </div>
</body>
</html>
"""


PREVIEW_TEMPLATE = """
<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <title>Предпросмотр Excel файла</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 2rem; background: #f6f7fb; color: #222; }
    .card { background: white; border-radius: 12px; padding: 1rem 1.4rem; margin-bottom: 1rem; box-shadow: 0 2px 8px rgba(0,0,0,.06); }
    table { border-collapse: collapse; width: 100%; font-size: .88rem; margin-top: .5rem; }
    th, td { border: 1px solid #ddd; padding: .4rem .5rem; text-align: left; }
    th { background: #f1f1f1; }
    .btn { display:inline-block; background:#eef4ff; color:#1f4a93; border-radius:999px; padding:.45rem .8rem; text-decoration:none; }
    .btn:hover { background:#dce9ff; }
  </style>
</head>
<body>
  <div class="card">
    <a class="btn" href="/">← Выбрать другой файл</a>
    <h1>Страница 2: предпросмотр файла</h1>
    <p><b>Имя файла:</b> {{ metadata.file_name }}</p>
    <p><b>Размер:</b> {{ metadata.size_kb }} KB | <b>Листов:</b> {{ metadata.sheets_count }}</p>
  </div>

  {% for sheet in sheets %}
    <div class="card">
      <h2>Лист: {{ sheet.name }}</h2>
      <p><b>Строк:</b> {{ sheet.rows }} | <b>Колонок:</b> {{ sheet.cols }}</p>
      <h3>Первые {{ preview_rows }} строк</h3>
      {{ sheet.preview_html | safe }}
    </div>
  {% endfor %}
</body>
</html>
"""


@app.route("/")
def select_page() -> str:
    error_message = session.pop("error_message", None)
    return render_template_string(UPLOAD_TEMPLATE, error_message=error_message)


@app.route("/upload", methods=["POST"])
def upload_excel():
    uploaded_file = request.files.get("excel_file")
    if uploaded_file is None or uploaded_file.filename == "":
        session["error_message"] = "Файл не выбран. Выберите Excel файл и повторите попытку."
        return redirect(url_for("select_page"))

    filename = uploaded_file.filename
    suffix = Path(filename).suffix.lower()
    if suffix not in {".xls", ".xlsx"}:
        session["error_message"] = "Поддерживаются только файлы формата .xls и .xlsx."
        return redirect(url_for("select_page"))

    payload = uploaded_file.read()
    session["file_name"] = filename
    session["excel_bytes"] = payload
    return redirect(url_for("preview_page"))


@app.route("/preview")
def preview_page():
    file_name = session.get("file_name")
    payload = session.get("excel_bytes")

    if not file_name or not payload:
        session["error_message"] = "Сначала выберите файл на первой странице."
        return redirect(url_for("select_page"))

    try:
        metadata, sheets = build_preview(file_name, payload)
    except Exception as exc:
        session["error_message"] = f"Ошибка чтения Excel: {exc}"
        return redirect(url_for("select_page"))

    return render_template_string(PREVIEW_TEMPLATE, metadata=metadata, sheets=sheets, preview_rows=PREVIEW_ROWS)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
