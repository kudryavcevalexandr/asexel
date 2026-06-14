from __future__ import annotations

import os
from pathlib import Path
from uuid import uuid4

import pandas as pd
from flask import Flask, flash, redirect, render_template, request, send_from_directory, session, url_for
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "change-me-in-production")
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024
UPLOAD_DIR = Path(__file__).parent / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)
DEFAULT_BATCH_SIZE = 100
MAX_BATCH_SIZE = 1000


def current_path() -> Path | None:
    name = session.get("current_file")
    if not name:
        return None
    path = UPLOAD_DIR / Path(name).name
    return path if path.is_file() else None


@app.get("/")
def index():
    return render_template("upload.html")


@app.post("/upload")
def upload():
    file = request.files.get("excel_file")
    if not file or not file.filename or Path(file.filename).suffix.lower() != ".xlsx":
        flash("Выберите корректный файл формата .xlsx.")
        return redirect(url_for("index"))
    name = f"{uuid4().hex}_{secure_filename(file.filename) or 'workbook.xlsx'}"
    path = UPLOAD_DIR / name
    file.save(path)
    try:
        pd.ExcelFile(path, engine="openpyxl")
    except Exception as e:
        path.unlink(missing_ok=True)
        flash(f"Файл поврежден или не является читаемой книгой Excel. Детали: {e}")
        return redirect(url_for("index"))
    session.clear()
    session["current_file"] = name
    return redirect(url_for("editor"))


@app.get("/editor")
def editor():
    path = current_path()
    if not path:
        flash("Сначала загрузите файл.")
        return redirect(url_for("index"))
    try:
        excel = pd.ExcelFile(path, engine="openpyxl")
        sheets = excel.sheet_names
        sheet = request.args.get("sheet")
        if sheet not in sheets:
            sheet = sheets[0]
        columns = pd.read_excel(path, sheet_name=sheet, nrows=0, engine="openpyxl").columns.tolist()
    except Exception:
        flash("Не удалось прочитать файл. Возможно, он поврежден.")
        return redirect(url_for("index"))
    return render_template("editor.html", filename=path.name.split("_", 1)[1], sheets=sheets, sheet=sheet, columns=columns, batch_size=DEFAULT_BATCH_SIZE, max_batch_size=MAX_BATCH_SIZE)


@app.get("/table")
def table():
    path = current_path()
    if not path:
        flash("Сначала загрузите файл.")
        return redirect(url_for("index"))
    try:
        excel = pd.ExcelFile(path, engine="openpyxl")
        sheet = request.args.get("sheet")
        if sheet not in excel.sheet_names:
            sheet = excel.sheet_names[0]
        frame = pd.read_excel(path, sheet_name=sheet, dtype=str, keep_default_na=False, engine="openpyxl")
        requested_columns = request.args.getlist("columns")
        visible = [column for column in frame.columns if column in requested_columns] if requested_columns else frame.columns.tolist()
        frame = frame.loc[:, visible]
        if request.args.get("transition") == "true" and "is_transition" in frame.columns:
            frame = frame[frame["is_transition"].astype(str).str.strip().str.lower().eq("true")]
        batch_size = max(1, min(int(request.args.get("batch_size", DEFAULT_BATCH_SIZE)), MAX_BATCH_SIZE))
        page = max(1, int(request.args.get("page", 1)))
    except (ValueError, TypeError):
        flash("Размер батча и номер страницы должны быть целыми числами.")
        return redirect(url_for("editor"))
    except Exception:
        flash("Не удалось прочитать файл. Возможно, он поврежден.")
        return redirect(url_for("editor"))
    total = len(frame)
    pages = max(1, (total + batch_size - 1) // batch_size)
    page = min(page, pages)
    offset = (page - 1) * batch_size
    rows = frame.iloc[offset:offset + batch_size].values.tolist()
    def page_url(number):
        return url_for("table", sheet=sheet, columns=visible, batch_size=batch_size, transition=request.args.get("transition", ""), page=number)
    editor_url = url_for("editor", sheet=sheet)
    return render_template("table.html", filename=path.name.split("_", 1)[1], columns=visible, rows=rows, page=page, pages=pages, total=total, range_start=offset + 1 if total else 0, range_end=min(offset + batch_size, total), page_url=page_url, editor_url=editor_url)


@app.post("/save")
def save_changes():
    path = current_path()
    if not path:
        return redirect(url_for("index"))
    sheet = request.form.get("sheet", "")
    try:
        books = pd.read_excel(path, sheet_name=None, dtype=str, keep_default_na=False, engine="openpyxl")
        if sheet not in books:
            raise ValueError("Лист не найден")
        frame = books[sheet]
        for key, value in request.form.items():
            if not key.startswith("cell_"):
                continue
            _, row, col = key.split("_")
            frame.iat[int(row), int(col)] = value
        books[sheet] = frame
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            for name, frame in books.items():
                frame.to_excel(writer, sheet_name=name, index=False)
        flash("Изменения сохранены.")
    except Exception as exc:
        flash(f"Не удалось сохранить изменения: {exc}")
    return redirect(url_for("editor", sheet=sheet, page=request.form.get("page", 1), transition=request.form.get("transition", "")))


@app.get("/download")
def download():
    path = current_path()
    if not path:
        return redirect(url_for("index"))
    return send_from_directory(UPLOAD_DIR, path.name, as_attachment=True, download_name=path.name.split("_", 1)[1])


@app.post("/delete")
def delete_file():
    path = current_path()
    if path:
        path.unlink(missing_ok=True)
    session.clear()
    flash("Файл удален.")
    return redirect(url_for("index"))


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=os.getenv("FLASK_DEBUG") == "1")
