from __future__ import annotations

import os
from pathlib import Path
from uuid import uuid4

import pandas as pd
from flask import Flask, flash, redirect, render_template_string, request, send_from_directory, session, url_for
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "change-me-in-production")
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024
UPLOAD_DIR = Path(__file__).parent / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)
DEFAULT_BATCH_SIZE = 100
MAX_BATCH_SIZE = 1000

BASE_CSS = """
:root{font-family:Inter,system-ui,sans-serif;color:#172033;background:#f4f7fb}*{box-sizing:border-box}body{margin:0}.wrap{max-width:1400px;margin:42px auto;padding:0 20px}.card{background:#fff;border:1px solid #e4eaf2;border-radius:18px;padding:24px;box-shadow:0 12px 35px #1b315014}.hero{max-width:620px;margin:12vh auto}.muted{color:#667085}.notice{padding:12px 15px;border-radius:10px;background:#fff4e5;color:#9a5700;margin:14px 0}.btn{display:inline-flex;align-items:center;gap:7px;border:0;border-radius:10px;padding:10px 16px;background:#2563eb;color:white;text-decoration:none;font-weight:650;cursor:pointer;font-size:14px}.btn.secondary{background:#eef2ff;color:#3346a8}.btn.danger{background:#fee2e2;color:#b42318}.actions{display:flex;gap:10px;flex-wrap:wrap;margin:18px 0}.drop{display:block;border:2px dashed #cbd5e1;border-radius:14px;padding:30px;text-align:center;margin:24px 0;background:#f8fafc}.meta{display:flex;gap:28px;flex-wrap:wrap;padding:15px 0}.meta b{display:block;font-size:18px;margin-top:4px}.form-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:18px;margin:24px 0}.field{display:grid;gap:7px;font-weight:650}.field select,.field input{width:100%;padding:11px;border:1px solid #cbd5e1;border-radius:9px;background:#fff;font:inherit}.check{display:flex;align-items:center;gap:9px;margin:20px 0}.table-box{overflow:auto;border:1px solid #dfe5ed;border-radius:12px}table{border-collapse:separate;border-spacing:0;width:100%;table-layout:fixed;font-size:14px}th{position:sticky;top:0;background:#f1f5f9;z-index:1}th,td{border-right:1px solid #e5e9ef;border-bottom:1px solid #e5e9ef;padding:10px;text-align:left;vertical-align:top;white-space:pre-wrap;overflow-wrap:anywhere;word-break:break-word}th{width:33.333%}.pagination{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin:16px 0}.pagination .muted{margin-right:auto}h1{margin-top:0}input[type=file]{max-width:100%}
"""

UPLOAD_HTML = """<!doctype html><html lang=ru><head><meta charset=utf-8><meta name=viewport content='width=device-width'><title>Excel Studio</title><style>{{ css }}</style></head><body><main class='wrap'><section class='card hero'><p class=muted>EXCEL STUDIO</p><h1>Редактируйте Excel прямо в браузере</h1><p class=muted>Загрузите книгу .xlsx, настройте фильтры и просмотрите данные.</p>{% for m in get_flashed_messages() %}<div class=notice>{{m}}</div>{% endfor %}<form method=post action='{{ url_for("upload") }}' enctype=multipart/form-data><label class=drop><b>Выберите файл .xlsx</b><br><br><input type=file name=excel_file accept=.xlsx required></label><button class=btn>Загрузить и открыть →</button></form></section></main></body></html>"""

FILTER_HTML = """<!doctype html><html lang=ru><head><meta charset=utf-8><meta name=viewport content='width=device-width'><title>Настройка выборки</title><style>{{ css }}</style></head><body><main class=wrap><section class='card hero'><p class=muted>ШАГ 1 ИЗ 2</p><h1>Настройте выборку</h1><p class=muted>{{ filename }}</p>{% for m in get_flashed_messages() %}<div class=notice>{{m}}</div>{% endfor %}<form method=get action='{{url_for("table")}}'><div class=form-grid><label class=field>Лист<select name=sheet>{% for sheet in sheets %}<option>{{sheet}}</option>{% endfor %}</select></label><label class=field>Размер батча<input type=number name=batch_size value={{batch_size}} min=1 max={{max_batch_size}} required></label></div><label class=check><input type=checkbox name=transition value=true> Только записи, где is_transition = True</label><div class=actions><button class=btn>Показать таблицу →</button><a class='btn danger' href='{{url_for("index")}}'>Выбрать другой файл</a></div></form></section></main></body></html>"""

TABLE_HTML = """<!doctype html><html lang=ru><head><meta charset=utf-8><meta name=viewport content='width=device-width'><title>Таблица</title><style>{{ css }}</style></head><body><main class=wrap><section class=card><p class=muted>ШАГ 2 ИЗ 2</p><h1>{{ filename }}</h1>{% for m in get_flashed_messages() %}<div class=notice>{{m}}</div>{% endfor %}<div class=actions><a class='btn secondary' href='{{url_for("editor")}}'>← Изменить фильтры</a><a class='btn secondary' href='{{url_for("download")}}'>Скачать файл</a></div><div class=table-box><table><thead><tr>{% for col in columns %}<th>{{col}}</th>{% endfor %}</tr></thead><tbody>{% for row in rows %}<tr>{% for cell in row %}<td>{{cell}}</td>{% endfor %}</tr>{% endfor %}</tbody></table></div><div class=pagination><span class=muted>Записи {{range_start}}–{{range_end}} из {{total}}</span>{% if page > 1 %}<a class='btn secondary' href='{{page_url(page-1)}}'>← Назад</a>{% endif %}<b>Страница {{page}} из {{pages}}</b>{% if page < pages %}<a class='btn secondary' href='{{page_url(page+1)}}'>Вперёд →</a>{% endif %}</div></section></main></body></html>"""


def current_path() -> Path | None:
    name = session.get("current_file")
    if not name:
        return None
    path = UPLOAD_DIR / Path(name).name
    return path if path.is_file() else None


@app.get("/")
def index():
    return render_template_string(UPLOAD_HTML, css=BASE_CSS)


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
        sheets = pd.ExcelFile(path, engine="openpyxl").sheet_names
    except Exception:
        flash("Не удалось прочитать файл. Возможно, он поврежден.")
        return redirect(url_for("index"))
    return render_template_string(FILTER_HTML, css=BASE_CSS, filename=path.name.split("_", 1)[1], sheets=sheets, batch_size=DEFAULT_BATCH_SIZE, max_batch_size=MAX_BATCH_SIZE)


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
        visible = [column for column in frame.columns if frame[column].astype(str).str.strip().ne("").any()][:3]
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
        return url_for("table", sheet=sheet, batch_size=batch_size, transition=request.args.get("transition", ""), page=number)
    return render_template_string(TABLE_HTML, css=BASE_CSS, filename=path.name.split("_", 1)[1], columns=visible, rows=rows, page=page, pages=pages, total=total, range_start=offset + 1 if total else 0, range_end=min(offset + batch_size, total), page_url=page_url)


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
