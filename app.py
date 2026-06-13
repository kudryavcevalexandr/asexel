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

BASE_CSS = """
:root{font-family:Inter,system-ui,sans-serif;color:#172033;background:#f4f7fb}*{box-sizing:border-box}body{margin:0}.wrap{max-width:1200px;margin:42px auto;padding:0 20px}.card{background:#fff;border:1px solid #e4eaf2;border-radius:18px;padding:24px;box-shadow:0 12px 35px #1b315014}.hero{max-width:620px;margin:12vh auto}.muted{color:#667085}.notice{padding:12px 15px;border-radius:10px;background:#fff4e5;color:#9a5700;margin:14px 0}.btn{display:inline-flex;align-items:center;gap:7px;border:0;border-radius:10px;padding:10px 16px;background:#2563eb;color:white;text-decoration:none;font-weight:650;cursor:pointer;font-size:14px}.btn.secondary{background:#eef2ff;color:#3346a8}.btn.danger{background:#fee2e2;color:#b42318}.actions{display:flex;gap:10px;flex-wrap:wrap;margin:18px 0}.drop{display:block;border:2px dashed #cbd5e1;border-radius:14px;padding:30px;text-align:center;margin:24px 0;background:#f8fafc}.meta{display:flex;gap:28px;flex-wrap:wrap;padding:15px 0}.meta b{display:block;font-size:18px;margin-top:4px}.tabs{display:flex;gap:7px;overflow:auto;margin:18px 0}.tab{padding:8px 13px;border-radius:9px;background:#eef2f6;color:#344054;text-decoration:none;white-space:nowrap}.tab.active{background:#2563eb;color:#fff}.table-box{overflow:auto;max-height:62vh;border:1px solid #dfe5ed;border-radius:12px}table{border-collapse:separate;border-spacing:0;width:100%;font-size:14px}th{position:sticky;top:0;background:#f1f5f9;z-index:1}th,td{border-right:1px solid #e5e9ef;border-bottom:1px solid #e5e9ef;padding:0;min-width:140px;text-align:left}th{padding:10px}td input{width:100%;border:0;padding:9px;background:white;font:inherit;outline-color:#2563eb}.rownum{min-width:48px!important;width:48px;text-align:center;color:#667085;background:#f8fafc;padding:9px}h1{margin-top:0}input[type=file]{max-width:100%}
"""

UPLOAD_HTML = """<!doctype html><html lang=ru><head><meta charset=utf-8><meta name=viewport content='width=device-width'><title>Excel Studio</title><style>{{ css }}</style></head><body><main class='wrap'><section class='card hero'><p class=muted>EXCEL STUDIO</p><h1>Редактируйте Excel прямо в браузере</h1><p class=muted>Загрузите книгу .xlsx, измените данные и скачайте готовый файл.</p>{% for m in get_flashed_messages() %}<div class=notice>{{m}}</div>{% endfor %}<form method=post action='{{ url_for("upload") }}' enctype=multipart/form-data><label class=drop><b>Выберите файл .xlsx</b><br><br><input type=file name=excel_file accept=.xlsx required></label><button class=btn>Загрузить и открыть →</button></form></section></main></body></html>"""

EDITOR_HTML = """<!doctype html><html lang=ru><head><meta charset=utf-8><meta name=viewport content='width=device-width'><title>Редактор Excel</title><style>{{ css }}</style></head><body><main class=wrap><section class=card><p class=muted>EXCEL STUDIO / РЕДАКТОР</p><h1>{{ filename }}</h1>{% for m in get_flashed_messages() %}<div class=notice>{{m}}</div>{% endfor %}<div class=meta><span class=muted>Размер<b>{{ size }}</b></span><span class=muted>Листов<b>{{ sheets|length }}</b></span><span class=muted>Активный лист<b>{{ active }}</b></span></div><div class=tabs>{% for sheet in sheets %}<a class='tab {% if sheet==active %}active{% endif %}' href='{{ url_for("editor", sheet=sheet) }}'>{{sheet}}</a>{% endfor %}</div><form method=post action='{{ url_for("save_changes") }}'><input type=hidden name=sheet value='{{active}}'><div class=table-box><table><thead><tr><th class=rownum>#</th>{% for col in columns %}<th>{{col}}</th>{% endfor %}</tr></thead><tbody>{% for row in rows %}<tr><td class=rownum>{{loop.index}}</td>{% for value in row %}<td><input name='cell_{{loop.index0}}_{{loop.index}}' value='{{value}}'></td>{% endfor %}</tr>{% endfor %}</tbody></table></div><input type=hidden name=rows value='{{rows|length}}'><input type=hidden name=cols value='{{columns|length}}'><div class=actions><button class=btn>Сохранить изменения</button><a class='btn secondary' href='{{url_for("download")}}'>Скачать готовый файл</a></form><form method=post action='{{url_for("delete_file")}}' onsubmit='return confirm("Удалить файл?")'><button class='btn danger'>Удалить файл</button></form></div></section></main></body></html>"""


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
        excel = pd.ExcelFile(path, engine="openpyxl")
        active = request.args.get("sheet")
        if active not in excel.sheet_names:
            active = excel.sheet_names[0]
        df = pd.read_excel(path, sheet_name=active, dtype=str, keep_default_na=False, engine="openpyxl")
    except Exception:
        flash("Не удалось прочитать файл. Возможно, он поврежден.")
        return redirect(url_for("index"))
    original = path.name.split("_", 1)[1]
    size = f"{path.stat().st_size / 1024:.1f} КБ"
    return render_template_string(EDITOR_HTML, css=BASE_CSS, filename=original, size=size, sheets=excel.sheet_names, active=active, columns=list(df.columns), rows=df.values.tolist())


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
        rows, cols = int(request.form["rows"]), int(request.form["cols"])
        data = [[request.form.get(f"cell_{r}_{c}", "") for c in range(cols)] for r in range(rows)]
        books[sheet] = pd.DataFrame(data, columns=books[sheet].columns)
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            for name, frame in books.items():
                frame.to_excel(writer, sheet_name=name, index=False)
        flash("Изменения сохранены.")
    except Exception as exc:
        flash(f"Не удалось сохранить изменения: {exc}")
    return redirect(url_for("editor", sheet=sheet))


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
