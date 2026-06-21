from __future__ import annotations

import os
import tempfile
from pathlib import Path
from uuid import uuid4
from zipfile import BadZipFile

import pandas as pd
from flask import Flask, flash, jsonify, redirect, render_template, request, send_from_directory, session, url_for
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "change-me-in-production")
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024
UPLOAD_DIR = Path(__file__).parent / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)
DEFAULT_BATCH_SIZE = 100
MAX_BATCH_SIZE = 1000
DEFAULT_TABLE_SCALE = 80
MIN_TABLE_SCALE = 50
MAX_TABLE_SCALE = 150
ANCHOR_COLUMNS = ("anchor_name", "anchor", "anchors")
TRUE_VALUES = {"true", "1", "yes", "y", "да", "истина"}
FALSE_VALUES = {"false", "0", "no", "n", "нет", "ложь"}


def workbook_error_details(action: str, path: Path, exc: Exception) -> str:
    reason = f"{type(exc).__name__}: {exc}"
    file_state = "файл отсутствует"
    if path.exists():
        try:
            file_state = f"размер файла: {path.stat().st_size} байт"
        except OSError as stat_exc:
            file_state = f"не удалось получить размер файла: {type(stat_exc).__name__}: {stat_exc}"

    hints = []
    if isinstance(exc, BadZipFile) or "not a zip file" in str(exc).lower():
        hints.append("openpyxl получил невалидный ZIP-контейнер .xlsx; часто это происходит, если вместо настоящего .xlsx загружен .xls/CSV/HTML с расширением .xlsx или предыдущая запись файла оборвалась")
    if isinstance(exc, PermissionError):
        hints.append("у процесса нет прав на запись/чтение файла или файл заблокирован операционной системой")
    if not hints:
        hints.append("проверьте формат .xlsx, выбранный лист, права на папку uploads и свободное место на диске")

    return f"Не удалось {action}. Технические детали: {reason}. Состояние файла на сервере: {file_state}. Что проверить: {'; '.join(hints)}."

def boolean_checkbox_values(value: str) -> tuple[bool, str, str] | None:
    normalized = str(value).strip().lower()
    if normalized in TRUE_VALUES:
        false_value = "0" if normalized == "1" else "Ложь" if normalized == "истина" else "False"
        return True, value, false_value
    if normalized in FALSE_VALUES:
        true_value = "1" if normalized == "0" else "Истина" if normalized == "ложь" else "True"
        return False, true_value, value
    return None


def parse_anchor_filter(raw: str) -> list[str]:
    anchors: list[str] = []
    seen: set[str] = set()
    for part in raw.replace("\r", "\n").replace(",", "\n").replace(";", "\n").split("\n"):
        value = part.strip()
        if value and value not in seen:
            anchors.append(value)
            seen.add(value)
    return anchors


def anchor_filter_column(columns) -> str | None:
    lowered = {str(column).strip().lower(): column for column in columns}
    for name in ANCHOR_COLUMNS:
        if name in lowered:
            return lowered[name]
    for column in columns:
        if "anchor" in str(column).lower():
            return column
    return None


def current_path() -> Path | None:
    name = session.get("current_file")
    if not name:
        return None
    path = UPLOAD_DIR / Path(name).name
    return path if path.is_file() else None


@app.get("/")
def index():
    return render_template("upload.html")


def wants_json_response() -> bool:
    return request.headers.get("X-Requested-With") == "XMLHttpRequest" or request.accept_mimetypes.best == "application/json"


def upload_error(message: str, status_code: int = 400):
    if wants_json_response():
        return jsonify({"ok": False, "message": message}), status_code
    flash(message)
    return redirect(url_for("index"))


@app.post("/upload")
def upload():
    file = request.files.get("excel_file")
    if not file or not file.filename:
        return upload_error("Файл не был передан браузером. Выберите файл .xlsx еще раз и повторите загрузку.")
    original_name = file.filename
    if Path(original_name).suffix.lower() != ".xlsx":
        return upload_error(f"Выбран файл «{original_name}», но поддерживаются только книги Excel в формате .xlsx.")
    name = f"{uuid4().hex}_{secure_filename(original_name) or 'workbook.xlsx'}"
    path = UPLOAD_DIR / name
    try:
        file.save(path)
    except Exception as exc:
        path.unlink(missing_ok=True)
        return upload_error(f"Не удалось получить выбранный файл «{original_name}» от браузера. Детали: {exc}", 500)
    try:
        pd.ExcelFile(path, engine="openpyxl")
    except Exception as exc:
        path.unlink(missing_ok=True)
        return upload_error(f"Файл «{original_name}» получен, но openpyxl не смог открыть его как .xlsx. Проверьте, что это именно книга Excel .xlsx, а не .xls/.xlsm/CSV с другим расширением, и что файл не зашифрован паролем. Технические детали: {type(exc).__name__}: {exc}")
    session.clear()
    session["current_file"] = name
    editor_url = url_for("editor")
    if wants_json_response():
        return jsonify({"ok": True, "redirect": editor_url})
    return redirect(editor_url)


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
    except Exception as exc:
        flash(workbook_error_details("прочитать загруженный файл", path, exc))
        return redirect(url_for("index"))
    default_columns = [column for column in columns if column in {"anchor_name", "name_fixed"}]
    if not default_columns:
        default_columns = columns
    return render_template("editor.html", filename=path.name.split("_", 1)[1], sheets=sheets, sheet=sheet, columns=columns, default_columns=default_columns, batch_size=DEFAULT_BATCH_SIZE, max_batch_size=MAX_BATCH_SIZE, table_scale=DEFAULT_TABLE_SCALE, min_table_scale=MIN_TABLE_SCALE, max_table_scale=MAX_TABLE_SCALE, default_transition=True)


def render_table_page(*, without_anchor: bool = False, anchor_filter: bool = False):
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
        if request.args.get("transition") == "true" and "is_transition" in frame.columns:
            frame = frame[frame["is_transition"].astype(str).str.strip().str.lower().eq("true")]
        anchors_raw = request.args.get("anchors", "")
        anchors = parse_anchor_filter(anchors_raw) if anchor_filter else []
        anchor_column = anchor_filter_column(frame.columns) if anchor_filter else None
        if anchor_filter:
            if not anchors:
                flash("Введите anchors для фильтра.")
                return redirect(url_for("editor", sheet=sheet))
            if anchor_column is None:
                flash("В выбранном листе не найдена колонка с anchor.")
                return redirect(url_for("editor", sheet=sheet))
            frame = frame[frame[anchor_column].astype(str).str.strip().isin(anchors)]
        requested_columns = request.args.getlist("columns")
        visible = [column for column in frame.columns if column in requested_columns] if requested_columns else frame.columns.tolist()
        if without_anchor:
            visible = [column for column in visible if "anchor" not in column.lower()]
        column_indexes = {column: int(frame.columns.get_loc(column)) for column in visible}
        frame = frame.loc[:, visible]
        batch_size = max(1, min(int(request.args.get("batch_size", DEFAULT_BATCH_SIZE)), MAX_BATCH_SIZE))
        page = max(1, int(request.args.get("page", 1)))
        table_scale = max(MIN_TABLE_SCALE, min(int(request.args.get("table_scale", DEFAULT_TABLE_SCALE)), MAX_TABLE_SCALE))
    except (ValueError, TypeError):
        flash("Размер батча и номер страницы должны быть целыми числами.")
        return redirect(url_for("editor"))
    except Exception as exc:
        flash(workbook_error_details("прочитать файл для таблицы", path, exc))
        return redirect(url_for("editor"))
    total = len(frame)
    pages = max(1, (total + batch_size - 1) // batch_size)
    page = min(page, pages)
    offset = (page - 1) * batch_size
    rows = []
    for row_index, row in frame.iloc[offset:offset + batch_size].iterrows():
        cells = []
        for column, value in row.items():
            boolean_values = boolean_checkbox_values(value)
            cells.append({
                "value": value,
                "column": column,
                "column_index": column_indexes[column],
                "is_boolean": boolean_values is not None,
                "checked": boolean_values[0] if boolean_values else False,
                "true_value": boolean_values[1] if boolean_values else "",
                "false_value": boolean_values[2] if boolean_values else "",
            })
        rows.append({"index": int(row_index), "cells": cells})
    def page_url(number):
        endpoint = "table_anchors" if anchor_filter else "table_without_anchor" if without_anchor else "table"
        return url_for(endpoint, sheet=sheet, columns=visible, batch_size=batch_size, table_scale=table_scale, transition=request.args.get("transition", ""), anchors=request.args.get("anchors", "") if anchor_filter else None, page=number)
    editor_url = url_for("editor", sheet=sheet)
    without_anchor_url = None
    if not without_anchor:
        without_anchor_url = url_for("table_without_anchor", sheet=sheet, columns=visible, batch_size=batch_size, table_scale=table_scale, transition=request.args.get("transition", ""), page=page)

    pagination_pages = []
    for number in range(1, pages + 1):
        if number == 1 or number == pages or abs(number - page) <= 2:
            pagination_pages.append(number)
        elif pagination_pages and pagination_pages[-1] != "…":
            pagination_pages.append("…")

    table_endpoint = "table_anchors" if anchor_filter else "table_without_anchor" if without_anchor else "table"
    return render_template("table.html", filename=path.name.split("_", 1)[1], sheet=sheet, columns=visible, rows=rows, page=page, pages=pages, total=total, range_start=offset + 1 if total else 0, range_end=min(offset + batch_size, total), batch_size=batch_size, page_url=page_url, pagination_pages=pagination_pages, table_scale=table_scale, transition=request.args.get("transition") == "true", editor_url=editor_url, without_anchor=without_anchor, without_anchor_url=without_anchor_url, table_endpoint=table_endpoint, anchor_filter=anchor_filter, anchors=request.args.get("anchors", ""))


@app.get("/table")
def table():
    return render_table_page()


@app.get("/table/without-anchor")
def table_without_anchor():
    return render_table_page(without_anchor=True)


@app.get("/table/anchors")
def table_anchors():
    return render_table_page(anchor_filter=True)


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
        pending_boolean_updates = {}
        for key, value in request.form.items():
            if key.startswith("bool_false_"):
                _, _, row, col = key.split("_")
                pending_boolean_updates[(int(row), int(col))] = value
            elif key.startswith("bool_"):
                _, row, col = key.split("_")
                pending_boolean_updates[(int(row), int(col))] = value
            elif key.startswith("cell_"):
                _, row, col = key.split("_")
                frame.iat[int(row), int(col)] = value
        for (row, col), value in pending_boolean_updates.items():
            frame.iat[row, col] = value
        books[sheet] = frame
        temp_name = None
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", dir=UPLOAD_DIR) as temp_file:
                temp_name = temp_file.name
            temp_path = Path(temp_name)
            with pd.ExcelWriter(temp_path, engine="openpyxl") as writer:
                for name, frame in books.items():
                    frame.to_excel(writer, sheet_name=name, index=False)
            with pd.ExcelFile(temp_path, engine="openpyxl"):
                pass
            temp_path.replace(path)
        finally:
            if temp_name:
                Path(temp_name).unlink(missing_ok=True)
        flash("Изменения сохранены.")
    except Exception as exc:
        flash(workbook_error_details("сохранить изменения", path, exc))
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
