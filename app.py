from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any
import os

import pandas as pd
from flask import Flask, render_template_string, request

app = Flask(__name__)

DEFAULT_DIR = Path(os.getenv("DOWNLOAD_DIR", "/sdcard/Download"))
EXCEL_PATH = Path(os.getenv("EXCEL_PATH", DEFAULT_DIR / "график ТЭЦ26 260424.xls"))
SECOND_EXCEL_PATH = Path(os.getenv("SECOND_EXCEL_PATH", DEFAULT_DIR / "nomenclature_parsed.xlsx"))
PREVIEW_ROWS = int(os.getenv("PREVIEW_ROWS", "10"))


@dataclass
class SheetInfo:
    name: str
    rows: int
    cols: int
    preview_html: str


@dataclass
class GroupedSheetInfo:
    name: str
    groups_tree_html: str
    groups_count: int
    total_rows: int


def split_wbs_id_columns(df: pd.DataFrame) -> pd.DataFrame:
    if "wbs_id" not in df.columns:
        return df

    parts = (
        df["wbs_id"]
        .astype("string")
        .fillna("")
        .str.split(".", n=3, expand=True, regex=False)
        .rename(columns=lambda idx: f"wbs_id_part_{idx + 1}")
    )

    return pd.concat([df, parts], axis=1)


def read_excel_file(path: Path) -> tuple[dict[str, Any], list[SheetInfo]]:
    stat = path.stat()
    metadata = {
        "file_name": path.name,
        "full_path": str(path),
        "size_mb": round(stat.st_size / (1024 * 1024), 2),
        "modified": pd.to_datetime(stat.st_mtime, unit="s").strftime("%Y-%m-%d %H:%M:%S"),
    }

    sheets: list[SheetInfo] = []
    excel = pd.ExcelFile(path, engine="xlrd")
    for sheet_name in excel.sheet_names:
        df = pd.read_excel(excel, sheet_name=sheet_name)
        df = split_wbs_id_columns(df)
        preview = df.head(PREVIEW_ROWS)
        sheets.append(
            SheetInfo(
                name=sheet_name,
                rows=len(df),
                cols=len(df.columns),
                preview_html=preview.to_html(classes="preview", index=False, border=0),
            )
        )

    return metadata, sheets


def read_grouped_by_wbs_part_3(path: Path) -> tuple[dict[str, Any], list[GroupedSheetInfo]]:
    metadata, _ = read_excel_file(path)
    grouped_sheets: list[GroupedSheetInfo] = []

    excel = pd.ExcelFile(path, engine="xlrd")
    for sheet_name in excel.sheet_names:
        df = pd.read_excel(excel, sheet_name=sheet_name)
        df = split_wbs_id_columns(df)

        if "wbs_id_part_3" not in df.columns:
            groups_tree_html = "<p>На этом листе нет колонки wbs_id_part_3.</p>"
            groups_count = 0
        else:
            details_parts: list[str] = ['<div class="tree-root">']
            grouped_dfs = sorted(
                df.groupby("wbs_id_part_3", dropna=False),
                key=lambda g: (str(g[0]).lower() if pd.notna(g[0]) else "zzz"),
            )
            groups_count = len(grouped_dfs)

            for group_value, group_df in grouped_dfs:
                group_name = str(group_value) if pd.notna(group_value) else "Пустое значение"
                details_parts.append(
                    "<details class='tree-node'>"
                    f"<summary><b>{group_name}</b> — строк: {len(group_df)}</summary>"
                    f"{group_df.to_html(classes='preview', index=False, border=0)}"
                    "</details>"
                )

            details_parts.append("</div>")
            groups_tree_html = "".join(details_parts)

        grouped_sheets.append(
            GroupedSheetInfo(
                name=sheet_name,
                groups_tree_html=groups_tree_html,
                groups_count=groups_count,
                total_rows=len(df),
            )
        )

    return metadata, grouped_sheets


def get_available_files() -> list[Path]:
    files = [EXCEL_PATH]
    if SECOND_EXCEL_PATH != EXCEL_PATH:
        files.append(SECOND_EXCEL_PATH)
    return files


def resolve_selected_file() -> tuple[Path | None, str | None]:
    selected_path = request.args.get("file")
    available_files = get_available_files()

    if selected_path:
        candidate = Path(selected_path)
        if any(candidate == file_path for file_path in available_files):
            return candidate, None
        return None, f"Недопустимый путь файла: {selected_path}"

    if EXCEL_PATH.exists():
        return EXCEL_PATH, None

    for file_path in available_files:
        if file_path.exists():
            return file_path, None

    return None, (
        f"Файл не найден: {EXCEL_PATH}<br>"
        "Укажите путь к файлу через переменную окружения EXCEL_PATH."
    )


NAV_TEMPLATE = """
<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <title>Навигация</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 2rem; background: #f7f7f7; color: #222; }
    .card { background: white; border-radius: 10px; padding: 1.2rem 1.5rem; margin-bottom: 1rem; box-shadow: 0 2px 8px rgba(0,0,0,.06); }
    .nav { display: flex; flex-wrap: wrap; gap: .7rem; margin-top: .8rem; }
    .btn { display:inline-block; background:#eef4ff; color:#1f4a93; border-radius:999px; padding:.5rem .9rem; text-decoration:none; }
    .btn:hover { background:#dce9ff; }
  </style>
</head>
<body>
  <div class="card">
    <h1>Навигация по страницам</h1>
    <p>Выберите файл на предпросмотр:</p>
    <div class="nav">
      <a class="btn" href="/preview?file=/sdcard/Download/график%20ТЭЦ26%20260424.xls">/sdcard/Download/график ТЭЦ26 260424.xls</a>
      <a class="btn" href="/preview?file=/sdcard/Download/nomenclature_parsed.xlsx">/sdcard/Download/nomenclature_parsed.xlsx</a>
    </div>
    <p style="margin-top:1rem;">Или откройте группировку для основного файла:</p>
    <div class="nav">
      <a class="btn" href="/grouped">Группировка по wbs_id_part_3</a>
    </div>
  </div>
</body>
</html>
"""

TEMPLATE = """
<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <title>Просмотр Excel</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 2rem; background: #f7f7f7; color: #222; }
    .card { background: white; border-radius: 10px; padding: 1rem 1.5rem; margin-bottom: 1rem; box-shadow: 0 2px 8px rgba(0,0,0,.06); }
    table { border-collapse: collapse; width: 100%; font-size: .88rem; margin-top: .45rem; }
    th, td { border: 1px solid #ddd; padding: .4rem .5rem; text-align: left; }
    th { background: #f0f0f0; }
    h1, h2 { margin-top: 0; }
    .meta p { margin: .2rem 0; }
  </style>
</head>
<body>
  <div class="card">
    <p><a href="/">← На страницу навигации</a> | <a href="/grouped">Группировка по wbs_id_part_3</a></p>
    <h1>Основные данные Excel-файлов</h1>
  </div>

  {% for file_preview in previews %}
    <div class="card">
      <h2>Файл: {{ file_preview.meta.file_name }}</h2>
      <div class="meta">
        <p><b>Путь:</b> {{ file_preview.meta.full_path }}</p>
        <p><b>Размер:</b> {{ file_preview.meta.size_mb }} МБ</p>
        <p><b>Изменён:</b> {{ file_preview.meta.modified }}</p>
        <p><b>Листов:</b> {{ file_preview.sheets|length }}</p>
      </div>
    </div>

    {% for s in file_preview.sheets %}
      <div class="card" id="file-{{ loop.parent_loop.index }}-sheet-{{ loop.index }}">
        <h3>Лист: {{ s.name }}</h3>
        <p><b>Строк:</b> {{ s.rows }} | <b>Колонок:</b> {{ s.cols }}</p>
        <h4>Первые {{ preview_rows }} строк</h4>
        {{ s.preview_html | safe }}
      </div>
    {% endfor %}
  {% endfor %}
</body>
</html>
"""


GROUPED_TEMPLATE = """
<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <title>Группировка Excel по wbs_id_part_3</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 2rem; background: #f7f7f7; color: #222; }
    .card { background: white; border-radius: 10px; padding: 1rem 1.5rem; margin-bottom: 1rem; box-shadow: 0 2px 8px rgba(0,0,0,.06); }
    table { border-collapse: collapse; width: 100%; font-size: .9rem; }
    th, td { border: 1px solid #ddd; padding: .4rem .5rem; text-align: left; }
    th { background: #f0f0f0; }
    h1, h2 { margin-top: 0; }
    .meta p { margin: .2rem 0; }
    .top-nav { margin-bottom: 1rem; display:flex; gap: .6rem; flex-wrap: wrap; }
    .btn { display:inline-block; background:#fff; border:1px solid #d8d8d8; border-radius:999px; padding:.45rem .8rem; text-decoration:none; color:#222; font-size:.92rem; }
    .btn:hover { background:#f1f1f1; }
    .sheet-nav { margin-top: .8rem; display:flex; gap:.45rem; flex-wrap: wrap; }
    .sheet-chip { display:inline-block; background:#eef4ff; color:#1f4a93; border-radius:999px; padding:.3rem .7rem; text-decoration:none; font-size:.86rem; }
    .sheet-chip:hover { background:#dce9ff; }
    .hint { color:#666; font-size: .9rem; margin-top:.4rem; }
    .tree-node { border: 1px solid #e5e5e5; border-radius: 8px; margin: .5rem 0; padding: .45rem .65rem; background:#fcfcfc; }
    .tree-node > summary { cursor: pointer; }
  </style>
</head>
<body>
  <div class="top-nav">
    <a class="btn" href="/">← К навигации</a>
    <a class="btn" href="/preview">К предпросмотру листов</a>
    <a class="btn" href="#grouped-sections">↓ К таблицам группировок</a>
  </div>

  <div class="card">
    <h1>Группировка строк по wbs_id_part_3</h1>
    <p class="hint">Это интерактивное дерево: раскройте нужную группу, чтобы посмотреть строки внутри неё.</p>
    <div class="meta">
      <p><b>Файл:</b> {{ meta.file_name }}</p>
      <p><b>Листов:</b> {{ grouped_sheets|length }}</p>
    </div>
    <div class="sheet-nav">
      {% for s in grouped_sheets %}
        <a class="sheet-chip" href="#sheet-{{ loop.index }}">{{ s.name }}</a>
      {% endfor %}
    </div>
  </div>

  <div id="grouped-sections"></div>
  {% for s in grouped_sheets %}
    <div class="card" id="sheet-{{ loop.index }}">
      <h2>Лист: {{ s.name }}</h2>
      <p><b>Групп:</b> {{ s.groups_count }} | <b>Всего строк:</b> {{ s.total_rows }}</p>
      {{ s.groups_tree_html | safe }}
    </div>
  {% endfor %}
</body>
</html>
"""


@app.route("/")
def navigation_page():
    return render_template_string(NAV_TEMPLATE)


@app.route("/preview")
def index():
    selected_file, error_message = resolve_selected_file()
    if error_message:
        return error_message

    try:
        metadata, sheets = read_excel_file(selected_file)
    except Exception as exc:
        return f"Ошибка чтения Excel ({selected_file.name}): {exc}"

    previews = [{"meta": metadata, "sheets": sheets}]
    return render_template_string(TEMPLATE, previews=previews, preview_rows=PREVIEW_ROWS)


@app.route("/grouped")
def grouped_page():
    if not EXCEL_PATH.exists():
        return (
            f"Файл не найден: {EXCEL_PATH}<br>"
            "Укажите путь к файлу через переменную окружения EXCEL_PATH."
        )

    try:
        metadata, grouped_sheets = read_grouped_by_wbs_part_3(EXCEL_PATH)
    except Exception as exc:
        return f"Ошибка чтения Excel: {exc}"

    return render_template_string(GROUPED_TEMPLATE, meta=metadata, grouped_sheets=grouped_sheets)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
