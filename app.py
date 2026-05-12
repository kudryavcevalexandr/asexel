from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any
import os

import pandas as pd
from flask import Flask, render_template_string

app = Flask(__name__)

EXCEL_PATH = Path(os.getenv("EXCEL_PATH", "/sdcard/Download/график ТЭЦ26 260424.xls"))
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
    groups_html: str
    groups_count: int


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


def read_grouped_by_wbs_pairs(path: Path) -> tuple[dict[str, Any], list[GroupedSheetInfo]]:
    metadata, _ = read_excel_file(path)
    grouped_sheets: list[GroupedSheetInfo] = []

    excel = pd.ExcelFile(path, engine="xlrd")
    for sheet_name in excel.sheet_names:
        df = pd.read_excel(excel, sheet_name=sheet_name)
        df = split_wbs_id_columns(df)

        required_columns = {"wbs_id_part_1", "wbs_id_part_2"}
        if not required_columns.issubset(df.columns):
            groups_html = "<p>На этом листе нет колонок wbs_id_part_1 и wbs_id_part_2.</p>"
            groups_count = 0
        else:
            grouped = (
                df.groupby(["wbs_id_part_2", "wbs_id_part_1"], dropna=False)
                .size()
                .reset_index(name="rows_count")
                .sort_values(["wbs_id_part_2", "wbs_id_part_1"], na_position="last")
            )
            groups_count = len(grouped)
            groups_html = grouped.to_html(classes="preview", index=False, border=0)

        grouped_sheets.append(
            GroupedSheetInfo(
                name=sheet_name,
                groups_html=groups_html,
                groups_count=groups_count,
            )
        )

    return metadata, grouped_sheets


TEMPLATE = """
<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <title>Просмотр Excel</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 2rem; background: #f7f7f7; color: #222; }
    .card { background: white; border-radius: 10px; padding: 1rem 1.5rem; margin-bottom: 1rem; box-shadow: 0 2px 8px rgba(0,0,0,.06); }
    table { border-collapse: collapse; width: 100%; font-size: .9rem; }
    th, td { border: 1px solid #ddd; padding: .4rem .5rem; text-align: left; }
    th { background: #f0f0f0; }
    h1, h2 { margin-top: 0; }
    .meta p { margin: .2rem 0; }
  </style>
</head>
<body>
  <div class="card">
    <h1>Основные данные Excel-файла</h1>
    <div class="meta">
      <p><b>Файл:</b> {{ meta.file_name }}</p>
      <p><b>Путь:</b> {{ meta.full_path }}</p>
      <p><b>Размер:</b> {{ meta.size_mb }} МБ</p>
      <p><b>Изменён:</b> {{ meta.modified }}</p>
      <p><b>Листов:</b> {{ sheets|length }}</p>
    </div>
  </div>

  {% for s in sheets %}
    <div class="card">
      <h2>Лист: {{ s.name }}</h2>
      <p><b>Строк:</b> {{ s.rows }} | <b>Колонок:</b> {{ s.cols }}</p>
      <h3>Первые {{ preview_rows }} строк</h3>
      {{ s.preview_html | safe }}
    </div>
  {% endfor %}
</body>
</html>
"""


GROUPED_TEMPLATE = """
<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <title>Группировка Excel по wbs_id_part_2 / wbs_id_part_1</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 2rem; background: #f7f7f7; color: #222; }
    .card { background: white; border-radius: 10px; padding: 1rem 1.5rem; margin-bottom: 1rem; box-shadow: 0 2px 8px rgba(0,0,0,.06); }
    table { border-collapse: collapse; width: 100%; font-size: .9rem; }
    th, td { border: 1px solid #ddd; padding: .4rem .5rem; text-align: left; }
    th { background: #f0f0f0; }
    h1, h2 { margin-top: 0; }
    .meta p { margin: .2rem 0; }
    .links { margin-bottom: 1rem; }
  </style>
</head>
<body>
  <div class="links">
    <a href="/">← Назад к предпросмотру листов</a>
  </div>

  <div class="card">
    <h1>Группировка строк по паре wbs_id_part_2 - wbs_id_part_1</h1>
    <div class="meta">
      <p><b>Файл:</b> {{ meta.file_name }}</p>
      <p><b>Листов:</b> {{ grouped_sheets|length }}</p>
    </div>
  </div>

  {% for s in grouped_sheets %}
    <div class="card">
      <h2>Лист: {{ s.name }}</h2>
      <p><b>Групп:</b> {{ s.groups_count }}</p>
      {{ s.groups_html | safe }}
    </div>
  {% endfor %}
</body>
</html>
"""


@app.route("/")
def index():
    if not EXCEL_PATH.exists():
        return (
            f"Файл не найден: {EXCEL_PATH}<br>"
            "Укажите путь к файлу через переменную окружения EXCEL_PATH."
        )

    try:
        metadata, sheets = read_excel_file(EXCEL_PATH)
    except Exception as exc:
        return f"Ошибка чтения Excel: {exc}"

    return render_template_string(TEMPLATE, meta=metadata, sheets=sheets, preview_rows=PREVIEW_ROWS)


@app.route("/grouped")
def grouped_page():
    if not EXCEL_PATH.exists():
        return (
            f"Файл не найден: {EXCEL_PATH}<br>"
            "Укажите путь к файлу через переменную окружения EXCEL_PATH."
        )

    try:
        metadata, grouped_sheets = read_grouped_by_wbs_pairs(EXCEL_PATH)
    except Exception as exc:
        return f"Ошибка чтения Excel: {exc}"

    return render_template_string(GROUPED_TEMPLATE, meta=metadata, grouped_sheets=grouped_sheets)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
