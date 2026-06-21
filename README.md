# Excel Studio

Flask-приложение для загрузки, просмотра, редактирования и скачивания файлов `.xlsx`. Файлы временно сохраняются в `uploads/`; если у процесса нет прав на запись в эту папку, приложение автоматически использует временную директорию системы. Cookie-сессия содержит только уникальное имя текущего файла.

## Запуск

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
FLASK_SECRET_KEY="replace-me" python app.py
```

Если приложение запущено в окружении без прав на запись рядом с кодом, задайте отдельную папку для временных файлов:

```bash
EXCEL_STUDIO_UPLOAD_DIR="/path/to/writable/uploads" FLASK_SECRET_KEY="replace-me" python app.py
```

Откройте `http://127.0.0.1:5000`. Максимальный размер загружаемого файла — 25 МБ.
