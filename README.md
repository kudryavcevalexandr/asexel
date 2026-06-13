# Excel Studio

Flask-приложение для загрузки, просмотра, редактирования и скачивания файлов `.xlsx`. Файлы временно сохраняются в `uploads/`; cookie-сессия содержит только уникальное имя текущего файла.

## Запуск

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
FLASK_SECRET_KEY="replace-me" python app.py
```

Откройте `http://127.0.0.1:5000`. Максимальный размер загружаемого файла — 25 МБ.
