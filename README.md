# Shop Analytics RU

Русскоязычное Flask-приложение для анализа продаж магазина, визуализации результатов, прогнозирования спроса и экспорта отчета в Excel.

## Запуск на Windows

```powershell
cd "C:\Users\DELL\Desktop\PRACTIC !\shop_analytics"
py -3.14 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python app.py
```

После запуска откройте `http://127.0.0.1:5000`. Демонстрационный CSV лежит в `data/sample_sales.csv`, также его можно скачать с главной страницы приложения.
