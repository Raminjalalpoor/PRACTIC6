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

## Формат данных

Обязательные столбцы: `Дата`, `Товар`, `Количество`, `Цена`.

Опциональный столбец: `Категория`. Если категории нет, приложение использует значение `Без категории` и показывает пояснение.

Поддерживаются:

- CSV с разделителями `,` и `;`;
- кодировки UTF-8, UTF-8-BOM и Windows-1251;
- десятичная запятая;
- даты в формате `ДД.ММ.ГГГГ` и ISO;
- Excel-файлы `.xlsx`.

## Проверка

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

Текущий результат проверки: `11 tests OK`.

## Отчет по практике

Готовые файлы для преподавателя:

- `docs/practice_report.docx` — отчет Word;
- `docs/practice_report.pdf` — отчет PDF;
- `docs/practice_report.html` — HTML-исходник PDF;
- `docs/screenshots/` — скриншоты главной страницы, дашборда, фильтра и титульного листа отчета.

Пересоздать Word/PDF можно командой:

```powershell
.\.venv\Scripts\python.exe tools\create_practice_report.py
```

Отчет оформлен по образцу лабораторной работы: титульный лист ИТМО, описание цели, задач, структуры проекта, пошаговой работы приложения, контрольных показателей, прогноза, проверки и выводов.
