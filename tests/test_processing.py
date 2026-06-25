import tempfile
import unittest
from pathlib import Path

import pandas as pd

from data_processing import (
    DataValidationError,
    aggregate_monthly,
    analyse,
    apply_filters,
    calculate_forecast,
    read_sales_file,
)


SAMPLE = """Дата;Товар;Количество;Цена;Категория
01.01.2025;Молоко;10;89,50;Молочные продукты
02.01.2025;Хлеб;20;45;Хлеб и выпечка
01.02.2025;Молоко;12;90;Молочные продукты
01.03.2025;Молоко;14;92;Молочные продукты
"""


class ProcessingTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.folder = Path(self.temp_dir.name)

    def tearDown(self):
        self.temp_dir.cleanup()

    def write_csv(self, content=SAMPLE, encoding="utf-8-sig", name="sales.csv"):
        path = self.folder / name
        path.write_bytes(content.encode(encoding))
        return path

    def test_reads_utf8_semicolon_and_decimal_comma(self):
        frame = read_sales_file(self.write_csv())
        self.assertEqual(len(frame), 4)
        self.assertAlmostEqual(frame.iloc[0]["Цена"], 89.5)
        self.assertTrue(frame.attrs["category_supplied"])

    def test_reads_windows_1251_and_optional_category(self):
        content = "Дата,Товар,Количество,Цена\n2025-01-01,Молоко,10,89\n"
        frame = read_sales_file(self.write_csv(content, "cp1251"))
        self.assertEqual(frame.iloc[0]["Категория"], "Без категории")
        self.assertFalse(frame.attrs["category_supplied"])

    def test_reads_xlsx(self):
        path = self.folder / "sales.xlsx"
        pd.DataFrame(
            {"Дата": ["01.01.2025"], "Товар": ["Молоко"], "Количество": [10], "Цена": [89]}
        ).to_excel(path, index=False)
        frame = read_sales_file(path)
        self.assertEqual(frame.iloc[0]["Выручка"], 890)

    def test_rejects_missing_columns(self):
        path = self.write_csv("Дата;Товар\n01.01.2025;Молоко\n")
        with self.assertRaisesRegex(DataValidationError, "Не хватает"):
            read_sales_file(path)

    def test_rejects_invalid_date_number_and_non_positive(self):
        cases = (
            ("нет;Молоко;1;10", "Некорректная дата"),
            ("01.01.2025;Молоко;много;10", "не число"),
            ("01.01.2025;Молоко;0;10", "больше нуля"),
            ("01.01.2025;Молоко;1;-10", "больше нуля"),
        )
        for index, (row, message) in enumerate(cases):
            with self.subTest(row=row):
                path = self.write_csv("Дата;Товар;Количество;Цена\n" + row + "\n", name=f"bad{index}.csv")
                with self.assertRaisesRegex(DataValidationError, message):
                    read_sales_file(path)

    def test_metrics_months_and_filters(self):
        frame = read_sales_file(self.write_csv())
        result = analyse(frame)
        self.assertAlmostEqual(result.metrics["revenue"], 4163)
        self.assertEqual(result.metrics["units"], 56)
        self.assertEqual(result.metrics["records"], 4)
        self.assertEqual(result.metrics["unique_products"], 2)
        self.assertEqual(result.metrics["popular_product"], "Молоко")
        self.assertEqual(aggregate_monthly(frame)["Месяц"].tolist(), ["янв 2025", "фев 2025", "мар 2025"])
        filtered = apply_filters(frame, date_from="2025-02-01", date_to="2025-03-01", product="Молоко")
        self.assertEqual(len(filtered), 2)
        self.assertEqual(filtered["Количество"].sum(), 26)

    def test_forecast_requires_three_month_span(self):
        short = "Дата;Товар;Количество;Цена\n01.01.2025;A;10;10\n01.02.2025;A;20;10\n"
        frame = read_sales_file(self.write_csv(short))
        self.assertFalse(calculate_forecast(frame)["available"])
        full = read_sales_file(self.write_csv())
        forecast = calculate_forecast(full)
        self.assertTrue(forecast["available"])
        self.assertGreater(forecast["total"], 0)

    def test_seasonality_requires_twelve_months(self):
        frame = read_sales_file(self.write_csv())
        self.assertFalse(analyse(frame).seasonality["available"])
        demo = read_sales_file(Path(__file__).parents[1] / "data" / "sample_sales.csv")
        seasonality = analyse(demo).seasonality
        self.assertTrue(seasonality["available"])
        self.assertEqual(seasonality["peak_month"], "июль")


if __name__ == "__main__":
    unittest.main()
