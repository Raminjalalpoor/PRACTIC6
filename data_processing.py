from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import numpy as np
import pandas as pd


REQUIRED_COLUMNS = ("Дата", "Товар", "Количество", "Цена")
OPTIONAL_COLUMNS = ("Категория",)
MONTHS_RU = {
    1: "январь",
    2: "февраль",
    3: "март",
    4: "апрель",
    5: "май",
    6: "июнь",
    7: "июль",
    8: "август",
    9: "сентябрь",
    10: "октябрь",
    11: "ноябрь",
    12: "декабрь",
}
MONTHS_SHORT_RU = {
    1: "янв",
    2: "фев",
    3: "мар",
    4: "апр",
    5: "май",
    6: "июн",
    7: "июл",
    8: "авг",
    9: "сен",
    10: "окт",
    11: "ноя",
    12: "дек",
}


class DataValidationError(ValueError):
    """Понятная пользователю ошибка структуры или содержимого файла."""


@dataclass(frozen=True)
class AnalysisResult:
    metrics: dict
    monthly: pd.DataFrame
    products: pd.DataFrame
    categories: pd.DataFrame
    forecast: dict
    seasonality: dict


def _read_csv(path: Path) -> pd.DataFrame:
    last_error = None
    for encoding in ("utf-8-sig", "cp1251"):
        try:
            return pd.read_csv(
                path,
                sep=None,
                engine="python",
                encoding=encoding,
                dtype=str,
                keep_default_na=False,
            )
        except (UnicodeDecodeError, pd.errors.ParserError) as exc:
            last_error = exc
    raise DataValidationError(
        "Не удалось прочитать CSV. Проверьте кодировку UTF-8/Windows-1251 и разделитель запятая/точка с запятой."
    ) from last_error


def _normalise_number(series: pd.Series) -> pd.Series:
    cleaned = (
        series.astype(str)
        .str.strip()
        .str.replace("\u00a0", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.replace("₽", "", regex=False)
        .str.replace("руб.", "", regex=False)
        .str.replace(",", ".", regex=False)
    )
    return pd.to_numeric(cleaned, errors="coerce")


def _row_list(mask: pd.Series) -> str:
    rows = (mask[mask].index + 2).tolist()
    shown = ", ".join(map(str, rows[:8]))
    return shown + ("…" if len(rows) > 8 else "")


def read_sales_file(path: str | Path) -> pd.DataFrame:
    path = Path(path)
    suffix = path.suffix.lower()
    try:
        if suffix == ".csv":
            frame = _read_csv(path)
        elif suffix == ".xlsx":
            frame = pd.read_excel(path, dtype=str, keep_default_na=False, engine="openpyxl")
        else:
            raise DataValidationError("Разрешены только файлы CSV и XLSX.")
    except DataValidationError:
        raise
    except Exception as exc:
        raise DataValidationError("Файл поврежден или имеет неподдерживаемую структуру.") from exc

    frame = frame.dropna(how="all").copy()
    if frame.empty:
        raise DataValidationError("Файл не содержит данных о продажах.")

    canonical = {column.lower(): column for column in (*REQUIRED_COLUMNS, *OPTIONAL_COLUMNS)}
    renamed = {}
    for column in frame.columns:
        key = str(column).replace("\ufeff", "").strip().lower()
        renamed[column] = canonical.get(key, str(column).strip())
    frame = frame.rename(columns=renamed)

    if frame.columns.duplicated().any():
        raise DataValidationError("В файле обнаружены повторяющиеся названия столбцов.")

    missing = [column for column in REQUIRED_COLUMNS if column not in frame.columns]
    if missing:
        raise DataValidationError("Не хватает обязательных столбцов: " + ", ".join(missing) + ".")

    category_supplied = "Категория" in frame.columns
    frame = frame[[*REQUIRED_COLUMNS, *(OPTIONAL_COLUMNS if category_supplied else ())]].copy()
    frame["Товар"] = frame["Товар"].astype(str).str.strip()
    empty_products = frame["Товар"].eq("")
    if empty_products.any():
        raise DataValidationError(f"Не заполнено название товара в строках: {_row_list(empty_products)}.")

    dates = pd.to_datetime(frame["Дата"].astype(str).str.strip(), format="mixed", dayfirst=True, errors="coerce")
    invalid_dates = dates.isna()
    if invalid_dates.any():
        raise DataValidationError(f"Некорректная дата в строках: {_row_list(invalid_dates)}.")
    frame["Дата"] = dates.dt.normalize()

    for column in ("Количество", "Цена"):
        values = _normalise_number(frame[column])
        invalid = values.isna()
        if invalid.any():
            raise DataValidationError(f"В столбце «{column}» не число в строках: {_row_list(invalid)}.")
        non_positive = values.le(0)
        if non_positive.any():
            raise DataValidationError(
                f"В столбце «{column}» значение должно быть больше нуля. Строки: {_row_list(non_positive)}."
            )
        frame[column] = values.astype(float)

    if category_supplied:
        frame["Категория"] = frame["Категория"].astype(str).str.strip().replace("", "Без категории")
    else:
        frame["Категория"] = "Без категории"

    frame["Выручка"] = frame["Количество"] * frame["Цена"]
    frame = frame.sort_values("Дата", kind="stable").reset_index(drop=True)
    frame.attrs["category_supplied"] = category_supplied
    return frame


def apply_filters(
    frame: pd.DataFrame,
    date_from: str = "",
    date_to: str = "",
    product: str = "",
) -> pd.DataFrame:
    result = frame.copy()
    for value, label, operator in (
        (date_from, "начала периода", "from"),
        (date_to, "окончания периода", "to"),
    ):
        if not value:
            continue
        parsed = pd.to_datetime(value, format="%Y-%m-%d", errors="coerce")
        if pd.isna(parsed):
            raise DataValidationError(f"Некорректная дата {label}.")
        result = result[result["Дата"].ge(parsed) if operator == "from" else result["Дата"].le(parsed)]
    if date_from and date_to and date_from > date_to:
        raise DataValidationError("Дата начала периода не может быть позже даты окончания.")
    if product:
        result = result[result["Товар"].eq(product)]
    result.attrs.update(frame.attrs)
    return result.reset_index(drop=True)


def calculate_metrics(frame: pd.DataFrame) -> dict:
    product_stats = aggregate_products(frame)
    return {
        "revenue": float(frame["Выручка"].sum()),
        "units": float(frame["Количество"].sum()),
        "records": int(len(frame)),
        "unique_products": int(frame["Товар"].nunique()),
        "average_operation": float(frame["Выручка"].mean()),
        "popular_product": str(product_stats.iloc[0]["Товар"]),
        "revenue_leader": str(product_stats.sort_values("Выручка", ascending=False).iloc[0]["Товар"]),
    }


def aggregate_monthly(frame: pd.DataFrame) -> pd.DataFrame:
    monthly = frame.assign(Период=frame["Дата"].dt.to_period("M")).groupby("Период", as_index=False).agg(
        Выручка=("Выручка", "sum"),
        Количество=("Количество", "sum"),
    )
    monthly["Месяц"] = monthly["Период"].map(
        lambda period: f"{MONTHS_SHORT_RU[period.month]} {period.year}"
    )
    return monthly


def aggregate_products(frame: pd.DataFrame) -> pd.DataFrame:
    return (
        frame.groupby("Товар", as_index=False)
        .agg(Количество=("Количество", "sum"), Выручка=("Выручка", "sum"), Записей=("Товар", "size"))
        .sort_values(["Количество", "Выручка", "Товар"], ascending=[False, False, True])
        .reset_index(drop=True)
    )


def aggregate_categories(frame: pd.DataFrame) -> pd.DataFrame:
    return (
        frame.groupby("Категория", as_index=False)
        .agg(Выручка=("Выручка", "sum"), Количество=("Количество", "sum"))
        .sort_values("Выручка", ascending=False)
        .reset_index(drop=True)
    )


def _monthly_quantity(frame: pd.DataFrame, periods: pd.PeriodIndex) -> pd.Series:
    grouped = frame.assign(Период=frame["Дата"].dt.to_period("M")).groupby("Период")["Количество"].sum()
    return grouped.reindex(periods, fill_value=0.0).astype(float)


def calculate_forecast(frame: pd.DataFrame) -> dict:
    min_period = frame["Дата"].min().to_period("M")
    max_period = frame["Дата"].max().to_period("M")
    periods = pd.period_range(min_period, max_period, freq="M")
    next_period = max_period + 1
    unavailable = {
        "available": False,
        "message": "Для прогноза нужно минимум три календарных месяца данных.",
        "next_month": f"{MONTHS_RU[next_period.month]} {next_period.year}",
        "total": None,
        "products": pd.DataFrame(columns=["Товар", "Прогноз", "Последний месяц", "Изменение"]),
    }
    if len(periods) < 3:
        return unavailable

    x = np.arange(len(periods), dtype=float)
    total_series = _monthly_quantity(frame, periods)
    total_prediction = max(0.0, float(np.polyval(np.polyfit(x, total_series.to_numpy(), 1), len(periods))))

    rows = []
    for product, product_frame in frame.groupby("Товар"):
        first_period = product_frame["Дата"].min().to_period("M")
        product_periods = pd.period_range(first_period, max_period, freq="M")
        if len(product_periods) < 3:
            continue
        series = _monthly_quantity(product_frame, product_periods)
        product_x = np.arange(len(series), dtype=float)
        prediction = max(0.0, float(np.polyval(np.polyfit(product_x, series.to_numpy(), 1), len(series))))
        last_value = float(series.iloc[-1])
        change = None if last_value == 0 else (prediction / last_value - 1) * 100
        rows.append(
            {"Товар": product, "Прогноз": prediction, "Последний месяц": last_value, "Изменение": change}
        )

    products = pd.DataFrame(rows, columns=["Товар", "Прогноз", "Последний месяц", "Изменение"])
    if not products.empty:
        products = products.sort_values("Прогноз", ascending=False).reset_index(drop=True)
    return {
        "available": True,
        "message": "Линейный прогноз по истории количества продаж; это ориентир, а не гарантия спроса.",
        "next_month": f"{MONTHS_RU[next_period.month]} {next_period.year}",
        "total": total_prediction,
        "products": products,
    }


def calculate_seasonality(frame: pd.DataFrame, forecast: dict) -> dict:
    min_period = frame["Дата"].min().to_period("M")
    max_period = frame["Дата"].max().to_period("M")
    periods = pd.period_range(min_period, max_period, freq="M")
    if len(periods) < 12:
        return {
            "available": False,
            "message": "Для анализа сезонности нужна история минимум за 12 календарных месяцев.",
            "peak_month": None,
            "insights": [],
            "recommendations": [],
        }

    total = _monthly_quantity(frame, periods)
    month_profile = total.groupby(total.index.month).mean()
    peak_month_number = int(month_profile.idxmax())
    insights = []
    recommendations = []

    for product, product_frame in frame.groupby("Товар"):
        series = _monthly_quantity(product_frame, periods)
        profile = series.groupby(series.index.month).mean().reindex(range(1, 13), fill_value=0.0)
        peak = int(profile.idxmax())
        other_average = float(profile.drop(peak).mean())
        if other_average <= 0:
            continue
        growth = (float(profile.loc[peak]) / other_average - 1) * 100
        if growth >= 20:
            insights.append({"product": product, "month": MONTHS_RU[peak], "growth": growth})

    insights = sorted(insights, key=lambda item: item["growth"], reverse=True)[:5]
    for insight in insights[:3]:
        recommendations.append(
            f"Рассмотрите увеличение запаса товара «{insight['product']}» перед месяцем «{insight['month']}»: "
            f"в истории спрос был выше среднего на {insight['growth']:.0f}%."
        )

    if forecast.get("available") and not forecast["products"].empty:
        growing = forecast["products"].dropna(subset=["Изменение"])
        growing = growing[growing["Изменение"] >= 15].head(2)
        for row in growing.itertuples(index=False):
            recommendations.append(
                f"Проверьте запас товара «{row.Товар}»: модель ожидает рост примерно на {row.Изменение:.0f}% к последнему месяцу."
            )

    if not recommendations:
        recommendations.append("Сильных сезонных сигналов не обнаружено; сохраняйте текущий уровень закупок и обновляйте прогноз ежемесячно.")

    return {
        "available": True,
        "message": "Сезонность рассчитана по среднему количеству продаж для каждого календарного месяца.",
        "peak_month": MONTHS_RU[peak_month_number],
        "insights": insights,
        "recommendations": recommendations[:5],
    }


def analyse(frame: pd.DataFrame) -> AnalysisResult:
    if frame.empty:
        raise DataValidationError("По выбранным фильтрам продаж не найдено.")
    forecast = calculate_forecast(frame)
    return AnalysisResult(
        metrics=calculate_metrics(frame),
        monthly=aggregate_monthly(frame),
        products=aggregate_products(frame),
        categories=aggregate_categories(frame),
        forecast=forecast,
        seasonality=calculate_seasonality(frame, forecast),
    )
