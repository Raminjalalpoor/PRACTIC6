from __future__ import annotations

import time
import uuid
from pathlib import Path

from flask import (
    Blueprint,
    current_app,
    flash,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
from plotly.offline import get_plotlyjs
from werkzeug.utils import secure_filename

from data_processing import DataValidationError, analyse, apply_filters, read_sales_file
from reporting import build_excel_report
from visualization import build_charts


main = Blueprint("main", __name__)
ALLOWED_EXTENSIONS = {".csv", ".xlsx"}


def cleanup_old_uploads(folder: Path, max_age_hours: int = 24):
    cutoff = time.time() - max_age_hours * 60 * 60
    for path in folder.iterdir():
        if path.is_file() and path.name != ".gitkeep" and path.stat().st_mtime < cutoff:
            try:
                path.unlink()
            except OSError:
                current_app.logger.warning("Не удалось удалить временный файл %s", path.name)


def _dataset_path() -> Path | None:
    filename = session.get("dataset_filename")
    if not filename:
        return None
    safe_name = Path(filename).name
    if safe_name != filename:
        return None
    path = Path(current_app.config["UPLOAD_FOLDER"]) / safe_name
    return path if path.is_file() and path.suffix.lower() in ALLOWED_EXTENSIONS else None


def _filters_from_request() -> dict:
    return {
        "date_from": request.args.get("date_from", "").strip(),
        "date_to": request.args.get("date_to", "").strip(),
        "product": request.args.get("product", "").strip(),
    }


@main.get("/")
def index():
    return render_template("index.html")


@main.post("/upload")
def upload():
    upload_folder = Path(current_app.config["UPLOAD_FOLDER"])
    cleanup_old_uploads(upload_folder)

    file = request.files.get("sales_file")
    if not file or not file.filename:
        return render_template("index.html", error="Выберите CSV- или XLSX-файл с данными продаж."), 400

    original_name = secure_filename(file.filename) or "sales"
    suffix = Path(file.filename).suffix.lower()
    if suffix not in ALLOWED_EXTENSIONS:
        return render_template("index.html", error="Неподдерживаемый формат. Разрешены только CSV и XLSX."), 400

    filename = f"{uuid.uuid4().hex}{suffix}"
    path = upload_folder / filename
    file.save(path)
    try:
        read_sales_file(path)
    except DataValidationError as exc:
        path.unlink(missing_ok=True)
        return render_template("index.html", error=str(exc)), 400

    old_path = _dataset_path()
    if old_path and old_path != path:
        old_path.unlink(missing_ok=True)
    session["dataset_filename"] = filename
    session["original_filename"] = file.filename
    flash("Файл успешно проверен. Аналитика готова.", "success")
    return redirect(url_for("main.dashboard"))


@main.get("/dashboard")
def dashboard():
    path = _dataset_path()
    if not path:
        flash("Сначала загрузите файл с данными продаж.", "warning")
        return redirect(url_for("main.index"))

    try:
        full_frame = read_sales_file(path)
        filters = _filters_from_request()
        filtered = apply_filters(full_frame, **filters)
    except DataValidationError as exc:
        flash(str(exc), "danger")
        return redirect(url_for("main.dashboard"))

    products = sorted(full_frame["Товар"].unique(), key=str.casefold)
    context = {
        "filename": session.get("original_filename", "Загруженный файл"),
        "filters": filters,
        "products": products,
        "date_min": full_frame["Дата"].min().strftime("%Y-%m-%d"),
        "date_max": full_frame["Дата"].max().strftime("%Y-%m-%d"),
        "category_supplied": bool(full_frame.attrs.get("category_supplied")),
        "total_rows": len(filtered),
    }
    if filtered.empty:
        context.update(has_data=False, empty_message="По выбранным фильтрам продаж не найдено.")
        return render_template("dashboard.html", **context)

    result = analyse(filtered)
    table = filtered.sort_values("Дата", ascending=False).head(100).copy()
    table["Дата"] = table["Дата"].dt.strftime("%d.%m.%Y")
    context.update(
        has_data=True,
        analysis=result,
        charts=build_charts(result),
        plotly_js=get_plotlyjs(),
        forecast_rows=result.forecast["products"].head(8).to_dict("records"),
        sales_table=table.to_dict("records"),
    )
    return render_template("dashboard.html", **context)


@main.get("/export")
def export_report():
    path = _dataset_path()
    if not path:
        flash("Сессия с данными завершена. Загрузите файл снова.", "warning")
        return redirect(url_for("main.index"))
    try:
        frame = apply_filters(read_sales_file(path), **_filters_from_request())
        result = analyse(frame)
    except DataValidationError as exc:
        flash(str(exc), "danger")
        return redirect(url_for("main.dashboard", **_filters_from_request()))

    report = build_excel_report(frame, result, _filters_from_request())
    return send_file(
        report,
        as_attachment=True,
        download_name="shop_analytics_report.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@main.get("/demo")
def demo_file():
    path = Path(current_app.config["DATA_FOLDER"]) / "sample_sales.csv"
    return send_file(path, as_attachment=True, download_name="primer_prodazh.csv", mimetype="text/csv")
