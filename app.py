import os
import secrets
from pathlib import Path

from flask import Flask, render_template

from routes import main


BASE_DIR = Path(__file__).resolve().parent


def create_app(test_config=None):
    app = Flask(__name__)
    app.config.from_mapping(
        SECRET_KEY=os.environ.get("SHOP_ANALYTICS_SECRET_KEY", secrets.token_hex(32)),
        MAX_CONTENT_LENGTH=16 * 1024 * 1024,
        UPLOAD_FOLDER=BASE_DIR / "uploads",
        DATA_FOLDER=BASE_DIR / "data",
    )
    if test_config:
        app.config.update(test_config)

    Path(app.config["UPLOAD_FOLDER"]).mkdir(parents=True, exist_ok=True)
    app.register_blueprint(main)

    @app.template_filter("rubles")
    def format_rubles(value):
        return f"{float(value):,.2f}".replace(",", " ").replace(".00", "") + " ₽"

    @app.template_filter("number_ru")
    def format_number(value):
        number = float(value)
        if number.is_integer():
            return f"{int(number):,}".replace(",", " ")
        return f"{number:,.2f}".replace(",", " ").replace(".", ",")

    @app.errorhandler(413)
    def file_too_large(_error):
        return render_template(
            "index.html",
            error="Файл слишком большой. Максимальный размер — 16 МБ.",
        ), 413

    return app


app = create_app()


if __name__ == "__main__":
    app.run(debug=True)
