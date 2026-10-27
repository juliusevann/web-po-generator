
import os

from flask import Flask, render_template, request, Response

from config import Config
from routes.po_routes import po_bp
from routes.barcode_routes import barcode_bp
from routes.corner_sale_routes import corner_sale_bp
from routes.promo_result_routes import promo_result_bp


app = Flask(__name__)

app.config["SECRET_KEY"] = Config.SECRET_KEY
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 25 MB


app.config["PROMO_RESULT_SCRIPT"] = Config.PROMO_RESULT_SCRIPT
app.config["PROMO_RESULT_INPUT_FOLDER"] = Config.PROMO_RESULT_INPUT_FOLDER
app.config["PROMO_RESULT_OUTPUT_FOLDER"] = Config.PROMO_RESULT_OUTPUT_FOLDER


# --- Create Required Folders ---
os.makedirs(Config.UPLOAD_FOLDER, exist_ok=True)
os.makedirs(Config.OUTPUT_FOLDER, exist_ok=True)
os.makedirs(Config.BARCODE_OUTPUT_FOLDER, exist_ok=True)
os.makedirs(Config.PROMO_RESULT_INPUT_FOLDER, exist_ok=True)
os.makedirs(Config.PROMO_RESULT_OUTPUT_FOLDER, exist_ok=True)


# --- Register Blueprints ---
app.register_blueprint(po_bp)
app.register_blueprint(barcode_bp)
app.register_blueprint(corner_sale_bp)
app.register_blueprint(promo_result_bp)


# --- Basic Auth ---
# Aktif otomatis jika APP_USERNAME dan APP_PASSWORD diisi di .env.
# Berlaku untuk semua halaman aplikasi.

def _auth_enabled():
    return bool(Config.APP_USERNAME and Config.APP_PASSWORD)


def _check_auth(username, password):
    return (
        username == Config.APP_USERNAME
        and password == Config.APP_PASSWORD
    )


def _auth_challenge():
    return Response(
        "Login diperlukan untuk mengakses aplikasi ini.",
        401,
        {"WWW-Authenticate": 'Basic realm="PO & Barcode Generator"'},
    )


@app.before_request
def _require_login():
    if not _auth_enabled():
        return None

    auth = request.authorization

    if not auth or not _check_auth(auth.username, auth.password):
        return _auth_challenge()

    return None


# --- Landing Page ---
@app.route("/")
def landing():
    return render_template("landing.html")


# --- Run Application ---
if __name__ == "__main__":
    app.run(
        host=Config.HOST,
        port=Config.PORT,
        debug=False,
        threaded=True,
    )
