import os

from flask import Flask, render_template, request, Response

from config import Config
from routes.po_routes import po_bp
from routes.barcode_routes import barcode_bp

app = Flask(__name__)
app.config["SECRET_KEY"] = Config.SECRET_KEY
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 25 MB

os.makedirs(Config.UPLOAD_FOLDER, exist_ok=True)
os.makedirs(Config.OUTPUT_FOLDER, exist_ok=True)
os.makedirs(Config.BARCODE_OUTPUT_FOLDER, exist_ok=True)

app.register_blueprint(po_bp)
app.register_blueprint(barcode_bp)


# --- Basic Auth: aktif otomatis kalau APP_USERNAME & APP_PASSWORD diisi di .env ---
# Wajib diisi kalau web app ini diakses lewat internet (misal via ngrok),
# supaya tidak sembarang orang yang punya link bisa masuk. Berlaku untuk
# SEMUA halaman (landing, PO, barcode) karena dicek di before_request global.
def _auth_enabled():
    return bool(Config.APP_USERNAME and Config.APP_PASSWORD)


def _check_auth(username, password):
    return username == Config.APP_USERNAME and password == Config.APP_PASSWORD


def _auth_challenge():
    return Response(
        "Login diperlukan untuk mengakses aplikasi ini.",
        401,
        {"WWW-Authenticate": 'Basic realm="PO & Barcode Generator"'},
    )


@app.before_request
def _require_login():
    if not _auth_enabled():
        return
    auth = request.authorization
    if not auth or not _check_auth(auth.username, auth.password):
        return _auth_challenge()


@app.route("/")
def landing():
    return render_template("landing.html")


if __name__ == "__main__":
    app.run(host=Config.HOST, port=Config.PORT, debug=False, threaded=True)
