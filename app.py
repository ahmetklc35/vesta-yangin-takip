from __future__ import annotations

import io
import os
import uuid
from datetime import datetime
from pathlib import Path

import qrcode
from qrcode.constants import ERROR_CORRECT_H
from flask import (
    Flask,
    abort,
    flash,
    redirect,
    render_template,
    request,
    send_file,
    url_for,
)
from openpyxl import Workbook
from PIL import Image, ImageDraw, ImageFont
from sqlalchemy import (
    Column,
    Float,
    ForeignKey,
    Integer,
    MetaData,
    String,
    Table,
    create_engine,
    desc,
    insert,
    select,
    update,
)
from sqlalchemy.engine import Engine
from sqlalchemy.exc import IntegrityError
from werkzeug.middleware.proxy_fix import ProxyFix


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_SQLITE_URL = f"sqlite:///{(BASE_DIR / 'database.db').as_posix()}"
BRAND_NAME = "Vesta Yangin"
LOGO_PATH = BASE_DIR / "static" / "vesta-logo.png"


def build_database_url() -> str:
    raw_url = os.environ.get("DATABASE_URL", DEFAULT_SQLITE_URL)
    if raw_url.startswith("postgres://"):
        return raw_url.replace("postgres://", "postgresql+psycopg://", 1)
    if raw_url.startswith("postgresql://"):
        return raw_url.replace("postgresql://", "postgresql+psycopg://", 1)
    return raw_url


DATABASE_URL = build_database_url()
CONNECT_ARGS = {"check_same_thread": False} if DATABASE_URL.startswith("sqlite") else {}
engine: Engine = create_engine(
    DATABASE_URL,
    future=True,
    pool_pre_ping=True,
    connect_args=CONNECT_ARGS,
)

metadata = MetaData()

extinguishers = Table(
    "extinguishers",
    metadata,
    Column("id", Integer, primary_key=True),
    Column("public_id", String(32), nullable=False, unique=True),
    Column("serial_number", String(128), nullable=False, unique=True),
    Column("company_name", String(255), nullable=False),
    Column("location_detail", String(255), nullable=False),
    Column("weight_kg", Float, nullable=False),
    Column("extinguisher_type", String(255), nullable=False),
    Column("pressure_status", String(255)),
    Column("notes", String),
    Column("last_service_date", String(32)),
    Column("next_service_date", String(32)),
    Column("created_at", String(32), nullable=False),
    Column("updated_at", String(32), nullable=False),
)

service_logs = Table(
    "service_logs",
    metadata,
    Column("id", Integer, primary_key=True),
    Column("extinguisher_id", Integer, ForeignKey("extinguishers.id"), nullable=False),
    Column("service_date", String(32), nullable=False),
    Column("technician_name", String(255), nullable=False),
    Column("operation_summary", String, nullable=False),
    Column("pressure_status", String(255)),
    Column("notes", String),
    Column("created_at", String(32), nullable=False),
)

metadata.create_all(engine)

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "change-this-secret-key")
app.wsgi_app = ProxyFix(app.wsgi_app, x_proto=1, x_host=1)  # type: ignore[assignment]


def fetch_all(statement):
    with engine.connect() as connection:
        result = connection.execute(statement)
        return [dict(row._mapping) for row in result]


def fetch_one(statement):
    with engine.connect() as connection:
        row = connection.execute(statement).mappings().first()
        return dict(row) if row else None


def parse_float(value: str, field_name: str) -> float:
    try:
        return float(value.replace(",", "."))
    except ValueError as exc:
        raise ValueError(f"{field_name} sayi olmali.") from exc


def parse_required_form(form: dict[str, str]) -> dict[str, str]:
    return {key: value.strip() for key, value in form.items()}


def get_extinguisher(public_id: str) -> dict:
    extinguisher = fetch_one(
        select(extinguishers).where(extinguishers.c.public_id == public_id)
    )
    if extinguisher is None:
        abort(404)
    return extinguisher


def autosize_worksheet(worksheet) -> None:
    for column_cells in worksheet.columns:
        max_length = 0
        column_letter = column_cells[0].column_letter
        for cell in column_cells:
            value = "" if cell.value is None else str(cell.value)
            if len(value) > max_length:
                max_length = len(value)
        worksheet.column_dimensions[column_letter].width = min(max_length + 2, 40)


def load_font(size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    font_candidates = [
        Path("C:/Windows/Fonts/arialbd.ttf"),
        Path("C:/Windows/Fonts/segoeuib.ttf"),
        Path("C:/Windows/Fonts/arial.ttf"),
    ]
    for font_path in font_candidates:
        if font_path.exists():
            try:
                return ImageFont.truetype(str(font_path), size)
            except OSError:
                continue
    return ImageFont.load_default()


def build_branded_qr(public_url: str) -> io.BytesIO:
    qr = qrcode.QRCode(
        version=None,
        error_correction=ERROR_CORRECT_H,
        box_size=10,
        border=4,
    )
    qr.add_data(public_url)
    qr.make(fit=True)
    qr_image = qr.make_image(fill_color="black", back_color="white").convert("RGBA")
    qr_width, qr_height = qr_image.size

    if LOGO_PATH.exists():
        logo = Image.open(LOGO_PATH).convert("RGBA")
        max_logo_size = int(qr_width * 0.24)
        logo.thumbnail((max_logo_size, max_logo_size))

        logo_bg_size = max(logo.width, logo.height) + 20
        logo_bg = Image.new("RGBA", (logo_bg_size, logo_bg_size), (255, 255, 255, 255))
        bg_x = (logo_bg_size - logo.width) // 2
        bg_y = (logo_bg_size - logo.height) // 2
        logo_bg.alpha_composite(logo, (bg_x, bg_y))

        logo_x = (qr_width - logo_bg_size) // 2
        logo_y = (qr_height - logo_bg_size) // 2
        qr_image.alpha_composite(logo_bg, (logo_x, logo_y))

    canvas = Image.new("RGBA", (qr_width, qr_height + 78), "white")
    canvas.paste(qr_image, (0, 0))

    draw = ImageDraw.Draw(canvas)
    font = load_font(28)
    text_bbox = draw.textbbox((0, 0), BRAND_NAME, font=font)
    text_width = text_bbox[2] - text_bbox[0]
    text_x = (qr_width - text_width) // 2
    text_y = qr_height + 24
    draw.text((text_x, text_y), BRAND_NAME, fill="black", font=font)

    buffer = io.BytesIO()
    canvas.save(buffer, format="PNG")
    buffer.seek(0)
    return buffer


@app.route("/")
def index():
    latest_log_date = (
        select(service_logs.c.service_date)
        .where(service_logs.c.extinguisher_id == extinguishers.c.id)
        .order_by(desc(service_logs.c.service_date), desc(service_logs.c.id))
        .limit(1)
        .scalar_subquery()
    )
    statement = (
        select(extinguishers, latest_log_date.label("latest_log_date"))
        .order_by(desc(extinguishers.c.updated_at))
    )
    return render_template("index.html", extinguishers=fetch_all(statement))


@app.route("/export/extinguishers.xlsx")
def export_extinguishers():
    rows = fetch_all(
        select(
            extinguishers.c.serial_number,
            extinguishers.c.company_name,
            extinguishers.c.location_detail,
            extinguishers.c.weight_kg,
            extinguishers.c.extinguisher_type,
            extinguishers.c.pressure_status,
            extinguishers.c.last_service_date,
            extinguishers.c.next_service_date,
            extinguishers.c.notes,
            extinguishers.c.created_at,
            extinguishers.c.updated_at,
        ).order_by(desc(extinguishers.c.updated_at))
    )

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Tupler"
    headers = [
        "Seri No",
        "Firma",
        "Konum",
        "Kg",
        "Tip",
        "Basinc",
        "Son Bakim",
        "Sonraki Bakim",
        "Notlar",
        "Kayit Tarihi",
        "Guncelleme Tarihi",
    ]
    sheet.append(headers)
    for row in rows:
        sheet.append(list(row.values()))

    autosize_worksheet(sheet)
    buffer = io.BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    return send_file(
        buffer,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="yangin-tupleri.xlsx",
    )


@app.route("/export/service-logs.xlsx")
def export_service_logs():
    statement = (
        select(
            extinguishers.c.serial_number,
            extinguishers.c.company_name,
            extinguishers.c.location_detail,
            service_logs.c.service_date,
            service_logs.c.technician_name,
            service_logs.c.operation_summary,
            service_logs.c.pressure_status,
            service_logs.c.notes,
            service_logs.c.created_at,
        )
        .join(service_logs, service_logs.c.extinguisher_id == extinguishers.c.id)
        .order_by(desc(service_logs.c.service_date), desc(service_logs.c.id))
    )
    rows = fetch_all(statement)

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Bakim Gecmisi"
    headers = [
        "Seri No",
        "Firma",
        "Konum",
        "Bakim Tarihi",
        "Teknisyen",
        "Yapilan Islem",
        "Basinc",
        "Notlar",
        "Kayit Zamani",
    ]
    sheet.append(headers)
    for row in rows:
        sheet.append(list(row.values()))

    autosize_worksheet(sheet)
    buffer = io.BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    return send_file(
        buffer,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="bakim-gecmisi.xlsx",
    )


@app.route("/extinguishers/new", methods=["GET", "POST"])
def create_extinguisher():
    if request.method == "POST":
        form = parse_required_form(request.form)
        required_fields = {
            "serial_number": "Seri numarasi",
            "company_name": "Firma adi",
            "location_detail": "Firma ici konum",
            "weight_kg": "Kg bilgisi",
            "extinguisher_type": "Tup tipi",
            "last_service_date": "Son bakim tarihi",
            "next_service_date": "Sonraki bakim tarihi",
            "technician_name": "Teknisyen",
            "operation_summary": "Yapilan islem",
        }
        missing = [label for key, label in required_fields.items() if not form.get(key)]
        if missing:
            flash(f"Eksik alanlar: {', '.join(missing)}", "error")
            return render_template("create_extinguisher.html", form=form)

        try:
            weight_kg = parse_float(form["weight_kg"], "Kg")
        except ValueError as exc:
            flash(str(exc), "error")
            return render_template("create_extinguisher.html", form=form)

        now = datetime.now().isoformat(timespec="seconds")
        public_id = uuid.uuid4().hex[:12]

        try:
            with engine.begin() as connection:
                result = connection.execute(
                    insert(extinguishers).values(
                        public_id=public_id,
                        serial_number=form["serial_number"],
                        company_name=form["company_name"],
                        location_detail=form["location_detail"],
                        weight_kg=weight_kg,
                        extinguisher_type=form["extinguisher_type"],
                        pressure_status=form.get("pressure_status"),
                        notes=form.get("notes"),
                        last_service_date=form["last_service_date"],
                        next_service_date=form["next_service_date"],
                        created_at=now,
                        updated_at=now,
                    )
                )
                extinguisher_id = result.inserted_primary_key[0]
                connection.execute(
                    insert(service_logs).values(
                        extinguisher_id=extinguisher_id,
                        service_date=form["last_service_date"],
                        technician_name=form["technician_name"],
                        operation_summary=form["operation_summary"],
                        pressure_status=form.get("pressure_status"),
                        notes=form.get("notes"),
                        created_at=now,
                    )
                )
        except IntegrityError:
            flash("Bu seri numarasi zaten kayitli.", "error")
            return render_template("create_extinguisher.html", form=form)

        flash("Tup kaydedildi ve QR olusturuldu.", "success")
        return redirect(url_for("extinguisher_detail", public_id=public_id))

    return render_template("create_extinguisher.html", form={})


@app.route("/extinguishers/<public_id>")
def extinguisher_detail(public_id: str):
    extinguisher = get_extinguisher(public_id)
    logs = fetch_all(
        select(service_logs)
        .where(service_logs.c.extinguisher_id == extinguisher["id"])
        .order_by(desc(service_logs.c.service_date), desc(service_logs.c.id))
    )
    return render_template(
        "extinguisher_detail.html",
        extinguisher=extinguisher,
        service_logs=logs,
    )


@app.route("/extinguishers/<public_id>/service", methods=["GET", "POST"])
def add_service_log(public_id: str):
    extinguisher = get_extinguisher(public_id)
    if request.method == "POST":
        form = parse_required_form(request.form)
        required_fields = {
            "service_date": "Bakim tarihi",
            "next_service_date": "Sonraki bakim tarihi",
            "technician_name": "Teknisyen",
            "operation_summary": "Yapilan islem",
        }
        missing = [label for key, label in required_fields.items() if not form.get(key)]
        if missing:
            flash(f"Eksik alanlar: {', '.join(missing)}", "error")
            return render_template(
                "service_log_form.html",
                extinguisher=extinguisher,
                form=form,
            )

        try:
            weight_kg = parse_float(
                form.get("weight_kg") or str(extinguisher["weight_kg"]),
                "Kg",
            )
        except ValueError as exc:
            flash(str(exc), "error")
            return render_template(
                "service_log_form.html",
                extinguisher=extinguisher,
                form=form,
            )

        now = datetime.now().isoformat(timespec="seconds")
        with engine.begin() as connection:
            connection.execute(
                insert(service_logs).values(
                    extinguisher_id=extinguisher["id"],
                    service_date=form["service_date"],
                    technician_name=form["technician_name"],
                    operation_summary=form["operation_summary"],
                    pressure_status=form.get("pressure_status"),
                    notes=form.get("notes"),
                    created_at=now,
                )
            )
            connection.execute(
                update(extinguishers)
                .where(extinguishers.c.id == extinguisher["id"])
                .values(
                    company_name=form.get("company_name") or extinguisher["company_name"],
                    location_detail=form.get("location_detail")
                    or extinguisher["location_detail"],
                    weight_kg=weight_kg,
                    extinguisher_type=form.get("extinguisher_type")
                    or extinguisher["extinguisher_type"],
                    pressure_status=form.get("pressure_status"),
                    notes=form.get("notes"),
                    last_service_date=form["service_date"],
                    next_service_date=form["next_service_date"],
                    updated_at=now,
                )
            )

        flash("Bakim kaydi eklendi.", "success")
        return redirect(url_for("extinguisher_detail", public_id=public_id))

    return render_template(
        "service_log_form.html",
        extinguisher=extinguisher,
        form={},
    )


@app.route("/public/<public_id>")
def public_detail(public_id: str):
    extinguisher = get_extinguisher(public_id)
    latest_log = fetch_one(
        select(service_logs)
        .where(service_logs.c.extinguisher_id == extinguisher["id"])
        .order_by(desc(service_logs.c.service_date), desc(service_logs.c.id))
        .limit(1)
    )
    return render_template(
        "public_detail.html",
        extinguisher=extinguisher,
        latest_log=latest_log,
    )


@app.route("/extinguishers/<public_id>/qr")
def extinguisher_qr(public_id: str):
    get_extinguisher(public_id)
    public_url = url_for("public_detail", public_id=public_id, _external=True)
    buffer = build_branded_qr(public_url)
    return send_file(buffer, mimetype="image/png", download_name=f"{public_id}.png")


@app.route("/extinguishers/<public_id>/label")
def extinguisher_label(public_id: str):
    get_extinguisher(public_id)
    return render_template("label.html", public_id=public_id)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "5000")), debug=True)
