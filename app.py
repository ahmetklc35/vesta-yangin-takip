from __future__ import annotations

import io
import json
import os
import uuid
from datetime import datetime
from pathlib import Path
from functools import wraps

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
    session,
    url_for,
)
from openpyxl import Workbook
from PIL import Image, ImageDraw, ImageFont
from sqlalchemy import (
    Boolean,
    Column,
    Float,
    ForeignKey,
    Integer,
    MetaData,
    String,
    Table,
    text,
    create_engine,
    desc,
    insert,
    select,
    update,
)
from sqlalchemy.engine import Engine
from sqlalchemy.exc import IntegrityError
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.middleware.proxy_fix import ProxyFix


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_SQLITE_URL = f"sqlite:///{(BASE_DIR / 'database.db').as_posix()}"
BRAND_NAME = "Vesta Yangin"
LOGO_PATH = BASE_DIR / "static" / "vesta qr.png"
EQUIPMENT_OPTIONS = [
    "Kuru Kimyevi Toz",
    "CO2",
    "Kopuk",
]
EQUIPMENT_PRESETS = {
    "Kopuk": {
        "title": "Taşınılabilir Yangın Söndürücü, 9 Litre Köpük",
        "summary": "YANGIN SINIFI: 21 A 183 B - 9 litrelik köpük yangın söndürücü , -30°C ila +60°C sıcaklık aralığında, UNI EN 3-7 standardına uygun olarak üretilmiş , MED 2014/90/EU Denizcilik Ekipmanları Direktifi onaylı, PED 2014/68/EU Basınçlı Ekipman Direktifi'ne göre sertifikalandırılmıştır . Üretim kontrolleri EN 3-10 standardına uygun olarak yapılmıştır . - Elektrik voltajı 1.000 V'a kadar olan yangınlarda, en az 1 metre mesafede kullanıma uygundur.",
        "features": "Özellikler: - SİLİNDİR: Yüksek mukavemetli alaşımlı çelikten, 3 parçalı klipsli, dış yüzeyi RAL 3000 Kırmızı toz boya ile boyanmış. -  SÖNDÜRÜCÜ MADDE: Su bazlı köpük. - İTİCİ AKIŞKAN : Nemi alınmış hava veya azot (N2). - VALF: M. 30x1.5, pirinç gövde, kolları RAL 6029 yeşil boya ile boyanmış. - KULLANIM: AB yangın sınıfı (katı maddeler, yanıcı sıvılar).",
        "technical_details": "Teknik Özellikler: -  YANGIN DAYANIKLILIĞI: 21 A 183 B -  SÖNDÜRÜCÜ MADDE: Köpük EW-30 (%100 Premix MG6-30) -  İTİCİ AKIŞKAN:  Nemi alınmış hava veya Azot (N2), 20°C'de 15 bar -  SICAKLIK ARALIĞI: -30°C / +60°C -  NOMİNAL HACİM:  9 Litre -  TOPLAM AĞIRLIK: ~ 15,1 Kg -  BOYUTLAR: Yükseklik (taban - valf) 635 +/- 5 mm - Çap (silindir) 170 +/- 2 mm -  BOŞALTMA SÜRESİ: ~ 34,8 saniye -  VALF SIKMA TORKU: Minimum 45 Nm, Maksimum 68 Nm -  SİLİNDİR BASINÇ  TESTİ : PT 27 bar -  SİLİNDİR HACMİ: 10,7 L. -  GÜVENLİK CİHAZI: Ayarlanmış 20 ile 26 bar arasında -  SİLİNDİR MALZEMESİ: Alaşımlı çelik -  İŞLEM: Dış yüzey: Kum püskürtme ve toz boyama, RAL 3000 - İç yüzey: Plastik kaplama",
        "image": "equipment/kopuk.png",
    },
    "CO2": {
        "title": "Taşınılabilir Yangın Söndürücü, 5 Kg CO₂",
        "summary": "YANGIN SINIFI: 3 7 113 B  - 5 kg Karbondioksitli Yangın Söndürücü, çalışma sıcaklığı -30°C ile +60°C arasında, UNI EN 3-7 (DM 7.1.2005) standardına uygun olarak üretilmiş, Denizcilik Ekipmanları Direktifi MED 2014/90/UE tarafından onaylanmış, Basınçlı Ekipman Direktifi PED 2014/68/EU'ya göre sertifikalandırılmıştır. EN 3-10 ile mutabık kalınan üretim kontrollerine göre üretilmiştir. - Elektrik voltajı 1.000 V'a kadar olan yangınlarda, en az 1 metre mesafede kullanıma uygundur.",
        "features": "Özellikler: - SİLİNDİR: Alaşımlı çelik, dış yüzey toz boya kaplama, renk Kırmızı RAL 3000. - SÖNDÜRÜCÜ MADDE: Karbondioksit %99,99 (CO2) - VALF: Levian M. 25x2, pirinç gövde, kol kırmızı RAL 3000. - KULLANIM: Yangın dayanıklılık sınıfı B (yanıcı sıvılar).",
        "technical_details": "Teknik Özellikler: • YANGIN DAYANIKLILIĞI 113 B • SÖNDÜRÜCÜ MADDE Karbondioksit %99,99 (CO2) • SICAKLIK ARALIĞI -30°C/ +60°C • 5 kg için nominal ücret. • TOPLAM AĞIRLIK ~ 13,65 Kg • BOYUTLAR Yükseklik (taban - valf) 740 +/- 10 mm - Çap (silindir) 136 +/- 2 mm • Boşalma süresi  ~ 15,6 saniye",
        "image": "equipment/co2.png",
    },
    "Kuru Kimyevi Toz": {
        "title": "Taşınılabilir Yangın Söndürücü, 6 Kg ABC Tozu",
        "summary": "YANGIN SINIFI: 34 A 233 B C  - 6 kg tozlu yangın söndürücü, -30°C ila +60°C sıcaklık aralığında, UNI EN 3-7 (DM 7.1.2005) standardına uygun olarak üretilmiş, PED 2014/68/EU basınçlı ekipman direktifine göre sertifikalandırılmıştır. Üretim kontrolleri EN 3-10 standardına uygun olarak gerçekleştirilmiştir. - Elektrik voltajı 1.000 V'a kadar olan yangınlarda, en az 1 metre mesafede kullanıma uygundur.",
        "features": "Özellikler: - Silindir: Alaşımlı çelik, derin çekme, klipsli, epoksipolyester toz boya, Kırmızı RAL 3000. - Söndürücü Madde: ABC Tozu - MAP %20. - İtici Gaz: Nemi alınmış hava veya N2. - Vana: M. 58x2, hafif alüminyum alaşım AA6061 gövde, harici çek valfli, kırmızı RAL 3000 boyalı kollar. - Kullanım: ABC yangın sınıflandırması (katı maddeler, yanıcı sıvılar, yanıcı gazlar).",
        "technical_details": "Teknik Özellikler: - Yangın Dayanımı: 34 A 233 BC - Söndürücü Madde: EPW 18462 (ABC Favorit Tertia) - ABC tozu - MAP %20 - İtici Gaz: Nemi alınmış hava veya N2, 20°C'de 15 Bar - Sıcaklık Aralığı: -30°C / +60°C - Nominal Şarj: 6 Kg - Tam Ağırlık: ~ 9,4 Kg - Boyutlar: Yükseklik 550 mm, Çap 160 mm - Boşaltma Süresi:  16,5 sn. - VALF SIKMA TORKU: Minimum 40 Nm, Maksimum 60 Nm - Silindir Basınç Testi: PT 27 bar - Silindir Hacmi: 7,8 L. - Emniyet Valfi: 22 ile 27 bar arasında ayarlanmıştır - Silindir Malzemesi: Alaşımlı çelik - Dış/İç İşlem: Kum püskürtme ve epoksipolyester toz boya, Kırmızı Ral 3000 rengi",
        "image": "equipment/kuru-kimyevi-toz.png",
    },
}
MONTHLY_CONTROL_ITEMS = [
    ("item_1", "17.M.1001.A.1 YSC Konumu Değiştirilmemiş (belirlenen yerde duruyor)"),
    ("item_2", "17.M.1001.A.2 YSC Kullanım Talimatı Kolay Okunabiliyor"),
    ("item_3", "17.M.1001.A.3 YSC Mühür ve Basınç Göstergesinin uygunluğu"),
    ("item_4", "17.M.1001.A.4 YSC Doluluk Durumu (el ile tartarak kontrol edilmeli)"),
    ("item_5", "17.M.1001.A.5 YSC Paslanmamış ve Nozulda Tıkanıklık veya sızdırma yok"),
    ("item_6", "17.M.1001.A.6 YSC Manometresinden Okunan Basınç Kabul Edilebilir aralıkta"),
]
MONTH_LABELS = [
    (1, "OCAK"),
    (2, "ŞUBAT"),
    (3, "MART"),
    (4, "NİSAN"),
    (5, "MAYIS"),
    (6, "HAZİRAN"),
    (7, "TEMMUZ"),
    (8, "AĞUSTOS"),
    (9, "EYLÜL"),
    (10, "EKİM"),
    (11, "KASIM"),
    (12, "ARALIK"),
]


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

users = Table(
    "users",
    metadata,
    Column("id", Integer, primary_key=True),
    Column("username", String(128), nullable=False, unique=True),
    Column("full_name", String(255), nullable=False),
    Column("password_hash", String(255), nullable=False),
    Column("is_admin", Boolean, nullable=False, default=False),
    Column("is_active", Boolean, nullable=False, default=True),
    Column("created_at", String(32), nullable=False),
    Column("updated_at", String(32), nullable=False),
)

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

monthly_inspections = Table(
    "monthly_inspections",
    metadata,
    Column("id", Integer, primary_key=True),
    Column("extinguisher_id", Integer, ForeignKey("extinguishers.id"), nullable=False),
    Column("inspection_date", String(32), nullable=False),
    Column("inspector_name", String(255), nullable=False),
    Column("item_1", Boolean, nullable=False),
    Column("item_2", Boolean, nullable=False),
    Column("item_3", Boolean, nullable=False),
    Column("item_4", Boolean, nullable=False),
    Column("item_5", Boolean, nullable=False),
    Column("item_6", Boolean, nullable=False),
    Column("notes", String),
    Column("created_at", String(32), nullable=False),
)

metadata.create_all(engine)

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "change-this-secret-key")
app.wsgi_app = ProxyFix(app.wsgi_app, x_proto=1, x_host=1)  # type: ignore[assignment]
ADMIN_USERNAME = os.environ.get("ADMIN_USERNAME", "vestaadmin")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "degistir-beni-2026")
DEFAULT_USER_PASSWORD = os.environ.get("DEFAULT_USER_PASSWORD", "Vesta123!")
DEFAULT_USERS = [
    {
        "username": ADMIN_USERNAME,
        "full_name": "Admin",
        "password": ADMIN_PASSWORD,
        "is_admin": True,
    },
    {
        "username": "yunus.emre.duman",
        "full_name": "Yunus Emre Duman",
        "password": DEFAULT_USER_PASSWORD,
        "is_admin": False,
    },
    {
        "username": "atil.bati",
        "full_name": "Atıl Batı",
        "password": DEFAULT_USER_PASSWORD,
        "is_admin": False,
    },
    {
        "username": "mustafa.kilic",
        "full_name": "Mustafa Kiliç",
        "password": DEFAULT_USER_PASSWORD,
        "is_admin": False,
    },
]


def seed_default_users() -> None:
    now = datetime.now().isoformat(timespec="seconds")
    with engine.begin() as connection:
        existing = {
            row[0]
            for row in connection.execute(select(users.c.username)).all()
        }
        for user in DEFAULT_USERS:
            if user["username"] in existing:
                continue
            connection.execute(
                insert(users).values(
                    username=user["username"],
                    full_name=user["full_name"],
                    password_hash=generate_password_hash(user["password"]),
                    is_admin=user["is_admin"],
                    is_active=True,
                    created_at=now,
                    updated_at=now,
                )
            )


def run_schema_migrations() -> None:
    if DATABASE_URL.startswith("sqlite"):
        with engine.begin() as connection:
            columns = {
                row[1] for row in connection.execute(text("PRAGMA table_info(users)")).fetchall()
            }
            if "is_active" not in columns:
                connection.execute(text("ALTER TABLE users ADD COLUMN is_active BOOLEAN NOT NULL DEFAULT 1"))
    else:
        with engine.begin() as connection:
            result = connection.execute(
                text(
                    """
                    SELECT column_name
                    FROM information_schema.columns
                    WHERE table_name = 'users' AND column_name = 'is_active'
                    """
                )
            ).fetchone()
            if result is None:
                connection.execute(
                    text("ALTER TABLE users ADD COLUMN is_active BOOLEAN NOT NULL DEFAULT TRUE")
                )


run_schema_migrations()
seed_default_users()


def is_authenticated() -> bool:
    return session.get("authenticated") is True


def current_user_full_name() -> str:
    return session.get("full_name", "")


def is_admin_user() -> bool:
    return session.get("is_admin") is True


def login_required(view_func):
    @wraps(view_func)
    def wrapped_view(*args, **kwargs):
        if not is_authenticated():
            return redirect(url_for("login", next=request.path))
        return view_func(*args, **kwargs)

    return wrapped_view


def admin_required(view_func):
    @wraps(view_func)
    def wrapped_view(*args, **kwargs):
        if not is_authenticated():
            return redirect(url_for("login", next=request.path))
        if not is_admin_user():
            abort(403)
        return view_func(*args, **kwargs)

    return wrapped_view


@app.before_request
def protect_private_routes():
    allowed_endpoints = {"login", "logout", "public_detail", "health", "static"}
    if request.endpoint in allowed_endpoints or request.endpoint is None:
        return None
    if not is_authenticated():
        return redirect(url_for("login", next=request.path))
    return None


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


def build_monthly_inspection_values(form_data) -> dict[str, bool]:
    return {key: form_data.get(key) == "on" for key, _label in MONTHLY_CONTROL_ITEMS}


def save_monthly_inspection(
    connection,
    extinguisher_id: int,
    inspection_date: str,
    inspector_name: str,
    notes: str | None,
    inspection_values: dict[str, bool],
    created_at: str,
) -> None:
    connection.execute(
        insert(monthly_inspections).values(
            extinguisher_id=extinguisher_id,
            inspection_date=inspection_date,
            inspector_name=inspector_name,
            notes=notes,
            created_at=created_at,
            **inspection_values,
        )
    )


def with_monthly_control_labels(rows: list[dict]) -> list[dict]:
    enriched_rows: list[dict] = []
    for row in rows:
        checks = []
        passed_count = 0
        for key, label in MONTHLY_CONTROL_ITEMS:
            passed = bool(row.get(key))
            if passed:
                passed_count += 1
            checks.append({"key": key, "label": label, "passed": passed})
        copied = dict(row)
        copied["checks"] = checks
        copied["passed_count"] = passed_count
        copied["total_count"] = len(MONTHLY_CONTROL_ITEMS)
        enriched_rows.append(copied)
    return enriched_rows


def get_equipment_preset(extinguisher_type: str | None) -> dict | None:
    if not extinguisher_type:
        return None
    return EQUIPMENT_PRESETS.get(extinguisher_type)


def build_monthly_table(rows: list[dict]) -> dict:
    if rows:
        sorted_rows = sorted(rows, key=lambda row: row["inspection_date"], reverse=True)
        target_year = datetime.strptime(sorted_rows[0]["inspection_date"], "%Y-%m-%d").year
    else:
        target_year = datetime.now().year

    latest_by_month: dict[int, dict] = {}
    for row in sorted(rows, key=lambda row: row["inspection_date"], reverse=True):
        inspection_date = datetime.strptime(row["inspection_date"], "%Y-%m-%d")
        if inspection_date.year != target_year:
            continue
        if inspection_date.month not in latest_by_month:
            latest_by_month[inspection_date.month] = row

    month_rows = []
    for month_number, month_label in MONTH_LABELS:
        source = latest_by_month.get(month_number)
        cells = []
        for key, label in MONTHLY_CONTROL_ITEMS:
            code = label.split(" ", 1)[0]
            value = None if source is None else bool(source.get(key))
            cells.append({"code": code, "value": value})
        month_rows.append(
            {
                "month_label": month_label,
                "inspection_date": source["inspection_date"] if source else None,
                "cells": cells,
            }
        )

    return {
        "year": target_year,
        "headers": [
            {"key": key, "code": label.split(" ", 1)[0]}
            for key, label in MONTHLY_CONTROL_ITEMS
        ],
        "rows": month_rows,
    }


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        user = fetch_one(select(users).where(users.c.username == username))
        if user and bool(user["is_active"]) and check_password_hash(user["password_hash"], password):
            session["authenticated"] = True
            session["user_id"] = user["id"]
            session["username"] = user["username"]
            session["full_name"] = user["full_name"]
            session["is_admin"] = bool(user["is_admin"])
            flash("Giris yapildi.", "success")
            next_url = request.args.get("next") or url_for("index")
            return redirect(next_url)
        flash("Kullanici adi veya sifre hatali.", "error")
    return render_template("login.html")


@app.route("/health")
def health():
    return {"status": "ok"}, 200


@app.route("/logout")
def logout():
    session.clear()
    flash("Cikis yapildi.", "success")
    return redirect(url_for("login"))


@app.route("/profile", methods=["GET", "POST"])
@login_required
def profile():
    current_user = fetch_one(select(users).where(users.c.id == session.get("user_id")))
    if current_user is None:
        session.clear()
        return redirect(url_for("login"))

    if request.method == "POST":
        current_password = request.form.get("current_password", "")
        new_password = request.form.get("new_password", "")
        confirm_password = request.form.get("confirm_password", "")

        if not check_password_hash(current_user["password_hash"], current_password):
            flash("Mevcut sifre hatali.", "error")
            return redirect(url_for("profile"))
        if not new_password:
            flash("Yeni sifre bos olamaz.", "error")
            return redirect(url_for("profile"))
        if new_password != confirm_password:
            flash("Yeni sifreler eslesmiyor.", "error")
            return redirect(url_for("profile"))

        now = datetime.now().isoformat(timespec="seconds")
        with engine.begin() as connection:
            connection.execute(
                update(users)
                .where(users.c.id == current_user["id"])
                .values(
                    password_hash=generate_password_hash(new_password),
                    updated_at=now,
                )
            )
        flash("Sifren guncellendi.", "success")
        return redirect(url_for("profile"))

    return render_template("profile.html", current_user=current_user)


@app.route("/users", methods=["GET"])
@admin_required
def user_management():
    user_rows = fetch_all(
        select(
            users.c.id,
            users.c.username,
            users.c.full_name,
            users.c.is_admin,
            users.c.is_active,
            users.c.created_at,
        ).order_by(users.c.full_name)
    )
    return render_template("user_management.html", users=user_rows)


@app.route("/users/create", methods=["POST"])
@admin_required
def create_user():
    username = request.form.get("username", "").strip()
    full_name = request.form.get("full_name", "").strip()
    password = request.form.get("password", "")
    is_admin = request.form.get("is_admin") == "on"

    if not username or not full_name or not password:
        flash("Kullanici adi, ad soyad ve sifre gerekli.", "error")
        return redirect(url_for("user_management"))

    now = datetime.now().isoformat(timespec="seconds")
    try:
        with engine.begin() as connection:
            connection.execute(
                insert(users).values(
                    username=username,
                    full_name=full_name,
                    password_hash=generate_password_hash(password),
                    is_admin=is_admin,
                    is_active=True,
                    created_at=now,
                    updated_at=now,
                )
            )
    except IntegrityError:
        flash("Bu kullanici adi zaten mevcut.", "error")
        return redirect(url_for("user_management"))

    flash("Kullanici olusturuldu.", "success")
    return redirect(url_for("user_management"))


@app.route("/users/<int:user_id>/password", methods=["POST"])
@admin_required
def update_user_password(user_id: int):
    password = request.form.get("password", "")
    if not password:
        flash("Yeni sifre bos olamaz.", "error")
        return redirect(url_for("user_management"))

    now = datetime.now().isoformat(timespec="seconds")
    with engine.begin() as connection:
        connection.execute(
            update(users)
            .where(users.c.id == user_id)
            .values(
                password_hash=generate_password_hash(password),
                updated_at=now,
            )
        )
    flash("Sifre guncellendi.", "success")
    return redirect(url_for("user_management"))


@app.route("/users/<int:user_id>/toggle-active", methods=["POST"])
@admin_required
def toggle_user_active(user_id: int):
    target_user = fetch_one(select(users).where(users.c.id == user_id))
    if not target_user:
        abort(404)
    if target_user["username"] == session.get("username"):
        flash("Kendi hesabini pasife alamazsin.", "error")
        return redirect(url_for("user_management"))

    now = datetime.now().isoformat(timespec="seconds")
    with engine.begin() as connection:
        connection.execute(
            update(users)
            .where(users.c.id == user_id)
            .values(
                is_active=not bool(target_user["is_active"]),
                updated_at=now,
            )
        )
    flash("Kullanici durumu guncellendi.", "success")
    return redirect(url_for("user_management"))


@app.route("/")
@login_required
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
@login_required
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
@login_required
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
@login_required
def create_extinguisher():
    if request.method == "POST":
        form = parse_required_form(request.form)
        form["technician_name"] = form.get("technician_name") or current_user_full_name()
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
            return render_template(
                "create_extinguisher.html",
                form=form,
                monthly_control_items=MONTHLY_CONTROL_ITEMS,
                equipment_options=EQUIPMENT_OPTIONS,
                equipment_presets=EQUIPMENT_PRESETS,
            )

        now = datetime.now().isoformat(timespec="seconds")
        public_id = uuid.uuid4().hex[:12]
        inspection_values = build_monthly_inspection_values(request.form)

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
                save_monthly_inspection(
                    connection=connection,
                    extinguisher_id=extinguisher_id,
                    inspection_date=form["last_service_date"],
                    inspector_name=form["technician_name"],
                    notes=form.get("notes"),
                    inspection_values=inspection_values,
                    created_at=now,
                )
        except IntegrityError:
            flash("Bu seri numarasi zaten kayitli.", "error")
            return render_template(
                "create_extinguisher.html",
                form=form,
                monthly_control_items=MONTHLY_CONTROL_ITEMS,
                equipment_options=EQUIPMENT_OPTIONS,
                equipment_presets=EQUIPMENT_PRESETS,
            )

        flash("Tup kaydedildi ve QR olusturuldu.", "success")
        return redirect(url_for("extinguisher_detail", public_id=public_id))

    return render_template(
        "create_extinguisher.html",
        form={"technician_name": current_user_full_name()},
        monthly_control_items=MONTHLY_CONTROL_ITEMS,
        equipment_options=EQUIPMENT_OPTIONS,
        equipment_presets=EQUIPMENT_PRESETS,
    )


@app.route("/extinguishers/<public_id>")
@login_required
def extinguisher_detail(public_id: str):
    extinguisher = get_extinguisher(public_id)
    service_history = fetch_all(
        select(service_logs)
        .where(service_logs.c.extinguisher_id == extinguisher["id"])
        .order_by(desc(service_logs.c.service_date), desc(service_logs.c.id))
    )
    monthly_history = with_monthly_control_labels(
        fetch_all(
            select(monthly_inspections)
            .where(monthly_inspections.c.extinguisher_id == extinguisher["id"])
            .order_by(
                desc(monthly_inspections.c.inspection_date),
                desc(monthly_inspections.c.id),
            )
        )
    )
    return render_template(
        "extinguisher_detail.html",
        extinguisher=extinguisher,
        service_logs=service_history,
        monthly_inspections=monthly_history,
        monthly_control_items=MONTHLY_CONTROL_ITEMS,
        equipment_preset=get_equipment_preset(extinguisher.get("extinguisher_type")),
    )


@app.route("/extinguishers/<public_id>/service", methods=["GET", "POST"])
@login_required
def add_service_log(public_id: str):
    extinguisher = get_extinguisher(public_id)
    if request.method == "POST":
        form = parse_required_form(request.form)
        form["technician_name"] = form.get("technician_name") or current_user_full_name()
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
                monthly_control_items=MONTHLY_CONTROL_ITEMS,
                equipment_options=EQUIPMENT_OPTIONS,
                equipment_presets=EQUIPMENT_PRESETS,
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
                monthly_control_items=MONTHLY_CONTROL_ITEMS,
                equipment_options=EQUIPMENT_OPTIONS,
                equipment_presets=EQUIPMENT_PRESETS,
            )

        now = datetime.now().isoformat(timespec="seconds")
        inspection_values = build_monthly_inspection_values(request.form)
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
            save_monthly_inspection(
                connection=connection,
                extinguisher_id=extinguisher["id"],
                inspection_date=form["service_date"],
                inspector_name=form["technician_name"],
                notes=form.get("notes"),
                inspection_values=inspection_values,
                created_at=now,
            )

        flash("Bakim kaydi eklendi.", "success")
        return redirect(url_for("extinguisher_detail", public_id=public_id))

    return render_template(
        "service_log_form.html",
        extinguisher=extinguisher,
        form={"technician_name": current_user_full_name()},
        monthly_control_items=MONTHLY_CONTROL_ITEMS,
        equipment_options=EQUIPMENT_OPTIONS,
        equipment_presets=EQUIPMENT_PRESETS,
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
    latest_monthly_inspection_raw = fetch_one(
        select(monthly_inspections)
        .where(monthly_inspections.c.extinguisher_id == extinguisher["id"])
        .order_by(
            desc(monthly_inspections.c.inspection_date),
            desc(monthly_inspections.c.id),
        )
        .limit(1)
    )
    monthly_history_raw = fetch_all(
        select(monthly_inspections)
        .where(monthly_inspections.c.extinguisher_id == extinguisher["id"])
        .order_by(
            desc(monthly_inspections.c.inspection_date),
            desc(monthly_inspections.c.id),
        )
    )
    latest_monthly_inspection = (
        with_monthly_control_labels([latest_monthly_inspection_raw])[0]
        if latest_monthly_inspection_raw
        else None
    )
    return render_template(
        "public_detail.html",
        extinguisher=extinguisher,
        latest_log=latest_log,
        latest_monthly_inspection=latest_monthly_inspection,
        equipment_preset=get_equipment_preset(extinguisher.get("extinguisher_type")),
        monthly_table=build_monthly_table(monthly_history_raw),
    )


@app.route("/extinguishers/<public_id>/monthly-inspection", methods=["GET", "POST"])
@login_required
def add_monthly_inspection(public_id: str):
    extinguisher = get_extinguisher(public_id)
    if request.method == "POST":
        form = parse_required_form(request.form)
        form["inspector_name"] = form.get("inspector_name") or current_user_full_name()
        required_fields = {
            "inspection_date": "Kontrol tarihi",
            "inspector_name": "Kontrol eden",
        }
        missing = [label for key, label in required_fields.items() if not form.get(key)]
        if missing:
            flash(f"Eksik alanlar: {', '.join(missing)}", "error")
            return render_template(
                "monthly_inspection_form.html",
                extinguisher=extinguisher,
                monthly_control_items=MONTHLY_CONTROL_ITEMS,
                form=form,
            )

        now = datetime.now().isoformat(timespec="seconds")
        inspection_values = build_monthly_inspection_values(request.form)
        with engine.begin() as connection:
            save_monthly_inspection(
                connection=connection,
                extinguisher_id=extinguisher["id"],
                inspection_date=form["inspection_date"],
                inspector_name=form["inspector_name"],
                notes=form.get("notes"),
                inspection_values=inspection_values,
                created_at=now,
            )

        flash("Aylık kontrol kaydı eklendi.", "success")
        return redirect(url_for("extinguisher_detail", public_id=public_id))

    return render_template(
        "monthly_inspection_form.html",
        extinguisher=extinguisher,
        monthly_control_items=MONTHLY_CONTROL_ITEMS,
        form={"inspector_name": current_user_full_name()},
    )


@app.route("/extinguishers/<public_id>/qr")
@login_required
def extinguisher_qr(public_id: str):
    get_extinguisher(public_id)
    public_url = url_for("public_detail", public_id=public_id, _external=True)
    buffer = build_branded_qr(public_url)
    return send_file(buffer, mimetype="image/png", download_name=f"{public_id}.png")


@app.route("/extinguishers/<public_id>/label")
@login_required
def extinguisher_label(public_id: str):
    get_extinguisher(public_id)
    return render_template("label.html", public_id=public_id)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "5000")), debug=True)
