from __future__ import annotations

import io
import json
import os
import tempfile
import uuid
from datetime import datetime
from pathlib import Path
from functools import wraps

import fitz
import qrcode
from qrcode.constants import ERROR_CORRECT_H
from flask import (
    Flask,
    abort,
    flash,
    redirect,
    render_template,
    render_template_string,
    request,
    send_file,
    session,
    url_for,
)
from openpyxl import Workbook, load_workbook
from PIL import Image, ImageDraw, ImageFont
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader, simpleSplit
from reportlab.platypus import Flowable, Image as PdfImage, Paragraph, SimpleDocTemplate, Spacer, Table as PdfTable, TableStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
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
from xhtml2pdf import pisa


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_SQLITE_URL = f"sqlite:///{(BASE_DIR / 'database.db').as_posix()}"
BRAND_NAME = "Vesta Yangin"
LOGO_PATH = BASE_DIR / "static" / "vesta qr.png"
VESTA_HEADER_LOGO_PATH = BASE_DIR / "static" / "vesta-form-amblem.png"
TSE_HYB_LOGO_PATH = BASE_DIR / "static" / "tse-form-amblem.png"
CONTROL_FORM_TEMPLATE_PATH = BASE_DIR / "assets" / "control-form-template.pdf"
CONTROL_FORM_TEMPLATE_IMAGE_PATH = BASE_DIR / "assets" / "control-form-template.png"
CONTROL_FORM_EXCEL_TEMPLATE_PATH = BASE_DIR / "assets" / "control-form-template.xlsx"
CONTROL_FORM_METHOD_TEXT = "İEKSGŞY, TS ISO 11602-2, TS 862-7 EN 3-7 + A1 ve TS EN 1866-1 standartlarına göre kontrol edilmiştir."
CONTROL_FORM_NOTES = [
    "NOT 1: Yangın söndürücünün kontrolü Madde a) ve b) bendindeki gibi listelenmiş koşullarda, bir eksikliği ortaya çıkardığı zaman, acil düzeltici faaliyet yapılmalıdır.",
    "NOT 2: Yangın söndürücünün kontrolü Madde c), d), e), f) veya g) bendindeki koşullarından herhangi birinde bir eksikliği ortaya çıkardığı zaman, söndürücü uygun bakım işlemlerine VESTA YANGIN tarafından tabi tutulmalıdır.",
    "NOT 3: Madde c), d), e), f) veya g) bendindeki koşullarından herhangi birinde, doldurulmayan tozlu söndürücünün kontrolü bir eksikliği ortaya çıkardığı zaman, bu söndürücü hizmetten kaldırılmalıdır.",
    "NOT 4: Bu muayene raporundaki bulgular muayene tarihindeki işletme koşulları için geçerlidir. Bu rapor 2 nüsha basılmıştır. Muayene raporu VESTA YANGIN onayı olmaksızın kopya edilemez",
]
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
CONTROL_FORM_ITEMS = [
    ("check_a", "a) BELİRLENEN YERE KONULDUĞU"),
    ("check_b", "b) KULLANMA TALİMATLARI GÖRÜNÜMÜ"),
    ("check_c", "c) KULLANMA TALİMATI OKUNURLUĞU"),
    ("check_d", "d) MÜHÜR VE BASINÇ GÖSTERGESİ UYGUNLUĞU"),
    ("check_e", "e) DOLULUK DURUMU (TARTARAK VEYA ELLE)"),
    ("check_f", "f) NOZUL UYGUNLUĞU (PASLANMA VB.)"),
    ("check_g", "g) BASINÇ GÖSTERGESİ İŞLEVSELLİĞİ"),
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
    Column("fire_class", String(255)),
    Column("manufacturer", String(255)),
    Column("hydrostatic_test_date", String(32)),
    Column("company_address", String(255)),
    Column("company_contact", String(255)),
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
    Column("check_a", Boolean, nullable=False, default=False),
    Column("check_b", Boolean, nullable=False, default=False),
    Column("check_c", Boolean, nullable=False, default=False),
    Column("check_d", Boolean, nullable=False, default=False),
    Column("check_e", Boolean, nullable=False, default=False),
    Column("check_f", Boolean, nullable=False, default=False),
    Column("check_g", Boolean, nullable=False, default=False),
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


class RotatedParagraph(Flowable):
    def __init__(self, text: str, style, width: float, height: float) -> None:
        super().__init__()
        self.text = text
        self.style = style
        self.width = width
        self.height = height
        self._paragraph = Paragraph(text, style)

    def wrap(self, availWidth, availHeight):
        return self.width, self.height

    def draw(self):
        self.canv.saveState()
        self.canv.translate(self.width, 0)
        self.canv.rotate(90)
        paragraph_width = self.height - 4
        paragraph_height = self.width - 4
        self._paragraph.wrapOn(self.canv, paragraph_width, paragraph_height)
        self._paragraph.drawOn(self.canv, 2, 2)
        self.canv.restoreState()


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
            extinguisher_columns = {
                row[1] for row in connection.execute(text("PRAGMA table_info(extinguishers)")).fetchall()
            }
            for column_name, column_def in [
                ("fire_class", "TEXT"),
                ("manufacturer", "TEXT"),
                ("hydrostatic_test_date", "TEXT"),
                ("company_address", "TEXT"),
                ("company_contact", "TEXT"),
            ]:
                if column_name not in extinguisher_columns:
                    connection.execute(text(f"ALTER TABLE extinguishers ADD COLUMN {column_name} {column_def}"))

            inspection_columns = {
                row[1] for row in connection.execute(text("PRAGMA table_info(monthly_inspections)")).fetchall()
            }
            for column_name in ["check_a", "check_b", "check_c", "check_d", "check_e", "check_f", "check_g"]:
                if column_name not in inspection_columns:
                    connection.execute(
                        text(f"ALTER TABLE monthly_inspections ADD COLUMN {column_name} BOOLEAN NOT NULL DEFAULT 0")
                    )
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
            for column_name in ["fire_class", "manufacturer", "hydrostatic_test_date", "company_address", "company_contact"]:
                result = connection.execute(
                    text(
                        f"""
                        SELECT column_name
                        FROM information_schema.columns
                        WHERE table_name = 'extinguishers' AND column_name = '{column_name}'
                        """
                    )
                ).fetchone()
                if result is None:
                    connection.execute(text(f"ALTER TABLE extinguishers ADD COLUMN {column_name} TEXT"))
            for column_name in ["check_a", "check_b", "check_c", "check_d", "check_e", "check_f", "check_g"]:
                result = connection.execute(
                    text(
                        f"""
                        SELECT column_name
                        FROM information_schema.columns
                        WHERE table_name = 'monthly_inspections' AND column_name = '{column_name}'
                        """
                    )
                ).fetchone()
                if result is None:
                    connection.execute(
                        text(f"ALTER TABLE monthly_inspections ADD COLUMN {column_name} BOOLEAN NOT NULL DEFAULT FALSE")
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
        Path("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
        Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
        Path("/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf"),
        Path("/usr/share/fonts/dejavu/DejaVuSans.ttf"),
    ]
    for font_path in font_candidates:
        if font_path.exists():
            try:
                return ImageFont.truetype(str(font_path), size)
            except OSError:
                continue
    return ImageFont.load_default()


def resolve_system_font(*candidates: str) -> str:
    for candidate in candidates:
        if Path(candidate).exists():
            return candidate
    raise FileNotFoundError(f"Uygun font bulunamadi: {', '.join(candidates)}")


def build_branded_qr(public_url: str, *, label_mode: bool = False) -> io.BytesIO:
    qr = qrcode.QRCode(
        version=None,
        error_correction=ERROR_CORRECT_H,
        box_size=10,
        border=2 if label_mode else 4,
    )
    qr.add_data(public_url)
    qr.make(fit=True)
    qr_image = qr.make_image(fill_color="black", back_color="white").convert("RGBA")
    qr_width, qr_height = qr_image.size

    if label_mode:
        draw_qr = ImageDraw.Draw(qr_image)
        center_font = load_font(max(104, qr_width // 4))
        v_text = "V"
        v_box = draw_qr.textbbox((0, 0), v_text, font=center_font)
        v_width = v_box[2] - v_box[0]
        v_height = v_box[3] - v_box[1]
        v_pad = 10
        v_bg_w = v_width + (v_pad * 2)
        v_bg_h = v_height + (v_pad * 2)
        v_x = (qr_width - v_bg_w) // 2
        v_y = (qr_height - v_bg_h) // 2
        draw_qr.rounded_rectangle(
            (v_x, v_y, v_x + v_bg_w, v_y + v_bg_h),
            radius=8,
            fill="white",
        )
        draw_qr.text(
            (v_x + v_pad, v_y + v_pad - 2),
            v_text,
            fill="black",
            font=center_font,
        )
    elif LOGO_PATH.exists():
        logo_source = LOGO_PATH
        logo = Image.open(logo_source).convert("RGBA")
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

    if label_mode:
        footer_height = 130
        canvas = Image.new("RGBA", (qr_width, qr_height + footer_height), "white")
        canvas.paste(qr_image, (0, 0))
        draw = ImageDraw.Draw(canvas)
        font = load_font(96)
        brand_text = "Vesta"
        text_bbox = draw.textbbox((0, 0), brand_text, font=font)
        text_width = text_bbox[2] - text_bbox[0]
        text_x = (qr_width - text_width) // 2
        text_y = qr_height + 14
        draw.text((text_x, text_y), brand_text, fill="black", font=font)
    else:
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


def build_control_form_values(form_data) -> dict[str, bool]:
    return {key: form_data.get(key) == "on" for key, _label in CONTROL_FORM_ITEMS}


def save_monthly_inspection(
    connection,
    extinguisher_id: int,
    inspection_date: str,
    inspector_name: str,
    notes: str | None,
    inspection_values: dict[str, bool],
    control_values: dict[str, bool],
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
            **control_values,
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


def build_control_form_pdf(company_name: str, extinguishers_for_company: list[dict], latest_inspections: dict[int, dict]) -> io.BytesIO:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        leftMargin=10 * mm,
        rightMargin=10 * mm,
        topMargin=10 * mm,
        bottomMargin=10 * mm,
    )
    styles = getSampleStyleSheet()
    story = []
    story.append(Paragraph("Taşınabilir Yangın Söndürme Cihazı 6 Aylık Bakım Kontrol Formu", styles["Title"]))
    story.append(Spacer(1, 4 * mm))

    first = extinguishers_for_company[0] if extinguishers_for_company else {}
    general_data = [
        ["FİRMA ADI", company_name, "KONTROL TARİHİ", datetime.now().strftime("%d.%m.%Y")],
        ["MUAYENE ADRESİ", first.get("company_address") or "-", "FİRMA YETKİLİ KİŞİ", first.get("company_contact") or "-"],
        ["PERİYODİK KONTROL PERSONELİ", current_user_full_name() or "-", "AÇIKLAMALAR", "-"],
    ]
    general_table = PdfTable(general_data, colWidths=[32 * mm, 86 * mm, 42 * mm, 90 * mm])
    general_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), colors.whitesmoke),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("BACKGROUND", (0, 0), (0, -1), colors.lightgrey),
                ("BACKGROUND", (2, 0), (2, -1), colors.lightgrey),
            ]
        )
    )
    story.append(general_table)
    story.append(Spacer(1, 6 * mm))

    header = [
        "CİHAZ NO",
        "YSC CİNSİ",
        "YSC SINIFI",
        "SERİ NO / KOD",
        "YSC ÜRETİCİ",
        "DOLUM TARİHİ",
        "HİDROSTATİK TEST TARİHİ",
        "BULUNDUĞU YER",
    ] + [label.split(")")[0] + ")" for _key, label in CONTROL_FORM_ITEMS]
    rows = [header]
    for index, extinguisher in enumerate(extinguishers_for_company, start=1):
        inspection = latest_inspections.get(extinguisher["id"], {})
        row = [
            str(index),
            extinguisher.get("extinguisher_type") or "-",
            extinguisher.get("fire_class") or "-",
            extinguisher.get("serial_number") or "-",
            extinguisher.get("manufacturer") or "-",
            extinguisher.get("last_service_date") or "-",
            extinguisher.get("hydrostatic_test_date") or "-",
            extinguisher.get("location_detail") or "-",
        ]
        for key, _label in CONTROL_FORM_ITEMS:
            value = inspection.get(key)
            row.append("✔" if value is True else "X" if value is False and inspection else "─")
        rows.append(row)

    table = PdfTable(rows, repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e9d6cf")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), 6.5),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]
        )
    )
    story.append(table)
    story.append(Spacer(1, 4 * mm))
    story.append(Paragraph("NOT 1: Yangın söndürücünün kontrolü Madde a) ve b) bendindeki gibi listelenmiş koşullarda, bir eksikliği ortaya çıkardığı zaman, acil düzeltici faaliyet yapılmalıdır.", styles["BodyText"]))
    story.append(Paragraph("NOT 2: Yangın söndürücünün kontrolü Madde c), d), e), f) veya g) bendindeki koşullarından herhangi birinde bir eksikliği ortaya çıkardığı zaman, söndürücü uygun bakım işlemlerine VESTA YANGIN tarafından tabi tutulmalıdır.", styles["BodyText"]))
    doc.build(story)
    buffer.seek(0)
    return buffer


def build_control_form_pdf_from_template(
    company_name: str,
    extinguishers_for_company: list[dict],
    latest_inspections: dict[int, dict],
) -> io.BytesIO:
    if not CONTROL_FORM_TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"Kontrol formu sablonu bulunamadi: {CONTROL_FORM_TEMPLATE_PATH}")

    template = fitz.open(CONTROL_FORM_TEMPLATE_PATH)
    output = fitz.open()

    regular_font = "vesta_regular"
    bold_font = "vesta_bold"
    regular_font_path = resolve_system_font(
        "C:/Windows/Fonts/arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/dejavu/DejaVuSans.ttf",
    )
    bold_font_path = resolve_system_font(
        "C:/Windows/Fonts/arialbd.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf",
    )
    rows_per_page = 15

    def draw_box_text(page, rect, text_value, *, fontname=regular_font, fontsize=7, align=0):
        value = str(text_value or "-").strip() or "-"
        page.insert_textbox(
            fitz.Rect(*rect),
            value,
            fontname=fontname,
            fontsize=fontsize,
            color=(0, 0, 0),
            align=align,
        )

    def draw_page(chunk: list[dict], page_index: int) -> None:
        output.insert_pdf(template, from_page=0, to_page=0)
        page = output[-1]
        page.insert_font(fontname=regular_font, fontfile=regular_font_path)
        page.insert_font(fontname=bold_font, fontfile=bold_font_path)

        first = chunk[0] if chunk else {}
        draw_box_text(page, (110, 78, 608, 92), company_name, fontname=bold_font, fontsize=8)
        draw_box_text(page, (779, 78, 828, 92), datetime.now().strftime("%d.%m.%Y"), fontsize=7)
        draw_box_text(page, (110, 95, 608, 109), first.get("company_address") or "-", fontsize=7)
        draw_box_text(page, (765, 95, 828, 109), first.get("company_contact") or "-", fontsize=7)
        draw_box_text(page, (165, 117, 450, 131), current_user_full_name() or "-", fontsize=7)

        table_top = 375
        row_height = 18.6
        column_edges = [
            (31, 52),
            (52, 96),
            (96, 133),
            (133, 214),
            (214, 278),
            (278, 321),
            (321, 383),
            (383, 524),
            (524, 559),
            (559, 594),
            (594, 628),
            (628, 662),
            (662, 697),
            (697, 731),
            (731, 765),
            (765, 798),
        ]

        for row_number, extinguisher in enumerate(chunk, start=1):
            inspection = latest_inspections.get(extinguisher["id"], {})
            top = table_top + ((row_number - 1) * row_height)
            bottom = top + row_height
            base_values = [
                str((page_index * rows_per_page) + row_number),
                extinguisher.get("extinguisher_type") or "-",
                extinguisher.get("fire_class") or "-",
                extinguisher.get("serial_number") or "-",
                extinguisher.get("manufacturer") or "-",
                extinguisher.get("last_service_date") or "-",
                extinguisher.get("hydrostatic_test_date") or "-",
                extinguisher.get("location_detail") or "-",
            ]

            for index, value in enumerate(base_values):
                x0, x1 = column_edges[index]
                align = 1 if index in {0, 5, 6} else 0
                fontsize = 6 if index in {1, 2, 3, 4, 7} else 7
                draw_box_text(page, (x0 + 2, top + 1, x1 - 2, bottom - 1), value, fontsize=fontsize, align=align)

            for check_index, (key, _label) in enumerate(CONTROL_FORM_ITEMS, start=8):
                x0, x1 = column_edges[check_index]
                symbol = ""
                if inspection:
                    symbol = "V" if inspection.get(key) else "X"
                draw_box_text(page, (x0, top + 1, x1, bottom - 1), symbol, fontname=bold_font, fontsize=8, align=1)

    for page_index, start in enumerate(range(0, len(extinguishers_for_company), rows_per_page)):
        draw_page(extinguishers_for_company[start : start + rows_per_page], page_index)

    if not extinguishers_for_company:
        draw_page([], 0)

    buffer = io.BytesIO(output.tobytes())
    buffer.seek(0)
    output.close()
    template.close()
    return buffer


def build_control_form_pdf_exact(
    company_name: str,
    extinguishers_for_company: list[dict],
    latest_inspections: dict[int, dict],
) -> io.BytesIO:
    if not CONTROL_FORM_TEMPLATE_IMAGE_PATH.exists():
        raise FileNotFoundError(f"Kontrol formu sablon gorseli bulunamadi: {CONTROL_FORM_TEMPLATE_IMAGE_PATH}")

    page_width, page_height = landscape(A4)
    rows_per_page = 15
    background = ImageReader(str(CONTROL_FORM_TEMPLATE_IMAGE_PATH))

    regular_font_path = resolve_system_font(
        "C:/Windows/Fonts/arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/dejavu/DejaVuSans.ttf",
    )
    bold_font_path = resolve_system_font(
        "C:/Windows/Fonts/arialbd.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf",
    )
    if "VestaArial" not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont("VestaArial", regular_font_path))
    if "VestaArialBold" not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont("VestaArialBold", bold_font_path))

    buffer = io.BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=(page_width, page_height))

    def fill_white_box(x0, y0, x1, y1):
        pdf.saveState()
        pdf.setFillColorRGB(1, 1, 1)
        pdf.rect(x0, page_height - y1, x1 - x0, y1 - y0, stroke=0, fill=1)
        pdf.restoreState()

    def draw_text_box(x0, y0, x1, y1, value, *, font="VestaArial", size=7, align="left"):
        text = str(value or "-").strip() or "-"
        box_width = x1 - x0
        lines = simpleSplit(text, font, size, max(box_width - 4, 10))
        max_lines = max(int((y1 - y0) / (size + 1)), 1)
        lines = lines[:max_lines]
        current_y = page_height - y0 - size - 1
        for line in lines:
            line_width = pdfmetrics.stringWidth(line, font, size)
            if align == "center":
                text_x = x0 + max((box_width - line_width) / 2, 0)
            elif align == "right":
                text_x = x1 - line_width - 2
            else:
                text_x = x0 + 2
            pdf.setFont(font, size)
            pdf.drawString(text_x, current_y, line)
            current_y -= size + 1

    def draw_page(chunk: list[dict], page_index: int) -> None:
        pdf.drawImage(background, 0, 0, width=page_width, height=page_height)
        first = chunk[0] if chunk else {}

        fill_white_box(112, 129, 610, 144)
        fill_white_box(748, 129, 812, 144)
        fill_white_box(112, 147, 610, 162)
        fill_white_box(748, 147, 812, 162)

        draw_text_box(112, 129, 610, 144, company_name, font="VestaArialBold", size=7.0)
        draw_text_box(748, 129, 812, 144, datetime.now().strftime("%d.%m.%Y"), size=5.5, align="center")
        draw_text_box(112, 147, 610, 162, first.get("company_address") or "-", size=6.2)
        draw_text_box(748, 147, 812, 162, first.get("company_contact") or "-", size=5.4, align="center")

        table_top = 264
        row_height = 18.55
        column_edges = [
            (32, 52),
            (52, 96),
            (96, 133),
            (133, 214),
            (214, 278),
            (278, 321),
            (321, 383),
            (383, 524),
            (524, 559),
            (559, 594),
            (594, 628),
            (628, 662),
            (662, 697),
            (697, 731),
            (731, 765),
            (765, 798),
        ]

        for row_number, extinguisher in enumerate(chunk, start=1):
            inspection = latest_inspections.get(extinguisher["id"], {})
            top = table_top + ((row_number - 1) * row_height)
            bottom = top + row_height
            values = [
                str((page_index * rows_per_page) + row_number),
                pdf_equipment_label(extinguisher.get("extinguisher_type")),
                extinguisher.get("fire_class") or "-",
                extinguisher.get("serial_number") or "-",
                extinguisher.get("manufacturer") or "-",
                extinguisher.get("last_service_date") or "-",
                extinguisher.get("hydrostatic_test_date") or "-",
                extinguisher.get("location_detail") or "-",
            ]

            for index, value in enumerate(values):
                x0, x1 = column_edges[index]
                align = "center" if index in {0, 5, 6} else "left"
                size = 5.4 if index in {1, 2, 3, 4, 7} else 6.3
                draw_text_box(x0, top, x1, bottom, value, size=size, align=align)

            for check_index, (key, _label) in enumerate(CONTROL_FORM_ITEMS, start=8):
                x0, x1 = column_edges[check_index]
                symbol = ""
                if inspection:
                    symbol = "V" if inspection.get(key) else "X"
                draw_text_box(x0, top, x1, bottom, symbol, font="VestaArialBold", size=8, align="center")

        pdf.showPage()

    for page_index, start in enumerate(range(0, len(extinguishers_for_company), rows_per_page)):
        draw_page(extinguishers_for_company[start : start + rows_per_page], page_index)

    if not extinguishers_for_company:
        draw_page([], 0)

    pdf.save()
    buffer.seek(0)
    return buffer


def build_company_filename(company_name: str) -> str:
    safe = "".join(char if char.isalnum() else "-" for char in company_name.lower()).strip("-")
    return safe or "kontrol-formu"


def pdf_equipment_label(extinguisher_type: str | None) -> str:
    mapping = {
        "Kuru Kimyevi Toz": "Kuru Kimyevi Toz",
        "CO2": "CO2",
        "Kopuk": "Köpük",
    }
    return mapping.get(extinguisher_type or "", extinguisher_type or "-")


def build_control_form_document_data(public_id: str) -> dict:
    extinguisher = get_extinguisher(public_id)
    company_name = extinguisher["company_name"]
    company_extinguishers = fetch_all(
        select(extinguishers)
        .where(extinguishers.c.company_name == company_name)
        .order_by(extinguishers.c.location_detail, extinguishers.c.serial_number)
    )
    extinguisher_ids = [row["id"] for row in company_extinguishers]
    latest_inspections: dict[int, dict] = {}
    if extinguisher_ids:
        inspection_rows = fetch_all(
            select(monthly_inspections)
            .where(monthly_inspections.c.extinguisher_id.in_(extinguisher_ids))
            .order_by(
                monthly_inspections.c.extinguisher_id,
                desc(monthly_inspections.c.inspection_date),
                desc(monthly_inspections.c.id),
            )
        )
        for row in inspection_rows:
            latest_inspections.setdefault(row["extinguisher_id"], row)

    control_rows = []
    for index, row in enumerate(company_extinguishers, start=1):
        inspection = latest_inspections.get(row["id"], {})
        control_rows.append(
            {
                "device_no": index,
                "extinguisher_type": pdf_equipment_label(row.get("extinguisher_type")),
                "fire_class": row.get("fire_class") or "-",
                "serial_number": row.get("serial_number") or "-",
                "manufacturer": row.get("manufacturer") or "-",
                "service_date": row.get("last_service_date") or "-",
                "hydrostatic_test_date": row.get("hydrostatic_test_date") or "-",
                "location_detail": row.get("location_detail") or "-",
                "checks": [
                    "V" if inspection.get(key) else "X" if inspection else "-"
                    for key, _label in CONTROL_FORM_ITEMS
                ],
            }
        )

    return {
        "company_name": company_name,
        "company_address": extinguisher.get("company_address") or "-",
        "company_contact": extinguisher.get("company_contact") or "-",
        "control_date": datetime.now().strftime("%d.%m.%Y"),
        "inspector_name": current_user_full_name() or "-",
        "method_text": CONTROL_FORM_METHOD_TEXT,
        "rows": control_rows,
        "check_headers": [label for _key, label in CONTROL_FORM_ITEMS],
        "notes": CONTROL_FORM_NOTES,
        "public_id": public_id,
    }


def build_control_form_pdf_reportlab(document_data: dict) -> io.BytesIO:
    main_col_widths = [10 * mm, 14 * mm, 14 * mm, 19 * mm, 16 * mm, 14 * mm, 16 * mm, 24 * mm, 11 * mm, 11 * mm, 11 * mm, 11 * mm, 11 * mm, 11 * mm, 11 * mm]
    form_total_width = sum(main_col_widths)

    regular_font_path = resolve_system_font(
        "C:/Windows/Fonts/arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/dejavu/DejaVuSans.ttf",
    )
    bold_font_path = resolve_system_font(
        "C:/Windows/Fonts/arialbd.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf",
    )

    if "VestaPDF" not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont("VestaPDF", regular_font_path))
    if "VestaPDFBold" not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont("VestaPDFBold", bold_font_path))

    styles = getSampleStyleSheet()
    body_style = styles["BodyText"].clone("vesta_body")
    body_style.fontName = "VestaPDF"
    body_style.fontSize = 7
    body_style.leading = 8

    small_style = styles["BodyText"].clone("vesta_small")
    small_style.fontName = "VestaPDF"
    small_style.fontSize = 6
    small_style.leading = 7

    tiny_style = styles["BodyText"].clone("vesta_tiny")
    tiny_style.fontName = "VestaPDF"
    tiny_style.fontSize = 5
    tiny_style.leading = 6

    tiny_bold_style = styles["BodyText"].clone("vesta_tiny_bold")
    tiny_bold_style.fontName = "VestaPDFBold"
    tiny_bold_style.fontSize = 5
    tiny_bold_style.leading = 6

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        leftMargin=6 * mm,
        rightMargin=6 * mm,
        topMargin=6 * mm,
        bottomMargin=6 * mm,
    )

    story = []

    logo_cells = []
    if TSE_HYB_LOGO_PATH.exists():
        logo_cells.append(PdfImage(str(TSE_HYB_LOGO_PATH), width=17 * mm, height=17 * mm))
    else:
        logo_cells.append(Paragraph("<b>TSE-HYB</b>", body_style))
    if VESTA_HEADER_LOGO_PATH.exists():
        logo_cells.append(PdfImage(str(VESTA_HEADER_LOGO_PATH), width=13 * mm, height=16 * mm))
    else:
        logo_cells.append(Paragraph("<b>VESTA</b>", body_style))

    logo_table = PdfTable([logo_cells], colWidths=[20 * mm, 20 * mm], rowHeights=[18 * mm])
    logo_table.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ]
        )
    )

    header_table = PdfTable(
        [
            [
                logo_table,
                Paragraph("<b>Onay: Şirket Müdürü<br/>Mustafa KİLİÇ</b>", body_style),
                Paragraph("<b>Bölüm</b>", body_style),
                Paragraph("<b>Revizyon No</b>", body_style),
                Paragraph("<b>Revizyon Tarihi</b>", body_style),
                Paragraph("<b>Sayfa</b>", body_style),
            ],
            [
                "",
                "",
                "F-06",
                "00",
                "01.03.2026",
                "1",
            ],
            [
                Paragraph("<b>FORM</b>", body_style),
                Paragraph("<b>Konu: Taşınabilir Yangın Söndürme Cihazı Kontrol Formu</b>", body_style),
                Paragraph("<b>Hazırlayan: Kalite Temsilcisi</b>", body_style),
                "",
                "",
                "",
            ],
        ],
        colWidths=[40 * mm, 60 * mm, 16 * mm, 22 * mm, 32 * mm, 34 * mm],
        hAlign="CENTER",
    )
    header_table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
                ("SPAN", (0, 0), (0, 1)),
                ("SPAN", (1, 0), (1, 1)),
                ("SPAN", (1, 2), (2, 2)),
                ("SPAN", (3, 2), (5, 2)),
                ("FONTNAME", (0, 0), (-1, -1), "VestaPDF"),
                ("FONTNAME", (0, 0), (-1, 0), "VestaPDFBold"),
                ("FONTNAME", (0, 2), (-1, 2), "VestaPDFBold"),
                ("FONTSIZE", (0, 0), (-1, -1), 7),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (2, 0), (-1, 1), "CENTER"),
            ]
        )
    )
    story.extend([header_table, Spacer(1, 1 * mm)])

    compact_small_style = styles["BodyText"].clone("vesta_compact_small")
    compact_small_style.fontName = "VestaPDF"
    compact_small_style.fontSize = 5
    compact_small_style.leading = 6
    compact_small_style.spaceBefore = 0
    compact_small_style.spaceAfter = 0

    compact_cell_style = styles["BodyText"].clone("vesta_compact_cell")
    compact_cell_style.fontName = "VestaPDF"
    compact_cell_style.fontSize = 4.4
    compact_cell_style.leading = 4.8
    compact_cell_style.spaceBefore = 0
    compact_cell_style.spaceAfter = 0
    compact_cell_style.alignment = 1

    method_paragraph = Paragraph(document_data["method_text"], compact_small_style)
    address_cell_style = styles["BodyText"].clone("vesta_address_cell")
    address_cell_style.fontName = "VestaPDF"
    address_cell_style.fontSize = 4.6
    address_cell_style.leading = 5
    address_cell_style.spaceBefore = 0
    address_cell_style.spaceAfter = 0

    info_table = PdfTable(
        [
            [Paragraph("<b>GENEL BİLGİLER</b>", body_style), "", "", "", "", ""],
            ["FIRMA ADI", document_data["company_name"], "", "", "KONTROL TARIHI", document_data["control_date"]],
            ["MUAYENE ADRESI", Paragraph(document_data["company_address"], address_cell_style), "", "", "FIRMA YETKILI KISI", document_data["company_contact"]],
            ["PERIYODIK KONTROL METODU", method_paragraph, "", "", "", ""],
        ],
        colWidths=[56 * mm, 58 * mm, 2 * mm, 2 * mm, 34 * mm, 52 * mm],
        hAlign="CENTER",
        rowHeights=[7 * mm, 7 * mm, 7 * mm, 9 * mm],
    )
    info_table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
                ("SPAN", (0, 0), (-1, 0)),
                ("SPAN", (1, 1), (3, 1)),
                ("SPAN", (1, 2), (3, 2)),
                ("SPAN", (1, 3), (-1, 3)),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#f1f1f1")),
                ("FONTNAME", (0, 0), (-1, -1), "VestaPDF"),
                ("FONTNAME", (0, 0), (-1, 0), "VestaPDFBold"),
                ("FONTNAME", (0, 1), (0, -1), "VestaPDFBold"),
                ("FONTNAME", (4, 1), (4, 2), "VestaPDFBold"),
                ("FONTSIZE", (0, 0), (-1, -1), 7),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                ("TOPPADDING", (0, 0), (-1, -1), 0.8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0.8),
                ("LEFTPADDING", (0, 0), (-1, -1), 2),
                ("RIGHTPADDING", (0, 0), (-1, -1), 2),
                ("LEFTPADDING", (1, 3), (1, 3), 10),
            ]
        )
    )
    story.extend([info_table, Spacer(1, 1 * mm)])

    header_row_1 = [
        Paragraph("<b>YANGIN SÖNDÜRME CİHAZI (YSC) BİLGİLERİ</b>", body_style),
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        Paragraph(
            "<b>TESPİT VE DEĞERLENDİRME</b><br/><font size='6'>(V: UYGUN, X: UYGUN DEĞİL, -: UYGULAMA YOK)</font>",
            tiny_bold_style,
        ),
        "",
        "",
        "",
        "",
        "",
        "",
    ]
    rotated_width = 11 * mm
    rotated_height = 30 * mm

    header_row_2 = [
        Paragraph("<b>CİHAZ</b><br/><b>NO</b>", tiny_bold_style),
        Paragraph("<b>YSC CİNSİ</b>", tiny_bold_style),
        Paragraph("<b>YSC SINIFI</b>", tiny_bold_style),
        Paragraph("<b>SERİ NO / KOD</b>", tiny_bold_style),
        Paragraph("<b>YSC ÜRETİCİ</b>", tiny_bold_style),
        Paragraph("<b>DOLUM TARİHİ</b>", tiny_bold_style),
        Paragraph("<b>HİDROSTATİK</b><br/><b>TEST TARİHİ</b>", tiny_bold_style),
        Paragraph("<b>BULUNDUĞU YER</b>", tiny_bold_style),
        *[
            RotatedParagraph(f"<b>{label}</b>", tiny_style, rotated_width, rotated_height)
            for label in document_data["check_headers"]
        ],
    ]

    data_rows = [header_row_1, header_row_2]
    for row in document_data["rows"]:
        extinguisher_type = str(row["extinguisher_type"]).replace("Kuru Kimyevi Toz", "Kuru<br/>Kimyevi Toz")
        data_rows.append(
            [
                row["device_no"],
                Paragraph(extinguisher_type, compact_cell_style),
                Paragraph(str(row["fire_class"]), compact_cell_style),
                Paragraph(str(row["serial_number"]), compact_cell_style),
                Paragraph(str(row["manufacturer"]), compact_cell_style),
                Paragraph(str(row["service_date"]), compact_cell_style),
                Paragraph(str(row["hydrostatic_test_date"]), compact_cell_style),
                Paragraph(str(row["location_detail"]), compact_cell_style),
                *row["checks"],
            ]
        )

    while len(data_rows) < 17:
        data_rows.append([""] * 15)

    main_table = PdfTable(
        data_rows,
        repeatRows=2,
        colWidths=main_col_widths,
        rowHeights=[6 * mm, 30 * mm] + [4.8 * mm] * (len(data_rows) - 2),
        hAlign="CENTER",
    )
    main_table.setStyle(
        TableStyle(
            [
                ("SPAN", (0, 0), (7, 0)),
                ("SPAN", (8, 0), (14, 0)),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                ("BACKGROUND", (0, 0), (-1, 1), colors.HexColor("#f7f7f7")),
                ("FONTNAME", (0, 0), (-1, 1), "VestaPDFBold"),
                ("FONTNAME", (0, 2), (-1, -1), "VestaPDF"),
                ("FONTSIZE", (0, 0), (-1, -1), 5),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]
        )
    )
    story.extend([main_table, Spacer(1, 1 * mm)])

    note_style = styles["BodyText"].clone("vesta_note")
    note_style.fontName = "VestaPDF"
    note_style.fontSize = 5
    note_style.leading = 6
    note_style.leftIndent = 0
    note_style.spaceBefore = 0
    note_style.spaceAfter = 0

    notes_table = PdfTable(
        [[Paragraph(note, note_style)] for note in document_data["notes"]],
        colWidths=[form_total_width],
        hAlign="CENTER",
    )
    notes_table.setStyle(
        TableStyle(
            [
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ]
        )
    )
    story.append(notes_table)

    doc.build(story)
    buffer.seek(0)
    return buffer


def build_control_form_excel_pdf(
    company_name: str,
    extinguishers_for_company: list[dict],
    latest_inspections: dict[int, dict],
) -> io.BytesIO | None:
    if os.name != "nt" or not CONTROL_FORM_EXCEL_TEMPLATE_PATH.exists():
        return None

    try:
        import pythoncom
        import win32com.client  # type: ignore[import-not-found]
    except ImportError:
        return None

    workbook = load_workbook(CONTROL_FORM_EXCEL_TEMPLATE_PATH)
    worksheet = workbook["KONTROL FORMU"]
    first = extinguishers_for_company[0] if extinguishers_for_company else {}

    worksheet["C4"] = company_name
    worksheet["N4"] = datetime.now().strftime("%d.%m.%Y")
    worksheet["C5"] = first.get("company_address") or "-"
    worksheet["N5"] = first.get("company_contact") or "-"

    start_row = 10
    max_rows = 14
    for offset in range(max_rows):
        row_no = start_row + offset
        extinguisher = extinguishers_for_company[offset] if offset < len(extinguishers_for_company) else None
        if extinguisher is None:
            for column in "ABCDEFGHIJKLMNO":
                worksheet[f"{column}{row_no}"] = None
            continue

        inspection = latest_inspections.get(extinguisher["id"], {})
        worksheet[f"A{row_no}"] = offset + 1
        worksheet[f"B{row_no}"] = pdf_equipment_label(extinguisher.get("extinguisher_type"))
        worksheet[f"C{row_no}"] = extinguisher.get("fire_class") or "-"
        worksheet[f"D{row_no}"] = extinguisher.get("serial_number") or "-"
        worksheet[f"E{row_no}"] = extinguisher.get("manufacturer") or "-"
        worksheet[f"F{row_no}"] = extinguisher.get("last_service_date") or "-"
        worksheet[f"G{row_no}"] = extinguisher.get("hydrostatic_test_date") or "-"
        worksheet[f"H{row_no}"] = extinguisher.get("location_detail") or "-"
        for idx, column in enumerate("IJKLMNO"):
            key = CONTROL_FORM_ITEMS[idx][0]
            if inspection:
                worksheet[f"{column}{row_no}"] = "✔" if inspection.get(key) else "X"
            else:
                worksheet[f"{column}{row_no}"] = "-"

    temp_dir = Path(tempfile.mkdtemp(prefix="vesta-control-form-"))
    xlsx_path = temp_dir / "control-form.xlsx"
    pdf_path = temp_dir / "control-form.pdf"
    workbook.save(xlsx_path)
    workbook.close()

    pythoncom.CoInitialize()
    excel = None
    excel_workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel_workbook = excel.Workbooks.Open(str(xlsx_path))
        worksheet_com = excel_workbook.Worksheets("KONTROL FORMU")
        worksheet_com.ExportAsFixedFormat(0, str(pdf_path))
        excel_workbook.Close(False)
        excel.Quit()
        pdf_bytes = pdf_path.read_bytes()
    finally:
        if excel_workbook is not None:
            try:
                excel_workbook.Close(False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

    buffer = io.BytesIO(pdf_bytes)
    buffer.seek(0)
    return buffer


def build_pdf_from_html(html: str) -> io.BytesIO:
    buffer = io.BytesIO()
    result = pisa.CreatePDF(src=html, dest=buffer, encoding="utf-8")
    if result.err:
        raise RuntimeError("HTML PDF olusturma basarisiz oldu.")
    buffer.seek(0)
    return buffer


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
            "company_address": "Muayene adresi",
            "company_contact": "Firma yetkili kisi",
            "location_detail": "Firma ici konum",
            "weight_kg": "Kg bilgisi",
            "extinguisher_type": "Tup tipi",
            "fire_class": "YSC sinifi",
            "manufacturer": "YSC uretici",
            "last_service_date": "Son bakim tarihi",
            "hydrostatic_test_date": "Hidrostatik test tarihi",
            "next_service_date": "Sonraki bakim tarihi",
            "technician_name": "Teknisyen",
            "operation_summary": "Yapilan islem",
        }
        missing = [label for key, label in required_fields.items() if not form.get(key)]
        if missing:
            flash(f"Eksik alanlar: {', '.join(missing)}", "error")
            return render_template(
                "create_extinguisher.html",
                form=form,
                monthly_control_items=MONTHLY_CONTROL_ITEMS,
                equipment_options=EQUIPMENT_OPTIONS,
                equipment_presets=EQUIPMENT_PRESETS,
            )

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
        control_values = build_control_form_values(request.form)

        try:
            with engine.begin() as connection:
                result = connection.execute(
                    insert(extinguishers).values(
                        public_id=public_id,
                        serial_number=form["serial_number"],
                        company_name=form["company_name"],
                        company_address=form["company_address"],
                        company_contact=form["company_contact"],
                        location_detail=form["location_detail"],
                        weight_kg=weight_kg,
                        extinguisher_type=form["extinguisher_type"],
                        fire_class=form["fire_class"],
                        manufacturer=form["manufacturer"],
                        hydrostatic_test_date=form["hydrostatic_test_date"],
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
                    control_values=control_values,
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


@app.route("/extinguishers/<public_id>/control-form.pdf")
@login_required
def control_form_pdf(public_id: str):
    document_data = build_control_form_document_data(public_id)
    pdf_buffer = build_control_form_pdf_reportlab(document_data)
    filename = f"{build_company_filename(document_data['company_name'])}-kontrol-formu.pdf"
    return send_file(
        pdf_buffer,
        mimetype="application/pdf",
        as_attachment=True,
        download_name=filename,
    )


@app.route("/public/<public_id>/control-form.pdf")
def public_control_form_pdf(public_id: str):
    document_data = build_control_form_document_data(public_id)
    pdf_buffer = build_control_form_pdf_reportlab(document_data)
    filename = f"{build_company_filename(document_data['company_name'])}-kontrol-formu.pdf"
    return send_file(
        pdf_buffer,
        mimetype="application/pdf",
        as_attachment=True,
        download_name=filename,
    )


@app.route("/extinguishers/<public_id>/control-form")
@login_required
def control_form_preview(public_id: str):
    document_data = build_control_form_document_data(public_id)
    return render_template("control_form_preview.html", **document_data)


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
            "company_name": "Firma adi",
            "company_address": "Muayene adresi",
            "company_contact": "Firma yetkili kisi",
            "location_detail": "Firma ici konum",
            "extinguisher_type": "Tup tipi",
            "fire_class": "YSC sinifi",
            "manufacturer": "YSC uretici",
            "hydrostatic_test_date": "Hidrostatik test tarihi",
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
        control_values = build_control_form_values(request.form)
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
                    company_address=form.get("company_address") or extinguisher.get("company_address"),
                    company_contact=form.get("company_contact") or extinguisher.get("company_contact"),
                    location_detail=form.get("location_detail")
                    or extinguisher["location_detail"],
                    weight_kg=weight_kg,
                    extinguisher_type=form.get("extinguisher_type")
                    or extinguisher["extinguisher_type"],
                    fire_class=form.get("fire_class") or extinguisher.get("fire_class"),
                    manufacturer=form.get("manufacturer") or extinguisher.get("manufacturer"),
                    hydrostatic_test_date=form.get("hydrostatic_test_date") or extinguisher.get("hydrostatic_test_date"),
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
                control_values=control_values,
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
                control_form_items=CONTROL_FORM_ITEMS,
                form=form,
            )

        now = datetime.now().isoformat(timespec="seconds")
        inspection_values = build_monthly_inspection_values(request.form)
        control_values = build_control_form_values(request.form)
        with engine.begin() as connection:
            save_monthly_inspection(
                connection=connection,
                extinguisher_id=extinguisher["id"],
                inspection_date=form["inspection_date"],
                inspector_name=form["inspector_name"],
                notes=form.get("notes"),
                inspection_values=inspection_values,
                control_values=control_values,
                created_at=now,
            )

        flash("Aylık kontrol kaydı eklendi.", "success")
        return redirect(url_for("extinguisher_detail", public_id=public_id))

    return render_template(
        "monthly_inspection_form.html",
        extinguisher=extinguisher,
        monthly_control_items=MONTHLY_CONTROL_ITEMS,
        control_form_items=CONTROL_FORM_ITEMS,
        form={"inspector_name": current_user_full_name()},
    )


@app.route("/extinguishers/<public_id>/qr")
@login_required
def extinguisher_qr(public_id: str):
    get_extinguisher(public_id)
    public_url = url_for("public_detail", public_id=public_id, _external=True)
    buffer = build_branded_qr(public_url)
    return send_file(buffer, mimetype="image/png", download_name=f"{public_id}.png")


@app.route("/extinguishers/<public_id>/label-qr")
@login_required
def extinguisher_label_qr(public_id: str):
    get_extinguisher(public_id)
    public_url = url_for("public_detail", public_id=public_id, _external=True)
    buffer = build_branded_qr(public_url, label_mode=True)
    return send_file(buffer, mimetype="image/png", download_name=f"{public_id}-label.png")


@app.route("/extinguishers/<public_id>/label")
@login_required
def extinguisher_label(public_id: str):
    get_extinguisher(public_id)
    return render_template("label.html", public_id=public_id)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "5000")), debug=True)
