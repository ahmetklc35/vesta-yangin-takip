from __future__ import annotations

import io
import json
import os
import tempfile
import re
import unicodedata
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
ASSET_CATEGORIES = [
    {
        "slug": "yangin-sondurme-cihazi",
        "label": "Yangin Sondurme Cihazi",
        "description": "Yangin tup ve sondurme cihazlarinin tamamini goruntule.",
    },
    {
        "slug": "yangin-elbisesi",
        "label": "Yangin Elbisesi",
        "description": "Yangin elbisesi kontrollerini ve kayitlarini listele.",
    },
    {
        "slug": "yangin-bareti",
        "label": "Yangin Bareti",
        "description": "Yangin baretlerine ait kontrolleri ve kayitlari goruntule.",
    },
    {
        "slug": "yangin-baltasi",
        "label": "Yangin Baltasi",
        "description": "Yangin baltasi ekipmanlarini listele.",
    },
    {
        "slug": "scba",
        "label": "SCBA",
        "description": "Bagimsiz solunum cihazi kayitlarini goruntule.",
    },
    {
        "slug": "eebd",
        "label": "EEBD",
        "description": "Acil kacis solunum setlerini goruntule.",
    },
    {
        "slug": "hava-tupu",
        "label": "Hava Tupu",
        "description": "Basincli hava solunum tupu kayitlarini goruntule.",
    },
    {
        "slug": "yangin-sondurme-dolabi",
        "label": "Yangin Sondurme Dolabi",
        "description": "Yangin dolabi periyodik kontrol kayitlarini listele.",
    },
    {
        "slug": "kopuklu-yangin-sondurme-dolabi",
        "label": "Kopuklu Yangin Sondurme Dolabi",
        "description": "Kopuklu dolap kontrol kayitlarini listele.",
    },
    {
        "slug": "yangin-hidranti",
        "label": "Yangin Hidranti",
        "description": "Yangin hidrant periyodik kontrol kayitlarini goruntule.",
    },
    {
        "slug": "elektrik-ic-tesisati",
        "label": "Elektrik Ic Tesisati",
        "description": "Elektrik ic tesisati periyodik kontrol raporlarini goruntule.",
    },
]
DEFAULT_ASSET_CATEGORY = ASSET_CATEGORIES[0]["label"]
ASSET_CATEGORY_BY_SLUG = {item["slug"]: item for item in ASSET_CATEGORIES}
ASSET_CATEGORY_BY_LABEL = {item["label"]: item for item in ASSET_CATEGORIES}
REGISTRATION_GROUPS = [
    {
        "slug": "yangin-sondurme-cihazi",
        "label": "Yangin Sondurme Cihazi",
        "description": "Mevcut calisan YSC kayit ekranini kullanir.",
        "status": "active",
    },
    {
        "slug": "yangin-elbisesi",
        "label": "Yangin Elbisesi",
        "description": "Yangin elbisesi kontrol formuna uygun yeni kayit akisi.",
        "status": "active",
    },
    {
        "slug": "yangin-bareti",
        "label": "Yangin Bareti",
        "description": "Yangin bareti kontrol formuna uygun kayit akisi.",
        "status": "active",
    },
    {
        "slug": "yangin-baltasi",
        "label": "Yangin Baltasi",
        "description": "Yangin baltasi kontrol formuna uygun kayit akisi.",
        "status": "active",
    },
    {
        "slug": "scba",
        "label": "SCBA",
        "description": "Bagimsiz solunum cihazi kontrol formuna uygun kayit akisi.",
        "status": "active",
    },
    {
        "slug": "eebd",
        "label": "EEBD",
        "description": "Acil kacis seti kontrol formuna uygun kayit akisi.",
        "status": "active",
    },
    {
        "slug": "hava-tupu",
        "label": "Hava Tupu",
        "description": "Basincli hava solunum tupu kontrol formuna uygun kayit akisi.",
        "status": "active",
    },
    {
        "slug": "yangin-sondurme-dolabi",
        "label": "Yangin Sondurme Dolabi",
        "description": "Yangin sondurme dolabi periyodik kontrol kayit akisi.",
        "status": "active",
    },
    {
        "slug": "kopuklu-yangin-sondurme-dolabi",
        "label": "Kopuklu Yangin Sondurme Dolabi",
        "description": "Kopuklu yangin sondurme dolabi kontrol kayit akisi.",
        "status": "active",
    },
    {
        "slug": "yangin-hidranti",
        "label": "Yangin Hidranti",
        "description": "Yangin hidranti periyodik kontrol kayit akisi.",
        "status": "active",
    },
    {
        "slug": "elektrik-ic-tesisati",
        "label": "Elektrik Ic Tesisati",
        "description": "Elektrik ic tesisati gozle kontrol ve fonksiyon testleri icin trial kayit akisi.",
        "status": "active",
    },
]
REGISTRATION_GROUP_BY_SLUG = {item["slug"]: item for item in REGISTRATION_GROUPS}
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
ELECTRICAL_NOTE_SECTIONS = [
    (
        "Firma Bilgileri",
        [
            "ISG-KATIP Sozlesme ID",
            "SGK Sicil Numarasi",
            "Periyodik Kontrol Metodu ve Kapsami",
        ],
    ),
    (
        "Detay Bilgiler",
        [
            "Tesise ait proje var mi",
            "Tek hat semasi var mi",
            "Kontrol nedeni",
            "Topraklayici tipi",
            "Yapi cinsi",
            "Son kontrol tarihi",
            "Faz iletkenlerinin sayisi ve tipi",
            "Temel topraklama direnci",
            "Ilave topraklama elektrotu detaylari",
            "Sistem topraklama iletkeni ve kesiti",
            "Ana espotansiyel iletkeni ve kesiti",
            "Besleme kaynagi karakteristikleri",
            "Ana RCD anma akimi",
            "Ana kesici karakteristikleri",
            "Ana RCD test akimi ve suresi",
        ],
    ),
    (
        "Tespit Edilen Bilgiler",
        [
            "Tesisatta kapsamli degisiklik var mi (>%20)",
            "Asiri gerilim koruma cihazlari kullanilmis mi",
            "Dogrudan dokunmaya karsi koruma onlemleri",
            "Bir onceki periyodik kontrol etiketi var mi",
            "Tespit edilen bilgiler",
        ],
    ),
    (
        "Termal Kamera ve Olcum Aletleri",
        [
            "Termal Kamera 1",
            "Termal Kamera 2",
            "Olcum Aleti 1",
            "Olcum Aleti 2",
            "Termal fotograf tarihi",
            "Termal fotograf no",
            "Kontak gevsakligi isinmasi",
            "Asiri yuk isinmasi",
        ],
    ),
    (
        "Test ve Sonuc",
        [
            "Kontrol Kriterleri ve Testler",
            "Olcum ve Dogrulama Metodu",
            "6.1 Notlari",
            "6.2 Notlari",
            "6.3 Notlari",
            "Kusur Aciklamalari",
            "Ekipman Fotograflari",
            "Genel Notlar",
            "Sonuc ve Kanaat",
            "Yetkili Kisi",
            "Nusha Sayisi",
        ],
    ),
]
FIRE_SUIT_CONTROL_ITEMS = [
    ("item_1", "17. Y .1001.A.1 Gorsel Kumas Kontrol yapildi"),
    ("item_2", "17. Y .1001.A.2 Fonksiyonel Fermuar, Cirt cirtlar ve Dugmeler kontrol edildi"),
    ("item_3", "17. Y .1001.A.3 Yansitici Bantlar Kontrol Edildi"),
    ("item_4", "17. Y .1001.A.4 Elbiselerin Temizligi Kontrol Edildi"),
    ("item_5", "17. Y .1001.A.5 Ic Astarin Durumu, Yalitim Ozelligi ve Dikisler Kontrol Edildi"),
    ("item_6", "17. Y .1001.A.6 Servis Etiketi Ekipmana Yapistirildi"),
]
HELMET_CONTROL_ITEMS = [
    ("item_1", "17.B.1001.A.1 Dis kisimda gorsel kontrol yapildi"),
    ("item_2", "17.B.1001.A.2 Ic kisimda gorsel kontrol yapildi"),
    ("item_3", "17.B.1001.A.3 Ayar mekanizmasinda kontrol yapildi"),
    ("item_4", "17.B.1001.A.4 Vizor ve goz korumasinda kontrol yapildi"),
    ("item_5", "17.B.1001.A.5 Boyun koruyucu kontrolu yapildi"),
    ("item_6", "17.B.1001.A.6 Servis etiketi ekipmana yapistirildi"),
]
AXE_CONTROL_ITEMS = [
    ("item_1", "17.YB.1001.A.1 Baltanin ahsap veya yalitkan sapi kontrol edildi"),
    ("item_2", "17.YB.1001.A.2 Metal kismi kontrol edildi"),
    ("item_3", "17.YB.1001.A.3 Agiz kismi kontrol edildi"),
    ("item_4", "17.YB.1001.A.4 Bulundugu yerde ulasilabilir durumda mi"),
    ("item_5", "17.YB.1001.A.5 Servis etiketi ekipmana yapistirildi"),
]
SCBA_CONTROL_ITEMS = [
    ("item_1", "17.S.1001.A.1 Yuz maskesi kontrol edildi"),
    ("item_2", "17.S.1001.A.2 Solunum valfi kontrol edildi"),
    ("item_3", "17.S.1001.A.3 Regulator unitesi kontrol edildi"),
    ("item_4", "17.S.1001.A.4 Kemer kontrol edildi"),
    ("item_5", "17.S.1001.A.5 Silindir kontrol edildi"),
    ("item_6", "17.S.1001.A.6 Servis etiketi ekipmana yapistirildi"),
]
EEBD_CONTROL_ITEMS = [
    ("item_1", "17.E.1001.A.1 Yuz maskesi kontrol edildi"),
    ("item_2", "17.E.1001.A.2 Solunum valfi kontrol edildi"),
    ("item_3", "17.E.1001.A.3 Regulator unitesi kontrol edildi"),
    ("item_4", "17.E.1001.A.4 Kemer kontrol edildi"),
    ("item_5", "17.E.1001.A.5 Silindir kontrol edildi"),
    ("item_6", "17.E.1001.A.6 Servis etiketi ekipmana yapistirildi"),
]
AIR_CYLINDER_CONTROL_ITEMS = [
    ("item_1", "17.C.1001.A.1 Valfi kontrol edildi"),
    ("item_2", "17.C.1001.A.2 Silindir kontrol edildi"),
    ("item_3", "17.C.1001.A.3 Servis etiketi ekipmana yapistirildi"),
]
CABINET_CONTROL_ITEMS = [
    ("item_1", "17.YD.1001.A.1 Erisebilirlik Kontrol Edildi (Dolap onu acik mi, istif veya engel malz. var mi)"),
    ("item_2", "17.YD.1001.A.2 Levhalar Kontrol Edildi (Yangin dolabi isareti ve talimati mevcut mu)"),
    ("item_3", "17.YD.1001.A.3 Kapak ve Kilit Kontrolu Yapildi (Kapak rahat aciliyor mu, kilit saglam mi)"),
    ("item_4", "17.YD.1001.A.4 Dolap Dis Yuzey Kontrol Edildi (Paslanma veya boya kabarmasi var mi)"),
    ("item_5", "17.YD.1001.A.5 Makara Kontrol Edildi (Kolayca acilabiliyor mu)"),
    ("item_6", "17.YD.1001.A.6 Hortum Kontrol Edildi (Catlama, kirilma, sertlesme veya kacak var mi)"),
    ("item_7", "17.YD.1001.A.7 Baglanti Rekorlari Kontrol Edildi (Hortum vana ve lans baglantilari siki mi)"),
    ("item_8", "17.YD.1001.A.8 Dolap ici ve disi temizlik kontrolu yapildi"),
    ("item_9", "17.YD.1001.A.9 Vana kontrol edildi (Vana kolu rahat donuyor mu, kacak veya sizdirma var mi)"),
    ("item_10", "17.YD.1001.A.10 Basinc Kontrol Edildi (Statik ve dinamik basinc degerleri uygun mu) (min.4 bar)"),
    ("item_11", "17.YD.1001.A.11 Lans Kontrol Edildi (Jet/Spray/Kapali konumlari islevsel mi)"),
]
FOAM_CABINET_CONTROL_ITEMS = [
    ("item_1", "17.KYD.1001.A.1 Erisebilirlik Kontrol Edildi (Dolap onu acik mi, istif veya engel malz. var mi)"),
    ("item_2", "17.KYD.1001.A.2 Levhalar Kontrol Edildi (Kopuklu yangin dolabi isareti ve talimati mevcut mu)"),
    ("item_3", "17.KYD.1001.A.3 Kapak ve Kilit Kontrolu Yapildi (Kapak rahat aciliyor mu, kilit saglam mi)"),
    ("item_4", "17.KYD.1001.A.4 Dolap Dis Yuzey Kontrol Edildi (Paslanma veya boya kabarmasi var mi)"),
    ("item_5", "17.KYD.1001.A.5 Makara Kontrol Edildi (Kolayca acilabiliyor mu)"),
    ("item_6", "17.KYD.1001.A.6 Hortum Kontrol Edildi (Catlama, kirilma, sertlesme veya kacak var mi)"),
    ("item_7", "17.KYD.1001.A.7 Baglanti Rekorlari Kontrol Edildi (Hortum vana ve lans baglantilari siki mi)"),
    ("item_8", "17.KYD.1001.A.8 Kopuk Doluluk Orani Kontrol Edildi (Seviye gostergesi kontrolu)"),
    ("item_9", "17.KYD.1001.A.9 Dolap ici ve disi temizlik kontrolu yapildi"),
    ("item_10", "17.KYD.1001.A.10 Kopuk Oranlayici Ayarlari Kontrol Edildi (Mix ayari dogru yuzde mi %1,%3 veya %6)"),
    ("item_11", "17.KYD.1001.A.11 Vana kontrol edildi (Ana su giris vanasi ve kopuk vanasi islevsel mi)"),
    ("item_12", "17.KYD.1001.A.12 Basinc Kontrol Edildi (Sistem calisma basinci kopuk olusumu icin yeterli mi) (min.5-6 bar)"),
    ("item_13", "17.KYD.1001.A.13 Kopuk Lans Kontrol Edildi (Kopuk yapici ozel lans saglam mi)"),
]
HYDRANT_CONTROL_ITEMS = [
    ("item_1", "17.H.1001.A.1 Erisebilirlik Kontrol Edildi (Hidrant cevresinde arac, malzeme engeli var mi)"),
    ("item_2", "17.H.1001.A.2 Gorunurlugu kontrol edildi (Hidrantin kirmizi boyasi canli mi, yonlendirme levhalari var mi)"),
    ("item_3", "17.H.1001.A.3 Kapak Kontrolu Yapildi (Cikis agzindaki kor tapalar/kapaklar takili mi, zincirleri saglam mi)"),
    ("item_4", "17.H.1001.A.4 Genel Dis Yuzey Kontrol Yapildi (Govdede catlak, korozyon veya darbe izi var mi)"),
    ("item_5", "17.H.1001.A.5 Makara Kontrol Edildi (Kolayca acilabiliyor mu)"),
    ("item_6", "17.H.1001.A.6 Acma kapama mili kontrol edildi (Hidrant anahtari ile mil rahatca donuyor mu)"),
    ("item_7", "17.H.1001.A.7 Vana Sizdirmazligi Kontrol Edildi (Hidrant kapaliyken cikis agzindan veya govde altindan su sizintisi var mi)"),
    ("item_8", "17.H.1001.A.8 Cikis agizlari kontrol edildi (Rekor dislerinde veya tirnaklarinda asinma veya deformasyon var mi)"),
]
ASSET_PROFILES = {
    "Yangin Sondurme Cihazi": {
        "label": "Yangin Sondurme Cihazi",
        "type_label": "Tip",
        "class_label": "YSC Sinifi",
        "brand_label": "YSC Uretici",
        "owner_label": "Firma Yetkilisi",
        "last_service_label": "Son Bakim",
        "next_service_label": "Sonraki Bakim",
        "service_input_label": "Son bakim tarihi",
        "next_service_input_label": "Sonraki bakim tarihi",
        "show_weight": True,
        "show_pressure": True,
        "show_hydrostatic": True,
        "control_form_enabled": True,
        "fixed_type": None,
        "monthly_control_items": MONTHLY_CONTROL_ITEMS,
        "control_form_items": CONTROL_FORM_ITEMS,
    },
    "Yangin Elbisesi": {
        "label": "Yangin Elbisesi",
        "type_label": "Ekipman Tipi",
        "class_label": "Kategori / Cinsi",
        "brand_label": "Marka",
        "owner_label": "Ekipman Yetkilisi",
        "last_service_label": "Son Kontrol",
        "next_service_label": "Sonraki Kontrol",
        "service_input_label": "Kontrol tarihi",
        "next_service_input_label": "Sonraki kontrol tarihi",
        "show_weight": True,
        "show_pressure": False,
        "show_hydrostatic": False,
        "control_form_enabled": False,
        "fixed_type": "Yangin Elbisesi",
        "monthly_control_items": FIRE_SUIT_CONTROL_ITEMS,
        "control_form_items": [],
    },
    "Yangin Bareti": {
        "label": "Yangin Bareti",
        "type_label": "Ekipman Tipi",
        "class_label": "Kategori / Cinsi",
        "brand_label": "Marka",
        "owner_label": "Ekipman Yetkilisi",
        "last_service_label": "Son Kontrol",
        "next_service_label": "Sonraki Kontrol",
        "service_input_label": "Kontrol tarihi",
        "next_service_input_label": "Sonraki kontrol tarihi",
        "show_weight": True,
        "show_pressure": False,
        "show_hydrostatic": False,
        "control_form_enabled": False,
        "fixed_type": "Yangin Bareti",
        "monthly_control_items": HELMET_CONTROL_ITEMS,
        "control_form_items": [],
    },
    "Yangin Baltasi": {
        "label": "Yangin Baltasi",
        "type_label": "Ekipman Tipi",
        "class_label": "Kategori / Cinsi",
        "brand_label": "Marka",
        "owner_label": "Ekipman Yetkilisi",
        "last_service_label": "Son Kontrol",
        "next_service_label": "Sonraki Kontrol",
        "service_input_label": "Kontrol tarihi",
        "next_service_input_label": "Sonraki kontrol tarihi",
        "show_weight": False,
        "show_pressure": False,
        "show_hydrostatic": False,
        "control_form_enabled": False,
        "fixed_type": "Yangin Baltasi",
        "monthly_control_items": AXE_CONTROL_ITEMS,
        "control_form_items": [],
    },
    "SCBA": {
        "label": "SCBA",
        "type_label": "Ekipman Tipi",
        "class_label": "Kategori / Cinsi",
        "brand_label": "Marka",
        "owner_label": "Ekipman Yetkilisi",
        "last_service_label": "Dolum Tarihi",
        "next_service_label": "Sonraki Kontrol",
        "service_input_label": "Dolum tarihi",
        "next_service_input_label": "Sonraki kontrol tarihi",
        "show_weight": True,
        "show_pressure": False,
        "show_hydrostatic": True,
        "control_form_enabled": False,
        "fixed_type": "SCBA",
        "monthly_control_items": SCBA_CONTROL_ITEMS,
        "control_form_items": [],
    },
    "EEBD": {
        "label": "EEBD",
        "type_label": "Ekipman Tipi",
        "class_label": "Kategori / Cinsi",
        "brand_label": "Marka",
        "owner_label": "Ekipman Yetkilisi",
        "last_service_label": "Dolum Tarihi",
        "next_service_label": "Sonraki Kontrol",
        "service_input_label": "Dolum tarihi",
        "next_service_input_label": "Sonraki kontrol tarihi",
        "show_weight": True,
        "show_pressure": False,
        "show_hydrostatic": True,
        "control_form_enabled": False,
        "fixed_type": "EEBD",
        "monthly_control_items": EEBD_CONTROL_ITEMS,
        "control_form_items": [],
    },
    "Hava Tupu": {
        "label": "Hava Tupu",
        "type_label": "Ekipman Tipi",
        "class_label": "Kategori / Cinsi",
        "brand_label": "Marka",
        "owner_label": "Ekipman Yetkilisi",
        "last_service_label": "Dolum Tarihi",
        "next_service_label": "Sonraki Kontrol",
        "service_input_label": "Dolum tarihi",
        "next_service_input_label": "Sonraki kontrol tarihi",
        "show_weight": True,
        "show_pressure": False,
        "show_hydrostatic": True,
        "control_form_enabled": False,
        "fixed_type": "Hava Tupu",
        "monthly_control_items": AIR_CYLINDER_CONTROL_ITEMS,
        "control_form_items": [],
    },
    "Yangin Sondurme Dolabi": {
        "label": "Yangin Sondurme Dolabi",
        "type_label": "Ekipman Tipi",
        "class_label": "Kategori / Cinsi",
        "brand_label": "Marka",
        "owner_label": "Ekipman Yetkilisi",
        "last_service_label": "Son Kontrol",
        "next_service_label": "Sonraki Kontrol",
        "service_input_label": "Kontrol tarihi",
        "next_service_input_label": "Sonraki kontrol tarihi",
        "show_weight": False,
        "show_pressure": False,
        "show_hydrostatic": False,
        "control_form_enabled": False,
        "fixed_type": "Yangin Sondurme Dolabi",
        "monthly_control_items": CABINET_CONTROL_ITEMS,
        "control_form_items": [],
    },
    "Kopuklu Yangin Sondurme Dolabi": {
        "label": "Kopuklu Yangin Sondurme Dolabi",
        "type_label": "Ekipman Tipi",
        "class_label": "Kategori / Cinsi",
        "brand_label": "Marka",
        "owner_label": "Ekipman Yetkilisi",
        "last_service_label": "Son Kontrol",
        "next_service_label": "Sonraki Kontrol",
        "service_input_label": "Kontrol tarihi",
        "next_service_input_label": "Sonraki kontrol tarihi",
        "show_weight": False,
        "show_pressure": False,
        "show_hydrostatic": False,
        "control_form_enabled": False,
        "fixed_type": "Kopuklu Yangin Sondurme Dolabi",
        "monthly_control_items": FOAM_CABINET_CONTROL_ITEMS,
        "control_form_items": [],
    },
    "Yangin Hidranti": {
        "label": "Yangin Hidranti",
        "type_label": "Ekipman Tipi",
        "class_label": "Kategori / Cinsi",
        "brand_label": "Marka",
        "owner_label": "Ekipman Yetkilisi",
        "last_service_label": "Son Kontrol",
        "next_service_label": "Sonraki Kontrol",
        "service_input_label": "Kontrol tarihi",
        "next_service_input_label": "Sonraki kontrol tarihi",
        "show_weight": False,
        "show_pressure": False,
        "show_hydrostatic": False,
        "control_form_enabled": False,
        "fixed_type": "Yangin Hidranti",
        "monthly_control_items": HYDRANT_CONTROL_ITEMS,
        "control_form_items": [],
    },
    "Elektrik Ic Tesisati": {
        "label": "Elektrik Ic Tesisati",
        "type_label": "Rapor Tipi",
        "class_label": "Sebeke Tipi",
        "brand_label": "Enerji Saglayan Kurulus",
        "owner_label": "Rapor Numarasi",
        "last_service_label": "Rapor Tarihi",
        "next_service_label": "Bir Sonraki Kontrol",
        "service_input_label": "Rapor tarihi",
        "next_service_input_label": "Bir sonraki periyodik kontrol tarihi",
        "show_weight": False,
        "show_pressure": False,
        "show_hydrostatic": False,
        "control_form_enabled": False,
        "fixed_type": "Elektrik Ic Tesisati",
        "monthly_control_items": [],
        "control_form_items": [],
    },
}
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

companies = Table(
    "companies",
    metadata,
    Column("id", Integer, primary_key=True),
    Column("public_id", String(32), nullable=False, unique=True),
    Column("slug", String(255), nullable=False, unique=True),
    Column("name", String(255), nullable=False, unique=True),
    Column("address", String(255), nullable=False),
    Column("contact_name", String(255), nullable=False),
    Column("created_at", String(32), nullable=False),
    Column("updated_at", String(32), nullable=False),
)

extinguishers = Table(
    "extinguishers",
    metadata,
    Column("id", Integer, primary_key=True),
    Column("public_id", String(32), nullable=False, unique=True),
    Column("serial_number", String(128), nullable=False, unique=True),
    Column("company_id", Integer, ForeignKey("companies.id")),
    Column("company_name", String(255), nullable=False),
    Column("location_detail", String(255), nullable=False),
    Column("weight_kg", Float, nullable=False),
    Column("asset_category", String(128), nullable=False, default=DEFAULT_ASSET_CATEGORY),
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
    Column("item_7", Boolean, nullable=False, default=False),
    Column("item_8", Boolean, nullable=False, default=False),
    Column("item_9", Boolean, nullable=False, default=False),
    Column("item_10", Boolean, nullable=False, default=False),
    Column("item_11", Boolean, nullable=False, default=False),
    Column("item_12", Boolean, nullable=False, default=False),
    Column("item_13", Boolean, nullable=False, default=False),
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


def ensure_monthly_inspection_columns() -> None:
    extra_boolean_columns = [
        "item_7",
        "item_8",
        "item_9",
        "item_10",
        "item_11",
        "item_12",
        "item_13",
    ]
    with engine.begin() as connection:
        if engine.dialect.name == "sqlite":
            existing_rows = connection.execute(text("PRAGMA table_info(monthly_inspections)")).fetchall()
            existing = {row[1] for row in existing_rows}
            for column_name in extra_boolean_columns:
                if column_name not in existing:
                    connection.execute(
                        text(
                            f"ALTER TABLE monthly_inspections ADD COLUMN {column_name} BOOLEAN NOT NULL DEFAULT 0"
                        )
                    )
        else:
            existing_rows = connection.execute(
                text(
                    "SELECT column_name FROM information_schema.columns "
                    "WHERE table_name = 'monthly_inspections'"
                )
            ).fetchall()
            existing = {row[0] for row in existing_rows}
            for column_name in extra_boolean_columns:
                if column_name not in existing:
                    connection.execute(
                        text(
                            f"ALTER TABLE monthly_inspections ADD COLUMN {column_name} BOOLEAN NOT NULL DEFAULT FALSE"
                        )
                    )


ensure_monthly_inspection_columns()

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
            company_tables = {
                row[0]
                for row in connection.execute(
                    text("SELECT name FROM sqlite_master WHERE type='table' AND name='companies'")
                ).fetchall()
            }
            if "companies" not in company_tables:
                connection.execute(
                    text(
                        """
                        CREATE TABLE companies (
                            id INTEGER PRIMARY KEY,
                            public_id VARCHAR(32) NOT NULL UNIQUE,
                            slug VARCHAR(255) NOT NULL UNIQUE,
                            name VARCHAR(255) NOT NULL UNIQUE,
                            address VARCHAR(255) NOT NULL,
                            contact_name VARCHAR(255) NOT NULL,
                            created_at VARCHAR(32) NOT NULL,
                            updated_at VARCHAR(32) NOT NULL
                        )
                        """
                    )
                )
            columns = {
                row[1] for row in connection.execute(text("PRAGMA table_info(users)")).fetchall()
            }
            if "is_active" not in columns:
                connection.execute(text("ALTER TABLE users ADD COLUMN is_active BOOLEAN NOT NULL DEFAULT 1"))
            extinguisher_columns = {
                row[1] for row in connection.execute(text("PRAGMA table_info(extinguishers)")).fetchall()
            }
            for column_name, column_def in [
                ("company_id", "INTEGER"),
                ("fire_class", "TEXT"),
                ("manufacturer", "TEXT"),
                ("hydrostatic_test_date", "TEXT"),
                ("company_address", "TEXT"),
                ("company_contact", "TEXT"),
                ("asset_category", "TEXT"),
            ]:
                if column_name not in extinguisher_columns:
                    connection.execute(text(f"ALTER TABLE extinguishers ADD COLUMN {column_name} {column_def}"))
            company_columns = {
                row[1] for row in connection.execute(text("PRAGMA table_info(companies)")).fetchall()
            }
            if "public_id" not in company_columns:
                connection.execute(text("ALTER TABLE companies ADD COLUMN public_id TEXT"))
            if "slug" not in company_columns:
                connection.execute(text("ALTER TABLE companies ADD COLUMN slug TEXT"))

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
            table_exists = connection.execute(
                text(
                    """
                    SELECT table_name
                    FROM information_schema.tables
                    WHERE table_name = 'companies'
                    """
                )
            ).fetchone()
            if table_exists is None:
                connection.execute(
                    text(
                        """
                        CREATE TABLE companies (
                            id SERIAL PRIMARY KEY,
                            public_id VARCHAR(32) NOT NULL UNIQUE,
                            slug VARCHAR(255) NOT NULL UNIQUE,
                            name VARCHAR(255) NOT NULL UNIQUE,
                            address VARCHAR(255) NOT NULL,
                            contact_name VARCHAR(255) NOT NULL,
                            created_at VARCHAR(32) NOT NULL,
                            updated_at VARCHAR(32) NOT NULL
                        )
                        """
                    )
                )
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
            for column_name in ["company_id", "fire_class", "manufacturer", "hydrostatic_test_date", "company_address", "company_contact", "asset_category"]:
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
                    column_type = "INTEGER" if column_name == "company_id" else "TEXT"
                    connection.execute(text(f"ALTER TABLE extinguishers ADD COLUMN {column_name} {column_type}"))
            result = connection.execute(
                text(
                    """
                    SELECT column_name
                    FROM information_schema.columns
                    WHERE table_name = 'companies' AND column_name = 'public_id'
                    """
                )
            ).fetchone()
            if result is None:
                connection.execute(text("ALTER TABLE companies ADD COLUMN public_id TEXT"))
            result = connection.execute(
                text(
                    """
                    SELECT column_name
                    FROM information_schema.columns
                    WHERE table_name = 'companies' AND column_name = 'slug'
                    """
                )
            ).fetchone()
            if result is None:
                connection.execute(text("ALTER TABLE companies ADD COLUMN slug TEXT"))
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


def seed_companies_from_extinguishers() -> None:
    now = datetime.now().isoformat(timespec="seconds")
    with engine.begin() as connection:
        company_rows = {
            row.name: {"id": row.id, "public_id": row.public_id, "slug": row.slug}
            for row in connection.execute(select(companies.c.id, companies.c.name, companies.c.public_id, companies.c.slug)).all()
        }
        extinguisher_rows = connection.execute(
            select(
                extinguishers.c.id,
                extinguishers.c.company_id,
                extinguishers.c.company_name,
                extinguishers.c.company_address,
                extinguishers.c.company_contact,
            )
        ).mappings().all()
        for row in extinguisher_rows:
            company_name = (row["company_name"] or "").strip()
            if not company_name:
                continue
            company_info = company_rows.get(company_name)
            if company_info is None:
                result = connection.execute(
                    insert(companies).values(
                        public_id=uuid.uuid4().hex[:12],
                        slug=build_unique_company_slug(connection, company_name),
                        name=company_name,
                        address=(row.get("company_address") or "-").strip() or "-",
                        contact_name=(row.get("company_contact") or "-").strip() or "-",
                        created_at=now,
                        updated_at=now,
                    )
                )
                company_id = result.inserted_primary_key[0]
                company_rows[company_name] = {"id": company_id, "public_id": None, "slug": None}
            else:
                company_id = company_info["id"]
            if row.get("company_id") != company_id:
                connection.execute(
                    update(extinguishers)
                    .where(extinguishers.c.id == row["id"])
                    .values(company_id=company_id)
                )
        existing_companies = connection.execute(
            select(companies.c.id, companies.c.public_id, companies.c.slug, companies.c.name)
        ).mappings().all()
        for company in existing_companies:
            values = {}
            if not company.get("public_id"):
                values["public_id"] = uuid.uuid4().hex[:12]
            if not company.get("slug"):
                values["slug"] = build_unique_company_slug(connection, company["name"], exclude_id=company["id"])
            if values:
                values["updated_at"] = now
                connection.execute(
                    update(companies)
                    .where(companies.c.id == company["id"])
                    .values(**values)
                )


def seed_asset_categories() -> None:
    now = datetime.now().isoformat(timespec="seconds")
    with engine.begin() as connection:
        rows = connection.execute(
            select(extinguishers.c.id, extinguishers.c.asset_category)
        ).mappings().all()
        for row in rows:
            if row.get("asset_category"):
                continue
            connection.execute(
                update(extinguishers)
                .where(extinguishers.c.id == row["id"])
                .values(asset_category=DEFAULT_ASSET_CATEGORY, updated_at=now)
            )


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
    allowed_endpoints = {
        "login",
        "logout",
        "public_detail",
        "public_control_form_pdf",
        "public_company_portal",
        "public_company_assets",
        "health",
        "static",
    }
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


def parse_structured_notes(notes: str | None) -> dict[str, str]:
    parsed: dict[str, str] = {}
    if not notes:
        return parsed
    for raw_line in notes.splitlines():
        line = raw_line.strip()
        if not line or ":" not in line:
            continue
        key, value = line.split(":", 1)
        parsed[key.strip()] = value.strip()
    return parsed


def build_electrical_note_sections(notes: str | None) -> list[dict]:
    parsed = parse_structured_notes(notes)
    sections: list[dict] = []
    for title, keys in ELECTRICAL_NOTE_SECTIONS:
        items = []
        for key in keys:
            value = parsed.get(key)
            if value:
                items.append({"label": key, "value": value})
        if items:
            sections.append({"title": title, "items": items})
    return sections


def get_extinguisher(public_id: str) -> dict:
    extinguisher = fetch_one(
        select(extinguishers).where(extinguishers.c.public_id == public_id)
    )
    if extinguisher is None:
        abort(404)
    return extinguisher


def get_company(company_id: int) -> dict:
    company = fetch_one(select(companies).where(companies.c.id == company_id))
    if company is None:
        abort(404)
    return company


def get_company_by_public_id(public_id: str) -> dict:
    company = fetch_one(select(companies).where(companies.c.public_id == public_id))
    if company is None:
        abort(404)
    return company


def get_company_by_slug(slug: str) -> dict:
    company = fetch_one(select(companies).where(companies.c.slug == slug))
    if company is None:
        abort(404)
    return company


def get_company_choices() -> list[dict]:
    return fetch_all(select(companies).order_by(companies.c.name))


def get_asset_category_choices() -> list[dict]:
    return ASSET_CATEGORIES


def get_asset_category(slug: str | None = None, *, label: str | None = None) -> dict | None:
    if slug:
        return ASSET_CATEGORY_BY_SLUG.get(slug)
    if label:
        return ASSET_CATEGORY_BY_LABEL.get(label)
    return None


def slugify_company_name(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode("ascii")
    slug = re.sub(r"[^a-zA-Z0-9]+", "-", normalized.lower()).strip("-")
    return slug or "firma"


def build_unique_company_slug(connection, name: str, exclude_id: int | None = None) -> str:
    base_slug = slugify_company_name(name)
    slug = base_slug
    index = 2
    while True:
        statement = select(companies.c.id).where(companies.c.slug == slug)
        row = connection.execute(statement).first()
        if row is None or (exclude_id is not None and row.id == exclude_id):
            return slug
        slug = f"{base_slug}-{index}"
        index += 1


def resolve_company_slug(connection, raw_slug: str, fallback_name: str, *, exclude_id: int | None = None) -> str:
    source = (raw_slug or "").strip() or fallback_name
    return build_unique_company_slug(connection, source, exclude_id=exclude_id)


run_schema_migrations()
seed_default_users()
seed_companies_from_extinguishers()
seed_asset_categories()


def build_company_portal_sections(company_id: int) -> list[dict]:
    assets = fetch_all(
        select(
            extinguishers.c.public_id,
            extinguishers.c.asset_category,
            extinguishers.c.serial_number,
            extinguishers.c.location_detail,
            extinguishers.c.extinguisher_type,
            extinguishers.c.last_service_date,
            extinguishers.c.next_service_date,
            extinguishers.c.pressure_status,
        )
        .where(extinguishers.c.company_id == company_id)
        .order_by(extinguishers.c.asset_category, extinguishers.c.location_detail, extinguishers.c.serial_number)
    )
    grouped: dict[str, list[dict]] = {}
    for asset in assets:
        label = asset.get("asset_category") or DEFAULT_ASSET_CATEGORY
        grouped.setdefault(label, []).append(asset)

    sections = []
    for category in ASSET_CATEGORIES:
        items = grouped.get(category["label"], [])
        sections.append(
            {
                **category,
                "count": len(items),
                "items": items,
            }
        )
    return sections


def get_registration_groups() -> list[dict]:
    return REGISTRATION_GROUPS


def get_registration_group(group_slug: str) -> dict:
    group = REGISTRATION_GROUP_BY_SLUG.get(group_slug)
    if group is None:
        abort(404)
    return group


def get_asset_profile(asset_category: str | None) -> dict:
    return ASSET_PROFILES.get(asset_category or "", ASSET_PROFILES["Yangin Sondurme Cihazi"])


def generate_prefixed_report_number(prefix: str) -> str:
    year = datetime.now().year
    like_pattern = f"{prefix}-{year}-%"
    existing = fetch_all(
        select(extinguishers.c.serial_number)
        .where(extinguishers.c.serial_number.like(like_pattern))
        .order_by(desc(extinguishers.c.serial_number))
    )
    max_number = 0
    for row in existing:
        value = row.get("serial_number") or ""
        match = re.match(rf"^{re.escape(prefix)}-{year}-(\d+)$", value)
        if match:
            max_number = max(max_number, int(match.group(1)))
    return f"{prefix}-{year}-{max_number + 1:04d}"


def render_profile_record_form(group_slug: str):
    group = get_registration_group(group_slug)
    if group_slug == "yangin-sondurme-cihazi":
        return redirect(url_for("create_extinguisher"))

    company_choices = get_company_choices()
    asset_profile = get_asset_profile(group["label"])

    if request.method == "POST":
        form = parse_required_form(request.form)
        form["technician_name"] = form.get("technician_name") or current_user_full_name()
        form["asset_category"] = group["label"]
        form["extinguisher_type"] = asset_profile.get("fixed_type") or group["label"]
        try:
            form, selected_company = sync_company_payload_from_selection(form)
        except ValueError as exc:
            flash(str(exc), "error")
            return render_template(
                "create_asset_profile.html",
                form=form,
                companies=company_choices,
                asset_profile=asset_profile,
                group=group,
            )

        required_fields = {
            "serial_number": "Seri numarasi",
            "company_id": "Cari secimi",
            "company_contact": asset_profile["owner_label"],
            "location_detail": "Bulundugu yer",
            "fire_class": asset_profile["class_label"],
            "manufacturer": asset_profile["brand_label"],
            "last_service_date": asset_profile["service_input_label"],
            "next_service_date": asset_profile["next_service_input_label"],
            "technician_name": "Teknisyen",
            "operation_summary": "Yapilan islem",
        }
        if asset_profile["show_weight"]:
            required_fields["weight_kg"] = "Kg"
        if asset_profile["show_hydrostatic"]:
            required_fields["hydrostatic_test_date"] = "Hidrostatik test tarihi"

        missing = [label for key, label in required_fields.items() if not form.get(key)]
        if missing:
            flash(f"Eksik alanlar: {', '.join(missing)}", "error")
            return render_template(
                "create_asset_profile.html",
                form=form,
                companies=company_choices,
                asset_profile=asset_profile,
                group=group,
            )

        try:
            weight_kg = parse_float(form["weight_kg"], "Kg") if asset_profile["show_weight"] else 0.0
        except ValueError as exc:
            flash(str(exc), "error")
            return render_template(
                "create_asset_profile.html",
                form=form,
                companies=company_choices,
                asset_profile=asset_profile,
                group=group,
            )

        now = datetime.now().isoformat(timespec="seconds")
        public_id = uuid.uuid4().hex[:12]
        inspection_values = build_monthly_inspection_values(request.form, asset_profile["monthly_control_items"])
        control_values = build_control_form_values({})

        try:
            with engine.begin() as connection:
                result = connection.execute(
                    insert(extinguishers).values(
                        public_id=public_id,
                        serial_number=form["serial_number"],
                        company_id=selected_company["id"],
                        company_name=form["company_name"],
                        company_address=form["company_address"],
                        company_contact=form["company_contact"],
                        asset_category=group["label"],
                        location_detail=form["location_detail"],
                        weight_kg=weight_kg,
                        extinguisher_type=form["extinguisher_type"],
                        fire_class=form["fire_class"],
                        manufacturer=form["manufacturer"],
                        hydrostatic_test_date=form.get("hydrostatic_test_date") if asset_profile["show_hydrostatic"] else None,
                        pressure_status=None,
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
                        pressure_status=None,
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
                "create_asset_profile.html",
                form=form,
                companies=company_choices,
                asset_profile=asset_profile,
                group=group,
            )

        flash(f"{group['label']} kaydedildi.", "success")
        return redirect(url_for("extinguisher_detail", public_id=public_id))

    return render_template(
        "create_asset_profile.html",
        form={
            "technician_name": current_user_full_name(),
            "asset_category": group["label"],
            "extinguisher_type": asset_profile.get("fixed_type") or group["label"],
        },
        companies=company_choices,
        asset_profile=asset_profile,
        group=group,
    )


@app.route("/records/new/elektrik-ic-tesisati", methods=["GET", "POST"])
@login_required
def create_electrical_installation():
    group = get_registration_group("elektrik-ic-tesisati")
    asset_profile = get_asset_profile(group["label"])
    company_choices = get_company_choices()

    if request.method == "POST":
        form = parse_required_form(request.form)
        form["technician_name"] = current_user_full_name()
        form["report_number"] = generate_prefixed_report_number("ELK")
        try:
            form, selected_company = sync_company_payload_from_selection(form)
        except ValueError as exc:
            flash(str(exc), "error")
            return render_template(
                "create_electrical_installation.html",
                form=form,
                companies=company_choices,
                group=group,
                asset_profile=asset_profile,
            )

        required_fields = {
            "company_id": "Cari secimi",
            "company_address": "Periyodik kontrol adresi",
            "report_date": "Rapor tarihi",
            "control_start": "Periyodik kontrol baslangic tarihi ve saati",
            "control_end": "Periyodik kontrol bitis tarihi ve saati",
            "next_service_date": "Bir sonraki periyodik kontrol tarihi",
            "energy_provider": "Enerji saglayan kurulus",
            "grid_type": "Sebeke tipi",
            "grid_voltage": "Sebeke gerilimi",
            "equipment_usage_purpose": "Ekipmanin kullanim amaci",
            "technician_name": "Teknisyen",
        }
        missing = [label for key, label in required_fields.items() if not form.get(key)]
        if missing:
            flash(f"Eksik alanlar: {', '.join(missing)}", "error")
            return render_template(
                "create_electrical_installation.html",
                form=form,
                companies=company_choices,
                group=group,
                asset_profile=asset_profile,
            )

        structured_notes = {
            "isg_katip_id": form.get("isg_katip_id", ""),
            "sgk_number": form.get("sgk_number", ""),
            "control_method": form.get("control_method", ""),
            "project_exists": form.get("project_exists", ""),
            "single_line_schema": form.get("single_line_schema", ""),
            "control_reason": form.get("control_reason", ""),
            "grounder_type": form.get("grounder_type", ""),
            "structure_type": form.get("structure_type", ""),
            "last_control_date": form.get("last_control_date", ""),
            "phase_conductor_type": form.get("phase_conductor_type", ""),
            "ground_resistance": form.get("ground_resistance", ""),
            "additional_ground_details": form.get("additional_ground_details", ""),
            "system_ground_conductor": form.get("system_ground_conductor", ""),
            "equipotential_conductor": form.get("equipotential_conductor", ""),
            "supply_characteristics": form.get("supply_characteristics", ""),
            "main_rcd_nominal": form.get("main_rcd_nominal", ""),
            "main_switch_characteristics": form.get("main_switch_characteristics", ""),
            "main_rcd_test": form.get("main_rcd_test", ""),
            "major_installation_change": form.get("major_installation_change", ""),
            "spd_used": form.get("spd_used", ""),
            "direct_contact_protections": form.get("direct_contact_protections", ""),
            "previous_control_label": form.get("previous_control_label", ""),
            "findings": form.get("findings", ""),
            "thermal_camera_1": form.get("thermal_camera_1", ""),
            "thermal_camera_2": form.get("thermal_camera_2", ""),
            "thermal_calibration_1": form.get("thermal_calibration_1", ""),
            "thermal_calibration_2": form.get("thermal_calibration_2", ""),
            "thermal_validity_1": form.get("thermal_validity_1", ""),
            "thermal_validity_2": form.get("thermal_validity_2", ""),
            "thermal_serial_1": form.get("thermal_serial_1", ""),
            "thermal_serial_2": form.get("thermal_serial_2", ""),
            "thermal_calibration_no_1": form.get("thermal_calibration_no_1", ""),
            "thermal_calibration_no_2": form.get("thermal_calibration_no_2", ""),
            "measurement_device_1": form.get("measurement_device_1", ""),
            "measurement_device_2": form.get("measurement_device_2", ""),
            "measurement_calibration_1": form.get("measurement_calibration_1", ""),
            "measurement_calibration_2": form.get("measurement_calibration_2", ""),
            "measurement_validity_1": form.get("measurement_validity_1", ""),
            "measurement_validity_2": form.get("measurement_validity_2", ""),
            "measurement_serial_1": form.get("measurement_serial_1", ""),
            "measurement_serial_2": form.get("measurement_serial_2", ""),
            "measurement_calibration_no_1": form.get("measurement_calibration_no_1", ""),
            "measurement_calibration_no_2": form.get("measurement_calibration_no_2", ""),
            "control_criteria_notes": form.get("control_criteria_notes", ""),
            "measurement_method": form.get("measurement_method", ""),
            "thermal_photo_date": form.get("thermal_photo_date", ""),
            "thermal_photo_number": form.get("thermal_photo_number", ""),
            "thermal_loose_contact_heating": form.get("thermal_loose_contact_heating", ""),
            "thermal_overload_heating": form.get("thermal_overload_heating", ""),
            "section_61_notes": form.get("section_61_notes", ""),
            "section_62_notes": form.get("section_62_notes", ""),
            "section_63_notes": form.get("section_63_notes", ""),
            "fault_notes": form.get("fault_notes", ""),
            "equipment_photos_notes": form.get("equipment_photos_notes", ""),
            "general_notes": form.get("general_notes", ""),
            "final_conclusion": form.get("final_conclusion", ""),
            "authorized_person_name": form.get("authorized_person_name", ""),
            "authorized_person_job": form.get("authorized_person_job", ""),
            "authorized_person_registry": form.get("authorized_person_registry", ""),
            "copy_count": form.get("copy_count", ""),
        }
        note_lines = [
            "Elektrik Ic Tesisati Tam Form Baslangic Kaydi",
            f"ISG-KATIP Sozlesme ID: {structured_notes['isg_katip_id'] or '-'}",
            f"SGK Sicil Numarasi: {structured_notes['sgk_number'] or '-'}",
            f"Periyodik Kontrol Metodu ve Kapsami: {structured_notes['control_method'] or '-'}",
            f"Tesise ait proje var mi: {structured_notes['project_exists'] or '-'}",
            f"Tek hat semasi var mi: {structured_notes['single_line_schema'] or '-'}",
            f"Kontrol nedeni: {structured_notes['control_reason'] or '-'}",
            f"Topraklayici tipi: {structured_notes['grounder_type'] or '-'}",
            f"Yapi cinsi: {structured_notes['structure_type'] or '-'}",
            f"Son kontrol tarihi: {structured_notes['last_control_date'] or '-'}",
            f"Faz iletkenlerinin sayisi ve tipi: {structured_notes['phase_conductor_type'] or '-'}",
            f"Temel topraklama direnci: {structured_notes['ground_resistance'] or '-'}",
            f"Ilave topraklama elektrotu detaylari: {structured_notes['additional_ground_details'] or '-'}",
            f"Sistem topraklama iletkeni ve kesiti: {structured_notes['system_ground_conductor'] or '-'}",
            f"Ana espotansiyel iletkeni ve kesiti: {structured_notes['equipotential_conductor'] or '-'}",
            f"Besleme kaynagi karakteristikleri: {structured_notes['supply_characteristics'] or '-'}",
            f"Ana RCD anma akimi: {structured_notes['main_rcd_nominal'] or '-'}",
            f"Ana kesici karakteristikleri: {structured_notes['main_switch_characteristics'] or '-'}",
            f"Ana RCD test akimi ve suresi: {structured_notes['main_rcd_test'] or '-'}",
            f"Tesisatta kapsamli degisiklik var mi (>%20): {structured_notes['major_installation_change'] or '-'}",
            f"Asiri gerilim koruma cihazlari kullanilmis mi: {structured_notes['spd_used'] or '-'}",
            f"Dogrudan dokunmaya karsi koruma onlemleri: {structured_notes['direct_contact_protections'] or '-'}",
            f"Bir onceki periyodik kontrol etiketi var mi: {structured_notes['previous_control_label'] or '-'}",
            f"Tespit edilen bilgiler: {structured_notes['findings'] or '-'}",
            f"Termal Kamera 1: {structured_notes['thermal_camera_1'] or '-'} / Seri: {structured_notes['thermal_serial_1'] or '-'} / Kal.No: {structured_notes['thermal_calibration_no_1'] or '-'}",
            f"Termal Kamera 2: {structured_notes['thermal_camera_2'] or '-'} / Seri: {structured_notes['thermal_serial_2'] or '-'} / Kal.No: {structured_notes['thermal_calibration_no_2'] or '-'}",
            f"Olcum Aleti 1: {structured_notes['measurement_device_1'] or '-'} / Seri: {structured_notes['measurement_serial_1'] or '-'} / Kal.No: {structured_notes['measurement_calibration_no_1'] or '-'}",
            f"Olcum Aleti 2: {structured_notes['measurement_device_2'] or '-'} / Seri: {structured_notes['measurement_serial_2'] or '-'} / Kal.No: {structured_notes['measurement_calibration_no_2'] or '-'}",
            f"Kontrol Kriterleri ve Testler: {structured_notes['control_criteria_notes'] or '-'}",
            f"Olcum ve Dogrulama Metodu: {structured_notes['measurement_method'] or '-'}",
            f"Termal fotograf tarihi: {structured_notes['thermal_photo_date'] or '-'}",
            f"Termal fotograf no: {structured_notes['thermal_photo_number'] or '-'}",
            f"Kontak gevsakligi isinmasi: {structured_notes['thermal_loose_contact_heating'] or '-'}",
            f"Asiri yuk isinmasi: {structured_notes['thermal_overload_heating'] or '-'}",
            f"6.1 Notlari: {structured_notes['section_61_notes'] or '-'}",
            f"6.2 Notlari: {structured_notes['section_62_notes'] or '-'}",
            f"6.3 Notlari: {structured_notes['section_63_notes'] or '-'}",
            f"Kusur Aciklamalari: {structured_notes['fault_notes'] or '-'}",
            f"Ekipman Fotograflari: {structured_notes['equipment_photos_notes'] or '-'}",
            f"Genel Notlar: {structured_notes['general_notes'] or '-'}",
            f"Sonuc ve Kanaat: {structured_notes['final_conclusion'] or '-'}",
            f"Yetkili Kisi: {structured_notes['authorized_person_name'] or '-'} / Meslek: {structured_notes['authorized_person_job'] or '-'} / Kayit No: {structured_notes['authorized_person_registry'] or '-'}",
            f"Nusha Sayisi: {structured_notes['copy_count'] or '-'}",
        ]

        now = datetime.now().isoformat(timespec="seconds")
        public_id = uuid.uuid4().hex[:12]
        with engine.begin() as connection:
            result = connection.execute(
                insert(extinguishers).values(
                    public_id=public_id,
                    serial_number=form["report_number"],
                    company_id=selected_company["id"],
                    company_name=form["company_name"],
                    company_address=form["company_address"],
                    company_contact=form["report_number"],
                    asset_category=group["label"],
                    location_detail=form["company_address"],
                    weight_kg=0.0,
                    extinguisher_type="Elektrik Ic Tesisati",
                    fire_class=form["grid_type"],
                    manufacturer=form["energy_provider"],
                    hydrostatic_test_date=None,
                    pressure_status=None,
                    notes="\n".join(note_lines),
                    last_service_date=form["report_date"],
                    next_service_date=form["next_service_date"],
                    created_at=now,
                    updated_at=now,
                )
            )
            extinguisher_id = result.inserted_primary_key[0]
            connection.execute(
                insert(service_logs).values(
                    extinguisher_id=extinguisher_id,
                    service_date=form["report_date"],
                    technician_name=form["technician_name"],
                    operation_summary=(
                        f"Elektrik ic tesisati periyodik kontrol raporu olusturuldu. "
                        f"Baslangic: {form['control_start']} / Bitis: {form['control_end']} / "
                        f"Kullanim amaci: {form['equipment_usage_purpose']}"
                    ),
                    pressure_status=None,
                    notes=form.get("findings"),
                    created_at=now,
                )
            )

        flash("Elektrik ic tesisati kaydi olusturuldu.", "success")
        return redirect(url_for("extinguisher_detail", public_id=public_id))

    report_number = generate_prefixed_report_number("ELK")
    return render_template(
        "create_electrical_installation.html",
        form={
            "technician_name": current_user_full_name(),
            "report_number": report_number,
            "control_method": (
                "- TS HD 60364-4-43 Alcak Gerilim Elektrik Tesisatlari - Bolum 4: Guvenlik Icin Koruma Grup 43 - Asiri Akima Karsi Koruma\n"
                "- TS HD 60364-6 Alcak Gerilim Elektrik Tesisatlari - Bolum 6: Dogrulama\n"
                "- Is Ekipmanlarinin Kullaniminda Saglik ve Guvenlik Sartlari Yonetmeligi\n"
                "- Elektrik Ic Tesisleri Yonetmeligi\n"
                "- Elektrik Tesislerinde Topraklamalar Yonetmeligi"
            ),
            "final_conclusion": "Periyodik kontrol tarihi itibari ile mevcut sartlar altinda kullanimi uygundur/uygun degildir.",
            "copy_count": "2",
        },
        companies=company_choices,
        group=group,
        asset_profile=asset_profile,
    )


def sync_company_payload_from_selection(form: dict[str, str]) -> tuple[dict[str, str], dict]:
    company_id_raw = (form.get("company_id") or "").strip()
    if not company_id_raw:
        raise ValueError("Firma secilmeden kayit olusturulamaz.")
    try:
        company_id = int(company_id_raw)
    except ValueError as exc:
        raise ValueError("Firma secimi gecersiz.") from exc
    company = get_company(company_id)
    form["company_id"] = str(company["id"])
    form["company_name"] = company["name"]
    form["company_address"] = company["address"]
    return form, company


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


def build_monthly_inspection_values(form_data, items: list[tuple[str, str]] | None = None) -> dict[str, bool]:
    source_items = items or MONTHLY_CONTROL_ITEMS
    return {key: form_data.get(key) == "on" for key, _label in source_items}


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


def with_monthly_control_labels(rows: list[dict], items: list[tuple[str, str]] | None = None) -> list[dict]:
    source_items = items or MONTHLY_CONTROL_ITEMS
    enriched_rows: list[dict] = []
    for row in rows:
        checks = []
        passed_count = 0
        for key, label in source_items:
            passed = bool(row.get(key))
            if passed:
                passed_count += 1
            checks.append({"key": key, "label": label, "passed": passed})
        copied = dict(row)
        copied["checks"] = checks
        copied["passed_count"] = passed_count
        copied["total_count"] = len(source_items)
        enriched_rows.append(copied)
    return enriched_rows


def get_equipment_preset(extinguisher_type: str | None) -> dict | None:
    if not extinguisher_type:
        return None
    return EQUIPMENT_PRESETS.get(extinguisher_type)


def build_monthly_table(rows: list[dict], items: list[tuple[str, str]] | None = None) -> dict:
    source_items = items or MONTHLY_CONTROL_ITEMS
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
        for key, label in source_items:
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
            for key, label in source_items
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


@app.route("/companies", methods=["GET"])
@admin_required
def company_management():
    company_rows = get_company_choices()
    return render_template("company_management.html", companies=company_rows)


@app.route("/companies/create", methods=["POST"])
@admin_required
def create_company():
    name = request.form.get("name", "").strip()
    address = request.form.get("address", "").strip()
    raw_slug = request.form.get("slug", "").strip()
    if not name or not address:
        flash("Firma adi ve adres gerekli.", "error")
        return redirect(url_for("company_management"))

    now = datetime.now().isoformat(timespec="seconds")
    try:
        with engine.begin() as connection:
            connection.execute(
                insert(companies).values(
                    public_id=uuid.uuid4().hex[:12],
                    slug=resolve_company_slug(connection, raw_slug, name),
                    name=name,
                    address=address,
                    contact_name="-",
                    created_at=now,
                    updated_at=now,
                )
            )
    except IntegrityError:
        flash("Bu firma zaten mevcut.", "error")
        return redirect(url_for("company_management"))

    flash("Cari kaydi olusturuldu.", "success")
    return redirect(url_for("company_management"))


@app.route("/companies/<int:company_id>/update", methods=["POST"])
@admin_required
def update_company(company_id: int):
    company = get_company(company_id)
    name = request.form.get("name", "").strip()
    address = request.form.get("address", "").strip()
    raw_slug = request.form.get("slug", "").strip()
    if not name or not address:
        flash("Firma adi ve adres gerekli.", "error")
        return redirect(url_for("company_management"))

    now = datetime.now().isoformat(timespec="seconds")
    try:
        with engine.begin() as connection:
            slug = resolve_company_slug(connection, raw_slug, name, exclude_id=company_id)
            connection.execute(
                update(companies)
                .where(companies.c.id == company_id)
                .values(
                    public_id=company.get("public_id") or uuid.uuid4().hex[:12],
                    slug=slug,
                    name=name,
                    address=address,
                    contact_name=company.get("contact_name") or "-",
                    updated_at=now,
                )
            )
            connection.execute(
                update(extinguishers)
                .where(extinguishers.c.company_id == company_id)
                .values(
                    company_name=name,
                    company_address=address,
                    updated_at=now,
                )
            )
    except IntegrityError:
        flash("Bu firma adi baska bir caride kullaniliyor.", "error")
        return redirect(url_for("company_management"))

    flash(f"{company['name']} guncellendi.", "success")
    return redirect(url_for("company_management"))


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
    company_choices = get_company_choices()
    asset_categories = get_asset_category_choices()
    if request.method == "POST":
        form = parse_required_form(request.form)
        form["technician_name"] = form.get("technician_name") or current_user_full_name()
        try:
            form, selected_company = sync_company_payload_from_selection(form)
        except ValueError as exc:
            flash(str(exc), "error")
            return render_template(
                "create_extinguisher.html",
                form=form,
                monthly_control_items=MONTHLY_CONTROL_ITEMS,
                equipment_options=EQUIPMENT_OPTIONS,
                equipment_presets=EQUIPMENT_PRESETS,
                companies=company_choices,
                asset_categories=asset_categories,
            )
        required_fields = {
            "serial_number": "Seri numarasi",
            "company_id": "Cari secimi",
            "asset_category": "Urun grubu",
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
                companies=company_choices,
                asset_categories=asset_categories,
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
                companies=company_choices,
                asset_categories=asset_categories,
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
                        company_id=selected_company["id"],
                        company_name=form["company_name"],
                        company_address=form["company_address"],
                        company_contact=form["company_contact"],
                        asset_category=form["asset_category"],
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
                companies=company_choices,
                asset_categories=asset_categories,
            )

        flash("Tup kaydedildi ve QR olusturuldu.", "success")
        return redirect(url_for("extinguisher_detail", public_id=public_id))

    return render_template(
        "create_extinguisher.html",
        form={
            "technician_name": current_user_full_name(),
            "asset_category": DEFAULT_ASSET_CATEGORY,
        },
        monthly_control_items=MONTHLY_CONTROL_ITEMS,
        equipment_options=EQUIPMENT_OPTIONS,
        equipment_presets=EQUIPMENT_PRESETS,
        companies=company_choices,
        asset_categories=asset_categories,
    )


@app.route("/records/new/yangin-elbisesi", methods=["GET", "POST"])
@login_required
def create_fire_suit():
    return render_profile_record_form("yangin-elbisesi")


@app.route("/records/new")
@login_required
def record_group_picker():
    return render_template(
        "record_group_picker.html",
        groups=get_registration_groups(),
    )


@app.route("/records/new/<group_slug>", methods=["GET", "POST"])
@login_required
def record_group_entry(group_slug: str):
    group = get_registration_group(group_slug)
    if group_slug == "elektrik-ic-tesisati":
        return create_electrical_installation()
    if group["status"] == "active":
        return render_profile_record_form(group_slug)
    return render_template(
        "record_group_placeholder.html",
        group=group,
    )


@app.route("/extinguishers/<public_id>")
@login_required
def extinguisher_detail(public_id: str):
    extinguisher = get_extinguisher(public_id)
    asset_profile = get_asset_profile(extinguisher.get("asset_category"))
    electrical_sections = (
        build_electrical_note_sections(extinguisher.get("notes"))
        if extinguisher.get("asset_category") == "Elektrik Ic Tesisati"
        else []
    )
    company_portal_url = None
    if extinguisher.get("company_id"):
        company = get_company(extinguisher["company_id"])
        company_portal_url = url_for(
            "public_company_portal",
            company_slug=company["slug"],
            _external=True,
        )
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
        ),
        asset_profile["monthly_control_items"],
    )
    return render_template(
        "extinguisher_detail.html",
        extinguisher=extinguisher,
        service_logs=service_history,
        monthly_inspections=monthly_history,
        monthly_control_items=asset_profile["monthly_control_items"],
        equipment_preset=get_equipment_preset(extinguisher.get("extinguisher_type")),
        company_portal_url=company_portal_url,
        asset_profile=asset_profile,
        electrical_sections=electrical_sections,
        can_delete=is_admin_user(),
    )


@app.route("/extinguishers/<public_id>/delete", methods=["POST"])
@admin_required
def delete_extinguisher(public_id: str):
    extinguisher = get_extinguisher(public_id)
    with engine.begin() as connection:
        connection.execute(
            service_logs.delete().where(service_logs.c.extinguisher_id == extinguisher["id"])
        )
        connection.execute(
            monthly_inspections.delete().where(monthly_inspections.c.extinguisher_id == extinguisher["id"])
        )
        connection.execute(
            extinguishers.delete().where(extinguishers.c.id == extinguisher["id"])
        )
    flash("Kayit silindi.", "success")
    return redirect(url_for("index"))


@app.route("/extinguishers/<public_id>/service/<int:log_id>/delete", methods=["POST"])
@admin_required
def delete_service_log(public_id: str, log_id: int):
    extinguisher = get_extinguisher(public_id)
    target_log = fetch_one(
        select(service_logs).where(
            service_logs.c.id == log_id,
            service_logs.c.extinguisher_id == extinguisher["id"],
        )
    )
    if target_log is None:
        abort(404)
    with engine.begin() as connection:
        connection.execute(
            service_logs.delete().where(service_logs.c.id == log_id)
        )
    flash("Bakim kaydi silindi.", "success")
    return redirect(url_for("extinguisher_detail", public_id=public_id))


@app.route("/extinguishers/<public_id>/monthly-inspection/<int:inspection_id>/delete", methods=["POST"])
@admin_required
def delete_monthly_inspection(public_id: str, inspection_id: int):
    extinguisher = get_extinguisher(public_id)
    target_inspection = fetch_one(
        select(monthly_inspections).where(
            monthly_inspections.c.id == inspection_id,
            monthly_inspections.c.extinguisher_id == extinguisher["id"],
        )
    )
    if target_inspection is None:
        abort(404)
    with engine.begin() as connection:
        connection.execute(
            monthly_inspections.delete().where(monthly_inspections.c.id == inspection_id)
        )
    flash("Aylik kontrol kaydi silindi.", "success")
    return redirect(url_for("extinguisher_detail", public_id=public_id))


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


@app.route("/firma/<company_slug>")
def public_company_portal(company_slug: str):
    company = get_company_by_slug(company_slug)
    sections = build_company_portal_sections(company["id"])
    return render_template(
        "public_company_portal.html",
        company=company,
        sections=sections,
        selected_category=None,
        selected_assets=[],
    )


@app.route("/firma/<company_slug>/<category_slug>")
def public_company_assets(company_slug: str, category_slug: str):
    company = get_company_by_slug(company_slug)
    selected_category = get_asset_category(slug=category_slug)
    if selected_category is None:
        abort(404)
    sections = build_company_portal_sections(company["id"])
    selected_section = next(
        (section for section in sections if section["slug"] == category_slug),
        None,
    )
    return render_template(
        "public_company_portal.html",
        company=company,
        sections=sections,
        selected_category=selected_category,
        selected_assets=selected_section["items"] if selected_section else [],
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
    company_choices = get_company_choices()
    asset_categories = get_asset_category_choices()
    asset_profile = get_asset_profile(extinguisher.get("asset_category"))
    if request.method == "POST":
        form = parse_required_form(request.form)
        form["technician_name"] = form.get("technician_name") or current_user_full_name()
        try:
            form, selected_company = sync_company_payload_from_selection(form)
        except ValueError as exc:
            flash(str(exc), "error")
            return render_template(
                "service_log_form.html",
                extinguisher=extinguisher,
                form=form,
                monthly_control_items=MONTHLY_CONTROL_ITEMS,
                equipment_options=EQUIPMENT_OPTIONS,
                equipment_presets=EQUIPMENT_PRESETS,
                companies=company_choices,
                asset_categories=asset_categories,
                asset_profile=asset_profile,
            )
        required_fields = {
            "service_date": asset_profile["last_service_label"],
            "next_service_date": asset_profile["next_service_label"],
            "technician_name": "Teknisyen",
            "company_id": "Cari secimi",
            "asset_category": "Urun grubu",
            "location_detail": "Bulundugu yer" if asset_profile["label"] != "Yangin Sondurme Cihazi" else "Firma ici konum",
            "fire_class": asset_profile["class_label"],
            "manufacturer": asset_profile["brand_label"],
            "operation_summary": "Yapilan islem",
            "company_contact": asset_profile["owner_label"],
        }
        if asset_profile["show_weight"]:
            required_fields["weight_kg"] = "Kg"
        if asset_profile["label"] == "Yangin Sondurme Cihazi":
            required_fields["extinguisher_type"] = "Tup tipi"
        if asset_profile["show_hydrostatic"]:
            required_fields["hydrostatic_test_date"] = "Hidrostatik test tarihi"
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
                companies=company_choices,
                asset_categories=asset_categories,
                asset_profile=asset_profile,
            )

        if asset_profile["show_weight"]:
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
                    companies=company_choices,
                    asset_categories=asset_categories,
                    asset_profile=asset_profile,
                )
        else:
            weight_kg = extinguisher.get("weight_kg") or 0.0

        now = datetime.now().isoformat(timespec="seconds")
        inspection_values = build_monthly_inspection_values(request.form, asset_profile["monthly_control_items"])
        control_values = build_control_form_values(request.form) if asset_profile["control_form_items"] else build_control_form_values({})
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
                    company_id=selected_company["id"],
                    company_name=form.get("company_name") or extinguisher["company_name"],
                    company_address=form.get("company_address") or extinguisher.get("company_address"),
                    company_contact=form.get("company_contact") or extinguisher.get("company_contact"),
                    asset_category=form.get("asset_category") or extinguisher.get("asset_category") or DEFAULT_ASSET_CATEGORY,
                    location_detail=form.get("location_detail")
                    or extinguisher["location_detail"],
                    weight_kg=weight_kg,
                    extinguisher_type=form.get("extinguisher_type")
                    or extinguisher["extinguisher_type"],
                    fire_class=form.get("fire_class") or extinguisher.get("fire_class"),
                    manufacturer=form.get("manufacturer") or extinguisher.get("manufacturer"),
                    hydrostatic_test_date=(form.get("hydrostatic_test_date") or extinguisher.get("hydrostatic_test_date")) if asset_profile["show_hydrostatic"] else None,
                    pressure_status=form.get("pressure_status") if asset_profile["show_pressure"] else None,
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
        form={
            "technician_name": current_user_full_name(),
            "company_id": str(extinguisher.get("company_id") or ""),
            "company_name": extinguisher.get("company_name") or "",
            "company_address": extinguisher.get("company_address") or "",
            "company_contact": extinguisher.get("company_contact") or "",
            "asset_category": extinguisher.get("asset_category") or DEFAULT_ASSET_CATEGORY,
        },
        monthly_control_items=MONTHLY_CONTROL_ITEMS,
        equipment_options=EQUIPMENT_OPTIONS,
        equipment_presets=EQUIPMENT_PRESETS,
        companies=company_choices,
        asset_categories=asset_categories,
        asset_profile=asset_profile,
    )


@app.route("/public/<public_id>")
def public_detail(public_id: str):
    extinguisher = get_extinguisher(public_id)
    asset_profile = get_asset_profile(extinguisher.get("asset_category"))
    electrical_sections = (
        build_electrical_note_sections(extinguisher.get("notes"))
        if extinguisher.get("asset_category") == "Elektrik Ic Tesisati"
        else []
    )
    company_portal_url = None
    if extinguisher.get("company_id"):
        company_portal_url = url_for(
            "public_company_portal",
            company_slug=get_company(extinguisher["company_id"])["slug"],
        )
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
        with_monthly_control_labels([latest_monthly_inspection_raw], asset_profile["monthly_control_items"])[0]
        if latest_monthly_inspection_raw
        else None
    )
    return render_template(
        "public_detail.html",
        extinguisher=extinguisher,
        latest_log=latest_log,
        latest_monthly_inspection=latest_monthly_inspection,
        equipment_preset=get_equipment_preset(extinguisher.get("extinguisher_type")),
        monthly_table=build_monthly_table(monthly_history_raw, asset_profile["monthly_control_items"]),
        company_portal_url=company_portal_url,
        asset_profile=asset_profile,
        electrical_sections=electrical_sections,
    )


@app.route("/extinguishers/<public_id>/monthly-inspection", methods=["GET", "POST"])
@login_required
def add_monthly_inspection(public_id: str):
    extinguisher = get_extinguisher(public_id)
    asset_profile = get_asset_profile(extinguisher.get("asset_category"))
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
                monthly_control_items=asset_profile["monthly_control_items"],
                control_form_items=asset_profile["control_form_items"],
                form=form,
                asset_profile=asset_profile,
            )

        now = datetime.now().isoformat(timespec="seconds")
        inspection_values = build_monthly_inspection_values(request.form, asset_profile["monthly_control_items"])
        control_values = build_control_form_values(request.form) if asset_profile["control_form_items"] else build_control_form_values({})
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
        monthly_control_items=asset_profile["monthly_control_items"],
        control_form_items=asset_profile["control_form_items"],
        form={"inspector_name": current_user_full_name()},
        asset_profile=asset_profile,
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
