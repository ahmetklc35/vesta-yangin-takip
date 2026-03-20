# Yangin Tupu Takip

Bu uygulama yangin tupu bakim kayitlarini tutar, her tup icin QR kod uretir ve
musterinin QR okutup son bakim bilgisine ulasmasini saglar.

## Ozellikler

- Ilk kayitta firma, konum, kg, tup tipi ve bakim bilgisi girme
- Her tup icin benzersiz QR kod uretme
- QR altinda `Vesta Yangin` markasi ile etiket hazirlama
- Sonraki bakimlarda yeni servis kaydi ekleme
- Musteri icin acik, sadece-okunur son bakim sayfasi
- Tup listesi ve bakim gecmisini Excel olarak disa aktarma
- Yerelde SQLite, yayinda PostgreSQL ile calisma

## Kurulum

```powershell
cd C:\Users\yunsu emre duman\Documents\Playground\yangin-tupu-takip
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Tarayicida `http://127.0.0.1:5000` adresini ac.

Uygulama `DATABASE_URL` yoksa otomatik olarak yerelde `SQLite` kullanir.

## Kolay Acilis

- `baslat.bat`: Var olan kurulumla uygulamayi direkt acar.
- `kurulum-ve-baslat.bat`: Gerekirse kutuphaneleri yukler, sonra uygulamayi acar.

## Kullanim

1. `Yeni Kayit` ile yangin tupu olustur.
2. Kayit sonrasinda detay sayfasinda QR gorunur.
3. `Etiket` sayfasindan baskiya hazir gorunum acilir.
4. `Bakim Guncelle` ile yeni servis kaydi ekle.
5. Ana ekrandan Excel listeleri indirebilirsin.
6. Musteri QR kodu okutunca `/public/<id>` sayfasina gider.

## Notlar

- Veritabani dosyasi `database.db` olarak proje klasorunde olusur.
- Uretim kullaniminda `SECRET_KEY` degerini degistir.
- Musterinin telefondan erisebilmesi icin uygulamanin sabit bir alan adinda veya
  yerel agda yayinlanmasi gerekir.

## Yayin ve PostgreSQL

- Canli ortamda `DATABASE_URL` bir PostgreSQL baglantisi olmali.
- `render.yaml` dosyasi Render uzerinde web servis + PostgreSQL olusturmak icin eklendi.
- Uygulama canlida `gunicorn wsgi:app` ile calisacak.
- Render `free` planinda web servis bos kalinca uyuyabilir. Musteri taramada ilk acilis biraz gecikebilir.

## Eski Veriyi PostgreSQL'e Tasima

Elindeki mevcut `SQLite` verisini yeni PostgreSQL veritabanina tasimak icin:

```powershell
cd C:\Users\yunsu emre duman\Documents\Playground\yangin-tupu-takip
$env:TARGET_DATABASE_URL="postgresql://kullanici:sifre@host:5432/veritabani"
.\.venv\Scripts\python migrate_sqlite_to_postgres.py
```

Render kullaniyorsan `TARGET_DATABASE_URL` yerine Render'in verdigi veritabani baglantisini yaz.
