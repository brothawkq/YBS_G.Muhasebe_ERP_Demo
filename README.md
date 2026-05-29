<div align="center">

# 🧾 YBS Genel Muhasebe ERP — Demo

**Türkiye Muhasebe Standartları'na uygun, web tabanlı ERP ve muhasebe yönetim sistemi**

![Python](https://img.shields.io/badge/Python-3.x-blue?style=flat-square&logo=python)
![Flask](https://img.shields.io/badge/Flask-Web%20Framework-black?style=flat-square&logo=flask)
![SQLite](https://img.shields.io/badge/Veritabanı-SQLite-003B57?style=flat-square&logo=sqlite)
![Platform](https://img.shields.io/badge/Platform-Web-lightgrey?style=flat-square)
![License](https://img.shields.io/badge/Kullanım-Demo-orange?style=flat-square)

</div>

---

## 📌 Proje Hakkında

YBS Genel Muhasebe ERP, Yönetim Bilişim Sistemleri (YBS) müfredatı kapsamında geliştirilmiş bir muhasebe otomasyon sistemidir. Türkiye Tek Düzen Hesap Planı (TDHP) ile tam uyumlu olup gerçek bir işletmenin tüm muhasebe döngüsünü dijital ortamda yönetmeyi sağlar.

**Geliştirici:** Baran İlgün

---

## ✨ Özellikler

### 📒 Temel Muhasebe

| Modül | Açıklama |
|-------|----------|
| **Yevmiye Defteri** | Tekli ve bileşik (3+ hesaplı) yevmiye kaydı, borç/alacak denge kontrolü |
| **Dövizli İşlem** | USD/EUR kur desteği, TCMB'den canlı kur çekme |
| **Sürekli Envanter** | Her satışta otomatik iki yevmiye: Satış + Maliyet (621/153) |
| **Bileşik Yevmiye** | Çok satırlı karmaşık kayıtlar (ör: 153+191 / 100+320) |
| **KDV Analizi** | 191 ve 391 hesap karşılaştırması, otomatik mahsup kaydı |
| **Dönem Sonu Kapanış** | 4 adımlı tam kapanış: Yansıtma → 690 Kapanış → Vergi → 590/591 |
| **TDHP** | 130+ hesap kodu, Nazım hesaplar dahil tam Tekdüzen Hesap Planı |

### 🏢 Varlık Yönetimi

| Modül | Açıklama |
|-------|----------|
| **Duran Varlık** | Demirbaş, taşıt, gayrimenkul kaydı ve takibi |
| **Amortisman** | Normal, Azalan Bakiyeler ve Kıst (binek otolar) yöntemleri |
| **Varlık Satışı** | Otomatik kar/zarar hesabı (679/689), net defter değeri analizi |
| **FIFO Stok** | Lot bazlı FIFO stok maliyet takibi |
| **Stok Modülü** | 153 Ticari Mallar yardımcı hesap dökümü ve envanter yönetimi |

### 📊 Raporlama ve Analiz

| Modül | Açıklama |
|-------|----------|
| **BI Dashboard** | Cari Oran, Likidite Oranı, Borç/Özkaynak rasyoları |
| **Gelir-Gider Analizi** | Dönemsel performans grafiği ve gelecek ay satış tahmini |
| **Excel Dışa Aktarım** | Mizan, Bilanço, Yevmiye, KDV, Duran Varlıklar, Stok → .xlsx |
| **PDF Fatura** | Yevmiye kaydından otomatik resmi fatura üretimi |
| **Nakit Akış Grafiği** | Haftalık kasa/banka hareketleri (100/102 hesapları) |

### 👥 CRM ve Kullanıcı Yönetimi

| Modül | Açıklama |
|-------|----------|
| **Cari Kart** | Müşteri ve satıcı kayıtları, VKN/TC no, bakiye takibi |
| **Avans/Depozito** | 126, 159, 326, 340 nolu hesap takibi |
| **Audit Log** | Tüm kullanıcı işlemlerinin denetim kaydı |
| **Kullanıcı Profili** | Profil fotoğrafı, bölüm, öğrenci numarası bilgileri |

### 🧭 Muhasebe Döngüsü Sihirbazı

12 adımlık muhasebe döngüsü yol göstericisi:

```
1. Açılış Bilançosu      7. Amortisman Kayıtları
2. Günlük Kayıtlar       8. Gelir-Gider Kapanışları
3. Maliyet Kayıtları     9. Kesin Mizan
4. Geçici Mizan         10. Gelir Tablosu
5. Envanter Değerleme   11. Bilanço
6. KDV Mahsubu          12. Kapanış Kayıtları
```

---

## 🛠 Teknoloji Yığını

| Katman | Teknoloji |
|--------|-----------|
| Dil | Python 3.x |
| Web Framework | Flask + Flask-Login |
| Veritabanı | SQLite |
| Raporlama | Pandas + XlsxWriter |
| PDF | fpdf |
| Döviz | TCMB XML API |
| Güvenlik | PBKDF2-SHA256 (Werkzeug), Audit Log, KVKK onayı |
| Frontend | Jinja2 + HTML/CSS |

---

## 🚀 Kurulum

```bash
# Repoyu klonla
git clone https://github.com/brothawkq/YBS_G.Muhasebe_ERP_Demo.git
cd YBS_G.Muhasebe_ERP_Demo

# Sanal ortam oluştur ve aktifleştir
python3 -m venv venv
source venv/bin/activate        # macOS / Linux
# venv\Scripts\activate         # Windows

# Bağımlılıkları kur
pip install -r requirements.txt

# Veritabanını başlat (TDHP yüklenir)
python setup_db.py

# Uygulamayı başlat
python app.py
```

Tarayıcıda `http://localhost:5000` adresini aç → Kayıt ol → Kullanmaya başla.

---

## 📁 Proje Yapısı

```
YBS_G.Muhasebe_ERP_Demo/
├── app.py              # Ana uygulama — tüm route ve iş mantığı
├── setup_db.py         # Veritabanı kurulumu ve TDHP yükleyici
├── muhasebe.db         # SQLite veritabanı (git'e gitmez)
├── requirements.txt    # Python bağımlılıkları
├── templates/          # Jinja2 HTML şablonları
│   ├── index.html          # Yevmiye kayıt merkezi
│   ├── dashboard.html      # BI Dashboard
│   ├── assets.html         # Duran varlık yönetimi
│   ├── inventory.html      # Stok ve envanter
│   ├── customers.html      # Cari kart (CRM)
│   ├── kdv_analiz.html     # KDV analiz ve mahsup
│   ├── kapanis.html        # Dönem sonu kapanış paneli
│   ├── wizard.html         # Muhasebe döngüsü sihirbazı
│   ├── avans_takip.html    # Avans/Depozito takibi
│   ├── login.html          # Giriş
│   └── register.html       # Kayıt
└── static/
    └── uploads/            # Profil fotoğrafları
```

---

## 📐 Hesap Planı (TDHP)

Uygulama, tam Tekdüzen Hesap Planı ile gelir. Temel gruplar:

| Grup | Kapsam |
|------|--------|
| **1xx** | Dönen Varlıklar (Kasa, Banka, Alacaklar, Stoklar, KDV) |
| **2xx** | Duran Varlıklar (Maddi, Maddi Olmayan, Özel Tükenme) |
| **3xx** | Kısa Vadeli Yabancı Kaynaklar |
| **4xx** | Uzun Vadeli Yabancı Kaynaklar |
| **5xx** | Öz Kaynaklar |
| **6xx** | Gelir Tablosu Hesapları |
| **7xx** | Maliyet Hesapları (7/A Yöntemi) |
| **9xx** | Nazım Hesaplar |

---

## ⚠️ Önemli Notlar

- Bu uygulama **demo ve eğitim amaçlıdır**
- Gerçek muhasebe veya vergi işlemleri için yetkili bir mali müşavir/muhasebeci ile çalışın
- Üretilen belgeler resmi nitelik taşımaz

---

<div align="center">
YBS Müfredat Projesi &nbsp;•&nbsp; Eğitim Amaçlıdır &nbsp;•&nbsp; Yatırım/Mali Tavsiye Değildir
</div>
