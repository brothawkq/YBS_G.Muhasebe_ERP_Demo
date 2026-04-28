"""
====================================================================
SAÜ YBS ERP - Sprint 1 Veritabanı Migration Script
====================================================================
Bu script MEVCUT muhasebe.db dosyasını korur ve sadece eksik şeyleri
ekler. setup_db.py'deki gibi DROP TABLE YAPMAZ.

Çalıştırma:  python3 migrate_db.py
====================================================================
"""
import sqlite3
import os
import sys

DB_PATH = 'muhasebe.db'

def migrate():
    if not os.path.exists(DB_PATH):
        print(f"⚠️  {DB_PATH} bulunamadı. İlk kurulum için 'python3 setup_db.py' çalıştırın.")
        sys.exit(1)
    
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    print("🔄 Migration başlatılıyor...")
    print("=" * 60)
    
    # === 1. SPRINT 1.4 - Source Module sütunu ===
    cursor.execute("PRAGMA table_info(transactions)")
    columns = [row['name'] for row in cursor.fetchall()]
    if 'source_module' not in columns:
        try:
            cursor.execute('ALTER TABLE transactions ADD COLUMN source_module TEXT')
            print("✅ transactions.source_module sütunu eklendi (Sürekli Envanter izleme için)")
        except sqlite3.OperationalError as e:
            print(f"⚠️  source_module ekleme hatası: {e}")
    else:
        print("✓ transactions.source_module zaten mevcut")
    
    # === 2. SPRINT 1.4 - Eksik Müfredat Hesapları ===
    # Müfredat eksikleri (önceki rapor K8 ve K10):
    mufredat_eksik_hesaplar = [
        # Komisyon ve Menkul Kıymet (4. Menkul Kıymetler.pdf)
        (111, '111 Özel Kesim Tahvil Senet ve Bonoları'),
        (112, '112 Kamu Kesimi Tahvil Senet ve Bonoları'),
        (118, '118 Diğer Menkul Kıymetler'),
        (119, '119 Menkul Kıymetler Değer Düşüklüğü Karşılığı (-)'),
        (122, '122 Alacak Senetleri Reeskontu (-)'),
        (127, '127 Diğer Ticari Alacaklar'),
        (128, '128 Şüpheli Ticari Alacaklar'),
        (129, '129 Şüpheli Ticari Alacaklar Karşılığı (-)'),
        (151, '151 Yarı Mamuller'),
        (152, '152 Mamuller'),
        (157, '157 Diğer Stoklar'),
        (158, '158 Stok Değer Düşüklüğü Karşılığı (-)'),
        (240, '240 Bağlı Menkul Kıymetler'),
        (242, '242 İştirakler'),
        (245, '245 Bağlı Ortaklıklar'),
        (251, '251 Yer Altı ve Yer Üstü Düzenleri'),
        (322, '322 Borç Senetleri Reeskontu (-)'),
        (329, '329 Diğer Ticari Borçlar'),
        (331, '331 Ortaklara Borçlar'),
        # Gelir Tablosu (9. PDF) - 7/A grubu yansıtma hesapları
        (611, '611 Satış İskontoları (-)'),
        (612, '612 Diğer İndirimler (-)'),
        (631, '631 Pazarlama Satış ve Dağıtım Giderleri'),
        (632, '632 Genel Yönetim Giderleri'),
        (645, '645 Menkul Kıymet Satış Karları'),
        (653, '653 Komisyon Giderleri (-)'),
        (654, '654 Karşılık Giderleri (-)'),
        (655, '655 Menkul Kıymet Satış Zararları (-)'),
        (660, '660 Kısa Vadeli Borçlanma Giderleri (-)'),
        (691, '691 Dönem Karı Vergi ve Diğer Yasal Yük. Karş. (-)'),
        (692, '692 Dönem Net Karı veya Zararı'),
        # 7/A grubu maliyet ve yansıtma hesapları (9. PDF)
        (761, '761 Pazarlama Sat. Dağ. Gid. Yansıtma Hs.'),
        (771, '771 Genel Yönetim Giderleri Yansıtma Hesabı'),
        (780, '780 Finansman Giderleri'),
        (781, '781 Finansman Giderleri Yansıtma Hesabı'),
    ]
    
    eklenen = 0
    var_olan = 0
    for kod, ad in mufredat_eksik_hesaplar:
        existing = cursor.execute('SELECT code FROM accounts WHERE code = ?', (kod,)).fetchone()
        if not existing:
            cursor.execute('INSERT INTO accounts (code, name) VALUES (?, ?)', (kod, ad))
            eklenen += 1
        else:
            var_olan += 1
    
    print(f"✅ Müfredat eksik hesapları: {eklenen} yeni eklendi, {var_olan} zaten mevcuttu.")
    
    # === 3. SPRINT 1 sonrası kullanıcılara bilgi ===
    cursor.execute('SELECT COUNT(*) as c FROM accounts')
    total_accounts = cursor.fetchone()['c']
    print(f"✅ TDHP toplam hesap sayısı: {total_accounts}")
    
    conn.commit()
    conn.close()
    
    print("=" * 60)
    print("✅ Migration tamamlandı! Sprint 1 değişiklikleri uygulandı.")
    print()
    print("🎯 YENİ ÖZELLİKLER:")
    print("  1. ✓ Azalan Bakiyeler amortismanı müfredata uygun (NDD×NAO×2)")
    print("  2. ✓ Bilanço'da kontra hesaplar (-) ile gösteriliyor")
    print("  3. ✓ Sürekli Envanter Yöntemi: 621/153 otomatik kayıt")
    print("  4. ✓ KDV oranı seçimli alış: %1, %10, %18, %20")
    print("  5. ✓ delete() yetkisiz erişim açığı kapatıldı")
    print("  6. ✓ amortisman idempotency: aşırı/tekrar amortisman engellendi")
    print()
    print("📌 TEST ETMEK İÇİN:")
    print("  - Yeni bir 30.000 TL'lik 'Azalan' yöntemli demirbaş ekle")
    print("  - Amortisman butonuna 5 kez bas → toplam 30.000 TL olmalı")
    print("  - Bilanço'ya git → 257 hesabı (kırmızı/parantez içinde) görünmeli")
    print("  - Index'te yeni iki form: KDV Otomatik Alış + Sürekli Env. Satış")

if __name__ == "__main__":
    migrate()
