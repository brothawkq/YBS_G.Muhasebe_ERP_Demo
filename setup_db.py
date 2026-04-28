import sqlite3
import os

def init_db():
    """
    YBS ERP Projesi - Veritabanı Mimari Yapılandırması
    Baran İlgün - 2026 Nihai Sürüm (Analitik & CRM Entegreli)
    """
    db_path = 'muhasebe.db'
    
    # Geliştirme aşamasında temiz kurulum için bağlantıyı açıyoruz
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    print("🔄 Tablolar yapılandırılıyor...")

    # --- 0. TEMİZLİK (Yeni eklenen 'customers' listeye dahil edildi) ---
    tables = ['users', 'accounts', 'transactions', 'journal_entries', 
              'audit_logs', 'assets', 'closing_steps', 'inventory_items', 
              'notifications', 'exchange_rates', 'stock_lots', 'invoices', 'customers']
    for table in tables:
        cursor.execute(f'DROP TABLE IF EXISTS {table}')

    # --- 1. KULLANICILAR VE YETKİLENDİRME (KORUNDU) ---
    cursor.execute('''CREATE TABLE IF NOT EXISTS users (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        username TEXT UNIQUE NOT NULL,
                        password_hash TEXT NOT NULL,
                        email TEXT,
                        full_name TEXT,       -- YENİ: Ad Soyad
                        department TEXT,      -- YENİ: Bölüm (Örn: YBS)
                        student_no TEXT,      -- YENİ: Öğrenci Numarası
                        birth_date DATE,      -- YENİ: Doğum Tarihi
                        bio TEXT,             -- YENİ: Kısa Özgeçmiş
                        profile_pic TEXT DEFAULT 'default.png',
                        role TEXT DEFAULT 'Admin',
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        verification_code TEXT,
                        is_verified INTEGER DEFAULT 0)''')

    # --- 2. HESAP PLANI (TDHP) (KORUNDU) ---
    cursor.execute('CREATE TABLE accounts (code REAL PRIMARY KEY, name TEXT)')
    
    # --- 3. İŞLEMLER (customer_id Eklendi - Eksiltme Yapılmadı) ---
    cursor.execute('''CREATE TABLE transactions (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        user_id INTEGER,
                        customer_id INTEGER,  -- YENİ: Cari Bağlantısı (CRM)
                        date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        description TEXT,
                        exchange_rate REAL DEFAULT 1.0,
                        currency TEXT DEFAULT 'TL',
                        source_module TEXT,
                        FOREIGN KEY(user_id) REFERENCES users(id),
                        FOREIGN KEY(customer_id) REFERENCES customers(id))''')

    # --- 4. YEVMİYE SATIRLARI (KORUNDU) ---
    cursor.execute('''CREATE TABLE journal_entries (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        transaction_id INTEGER,
                        account_code REAL,
                        debit REAL DEFAULT 0,
                        credit REAL DEFAULT 0,
                        entry_type TEXT DEFAULT 'Standart',
                        FOREIGN KEY(transaction_id) REFERENCES transactions(id))''')

    # --- 5. AUDIT LOGS (Denetim) (KORUNDU) ---
    cursor.execute('''CREATE TABLE audit_logs (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        user_id INTEGER,
                        action TEXT,
                        details TEXT,
                        ip_address TEXT,
                        timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        FOREIGN KEY(user_id) REFERENCES users(id))''')

    # --- 6. DURAN VARLIKLAR (KORUNDU) ---
    cursor.execute('''CREATE TABLE assets (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        user_id INTEGER,
                        name TEXT,
                        code REAL,
                        purchase_date DATE,
                        purchase_value REAL,
                        economic_life INTEGER,
                        dep_method TEXT DEFAULT 'Normal',
                        status TEXT DEFAULT 'Aktif',
                        accumulated_depreciation REAL DEFAULT 0,
                        FOREIGN KEY(user_id) REFERENCES users(id))''')

    # --- 7. İŞLEM SIRASI (Wizard) ---
    cursor.execute('''CREATE TABLE closing_steps (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        user_id INTEGER,
                        step_index INTEGER, 
                        is_completed INTEGER DEFAULT 0,
                        completion_date TIMESTAMP,
                        FOREIGN KEY(user_id) REFERENCES users(id),
                        UNIQUE(user_id, step_index))''')

    # --- 8. STOK VE ÜRÜN YÖNETİMİ (KORUNDU) ---
    cursor.execute('''CREATE TABLE inventory_items (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        user_id INTEGER,
                        sku TEXT UNIQUE,
                        product_name TEXT,
                        unit TEXT DEFAULT 'Adet',
                        quantity REAL DEFAULT 0,
                        unit_price REAL DEFAULT 0,
                        total_value REAL DEFAULT 0,
                        account_code REAL DEFAULT 153,
                        FOREIGN KEY(user_id) REFERENCES users(id))''')

    # --- 9. SİSTEM BİLDİRİMLERİ (KORUNDU) ---
    cursor.execute('''CREATE TABLE notifications (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        user_id INTEGER,
                        title TEXT,
                        message TEXT,
                        type TEXT DEFAULT 'info',
                        is_read INTEGER DEFAULT 0,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        FOREIGN KEY(user_id) REFERENCES users(id))''')

    # --- 10. CANLI DÖVİZ KURLARI (KORUNDU) ---
    cursor.execute('''CREATE TABLE IF NOT EXISTS exchange_rates (
                        id INTEGER PRIMARY KEY,
                        currency_code TEXT UNIQUE,
                        rate REAL,
                        last_update TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')

    # --- 11. FIFO STOK MALİYET TAKİBİ (KORUNDU) ---
    cursor.execute('''CREATE TABLE IF NOT EXISTS stock_lots (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        user_id INTEGER,
                        product_code REAL,
                        quantity REAL,
                        unit_price REAL,
                        purchase_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')

    # --- 12. OTOMATİK FATURA VE RESMİ BELGE TAKİBİ (KORUNDU) ---
    cursor.execute('''CREATE TABLE IF NOT EXISTS invoices (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        transaction_id INTEGER,
                        invoice_number TEXT UNIQUE,
                        customer_name TEXT,
                        tax_office TEXT,
                        grand_total REAL,
                        FOREIGN KEY(transaction_id) REFERENCES transactions(id))''')

    # --- 13. YENİ BÖLÜM: CARİ KARTLAR (CRM MODÜLÜ) ---
    cursor.execute('''CREATE TABLE IF NOT EXISTS customers (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        user_id INTEGER,
                        name TEXT NOT NULL,       -- Firma veya Şahıs Adı
                        tax_no TEXT,              -- VKN / TC Kimlik No
                        type TEXT,                -- 'Müşteri' veya 'Satıcı'
                        phone TEXT,               
                        email TEXT,               
                        address TEXT,             
                        current_balance REAL DEFAULT 0,
                        FOREIGN KEY(user_id) REFERENCES users(id))''')

    # --- GENİŞLETİLMİŞ TDHP VERİSİ (Müfredat PDF'leri tam kapsanıyor) ---
    tdhp = [
        # ── 1. DÖNEN VARLIKLAR ─────────────────────────────────────────────
        (100, '100 Kasa'),
        (101, '101 Alınan Çekler'),
        (102, '102 Bankalar'),
        (103, '103 Verilen Çekler ve Ödeme Emirleri (-)'),
        (108, '108 Diğer Hazır Değerler'),
        (110, '110 Hisse Senetleri'),
        (111, '111 Özel Kesim Tahvil Senet ve Bonoları'),
        (112, '112 Kamu Kesimi Tahvil Senet ve Bonoları'),
        (118, '118 Diğer Menkul Kıymetler'),
        (119, '119 Menkul Kıymetler Değer Düşüklüğü Karşılığı (-)'),
        (120, '120 Alıcılar'),
        (121, '121 Alacak Senetleri'),
        (122, '122 Alacak Senetleri Reeskontu (-)'),
        (126, '126 Verilen Depozito ve Teminatlar'),
        (128, '128 Şüpheli Ticari Alacaklar'),
        (129, '129 Şüpheli Ticari Alacaklar Karşılığı (-)'),
        (131, '131 Ortaklardan Alacaklar'),
        (150, '150 İlk Madde ve Malzeme'),
        (153, '153 Ticari Mallar'),
        (153.1, '153.01 A Malı Stok'),
        (153.2, '153.02 B Malı Stok'),
        (158, '158 Stok Değer Düşüklüğü Karşılığı (-)'),
        (159, '159 Verilen Sipariş Avansları'),
        (180, '180 Gelecek Aylara Ait Giderler'),
        (181, '181 Gelir Tahakkukları'),
        (190, '190 Devreden KDV'),
        (191, '191 İndirilecek KDV'),
        (193, '193 Peşin Ödenen Vergi ve Fonlar'),

        # ── 2. DURAN VARLIKLAR ─────────────────────────────────────────────
        # 22 Ticari Alacaklar (Uzun Vadeli)
        (220, '220 Alıcılar'),
        (221, '221 Alacak Senetleri'),
        (226, '226 Verilen Depozito ve Teminatlar'),
        # 24 Mali Duran Varlıklar
        (240, '240 Bağlı Menkul Kıymetler'),
        (242, '242 İştirakler'),
        (245, '245 Bağlı Ortaklıklar'),
        # 25 Maddi Duran Varlıklar (7. Duran Varlıklar.pdf)
        (250, '250 Arazi ve Arsalar'),
        (251, '251 Yeraltı ve Yerüstü Düzenleri'),
        (252, '252 Binalar'),
        (253, '253 Tesis, Makine ve Cihazlar'),
        (254, '254 Taşıtlar'),
        (255, '255 Demirbaşlar'),
        (257, '257 Birikmiş Amortismanlar (-)'),
        (258, '258 Yapılmakta Olan Yatırımlar'),
        (259, '259 Verilen Avanslar'),
        # 26 Maddi Olmayan Duran Varlıklar (7. Duran Varlıklar.pdf + 8. Amortisman.pdf)
        (260, '260 Haklar'),
        (261, '261 Şerefiye'),
        (262, '262 Kuruluş ve Örgütlenme Giderleri'),
        (263, '263 Araştırma ve Geliştirme Giderleri'),
        (264, '264 Özel Maliyetler'),
        (267, '267 Diğer Maddi Olmayan Duran Varlıklar'),
        (268, '268 Birikmiş Amortismanlar — Maddi Olmayan (-)'),
        (269, '269 Verilen Avanslar — Maddi Olmayan'),
        # 27 Özel Tükenmeye Tabi Varlıklar (7. Duran Varlıklar.pdf)
        (271, '271 Arama Giderleri'),
        (272, '272 Hazırlık ve Geliştirme Giderleri'),
        (277, '277 Diğer Özel Tükenmeye Tabi Varlıklar'),
        (278, '278 Birikmiş Tükenme Payları (-)'),
        (279, '279 Verilen Avanslar — Özel Tükenme'),
        # 28 Gelecek Yıllara Ait Giderler
        (280, '280 Gelecek Yıllara Ait Giderler'),
        (281, '281 Gelir Tahakkukları — Uzun Vadeli'),

        # ── 3. KISA VADELİ YABANCI KAYNAKLAR ──────────────────────────────
        (300, '300 Banka Kredileri'),
        (301, '301 Finansman Bonoları'),
        (305, '305 Çıkarılmış Bonolar ve Senetler'),
        (320, '320 Satıcılar'),
        (321, '321 Borç Senetleri'),
        (322, '322 Borç Senetleri Reeskontu (-)'),
        (326, '326 Alınan Depozito ve Teminatlar'),
        (331, '331 Ortaklara Borçlar'),
        (340, '340 Alınan Sipariş Avansları'),
        (360, '360 Ödenecek Vergi ve Fonlar'),
        (361, '361 Ödenecek Sosyal Güvenlik Kesintileri'),
        (368, '368 Vadesi Geçmiş Ertelenmiş veya Taksitlendirilmiş Vergi'),
        (370, '370 Dönem Kârı Vergi ve Diğer Yasal Yük. Karşılığı'),
        (380, '380 Gelecek Aylara Ait Gelirler'),
        (381, '381 Gider Tahakkukları'),
        (391, '391 Hesaplanan KDV'),
        (392, '392 Diğer KDV'),

        # ── 4. UZUN VADELİ YABANCI KAYNAKLAR (Müfredat Bilanço Şeması) ────
        (400, '400 Banka Kredileri — Uzun Vadeli'),
        (405, '405 Çıkarılmış Tahviller'),
        (420, '420 Satıcılar — Uzun Vadeli'),
        (421, '421 Borç Senetleri — Uzun Vadeli'),
        (431, '431 Ortaklara Borçlar — Uzun Vadeli'),
        (440, '440 Alınan Depozito ve Teminatlar — UV'),
        (472, '472 Kıdem Tazminatı Karşılığı'),
        (480, '480 Gelecek Yıllara Ait Gelirler'),
        (481, '481 Gider Tahakkukları — Uzun Vadeli'),

        # ── 5. ÖZ KAYNAKLAR ────────────────────────────────────────────────
        (500, '500 Sermaye'),
        (501, '501 Ödenmemiş Sermaye (-)'),
        (510, '510 Hisse Senedi İhraç Primleri'),
        (520, '520 Yasal Yedekler'),
        (521, '521 Statü Yedekleri'),
        (522, '522 Olağanüstü Yedekler'),
        (540, '540 Yasal Yedekler'),
        (570, '570 Geçmiş Yıllar Kârları'),
        (580, '580 Geçmiş Yıllar Zararları (-)'),
        (590, '590 Dönem Net Kârı'),
        (591, '591 Dönem Net Zararı (-)'),

        # ── 6. GELİR TABLOSU HESAPLARI ─────────────────────────────────────
        (600, '600 Yurt İçi Satışlar'),
        (601, '601 Yurt Dışı Satışlar'),
        (610, '610 Satıştan İadeler (-)'),
        (611, '611 Satış İskontoları (-)'),
        (612, '612 Diğer İndirimler (-)'),
        (620, '620 Satılan Mamullerin Maliyeti (-)'),
        (621, '621 Satılan Ticari Malların Maliyeti (-)'),
        (622, '622 Satılan Hizmet Maliyeti (-)'),
        (631, '631 Pazarlama Satış ve Dağıtım Giderleri (-)'),
        (632, '632 Genel Yönetim Giderleri (-)'),
        (640, '640 İştiraklerden Temettü Gelirleri'),
        (641, '641 Bağlı Ortaklıklardan Temettü Gelirleri'),
        (642, '642 Faiz Gelirleri'),
        (643, '643 Komisyon Gelirleri'),
        (644, '644 Konusu Kalmayan Karşılıklar'),
        (645, '645 Menkul Kıymet Satış Kârları'),
        (646, '646 Kambiyo Kârları'),
        (647, '647 Reeskont Faiz Gelirleri'),
        (649, '649 Diğer Olağan Gelir ve Kârlar'),
        (653, '653 Komisyon Giderleri (-)'),
        (654, '654 Karşılık Giderleri (-)'),
        (655, '655 Menkul Kıymet Satış Zararları (-)'),
        (656, '656 Kambiyo Zararları (-)'),
        (657, '657 Reeskont Faiz Giderleri (-)'),
        (659, '659 Diğer Olağan Gider ve Zararlar (-)'),
        (660, '660 Kısa Vadeli Borçlanma Giderleri (-)'),
        (661, '661 Uzun Vadeli Borçlanma Giderleri (-)'),
        (671, '671 Önceki Dönem Gelir ve Kârları'),
        (672, '672 Diğer Olağandışı Gelirler'),
        (679, '679 Diğer Olağandışı Gelir ve Kârlar'),
        (680, '680 Çalışmayan Kısım Gider ve Zararları (-)'),
        (681, '681 Önceki Dönem Gider ve Zararları (-)'),
        (689, '689 Diğer Olağandışı Gider ve Zararlar (-)'),
        (690, '690 Dönem Kârı veya Zararı'),
        (691, '691 Dönem Kârı Vergi ve Yasal Yük. Karşılığı (-)'),
        (692, '692 Dönem Net Kârı veya Zararı'),

        # ── 7. MALİYET HESAPLARI (7/A — Ticaret İşletmesi) ─────────────────
        (760, '760 Pazarlama Satış ve Dağıtım Giderleri'),
        (761, '761 Pazarlama Sat. Dağ. Gid. Yansıtma Hesabı'),
        (770, '770 Genel Yönetim Giderleri'),
        (771, '771 Genel Yönetim Giderleri Yansıtma Hesabı'),
        (780, '780 Finansman Giderleri'),
        (781, '781 Finansman Giderleri Yansıtma Hesabı'),

        # ── 9. NAZIM HESAPLAR (Muhasebeye Giriş Notu) ──────────────────────
        # Nazım hesaplar: varlık/kaynak niteliği taşımayan,
        # ilerde koşullara bağlı etki doğurabilecek kalemler
        (900, '900 Teminat Mektupları (Alınan)'),
        (901, '901 Teminat Mektupları Karşılığı'),
        (910, '910 Kefil Olunan Borçlar'),
        (911, '911 Kefil Olunan Borçlar Karşılığı'),
        (920, '920 Müşteri Adına Saklanan Kıymetler'),
        (921, '921 Müşteri Adına Saklanan Kıymetler Karşılığı'),
        (940, '940 Vadesi Geçmiş Alacaklar'),
        (941, '941 Vadesi Geçmiş Alacaklar Karşılığı'),
        (950, '950 İhracat Taahhütleri'),
        (951, '951 İhracat Taahhütleri Karşılığı'),
    ]

    cursor.executemany('INSERT INTO accounts (code, name) VALUES (?,?)', tdhp)
    
    conn.commit()
    conn.close()
    print("✅ YBS ERP — TDHP tam müfredat: 26x Maddi Olmayan, 27x Özel Tükenme, 4xx UVYK, 9xx Nazım dahil!")
    print(f"   → Toplam {len(tdhp)} hesap kodu yüklendi.")

if __name__ == "__main__":
    init_db()