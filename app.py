import os
import random
import sqlite3
import pandas as pd  # Yeni: YBS Veri Analitiği ve Excel Raporlama Modülü
import re            # YENİ: Şifre Karmaşıklığı (Regex)
import requests      # YENİ: Canlı Döviz API
import xml.etree.ElementTree as ET # YENİ: TCMB Veri İşleme
from fpdf import FPDF # YENİ: PDF Fatura Üretimi
from io import BytesIO
# session içe aktarıldı
from flask import Flask, render_template, request, redirect, url_for, flash, send_file, make_response, session 
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename

app = Flask(__name__)

# B3 GÜVENLİK: secret_key artık .env dosyasından veya os.urandom'dan alınır
# Production'da: SECRET_KEY=<random_string> içeren .env dosyası oluşturun
import os as _os
_SECRET_KEY = None
_env_path = _os.path.join(_os.path.dirname(__file__), '.env')
if _os.path.exists(_env_path):
    for _line in open(_env_path):
        _line = _line.strip()
        if _line.startswith('SECRET_KEY='):
            _SECRET_KEY = _line.split('=', 1)[1].strip().strip('"\'')
if not _SECRET_KEY:
    _SECRET_KEY = _os.environ.get('SECRET_KEY') or _os.urandom(32).hex()
app.secret_key = _SECRET_KEY

# --- YENİ EKLENDİ: PDF TÜRKÇE KARAKTER HATASI ÇÖZÜCÜ ---
def tr_fix(text):
    """ı, ş, ğ gibi karakterleri PDF kütüphanesinin (Latin-1) anlayacağı dile çevirir."""
    if not text: return ""
    search = "İığĞüÜşŞöÖçÇ"
    replace = "IigGuUsSoOcC"
    translation_table = str.maketrans(search, replace)
    return str(text).translate(translation_table)

# --- DOSYA YÜKLEME AYARLARI ---
# Kullanıcı profil resimleri için sunucuda klasör tanımlanır
UPLOAD_FOLDER = 'static/uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif'}

def allowed_file(filename):
    """Dosya uzantısının güvenli olup olmadığını kontrol eder"""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# --- LOGIN MANAGER VE USER SINIFI ---
login_manager = LoginManager(app)
login_manager.login_view = "login"

class User(UserMixin):
    """Kullanıcı oturum yönetimi için User sınıfı"""
    def __init__(self, id, username, profile_pic):
        self.id = id
        self.username = username
        self.profile_pic = profile_pic 

@login_manager.user_loader
def load_user(user_id):
    """Veritabanından kullanıcı bilgilerini yükler"""
    db = sqlite3.connect('muhasebe.db')
    u = db.execute('SELECT id, username, profile_pic FROM users WHERE id = ?', (user_id,)).fetchone()
    db.close()
    if u:
        return User(id=u[0], username=u[1], profile_pic=u[2])
    return None

def get_db():
    """SQLite veritabanı bağlantısı oluşturur"""
    conn = sqlite3.connect('muhasebe.db')
    conn.row_factory = sqlite3.Row
    return conn

def money(value, decimals=2):
    """
    B13 Para Birimi Yuvarlama Koruyucu (float yerine güvenli round)
    Müfredatta tüm tutarlar 2 ondalıkla gösterilir.
    0.1 + 0.2 = 0.30000000000000004 gibi float hatalarını önler.
    """
    try:
        return round(float(value or 0), decimals)
    except (TypeError, ValueError):
        return 0.0


def log_action(action, details, existing_cursor=None):
    """Siber Güvenlik ve İç Denetim için tüm eylemleri loglar (Audit Logs)"""
    if existing_cursor:
        # Mevcut transaction içinde kullan (DB kilitlenmesini önler)
        try:
            existing_cursor.execute('INSERT INTO audit_logs (user_id, action, details) VALUES (?, ?, ?)',
                                    (current_user.id, action, details))
        except Exception:
            pass  # Log hatası ana işlemi durdurmamalı
    else:
        db = get_db()
        db.execute('INSERT INTO audit_logs (user_id, action, details) VALUES (?, ?, ?)',
                   (current_user.id, action, details))
        db.commit()
        db.close()

# --- YENİ: GLOBALİZASYON VE DİL YÖNETİMİ ---
@app.route('/change_lang/<lang>')
def change_lang(lang):
    """Sistem dilini değiştirir ve kullanıcıyı mevcut sayfada tutar"""
    session['lang'] = lang
    return redirect(request.referrer or url_for('index'))

@app.before_request
def check_lang():
    """Her istekten önce varsayılan dilin 'tr' olduğunu doğrular"""
    if 'lang' not in session:
        session['lang'] = 'tr'

# --- YBS ANALİTİK YARDIMCI FONKSİYONU: SATIŞ TAHMİNLEME ---
def get_sales_prediction(user_id):
    """Gelecek ay beklenen ciro tahmini (Lineer Trend Analizi)"""
    db = get_db()
    # 600 nolu hesabın aylık toplamlarını çek
    sales = db.execute('''SELECT SUM(j.credit) as total FROM journal_entries j 
                          JOIN transactions t ON j.transaction_id = t.id 
                          WHERE t.user_id = ? AND j.account_code = 600 
                          GROUP BY strftime('%m', t.date) ORDER BY t.date DESC LIMIT 3''', (user_id,)).fetchall()
    db.close()
    data = [row['total'] for row in sales] if sales else [0]
    if len(data) < 2: return sum(data)
    growth_rate = (data[0] - data[-1]) / len(data)
    prediction = data[0] + growth_rate
    return round(prediction, 2)

# --- 1. MODÜL: KDV ANALİZ VE MAHSUP İŞLEMLERİ (6. KDV.pdf) ---
@app.route('/kdv_analiz')
@login_required
def kdv_analiz():
    """Ay sonu KDV hesaplarının karşılaştırılması (191 vs 391)"""
    db = get_db()
    
    # 191 İndirilecek KDV toplam bakiyesi (Borç - Alacak)
    ind_sql = "SELECT SUM(debit - credit) as bakiye FROM journal_entries j JOIN transactions t ON j.transaction_id = t.id WHERE t.user_id = ? AND account_code = 191"
    ind_kdv = db.execute(ind_sql, (current_user.id,)).fetchone()['bakiye'] or 0
    
    # 391 Hesaplanan KDV toplam bakiyesi (Alacak - Borç)
    hes_sql = "SELECT SUM(credit - debit) as bakiye FROM journal_entries j JOIN transactions t ON j.transaction_id = t.id WHERE t.user_id = ? AND account_code = 391"
    hes_kdv = db.execute(hes_sql, (current_user.id,)).fetchone()['bakiye'] or 0
    
    fark = ind_kdv - hes_kdv
    
    # KDV Mahsup kuralına göre durum tespiti
    if fark > 0:
        durum = "Devreden KDV (190)"
    else:
        durum = "Ödenecek Vergi (360)"
        
    db.close()
    return render_template('kdv_analiz.html', ind=ind_kdv, hes=hes_kdv, fark=abs(fark), durum=durum)

@app.route('/kdv_kapat', methods=['POST'])
@login_required
def kdv_kapat():
    """KDV mahsup yevmiye kaydını otomatik oluşturur (Ters Kayıt Mantığı)"""
    ind = float(request.form.get('ind', 0))
    hes = float(request.form.get('hes', 0))
    fark = ind - hes
    
    db = get_db()
    cursor = db.cursor()
    
    # Ana işlem kaydı oluşturma
    cursor.execute('INSERT INTO transactions (user_id, description) VALUES (?, ?)', 
                   (current_user.id, f"Ay Sonu KDV Mahsup Kaydı - Fark: {abs(fark)} TL"))
    t_id = cursor.lastrowid
    
    # 191 ve 391 hesapların ters kayıtla kapatılması (Mizanı sıfırlar)
    cursor.execute('INSERT INTO journal_entries (transaction_id, account_code, credit) VALUES (?, 191, ?)', (t_id, ind))
    cursor.execute('INSERT INTO journal_entries (transaction_id, account_code, debit) VALUES (?, 391, ?)', (t_id, hes))
    
    if fark > 0:
        # Aradaki fark işletme lehine ise 190 nolu hesaba aktarılır
        cursor.execute('INSERT INTO journal_entries (transaction_id, account_code, debit) VALUES (?, 190, ?)', (t_id, fark))
    else:
        # Aradaki fark maliye lehine ise 360 nolu hesaba aktarılır
        cursor.execute('INSERT INTO journal_entries (transaction_id, account_code, credit) VALUES (?, 360, ?)', (t_id, abs(fark)))
    
    db.commit()
    db.close()
    flash("KDV Mahsup İşlemi Başarıyla Tamamlandı!")
    return redirect(url_for('index'))

# --- 2 & 3. MODÜL: DURAN VARLIK VE AMORTİSMAN (7 & 8. PDF) ---
@app.route('/assets', methods=['GET', 'POST'])
@login_required
def assets():
    """Duran varlık listesi ve yeni varlık (Demirbaş) ekleme"""
    db = get_db()
    if request.method == 'POST':
        name = request.form.get('name')
        code = int(request.form.get('code'))
        val = float(request.form.get('value'))
        life = int(request.form.get('life'))
        date = request.form.get('date')
        method = request.form.get('method', 'Normal')
        
        db.execute('''INSERT INTO assets (user_id, name, code, purchase_date, purchase_value, economic_life, dep_method, status) 
                      VALUES (?,?,?,?,?,?,?,?)''',
                   (current_user.id, name, code, date, val, life, method, 'Aktif'))
        db.commit()
        flash(f"{name} isimli varlık sisteme eklendi.")
    
    my_assets = db.execute('SELECT * FROM assets WHERE user_id = ? AND status = "Aktif"', (current_user.id,)).fetchall()
    db.close()
    return render_template('assets.html', assets=my_assets)

@app.route('/calculate_depreciation/<int:asset_id>', methods=['POST'])
@login_required
def calculate_depreciation(asset_id):
    """
    AKADEMİK AMORTİSMAN HESAPLAMA (8. Amortisman.pdf MÜFREDATI)
    
    NORMAL YÖNTEM (Eşit tutarlı):
        Her yıl: AA = DD / EKÖ  (sabit tutar, tüm yıllarda aynı)
    
    AZALAN BAKİYELER YÖNTEMİ:
        Her yıl: AA = NDD × (NAO × 2)
        NDD = DD - BA (Net Defter Değeri = Defter Değeri - Birikmiş Amortisman)
        Son yıl: kalan tüm NDD amortismana atılır.
        
    Müfredat örneği (s. amortisman): 30.000 TL, 5 yıl ömür (NAO=%20)
        Yıl 1: 30.000 × %40 = 12.000  → BA: 12.000, NDD: 18.000
        Yıl 2: 18.000 × %40 = 7.200   → BA: 19.200, NDD: 10.800
        Yıl 3: 10.800 × %40 = 4.320   → BA: 23.520, NDD: 6.480
        ...
        Son yıl: kalan NDD tamamı amortismana atılır.
    """
    db = get_db()
    
    # GÜVENLİK: Ownership kontrolü (yetkisiz erişim önleme)
    asset = db.execute('SELECT * FROM assets WHERE id = ? AND user_id = ?', 
                       (asset_id, current_user.id)).fetchone()
    if not asset:
        db.close()
        flash("Varlık bulunamadı veya bu işlem için yetkiniz yok.")
        return redirect(url_for('assets'))
    
    # Güvenlik: Sadece Aktif varlıklara amortisman ayrılır
    if asset['status'] != 'Aktif':
        db.close()
        flash(f"{asset['name']} aktif değil, amortisman ayrılamaz.")
        return redirect(url_for('assets'))
    
    purchase_value = float(asset['purchase_value'])
    economic_life = int(asset['economic_life'])
    accumulated_so_far = float(asset['accumulated_depreciation'] or 0)
    
    # Amortisman tamamlanmış mı kontrolü (idempotency)
    if accumulated_so_far >= purchase_value - 0.01:
        db.close()
        flash(f"{asset['name']} için amortisman tamamlanmış. Yeni kayıt atılamaz.")
        return redirect(url_for('assets'))
    
    # Bu varlık için bugüne kadar kaç yıl amortisman ayrıldı? (yevmiye sayısından)
    year_count = db.execute('''SELECT COUNT(*) as c FROM transactions t 
                               JOIN journal_entries j ON t.id = j.transaction_id
                               WHERE t.user_id = ? AND j.account_code = 257 
                               AND t.description LIKE ?''',
                            (current_user.id, f"%{asset['name']}%Amortisman%")).fetchone()['c']
    current_year = year_count + 1  # Bu kaydedilecek olan yıl
    
    # Net Defter Değeri (NDD)
    ndd = purchase_value - accumulated_so_far

    # ─── KIST AMORTİSMAN KONTROLÜ (8. Amortisman.pdf Müfredatı) ────────────────
    # "Binek otolarda aktife girdiği ay önemlidir. Kıst Amortisman uygulanır."
    # Kural: İlk yıl = (Yıllık AA / 12) × (Alım ayından yıl sonuna kalan ay)
    #        Son yıl = Normal son yıl + (1. yılda ayrılmayan kalan kısım)
    # Kıst sadece: dep_method == 'Kıst' VEYA asset.code == 254 (Taşıtlar)
    is_kist = (asset['dep_method'] == 'Kıst') or (str(int(asset['code'] or 0)) == '254')
    kist_note = ""

    if is_kist and current_year == 1 and asset['purchase_date']:
        # Alım ayından yıl sonuna kaç ay kaldığını hesapla (ay kesri TAM sayılır)
        try:
            from datetime import datetime
            purchase_date = datetime.strptime(str(asset['purchase_date'])[:10], '%Y-%m-%d')
            purchase_month = purchase_date.month  # 1=Ocak ... 12=Aralık
            # Kasım=11 → 12 ay - 11 + 1 = 2 ay (Kasım + Aralık)
            kalan_ay = 12 - purchase_month + 1
        except Exception:
            kalan_ay = 12  # parse hatası → tam yıl uygula

        yillik_aa_normal = purchase_value / economic_life  # Normal yöntem için yıllık
        kist_aa = round((yillik_aa_normal / 12) * kalan_ay, 2)
        dep_amount = kist_aa
        kist_note = f" [KIST: {kalan_ay} ay × {yillik_aa_normal/12:.2f}]"
        method_note = f"Yıl 1/{economic_life} - KIST ({kalan_ay} ay)" + kist_note

    elif is_kist and current_year >= economic_life:
        # Son yıl: kalan NDD tamamını al (ilk yıl eksik kalan kısım burada telafi edilir)
        dep_amount = ndd
        method_note = f"Yıl {current_year}/{economic_life} - KIST SON YIL (kalan NDD)"
    # ─────────────────────────────────────────────────────────────────────────────
    elif asset['dep_method'] == 'Azalan':
        # AZALAN BAKİYELER YÖNTEMİ: NDD × (NAO×2)
        rate = (1.0 / economic_life) * 2.0  # NAO×2
        
        # SON YIL KURALI: Eğer ekonomik ömrün son yılındaysak, kalan tüm NDD'yi at
        if current_year >= economic_life:
            dep_amount = ndd
            method_note = f"Yıl {current_year}/{economic_life} - SON YIL (Tüm kalan NDD)"
        else:
            dep_amount = ndd * rate
            method_note = f"Yıl {current_year}/{economic_life} - NDD ({ndd:.2f}) × %{rate*100:.0f}"
    else:
        # NORMAL YÖNTEM: DD / EKÖ (her yıl sabit)
        dep_amount = purchase_value / economic_life
        method_note = f"Yıl {current_year}/{economic_life} - Eşit Pay"
        
        # Son yıl yuvarlama düzeltmesi
        if current_year >= economic_life:
            dep_amount = ndd  # Kalan tutarı al (yuvarlama hatası olmasın)
            method_note += " (Son yıl, kalan bakiye)"
    
    # Aşırı amortisman kontrolü (NDD'den fazla amortisman ayrılmasın)
    if dep_amount > ndd:
        dep_amount = ndd
    
    # Yuvarlama (2 ondalık)
    dep_amount = round(dep_amount, 2)
    
    if dep_amount <= 0:
        db.close()
        flash(f"{asset['name']} için ayrılacak amortisman tutarı sıfır.")
        return redirect(url_for('assets'))
        
    cursor = db.cursor()
    description = f"{asset['name']} ({asset['dep_method']}) Amortisman Gider Kaydı - {method_note}"
    cursor.execute('INSERT INTO transactions (user_id, description) VALUES (?, ?)', 
                   (current_user.id, description))
    t_id = cursor.lastrowid

    # MÜFREDAt: 8. Amortisman.pdf — hesap koduna göre birikmiş amortisman seçimi
    # 25x Maddi → 257 | 26x Maddi Olmayan → 268 | 27x Özel Tükenme → 278
    _prefix = str(int(float(asset['code'] or 0)))[:2]
    _birikim = 268 if _prefix == '26' else (278 if _prefix == '27' else 257)
    cursor.execute('INSERT INTO journal_entries (transaction_id, account_code, debit) VALUES (?, 770, ?)', (t_id, dep_amount))
    cursor.execute('INSERT INTO journal_entries (transaction_id, account_code, credit) VALUES (?, ?, ?)', (t_id, _birikim, dep_amount))
    
    # Varlık bazlı birikmiş amortisman güncellemesi
    cursor.execute('UPDATE assets SET accumulated_depreciation = accumulated_depreciation + ? WHERE id = ?', 
                   (dep_amount, asset_id))
    
    db.commit()
    db.close()
    flash(f"{asset['name']} - Yıl {current_year}: {dep_amount:,.2f} TL amortisman kaydedildi. (Yöntem: {asset['dep_method']})")
    return redirect(url_for('assets'))

@app.route('/sell_asset/<int:asset_id>', methods=['POST'])
@login_required
def sell_asset(asset_id):
    """Duran varlık satışı ve Satış Kar/Zarar hesabı (7. Duran Varlıklar.pdf)"""
    price = float(request.form.get('sell_price'))
    db = get_db()
    asset = db.execute('SELECT * FROM assets WHERE id = ?', (asset_id,)).fetchone()
    
    # BUG FIX: Birikmiş amortisman, varlığa özgü açıklama üzerinden filtreleniyor
    # Tüm kullanıcı bazındaki 257 yerine bu varlığa ait kayıtlar ayrıştırılıyor
    acc_sql = """SELECT SUM(j.credit - j.debit) as t 
                 FROM journal_entries j 
                 JOIN transactions t ON j.transaction_id = t.id 
                 WHERE t.user_id = ? AND j.account_code = ?
                 AND t.description LIKE ?"""
    _sp = str(int(float(asset['code'] or 0)))[:2]
    _bk = 268 if _sp == '26' else (278 if _sp == '27' else 257)
    accumulated = db.execute(acc_sql, (current_user.id, _bk, f"%{asset['name']}%")).fetchone()['t'] or 0
    accumulated = max(accumulated, asset['accumulated_depreciation'] or 0)
    
    net_value = asset['purchase_value'] - accumulated
    profit_loss = price - net_value
    
    cursor = db.cursor()
    cursor.execute('INSERT INTO transactions (user_id, description) VALUES (?, ?)', (current_user.id, f"{asset['name']} Satış İşlemi"))
    t_id = cursor.lastrowid
    
    # Satış yevmiye kaydı silsilesi: Kasa(B), Birikmiş Amortisman(B), Duran Varlık(A)
    cursor.execute('INSERT INTO journal_entries (transaction_id, account_code, debit) VALUES (?, 100, ?)', (t_id, price))
    cursor.execute('INSERT INTO journal_entries (transaction_id, account_code, debit) VALUES (?, _bk, ?)', (t_id, accumulated))
    cursor.execute('INSERT INTO journal_entries (transaction_id, account_code, credit) VALUES (?, ?, ?)', (t_id, asset['code'], asset['purchase_value']))
    
    if profit_loss > 0:
        # 679 Diğer Olağandışı Gelir ve Karlar
        cursor.execute('INSERT INTO journal_entries (transaction_id, account_code, credit) VALUES (?, 679, ?)', (t_id, profit_loss))
    else:
        # 689 Diğer Olağandışı Gider ve Zararlar
        cursor.execute('INSERT INTO journal_entries (transaction_id, account_code, debit) VALUES (?, 689, ?)', (t_id, abs(profit_loss)))
        
    db.execute('UPDATE assets SET status = "Satıldı" WHERE id = ?', (asset_id,))
    db.commit()
    db.close()
    flash("Varlık satışı ve kar/zarar analizi sisteme işlendi.")
    return redirect(url_for('assets'))

# --- 4. MODÜL: İŞ ZEKASI (BI) VE DASHBOARD ANALİZİ ---
@app.route('/dashboard')
@login_required
def dashboard():
    """YBS Karar Destek Sistemi: Finansal Rasyo (Cari Oran, Likidite) ve Grafik Analizleri"""
    db = get_db()
    
    # 1. Pasta Grafik Verisi: Aktif Varlık Dağılımı (1xx ve 2xx)
    a_sql = '''SELECT a.name, SUM(j.debit - j.credit) as bakiye FROM accounts a 
               JOIN journal_entries j ON a.code = j.account_code JOIN transactions t ON j.transaction_id = t.id 
               WHERE t.user_id = ? AND (a.code LIKE '1%' OR a.code LIKE '2%') GROUP BY a.code HAVING bakiye > 0'''
    assets_data = db.execute(a_sql, (current_user.id,)).fetchall()
    
    # 2. Rasyo Analiz Verileri (Dönen Varlıklar, KVYK, Stoklar)
    m_sql = '''SELECT a.code, (SUM(j.debit) - SUM(j.credit)) as bakiye FROM accounts a 
               JOIN journal_entries j ON a.code = j.account_code JOIN transactions t ON j.transaction_id = t.id 
               WHERE t.user_id = ? GROUP BY a.code'''
    mali_data = db.execute(m_sql, (current_user.id,)).fetchall()
    
    donen_v = sum(r['bakiye'] for r in mali_data if str(r['code']).startswith('1'))
    stoklar = sum(r['bakiye'] for r in mali_data if str(r['code']).startswith('153'))
    kvyk = abs(sum(r['bakiye'] for r in mali_data if str(r['code']).startswith('3')))
    oz_kaynak = abs(sum(r['bakiye'] for r in mali_data if str(r['code']).startswith('5')))
    
    # YBS Finansal Analiz Rasyoları
    cari_oran = round(donen_v / kvyk, 2) if kvyk > 0 else 0
    likidite = round((donen_v - stoklar) / kvyk, 2) if kvyk > 0 else 0
    borc_oran = round((kvyk) / oz_kaynak, 2) if oz_kaynak > 0 else 0
    
    # Performans Analizi (Gelir vs Gider)
    perf_sql = '''SELECT SUM(CASE WHEN a.code LIKE '6%' THEN (j.credit - j.debit) ELSE 0 END) as gelir,
                         SUM(CASE WHEN a.code LIKE '7%' THEN (j.debit - j.credit) ELSE 0 END) as gider
                         FROM accounts a JOIN journal_entries j ON a.code = j.account_code 
                         JOIN transactions t ON j.transaction_id = t.id WHERE t.user_id = ?'''
    perf_row = db.execute(perf_sql, (current_user.id,)).fetchone()
    
    # YENİ: Dashboard döviz kurları
    rates = db.execute('SELECT * FROM exchange_rates').fetchall()

    # YENİ: SATIŞ TAHMİNLEME VE CRM İSTATİSTİKLERİ
    sales_prediction = get_sales_prediction(current_user.id)
    cust_count = db.execute('SELECT COUNT(*) as c FROM customers WHERE user_id = ?', (current_user.id,)).fetchone()['c']
    
    # BUG FIX: Nakit akış grafiği için gerçek haftalık veri (100+102 hesapları)
    cash_flow_sql = '''SELECT strftime('%w', t.date) as gun, SUM(j.debit) as giris
                       FROM journal_entries j JOIN transactions t ON j.transaction_id = t.id
                       WHERE t.user_id = ? AND j.account_code IN (100, 102)
                       AND t.date >= date('now', '-7 days')
                       GROUP BY strftime('%w', t.date)'''
    cash_rows = db.execute(cash_flow_sql, (current_user.id,)).fetchall()
    # Gün bazlı dict oluştur (0=Paz, 1=Pzt ... 6=Cmt)
    cash_map = {str(r['gun']): r['giris'] for r in cash_rows}
    # Haftalık sıra: Pzt(1) → Paz(0)
    cash_data = [cash_map.get(str(i), 0) for i in range(1, 7)] + [cash_map.get('0', 0)]
    
    db.close()
    return render_template('dashboard.html', assets=assets_data, perf=perf_row, 
                           cari=cari_oran, likidite=likidite, borc=borc_oran, 
                           rates=rates, prediction=sales_prediction, cust_count=cust_count,
                           cash_data=cash_data)

# --- YENİ: CANLI DÖVİZ KURU SERVİSİ ---
@app.route('/refresh_rates')
@login_required
def refresh_rates():
    """TCMB üzerinden canlı kurları çeker ve veritabanına mühürler"""
    try:
        r = requests.get("https://www.tcmb.gov.tr/kurlar/today.xml", timeout=5)
        root = ET.fromstring(r.content)
        db = get_db()
        for currency in root.findall('Currency'):
            code = currency.get('CurrencyCode')
            if code in ['USD', 'EUR']:
                rate = float(currency.find('ForexBuying').text)
                db.execute('INSERT OR REPLACE INTO exchange_rates (currency_code, rate) VALUES (?, ?)', (code, rate))
        db.commit()
        db.close()
        flash("Canlı Piyasa Verileri Güncellendi.")
    except Exception as e:
        flash(f"Kur Servis Hatası: {str(e)}")
    return redirect(url_for('dashboard'))

# --- 5. MODÜL: EXCEL RAPORLAMA (Pandas Entegrasyonu) ---
@app.route('/export_excel/<string:report_type>')
@login_required
def export_excel(report_type):
    """Mali tabloları (Mizan, Yevmiye vb.) profesyonel Excel formatında dışa aktarır"""
    db = get_db()
    sheet_name = 'YBS_Rapor'
    
    if report_type == "mizan":
        query = '''SELECT a.code as 'Hesap Kodu', a.name as 'Hesap Adı', 
                   SUM(j.debit) as 'Borç Toplamı', SUM(j.credit) as 'Alacak Toplamı', 
                   (SUM(j.debit) - SUM(j.credit)) as 'Bakiye' 
                   FROM accounts a JOIN journal_entries j ON a.code = j.account_code 
                   WHERE j.transaction_id IN (SELECT id FROM transactions WHERE user_id = ?) 
                   GROUP BY a.code'''
        sheet_name = 'Mizan'
    elif report_type == "bilanco":
        query = '''SELECT a.code as 'Hesap Kodu', a.name as 'Hesap Adı',
                   (SUM(j.debit) - SUM(j.credit)) as 'Bakiye'
                   FROM accounts a JOIN journal_entries j ON a.code = j.account_code
                   JOIN transactions t ON j.transaction_id = t.id
                   WHERE t.user_id = ? GROUP BY a.code ORDER BY a.code'''
        sheet_name = 'Bilanco'
    elif report_type == "kdv":
        query = '''SELECT a.code as 'KDV Hesabı', a.name as 'Açıklama',
                   SUM(j.debit) as 'Borç', SUM(j.credit) as 'Alacak',
                   (SUM(j.debit) - SUM(j.credit)) as 'Net Bakiye'
                   FROM accounts a JOIN journal_entries j ON a.code = j.account_code
                   JOIN transactions t ON j.transaction_id = t.id
                   WHERE t.user_id = ? AND a.code IN (191, 391, 190, 360) GROUP BY a.code'''
        sheet_name = 'KDV_Analiz'
    elif report_type == "assets":
        query = '''SELECT name as 'Varlık Adı', code as 'Hesap Kodu', purchase_date as 'Alış Tarihi',
                   purchase_value as 'Maliyet Bedeli', economic_life as 'Ömür (Yıl)',
                   dep_method as 'Amortisman Yöntemi', accumulated_depreciation as 'Birikmiş Amortisman',
                   status as 'Durum'
                   FROM assets WHERE user_id = ?'''
        sheet_name = 'Duran_Varliklar'
    elif report_type == "inventory":
        query = '''SELECT a.code as 'Hesap Kodu', a.name as 'Malzeme Adı',
                   SUM(j.debit - j.credit) as 'Stok Değeri (TL)'
                   FROM accounts a JOIN journal_entries j ON a.code = j.account_code
                   JOIN transactions t ON j.transaction_id = t.id
                   WHERE t.user_id = ? AND CAST(a.code as TEXT) LIKE '153%' GROUP BY a.code'''
        sheet_name = 'Stok_Envanter'
    elif report_type == "avans":
        query = '''SELECT a.code as 'Hesap Kodu', a.name as 'Hesap Adı',
                   (SUM(j.debit) - SUM(j.credit)) as 'Bakiye'
                   FROM accounts a JOIN journal_entries j ON a.code = j.account_code
                   JOIN transactions t ON j.transaction_id = t.id
                   WHERE t.user_id = ? AND a.code IN (126, 159, 326, 340) GROUP BY a.code'''
        sheet_name = 'Avans_Depozito'
    else:
        # Varsayılan: Yevmiye Defteri
        query = '''SELECT t.date as 'Tarih', t.description as 'Açıklama', 
                   j.account_code as 'Hesap Kodu', j.debit as 'Borç', j.credit as 'Alacak'
                   FROM transactions t JOIN journal_entries j ON t.id = j.transaction_id 
                   WHERE t.user_id = ? ORDER BY t.id DESC'''
        sheet_name = 'Yevmiye_Defteri'
        
    df = pd.read_sql_query(query, db, params=(current_user.id,))
    db.close()
    
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        # Sütun genişliklerini otomatik ayarla
        worksheet = writer.sheets[sheet_name]
        for i, col in enumerate(df.columns):
            # BUG FIX: Boş DF'de max() NaN döner → int() crash önlenir
            col_max = df[col].astype(str).map(len).max() if not df.empty else 0
            max_len = max(int(col_max) if col_max == col_max else 0, len(str(col))) + 2
            worksheet.set_column(i, i, min(max_len, 40))
    output.seek(0)
    
    return send_file(output, download_name=f"YBS_ERP_{report_type}_{sheet_name}.xlsx", as_attachment=True)

# --- 6. MODÜL: STOK VE ENVANTER TAKİBİ (5. Stoklar.pdf) ---
@app.route('/inventory')
@login_required
def inventory():
    """153 Ticari Mallar hesabının yardımcı hesap dökümü ve FIFO analizi"""
    db = get_db()
    query = '''SELECT a.code, a.name, SUM(j.debit - j.credit) as bakiye FROM accounts a 
               JOIN journal_entries j ON a.code = j.account_code JOIN transactions t ON j.transaction_id = t.id 
               WHERE t.user_id = ? AND a.code LIKE '153%' GROUP BY a.code'''
    stocks = db.execute(query, (current_user.id,)).fetchall()
    # FIFO lotlarını çekiyoruz
    lots = db.execute('SELECT * FROM stock_lots WHERE user_id = ? ORDER BY purchase_date ASC', (current_user.id,)).fetchall()
    db.close()
    return render_template('inventory.html', stocks=stocks, lots=lots)

# YENİ: FIFO STOK GİRİŞİ
@app.route('/add_stock', methods=['POST'])
@login_required
def add_stock():
    db = get_db()
    db.execute('INSERT INTO stock_lots (user_id, product_code, quantity, unit_price) VALUES (?,?,?,?)',
               (current_user.id, request.form.get('code'), request.form.get('qty'), request.form.get('price')))
    db.commit()
    db.close()
    flash("Stok girişi FIFO lotu olarak kaydedildi.")
    return redirect(url_for('inventory'))

# --- YENİ: AVANS VE DEPOZİTO TAKİBİ ---
@app.route('/avans_takip')
@login_required
def avans_takip():
    db = get_db()
    sql = '''SELECT a.code, a.name, (SUM(j.debit) - SUM(j.credit)) as bakiye FROM accounts a 
             JOIN journal_entries j ON a.code = j.account_code JOIN transactions t ON j.transaction_id = t.id 
             WHERE t.user_id = ? AND a.code IN (126, 159, 326, 340) GROUP BY a.code'''
    data = db.execute(sql, (current_user.id,)).fetchall()
    db.close()
    return render_template('avans_takip.html', data=data)

# --- YENİ: CARİ KART YÖNETİMİ (CRM) ---
@app.route('/customers', methods=['GET', 'POST'])
@login_required
def customers():
    """CRM: Müşteri ve Satıcı Kartları Yönetimi"""
    db = get_db()
    if request.method == 'POST':
        db.execute('''INSERT INTO customers (user_id, name, tax_no, type, phone) 
                      VALUES (?,?,?,?,?)''', (current_user.id, request.form['name'], 
                      request.form['tax_no'], request.form['type'], request.form['phone']))
        db.commit()
        flash("Yeni Cari Kart başarıyla mühürlendi.")
    
    customer_list = db.execute('SELECT * FROM customers WHERE user_id = ?', (current_user.id,)).fetchall()
    db.close()
    return render_template('customers.html', customers=customer_list)

@app.route('/customers/edit', methods=['POST'])
@login_required
def edit_customer():
    """Cari kart güncelleme"""
    db = get_db()
    db.execute('''UPDATE customers SET name=?, type=?, tax_no=?, phone=?, email=?, address=?
                  WHERE id=? AND user_id=?''',
               (request.form['name'], request.form['type'], request.form.get('tax_no'),
                request.form.get('phone'), request.form.get('email'), request.form.get('address'),
                request.form['customer_id'], current_user.id))
    db.commit()
    db.close()
    log_action("Düzenleme", f"Cari kart güncellendi: {request.form['name']}")
    flash("Cari kart başarıyla güncellendi.")
    return redirect(url_for('customers'))

@app.route('/customers/delete', methods=['POST'])
@login_required
def delete_customer():
    """Cari kart silme"""
    db = get_db()
    cari = db.execute('SELECT name FROM customers WHERE id=? AND user_id=?', 
                      (request.form['customer_id'], current_user.id)).fetchone()
    db.execute('DELETE FROM customers WHERE id=? AND user_id=?', 
               (request.form['customer_id'], current_user.id))
    db.commit()
    db.close()
    if cari:
        log_action("Silme", f"Cari kart silindi: {cari['name']}")
    flash("Cari kart başarıyla silindi.")
    return redirect(url_for('customers'))

# --- 7. MODÜL: DÖNEM SONU KAPANIŞ (9. Gelir Tablosu.pdf) ---
@app.route('/kapanis_paneli')
@login_required
def kapanis_paneli():
    """
    Kapanış öncesi ön izleme paneli.
    9. Gelir Tablosu PDF müfredatına göre hem 6'lı hem 7'li hesapları listeler.
    """
    db = get_db()

    # 6'lı Gelir/Gider hesapları (net bakiye = Alacak - Borç)
    q6 = """SELECT a.code, a.name,
                   SUM(j.credit - j.debit) as bakiye,
                   SUM(j.credit - j.debit) as b
            FROM accounts a
            JOIN journal_entries j ON a.code = j.account_code
            JOIN transactions t ON j.transaction_id = t.id
            WHERE t.user_id = ? AND CAST(a.code AS TEXT) LIKE '6%'
            GROUP BY a.code HAVING bakiye != 0"""
    data6 = db.execute(q6, (current_user.id,)).fetchall()

    # 7'li Maliyet hesapları (Borç - Alacak → borç bakiyeli, pozitif = gider)
    q7 = """SELECT a.code, a.name,
                   SUM(j.debit - j.credit) as bakiye,
                   SUM(j.debit - j.credit) as b
            FROM accounts a
            JOIN journal_entries j ON a.code = j.account_code
            JOIN transactions t ON j.transaction_id = t.id
            WHERE t.user_id = ? AND CAST(a.code AS TEXT) LIKE '7%'
            AND CAST(a.code AS TEXT) NOT LIKE '76%'  -- yansıtma değil
            AND CAST(a.code AS TEXT) NOT LIKE '77%'
            AND CAST(a.code AS TEXT) NOT LIKE '78%'
            GROUP BY a.code HAVING bakiye != 0"""
    data7 = db.execute(q7, (current_user.id,)).fetchall()

    db.close()
    return render_template('kapanis.html', data=data6, data7=data7)


@app.route('/close_period', methods=['POST'])
@login_required
def close_period():
    """
    MÜFREDAta UYGUN TAM KAPANIŞ ZİNCİRİ (9. Gelir Tablosu PDF)

    7/A Yöntemi kapanış adımları:
    ─────────────────────────────────────────────────────────────
    ADIM 1 — Yansıtma (dönem sonu maliyet → 6xx yansıtma):
        761(B) / 760(A)   Pazarlama giderleri yansıtma
        771(B) / 770(A)   Genel yönetim giderleri yansıtma
        781(B) / 780(A)   Finansman giderleri yansıtma

    ADIM 2 — Yansıtma hesapları ile 6xx gider hesaplarını eşleştir:
        631(B) / 761(A)   Paz. gid. → 6'ya taşı
        632(B) / 771(A)   Gen.yön.gid. → 6'ya taşı
        660(B) / 781(A)   Fin. gid. → 6'ya taşı

    ADIM 3 — Tüm 6'lı hesapları 690'a kapat:
        Gelirler (bakiye > 0): 6xx (B) / 690 (A)
        Giderler (bakiye < 0): 690 (B) / 6xx (A)

    ADIM 4 — Vergi ve Net Kâr/Zarar:
        KÂR durumunda:
            691 (B) / 360 (A)    Kurumlar Vergisi (%22)
            690 (B) / 691 (A)    Vergiyi 690'dan düş
            690 (B) / 692 (A)    Net kârı 692'ye aktar
            692 (B) / 590 (A)    Net kârı özkaynaklara taşı

        ZARAR durumunda:
            690 (B) / 692 (A)    Zararı 692'ye aktar
            591 (B) / 692 (A)    Net zararı özkaynaklara taşı
    ─────────────────────────────────────────────────────────────
    """
    db = get_db()
    cursor = db.cursor()
    uid = current_user.id

    def get_balance(code, balance_type='credit_minus_debit'):
        """Belirli hesabın net bakiyesini döndür — integer kod ile çalışır"""
        if balance_type == 'credit_minus_debit':
            expr = 'SUM(j.credit - j.debit)'
        else:
            expr = 'SUM(j.debit - j.credit)'
        # accounts.code REAL olarak saklandığı için doğrudan sayı ile karşılaştır
        sql = f"""SELECT {expr} as bal FROM journal_entries j
                  JOIN transactions t ON j.transaction_id = t.id
                  WHERE t.user_id = ? AND j.account_code = ?"""
        result = db.execute(sql, (uid, code)).fetchone()
        return float(result['bal'] or 0)

    def ins(t_id, code, side, amount):
        """Yevmiye satırı ekle"""
        if side == 'D':
            cursor.execute('INSERT INTO journal_entries (transaction_id, account_code, debit) VALUES (?,?,?)',
                           (t_id, code, round(abs(amount), 2)))
        else:
            cursor.execute('INSERT INTO journal_entries (transaction_id, account_code, credit) VALUES (?,?,?)',
                           (t_id, code, round(abs(amount), 2)))

    def new_tx(desc, module='Dönem Sonu Kapanış'):
        cursor.execute('INSERT INTO transactions (user_id, description, source_module) VALUES (?,?,?)',
                       (uid, desc, module))
        return cursor.lastrowid

    # ─────────────────────────────────────────────────────────────
    # ADIM 1: YANSITMA — 7'ler → Yansıtma hesapları (7x1 borçlu, 7x0 alacaklı)
    # ─────────────────────────────────────────────────────────────
    # Yansıtma çiftleri: (kaynak_maliyet, yansıtma, hedef_6_kodu)
    YANSITMA_CIFTLERI = [
        (760, 761, 631),  # Pazarlama
        (770, 771, 632),  # Genel Yönetim
        (780, 781, 660),  # Finansman
    ]

    yansitma_yapildi = False
    for kaynak, yansitma, hedef6 in YANSITMA_CIFTLERI:
        # 7x0 hesabının borç bakiyesi (debit - credit)
        kaynak_bakiye = get_balance(f'{kaynak}', 'debit_minus_credit')
        if abs(kaynak_bakiye) < 0.01:
            continue

        yansitma_yapildi = True

        # ADIM 1a: Yansıtma kaydı — 7x1(B) / 7x0(A)
        t = new_tx(f'Yansıtma: {kaynak} → {yansitma}')
        ins(t, yansitma, 'D', kaynak_bakiye)   # 761/771/781 Borçlanır
        ins(t, kaynak, 'C', kaynak_bakiye)      # 760/770/780 Alacaklanır

        # ADIM 1b: Yansıtma → 6'ya transfer — 6xx(B) / 7x1(A)
        t2 = new_tx(f'6\'ya transfer: {yansitma} → {hedef6}')
        ins(t2, hedef6, 'D', kaynak_bakiye)    # 631/632/660 Borçlanır
        ins(t2, yansitma, 'C', kaynak_bakiye)  # 761/771/781 Alacaklanır

    # ─────────────────────────────────────────────────────────────
    # ADIM 2: 6'lı hesapları 690'a kapat
    # ─────────────────────────────────────────────────────────────
    balances_6 = db.execute("""
        SELECT a.code, a.name, SUM(j.credit - j.debit) as bakiye
        FROM accounts a
        JOIN journal_entries j ON a.code = j.account_code
        JOIN transactions t ON j.transaction_id = t.id
        WHERE t.user_id = ? AND CAST(a.code AS TEXT) LIKE '6%'
        GROUP BY a.code HAVING ABS(bakiye) > 0.005
    """, (uid,)).fetchall()

    if not balances_6:
        db.close()
        flash("Kapatılacak 6'lı hesap bulunamadı. Önce gelir/gider kayıtları yapın.")
        return redirect(url_for('kapanis_paneli'))

    t_kapat = new_tx('Dönem Sonu: 6\'lı Hesaplar → 690 Kapanış')
    net_690 = 0.0

    for row in balances_6:
        b = float(row['bakiye'])
        if abs(b) < 0.005:
            continue
        net_690 += b
        if b > 0:  # Gelir hesabı: borçlandırarak kapat
            ins(t_kapat, row['code'], 'D', b)
        else:       # Gider hesabı: alacaklandırarak kapat
            ins(t_kapat, row['code'], 'C', abs(b))

    # 690'a karşı kayıt
    if net_690 > 0:
        ins(t_kapat, 690, 'C', net_690)   # Net kâr → 690 alacaklı
    else:
        ins(t_kapat, 690, 'D', abs(net_690))  # Net zarar → 690 borçlu

    # ─────────────────────────────────────────────────────────────
    # ADIM 3: Vergi + Net Kâr/Zarar devri
    # ─────────────────────────────────────────────────────────────
    VERGI_ORANI = 0.22  # Kurumlar Vergisi %22 (müfredat)

    if net_690 > 0:
        # ── KÂR durumu ──────────────────────────────────────────
        vergi = round(net_690 * VERGI_ORANI, 2)
        net_kar = round(net_690 - vergi, 2)

        # ADIM 3a: Vergi karşılığı — 691(B) / 360(A)
        t_vergi = new_tx('Kurumlar Vergisi Karşılığı (%22) — 691/360')
        ins(t_vergi, 691, 'D', vergi)
        ins(t_vergi, 360, 'C', vergi)

        # ADIM 3b: 690 → 691 → 692 kapanışı
        t_net = new_tx('Dönem Net Kârı Devri — 690→691→692→590')
        ins(t_net, 690, 'D', net_690)    # 690 borçlanır (kapatılır)
        ins(t_net, 691, 'C', vergi)      # 691 alacaklanır (kapatılır)
        ins(t_net, 692, 'C', net_kar)    # 692 Net Kâr alacaklanır

        # ADIM 3c: 692 → 590 (Öz Kaynaklara)
        t_590 = new_tx('Net Kâr Öz Kaynaklara — 692→590')
        ins(t_590, 692, 'D', net_kar)
        ins(t_590, 590, 'C', net_kar)

        flash_msg = (f"✅ Dönem kapanışı tamamlandı | "
                     f"Dönem Kârı: {net_690:,.2f} TL | "
                     f"Vergi (%22): {vergi:,.2f} TL | "
                     f"Net Kâr (590): {net_kar:,.2f} TL")
    else:
        # ── ZARAR durumu ─────────────────────────────────────────
        net_zarar = abs(net_690)

        # ADIM 3a: 690 → 692 kapanışı
        t_net = new_tx('Dönem Net Zararı Devri — 690→692→591')
        ins(t_net, 692, 'D', net_zarar)   # 692 Net Zarar borçlanır
        ins(t_net, 690, 'C', net_zarar)   # 690 alacaklanır (kapatılır)

        # ADIM 3b: 692 → 591 (Öz Kaynaklara)
        t_591 = new_tx('Net Zarar Öz Kaynaklara — 692→591')
        ins(t_591, 591, 'D', net_zarar)
        ins(t_591, 692, 'C', net_zarar)

        flash_msg = (f"⚠️ Dönem kapanışı tamamlandı | "
                     f"Dönem Zararı: {net_zarar:,.2f} TL → 591 Hesabına aktarıldı")

    log_action("Kapanış", flash_msg, existing_cursor=cursor)
    db.commit()
    db.close()
    flash(flash_msg)
    return redirect(url_for('gelir_tablosu'))

# --- 8. MODÜL: İŞLEM SIRASI SİHİRBAZI (10. PDF) ---
@app.route('/wizard')
@login_required
def wizard():
    """12 Adımlık muhasebe döngüsü ve veritabanı kontrollü denetimler"""
    db = get_db()
    
    # Akıllı Denetim: KDV Mahsubu adımı için 191 ve 391 bakiyeleri sıfır mı?
    kdv_check = db.execute('''SELECT SUM(debit - credit) as bakiye FROM journal_entries j 
                              JOIN transactions t ON j.transaction_id = t.id 
                              WHERE t.user_id = ? AND account_code IN (191, 391)''', (current_user.id,)).fetchone()['bakiye'] or 0
    kdv_auto = 1 if abs(kdv_check) < 0.01 else 0
    
    # Kullanıcının daha önce işaretlediği adımları çek
    db_steps = db.execute('SELECT * FROM closing_steps WHERE user_id = ?', (current_user.id,)).fetchall()
    status_map = {row['step_index']: row['is_completed'] for row in db_steps}
    
    steps = ["Açılış Bilançosu", "Günlük Kayıtlar", "Maliyet Kayıtları", "Geçici Mizan", "Envanter ve Değerleme", "KDV Mahsubu", "Amortisman Kayıtları", "Gelir-Gider Kapanışları", "Kesin Mizan", "Gelir Tablosu", "Bilanço", "Kapanış Kayıtları"]
    db.close()
    return render_template('wizard.html', steps=steps, status=status_map, kdv_done=kdv_auto)

@app.route('/update_wizard_step', methods=['POST'])
@login_required
def update_wizard_step():
    """Sihirbazdaki kutucuk işaretlemelerini anlık olarak veritabanına kaydeder"""
    s_idx = int(request.form.get('step_index'))
    is_c = int(request.form.get('is_checked', 0))
    db = get_db()
    db.execute('INSERT OR REPLACE INTO closing_steps (user_id, step_index, is_completed) VALUES (?, ?, ?)', (current_user.id, s_idx, is_c))
    db.commit()
    db.close()
    return "OK"

# --- AUTH VE STANDART ROTALAR ---
@app.route('/register', methods=['GET', 'POST'])
def register():
    """Yeni kullanıcı kaydı (Güvenli Hashing, Şifre Kontrolü ve KVKK)"""
    if request.method == 'POST':
        u_name = request.form.get('username')
        p_word = request.form.get('password')
        kvkk = request.form.get('kvkk_check')

        # GÜVENLİK PROTOKOLÜ: Şifre Kuralları
        if len(p_word) < 8 or not any(x.isupper() for x in p_word) or not any(x.isdigit() for x in p_word):
            flash("Güvenlik Hatası: Şifre en az 8 karakter, 1 büyük harf ve 1 rakam içermelidir!")
            return redirect(url_for('register'))
        
        if not kvkk:
            flash("Yasal Uyarı: KVKK ve Aydınlatma metnini onaylamalısınız!")
            return redirect(url_for('register'))

        hashed_pw = generate_password_hash(p_word, method='pbkdf2:sha256')
        db = get_db()
        try:
            db.execute('INSERT INTO users (username, password_hash) VALUES (?, ?)', (u_name, hashed_pw))
            db.commit()
            # BUG FIX: log_action current_user gerektiriyor, kayıt sırasında henüz auth yok
            # Audit logu doğrudan yazıyoruz
            db.execute('INSERT INTO audit_logs (user_id, action, details) VALUES (?, ?, ?)',
                       (None, 'Kayıt', f"Yeni kullanıcı oluşturuldu: {u_name}. KVKK onaylandı."))
            db.commit()
            return redirect(url_for('login'))
        except Exception:
            flash("Seçilen kullanıcı adı sistemde zaten kayıtlı!")
        finally:
            db.close()
    return render_template('register.html')

@app.route('/login', methods=['GET', 'POST'])
def login():
    """Sisteme giriş portalı (Denetim logu tutulur)"""
    if request.method == 'POST':
        conn = get_db()
        user = conn.execute('SELECT * FROM users WHERE username = ?', (request.form['username'],)).fetchone()
        conn.close()
        if user and check_password_hash(user['password_hash'], request.form['password']):
            login_user(User(user['id'], user['username'], user['profile_pic']))
            log_action("Giriş", "Başarılı oturum açıldı.")
            return redirect(url_for('index'))
        flash("Hatalı kimlik bilgileri!")
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    """Güvenli çıkış ve loglama"""
    log_action("Çıkış", "Oturum sonlandırıldı.")
    logout_user()
    return redirect(url_for('login'))

@app.route('/')
@login_required
def index():
    """Ana Sayfa: Yevmiye Kayıt Merkezi ve Son İşlemler Paneli"""
    db = get_db()
    accounts = db.execute('SELECT * FROM accounts ORDER BY code ASC').fetchall()
    # LEFT JOIN ile tanımlanmamış hesapları da koruyoruz
    query = '''SELECT j.*, a.name, t.description, t.date, t.id as t_id, t.exchange_rate FROM journal_entries j 
               LEFT JOIN accounts a ON j.account_code = a.code JOIN transactions t ON j.transaction_id = t.id 
               WHERE t.user_id = ? ORDER BY t.id DESC'''
    entries = db.execute(query, (current_user.id,)).fetchall()
    rates = db.execute('SELECT * FROM exchange_rates').fetchall()
    
    # YENİ: Kayıt formunda cari seçimi için müşteri listesini gönder
    customers = db.execute('SELECT id, name FROM customers WHERE user_id = ?', (current_user.id,)).fetchall()
    
    db.close()
    return render_template('index.html', accounts=accounts, entries=entries, rates=rates, customers=customers)

# YENİ: OTOMATİK PDF FATURA (TAMİR EDİLDİ: TÜRKÇE KARAKTER FIX)
@app.route('/generate_invoice/<int:t_id>')
@login_required
def generate_invoice(t_id):
    db = get_db()
    entry = db.execute('SELECT * FROM transactions WHERE id = ?', (t_id,)).fetchone()
    details = db.execute('SELECT * FROM journal_entries WHERE transaction_id = ?', (t_id,)).fetchall()
    db.close()
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    # tr_fix eklendi
    pdf.cell(200, 10, txt=tr_fix("BARAN ILGUN ERP - SATIS FATURASI"), ln=True, align='C')
    pdf.set_font("Arial", size=10)
    pdf.cell(200, 10, txt=tr_fix(f"Fatura No: INV-{t_id} | Tarih: {entry['date']}"), ln=True)
    for d in details:
        desc_text = tr_fix(entry['description'][:40] if entry['description'] else "Islem Detayi")
        pdf.cell(100, 10, txt=desc_text, border=1)
        pdf.cell(40, 10, txt=str(d['account_code']), border=1)
        pdf.cell(50, 10, txt=f"{max(d['debit'], d['credit']):,.2f}", border=1, ln=True)
    # encode latin-1 ignore eklendi
    response = make_response(pdf.output(dest='S').encode('latin-1', errors='ignore'))
    response.headers.set('Content-Type', 'application/pdf')
    response.headers.set('Content-Disposition', 'attachment', filename=f'Fatura_{t_id}.pdf')
    return response

@app.route('/save', methods=['POST'])
@login_required
def save():
    """Yeni Yevmiye Fişi Kaydı ve Dövizli İşlem Desteği (2. Hazır Değerler.pdf)"""
    rate = float(request.form.get('rate', 1)) 
    debit_v = float(request.form.get('debit', 0)) * rate
    credit_v = float(request.form.get('credit', 0)) * rate
    acc1 = request.form.get('acc1')
    acc2 = request.form.get('acc2')
    customer_id = request.form.get('customer_id') # YENİ: Cari ID

    # BUG FIX: Hesap kodu boş gelirse NULL kayıt önle
    if not acc1 or not acc2:
        flash("Hata: Borç ve Alacak hesap kodları seçilmeden kayıt yapılamaz!")
        return redirect(url_for('index'))

    # Muhasebe Temel Kuralı: Borç = Alacak
    if debit_v != credit_v or debit_v <= 0:
        flash("Hata: Çift taraflı kayıt ilkesi (Borç/Alacak dengesi) sağlanamadı!")
        return redirect(url_for('index'))
    
    db = get_db()
    cursor = db.cursor()
    # YENİ: Hesap kodu yoksa otomatik ekleme
    for c in [acc1, acc2]:
        if not cursor.execute('SELECT code FROM accounts WHERE code = ?', (c,)).fetchone():
            cursor.execute('INSERT INTO accounts (code, name) VALUES (?, ?)', (c, f"{c} - Yeni Hesap"))

    # YENİ: customer_id veritabanına mühürleniyor
    cursor.execute('INSERT INTO transactions (user_id, description, exchange_rate, customer_id) VALUES (?, ?, ?, ?)', 
                   (current_user.id, request.form['description'], rate, customer_id))
    t_id = cursor.lastrowid
    
    # Hesaplara kayıt atma (B/A)
    cursor.execute('INSERT INTO journal_entries (transaction_id, account_code, debit) VALUES (?,?,?)', (t_id, acc1, debit_v))
    cursor.execute('INSERT INTO journal_entries (transaction_id, account_code, credit) VALUES (?,?,?)', (t_id, acc2, credit_v))
    
    db.commit()
    log_action("Kayıt", f"{request.form['description']} işlemi mizan ve defter-i kebir kayıtlarına eklendi.")
    db.close()
    return redirect(url_for('index'))


@app.route('/save_compound', methods=['POST'])
@login_required
def save_compound():
    """
    BİLEŞİK YEVMİYE KAYDI (M1 — 3+ Hesaplı)

    Müfredat örneği (5. Stoklar.pdf):
        153 Ticari Mallar (B) 13.000
        191 İndirilecek KDV  (B)  1.300
            100 Kasa (A)          7.150
            320 Satıcılar (A)     7.150

    Form yapısı (çoklu satır):
        borc_code[]  = [153, 191]
        borc_tutar[] = [13000, 1300]
        alacak_code[]  = [100, 320]
        alacak_tutar[] = [7150, 7150]
        description = "..."
    """
    description = request.form.get('description', 'Bileşik Yevmiye')
    source_module = request.form.get('source_module', 'Bileşik Yevmiye')

    # HTML form'ları name[] ile gönderir, test client name[] veya name ile gönderebilir
    borc_codes   = request.form.getlist('borc_code[]') or request.form.getlist('borc_code')
    borc_tutars  = request.form.getlist('borc_tutar[]') or request.form.getlist('borc_tutar')
    alacak_codes  = request.form.getlist('alacak_code[]') or request.form.getlist('alacak_code')
    alacak_tutars = request.form.getlist('alacak_tutar[]') or request.form.getlist('alacak_tutar')

    # Boş satırları temizle
    borçlar  = [(c, round(float(t), 2)) for c, t in zip(borc_codes, borc_tutars)
                if c and t and float(t) > 0]
    alacaklar = [(c, round(float(t), 2)) for c, t in zip(alacak_codes, alacak_tutars)
                 if c and t and float(t) > 0]

    if not borçlar or not alacaklar:
        flash("Hata: En az bir borç ve bir alacak satırı gereklidir.")
        return redirect(url_for('index'))

    total_borc   = round(sum(t for _, t in borçlar), 2)
    total_alacak = round(sum(t for _, t in alacaklar), 2)

    # ÇİFT KAYIT DENGESİ KONTROLÜ
    if abs(total_borc - total_alacak) > 0.01:
        flash(f"❌ Bileşik yevmiye dengesi sağlanamadı: "
              f"Borç {total_borc:,.2f} ≠ Alacak {total_alacak:,.2f} TL "
              f"(Fark: {abs(total_borc-total_alacak):,.2f})")
        return redirect(url_for('index'))

    db = get_db()
    cursor = db.cursor()

    # Eksik hesapları otomatik ekle
    all_codes = [c for c, _ in borçlar] + [c for c, _ in alacaklar]
    for c in all_codes:
        if not cursor.execute('SELECT code FROM accounts WHERE code = ?', (float(c),)).fetchone():
            cursor.execute('INSERT INTO accounts (code, name) VALUES (?, ?)',
                           (float(c), f"{c} - Yeni Hesap"))

    # Transaction oluştur
    cursor.execute(
        'INSERT INTO transactions (user_id, description, source_module) VALUES (?,?,?)',
        (current_user.id, description, source_module)
    )
    t_id = cursor.lastrowid

    # Borç satırları
    for code, tutar in borçlar:
        cursor.execute(
            'INSERT INTO journal_entries (transaction_id, account_code, debit) VALUES (?,?,?)',
            (t_id, float(code), tutar)
        )

    # Alacak satırları
    for code, tutar in alacaklar:
        cursor.execute(
            'INSERT INTO journal_entries (transaction_id, account_code, credit) VALUES (?,?,?)',
            (t_id, float(code), tutar)
        )

    db.commit()
    log_action("Bileşik Yevmiye",
               f"{description} | Borç: {total_borc:,.2f} = Alacak: {total_alacak:,.2f} TL | "
               f"{len(borçlar)} borç + {len(alacaklar)} alacak satırı")
    db.close()
    flash(f"✅ Bileşik yevmiye kaydedildi: {len(borçlar)+len(alacaklar)} hesap, "
          f"{total_borc:,.2f} TL (Borç = Alacak)")
    return redirect(url_for('index'))


# ============================================================
# K4 - SÜREKLİ ENVANTER YÖNTEMİ - SATIŞ KAYDI (5. Stoklar.pdf MÜFREDATI)
# ============================================================
# Müfredat: 8. hafta + 13. hafta çalışma soruları:
# "İşletme Sürekli Envanter Yöntemini kullanmaktadır. Satış+Maliyet Kaydı"
#
# SÜREKLİ ENVANTER YÖNTEMİ - HER SATIŞTA İKİ YEVMİYE ATILIR:
# 
# 1) SATIŞ KAYDI:
#     100/120/101 (B) Brüt Satış Bedeli
#         600 Yurt İçi Satışlar (A)  Net Satış
#         391 Hesaplanan KDV (A)     KDV Tutarı
#
# 2) MALİYET KAYDI:
#     621 STMM (B)              Maliyet Tutarı
#         153 Ticari Mallar (A) Maliyet Tutarı
# ============================================================
@app.route('/sale_perpetual', methods=['POST'])
@login_required
def sale_perpetual():
    """
    Sürekli Envanter Yöntemi ile Satış Kaydı.
    Otomatik olarak iki yevmiye atar: Satış + Maliyet (621/153).
    
    Form parametreleri:
        net_amount       : KDV hariç net satış tutarı (TL)
        cost_amount      : Satılan malların maliyeti (TL) → 621/153 için
        kdv_rate         : KDV oranı (%) - 1, 10, 18, 20
        payment_account  : Tahsilat hesabı kodu (100, 102, 120, 101)
        sales_account    : Satış geliri hesabı (default: 600)
        cost_account     : Stok maliyet hesabı (default: 153)
        description      : İşlem açıklaması
        customer_id      : (opsiyonel) Cari kart ID
    """
    try:
        net_amount = float(request.form.get('net_amount', 0))
        cost_amount = float(request.form.get('cost_amount', 0))
        kdv_rate = float(request.form.get('kdv_rate', 20))  # Default %20 (2024+)
        payment_account = request.form.get('payment_account', '100')
        sales_account = request.form.get('sales_account', '600')
        cost_account = request.form.get('cost_account', '153')
        description = request.form.get('description', 'Sürekli Envanter Satış Kaydı')
        customer_id = request.form.get('customer_id') or None
    except (TypeError, ValueError):
        flash("Hata: Geçersiz tutar veya oran girdiniz.")
        return redirect(url_for('index'))
    
    # Validasyonlar
    if net_amount <= 0:
        flash("Hata: Net satış tutarı sıfırdan büyük olmalıdır.")
        return redirect(url_for('index'))
    if cost_amount < 0:
        flash("Hata: Maliyet tutarı negatif olamaz.")
        return redirect(url_for('index'))
    if cost_amount > net_amount * 1.5:
        # Sanity check: maliyet satıştan çok daha yüksekse kullanıcı uyarılır
        flash(f"⚠️ Uyarı: Maliyet ({cost_amount:.2f}) satış tutarından (%50+) yüksek görünüyor. Kayıt yapıldı.")
    if kdv_rate not in [0, 1, 10, 18, 20]:
        flash("Hata: KDV oranı yalnızca %0, %1, %10, %18 veya %20 olabilir.")
        return redirect(url_for('index'))
    
    # KDV ve brüt tutarı hesapla
    kdv_amount = round(net_amount * (kdv_rate / 100.0), 2)
    gross_amount = round(net_amount + kdv_amount, 2)
    net_amount = round(net_amount, 2)
    cost_amount = round(cost_amount, 2)
    
    db = get_db()
    cursor = db.cursor()
    
    try:
        # =================================
        # YEVMİYE 1: SATIŞ KAYDI
        # =================================
        full_desc = f"[Sürekli Env.] {description} | KDV %{int(kdv_rate)} | Net: {net_amount:,.2f} TL"
        cursor.execute('INSERT INTO transactions (user_id, description, customer_id, source_module) VALUES (?, ?, ?, ?)',
                       (current_user.id, full_desc, customer_id, 'Sürekli Envanter - Satış'))
        t_id_sale = cursor.lastrowid
        
        # 100/120/102 (B) Brüt Tutar
        cursor.execute('INSERT INTO journal_entries (transaction_id, account_code, debit) VALUES (?, ?, ?)',
                       (t_id_sale, payment_account, gross_amount))
        # 600 Yurtiçi Satışlar (A) Net
        cursor.execute('INSERT INTO journal_entries (transaction_id, account_code, credit) VALUES (?, ?, ?)',
                       (t_id_sale, sales_account, net_amount))
        # 391 Hesaplanan KDV (A) - KDV varsa
        if kdv_amount > 0:
            cursor.execute('INSERT INTO journal_entries (transaction_id, account_code, credit) VALUES (?, 391, ?)',
                           (t_id_sale, kdv_amount))
        
        # =================================
        # YEVMİYE 2: MALİYET KAYDI (621/153) - Sürekli Envanter farkı
        # =================================
        if cost_amount > 0:
            cost_desc = f"[Sürekli Env. - MALİYET] {description} | STMM: {cost_amount:,.2f} TL"
            cursor.execute('INSERT INTO transactions (user_id, description, customer_id, source_module) VALUES (?, ?, ?, ?)',
                           (current_user.id, cost_desc, customer_id, 'Sürekli Envanter - Maliyet'))
            t_id_cost = cursor.lastrowid
            
            # 621 STMM (B) Maliyet
            cursor.execute('INSERT INTO journal_entries (transaction_id, account_code, debit) VALUES (?, 621, ?)',
                           (t_id_cost, cost_amount))
            # 153 Ticari Mallar (A) Maliyet  
            cursor.execute('INSERT INTO journal_entries (transaction_id, account_code, credit) VALUES (?, ?, ?)',
                           (t_id_cost, cost_account, cost_amount))
        
        db.commit()
        log_action("Sürekli Envanter Satışı",
                   f"Satış: {net_amount:,.2f} + KDV {kdv_amount:,.2f} = {gross_amount:,.2f} TL | Maliyet: {cost_amount:,.2f}")
        flash(f"✅ Sürekli Envanter Satış Kaydı tamamlandı: Brüt {gross_amount:,.2f} TL satış + {cost_amount:,.2f} TL maliyet kaydı atıldı.")
    except Exception as e:
        db.rollback()
        flash(f"Hata: Satış kaydı oluşturulamadı - {str(e)}")
    finally:
        db.close()
    
    return redirect(url_for('index'))


# ============================================================
# K12 - KDV ORANI SEÇİMLİ ALIŞ KAYDI (6. KDV.pdf MÜFREDATI)
# ============================================================
# Müfredat: %1 / %10 / %18 / %20 farklı oranlarda KDV
# Müfredat formülü (5. Stoklar.pdf):
#     153 Ticari Mallar (B)     Net Tutar
#     191 İndirilecek KDV (B)   KDV Tutarı
#         100/102/320 (A)        Brüt Tutar
# ============================================================
@app.route('/purchase_with_vat', methods=['POST'])
@login_required
def purchase_with_vat():
    """
    KDV oranı seçimli alış kaydı (otomatik 191 KDV hesabı).
    
    Form parametreleri:
        net_amount       : KDV hariç tutar
        kdv_rate         : %1, %10, %18, %20 (2024+ için %20)
        payment_account  : Ödeme hesabı (100, 102, 320, 321)
        purchase_account : Alış hesabı (153, 150, 255, 252 vb.)
        description      : Açıklama
        customer_id      : (opsiyonel)
    """
    try:
        net_amount = float(request.form.get('net_amount', 0))
        kdv_rate = float(request.form.get('kdv_rate', 20))
        payment_account = request.form.get('payment_account', '100')
        purchase_account = request.form.get('purchase_account', '153')
        description = request.form.get('description', 'KDV Otomatik Alış Kaydı')
        customer_id = request.form.get('customer_id') or None
    except (TypeError, ValueError):
        flash("Hata: Geçersiz tutar.")
        return redirect(url_for('index'))
    
    if net_amount <= 0:
        flash("Hata: Net tutar pozitif olmalıdır.")
        return redirect(url_for('index'))
    if kdv_rate not in [0, 1, 10, 18, 20]:
        flash("Hata: KDV oranı %0, %1, %10, %18, %20 dışında olamaz.")
        return redirect(url_for('index'))
    
    kdv_amount = round(net_amount * (kdv_rate / 100.0), 2)
    gross_amount = round(net_amount + kdv_amount, 2)
    net_amount = round(net_amount, 2)
    
    db = get_db()
    cursor = db.cursor()
    
    try:
        full_desc = f"[KDV %{int(kdv_rate)}] {description} | Net: {net_amount:,.2f} TL"
        cursor.execute('INSERT INTO transactions (user_id, description, customer_id, source_module) VALUES (?, ?, ?, ?)',
                       (current_user.id, full_desc, customer_id, 'KDV Otomatik Alış'))
        t_id = cursor.lastrowid
        
        # XX (B) Net (153, 150, 255, 252, ...)
        cursor.execute('INSERT INTO journal_entries (transaction_id, account_code, debit) VALUES (?, ?, ?)',
                       (t_id, purchase_account, net_amount))
        # 191 İndirilecek KDV (B) - varsa
        if kdv_amount > 0:
            cursor.execute('INSERT INTO journal_entries (transaction_id, account_code, debit) VALUES (?, 191, ?)',
                           (t_id, kdv_amount))
        # 100/102/320 (A) Brüt
        cursor.execute('INSERT INTO journal_entries (transaction_id, account_code, credit) VALUES (?, ?, ?)',
                       (t_id, payment_account, gross_amount))
        
        db.commit()
        log_action("KDV Alış",
                   f"Alış: Net {net_amount:,.2f} + KDV {kdv_amount:,.2f} = {gross_amount:,.2f} TL")
        flash(f"✅ KDV %{int(kdv_rate)} ile alış kaydı tamamlandı: {gross_amount:,.2f} TL")
    except Exception as e:
        db.rollback()
        flash(f"Hata: {str(e)}")
    finally:
        db.close()
    
    return redirect(url_for('index'))


# ============================================================
# K10 - ŞÜPHELİ ALACAK KARŞILIĞI (4. Alacak ve Borçlar.pdf + Kesin Mizan Örneği)
# ============================================================
# Kesin Mizan (10. PDF): 128: 7.000 / 129: 7.000 / 654: 7.000
# 
# İŞLEM 1 - Karşılık Ayrılması:
#   120/121 alacak → 128 Şüpheli TicariAlacak (transfer)
#   654 Karşılık Giderleri (B) / 129 Şüpheli T.A. Karşılığı (A)
#
# İŞLEM 2 - Tahsil (kısmen veya tamamen):
#   100/102 (B) / 128 (A) → Tahsil edilen kısım
#   129 (B) / 654 (A) → Karşılık iptali
#
# İŞLEM 3 - Değersiz Alacak (tamamen siliş):
#   129 (B) / 128 (A) → Karşılık + alacak silinir
# ============================================================
@app.route('/supheli_alacak', methods=['POST'])
@login_required
def supheli_alacak():
    """
    Şüpheli Alacak İşlem Merkezi
    
    islem_turu:
      'karsili_ayir' : 120/121 → 128 + 654/129 karşılık
      'tahsil_et'    : 100/102 / 128 + 129 / 654 iptal
      'degersis_sil' : 129 / 128 değersiz alacak silme
    """
    tur = request.form.get('islem_turu', 'karsili_ayir')
    tutar = round(float(request.form.get('tutar', 0) or 0), 2)
    kaynak_hesap = request.form.get('kaynak_hesap', '120')
    tahsilat_hesap = request.form.get('tahsilat_hesap', '100')
    aciklama = request.form.get('aciklama', 'Şüpheli Alacak İşlemi')

    if tutar <= 0:
        flash("Hata: Tutar sıfırdan büyük olmalıdır.")
        return redirect(url_for('index'))

    db = get_db()
    cursor = db.cursor()

    def ins(tid, code, side, amt):
        col = 'debit' if side == 'D' else 'credit'
        cursor.execute(f'INSERT INTO journal_entries (transaction_id, account_code, {col}) VALUES (?,?,?)',
                       (tid, code, round(abs(amt), 2)))

    try:
        if tur == 'karsili_ayir':
            # ADIM 1: 120/121 → 128 (alacak transferi)
            cursor.execute('INSERT INTO transactions (user_id, description, source_module) VALUES (?,?,?)',
                           (current_user.id, f'[Şüph.Alacak Transfer] {aciklama}', 'Şüpheli Alacak'))
            t1 = cursor.lastrowid
            ins(t1, 128, 'D', tutar)              # 128 Şüpheli T.A. (B)
            ins(t1, int(kaynak_hesap), 'C', tutar) # 120 veya 121 (A)

            # ADIM 2: Karşılık gideri
            cursor.execute('INSERT INTO transactions (user_id, description, source_module) VALUES (?,?,?)',
                           (current_user.id, f'[Karşılık Gideri] {aciklama}', 'Şüpheli Alacak'))
            t2 = cursor.lastrowid
            ins(t2, 654, 'D', tutar)  # 654 Karşılık Giderleri (B) → 7A gideri
            ins(t2, 129, 'C', tutar)  # 129 Şüpheli T.A. Karşılığı (A)

            flash(f"✅ Şüpheli alacak karşılığı: {tutar:,.2f} TL | "
                  f"{kaynak_hesap}→128 transfer + 654 Karşılık Gideri / 129 Karşılık")

        elif tur == 'tahsil_et':
            # Tahsil: para geldi, karşılık iptal
            cursor.execute('INSERT INTO transactions (user_id, description, source_module) VALUES (?,?,?)',
                           (current_user.id, f'[Şüph.Alacak Tahsil] {aciklama}', 'Şüpheli Alacak'))
            t1 = cursor.lastrowid
            ins(t1, int(tahsilat_hesap), 'D', tutar)  # Kasa/Banka (B)
            ins(t1, 128, 'C', tutar)                  # 128 (A) — alacak kapandı

            # 129 karşılığı iptal et
            cursor.execute('INSERT INTO transactions (user_id, description, source_module) VALUES (?,?,?)',
                           (current_user.id, f'[Karşılık İptal] {aciklama}', 'Şüpheli Alacak'))
            t2 = cursor.lastrowid
            ins(t2, 129, 'D', tutar)  # 129 (B) — karşılık kapandı
            ins(t2, 654, 'C', tutar)  # 654 (A) — gelir düzeltmesi

            flash(f"✅ Şüpheli alacak tahsil edildi: {tutar:,.2f} TL | 129 karşılık iptal edildi")

        elif tur == 'degersis_sil':
            # Alacak tamamen değersiz: 128 ve 129 silinir, birbirini kapatır
            cursor.execute('INSERT INTO transactions (user_id, description, source_module) VALUES (?,?,?)',
                           (current_user.id, f'[Değersiz Alacak Silme] {aciklama}', 'Şüpheli Alacak'))
            t1 = cursor.lastrowid
            ins(t1, 129, 'D', tutar)  # 129 Karşılık (B) kapandı
            ins(t1, 128, 'C', tutar)  # 128 Şüpheli T.A. (A) kapandı

            flash(f"✅ Değersiz alacak silindi: {tutar:,.2f} TL | 128 ve 129 hesapları kapatıldı")

        db.commit()
        log_action("Şüpheli Alacak", f"{tur}: {tutar:,.2f} TL | {aciklama}")
    except Exception as e:
        db.rollback()
        flash(f"Hata: {e}")
    finally:
        db.close()

    return redirect(url_for('index'))


# ============================================================
# K7 - SENET YÖNETİM MODÜLÜ (4. Alacak ve Borçlar.pdf)
# ============================================================
# Desteklenen işlemler:
#   1. Alacak senedi yenileme  → 121/121 + 642
#   2. Alacak senedi bankaya tahsile verme → 121.02/121.01
#   3. Alacak senedi tahsil edilmesi → 102/121.02
#   4. Borç senedi yenileme  → 321/321 + 780
# ============================================================
@app.route('/senet_islem', methods=['POST'])
@login_required
def senet_islem():
    """
    Senet İşlem Merkezi (4. Alacak ve Borçlar.pdf Müfredatı)

    işlem_turu parametresine göre:
    - yenileme_alacak : Alacak senedi yenileme (121/121 + 642 Faiz Geliri)
    - tahsile_ver     : Senet bankaya tahsile verilmesi (121.02/121.01)
    - tahsil_et       : Bankadan tahsil (102/121.02)
    - yenileme_borc   : Borç senedi yenileme (321/321 + 780 Fin. Gid.)
    - ciro            : Senet cirolu ödeme (XX/121.01 yeni senet varlık devri)
    """
    tur = request.form.get('islem_turu')
    eski_tutar = float(request.form.get('eski_tutar', 0))
    yeni_tutar = float(request.form.get('yeni_tutar', 0) or eski_tutar)
    faiz_tutar = round(yeni_tutar - eski_tutar, 2)
    aciklama   = request.form.get('aciklama', 'Senet İşlemi')
    odeme_hsp  = request.form.get('odeme_hesap', '100')
    
    if eski_tutar <= 0:
        flash("Hata: Senet tutarı sıfırdan büyük olmalıdır.")
        return redirect(url_for('index'))

    db = get_db()
    cursor = db.cursor()

    def ins(tid, code, side, amt):
        col = 'debit' if side == 'D' else 'credit'
        cursor.execute(
            f'INSERT INTO journal_entries (transaction_id, account_code, {col}) VALUES (?,?,?)',
            (tid, code, round(abs(amt), 2))
        )

    cursor.execute('INSERT INTO transactions (user_id, description, source_module) VALUES (?,?,?)',
                   (current_user.id, f'[Senet] {aciklama}', f'Senet-{tur}'))
    tid = cursor.lastrowid

    try:
        if tur == 'yenileme_alacak':
            # Müfredat: Yeni senet (B) / Eski senet (A) + 642 Faiz Geliri (A)
            # Örnek: 2.000 TL senet → 2.500 TL yeni senet | 642: 500 TL faiz
            ins(tid, 121, 'D', yeni_tutar)   # 121 Yeni senet (borçlu)
            ins(tid, 121, 'C', eski_tutar)   # 121 Eski senet (alacaklı = kapatılır)
            if faiz_tutar > 0:
                ins(tid, 642, 'C', faiz_tutar)  # 642 Faiz Geliri
            flash(f"✅ Alacak senedi yenilendi: {eski_tutar:,.2f} → {yeni_tutar:,.2f} TL "
                  f"{'(Faiz: ' + str(faiz_tutar) + ' TL)' if faiz_tutar > 0 else ''}")

        elif tur == 'tahsile_ver':
            # Müfredat: 121.02 Tahsile verilen (B) / 121.01 Portföydeki (A)
            ins(tid, 121, 'D', eski_tutar)   # 121.02 Tahsile verilen → Borç
            ins(tid, 121, 'C', eski_tutar)   # 121.01 Portföy → Alacak
            flash(f"✅ {eski_tutar:,.2f} TL'lik senet bankaya tahsile verildi.")

        elif tur == 'tahsil_et':
            # Müfredat: 102 Bankalar (B) / 121.02 Tahsile verilen (A)
            ins(tid, 102, 'D', eski_tutar)   # Banka borçlu (para geldi)
            ins(tid, 121, 'C', eski_tutar)   # 121.02 Alacak (senet kapatıldı)
            flash(f"✅ {eski_tutar:,.2f} TL senet banka aracılığıyla tahsil edildi.")

        elif tur == 'elden_tahsil':
            # 100 Kasa (B) / 121 Alacak Senetleri (A)
            ins(tid, 100, 'D', eski_tutar)
            ins(tid, 121, 'C', eski_tutar)
            flash(f"✅ {eski_tutar:,.2f} TL'lik senet elden tahsil edildi → 100 Kasa.")

        elif tur == 'yenileme_borc':
            # Müfredat: Eski 321 (B) + 780 Fin.Gid. (B) / Yeni 321 (A)
            # Örnek: 4.000 → 4.300 | 780: 300 TL finansman gideri
            ins(tid, 321, 'D', eski_tutar)   # Eski borç senedi kapandı
            if faiz_tutar > 0:
                ins(tid, 780, 'D', faiz_tutar)   # 780 Finansman Giderleri
            ins(tid, 321, 'C', yeni_tutar)   # Yeni borç senedi oluştu
            flash(f"✅ Borç senedi yenilendi: {eski_tutar:,.2f} → {yeni_tutar:,.2f} TL "
                  f"{'(780 Fin.Gid.: ' + str(faiz_tutar) + ' TL)' if faiz_tutar > 0 else ''}")

        elif tur == 'ciro':
            # Müfredat: Ciro = Alıcıya senet devri ile borç ödeme
            # BUG FIX: Çifte borç kaldırıldı — sadece odeme_hsp borçlanır
            ins(tid, int(odeme_hsp), 'D', eski_tutar)  # 320/321 vb. (B)
            ins(tid, 121, 'C', eski_tutar)              # 121 Alacak senet devredildi (A)
            flash(f"✅ {eski_tutar:,.2f} TL'lik senet {odeme_hsp} hesabına ciro edildi.")

        else:
            flash(f"Hata: Bilinmeyen senet işlem türü: {tur}")
            db.rollback()
            db.close()
            return redirect(url_for('index'))

        db.commit()
        log_action("Senet İşlemi", f"{tur}: {eski_tutar:,.2f} → {yeni_tutar:,.2f} TL")
    except Exception as e:
        db.rollback()
        flash(f"Senet işlemi hatası: {e}")
    finally:
        db.close()

    return redirect(url_for('index'))


# ============================================================
# K8 - MENKUL KIYMET İŞLEMLERİ (4. Menkul Kıymetler.pdf)
# ============================================================
# Müfredat:
#   ALIŞ: 110(B) + 653 Komisyon(B) / 102(A) [brüt]
#   SATIŞ KÂR: 102(B) [net] + 653 Komisyon(B) / 110(A) [maliyet] + 645 Kâr(A)
#   SATIŞ ZARAR: 102(B) [net] + 653 Komisyon(B) + 655 Zarar(B) / 110(A) [maliyet]
# ============================================================
@app.route('/menkul_kiymet', methods=['POST'])
@login_required
def menkul_kiymet():
    """
    Menkul Kıymet Alış/Satış (4. Menkul Kıymetler.pdf Müfredatı)

    Form parametreleri:
        islem_turu  : 'alis' | 'satis'
        adet        : Hisse adedi
        fiyat       : Birim fiyat (TL)
        maliyet_fiyat: Alış birim maliyeti (satış için)
        komisyon    : Komisyon tutarı (TL) → 653
        hisse_kod   : 110 | 111 | 112 | 240 | 242 | 245
        odeme_hsp   : 100 | 102
    """
    tur          = request.form.get('islem_turu', 'alis')
    adet         = int(request.form.get('adet', 0) or 0)
    fiyat        = float(request.form.get('fiyat', 0) or 0)
    komisyon     = float(request.form.get('komisyon', 0) or 0)
    hisse_kod    = int(request.form.get('hisse_kod', 110))
    odeme_hsp    = int(request.form.get('odeme_hsp', 102))
    aciklama     = request.form.get('aciklama', 'Menkul Kıymet İşlemi')

    if adet <= 0 or fiyat <= 0:
        flash("Hata: Adet ve fiyat sıfırdan büyük olmalıdır.")
        return redirect(url_for('index'))

    toplam_islem = round(adet * fiyat, 2)
    komisyon     = round(komisyon, 2)

    db = get_db()
    cursor = db.cursor()

    def ins(tid, code, side, amt):
        col = 'debit' if side == 'D' else 'credit'
        cursor.execute(
            f'INSERT INTO journal_entries (transaction_id, account_code, {col}) VALUES (?,?,?)',
            (tid, code, round(abs(amt), 2))
        )

    try:
        if tur == 'alis':
            # ALIŞ: 110(B) maliyet + 653(B) komisyon / 102(A) brüt
            # Müfredat örneği: 2000 adet × 12 TL = 24.000 + 480 komisyon → Banka 24.480
            cursor.execute('INSERT INTO transactions (user_id, description, source_module) VALUES (?,?,?)',
                           (current_user.id, f'[MK Alış] {adet} adet × {fiyat} TL | {aciklama}', 'Menkul Kıymet-Alış'))
            tid = cursor.lastrowid

            ins(tid, hisse_kod, 'D', toplam_islem)   # 110 Hisse Senetleri (B)
            if komisyon > 0:
                ins(tid, 653, 'D', komisyon)          # 653 Komisyon Giderleri (B)
            ins(tid, odeme_hsp, 'C', toplam_islem + komisyon)  # Banka (A) brüt

            flash(f"✅ Menkul Kıymet Alış: {adet} adet × {fiyat:,.2f} TL = {toplam_islem:,.2f} TL "
                  f"+ Komisyon {komisyon:,.2f} TL | {hisse_kod} Hesabı güncellendi.")

        elif tur == 'satis':
            # SATIŞ: Maliyet birim fiyatını al
            maliyet_fiyat = float(request.form.get('maliyet_fiyat', fiyat) or fiyat)
            toplam_maliyet = round(adet * maliyet_fiyat, 2)
            net_tahsilat   = round(toplam_islem - komisyon, 2)
            kar_zarar      = round(net_tahsilat - toplam_maliyet, 2)

            cursor.execute('INSERT INTO transactions (user_id, description, source_module) VALUES (?,?,?)',
                           (current_user.id, f'[MK Satış] {adet} adet × {fiyat} TL | {aciklama}', 'Menkul Kıymet-Satış'))
            tid = cursor.lastrowid

            # BUG FIX: Satış yevmiye dengesi
            # Müfredat: 102(B)=BRÜT, 653(B)=komisyon, 110(A)=maliyet, 645/655=kar-zarar
            kar_zarar = round(toplam_islem - komisyon - toplam_maliyet, 2)
            ins(tid, odeme_hsp, 'D', toplam_islem)      # 102 Banka BRÜT satış
            if komisyon > 0:
                ins(tid, 653, 'D', komisyon)             # 653 Komisyon
            ins(tid, hisse_kod, 'C', toplam_maliyet)    # 110 Maliyet

            if kar_zarar > 0:
                ins(tid, 645, 'C', kar_zarar)            # 645 Kâr
                flash(f"✅ MK Satış: {adet} adet × {fiyat:,.2f} TL = {toplam_islem:,.2f} TL | "
                      f"Maliyet: {toplam_maliyet:,.2f} | KÂR: {kar_zarar:,.2f} TL → 645")
            elif kar_zarar < 0:
                # 655 Menkul Kıymet Satış Zararları (B)
                ins(tid, 655, 'D', abs(kar_zarar))
                flash(f"⚠️ MK Satış: ZARAR {abs(kar_zarar):,.2f} TL → 655")
            else:
                flash(f"✅ MK Satış: Kâr/Zarar yok.")

        db.commit()
        log_action("Menkul Kıymet", f"{tur}: {adet}×{fiyat} TL | Komisyon: {komisyon}")
    except Exception as e:
        db.rollback()
        flash(f"Menkul Kıymet işlem hatası: {e}")
    finally:
        db.close()

    return redirect(url_for('index'))

@app.route('/delete/<int:t_id>')
@login_required
def delete(t_id):
    """İşlem silme modülü (İç Denetim Logu ile birlikte) - GÜVENLİK: ownership kontrollü"""
    db = get_db()
    
    # GÜVENLİK B5: Bu işlem gerçekten bu kullanıcıya mı ait?
    owner_check = db.execute('SELECT user_id FROM transactions WHERE id = ?', (t_id,)).fetchone()
    if not owner_check:
        flash("İşlem bulunamadı.")
        db.close()
        return redirect(url_for('index'))
    if owner_check['user_id'] != current_user.id:
        flash("Bu işlemi silme yetkiniz yok.")
        log_action("Yetkisiz Erişim", f"Kullanıcı başkasının işlemini silmeyi denedi (ID: {t_id})")
        db.close()
        return redirect(url_for('index'))
    
    db.execute('DELETE FROM journal_entries WHERE transaction_id = ?', (t_id,))
    db.execute('DELETE FROM transactions WHERE id = ? AND user_id = ?', (t_id, current_user.id))
    db.commit()
    log_action("Silme", f"Kayıt silindi (ID: {t_id}). Bu işlem mali tabloları güncelledi.")
    db.close()
    return redirect(url_for('index'))

# --- RAPORLAMA VE ANALİZ ROTALARI ---
@app.route('/mizan')
@login_required
def mizan():
    """Mizan Tablosu: Hesapların matematiksel kontrol noktası"""
    db = get_db()
    sql = '''SELECT a.code, a.name, SUM(j.debit) as t_borc, SUM(j.credit) as t_alacak, (SUM(j.debit) - SUM(j.credit)) as bakiye 
             FROM accounts a JOIN journal_entries j ON a.code = j.account_code 
             JOIN transactions t ON j.transaction_id = t.id WHERE t.user_id = ? GROUP BY a.code'''
    data = db.execute(sql, (current_user.id,)).fetchall()
    db.close()
    return render_template('mizan.html', mizan=data)


@app.route('/ledger')
@app.route('/ledger/<path:code>')
@login_required
def ledger(code=None):
    """
    DEFTERİ KEBİR (T-HESABI) GÖRÜNÜMÜ (Muhasebeye Giriş Notu Müfredatı)

    Müfredat: "Yevmiye kayıtları dışında defteri kebir kayıtları vardır,
    o da her hesabın borç ve alacaklarının yazıldığı kayıtlardır."

    /ledger         → tüm kullanılmış hesapların listesi (seçim ekranı)
    /ledger/100     → 100 Kasa'nın T-hesabı görünümü
    /ledger/153     → 153 Ticari Mallar'ın T-hesabı
    """
    db = get_db()

    # Kullanıcının hareket görmüş tüm hesapları
    used_accounts_sql = '''
        SELECT a.code, a.name,
               ROUND(SUM(j.debit), 2) as t_borc,
               ROUND(SUM(j.credit), 2) as t_alacak,
               ROUND(SUM(j.debit) - SUM(j.credit), 2) as net_bakiye
        FROM accounts a
        JOIN journal_entries j ON a.code = j.account_code
        JOIN transactions t ON j.transaction_id = t.id
        WHERE t.user_id = ?
        GROUP BY a.code
        ORDER BY CAST(a.code AS REAL)
    '''
    all_accounts = db.execute(used_accounts_sql, (current_user.id,)).fetchall()

    selected = None
    movements = []

    if code:
        # Seçili hesabın tüm hareketleri (kronolojik)
        mov_sql = '''
            SELECT t.date, t.description, t.source_module,
                   ROUND(j.debit, 2) as borc,
                   ROUND(j.credit, 2) as alacak
            FROM journal_entries j
            JOIN transactions t ON j.transaction_id = t.id
            WHERE t.user_id = ? AND CAST(j.account_code AS TEXT) = ?
            ORDER BY t.id ASC
        '''
        movements = db.execute(mov_sql, (current_user.id, str(code))).fetchall()

        # Hesap bilgisi
        selected = db.execute(
            'SELECT code, name FROM accounts WHERE CAST(code AS TEXT) = ?', (str(code),)
        ).fetchone()

        # Bulunamazsa kayan kodla dene (153.1 vs 153.01 gibi)
        if not selected and movements:
            selected = db.execute(
                'SELECT code, name FROM accounts WHERE code = ?', (float(code),)
            ).fetchone()

    db.close()
    return render_template('ledger.html',
                           all_accounts=all_accounts,
                           selected=selected,
                           movements=movements,
                           current_code=code)

@app.route('/acilis', methods=['GET', 'POST'])
@login_required
def acilis():
    """
    AÇILIŞ BİLANÇOSU SİHİRBAZI (10. İşlem Sırası PDF — Adım 4)

    Müfredat: "Başlangıç/İşe Başlama Bilançosuna dayanarak Açılış Kaydının yapılması.
    Aktifindeki hesaplar yevmiye maddesinin BORÇLU kısmına,
    pasifindeki hesaplar ALACAK kısmına yazılır.
    Aktif düzenleyici hesaplar (103, 257 vb.) ALACAK kısmına,
    pasif düzenleyici hesaplar BORÇ kısmına yazılır."

    Kısa Sınav Sorusu (5p.): "İşletmenin sermayesi kaç TL'dir?"
    YBS Genel Muhasebe Kısa Sınavı: 8 kalem verilmiş, sermaye + açılış kaydı isteniyor.
    """
    db = get_db()
    accounts = db.execute('SELECT code, name FROM accounts ORDER BY CAST(code AS REAL)').fetchall()

    # Mevcut açılış kaydı var mı?
    existing = db.execute("""
        SELECT COUNT(*) as c FROM transactions
        WHERE user_id = ? AND source_module = 'Açılış Bilançosu'
    """, (current_user.id,)).fetchone()['c']

    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'post_opening':
            # Form verisini parse et
            # Format: aktif_code[], aktif_tutar[], pasif_code[], pasif_tutar[]
            aktif_codes  = request.form.getlist('aktif_code')
            aktif_tutars = request.form.getlist('aktif_tutar')
            pasif_codes  = request.form.getlist('pasif_code')
            pasif_tutars = request.form.getlist('pasif_tutar')

            # Boş satırları filtrele
            aktifler = [(c, float(t)) for c, t in zip(aktif_codes, aktif_tutars)
                        if c and t and float(t) > 0]
            pasifler = [(c, float(t)) for c, t in zip(pasif_codes, pasif_tutars)
                        if c and t and float(t) > 0]

            if not aktifler and not pasifler:
                flash("Hata: En az bir aktif veya pasif kalem girilmelidir.")
                db.close()
                return redirect(url_for('acilis'))

            # Denklik Kontrolü
            toplam_aktif = round(sum(t for _, t in aktifler), 2)
            toplam_pasif = round(sum(t for _, t in pasifler), 2)

            if abs(toplam_aktif - toplam_pasif) > 0.01:
                flash(f"⚠️ Aktif ({toplam_aktif:,.2f}) ≠ Pasif ({toplam_pasif:,.2f}) — "
                      f"Fark: {abs(toplam_aktif - toplam_pasif):,.2f} TL. Sermaye hesabını kontrol edin.")
                db.close()
                return redirect(url_for('acilis'))

            # AÇILIŞ YEVMİYESİ OLUŞTUR
            # Müfredat: Aktifler BORÇLU, Pasifler ALACAKLI
            # Aktif kontralar (103, 257 vb.) ALACAKLI tarafta
            AKTIF_KONTRA = {103, 119, 122, 129, 158, 241, 244, 247, 257, 268, 278, 299}

            cursor = db.cursor()
            cursor.execute(
                'INSERT INTO transactions (user_id, description, source_module) VALUES (?,?,?)',
                (current_user.id, 'Dönembaşı Açılış Kaydı', 'Açılış Bilançosu')
            )
            t_id = cursor.lastrowid

            for code, tutar in aktifler:
                try:
                    code_int = int(float(code))
                except ValueError:
                    continue
                if code_int in AKTIF_KONTRA:
                    # Kontra aktif → ALACAK tarafa
                    cursor.execute(
                        'INSERT INTO journal_entries (transaction_id, account_code, credit) VALUES (?,?,?)',
                        (t_id, float(code), tutar)
                    )
                else:
                    # Normal aktif → BORÇ tarafa
                    cursor.execute(
                        'INSERT INTO journal_entries (transaction_id, account_code, debit) VALUES (?,?,?)',
                        (t_id, float(code), tutar)
                    )

            for code, tutar in pasifler:
                # Pasifler → ALACAK tarafa
                cursor.execute(
                    'INSERT INTO journal_entries (transaction_id, account_code, credit) VALUES (?,?,?)',
                    (t_id, float(code), tutar)
                )

            log_action("Açılış Kaydı",
                       f"Aktif: {toplam_aktif:,.2f} TL | Pasif: {toplam_pasif:,.2f} TL",
                       existing_cursor=cursor)
            db.commit()  # BUG FIX: tek commit
            flash(f"✅ Açılış kaydı oluşturuldu! Aktif = Pasif = {toplam_aktif:,.2f} TL | "
                  f"Defter-i Kebir hesapları açıldı.")
            db.close()
            return redirect(url_for('ledger'))

    db.close()
    return render_template('acilis.html', accounts=accounts, existing=existing)


@app.route('/bilanco')
@login_required
def bilanco():
    """
    AKADEMİK BİLANÇO (Muhasebeye Giriş Notu, Deneme Sınavı 1 Q5 müfredatı)
    
    DOĞRU FORMÜL:
        Aktif hesaplar (1, 2): bakiye = Borç - Alacak (doğal borç bakiyeli)
        Pasif hesaplar (3, 4, 5): bakiye = Alacak - Borç (doğal alacak bakiyeli)
    
    KONTRA HESAPLAR (Düzenleyici) — müfredatta eksi (-) işaretli:
        AKTİF KONTRA: 103, 119, 122, 129, 257, 158, 241, 244, 247, 268, 278
            → Bu hesaplar AKTİFTE (-) olarak gösterilir, aktif toplamından düşülür
        PASİF KONTRA: 322, 591, 610, 611, 612 (kontra düzenleyici)
            → 591 Dönem Net Zararı pasifte (-), 610/611 satıştan iade (gelir tablosu kontra)
    
    Müfredat referans (Deneme Sınavı 1 - Soru 1):
        "Verilen Çekler ve Ödeme Emirleri Hs(-): 1.000 TL" → 103 hesabı eksi
    """
    db = get_db()
    sql = '''SELECT a.code, a.name, 
             COALESCE(SUM(j.debit), 0) as t_borc,
             COALESCE(SUM(j.credit), 0) as t_alacak,
             (COALESCE(SUM(j.debit), 0) - COALESCE(SUM(j.credit), 0)) as net_bakiye
             FROM accounts a 
             JOIN journal_entries j ON a.code = j.account_code 
             JOIN transactions t ON j.transaction_id = t.id 
             WHERE t.user_id = ? GROUP BY a.code'''
    data = db.execute(sql, (current_user.id,)).fetchall()
    
    # AKTİF KONTRA HESAPLAR (Müfredat: bilançoda eksi (-) ile gösterilir)
    AKTIF_KONTRA = {103, 119, 122, 129, 158, 241, 244, 247, 257, 268, 278, 299}
    # PASİF KONTRA HESAPLAR
    PASIF_KONTRA = {322, 591}
    
    aktifler = []
    pasifler = []
    aktif_toplam = 0.0
    pasif_toplam = 0.0
    
    for r in data:
        code = r['code']
        code_int = int(code)  # 153.1 → 153 dönüşümü için
        code_str = str(int(code)) if float(code) == int(code) else str(code)
        first_digit = code_str[0]
        
        # Mizan formülü: Borç - Alacak (varlıklar için pozitif olmalı)
        net = float(r['net_bakiye'])
        
        # AKTİF GRUBU (Sınıf 1 ve 2)
        if first_digit in ('1', '2'):
            is_kontra = code_int in AKTIF_KONTRA
            
            if is_kontra:
                # Kontra aktif: doğal alacak bakiyeli, eksi gösterilir
                # Tutar = Alacak - Borç (pozitif değer), bilançoda (-) olarak görünür
                display_value = -net  # Çünkü kontra hesabın bakiyesi alacak yönünde (negatif borç-alacak)
                aktifler.append({
                    'code': code,
                    'name': r['name'],
                    'bakiye': display_value,
                    'is_kontra': True,
                    'sign': '(-)'
                })
                aktif_toplam -= display_value  # Aktif toplamından DÜŞ
            else:
                # Normal aktif: borç bakiyesi pozitif olmalı
                if net != 0:  # Sıfır bakiyeli hesapları gösterme
                    aktifler.append({
                        'code': code,
                        'name': r['name'],
                        'bakiye': net,
                        'is_kontra': False,
                        'sign': ''
                    })
                    aktif_toplam += net
        
        # PASİF GRUBU (Sınıf 3, 4, 5)
        elif first_digit in ('3', '4', '5'):
            is_kontra = code_int in PASIF_KONTRA
            
            # Pasif hesapların doğal bakiyesi alacak → pozitif gösterim için (Alacak - Borç)
            pasif_value = -net  # = (Alacak - Borç)
            
            if is_kontra:
                # 591 Dönem Net Zararı pasifte (-) gösterilir
                pasifler.append({
                    'code': code,
                    'name': r['name'],
                    'bakiye': abs(pasif_value),
                    'is_kontra': True,
                    'sign': '(-)'
                })
                pasif_toplam -= abs(pasif_value)
            else:
                if pasif_value != 0:
                    pasifler.append({
                        'code': code,
                        'name': r['name'],
                        'bakiye': pasif_value,
                        'is_kontra': False,
                        'sign': ''
                    })
                    pasif_toplam += pasif_value
    
    # Hesap kodlarına göre sırala
    aktifler.sort(key=lambda x: float(x['code']))
    pasifler.sort(key=lambda x: float(x['code']))
    
    # Bilanço Denkliği Kontrolü (Aktif = Pasif + Kar/Zarar)
    fark = round(aktif_toplam - pasif_toplam, 2)
    denk = abs(fark) < 0.01
    
    db.close()
    return render_template('bilanco.html', 
                           aktifler=aktifler, 
                           pasifler=pasifler,
                           t_a=aktif_toplam, 
                           t_p=pasif_toplam,
                           fark=fark,
                           denk=denk)

@app.route('/gelir-tablosu')
@login_required
def gelir_tablosu():
    """
    GELİR TABLOSU (9. Gelir Tablosu PDF müfredatı — 7/A yapısı)

    Hiyerarşi:
    ─────────────────────────────────────────────────────────────
    600  Yurtiçi Satışlar
    610  Satıştan İadeler (-)
    611  Satış İskontoları (-)
    ─ = NET SATIŞLAR
    621  STMM (-)
    ─ = BRÜT SATIŞ KÂRI
    631  Paz. Sat. Dağ. Giderleri (-)
    632  Genel Yönetim Giderleri (-)
    ─ = FAALİYET KÂRI/ZARARI
    64x  Diğer Faaliyet Gelirleri
    65x  Diğer Faaliyet Giderleri
    ─ = OLAĞAN KÂRI/ZARARI
    67x  Olağandışı Gelirler
    68x  Olağandışı Giderler
    ─ = DÖNEM KÂRI/ZARARI (690)
    691  Vergi Karşılığı (%22) (-)
    ─ = NET DÖNEM KÂRI/ZARARI (692 → 590/591)
    ─────────────────────────────────────────────────────────────
    """
    db = get_db()
    sql = '''SELECT a.code, a.name,
                    ROUND(SUM(j.credit) - SUM(j.debit), 2) as bakiye
             FROM accounts a
             JOIN journal_entries j ON a.code = j.account_code
             JOIN transactions t ON j.transaction_id = t.id
             WHERE t.user_id = ? AND CAST(a.code AS TEXT) LIKE '6%'
             GROUP BY a.code
             ORDER BY CAST(a.code AS REAL)'''
    data_raw = db.execute(sql, (current_user.id,)).fetchall()

    # Hesap kodu → bakiye haritası
    kod_bakiye = {int(float(r['code'])): float(r['bakiye'] or 0) for r in data_raw}

    def b(kod):
        """Bakiye al, yoksa 0"""
        return kod_bakiye.get(kod, 0)

    # ── Gelir Tablosu Hiyerarşisi ──────────────────────────────
    satis_hasilati = b(600)
    satis_iadeleri = abs(b(610)) + abs(b(611)) + abs(b(612))  # Negatif olan iadeler
    net_satislar   = satis_hasilati - satis_iadeleri

    stmm           = abs(b(621))
    brut_kar       = net_satislar - stmm

    paz_gid        = abs(b(631))
    gyn_gid        = abs(b(632))
    faaliyet_kar   = brut_kar - paz_gid - gyn_gid

    diger_gel      = b(642) + b(643) + b(644) + b(645) + b(646) + b(647) + b(648) + b(649)
    diger_gid      = abs(b(653)) + abs(b(654)) + abs(b(655)) + abs(b(656)) + abs(b(657)) + abs(b(658)) + abs(b(659))
    fin_gid        = abs(b(660)) + abs(b(661))
    olagan_kar     = faaliyet_kar + diger_gel - diger_gid - fin_gid

    olagandusu_gel = b(671) + b(672) + b(679)
    olagandusu_gid = abs(b(680)) + abs(b(681)) + abs(b(689))
    donem_kar      = olagan_kar + olagandusu_gel - olagandusu_gid

    vergi_karsili  = round(donem_kar * 0.22, 2) if donem_kar > 0 else 0
    net_kar        = round(donem_kar - vergi_karsili, 2)

    gt = {
        'satis_hasilati': satis_hasilati,
        'satis_iadeleri': satis_iadeleri,
        'net_satislar':   net_satislar,
        'stmm':           stmm,
        'brut_kar':       brut_kar,
        'paz_gid':        paz_gid,
        'gyn_gid':        gyn_gid,
        'faaliyet_kar':   faaliyet_kar,
        'diger_gel':      diger_gel,
        'diger_gid':      diger_gid,
        'fin_gid':        fin_gid,
        'olagan_kar':     olagan_kar,
        'olagandusu_gel': olagandusu_gel,
        'olagandusu_gid': olagandusu_gid,
        'donem_kar':      donem_kar,
        'vergi_karsili':  vergi_karsili,
        'net_kar':        net_kar,
    }

    # Eski template uyumu için
    gelir_t = sum(r['bakiye'] for r in data_raw if r['bakiye'] and r['bakiye'] > 0)
    gider_t = sum(abs(r['bakiye']) for r in data_raw if r['bakiye'] and r['bakiye'] < 0)

    db.close()
    return render_template('gelir_tablosu.html',
                           data=data_raw, gt=gt,
                           gelir=gelir_t, gider=gider_t, net=gelir_t - gider_t)

# --- KULLANICI AYARLARI VE GÜVENLİK ---
@app.route('/profile')
@login_required
def profile():
    """Kullanıcı profil merkezi ve hızlı istatistikler"""
    db = get_db()
    u_data = db.execute('SELECT * FROM users WHERE id = ?', (current_user.id,)).fetchone()
    stats = db.execute('SELECT COUNT(*) as count FROM transactions WHERE user_id = ?', (current_user.id,)).fetchone()
    logs_data = db.execute('SELECT * FROM audit_logs WHERE user_id = ? ORDER BY id DESC LIMIT 5', (current_user.id,)).fetchall()
    db.close()
    return render_template('profile.html', user=u_data, stats=stats, logs=logs_data)

@app.route('/settings', methods=['GET', 'POST'])
@login_required
def settings():
    """Güvenlik ayarları, şifreleme ve yetkilendirme modülü"""
    db = get_db()
    u_data = db.execute('SELECT * FROM users WHERE id = ?', (current_user.id,)).fetchone()
    if request.method == 'POST':
        new_p = request.form.get('new_password')
        if new_p:
            hashed = generate_password_hash(new_p, method='pbkdf2:sha256')
            db.execute('UPDATE users SET password_hash = ? WHERE id = ?', (hashed, current_user.id))
            db.commit()
            log_action("Güvenlik", "Sistem erişim şifresi başarıyla güncellendi.")
            flash("Güvenlik şifreniz güncellendi. Yeni şifrenizle devam edebilirsiniz.")
            return redirect(url_for('profile'))
    db.close()
    return render_template('settings.html', user=u_data)

# --- app.py içindeki update_settings rotasına ekleme ---
@app.route('/update_settings', methods=['POST'])
@login_required
def update_settings():
    db = get_db()
    # Mevcut e-posta ve resim kodlarını koruyarak yeni verileri çekiyoruz
    email = request.form.get('email')
    full_name = request.form.get('full_name')    # YENİ
    department = request.form.get('department')  # YENİ
    student_no = request.form.get('student_no')  # YENİ
    birth_date = request.form.get('birth_date')  # YENİ
    bio = request.form.get('bio')                # YENİ

    # Veritabanını güncelle (mevcut mantığa ekleme yapıldı)
    db.execute('''UPDATE users SET email = ?, full_name = ?, department = ?, 
                  student_no = ?, birth_date = ?, bio = ? WHERE id = ?''', 
               (email, full_name, department, student_no, birth_date, bio, current_user.id))
    
    # Mevcut profil resmi yükleme kodlarını buraya (altına) aynen ekle...
    file = request.files.get('profile_pic')
    if file and allowed_file(file.filename):
        filename = secure_filename(f"user_{current_user.id}.jpg")
        file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
        db.execute('UPDATE users SET profile_pic = ? WHERE id = ?', (filename, current_user.id))
    
    db.commit()
    db.close()
    flash("Profil bilgileriniz başarıyla detaylandırıldı.")
    return redirect(url_for('profile'))

@app.route('/send_verify_code')
@login_required
def send_verify_code():
    """Siber Güvenlik: 2 adımlı doğrulama kodu üretimi"""
    code = str(random.randint(100000, 999999))
    db = get_db()
    db.execute('UPDATE users SET verification_code = ? WHERE id = ?', (code, current_user.id))
    db.commit()
    db.close()
    flash(f"GÜVENLİK: Yeni doğrulama kodunuz oluşturuldu ve sisteme mühürlendi: {code}")
    return redirect(url_for('settings'))

@app.route('/verify_now', methods=['POST'])
@login_required
def verify_now():
    """Güvenlik kodu doğrulaması ve kullanıcı yetkilendirme"""
    inp = request.form.get('code')
    db = get_db()
    user = db.execute('SELECT verification_code FROM users WHERE id = ?', (current_user.id,)).fetchone()
    if user and user['verification_code'] == inp:
        db.execute('UPDATE users SET is_verified = 1 WHERE id = ?', (current_user.id,))
        db.commit()
        flash("Hesabınız başarıyla doğrulandı ve kurumsal yetki seviyesi yükseltildi.")
    else:
        flash("Hatalı güvenlik kodu girişi yapıldı!")
    db.close()
    return redirect(url_for('profile'))

@app.route('/logs')
@login_required
def logs():
    """Audit Log: Tüm sistem geçmişinin listelenmesi"""
    db = get_db()
    data = db.execute('SELECT * FROM audit_logs WHERE user_id = ? ORDER BY id DESC', (current_user.id,)).fetchall()
    db.close()
    return render_template('logs.html', logs=data)


# ============================================================
# S4.4 — EĞİTİM MODU: SINAV SENARYOSU DOĞRULAYICI
# ============================================================
# Öğrenci yevmiye kaydı girer, sistem müfredat uyumunu kontrol eder.
# YBS Kısa Sınavı + Deneme Sınavları + 8. Hafta Soruları için.
# ============================================================
@app.route('/egitim')
@login_required
def egitim():
    """Eğitim Modu Ana Sayfası — Senaryolar listesi"""
    return render_template('egitim.html')


@app.route('/egitim/dogrula', methods=['POST'])
@login_required
def egitim_dogrula():
    """
    Öğrencinin girdiği yevmiye kaydını müfredat anahtarıyla karşılaştırır.
    
    Doğrulama mantığı:
    1. Öğrenci borç/alacak hesap kodları + tutarları girer
    2. Sistem senaryo ID'sine göre beklenen cevabı bilir
    3. Borç=Alacak dengesi + doğru hesap kodları + doğru tutarlar → ✅
    4. Yanlışlar için müfredat referansıyla açıklamalı feedback döner
    """
    senaryo_id = request.form.get('senaryo_id', '')
    borc_codes  = request.form.getlist('borc_code') or request.form.getlist('borc_code[]')
    borc_tutars = request.form.getlist('borc_tutar') or request.form.getlist('borc_tutar[]')
    alacak_codes  = request.form.getlist('alacak_code') or request.form.getlist('alacak_code[]')
    alacak_tutars = request.form.getlist('alacak_tutar') or request.form.getlist('alacak_tutar[]')

    # Öğrencinin cevabını parse et
    ogrenci_borc   = {int(c): money(t) for c, t in zip(borc_codes, borc_tutars)
                      if c and t and money(t) > 0}
    ogrenci_alacak = {int(c): money(t) for c, t in zip(alacak_codes, alacak_tutars)
                      if c and t and money(t) > 0}

    # Müfredat cevap anahtarı (8. hafta + Kısa Sınav sorularından)
    SENARYO_ANAHTARI = {
        # 8. hafta S1 — %18 KDV'li mal alımı (yarı peşin/yarı veresiye)
        'S8_1': {
            'aciklama': '%18 KDV\'li 13.000 TL mal alımı (7.150 peşin, 6.500 veresiye)',
            'kaynak': '5. Stoklar.pdf + 6. KDV.pdf',
            'borc':   {153: 13000.0, 191: 2340.0},
            'alacak': {100: 8587.0, 320: 6753.0},
            'tolerans': 1.0,
        },
        # 8. hafta S2 — Sürekli Envanter satış (%18 KDV dahil, senet+çek)
        'S8_2': {
            'aciklama': '1.000 TL maliyetli mal 2.360 TL\'ye satış (%18 KDV dahil)',
            'kaynak': '8. hafta Çalışma Soruları',
            'borc':   {100: 360.0, 101: 1000.0, 121: 1000.0, 621: 1000.0},
            'alacak': {600: 2000.0, 391: 360.0, 153: 1000.0},
            'tolerans': 1.0,
        },
        # Kısa sınav S1 — Açılış bilançosu sermaye hesaplama
        'KS_1': {
            'aciklama': 'YBS Kısa Sınavı: Sermaye = 965.000 TL açılış kaydı',
            'kaynak': 'YBS Genel Muhasebe Kısa Sınavı',
            'borc':   {100: 105000.0, 102: 220000.0, 121: 15000.0,
                       153: 125000.0, 254: 500000.0, 255: 270000.0},
            'alacak': {300: 250000.0, 321: 20000.0, 500: 965000.0},
            'tolerans': 1.0,
        },
        # 8. hafta S3 — Hisse senedi alış
        'S8_3': {
            'aciklama': 'MOZAİK A.Ş. 5 adet hisse senedi alımı (50 TL/adet + 100 TL komisyon)',
            'kaynak': '8. hafta Çalışma Soruları + 4. Menkul Kıymetler',
            'borc':   {110: 250.0, 653: 100.0},
            'alacak': {102: 350.0},
            'tolerans': 0.5,
        },
        # Amortisman — Normal yöntem
        'AMR_1': {
            'aciklama': '30.000 TL demirbaş, 5 yıl Normal Amortisman (yıllık 6.000 TL)',
            'kaynak': '8. Amortisman.pdf',
            'borc':   {770: 6000.0},
            'alacak': {257: 6000.0},
            'tolerans': 0.5,
        },
    }

    if senaryo_id not in SENARYO_ANAHTARI:
        return {'hata': f'Bilinmeyen senaryo: {senaryo_id}'}, 400

    senaryo = SENARYO_ANAHTARI[senaryo_id]
    beklenen_borc   = senaryo['borc']
    beklenen_alacak = senaryo['alacak']
    tolerans = senaryo.get('tolerans', 1.0)

    sonuclar = []
    toplam_puan = 0
    max_puan = 0

    # Borç=Alacak dengesi kontrolü
    ogrenci_borc_top   = sum(ogrenci_borc.values())
    ogrenci_alacak_top = sum(ogrenci_alacak.values())
    beklenen_borc_top  = sum(beklenen_borc.values())

    max_puan += 20
    if abs(ogrenci_borc_top - ogrenci_alacak_top) < 0.01:
        toplam_puan += 20
        sonuclar.append({'tip': 'ok', 'mesaj': f'✅ Borç = Alacak dengesi sağlandı ({ogrenci_borc_top:,.2f} TL)'})
    else:
        sonuclar.append({'tip': 'hata', 'mesaj': f'❌ Borç ({ogrenci_borc_top:,.2f}) ≠ Alacak ({ogrenci_alacak_top:,.2f}) — Fark: {abs(ogrenci_borc_top-ogrenci_alacak_top):,.2f} TL'})

    # Borç hesapları kontrolü
    for kod, beklenen_tutar in beklenen_borc.items():
        max_puan += 20
        ogrenci_tutar = ogrenci_borc.get(kod, 0)
        if abs(ogrenci_tutar - beklenen_tutar) <= tolerans:
            toplam_puan += 20
            sonuclar.append({'tip': 'ok', 'mesaj': f'✅ {kod} Borç: {ogrenci_tutar:,.2f} TL (beklenen: {beklenen_tutar:,.2f})'})
        elif kod in ogrenci_borc:
            sonuclar.append({'tip': 'yanlis_tutar', 'mesaj': f'⚠️ {kod} Borç: {ogrenci_tutar:,.2f} girdiniz, beklenen: {beklenen_tutar:,.2f} TL'})
        else:
            sonuclar.append({'tip': 'eksik', 'mesaj': f'❌ {kod} hesabı borçlu tarafta eksik (beklenen: {beklenen_tutar:,.2f} TL)'})

    # Alacak hesapları kontrolü
    for kod, beklenen_tutar in beklenen_alacak.items():
        max_puan += 20
        ogrenci_tutar = ogrenci_alacak.get(kod, 0)
        if abs(ogrenci_tutar - beklenen_tutar) <= tolerans:
            toplam_puan += 20
            sonuclar.append({'tip': 'ok', 'mesaj': f'✅ {kod} Alacak: {ogrenci_tutar:,.2f} TL (beklenen: {beklenen_tutar:,.2f})'})
        elif kod in ogrenci_alacak:
            sonuclar.append({'tip': 'yanlis_tutar', 'mesaj': f'⚠️ {kod} Alacak: {ogrenci_tutar:,.2f} girdiniz, beklenen: {beklenen_tutar:,.2f} TL'})
        else:
            sonuclar.append({'tip': 'eksik', 'mesaj': f'❌ {kod} hesabı alacaklı tarafta eksik (beklenen: {beklenen_tutar:,.2f} TL)'})

    # Fazladan girilen hesaplar (müfredatta yok)
    for kod in ogrenci_borc:
        if kod not in beklenen_borc:
            sonuclar.append({'tip': 'fazla', 'mesaj': f'⚠️ {kod} borç tarafında fazla görünüyor (müfredat cevabında yok)'})
    for kod in ogrenci_alacak:
        if kod not in beklenen_alacak:
            sonuclar.append({'tip': 'fazla', 'mesaj': f'⚠️ {kod} alacak tarafında fazla görünüyor'})

    yuzde = round((toplam_puan / max_puan * 100) if max_puan > 0 else 0, 1)

    from flask import jsonify
    return jsonify({
        'senaryo': senaryo['aciklama'],
        'kaynak': senaryo['kaynak'],
        'puan': toplam_puan,
        'max_puan': max_puan,
        'yuzde': yuzde,
        'sonuclar': sonuclar,
        'beklenen': {
            'borc':   {str(k): v for k, v in beklenen_borc.items()},
            'alacak': {str(k): v for k, v in beklenen_alacak.items()}
        }
    })


# ============================================================
# YENİ MODÜL 1 — NAZIM HESAPLAR (9xx Sınıfı)
# Muhasebeye Giriş Notu: "Varlık/kaynak niteliği taşımayan
# koşullu taahhütler"  Her 9xx(B) kaydının 9xx+1(A) karşılığı var.
# ============================================================
@app.route('/nazim_hesap', methods=['GET', 'POST'])
@login_required
def nazim_hesap():
    db = get_db()
    NAZIM_CESITLERI = [
        (900, '900/901 — Teminat Mektupları (Alınan)'),
        (910, '910/911 — Kefil Olunan Borçlar'),
        (920, '920/921 — Müşteri Adına Saklanan Kıymetler'),
        (940, '940/941 — Vadesi Geçmiş Alacaklar'),
        (950, '950/951 — İhracat Taahhütleri'),
    ]
    if request.method == 'POST':
        borc_kodu   = int(request.form.get('borc_kodu', 900))
        alacak_kodu = borc_kodu + 1
        tutar       = money(request.form.get('tutar', 0))
        aciklama    = request.form.get('aciklama', 'Nazım Hesap Kaydı')
        if tutar <= 0:
            flash("Hata: Tutar pozitif olmalıdır.")
            db.close()
            return redirect(url_for('nazim_hesap'))
        cursor = db.cursor()
        cursor.execute('INSERT INTO transactions (user_id, description, source_module) VALUES (?,?,?)',
                       (current_user.id, aciklama, 'Nazım Hesap'))
        t_id = cursor.lastrowid
        cursor.execute('INSERT INTO journal_entries (transaction_id, account_code, debit) VALUES (?,?,?)',
                       (t_id, borc_kodu, tutar))
        cursor.execute('INSERT INTO journal_entries (transaction_id, account_code, credit) VALUES (?,?,?)',
                       (t_id, alacak_kodu, tutar))
        db.commit()
        log_action("Nazım Hesap", f"{borc_kodu}/{alacak_kodu}: {tutar:,.2f} TL | {aciklama}")
        flash(f"✅ Nazım kayıt: {borc_kodu}/{alacak_kodu} = {tutar:,.2f} TL")
    sql = """SELECT a.code, a.name,
                    ROUND(SUM(j.debit),2) as t_borc,
                    ROUND(SUM(j.credit),2) as t_alacak
             FROM accounts a
             JOIN journal_entries j ON a.code = j.account_code
             JOIN transactions t ON j.transaction_id = t.id
             WHERE t.user_id = ? AND CAST(a.code AS TEXT) LIKE '9%'
             GROUP BY a.code ORDER BY a.code"""
    nazim_data = db.execute(sql, (current_user.id,)).fetchall()
    db.close()
    return render_template('nazim.html', data=nazim_data, cesitler=NAZIM_CESITLERI)


@app.route('/nazim_kapat', methods=['POST'])
@login_required
def nazim_kapat():
    borc_kodu   = int(request.form.get('borc_kodu', 900))
    alacak_kodu = borc_kodu + 1
    tutar       = money(request.form.get('tutar', 0))
    aciklama    = request.form.get('aciklama', 'Nazım Hesap Kapatma')
    if tutar <= 0:
        flash("Hata: Tutar pozitif olmalıdır.")
        return redirect(url_for('nazim_hesap'))
    db = get_db()
    cursor = db.cursor()
    cursor.execute('INSERT INTO transactions (user_id, description, source_module) VALUES (?,?,?)',
                   (current_user.id, f'[Kapatma] {aciklama}', 'Nazım Hesap'))
    t_id = cursor.lastrowid
    cursor.execute('INSERT INTO journal_entries (transaction_id, account_code, credit) VALUES (?,?,?)',
                   (t_id, borc_kodu, tutar))
    cursor.execute('INSERT INTO journal_entries (transaction_id, account_code, debit) VALUES (?,?,?)',
                   (t_id, alacak_kodu, tutar))
    db.commit()
    log_action("Nazım Kapatma", f"{borc_kodu}/{alacak_kodu}: {tutar:,.2f} TL")
    flash(f"✅ Nazım hesap kapatıldı: {borc_kodu}/{alacak_kodu}")
    db.close()
    return redirect(url_for('nazim_hesap'))


# ============================================================
# YENİ MODÜL 2 — DÖNEMSELLİK AYARLAMA KAYITLARI
# Muhasebeye Giriş Notu: 180/181/380/381 hesapları
# 4 tür: gid_gelecek | gid_tahakkuk | gel_gelecek | gel_tahakkuk
# ============================================================
@app.route('/donemsellik', methods=['GET', 'POST'])
@login_required
def donemsellik():
    db = get_db()
    if request.method == 'POST':
        tur        = request.form.get('tur')
        tutar      = money(request.form.get('tutar', 0))
        ilgili_hsp = request.form.get('ilgili_hsp', '770')
        aciklama   = request.form.get('aciklama', 'Dönemsellik Ayarlama')
        if tutar <= 0:
            flash("Hata: Tutar pozitif olmalıdır.")
            db.close()
            return redirect(url_for('donemsellik'))
        cursor = db.cursor()
        cursor.execute('INSERT INTO transactions (user_id, description, source_module) VALUES (?,?,?)',
                       (current_user.id, f'[Dönemsellik-{tur}] {aciklama}', 'Dönemsellik'))
        t_id = cursor.lastrowid
        def ins(code, side, amt):
            col = 'debit' if side == 'D' else 'credit'
            cursor.execute(f'INSERT INTO journal_entries (transaction_id, account_code, {col}) VALUES (?,?,?)',
                           (t_id, int(code), round(abs(amt), 2)))
        if tur == 'gid_gelecek':
            ins(180, 'D', tutar); ins(ilgili_hsp, 'C', tutar)
            flash(f"✅ Gelecek Aylara Ait Gider: 180 (B) / {ilgili_hsp} (A) = {tutar:,.2f} TL")
        elif tur == 'gid_tahakkuk':
            ins(ilgili_hsp, 'D', tutar); ins(381, 'C', tutar)
            flash(f"✅ Gider Tahakkuku: {ilgili_hsp} (B) / 381 (A) = {tutar:,.2f} TL")
        elif tur == 'gel_gelecek':
            ins(ilgili_hsp, 'D', tutar); ins(380, 'C', tutar)
            flash(f"✅ Gelecek Aylara Ait Gelir: {ilgili_hsp} (B) / 380 (A) = {tutar:,.2f} TL")
        elif tur == 'gel_tahakkuk':
            ins(181, 'D', tutar); ins(ilgili_hsp, 'C', tutar)
            flash(f"✅ Gelir Tahakkuku: 181 (B) / {ilgili_hsp} (A) = {tutar:,.2f} TL")
        else:
            flash("Hata: Geçersiz dönemsellik türü.")
            db.close()
            return redirect(url_for('donemsellik'))
        db.commit()
        log_action("Dönemsellik", f"{tur}: {tutar:,.2f} TL | {aciklama}")
    hsp_kodlari = [180, 181, 280, 281, 380, 381, 480, 481]
    donem_data = []
    for kod in hsp_kodlari:
        row = db.execute("""SELECT ROUND(SUM(j.debit)-SUM(j.credit),2) as bakiye
                            FROM journal_entries j JOIN transactions t ON j.transaction_id=t.id
                            WHERE t.user_id=? AND j.account_code=?""",
                         (current_user.id, kod)).fetchone()
        bakiye = float(row['bakiye'] or 0)
        if abs(bakiye) > 0.001:
            acc = db.execute('SELECT name FROM accounts WHERE code=?', (kod,)).fetchone()
            donem_data.append({'code': kod, 'name': acc['name'] if acc else str(kod), 'bakiye': bakiye})
    db.close()
    return render_template('donemsellik.html', data=donem_data)


# ============================================================
# YENİ MODÜL 3 — UZUN VADELİ YABANCI KAYNAK (4xx Sınıfı)
# Bilanço pasif şeması: 3.KVYK / 4.UVYK / 5.Öz Kaynak
# İşlemler: UV kredi, KV transfer, kıdem karşılığı, ödeme
# ============================================================
@app.route('/uvyk', methods=['GET', 'POST'])
@login_required
def uvyk():
    db = get_db()
    UV_HESAPLARI = [
        ('400', '400 — UV Banka Kredileri'),
        ('405', '405 — Çıkarılmış Tahviller'),
        ('420', '420 — Satıcılar UV'),
        ('421', '421 — Borç Senetleri UV'),
        ('472', '472 — Kıdem Tazminatı Karşılığı'),
        ('480', '480 — Gelecek Yıllara Ait Gelirler'),
    ]
    if request.method == 'POST':
        tur      = request.form.get('tur')
        tutar    = money(request.form.get('tutar', 0))
        aciklama = request.form.get('aciklama', 'UVYK İşlemi')
        hsp      = request.form.get('hsp', '400')
        if tutar <= 0:
            flash("Hata: Tutar pozitif olmalıdır.")
            db.close()
            return redirect(url_for('uvyk'))
        cursor = db.cursor()
        cursor.execute('INSERT INTO transactions (user_id, description, source_module) VALUES (?,?,?)',
                       (current_user.id, f'[UVYK-{tur}] {aciklama}', 'UVYK'))
        t_id = cursor.lastrowid
        def ins(code, side, amt):
            col = 'debit' if side == 'D' else 'credit'
            cursor.execute(f'INSERT INTO journal_entries (transaction_id, account_code, {col}) VALUES (?,?,?)',
                           (t_id, int(code), round(abs(amt), 2)))
        if tur == 'uv_kredi_al':
            ins(102, 'D', tutar); ins(hsp, 'C', tutar)
            flash(f"✅ UV Kredi: 102 (B) / {hsp} (A) = {tutar:,.2f} TL")
        elif tur == 'kv_transfer':
            kv_hsp = int(hsp) - 100
            ins(hsp, 'D', tutar); ins(kv_hsp, 'C', tutar)
            flash(f"✅ KV Transfer: {hsp} (B) / {kv_hsp} (A) = {tutar:,.2f} TL — Yeniden sınıflandırma")
        elif tur == 'kidem':
            ins(770, 'D', tutar); ins(472, 'C', tutar)
            flash(f"✅ Kıdem Tazminatı Karşılığı: 770 (B) / 472 (A) = {tutar:,.2f} TL")
        elif tur == 'uv_odeme':
            ins(hsp, 'D', tutar); ins(102, 'C', tutar)
            flash(f"✅ UV Borç Ödeme: {hsp} (B) / 102 (A) = {tutar:,.2f} TL")
        else:
            flash("Hata: Geçersiz UVYK işlem türü.")
            db.close()
            return redirect(url_for('uvyk'))
        db.commit()
        log_action("UVYK", f"{tur}: {tutar:,.2f} TL | {aciklama}")
    sql = """SELECT a.code, a.name,
                    ROUND(SUM(j.credit) - SUM(j.debit), 2) as bakiye
             FROM accounts a
             JOIN journal_entries j ON a.code = j.account_code
             JOIN transactions t ON j.transaction_id = t.id
             WHERE t.user_id = ? AND CAST(a.code AS TEXT) LIKE '4%'
             GROUP BY a.code HAVING abs(bakiye) > 0 ORDER BY a.code"""
    uvyk_data = db.execute(sql, (current_user.id,)).fetchall()
    db.close()
    return render_template('uvyk.html', data=uvyk_data, hesaplar=UV_HESAPLARI)

# --- SİSTEM ÇALIŞTIRMA ---
if __name__ == '__main__':
    # Gerekli klasörlerin varlığı kontrol edilir
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    # MacBook Air M4 ve MacOS port yönetimi (5001 tercih edilir)
    app.run(debug=True, port=5001)