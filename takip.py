import streamlit as st
import pandas as pd
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# --- SAYFA AYARLARI ---
st.set_page_config(page_title="Doktora Süreç Takip", layout="wide")

# --- GÜVENLİK VE GİRİŞ EKRANI (LOGIN) ---
if "giris_yapildi" not in st.session_state:
    st.session_state.giris_yapildi = False

if not st.session_state.giris_yapildi:
    st.title("🔒 Doktora Takip Sistemine Giriş")
    kullanici = st.text_input("Kullanıcı Adı")
    sifre = st.text_input("Şifre", type="password")
    
    if st.button("Giriş Yap", type="primary"):
        # Bilgileri Streamlit Secrets içinden kontrol et
        if kullanici == st.secrets["kimlik"]["kullanici_adi"] and sifre == st.secrets["kimlik"]["sifre"]:
            st.session_state.giris_yapildi = True
            st.rerun()
        else:
            st.error("Hatalı kullanıcı adı veya şifre!")
    st.stop() # Giriş yapılmadıysa kodun geri kalanını çalıştırma!

# === BUNDAN SONRASI SADECE GİRİŞ YAPILINCA ÇALIŞIR ===

# Email bilgilerini artık sol menüden değil, güvenli kasadan (Secrets) alıyoruz
gonderici_mail = st.secrets["email"]["adres"]
gonderici_sifre = st.secrets["email"]["sifre"]

st.title("🎓 Doktora Öğrenci Süreç Takip Sistemi")
st.markdown("Yeterlik, Komite Belirleme, Tez Önerisi ve TİK süreçlerinizi buradan takip edebilirsiniz.")

# Sağ üste Çıkış Yap butonu
if st.button("🚪 Çıkış Yap"):
    st.session_state.giris_yapildi = False
    st.rerun()

# --- 1. EXCEL'İ OKUMA VE HAZIRLIK ---
try:
    df = pd.read_excel("ogrenciler.xlsx", engine="openpyxl")
except FileNotFoundError:
    st.error("HATA: 'ogrenciler.xlsx' dosyası bulunamadı. Lütfen sisteme yüklediğinizden emin olun.")
    st.stop()

tarih_sutunlari = ['Yeterlik Tarihi', 'TİK Komite belirleme tarihi', 'Tez Önerisi Tarihi'] + [f"{i}. Son TİK Tarihi" for i in range(1, 9)]

for sutun in tarih_sutunlari:
    if sutun in df.columns:
        df[sutun] = pd.to_datetime(df[sutun], format='%d.%m.%Y', errors='coerce')

bugun = pd.to_datetime("today")

# --- 2. HEDEF TARİHLERİ HESAPLAMA ---
df['Hedef Komite'] = df['Yeterlik Tarihi'] + pd.DateOffset(days=30)
df['Hedef Tez Önerisi'] = df['Yeterlik Tarihi'] + pd.DateOffset(months=6)
df['Hedef 1. TİK'] = df['Tez Önerisi Tarihi'] + pd.DateOffset(months=6)
for i in range(1, 8):
    df[f'Hedef {i+1}. TİK'] = df[f'{i}. Son TİK Tarihi'] + pd.DateOffset(months=6)

def siradaki_tik_bul(row):
    if pd.isna(row['Tez Önerisi Tarihi']):
        return None, pd.NaT 
    if pd.isna(row['1. Son TİK Tarihi']):
        return "1. TİK", row['Hedef 1. TİK']
    for i in range(1, 8):
        if pd.isna(row[f'{i+1}. Son TİK Tarihi']):
            return f"{i+1}. TİK", row[f'Hedef {i+1}. TİK']
    return "Tüm TİK'ler Bitti", pd.NaT

df[['Beklenen TİK Aşaması', 'Hedef TİK Tarihi']] = df.apply(siradaki_tik_bul, axis=1, result_type="expand")

# --- RENKLENDİRME FONKSİYONU ---
def satiri_renklendir(row):
    if row['Kalan Gün'] < 0:
        return ['color: #ff3333; font-weight: bold'] * len(row)
    else:
        return ['color: #ff9900; font-weight: bold'] * len(row)

eposta_gonderilecekler = []

def eposta_listesine_ekle(row, islem_adi, tarih_str, kalan_gun):
    alicilar = []
    if 'Danışman E-mail' in row and pd.notna(row['Danışman E-mail']):
        alicilar.append(str(row['Danışman E-mail']).strip())
    if 'Ogrenci e-posta' in row and pd.notna(row['Ogrenci e-posta']):
        alicilar.append(str(row['Ogrenci e-posta']).strip())
        
    if len(alicilar) > 0:
        eposta_gonderilecekler.append({
            "danisman": row['Danışman Adı'], 
            "ogrenci": row['Ad Soyad'],
            "alicilar": alicilar,
            "islem": islem_adi, 
            "kalan": kalan_gun, 
            "tarih": tarih_str
        })

# --- 3. UYARI PANELLERİ ---
st.markdown("---")
st.subheader("🚨 Acil Eylem Bekleyen İşlemler")
col1, col2, col3 = st.columns(3)

with col1:
    st.error("👥 Komite Belirleme")
    komite_bekleyenler = df[pd.notna(df['Yeterlik Tarihi']) & pd.isna(df['TİK Komite belirleme tarihi'])].copy()
    if not komite_bekleyenler.empty:
        komite_bekleyenler['Kalan Gün'] = (komite_bekleyenler['Hedef Komite'] - bugun).dt.days
        acil_komite = komite_bekleyenler[komite_bekleyenler['Kalan Gün'] <= 30].copy()
        if not acil_komite.empty:
            acil_komite['Tarih'] = acil_komite['Hedef Komite'].dt.strftime('%d.%m.%Y')
            gosterim = acil_komite[['Ad Soyad', 'Danışman Adı', 'Tarih', 'Kalan Gün']]
            st.dataframe(gosterim.style.apply(satiri_renklendir, axis=1), hide_index=True)
            for _, row in acil_komite.iterrows():
                eposta_listesine_ekle(row, "TİK Komite Belirleme", row['Tarih'], row['Kalan Gün'])
        else:
            st.success("Yaklaşan yok.")
    else:
        st.success("Tüm komiteler belirlenmiş.")

with col2:
    st.warning("📄 Tez Önerisi")
    tez_bekleyenler = df[pd.notna(df['Yeterlik Tarihi']) & pd.isna(df['Tez Önerisi Tarihi'])].copy()
    if not tez_bekleyenler.empty:
        tez_bekleyenler['Kalan Gün'] = (tez_bekleyenler['Hedef Tez Önerisi'] - bugun).dt.days
        acil_tez = tez_bekleyenler[tez_bekleyenler['Kalan Gün'] <= 30].copy()
        if not acil_tez.empty:
            acil_tez['Tarih'] = acil_tez['Hedef Tez Önerisi'].dt.strftime('%d.%m.%Y')
            gosterim = acil_tez[['Ad Soyad', 'Danışman Adı', 'Tarih', 'Kalan Gün']]
            st.dataframe(gosterim.style.apply(satiri_renklendir, axis=1), hide_index=True)
            for _, row in acil_tez.iterrows():
                eposta_listesine_ekle(row, "Tez Önerisi Savunması", row['Tarih'], row['Kalan Gün'])
        else:
            st.success("Yaklaşan yok.")
    else:
        st.success("Tüm tez önerileri verilmiş.")

with col3:
    st.info("📊 TİK Raporu")
    tik_bekleyenler = df[pd.notna(df['Hedef TİK Tarihi'])].copy()
    if not tik_bekleyenler.empty:
        tik_bekleyenler['Kalan Gün'] = (tik_bekleyenler['Hedef TİK Tarihi'] - bugun).dt.days
        acil_tik = tik_bekleyenler[tik_bekleyenler['Kalan Gün'] <= 30].copy()
        if not acil_tik.empty:
            acil_tik['Tarih'] = acil_tik['Hedef TİK Tarihi'].dt.strftime('%d.%m.%Y')
            gosterim = acil_tik[['Ad Soyad', 'Beklenen TİK Aşaması', 'Tarih', 'Kalan Gün']]
            st.dataframe(gosterim.style.apply(satiri_renklendir, axis=1), hide_index=True)
            for _, row in acil_tik.iterrows():
                eposta_listesine_ekle(row, row['Beklenen TİK Aşaması'], row['Tarih'], row['Kalan Gün'])
        else:
            st.success("Yaklaşan yok.")
    else:
        st.success("TİK bekleyen yok.")

# --- E-POSTA GÖNDERME İŞLEMİ ---
st.markdown("---")
st.subheader("✉️ Danışmanlara ve Öğrencilere E-Posta Gönder")

if len(eposta_gonderilecekler) > 0:
    st.write(f"Yukarıdaki tablolarda süresi yaklaşan/geçen toplam **{len(eposta_gonderilecekler)}** işlem için uyarı maili atılacaktır.")
    
    if st.button("🚀 Tüm Uyarı Maillerini Gönder", type="primary"):
        with st.spinner("E-postalar gönderiliyor, lütfen bekleyin..."):
            try:
                server = smtplib.SMTP("smtp.gmail.com", 587)
                server.starttls()
                server.login(gonderici_mail, gonderici_sifre)
                
                basarili = 0
                for bilgi in eposta_gonderilecekler:
                    msg = MIMEMultipart()
                    msg['From'] = gonderici_mail
                    msg['To'] = ", ".join(bilgi['alicilar'])
                    msg['Subject'] = f"Uyarı: {bilgi['ogrenci']} - {bilgi['islem']} Süreci"
                    
                    durum_metni = "gecikmiştir!" if bilgi['kalan'] < 0 else f"{bilgi['kalan']} gün kalmıştır."
                    
                    govde = f"""
                    Sayın {bilgi['danisman']} ve Sevgili Öğrencimiz {bilgi['ogrenci']},
                    
                    {bilgi['ogrenci']} isimli öğrencinin '{bilgi['islem']}' süreci için son işlem tarihi {bilgi['tarih']} olarak hesaplanmıştır. 
                    
                    Bu işlem için süreniz {durum_metni} 
                    
                    Gereğini bilgilerinize arz ederiz.
                    Enstitü / Bölüm Başkanlığı
                    """
                    msg.attach(MIMEText(govde, 'plain', 'utf-8'))
                    server.send_message(msg)
                    basarili += 1
                    
                server.quit()
                st.success(f"✅ Başarıyla {basarili} gruba uyarı e-postası gönderildi.")
                
            except Exception as e:
                st.error(f"❌ Bir hata oluştu: Lütfen Secrets şifrenizi doğru girdiğinizden emin olun. Hata: {e}")
else:
    st.info("Şu an e-posta atılacak acil bir durum bulunmamaktadır.")

st.markdown("---")
st.subheader("📋 Genel Öğrenci Listesi")
gosterim_sutunlari = ['Anabilim', 'Ad Soyad', 'Ogrenci e-posta', 'Yeterlik Tarihi', 'TİK Komite belirleme tarihi', 'Tez Önerisi Tarihi', 'Danışman Adı', 'Danışman E-mail']
gosterilebilir_sutunlar = [col for col in gosterim_sutunlari if col in df.columns]

gosterim_df = df[gosterilebilir_sutunlar].copy()
for sutun in ['Yeterlik Tarihi', 'TİK Komite belirleme tarihi', 'Tez Önerisi Tarihi']:
    if sutun in gosterim_df.columns:
        gosterim_df[sutun] = gosterim_df[sutun].dt.strftime('%d.%m.%Y').fillna('-')
st.dataframe(gosterim_df, use_container_width=True)