import streamlit as st
import pandas as pd
import numpy as np
import os
from io import BytesIO  # Excel dosyasını bellekte tutmak için
import requests
import plotly.express as px
import plotly.graph_objects as go
import time


 
def atama_yap(vardiya_plani_df, personel_listesi):
    personel_programi = {personel: {'Pazartesi': [], 'Salı': [], 'Çarşamba': [], 'Perşembe': [], 'Cuma': [], 'Cumartesi': [], 'Pazar': []} for personel in personel_listesi}
    gunler = vardiya_plani_df.index.tolist()
    saatler = vardiya_plani_df.columns.tolist()

    off_gun_secenekleri = ['Pazartesi', 'Salı', 'Çarşamba', 'Perşembe', 'Cuma']
    personel_off_gunleri = {personel: np.random.choice(off_gun_secenekleri) for personel in personel_listesi}
 
    for personel in personel_listesi:
        off_gun = personel_off_gunleri[personel]
        for gun in gunler:
            if gun == off_gun:
                continue  # OFF günlerinde hiçbir şey yapma, listeyi boş bırak
            else:
                for i in range(len(saatler)):
                    if i + 9 > len(saatler):
                        break
                    if all(vardiya_plani_df.at[gun, saatler[j]] > 0 for j in range(i, i + 9)):
                        personel_programi[personel][gun] = [saatler[j] for j in range(i, i + 9)]
                        for j in range(i, i + 9):
                            vardiya_plani_df.at[gun, saatler[j]] -= 1
                        break

    return personel_programi
    


def sonuclari_excel_olarak_indir(personel_programi):
    tum_personellerin_programi = pd.DataFrame()
    toplam_calisma_saatleri = []
    havuz_personel_listesi = []  # Havuz personel listesi
    
    
    for personel, gunler in personel_programi.items():
        saat_dilimleri = sorted(list(set([saat for gun in gunler.values() for saat in gun])))
        data = {'Personel': personel, 'Gün': [], **{saat: [] for saat in saat_dilimleri}}
        toplam_saat = sum(len(saatler) for saatler in gunler.values())
        toplam_calisma_saatleri.append({'Personel': personel, 'Toplam Çalışma Saati': toplam_saat})
 
        eksik_saat = max(0, 54 - toplam_saat)  # Eksik saat hesaplama
        if toplam_saat < 54:  # Haftalık 63 saat dolduramayanlar için kontrol
            havuz_personel_listesi.append({'Personel': personel, 'Durum': 'Havuz Personel', 'Eksik Saat': eksik_saat})
 
        for gun in ['Pazartesi', 'Salı', 'Çarşamba', 'Perşembe', 'Cuma', 'Cumartesi', 'Pazar']:
            data['Gün'].append(gun)
            for saat in saat_dilimleri:
                data[saat].append('X' if saat in gunler[gun] else '')
       
        personel_df = pd.DataFrame(data)
        tum_personellerin_programi = pd.concat([tum_personellerin_programi, personel_df, pd.DataFrame([['']*(len(saat_dilimleri)+2)], columns=['Personel', 'Gün', *saat_dilimleri])])
    
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        tum_personellerin_programi.to_excel(writer, index=False, sheet_name='Vardiya Planı')
        pd.DataFrame(toplam_calisma_saatleri).to_excel(writer, index=False, sheet_name='Toplam Çalışma Saatleri')
        pd.DataFrame(havuz_personel_listesi).to_excel(writer, index=False, sheet_name='Havuz Personelleri')  # Havuz personelleri sayfasına 'Eksik Saat' sütunu ekle

    processed_data = output.getvalue()
    
    return processed_data



#Uygulama
st.image("https://www.filmon.com.tr/wp-content/uploads/2022/01/divan-logo.png", width=200)
st.title('Smart Shift Planner')

# Kullanıcı adı ve şifre için benzersiz key'ler ekliyoruz
user = st.text_input('Kullanıcı Adı', key='user_input')
password = st.text_input('Şifre', type='password', key='password_input')

if 'login_successful' not in st.session_state:
    st.session_state['login_successful'] = False

# Kullanıcı giriş bilgileri
kullanici_bilgileri = {
    'admin': 'adm2024D*',
    'dogana' : 'adm2024D*',
    'akasyabrasserie' : 'YY9D1jMy*',
    'kalamisbrasserie': 'jCUqvXUL.'
}
    

if st.button('Giriş Yap') or st.session_state['login_successful']:
    # Kullanıcı adının ve şifrenin doğruluğunu kontrol et
    if user in kullanici_bilgileri and password == kullanici_bilgileri[user]:
        st.session_state['login_successful'] = True
        st.success('Giriş başarılı!')
        secim = st.selectbox("Yapmak istediğiniz işlemi seçiniz:", ('Rapor Görüntüle', 'Vardiya Planı Yap'))
        
        # Filtreleme işlemi burada yapılacak
        if user == 'akasyabrasserie':
            location_filter = 'AKASYA BRASSERIE'
        elif user == 'kalamisbrasserie':
            location_filter = 'KALAMIŞ BRASSERIE'
        else:
            location_filter = 'AKASYA BRASSERIE'

        if secim == 'Rapor Görüntüle':
            st.write("Yakın Zamanda Hizmetinizde")
              
        elif secim == 'Vardiya Planı Yap':
            uploaded_personel_listesi = st.file_uploader("Çalışanların Excel dosyasını yükle", type=['xlsx'], key="personel_uploader")
              
            if uploaded_personel_listesi is not None:
                df_uploaded_personel = pd.read_excel(uploaded_personel_listesi, usecols=['Ad Soyad'])

                excel_url = "https://github.com/dijitaldonusum/smartshiftplanproject/raw/main/7_gunluk_vardiya_plani.xlsx"
                excel_location = pd.read_excel(excel_url, header=0, index_col=0) 


                # Kullanıcıya göre LOCATIONNAME sütununu filtrele
                if location_filter:
                    excel_location = excel_location[excel_location['LOCATIONNAME'] == location_filter]
                
                # LOCATIONNAME sütununu gizle
                excel_location = excel_location.drop(columns=['LOCATIONNAME'])
                
                st.write('Yüklenen personel listesi başarıyla alındı.')
                
                st.dataframe(df_uploaded_personel)

                df_vardiya_plani = excel_location
                
                df_vardiya_plani.columns = ['00:00', '01:00', '02:00', '03:00', '04:00', '05:00', '06:00', '07:00', '08:00', '09:00', '10:00', '11:00', '12:00', '13:00', '14:00', '15:00', '16:00', '17:00', '18:00', '19:00', '20:00', '21:00', '22:00', '23:00']
                st.success('7 günlük vardiya planı dosyası başarıyla okundu.')
 
                personel_programi = atama_yap(df_vardiya_plani, df_uploaded_personel['Ad Soyad'].tolist())
                excel_data = sonuclari_excel_olarak_indir(personel_programi)
                

                with st.spinner('Yapay Zeka ile planınız hazırlanıyor...'):
                    time.sleep(10)  # 5 saniye bekleme efekti 
                
                
                st.dataframe(personel_programi)
                st.download_button(label="Atama Sonuçlarını Excel olarak indir",
                                   data=excel_data,
                                   file_name="vardiya_planı.xlsx",
                                   mime="application/vnd.ms-excel")
    else:
        st.session_state['login_successful'] = False
        st.error('Giriş başarısız. Lütfen tekrar deneyin.')

    