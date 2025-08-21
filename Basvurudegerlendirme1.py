import pandas as pd
from datetime import datetime
import sys
import os
import json

# Komut satırından dosya yolu al
if len(sys.argv) < 4:
    print("Kullanım: python Basvurudegerlendirme1.py <girdi.xlsx> <cikti.xlsx> <kresler.json>")
    sys.exit(1)

girdi_dosyasi = sys.argv[1]
cikti_dosyasi = sys.argv[2]
kres_json_dosyasi = sys.argv[3]

# Kreş bilgilerini JSON'dan al
with open(kres_json_dosyasi, 'r', encoding='utf-8') as f:
    kres_listesi = json.load(f)

kontenjan = {item["KresAdi"]: item["Kontenjan"] for item in kres_listesi}

df = pd.read_excel(girdi_dosyasi, engine='openpyxl')

def yas_hesapla(dogum_tarihi):
    if pd.isna(dogum_tarihi):
        return None
    bugun = datetime.today()
    return bugun.year - dogum_tarihi.year - ((bugun.month, bugun.day) < (dogum_tarihi.month, dogum_tarihi.day))

def puan_ve_elendi_bul(row):
    elendi = False
    elendi_sebep = ""
    puan = 0

    yas = yas_hesapla(row["Öğrenci Doğum Tarihi"])
    if (yas is None) or (yas < 4) or (yas > 5):
        return 0, True, "Yaş uygun değil (4-5 yaş aralığında değil)"

    tuvalet_egitimi = str(row["Öğrenci tuvalet eğitimi var mı?"]).strip().lower()
    if tuvalet_egitimi == "bez kullanıyor":
        return 0, True, "Tuvalet eğitimi yok (Bez kullanıyor)"

    if str(row["Okul deneyimi var mı?"]).strip().lower() == "evet":
        puan += 5

    yetim_durumu = str(row["Öğrenci Yetim veya Öksüz mü?"]).strip().lower()
    if yetim_durumu == "babası ölü" or yetim_durumu == "annesi ölü":
        puan += 10
    elif yetim_durumu == "hem annesi hem babası ölü":
        puan += 20

    gelir = str(row["Aylık Net Gelir?"]).strip().lower()
    gelir_puanlari = {
        "< 10.000": 25,
        "< 20.000": 20,
        "< 30.000": 15,
        "< 40.000": 10,
        "< 50.000": 5,
        "< 60.000": 0
    }
    puan += gelir_puanlari.get(gelir, 0)

    cocuk_sayisi = str(row["Ailedeki Çocuk Sayısı?"]).strip()
    cocuk_puanlari = {
        "1": 0,
        "2": 10,
        "3": 15,
        "4": 20,
        "5+": 25
    }
    puan += cocuk_puanlari.get(cocuk_sayisi, 0)

    konut = str(row["İkamet Edilen Konut"]).strip().lower()
    if konut == "kiracı":
        puan += 20

    return puan, elendi, elendi_sebep

puanlar, elendiler, sebepler = [], [], []
for _, row in df.iterrows():
    p, e, s = puan_ve_elendi_bul(row)
    puanlar.append(p)
    elendiler.append("Evet" if e else "Hayır")
    sebepler.append(s)

df["Puan"] = puanlar
df["Elendi mi?"] = elendiler
df["Elendi Sebebi"] = sebepler
df["Yerleştiği Kreş"] = ""

df_sirali = df[df["Elendi mi?"] == "Hayır"].sort_values("Puan", ascending=False).copy()
tercih_sutunlari = ["1.Kreş Tercihiniz?", "2.Kreş Tercihiniz?", "3.Kreş Tercihiniz?", "4.Kreş Tercihiniz?"]

for idx, row in df_sirali.iterrows():
    yerlesme = False
    for tercih in tercih_sutunlari:
        kres = row[tercih]
        if pd.isna(kres):
            continue
        if kres in kontenjan and kontenjan[kres] > 0:
            df.loc[idx, "Yerleştiği Kreş"] = kres
            kontenjan[kres] -= 1
            yerlesme = True
            break
    if not yerlesme:
        df.loc[idx, "Yerleştiği Kreş"] = "Yerleşemedi"

df.loc[df["Elendi mi?"] == "Evet", "Yerleştiği Kreş"] = "Elendi"

# Sıralama
sirali_df_listesi = []
kresler = list(kontenjan.keys())
for kres in kresler:
    df_kres = df[df["Yerleştiği Kreş"] == kres]
    if not df_kres.empty:
        sirali_df_listesi.append(df_kres)
        sirali_df_listesi.append(pd.DataFrame([[""] * len(df.columns)] * 2, columns=df.columns))

df_elenen = df[df["Yerleştiği Kreş"] == "Elendi"]
if not df_elenen.empty:
    sirali_df_listesi.append(df_elenen)

df_sonuc = pd.concat(sirali_df_listesi, ignore_index=True)
df_sonuc.to_excel(cikti_dosyasi, index=False)
