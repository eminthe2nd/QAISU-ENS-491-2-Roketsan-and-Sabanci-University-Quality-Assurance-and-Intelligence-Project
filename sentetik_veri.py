#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Oct 21 12:52:57 2025

@author: yakar
"""

import pandas as pd

dosya_adi = 'Sabitlenmiş_Random_Sentetik1.xlsx' 

# Excel dosyasını okuyarak bir Pandas DataFrame'e yükleme
try:
    df = pd.read_excel(dosya_adi)
    print("Veri seti başarıyla yüklendi.")
    print(f"Satır sayısı: {df.shape[0]}, Sütun sayısı: {df.shape[1]}")
    
except FileNotFoundError:
    print(f"HATA: '{dosya_adi}' dosyası bulunamadı. Lütfen dosya adını ve yolunu kontrol edin.")
    df = None # Hata durumunda df'i None olarak ayarla

# Verinin ilk 5 satırını kontrol edelim
if df is not None:
    print("\nİlk 5 satır:")
    print(df.head())
    
# Sütun isimleri, veri tipleri ve boş olmayan değer sayısını kontrol etme ----
print("\n--- Sütun Bilgileri (df.info()) ---")
df.info()

# Her sütundaki eksik değer sayısını ve yüzdesini hesaplama ---
print("\n--- Eksik Değer Analizi ---")

missing_values_count = df.isnull().sum()
total_rows = len(df)
missing_values_percentage = (missing_values_count / total_rows) * 100

missing_info = pd.DataFrame({
    'Eksik Sayısı': missing_values_count,
    'Eksik Yüzdesi (%)': missing_values_percentage
})

# Sadece eksik değeri olan sütunları gösterme
# Eksik Yüzdesi %0.0'dan büyük olanları filtreleyelim
missing_info_filtered = missing_info[missing_info['Eksik Yüzdesi (%)'] > 0]

# Eğer sadece "Analiz 1" ve "Analiz 2" eksikse, tüm listeyi göstermek daha açıklayıcı olabilir.
if missing_info_filtered.empty:
     print("Temel veri sütunlarında eksik değer bulunmamaktadır.")
else:
     # Yüzdeye göre azalan sırayla yazdırma
     print("Eksik değere sahip sütunlar (Yüzdeye göre sıralı):")
     print(missing_info_filtered.sort_values(by='Eksik Yüzdesi (%)', ascending=False))
     
# Sayısal sütunların özet istatistiklerini hesaplama
print("\n--- Sayısal Sütunlar Özet İstatistikler (df.describe()) ---")
print(df.describe())

# Sadece 'object' tipindeki (kategorik) sütunları seçme----
object_cols = df.select_dtypes(include='object').columns

# Kategorik sütunların özet istatistiklerini hesaplama
print("\n--- Kategorik Sütunlar Özet İstatistikler (df.describe(include='object')) ---")
categorical_summary = df[object_cols].describe()
print(categorical_summary)

# Bu tablo Analiz 1'e eklenecektir.

# Seçilen kritik kategorik sütunların frekans dağılımı--
critical_cols = ['HATA_SINIFI', 'HATA_TURU', 'ONAY_STATUSU', 'PROJE_ADI']

print("\n--- Kritik Kategorik Sütunların Frekans Dağılımı ---")

for col in critical_cols:
    if col in df.columns:
        print(f"\n***** Sütun: {col} *****")
        # Benzersiz değerlerin sayısını gösteren tablo
        print(df[col].value_counts())
        print("-" * 30)

# Analizlerimizi yeni bir Excel dosyasına kaydetme---
ANALIZ_OUTPUT_DOSYA_ADI = 'Analiz_Rapor_Tablolari.xlsx'

# 1. Sayısal Özet (df.describe())
numerical_summary = df.describe()

# 2. Kategorik Özet (df.describe(include='object'))
object_cols = df.select_dtypes(include='object').columns
categorical_summary = df[object_cols].describe()

# 3. Kritik Kategorik Dağılımları (Hata Türü ve Onay Statüsü)
hata_turu_counts = df['HATA_TURU'].value_counts().rename('HATA_TURU_COUNT')
onay_statusu_counts = df['ONAY_STATUSU'].value_counts().rename('ONAY_STATUSU_COUNT')
kritik_counts = pd.concat([hata_turu_counts, onay_statusu_counts], axis=1).fillna(0)


# Yeni bir Excel dosyası oluşturarak analizleri sekmeler halinde yazma
with pd.ExcelWriter(ANALIZ_OUTPUT_DOSYA_ADI, engine='xlsxwriter') as writer:
    numerical_summary.to_excel(writer, sheet_name='Analiz 1_Sayisal_Ozet')
    categorical_summary.to_excel(writer, sheet_name='Analiz 1_Kategorik_Ozet')
    kritik_counts.to_excel(writer, sheet_name='Analiz 1_Kritik_Frekans')
    
print(f"\n✅ Analiz 1 özet tabloları, '{ANALIZ_OUTPUT_DOSYA_ADI}' adında YENİ bir dosyaya yazıldı.")

#---
import matplotlib.pyplot as plt
import seaborn as sns

# Grafiklerin düzgün görünmesi için stil ayarı
sns.set_style("whitegrid")

# Görsel 1: HATA_TURU Dağılımı (Kategorik)
plt.figure(figsize=(10, 6))
sns.countplot(y='HATA_TURU', data=df, order=df['HATA_TURU'].value_counts().index, palette='viridis')
plt.title('HATA_TURU Değişkeninin Frekans Dağılımı')
plt.xlabel('Adet')
plt.ylabel('Hata Türü')
plt.tight_layout()
plt.savefig('Analiz 2_Görsel 1_HATA_TURU_Dagilimi.png')
plt.show()

# Görsel 2: ALIKONULAN_MIKTAR Dağılımı (Sayısal)
plt.figure(figsize=(8, 5))
sns.histplot(df['ALIKONULAN_MIKTAR'], kde=True, bins=20, color='skyblue')
plt.title('ALIKONULAN_MIKTAR Dağılımı (Rastgele Veri Doğrulaması)')
plt.xlabel('Alıkonulan Miktar')
plt.ylabel('Frekans')
plt.tight_layout()
plt.savefig('Analiz 2_Görsel 2_ALIKONULAN_MIKTAR_Dagilimi.png')
plt.show()


# Görsel 3: İki Sayısal Değişken Arasındaki İlişki (Saçılım Grafiği)
plt.figure(figsize=(8, 6))
sns.scatterplot(x='ALIKONULAN_MIKTAR', y='RET_MIKTAR', data=df, alpha=0.6, color='darkred')
plt.title('ALIKONULAN_MIKTAR ve RET_MIKTAR İlişkisi')
plt.xlabel('Alıkonulan Miktar')
plt.ylabel('Ret Miktarı')
plt.tight_layout()
plt.savefig('Analiz 2_Görsel 3_Miktar_Iliskisi.png')
plt.show()

print("\nAnaliz 2 için gerekli görseller PNG dosyaları olarak kaydedildi.")
print("Bu görselleri 'Analiz 2' sekmesine ekleyebilir ve altına yorumlarınızı yazabilirsiniz.")