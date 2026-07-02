# NAVIPACK – Sipariş & Hacim Hesaplama Tablosu

Müşterinin siparişlerini **koli / adet** girerek **palet sayısı, hacim (m³),
ağırlık (kg) ve tutarı (USD)** otomatik gören Excel aracı.

## Dosyalar
- `NAVIPACK_Siparis_Hacim_Hesaplama.xlsx` — müşteriye gönderilecek tablo.
- `generate_workbook.py` — tabloyu üreten Python (openpyxl) betiği. Ürün/fiyat/palet
  verisi değişirse bu betikteki `P` listesini güncelleyip yeniden çalıştırın.

## Sayfalar
1. **SIPARIS - ORDER** — Müşteri yalnızca sarı **SİPARİŞ (KOLİ)** sütununu doldurur.
   Diğer her şey otomatik hesaplanır. Üstten Fiyat Kademesi (1/2/3) seçilir.
2. **PALET OZETI - PALLETS** — Yarım paletleri ürün ailesi bazında birleştirip
   (Diamond'lar Diamond'larla, Premium'lar Premium'larla …) **gerçek palet
   sayısını** verir.
3. **ACIKLAMA - NOTES** — Kullanım ve mantık açıklaması (TR/EN).

## Temel Kurallar
- Sipariş **KOLİ** cinsinden ve her ürünün **KOLİ/SIRA** değerinin **KATI** olmalı
  (1 palet sırası = 1 tam sıra). Yanlış girişte Excel uyarı verir.
  Örn. Diamond Fork (12/sıra): 12, 24, 36, 48 … → 45 yerine 48 koli.
- Palet: 100×120 cm.
- Turuncu hücreler (Smart Fork, şarap kadehleri): palet dizilimi eldeki volume
  tablolarında yok, **tahmin**dir — fabrikadan teyit edilmelidir.

## Kaynaklar
- Ürün listesi & fiyatlar: NAVIPACK Price Offer 10.06.2026
- Palet dizilimleri: NAVIPACK Volume Calculation 11.08.2025 & 01.09.2025
