# NAVIPACK – Sipariş & Hacim Hesaplama Tablosu (v2)

Müşterinin siparişlerini **koli** girerek **palet sayısı, TIR doluluğu, hacim (m³),
ağırlık (kg) ve tutarı (USD)** otomatik gören Excel aracı.

## Dosyalar
- `NAVIPACK_Siparis_Hacim_Hesaplama.xlsx` — müşteriye gönderilecek tablo.
- `generate_workbook.py` — tabloyu üreten Python (openpyxl) betiği. Ürün/fiyat/palet
  verisi değişirse betikteki `P` listesini güncelleyip yeniden çalıştırın.

## Sayfalar
1. **SIPARIS - ORDER** — Müşteri yalnızca sarı **SİPARİŞ (KOLİ)** sütununu doldurur.
   Üstte anlık KPI bandı: **Toplam Hacim, TIR Doluluk %, Gereken TIR adedi, Toplam Tutar**.
   TIR kapasitesi düzenlenebilir (varsayılan 90 m³).
2. **PALET & TIR OZETI** — Yarım paletleri ürün ailesi bazında birleştirir
   (Diamond'lar Diamond'larla, Premium'lar Premium'larla …), **gerçek palet sayısını**
   ve **TIR yüklemesini (hacme göre)** verir.
3. **ACIKLAMA - NOTES** — Kullanım ve mantık açıklaması (TR/EN).

## Temel Kurallar
- Sipariş **KOLİ** cinsinden ve her ürünün **KOLİ/SIRA** (box/layer) değerinin **KATI**
  olmalı (1 palet sırası = 1 tam sıra). Yanlış girişte Excel uyarı verir.
  Örn. Diamond Fork (12/sıra): 12, 24, 36, 48 … → 45 yerine 48 koli.
- Palet: 100×120 cm.
- **TIR (dolum HACME göre):** NAVIPACK standart = **90 m³**, çift dorse = **120 m³**.
- Çok az sayıdaki kalan koliler **"palet üstü dökme"** olarak mevcut paletlerin üstüne
  yüklenebilir (bkz. 07.07.2026 volume tablosu).

## v2'de Neler Değişti
- Veriler **NAVIPACK Price Offer 20.06.2026**'ya güncellendi (tek fiyat kolonu,
  Premium koliler 4000 adet, renklere göre ayrı satırlar).
- Palet dizilimleri artık **resmi tekliften** geliyor — tahmin yok.
- **TIR kapasitesi / doluluk / gereken TIR adedi** KPI'ları eklendi (NAVIPACK 90 m³).
- Palet özeti sayfasına TIR yükleme bloğu eklendi.

## Bulk (Dökme) Karşılaştırması
FERPROM gibi **tamamen dökme** yükleyen firmalarda palet yoktur; yükleme yalnızca
**toplam hacmin TIR'a sığması** ile hesaplanır (FERPROM çift dorse = 120 m³).
NAVIPACK paletli olduğu için bu araç palet + sıra mantığını da içerir.

## Kaynaklar
- Ürün listesi, fiyatlar, palet dizilimleri: NAVIPACK Price Offer 20.06.2026
- Palet doğrulama & palet üstü dökme: NAVIPACK Volume Calculation 07.07.2026
- TIR kapasiteleri: NAVIPACK standart 90 m³, FERPROM çift dorse 120 m³
