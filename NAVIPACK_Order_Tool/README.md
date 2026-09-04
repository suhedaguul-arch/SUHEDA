# Sipariş & Hacim Hesaplama Araçları

İki firma için sipariş + hacim + yükleme hesaplama Excel araçları.

## Dosyalar
| Dosya | Firma | Yükleme |
|---|---|---|
| `NAVIPACK_Siparis_Hacim_Hesaplama.xlsx` | NAVIPACK | **Paletli** (100×120 cm), TIR = 90 m³ |
| `FERPROM_Dokme_Siparis_Hacim.xlsx` | FERPROM | **Dökme / bulk** (palet yok), çift dorse = 120 m³ |
| `generate_workbook.py` | — | NAVIPACK tablosunu üreten betik |
| `generate_ferprom.py` + `ferprom_master.json` | — | FERPROM tablosunu üreten betik + ürün verisi |

---

## NAVIPACK (paletli) — v3
**SIPARIS - ORDER**: Müşteri sarı **SİPARİŞ (KOLİ)** sütununu doldurur. Üstte KPI:
Toplam Hacim, TIR Doluluk %, Gereken TIR adedi, Toplam Tutar. TIR kapasitesi
düzenlenebilir (varsayılan **90 m³**).

**Yükleme tipi** sütunu (yeni): her satır **PALET** veya **DÖKME**.
- **PALET**: Sipariş, ürünün **KOLİ/SIRA** değerinin **KATI** olmalı (tam sıra).
  Örn. Diamond Fork 12/sıra → 12, 24, 36, 48 … Yanlış katta Excel uyarır.
- **DÖKME** (palet üstü dökme / bulk): serbest adet; palet oluşturmaz, sadece
  **hacme** katkı verir (mevcut paletlerin üstüne yüklenir).

**PALET & TIR OZETI**: yarım paletleri aile bazında birleştirir (Diamond↔Diamond,
Premium↔Premium), gerçek palet sayısını + TIR yüklemesini verir; dökme ürünler için
ayrı satır.

**Ekstra ürünler** (NAVIPACK volume tablosunda var, fiyat teklifinde yok): PANDA Ice
Cream Spoon (palet), 95 mm Flat Ice Cream Spoon (dökme), Shot Glass (dökme). Bunların
**FİYAT/1000** hücresi sarıdır — siparişe göre elle girilir.

## FERPROM (dökme / bulk)
FERPROM tamamen dökme yükler — **palet yoktur**. Müşteri her ürün için istediği koli
adedini serbest girer; yükleme yalnızca **toplam hacmin TIR'a sığması** ile hesaplanır
(çift dorse = **120 m³**). 49 ürünlük katalog (FERPROM Price Offer 20.07.2026 +
Proforma 26094).

---

## Temel Fark (neden iki ayrı araç)
- **NAVIPACK paletli** → koli/sıra, palet ve tam-sıra kısıtı var.
- **FERPROM dökme** → sadece hacim önemli, palet mantığı yok.

## Kaynaklar
- NAVIPACK: Price Offer 20.06.2026 + Volume Calculation 07.07.2026 (TIR 90 m³)
- FERPROM: Price Offer 20.07.2026 + Proforma 26094 (çift dorse 120 m³)
