# SmartPack Export Dashboard — Data Agent (Veri Boru Hattı)

Bu pipeline, ham fatura verisi + stok taksonomisi girildiğinde dashboard
verisini (`const D`) **StokKodu üzerinden** sıfırdan derler ve
`dashboard.html` içine enjekte eder.

## 🔑 Temel Kural

> **Ürün kimliği (Genel Ürün Adı, Tip, Seri, Renk) SADECE `StokKodu`
> üzerinden, taksonomi tablosundan gelir.**
> Faturadaki `UrunAdi` / `Kategori` sütunları YOK SAYILIR.
> `StokKodu` asla değiştirilmez — o, değişmez birleştirme (join) anahtarıdır.

Yeni veri her beslendiğinde: faturadaki StokKodu → taksonomideki Genel Ürün
Adı ile eşleştirilir. Taksonomide olmayan kodlar **atlanır ve raporlanır**
(sessizce kaybolmaz).

## 📁 Klasör Yapısı

```
SUHEDA/
├── dashboard.html              ← nihai çıktı (D buraya enjekte edilir)
├── data/
│   ├── ham_veri.xlsx           ← fatura satırları (fact table)  [SİZ GÜNCELLERSİNİZ]
│   └── taksonomi.xlsx          ← StokKodu master tablosu         [SİZ GÜNCELLERSİNİZ]
└── pipeline/
    ├── build_dashboard.py      ← ana script (agent)
    ├── README.md               ← bu dosya
    └── reference/              ← enrichment sözlükleri (nadiren değişir)
        ├── product_en.json     ← Genel Ürün Adı → İngilizce ad
        ├── seri_cat.json       ← Seri → seri kategorisi
        ├── country_geo.json    ← Ülke → {lat, lon, İngilizce ad}
        └── exchange_rates.json ← Döviz → USD kuru
```

## 🚀 Kullanım

```bash
# Sadece doğrula (dosyaya yazmaz) — önce bunu çalıştırın
python3 pipeline/build_dashboard.py --check

# Derle ve dashboard.html'i güncelle
python3 pipeline/build_dashboard.py

# Farklı bir çıktıya yaz (mevcut dashboard'a dokunmadan)
python3 pipeline/build_dashboard.py --out dashboard_yeni.html
```

## 📥 Girdi Dosya Formatları

### `data/ham_veri.xlsx` — "📋 Ham Veri (Power BI)" sayfası
Her satır bir fatura kalemi. Kolonlar (sıra önemli):

| # | Kolon | Açıklama |
|---|-------|----------|
| 0 | Tarih | Fatura tarihi (YYYY-MM-DD) |
| 1 | Yil | Yıl |
| 2 | Ay | Ay (1-12) |
| 3 | CariKodu | Müşteri kodu |
| 4 | Unvan | Müşteri adı |
| 5 | Ulke | Ülke (TR) |
| 6 | Ulke_EN | Ülke (EN) |
| 7 | **StokKodu** | **Ürün kodu — join anahtarı** |
| 8 | UrunAdi | (yok sayılır) |
| 9 | Kategori | (yok sayılır) |
| 10 | Miktar | Adet |
| 11 | Birim | AD/KG |
| 12 | BirimFiyat | — |
| 13 | ToplamTutar | Döviz cinsinden tutar |
| 14 | Doviz | USD/EUR/GBP/TRY |
| 15 | USD_Tutar | USD tutar (boşsa kurdan hesaplanır) |

### `data/taksonomi.xlsx` — "STOK TAKSONOMİSİ" sayfası
Her satır bir stok kodu. Kolonlar:

| # | Kolon | Açıklama |
|---|-------|----------|
| 0 | **Stok Kodu** | **Değişmez anahtar** |
| 1 | Ürün Adı (DIA) | DIA'daki ham ad (referans) |
| 2 | **Genel Ürün Adı** | **Dashboard'da görünen ürün adı** |
| 3 | Hammadde | PS/PP/PLA... |
| 4 | Ürün Tipi | Çatal, Kaşık, Tabak... (kategori grafiği) |
| 5 | Seri | Premium, Diamond, Lux... |
| 6 | Renk | Şeffaf, Beyaz, Bej... (renk analizi) |

## ➕ Yeni Ürün / Veri Ekleme

1. **Yeni fatura** geldiğinde: `data/ham_veri.xlsx`'e satırları ekleyin
   (veya DIA'dan yeni export ile değiştirin).
2. **Yeni stok kodu** varsa: `data/taksonomi.xlsx`'e satır ekleyin
   (Stok Kodu + Genel Ürün Adı + Tip + Seri + Renk). Mevcut kodları
   **değiştirmeyin**.
3. `python3 pipeline/build_dashboard.py --check` çalıştırın →
   "taksonomide olmayan stok kodu" uyarısı varsa o kodları taksonomiye ekleyin.
4. Temizse `python3 pipeline/build_dashboard.py` ile derleyin.
5. Yeni ürünün İngilizce adı yoksa `reference/product_en.json`'a ekleyin.

## ✅ Doğrulama Raporu

Her çalıştırmada script şunları raporlar:
- Yıllık ciro (USD) ve toplam
- Ülke / Müşteri / Firma / Ürün sayıları
- **Taksonomide olmayan stok kodları** (eklenmesi gerekenler)
- Eksik İngilizce ad / koordinat uyarıları
