# foy-landing — appfoy.com

FÖY'ün tanıtım sitesi. Statik HTML; derleme adımı yok. **`main`'e giren her şey Vercel'de
anında canlıya çıkar** — değişiklikleri önce dalda yap, PR aç, önizlemeyi kontrol et.

## Yayındaki (ziyaretçinin gördüğü) dosyalar

| Yol | Ne |
|---|---|
| `index.html` | Ana sayfa (`appfoy.com/`); `#tanitim-filmleri`, dil grupları `#film-tr` / `#film-en`, film kartları `#film-tanitim-*` / `#film-nakit-*` |
| `portfoy-takip-excel/` · `finansal-ozgurluk-hesaplama/` · `bilesik-getiri-hesaplama/` | Üç ücretsiz araç sayfası (adresleri `sitemap.xml`'de) |
| `gizlilik.html` · `kosullar.html` | Gizlilik Politikası ve Kullanım Koşulları (uygulamadan ve mağaza kayıtlarından bağlı) |
| `foy-portfoy-takip-excel.xlsx` | İndirilen Excel şablonu (`tools/excel-uret.mjs` üretir) |
| `assets/` · `assets/film/` | Görseller ve 45 sn tanıtım filmleri (sürümlü adlı, uzun önbellek: `vercel.json`) |
| `og.png` | Paylaşım görseli (5 sayfa `/og.png` olarak bağlı) |
| `googlec6a11880ca2ccc02.html` | Google Search Console doğrulaması — **silme, kökte kalsın** |
| `robots.txt` · `sitemap.xml` · `vercel.json` | Arama motoru ve Vercel yapılandırması (kökte durmalı) |

## DOKUNULMAZ adresler

Bunlar dışarıdan bağlı; taşınırsa ya da adı değişirse kırılır:
`/` · `/portfoy-takip-excel` · `/finansal-ozgurluk-hesaplama` · `/bilesik-getiri-hesaplama` ·
`/gizlilik.html` · `/kosullar.html` · `/og.png` · `/foy-portfoy-takip-excel.xlsx` ·
`/#film-tanitim-tr` · `/#film-nakit-tr` · `/#film-tanitim-en` · `/#film-nakit-en`
(uygulamadaki Menü → Tanıtım Videolarını İzle sayfası bu dört film kartını açar; karşılığı
`FinansalOzgurluk` deposunda `constants/linkler.js`). `#film-tr` / `#film-en` dil grubu başlıkları da
yerinde kalsın.

## Yerel araçlar (`tools/` — sitede yayınlanmaz, `.vercelignore`)

```
node tools/check-pricing.mjs     # fiyat iddiaları (X AY BEDAVA, %X tasarruf) tutarlı mı
node tools/excel-uret.mjs        # Excel şablonunu yeniden üretir (kökteki .xlsx'in üzerine yazar)
node tools/excel-dogrula.mjs     # üretilen şablonun hesaplarını doğrular
```

Gizlilik/Koşullar metinleri uygulamadaki `constants/texts.js` ile aynı anlamda tutulur.
