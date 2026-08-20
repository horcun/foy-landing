// ═══════════════════════════════════════════════════════════════════════════
// FÖY — PORTFÖY TAKİP EXCEL ŞABLONU ÜRETİCİSİ
//
// Çalıştır:  node tools/excel-uret.mjs
// Çıktı:     foy-portfoy-takip-excel.xlsx  (depo kökü → /foy-portfoy-takip-excel.xlsx)
//
// SIFIR BAĞIMLILIK. Yalnız Node'un kendi `zlib` ve `fs` modülleri kullanılır.
// .xlsx zaten bir ZIP arşividir; içindeki XML'leri elle yazıp paketliyoruz.
// Neden kütüphane yok: bu proje "bir kere yap, uzun süre kullan" ilkesiyle
// kuruldu. npm bağımlılığı eklemek, dosyayı bir daha üretmek istediğimizde
// çalışmayabilecek bir kurulum adımı yaratırdı.
//
// ⚠️ CANLI VERİ YOK — BİLİNÇLİ KARAR (V1)
// Şablon hiçbir dış kaynaktan fiyat çekmez: TEFAS/BIST/Yahoo yok, Power Query
// yok, makro yok, VBA yok. Güncel fiyatı kullanıcı elle girer. Böylece dosya
// bugün üretilir ve yıllarca bakım istemeden çalışır. Dış veri bağlantısı olan
// bir şablon, kaynak ilk değiştiğinde sessizce yanlış sayı göstermeye başlar.
// ═══════════════════════════════════════════════════════════════════════════

import { deflateRawSync } from 'node:zlib';
import { writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const KOK = join(dirname(fileURLToPath(import.meta.url)), '..');

// ── ZIP YAZICI ────────────────────────────────────────────────────────────
const CRC_TABLO = (() => {
  const t = new Int32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c;
  }
  return t;
})();
const crc32 = (buf) => {
  let c = -1;
  for (let i = 0; i < buf.length; i++) c = CRC_TABLO[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ -1) >>> 0;
};

function zipYaz(girdiler) {
  const yerel = [], merkez = [];
  let ofset = 0;
  for (const { ad, veri } of girdiler) {
    const adB = Buffer.from(ad, 'utf8');
    const ham = Buffer.isBuffer(veri) ? veri : Buffer.from(veri, 'utf8');
    const sikis = deflateRawSync(ham, { level: 9 });
    const crc = crc32(ham);

    const lh = Buffer.alloc(30);
    lh.writeUInt32LE(0x04034b50, 0);   // imza
    lh.writeUInt16LE(20, 4);           // gereken sürüm
    lh.writeUInt16LE(0x0800, 6);       // bayrak: UTF-8 dosya adı
    lh.writeUInt16LE(8, 8);            // yöntem: deflate
    lh.writeUInt16LE(0, 10); lh.writeUInt16LE(0x21, 12); // sabit tarih/saat (yeniden üretilebilirlik)
    lh.writeUInt32LE(crc, 14);
    lh.writeUInt32LE(sikis.length, 18);
    lh.writeUInt32LE(ham.length, 22);
    lh.writeUInt16LE(adB.length, 26);
    lh.writeUInt16LE(0, 28);
    yerel.push(lh, adB, sikis);

    const ch = Buffer.alloc(46);
    ch.writeUInt32LE(0x02014b50, 0);
    ch.writeUInt16LE(20, 4); ch.writeUInt16LE(20, 6);
    ch.writeUInt16LE(0x0800, 8); ch.writeUInt16LE(8, 10);
    ch.writeUInt16LE(0, 12); ch.writeUInt16LE(0x21, 14);
    ch.writeUInt32LE(crc, 16);
    ch.writeUInt32LE(sikis.length, 20);
    ch.writeUInt32LE(ham.length, 24);
    ch.writeUInt16LE(adB.length, 28);
    ch.writeUInt32LE(0, 42);           // yerel başlık ofseti (aşağıda düzeltilir)
    ch.writeUInt32LE(ofset, 42);
    merkez.push(ch, adB);

    ofset += lh.length + adB.length + sikis.length;
  }
  const merkezBuf = Buffer.concat(merkez);
  const son = Buffer.alloc(22);
  son.writeUInt32LE(0x06054b50, 0);
  son.writeUInt16LE(girdiler.length, 8);
  son.writeUInt16LE(girdiler.length, 10);
  son.writeUInt32LE(merkezBuf.length, 12);
  son.writeUInt32LE(ofset, 16);
  return Buffer.concat([Buffer.concat(yerel), merkezBuf, son]);
}

// ── XML YARDIMCILARI ──────────────────────────────────────────────────────
const esc = (s) => String(s)
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
  .replace(/"/g, '&quot;').replace(/'/g, '&apos;');

const sutunAdi = (n) => { // 1 → A
  let s = '';
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
};

// Hücre üretimi. tip: 'metin' | 'sayi' | 'formul' | 'bos'
function hucre(sutun, satir, { tip = 'bos', deger = null, stil = 0 }) {
  const ref = `${sutunAdi(sutun)}${satir}`;
  const s = stil ? ` s="${stil}"` : '';
  if (tip === 'metin')  return `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${esc(deger)}</t></is></c>`;
  if (tip === 'sayi')   return `<c r="${ref}"${s}><v>${deger}</v></c>`;
  // Formülde önbellek DEĞERİ yazmıyoruz; workbook `fullCalcOnLoad` ile
  // açılışta her şeyi yeniden hesaplıyor. Böylece burada üretilmiş bir sayı
  // Excel'in sonucuyla asla çelişemez.
  if (tip === 'formul') return `<c r="${ref}"${s}><f>${esc(deger)}</f></c>`;
  return `<c r="${ref}"${s}/>`;
}

function satirXml(no, hucreler, ozel = '') {
  const dolu = hucreler.filter(Boolean).join('');
  return `<row r="${no}"${ozel}>${dolu}</row>`;
}

// ── STİLLER ───────────────────────────────────────────────────────────────
// Renkler FÖY marka paletinden: navy #121B31, gold #C9A65A, cream #F2EFE9.
// Sayfa açık zeminli tutuldu — ekranda da baskıda da okunur.
const STIL = {
  VARSAYILAN: 0,
  BASLIK_BUYUK: 1,   // sayfa başlığı
  ACIKLAMA: 2,       // gri küçük not
  TABLO_BASLIK: 3,   // koyu zemin, beyaz kalın
  GIRIS_METIN: 4,    // kullanıcı yazar — krem zemin
  GIRIS_ADET: 5,
  GIRIS_FIYAT: 6,
  HESAP_TL: 7,       // formül — TL
  HESAP_YUZDE: 8,    // formül — yüzde
  TOPLAM_ETIKET: 9,
  TOPLAM_TL: 10,
  TOPLAM_YUZDE: 11,
  BOLUM: 12,         // altın renkli bölüm başlığı
  ETIKET: 13,        // özet satır etiketi
  DEGER_TL: 14,      // özet büyük değer
  DEGER_YUZDE: 15,
  ADIM: 16,          // kullanım adımı
  UYARI: 17,         // dürüstlük notu — bordo
};

const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="4">
  <numFmt numFmtId="164" formatCode="#,##0.00\\ &quot;₺&quot;"/>
  <numFmt numFmtId="165" formatCode="0.00%"/>
  <numFmt numFmtId="166" formatCode="#,##0.00####"/>
  <numFmt numFmtId="167" formatCode="#,##0.########"/>
</numFmts>
<fonts count="9">
  <font><sz val="11"/><color rgb="FF1A2130"/><name val="Calibri"/></font>
  <font><b/><sz val="18"/><color rgb="FF121B31"/><name val="Calibri"/></font>
  <font><sz val="10"/><color rgb="FF7A8296"/><name val="Calibri"/></font>
  <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
  <font><b/><sz val="11"/><color rgb="FF121B31"/><name val="Calibri"/></font>
  <font><b/><sz val="12"/><color rgb="FF8A6D22"/><name val="Calibri"/></font>
  <font><b/><sz val="14"/><color rgb="FF121B31"/><name val="Calibri"/></font>
  <font><sz val="11"/><color rgb="FF1A2130"/><name val="Calibri"/></font>
  <font><b/><sz val="11"/><color rgb="FF9E2F45"/><name val="Calibri"/></font>
</fonts>
<fills count="6">
  <fill><patternFill patternType="none"/></fill>
  <fill><patternFill patternType="gray125"/></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FF121B31"/><bgColor indexed="64"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFFDF7E9"/><bgColor indexed="64"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFF2EFE9"/><bgColor indexed="64"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="3">
  <border><left/><right/><top/><bottom/><diagonal/></border>
  <border>
    <left style="thin"><color rgb="FFD8D3C7"/></left><right style="thin"><color rgb="FFD8D3C7"/></right>
    <top style="thin"><color rgb="FFD8D3C7"/></top><bottom style="thin"><color rgb="FFD8D3C7"/></bottom><diagonal/>
  </border>
  <border>
    <left style="thin"><color rgb="FFC9A65A"/></left><right style="thin"><color rgb="FFC9A65A"/></right>
    <top style="medium"><color rgb="FFC9A65A"/></top><bottom style="thin"><color rgb="FFC9A65A"/></bottom><diagonal/>
  </border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="18">
  <xf numFmtId="0"   fontId="0" fillId="0" borderId="0" xfId="0"/>
  <xf numFmtId="0"   fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
  <xf numFmtId="0"   fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
  <xf numFmtId="0"   fontId="3" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
  <xf numFmtId="0"   fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1"/>
  <xf numFmtId="167" fontId="0" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>
  <xf numFmtId="166" fontId="0" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>
  <xf numFmtId="164" fontId="0" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>
  <xf numFmtId="165" fontId="0" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>
  <xf numFmtId="0"   fontId="4" fillId="4" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
  <xf numFmtId="164" fontId="4" fillId="4" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
  <xf numFmtId="165" fontId="4" fillId="4" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
  <xf numFmtId="0"   fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1"/>
  <xf numFmtId="0"   fontId="7" fillId="0" borderId="0" xfId="0" applyFont="1"/>
  <xf numFmtId="164" fontId="6" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>
  <xf numFmtId="165" fontId="6" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>
  <xf numFmtId="0"   fontId="7" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
  <xf numFmtId="0"   fontId="8" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

// ── VERİ ──────────────────────────────────────────────────────────────────
export const TURLER = ['BIST Hisse', 'TEFAS Fonu', 'ABD Hisse/ETF', 'Kripto', 'Altın/Emtia', 'Döviz', 'Nakit'];

// Örnek portföy — kullanıcı dosyayı açar açmaz çalıştığını görsün diye.
// Kullanım sayfasında ve web sayfasında "silebilirsiniz" diye yazıyor.
// Bilinçli olarak bir satır ZARARDA (BTC): negatif K/Z'nin doğru görüldüğü
// dosyayla birlikte teslim edilsin.
export const ORNEK = [
  { ad: 'Aselsan',     kod: 'ASELS', tur: 'BIST Hisse',    adet: 150,   alis: 172.40,     guncel: 181.20 },
  { ad: 'İş Emtia Fonu', kod: 'TGE', tur: 'TEFAS Fonu',    adet: 12000, alis: 0.2715,     guncel: 0.293684 },
  { ad: 'Apple',       kod: 'AAPL',  tur: 'ABD Hisse/ETF', adet: 8,     alis: 7450.00,    guncel: 7820.00 },
  { ad: 'Gram Altın',  kod: 'XAU',   tur: 'Altın/Emtia',   adet: 25,    alis: 4180.00,    guncel: 4355.00 },
  { ad: 'Bitcoin',     kod: 'BTC',   tur: 'Kripto',        adet: 0.05,  alis: 3850000.00, guncel: 3610000.00 },
];

export const ILK_VERI_SATIRI = 4;
export const SON_VERI_SATIRI = 43;   // 40 satırlık çalışma alanı
export const TOPLAM_SATIRI = SON_VERI_SATIRI + 1;

// ── SAYFA 1 — PORTFÖY ─────────────────────────────────────────────────────
function portfoySayfasi() {
  const B = ILK_VERI_SATIRI, S = SON_VERI_SATIRI, T = TOPLAM_SATIRI;
  const satirlar = [];

  satirlar.push(satirXml(1, [
    hucre(1, 1, { tip: 'metin', deger: 'FÖY — Portföy Takip Şablonu', stil: STIL.BASLIK_BUYUK }),
  ], ' ht="26" customHeight="1"'));

  satirlar.push(satirXml(2, [
    hucre(1, 2, { tip: 'metin', stil: STIL.ACIKLAMA,
      deger: 'Krem renkli hücreleri siz doldurun; beyaz hücreleri Excel hesaplar. Güncel fiyatlar elle girilir — şablon dışarıdan veri çekmez.' }),
  ], ' ht="18" customHeight="1"'));

  const basliklar = ['Varlık', 'Kod', 'Tür', 'Adet', 'Ort. Alış', 'Maliyet', 'Güncel Fiyat', 'Güncel Değer', 'K/Z', 'K/Z %', 'Ağırlık'];
  satirlar.push(satirXml(3, basliklar.map((b, i) =>
    hucre(i + 1, 3, { tip: 'metin', deger: b, stil: STIL.TABLO_BASLIK })), ' ht="30" customHeight="1"'));

  for (let r = B; r <= S; r++) {
    const o = ORNEK[r - B];
    satirlar.push(satirXml(r, [
      hucre(1, r, o ? { tip: 'metin', deger: o.ad,  stil: STIL.GIRIS_METIN } : { stil: STIL.GIRIS_METIN }),
      hucre(2, r, o ? { tip: 'metin', deger: o.kod, stil: STIL.GIRIS_METIN } : { stil: STIL.GIRIS_METIN }),
      hucre(3, r, o ? { tip: 'metin', deger: o.tur, stil: STIL.GIRIS_METIN } : { stil: STIL.GIRIS_METIN }),
      hucre(4, r, o ? { tip: 'sayi',  deger: o.adet,   stil: STIL.GIRIS_ADET }  : { stil: STIL.GIRIS_ADET }),
      hucre(5, r, o ? { tip: 'sayi',  deger: o.alis,   stil: STIL.GIRIS_FIYAT } : { stil: STIL.GIRIS_FIYAT }),
      hucre(6, r, { tip: 'formul', stil: STIL.HESAP_TL,    deger: `IF(OR(D${r}="",E${r}=""),"",D${r}*E${r})` }),
      hucre(7, r, o ? { tip: 'sayi', deger: o.guncel, stil: STIL.GIRIS_FIYAT } : { stil: STIL.GIRIS_FIYAT }),
      hucre(8, r, { tip: 'formul', stil: STIL.HESAP_TL,    deger: `IF(OR(D${r}="",G${r}=""),"",D${r}*G${r})` }),
      hucre(9, r, { tip: 'formul', stil: STIL.HESAP_TL,    deger: `IF(OR(F${r}="",H${r}=""),"",H${r}-F${r})` }),
      hucre(10, r, { tip: 'formul', stil: STIL.HESAP_YUZDE, deger: `IF(OR(F${r}="",F${r}=0,H${r}=""),"",I${r}/F${r})` }),
      hucre(11, r, { tip: 'formul', stil: STIL.HESAP_YUZDE, deger: `IF(OR(H${r}="",$H$${T}=0,$H$${T}=""),"",H${r}/$H$${T})` }),
    ]));
  }

  satirlar.push(satirXml(T, [
    hucre(1, T, { tip: 'metin', deger: 'TOPLAM', stil: STIL.TOPLAM_ETIKET }),
    hucre(2, T, { stil: STIL.TOPLAM_ETIKET }),
    hucre(3, T, { stil: STIL.TOPLAM_ETIKET }),
    hucre(4, T, { stil: STIL.TOPLAM_ETIKET }),
    hucre(5, T, { stil: STIL.TOPLAM_ETIKET }),
    hucre(6, T, { tip: 'formul', stil: STIL.TOPLAM_TL,     deger: `SUM(F${B}:F${S})` }),
    hucre(7, T, { stil: STIL.TOPLAM_ETIKET }),
    hucre(8, T, { tip: 'formul', stil: STIL.TOPLAM_TL,     deger: `SUM(H${B}:H${S})` }),
    hucre(9, T, { tip: 'formul', stil: STIL.TOPLAM_TL,     deger: `SUM(I${B}:I${S})` }),
    hucre(10, T, { tip: 'formul', stil: STIL.TOPLAM_YUZDE, deger: `IF(F${T}=0,"",I${T}/F${T})` }),
    hucre(11, T, { tip: 'formul', stil: STIL.TOPLAM_YUZDE, deger: `IF(H${T}=0,"",SUM(K${B}:K${S}))` }),
  ], ' ht="22" customHeight="1"'));

  const dogrulama = `<dataValidations count="1">
    <dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="0" sqref="C${B}:C${S}">
      <formula1>"${TURLER.join(',')}"</formula1>
    </dataValidation>
  </dataValidations>`;

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews><sheetView tabSelected="1" showGridLines="0" workbookViewId="0">
  <pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/>
  <selection pane="bottomLeft" activeCell="A${B}" sqref="A${B}"/>
</sheetView></sheetViews>
<sheetFormatPr defaultRowHeight="16.5"/>
<cols>
  <col min="1" max="1" width="24" customWidth="1"/>
  <col min="2" max="2" width="10" customWidth="1"/>
  <col min="3" max="3" width="16" customWidth="1"/>
  <col min="4" max="4" width="12" customWidth="1"/>
  <col min="5" max="5" width="14" customWidth="1"/>
  <col min="6" max="6" width="16" customWidth="1"/>
  <col min="7" max="7" width="14" customWidth="1"/>
  <col min="8" max="8" width="16" customWidth="1"/>
  <col min="9" max="9" width="15" customWidth="1"/>
  <col min="10" max="10" width="10" customWidth="1"/>
  <col min="11" max="11" width="10" customWidth="1"/>
</cols>
<sheetData>${satirlar.join('')}</sheetData>
${dogrulama}
<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>
</worksheet>`;
}

// ── SAYFA 2 — ÖZET ────────────────────────────────────────────────────────
export const OZET_ILK_TUR = 11;
export const OZET_SON_TUR = OZET_ILK_TUR + TURLER.length - 1;   // 17
export const OZET_TOPLAM  = OZET_SON_TUR + 1;                   // 18

function ozetSayfasi() {
  const T = TOPLAM_SATIRI, B = ILK_VERI_SATIRI, S = SON_VERI_SATIRI;
  const r = [];
  const P = `'Portföy'`;

  r.push(satirXml(1, [hucre(1, 1, { tip: 'metin', deger: 'Portföy Özeti', stil: STIL.BASLIK_BUYUK })], ' ht="26" customHeight="1"'));
  r.push(satirXml(2, [hucre(1, 2, { tip: 'metin', stil: STIL.ACIKLAMA, deger: 'Bu sayfadaki her sayı Portföy sayfasından otomatik gelir. Buraya elle bir şey yazmanız gerekmez.' })], ' ht="18" customHeight="1"'));

  r.push(satirXml(4, [hucre(1, 4, { tip: 'metin', deger: 'GENEL DURUM', stil: STIL.BOLUM })]));
  const genel = [
    ['Toplam Maliyet',   `${P}!F${T}`, STIL.DEGER_TL],
    ['Toplam Değer',     `${P}!H${T}`, STIL.DEGER_TL],
    ['Toplam Kâr/Zarar', `${P}!I${T}`, STIL.DEGER_TL],
    ['Toplam K/Z %',     `${P}!J${T}`, STIL.DEGER_YUZDE],
  ];
  genel.forEach(([etiket, formul, stil], i) => {
    const satir = 5 + i;
    r.push(satirXml(satir, [
      hucre(1, satir, { tip: 'metin', deger: etiket, stil: STIL.ETIKET }),
      hucre(2, satir, { tip: 'formul', deger: formul, stil }),
    ], ' ht="20" customHeight="1"'));
  });

  r.push(satirXml(10, [
    hucre(1, 10, { tip: 'metin', deger: 'Tür',     stil: STIL.TABLO_BASLIK }),
    hucre(2, 10, { tip: 'metin', deger: 'Değer',   stil: STIL.TABLO_BASLIK }),
    hucre(3, 10, { tip: 'metin', deger: 'Ağırlık', stil: STIL.TABLO_BASLIK }),
  ], ' ht="24" customHeight="1"'));

  TURLER.forEach((tur, i) => {
    const satir = OZET_ILK_TUR + i;
    r.push(satirXml(satir, [
      hucre(1, satir, { tip: 'metin', deger: tur, stil: STIL.GIRIS_METIN }),
      hucre(2, satir, { tip: 'formul', stil: STIL.HESAP_TL,
        deger: `SUMIF(${P}!$C$${B}:$C$${S},$A${satir},${P}!$H$${B}:$H$${S})` }),
      hucre(3, satir, { tip: 'formul', stil: STIL.HESAP_YUZDE,
        deger: `IF($B$${OZET_TOPLAM}=0,"",B${satir}/$B$${OZET_TOPLAM})` }),
    ]));
  });

  r.push(satirXml(OZET_TOPLAM, [
    hucre(1, OZET_TOPLAM, { tip: 'metin', deger: 'TOPLAM', stil: STIL.TOPLAM_ETIKET }),
    hucre(2, OZET_TOPLAM, { tip: 'formul', stil: STIL.TOPLAM_TL,    deger: `SUM(B${OZET_ILK_TUR}:B${OZET_SON_TUR})` }),
    hucre(3, OZET_TOPLAM, { tip: 'formul', stil: STIL.TOPLAM_YUZDE, deger: `IF(B${OZET_TOPLAM}=0,"",SUM(C${OZET_ILK_TUR}:C${OZET_SON_TUR}))` }),
  ], ' ht="22" customHeight="1"'));

  r.push(satirXml(21, [hucre(1, 21, { tip: 'metin', deger: 'DAHA PRATİK BİR YOL', stil: STIL.BOLUM })]));
  r.push(satirXml(22, [hucre(1, 22, { tip: 'metin', stil: STIL.ADIM,
    deger: 'Bu şablonda güncel fiyatları siz girersiniz. Fiyatları otomatik takip eden, işlem geçmişini ve ortalama maliyeti kendisi tutan bir uygulama isterseniz FÖY\'e bakabilirsiniz — hesap açmadan çalışır.' })], ' ht="34" customHeight="1"'));
  r.push(satirXml(23, [hucre(1, 23, { tip: 'metin', deger: 'appfoy.com', stil: STIL.BOLUM })]));

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews>
<sheetFormatPr defaultRowHeight="16.5"/>
<cols>
  <col min="1" max="1" width="30" customWidth="1"/>
  <col min="2" max="2" width="20" customWidth="1"/>
  <col min="3" max="3" width="14" customWidth="1"/>
</cols>
<sheetData>${r.join('')}</sheetData>
<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>
</worksheet>`;
}

// ── SAYFA 3 — KULLANIM ────────────────────────────────────────────────────
function kullanimSayfasi() {
  const r = [];
  const yaz = (satir, metin, stil = STIL.ADIM, ht = '20') =>
    r.push(satirXml(satir, [hucre(1, satir, { tip: 'metin', deger: metin, stil })], ` ht="${ht}" customHeight="1"`));

  yaz(1, 'Nasıl Kullanılır?', STIL.BASLIK_BUYUK, '26');
  yaz(3, 'BEŞ ADIM', STIL.BOLUM);
  yaz(4, '1.  Portföy sayfasında bir satıra varlığın adını ve kodunu yazın.');
  yaz(5, '2.  Tür sütunundan varlık türünü seçin (hücreye tıklayınca liste açılır).');
  yaz(6, '3.  Adet sütununa kaç adet/lot/pay aldığınızı girin.');
  yaz(7, '4.  Ort. Alış sütununa ortalama alış fiyatınızı girin.');
  yaz(8, '5.  Güncel Fiyat sütununa bugünkü fiyatı girin.');
  yaz(9, 'Maliyet, Güncel Değer, K/Z, K/Z % ve Ağırlık sütunlarını Excel kendisi hesaplar. Bu sütunlara yazmayın.', STIL.ADIM, '30');

  yaz(11, 'RENKLER NE ANLAMA GELİYOR', STIL.BOLUM);
  yaz(12, 'Krem zeminli hücreler  →  siz doldurursunuz.');
  yaz(13, 'Beyaz zeminli hücreler →  Excel hesaplar, elle değiştirmeyin.');

  yaz(15, 'ÖRNEK SATIRLAR', STIL.BOLUM);
  yaz(16, 'Dosyada 5 örnek satır var; şablonun çalıştığını görmeniz için kondu. Kendi portföyünüzü girmeden önce bu satırları silebilirsiniz.', STIL.ADIM, '30');

  yaz(18, 'GÜNCEL FİYATLAR', STIL.BOLUM);
  yaz(19, 'Bu şablon internetten fiyat çekmez. Güncel fiyatı siz girersiniz. Böylece dosya hiçbir dış servise bağlı değildir ve yıllar sonra da aynı şekilde çalışır.', STIL.UYARI, '32');

  yaz(21, 'SINIRLAR', STIL.BOLUM);
  yaz(22, '•  40 satırlık çalışma alanı vardır. Daha fazlası için satır ekleyip formülleri aşağı sürükleyin.');
  yaz(23, '•  Temettü, bedelsiz ve kısmi satış gibi işlemler için ayrı bir kayıt tutmanız gerekir.');
  yaz(24, '•  Fiyatları elle güncellemeyi unuttuğunuzda tablo eski değeri göstermeye devam eder.');

  yaz(26, 'DAHA PRATİK BİR YOL', STIL.BOLUM);
  yaz(27, 'FÖY; fiyatları kendisi günceller, her alım-satımı ve temettüyü kalıcı olarak kaydeder, ortalama maliyeti sizin girdiğiniz gerçek alış fiyatından hesaplar. Hesap açmanız gerekmez, veriler cihazınızda kalır.', STIL.ADIM, '34');
  yaz(28, 'appfoy.com', STIL.BOLUM);

  yaz(30, 'FÖY bir portföy takip aracıdır; yatırım danışmanlığı değildir. Bu dosyadaki örnek sayılar gerçek bir portföyü temsil etmez.', STIL.ACIKLAMA, '28');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews>
<sheetFormatPr defaultRowHeight="16.5"/>
<cols><col min="1" max="1" width="110" customWidth="1"/></cols>
<sheetData>${r.join('')}</sheetData>
<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>
</worksheet>`;
}

// ── PAKETLE ───────────────────────────────────────────────────────────────
const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;

const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;

// fullCalcOnLoad: Excel dosyayı açar açmaz TÜM formülleri yeniden hesaplar.
// Bu yüzden dosyaya önbellek değeri yazmıyoruz — burada üretilmiş bir sayının
// Excel'in sonucuyla çelişme ihtimali sıfırlanıyor.
const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<workbookPr/>
<sheets>
  <sheet name="Portföy"  sheetId="1" r:id="rId1"/>
  <sheet name="Özet"     sheetId="2" r:id="rId2"/>
  <sheet name="Kullanım" sheetId="3" r:id="rId3"/>
</sheets>
<calcPr calcId="0" fullCalcOnLoad="1"/>
</workbook>`;

const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

const coreXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
 xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:title>FÖY — Portföy Takip Excel Şablonu</dc:title>
<dc:creator>FÖY — appfoy.com</dc:creator>
<cp:lastModifiedBy>FÖY</cp:lastModifiedBy>
<dc:description>Ücretsiz portföy takip şablonu. Güncel fiyatlar elle girilir; şablon dış kaynaktan veri çekmez. appfoy.com</dc:description>
</cp:coreProperties>`;

const appXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
 xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
<Application>FÖY</Application>
<Company>appfoy.com</Company>
</Properties>`;

const parcalar = [
  { ad: '[Content_Types].xml',        veri: contentTypes },
  { ad: '_rels/.rels',                veri: rootRels },
  { ad: 'docProps/core.xml',          veri: coreXml },
  { ad: 'docProps/app.xml',           veri: appXml },
  { ad: 'xl/workbook.xml',            veri: workbookXml },
  { ad: 'xl/_rels/workbook.xml.rels', veri: workbookRels },
  { ad: 'xl/styles.xml',              veri: stylesXml },
  { ad: 'xl/worksheets/sheet1.xml',   veri: portfoySayfasi() },
  { ad: 'xl/worksheets/sheet2.xml',   veri: ozetSayfasi() },
  { ad: 'xl/worksheets/sheet3.xml',   veri: kullanimSayfasi() },
];

const HEDEF = join(KOK, 'foy-portfoy-takip-excel.xlsx');

if (process.argv[1] && process.argv[1].endsWith('excel-uret.mjs')) {
  const buf = zipYaz(parcalar);
  writeFileSync(HEDEF, buf);
  console.log(`✓ ${HEDEF}`);
  console.log(`  ${parcalar.length} parça · ${(buf.length / 1024).toFixed(1)} KB`);
  console.log(`  ${SON_VERI_SATIRI - ILK_VERI_SATIRI + 1} satırlık çalışma alanı · ${ORNEK.length} örnek satır · ${TURLER.length} varlık türü`);
}
