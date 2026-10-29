// ═══════════════════════════════════════════════════════════════════════════
// FÖY — NAKİT AKIŞI TAKİP EXCEL ŞABLONU ÜRETİCİSİ
//
// Çalıştır:  node tools/nakit-akisi-excel-uret.mjs
// Çıktı:     foy-nakit-akisi-takip-excel.xlsx  (depo kökü → /foy-nakit-akisi-takip-excel.xlsx)
//
// SIFIR BAĞIMLILIK — portföy şablonuyla (tools/excel-uret.mjs) aynı yaklaşım:
// .xlsx bir ZIP arşividir; XML'leri elle yazıp Node'un kendi `zlib`'iyle paketliyoruz.
//
// ⚠️ ORTAK KOD KOPYALANDI, İMPORT EDİLMEDİ — BİLİNÇLİ KARAR
// ZIP yazıcı ve hücre yardımcıları excel-uret.mjs'den aynen alındı. İmport
// edilemezdi: o dosyanın "doğrudan çalıştırıldım mı" kontrolü
// `process.argv[1].endsWith('excel-uret.mjs')` ve BU dosyanın adı da öyle
// bitiyor ("nakit-akisi-EXCEL-URET.MJS"). İmport etseydik bu betik her
// çalıştığında portföy şablonunu da yeniden yazardı. O dosyayı değiştirmek
// ise yayındaki şablonun bayt bayt aynı kalması şartını riske atardı.
//
// İÇERİK — FÖY uygulamasının Nakit Akışı ekranına sadık:
//   gelir üç kalemdir → aktif gelir (emek) · pasif gelir (kira, temettü vb.)
//   · portföy geliri (değerleme farkı; EKSİ olabilir).
//
// ⚠️ CANLI VERİ YOK — makro yok, dış bağlantı yok, Power Query yok. Her sayıyı
// kullanıcı elle girer; dosya yıllarca bakım istemeden aynı şekilde çalışır.
// ═══════════════════════════════════════════════════════════════════════════

import { deflateRawSync } from 'node:zlib';
import { writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const KOK = join(dirname(fileURLToPath(import.meta.url)), '..');

// ── ZIP YAZICI (excel-uret.mjs'den aynen) ─────────────────────────────────
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
    ch.writeUInt32LE(ofset, 42);       // yerel başlık ofseti
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

// ── XML YARDIMCILARI (excel-uret.mjs'den aynen) ───────────────────────────
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
  // açılışta her şeyi yeniden hesaplıyor.
  if (tip === 'formul') return `<c r="${ref}"${s}><f>${esc(deger)}</f></c>`;
  return `<c r="${ref}"${s}/>`;
}

function satirXml(no, hucreler, ozel = '') {
  const dolu = hucreler.filter(Boolean).join('');
  return `<row r="${no}"${ozel}>${dolu}</row>`;
}

// ── STİLLER ───────────────────────────────────────────────────────────────
// Renkler ve yazı tipleri portföy şablonuyla AYNI (FÖY paleti: navy #121B31,
// gold #C9A65A, cream #F2EFE9). Tek fark sayı biçimlerinde:
//
// ⚠️ EKSİ SAYI YALNIZ RENKLE GÖSTERİLMEZ. Portföy geliri eksi olabilir; biçim
// kodunun ikinci bölümü `[Red]\-…` → eksi değer hem KIRMIZI hem başında
// EKSİ İŞARETİYLE görünür. Renk körü bir kullanıcı ya da siyah-beyaz çıktı
// da "-9.600,00 ₺" okur. Aynı kural yüzde biçiminde de var (pay eksi olabilir).
const STIL = {
  VARSAYILAN: 0,
  BASLIK_BUYUK: 1,   // sayfa başlığı
  ACIKLAMA: 2,       // gri küçük not (kaydırmalı)
  TABLO_BASLIK: 3,   // koyu zemin, beyaz kalın
  GIRIS_METIN: 4,    // kullanıcı yazar — krem zemin (Ay sütunu)
  GIRIS_TL: 5,       // kullanıcı yazar — krem zemin, ₺ (eksi işaretli)
  HESAP_TL: 6,       // formül — beyaz zemin, ₺
  HESAP_YUZDE: 7,    // formül — beyaz zemin, %
  TOPLAM_ETIKET: 8,
  TOPLAM_TL: 9,
  TOPLAM_YUZDE: 10,
  BOLUM: 11,         // altın renkli bölüm başlığı
  ETIKET: 12,        // özet satır etiketi
  DEGER_TL: 13,      // özet büyük değer
  DEGER_YUZDE: 14,
  DEGER_SAYI: 15,    // özet büyük değer — düz sayı (ay adedi)
  ADIM: 16,          // kullanım adımı
  UYARI: 17,         // dürüstlük notu — bordo
  // Tek satırlık, KAYDIRMASIZ not (portföy şablonundaki NOT ile aynı ders:
  // kaydırmalı not sabit yükseklikli satırda ortasından kırpılıyordu).
  NOT: 18,
};

const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="2">
  <numFmt numFmtId="164" formatCode="#,##0.00\\ &quot;₺&quot;;[Red]\\-#,##0.00\\ &quot;₺&quot;"/>
  <numFmt numFmtId="165" formatCode="0.00%;[Red]\\-0.00%"/>
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
<cellXfs count="19">
  <xf numFmtId="0"   fontId="0" fillId="0" borderId="0" xfId="0"/>
  <xf numFmtId="0"   fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
  <xf numFmtId="0"   fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
  <xf numFmtId="0"   fontId="3" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
  <xf numFmtId="0"   fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1"/>
  <xf numFmtId="164" fontId="0" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>
  <xf numFmtId="164" fontId="0" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>
  <xf numFmtId="165" fontId="0" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>
  <xf numFmtId="0"   fontId="4" fillId="4" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
  <xf numFmtId="164" fontId="4" fillId="4" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
  <xf numFmtId="165" fontId="4" fillId="4" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
  <xf numFmtId="0"   fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1"/>
  <xf numFmtId="0"   fontId="7" fillId="0" borderId="0" xfId="0" applyFont="1"/>
  <xf numFmtId="164" fontId="6" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>
  <xf numFmtId="165" fontId="6" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>
  <xf numFmtId="1"   fontId="6" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyAlignment="1"><alignment horizontal="right"/></xf>
  <xf numFmtId="0"   fontId="7" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
  <xf numFmtId="0"   fontId="8" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
  <xf numFmtId="0"   fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="center"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

// ── VERİ ──────────────────────────────────────────────────────────────────
// Örnek aylar — kullanıcı dosyayı açar açmaz çalıştığını görsün diye.
// Kullanım sayfasında "içeriğini temizleyebilirsiniz" diye yazıyor.
// Bilinçli seçimler:
//   • Şubat'ta portföy geliri EKSİ → eksi değerin işaretiyle göründüğü ve
//     toplamı doğru düşürdüğü dosyayla birlikte teslim edilsin.
//   • Mart'ta gider BOŞ → "gider isteğe bağlı" davranışı görünsün: karşılama
//     hücresi boş kalır, ortalama karşılama yalnız gideri girilmiş 2 aydan çıkar.
export const ORNEK = [
  { ay: 'Ocak 2026',  aktif: 85000, pasif: 12500, portfoy: 18400,  gider: 62000 },
  { ay: 'Şubat 2026', aktif: 85000, pasif: 13250, portfoy: -9600,  gider: 58500 },
  { ay: 'Mart 2026',  aktif: 85000, pasif: 14100, portfoy: 22750,  gider: null },
];

export const ILK_VERI_SATIRI = 4;
export const SON_VERI_SATIRI = 27;   // 24 aylık çalışma alanı
export const TOPLAM_SATIRI = SON_VERI_SATIRI + 1;   // 28

// ── SAYFA 1 — AYLIK ───────────────────────────────────────────────────────
// A Ay (metin, girdi) · B Aktif gelir · C Pasif gelir · D Portföy geliri (girdi)
// E Toplam = B+C+D · F Pasif payı = C/E · G Aylık gider (girdi, isteğe bağlı)
// H Pasif gelirin gideri karşılama oranı = C/G
function aylikSayfasi() {
  const B = ILK_VERI_SATIRI, S = SON_VERI_SATIRI, T = TOPLAM_SATIRI;
  const satirlar = [];

  satirlar.push(satirXml(1, [
    hucre(1, 1, { tip: 'metin', deger: 'FÖY — Nakit Akışı Takip Şablonu', stil: STIL.BASLIK_BUYUK }),
  ], ' ht="26" customHeight="1"'));

  satirlar.push(satirXml(2, [
    hucre(1, 2, { tip: 'metin', stil: STIL.NOT,
      deger: 'Krem hücreleri siz doldurun, beyaz hücreleri Excel hesaplar. Portföyünüz o ay değer kaybettiyse eksi yazın (örn. -9600).' }),
  ], ' ht="18" customHeight="1"'));

  const basliklar = ['Ay', 'Aktif gelir', 'Pasif gelir', 'Portföy geliri', 'Toplam gelir',
    'Pasif payı %', 'Aylık gider (isteğe bağlı)', 'Pasif gelirin gideri karşılama %'];
  satirlar.push(satirXml(3, basliklar.map((b, i) =>
    hucre(i + 1, 3, { tip: 'metin', deger: b, stil: STIL.TABLO_BASLIK })), ' ht="36" customHeight="1"'));

  const tl = (sutun, r, v) => hucre(sutun, r, v == null ? { stil: STIL.GIRIS_TL } : { tip: 'sayi', deger: v, stil: STIL.GIRIS_TL });

  for (let r = B; r <= S; r++) {
    const o = ORNEK[r - B];
    satirlar.push(satirXml(r, [
      hucre(1, r, o ? { tip: 'metin', deger: o.ay, stil: STIL.GIRIS_METIN } : { stil: STIL.GIRIS_METIN }),
      tl(2, r, o?.aktif),
      tl(3, r, o?.pasif),
      tl(4, r, o?.portfoy),
      // Üç gelir kalemi de boşsa satır boş görünür (0 değil) — toplamlar bozulmaz.
      hucre(5, r, { tip: 'formul', stil: STIL.HESAP_TL,    deger: `IF(AND(B${r}="",C${r}="",D${r}=""),"",B${r}+C${r}+D${r})` }),
      // ⚠️ `E="" ` kontrolü ŞART, `E<=0` tek başına yetmez: Excel'de metin her
      // sayıdan BÜYÜK sayılır, yani boş toplam ("") `>0` testini GEÇER ve C/""
      // #DEĞER! verir. OR içinde ikisi birlikte → boş ya da sıfır/eksi toplamda "".
      hucre(6, r, { tip: 'formul', stil: STIL.HESAP_YUZDE, deger: `IF(OR(E${r}="",E${r}<=0),"",C${r}/E${r})` }),
      tl(7, r, o?.gider),
      hucre(8, r, { tip: 'formul', stil: STIL.HESAP_YUZDE, deger: `IF(OR(G${r}="",G${r}<=0),"",C${r}/G${r})` }),
    ]));
  }

  // TOPLAM satırı. Yüzde sütunları "toplam" değil, Özet sayfasıyla AYNI sayıdır:
  //   F = toplam pasif ÷ toplam gelir   (Özet'teki pasif payı)
  //   H = gideri girilmiş ayların karşılama oranlarının ortalaması (Özet'teki
  //       "ortalama karşılama"). Aynı dosyada iki farklı "karşılama" sayısı
  //       görmek kullanıcıyı şaşırtırdı; tek tanım var.
  satirlar.push(satirXml(T, [
    hucre(1, T, { tip: 'metin', deger: 'TOPLAM', stil: STIL.TOPLAM_ETIKET }),
    hucre(2, T, { tip: 'formul', stil: STIL.TOPLAM_TL,    deger: `SUM(B${B}:B${S})` }),
    hucre(3, T, { tip: 'formul', stil: STIL.TOPLAM_TL,    deger: `SUM(C${B}:C${S})` }),
    hucre(4, T, { tip: 'formul', stil: STIL.TOPLAM_TL,    deger: `SUM(D${B}:D${S})` }),
    hucre(5, T, { tip: 'formul', stil: STIL.TOPLAM_TL,    deger: `SUM(E${B}:E${S})` }),
    hucre(6, T, { tip: 'formul', stil: STIL.TOPLAM_YUZDE, deger: `IF(E${T}>0,C${T}/E${T},"")` }),
    hucre(7, T, { tip: 'formul', stil: STIL.TOPLAM_TL,    deger: `SUM(G${B}:G${S})` }),
    hucre(8, T, { tip: 'formul', stil: STIL.TOPLAM_YUZDE,
      deger: `IF(COUNTIF(G${B}:G${S},">0")=0,"",AVERAGEIF(G${B}:G${S},">0",H${B}:H${S}))` }),
  ], ' ht="22" customHeight="1"'));

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews><sheetView tabSelected="1" showGridLines="0" workbookViewId="0">
  <pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/>
  <selection pane="bottomLeft" activeCell="A${B}" sqref="A${B}"/>
</sheetView></sheetViews>
<sheetFormatPr defaultRowHeight="16.5"/>
<cols>
  <col min="1" max="1" width="16" customWidth="1"/>
  <col min="2" max="2" width="16" customWidth="1"/>
  <col min="3" max="3" width="16" customWidth="1"/>
  <col min="4" max="4" width="16" customWidth="1"/>
  <col min="5" max="5" width="17" customWidth="1"/>
  <col min="6" max="6" width="12" customWidth="1"/>
  <col min="7" max="7" width="17" customWidth="1"/>
  <col min="8" max="8" width="22" customWidth="1"/>
</cols>
<sheetData>${satirlar.join('')}</sheetData>
<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>
</worksheet>`;
}

// ── SAYFA 2 — ÖZET ────────────────────────────────────────────────────────
export const OZET_ILK_KAYNAK = 12;
export const OZET_TOPLAM = 15;

function ozetSayfasi() {
  const T = TOPLAM_SATIRI, B = ILK_VERI_SATIRI, S = SON_VERI_SATIRI;
  const r = [];
  const A = `'Aylık'`;

  r.push(satirXml(1, [hucre(1, 1, { tip: 'metin', deger: 'Nakit Akışı Özeti', stil: STIL.BASLIK_BUYUK })], ' ht="26" customHeight="1"'));
  r.push(satirXml(2, [hucre(1, 2, { tip: 'metin', stil: STIL.NOT, deger: 'Bu sayfadaki her sayı Aylık sayfasından gelir; buraya elle bir şey yazmanız gerekmez.' })], ' ht="18" customHeight="1"'));

  r.push(satirXml(4, [hucre(1, 4, { tip: 'metin', deger: 'GENEL DURUM', stil: STIL.BOLUM })]));
  // "Dolu ay" = üç gelir kaleminden en az biri girilmiş ay. Toplam sütunu (E)
  // tam da o zaman SAYI, aksi hâlde "" (metin) üretir; COUNT yalnız sayıları
  // sayar. Ay adını önceden yazıp geliri boş bırakmak ortalamayı bozmaz.
  const genel = [
    ['Dolu ay sayısı',                 `COUNT(${A}!E${B}:E${S})`,                         STIL.DEGER_SAYI],
    ['Toplam gelir',                   `${A}!E${T}`,                                      STIL.DEGER_TL],
    ['Aylık ortalama gelir',           `IF(B5=0,"",B6/B5)`,                               STIL.DEGER_TL],
    ['Gideri girilmiş ay sayısı',      `COUNTIF(${A}!G${B}:G${S},">0")`,                  STIL.DEGER_SAYI],
    ['Ortalama karşılama oranı',       `IF(B8=0,"",AVERAGEIF(${A}!G${B}:G${S},">0",${A}!H${B}:H${S}))`, STIL.DEGER_YUZDE],
  ];
  genel.forEach(([etiket, formul, stil], i) => {
    const satir = 5 + i;
    r.push(satirXml(satir, [
      hucre(1, satir, { tip: 'metin', deger: etiket, stil: STIL.ETIKET }),
      hucre(2, satir, { tip: 'formul', deger: formul, stil }),
    ], ' ht="20" customHeight="1"'));
  });
  r.push(satirXml(10, [hucre(1, 10, { tip: 'metin', stil: STIL.NOT,
    deger: 'Karşılama oranı: pasif gelirin o ayki gideri ne kadar karşıladığı. Yalnız gideri girilmiş aylardan hesaplanır.' })], ' ht="18" customHeight="1"'));

  r.push(satirXml(11, [
    hucre(1, 11, { tip: 'metin', deger: 'Gelir kaynağı', stil: STIL.TABLO_BASLIK }),
    hucre(2, 11, { tip: 'metin', deger: 'Toplam',        stil: STIL.TABLO_BASLIK }),
    hucre(3, 11, { tip: 'metin', deger: 'Pay',           stil: STIL.TABLO_BASLIK }),
  ], ' ht="24" customHeight="1"'));

  const kaynaklar = [
    ['Aktif gelir (emek)',              'B'],
    ['Pasif gelir (kira, temettü vb.)', 'C'],
    ['Portföy geliri (değerleme)',      'D'],
  ];
  kaynaklar.forEach(([ad, sutun], i) => {
    const satir = OZET_ILK_KAYNAK + i;
    r.push(satirXml(satir, [
      hucre(1, satir, { tip: 'metin', deger: ad, stil: STIL.GIRIS_METIN }),
      hucre(2, satir, { tip: 'formul', stil: STIL.HESAP_TL,    deger: `${A}!${sutun}${T}` }),
      hucre(3, satir, { tip: 'formul', stil: STIL.HESAP_YUZDE, deger: `IF($B$${OZET_TOPLAM}<=0,"",B${satir}/$B$${OZET_TOPLAM})` }),
    ]));
  });

  r.push(satirXml(OZET_TOPLAM, [
    hucre(1, OZET_TOPLAM, { tip: 'metin', deger: 'TOPLAM', stil: STIL.TOPLAM_ETIKET }),
    hucre(2, OZET_TOPLAM, { tip: 'formul', stil: STIL.TOPLAM_TL,    deger: `SUM(B${OZET_ILK_KAYNAK}:B${OZET_TOPLAM - 1})` }),
    hucre(3, OZET_TOPLAM, { tip: 'formul', stil: STIL.TOPLAM_YUZDE, deger: `IF(B${OZET_TOPLAM}<=0,"",SUM(C${OZET_ILK_KAYNAK}:C${OZET_TOPLAM - 1}))` }),
  ], ' ht="22" customHeight="1"'));
  r.push(satirXml(16, [hucre(1, 16, { tip: 'metin', stil: STIL.NOT,
    deger: 'Portföy geliri eksiyse payı da eksi görünür; o zaman diğer iki payın toplamı %100\'ü aşar.' })], ' ht="18" customHeight="1"'));

  r.push(satirXml(18, [hucre(1, 18, { tip: 'metin', deger: 'DAHA PRATİK BİR YOL', stil: STIL.BOLUM })]));
  r.push(satirXml(19, [hucre(1, 19, { tip: 'metin', stil: STIL.ADIM,
    deger: 'Bu şablonda her ayın gelirini siz girersiniz. Portföy gelirini yatırımlarınızın fiyatından kendisi hesaplayan bir uygulama isterseniz FÖY\'e bakabilirsiniz — hesap açmadan çalışır.' })], ' ht="34" customHeight="1"'));
  r.push(satirXml(20, [hucre(1, 20, { tip: 'metin', deger: 'appfoy.com', stil: STIL.BOLUM })]));

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews>
<sheetFormatPr defaultRowHeight="16.5"/>
<cols>
  <col min="1" max="1" width="34" customWidth="1"/>
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
  yaz(4, '1.  Aylık sayfasında boş bir satıra ayın adını yazın (örn. "Nisan 2026").');
  yaz(5, '2.  Aktif gelir sütununa emeğinizle kazandığınızı girin: maaş, serbest iş, ek iş.');
  yaz(6, '3.  Pasif gelir sütununa çalışmadan gelen geliri girin: kira, temettü, faiz.');
  yaz(7, '4.  Portföy geliri sütununa yatırımlarınızın o ayki değer değişimini girin. Değer düştüyse eksi yazın.');
  yaz(8, '5.  İsterseniz Aylık gider sütununa o ayki harcamanızı girin. Boş bırakabilirsiniz.');
  yaz(9, 'Toplam gelir, Pasif payı ve karşılama sütunlarını Excel kendisi hesaplar. Bu sütunlara yazmayın.', STIL.ADIM, '30');

  yaz(11, 'RENKLER NE ANLAMA GELİYOR', STIL.BOLUM);
  yaz(12, 'Krem zeminli hücreler  →  siz doldurursunuz.');
  yaz(13, 'Beyaz zeminli hücreler →  Excel hesaplar, elle değiştirmeyin.');
  yaz(14, 'Eksi sayılar başında eksi işaretiyle ve kırmızı görünür (örn. -9.600,00 ₺).');

  yaz(16, 'ÖRNEK SATIRLAR', STIL.BOLUM);
  yaz(17, 'Dosyada 3 örnek ay var; şablonun çalıştığını görmeniz için kondu. Kendi aylarınızı girmeden önce bu satırların içeriğini temizleyebilirsiniz (satırı silmeyin, 24 aylık alan daralır).', STIL.ADIM, '30');

  yaz(19, 'PASİF PAYI VE KARŞILAMA', STIL.BOLUM);
  yaz(20, 'Pasif payı: o ayki toplam gelirin ne kadarının pasif gelirden geldiği.');
  yaz(21, 'Karşılama: pasif gelirin o ayki gideri ne kadar karşıladığı. %100 olduğunda pasif gelir gideri tek başına öder.', STIL.ADIM, '30');

  yaz(23, 'SINIRLAR', STIL.BOLUM);
  yaz(24, '•  24 aylık çalışma alanı vardır. Daha fazlası için satır ekleyip formülleri aşağı sürükleyin.');
  yaz(25, '•  Makro yoktur, dış veri bağlantısı yoktur. Dosya internete bağlanmaz; her sayıyı siz girersiniz.');
  yaz(26, '•  Google E-Tablolar ile de açılır; yalnız temel işlevler kullanılır.');
  yaz(27, '•  İşlem defteri tutmaz: tek tek alım-satım ya da temettü kaydı için ayrı bir kayıt gerekir.');

  yaz(29, 'DAHA PRATİK BİR YOL', STIL.BOLUM);
  yaz(30, 'FÖY; aktif ve pasif gelirinizi ay ay saklar, portföy gelirini yatırımlarınızın fiyatından kendisi hesaplar. Hesap açmanız gerekmez, veriler cihazınızda kalır.', STIL.ADIM, '34');
  yaz(31, 'appfoy.com', STIL.BOLUM);

  yaz(33, 'FÖY bir nakit akışı ve portföy takip aracıdır; yatırım danışmanlığı değildir. Bu dosyadaki örnek sayılar gerçek bir kişiyi temsil etmez.', STIL.ACIKLAMA, '28');

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

// fullCalcOnLoad: Excel dosyayı açar açmaz TÜM formülleri yeniden hesaplar;
// dosyaya önbellek değeri yazılmadığı için burada üretilmiş bir sayının
// Excel'in sonucuyla çelişme ihtimali yok.
const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<workbookPr/>
<sheets>
  <sheet name="Aylık"    sheetId="1" r:id="rId1"/>
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
<dc:title>FÖY — Nakit Akışı Takip Excel Şablonu</dc:title>
<dc:creator>FÖY — appfoy.com</dc:creator>
<cp:lastModifiedBy>FÖY</cp:lastModifiedBy>
<dc:description>Ücretsiz nakit akışı takip şablonu: aktif, pasif ve portföy geliri ay ay. Sayılar elle girilir; şablon dış kaynaktan veri çekmez. appfoy.com</dc:description>
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
  { ad: 'xl/worksheets/sheet1.xml',   veri: aylikSayfasi() },
  { ad: 'xl/worksheets/sheet2.xml',   veri: ozetSayfasi() },
  { ad: 'xl/worksheets/sheet3.xml',   veri: kullanimSayfasi() },
];

const HEDEF = join(KOK, 'foy-nakit-akisi-takip-excel.xlsx');

if (process.argv[1] && process.argv[1].endsWith('nakit-akisi-excel-uret.mjs')) {
  const buf = zipYaz(parcalar);
  writeFileSync(HEDEF, buf);
  console.log(`✓ ${HEDEF}`);
  console.log(`  ${parcalar.length} parça · ${(buf.length / 1024).toFixed(1)} KB`);
  console.log(`  ${SON_VERI_SATIRI - ILK_VERI_SATIRI + 1} aylık çalışma alanı · ${ORNEK.length} örnek ay`);
}
