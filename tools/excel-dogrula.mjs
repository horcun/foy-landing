// ═══════════════════════════════════════════════════════════════════════════
// FÖY — EXCEL ŞABLONU DOĞRULAMA
//
// Çalıştır:  node tools/excel-dogrula.mjs
//
// NEDEN VAR: Bu makinede Excel yok. "Dosyayı ürettim" ile "hesap motoru
// doğru çalışıyor" arasındaki farkı kapatmak için üretilen .xlsx AÇILIR
// (ZIP → XML), her formül OKUNUR ve küçük bir hesap motoruyla ÇALIŞTIRILIR.
// Çıkan sayılar, elle hesaplanmış beklentilerle karşılaştırılır.
//
// Yani doğrulanan şey "kod ne yazdı" değil, "dosyadaki formüller ne üretiyor".
// ⚠️ Sınır: bu Excel'in kendisi değildir. Kanıtladığı şey formüllerin
// referansları ve aritmetiği doğru; kanıtlamadığı şey Excel'in dosyayı
// biçimsel olarak kusursuz açtığı. Onu gerçek Excel/Sheets açılışı gösterir.
// ═══════════════════════════════════════════════════════════════════════════

import { inflateRawSync } from 'node:zlib';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const KOK = join(dirname(fileURLToPath(import.meta.url)), '..');
const DOSYA = join(KOK, 'foy-portfoy-takip-excel.xlsx');

let gecen = 0, kalan = 0; const cikti = [];
const ok = (ad, kosul, detay = '') => {
  if (kosul) { gecen++; cikti.push(`  ✓ ${ad}`); }
  else { kalan++; cikti.push(`  ✗ ${ad}   ${detay}`); }
};
const yakin = (a, b, tol = 1e-6) => typeof a === 'number' && Math.abs(a - b) < tol;

// ── ZIP OKUYUCU ───────────────────────────────────────────────────────────
function zipOku(buf) {
  const dosyalar = {};
  // Merkezi dizin sonunu bul
  let eocd = -1;
  for (let i = buf.length - 22; i >= 0; i--) { if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; } }
  if (eocd < 0) throw new Error('ZIP merkezi dizini bulunamadı');
  const adet = buf.readUInt16LE(eocd + 10);
  let p = buf.readUInt32LE(eocd + 16);
  for (let i = 0; i < adet; i++) {
    if (buf.readUInt32LE(p) !== 0x02014b50) throw new Error('merkezi dizin kaydı bozuk');
    const yontem = buf.readUInt16LE(p + 10);
    const sikisBoy = buf.readUInt32LE(p + 20);
    const adBoy = buf.readUInt16LE(p + 28);
    const ekBoy = buf.readUInt16LE(p + 30);
    const yorumBoy = buf.readUInt16LE(p + 32);
    const yerelOfset = buf.readUInt32LE(p + 42);
    const ad = buf.slice(p + 46, p + 46 + adBoy).toString('utf8');
    // Yerel başlıktan gerçek veri başlangıcını bul
    const lAdBoy = buf.readUInt16LE(yerelOfset + 26);
    const lEkBoy = buf.readUInt16LE(yerelOfset + 28);
    const veriBas = yerelOfset + 30 + lAdBoy + lEkBoy;
    const ham = buf.slice(veriBas, veriBas + sikisBoy);
    dosyalar[ad] = yontem === 8 ? inflateRawSync(ham) : ham;
    p += 46 + adBoy + ekBoy + yorumBoy;
  }
  return dosyalar;
}

// ── SAYFA XML → HÜCRE HARİTASI ────────────────────────────────────────────
const cozXml = (s) => String(s)
  .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
  .replace(/&apos;/g, "'").replace(/&amp;/g, '&');

function sayfaOku(xml) {
  const hucreler = {};
  // ⚠️ Kendinden kapanan hücreler (`<c r="A44" s="9"/>`) ayrı ele alınmak
  // ZORUNDA. İlk yazımda tek desen kullanıldı ve `>` gördüğü yerde gövde
  // aramaya başlayıp BİR SONRAKİ hücrenin içeriğini yuttu; sonuç olarak
  // toplam satırındaki SUM formülü hiç okunmadı ve motor bölme hatası verdi.
  // Bu, testin kendi kusuruydu — şablonun değil.
  for (const m of xml.matchAll(/<c r="([A-Z]+\d+)"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
    const ref = m[1], govde = m[3];
    if (govde === undefined) { hucreler[ref] = { deger: '' }; continue; }
    const f = /<f>([\s\S]*?)<\/f>/.exec(govde);
    const v = /<v>([\s\S]*?)<\/v>/.exec(govde);
    const t = /<is><t[^>]*>([\s\S]*?)<\/t><\/is>/.exec(govde);
    if (f) hucreler[ref] = { formul: cozXml(f[1]) };
    else if (t) hucreler[ref] = { deger: cozXml(t[1]) };
    else if (v) hucreler[ref] = { deger: Number(v[1]) };
    else hucreler[ref] = { deger: '' };
  }
  return hucreler;
}

// ── FORMÜL MOTORU ─────────────────────────────────────────────────────────
// Excel'in tamamı değil; bu şablonun kullandığı alt küme:
//   IF · OR · SUM · SUMIF · + - * / · = <> < > <= >= · hücre ve aralık başvurusu
function motorKur(sayfalar, varsayilanSayfa) {
  const onbellek = new Map();
  const zincir = new Set();

  const sutunNo = (s) => { let n = 0; for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64); return n; };

  function hucreDeger(sayfa, ref) {
    const anahtar = `${sayfa}!${ref}`;
    if (onbellek.has(anahtar)) return onbellek.get(anahtar);
    if (zincir.has(anahtar)) throw new Error('döngüsel başvuru: ' + anahtar);
    const h = (sayfalar[sayfa] || {})[ref];
    if (!h) { onbellek.set(anahtar, ''); return ''; }
    if ('deger' in h) { onbellek.set(anahtar, h.deger); return h.deger; }
    zincir.add(anahtar);
    const sonuc = hesapla(h.formul, sayfa);
    zincir.delete(anahtar);
    onbellek.set(anahtar, sonuc);
    return sonuc;
  }

  function aralikDegerleri(sayfa, bas, son) {
    const [, bs, br] = /^([A-Z]+)(\d+)$/.exec(bas);
    const [, ss, sr] = /^([A-Z]+)(\d+)$/.exec(son);
    const s1 = sutunNo(bs), s2 = sutunNo(ss);
    const r1 = Number(br), r2 = Number(sr);
    const out = [];
    for (let s = Math.min(s1, s2); s <= Math.max(s1, s2); s++) {
      let ad = '', n = s;
      while (n > 0) { const r = (n - 1) % 26; ad = String.fromCharCode(65 + r) + ad; n = Math.floor((n - 1) / 26); }
      for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++) out.push(hucreDeger(sayfa, ad + r));
    }
    return out;
  }

  function hesapla(formul, sayfa) {
    let i = 0;
    const src = formul;
    const bosla = () => { while (i < src.length && src[i] === ' ') i++; };

    function ifade() {
      let sol = toplama();
      bosla();
      while (i < src.length) {
        const iki = src.slice(i, i + 2);
        let op = null;
        if (iki === '<>' || iki === '<=' || iki === '>=') { op = iki; i += 2; }
        else if ('=<>'.includes(src[i])) { op = src[i]; i += 1; }
        else break;
        const sag = toplama();
        const a = sol, b = sag;
        const say = (x) => (x === '' ? 0 : x);
        if (op === '=')  sol = (typeof a === 'string' || typeof b === 'string') ? String(a) === String(b) : a === b;
        else if (op === '<>') sol = (typeof a === 'string' || typeof b === 'string') ? String(a) !== String(b) : a !== b;
        else if (op === '<')  sol = say(a) < say(b);
        else if (op === '>')  sol = say(a) > say(b);
        else if (op === '<=') sol = say(a) <= say(b);
        else if (op === '>=') sol = say(a) >= say(b);
        bosla();
      }
      return sol;
    }

    function toplama() {
      let sol = carpma();
      bosla();
      while (i < src.length && (src[i] === '+' || src[i] === '-')) {
        const op = src[i++];
        const sag = carpma();
        const a = sol === '' ? 0 : sol, b = sag === '' ? 0 : sag;
        sol = op === '+' ? a + b : a - b;
        bosla();
      }
      return sol;
    }

    function carpma() {
      let sol = tekli();
      bosla();
      while (i < src.length && (src[i] === '*' || src[i] === '/')) {
        const op = src[i++];
        const sag = tekli();
        const a = sol === '' ? 0 : sol, b = sag === '' ? 0 : sag;
        if (op === '/' && b === 0) throw new Error('#DIV/0! — ' + src);
        sol = op === '*' ? a * b : a / b;
        bosla();
      }
      return sol;
    }

    function tekli() {
      bosla();
      if (src[i] === '-') { i++; const v = tekli(); return (v === '' ? 0 : v) * -1; }
      return birincil();
    }

    function birincil() {
      bosla();
      if (src[i] === '(') { i++; const v = ifade(); bosla(); i++; return v; }
      if (src[i] === '"') {
        i++; let s = '';
        while (i < src.length) {
          if (src[i] === '"' && src[i + 1] === '"') { s += '"'; i += 2; continue; }
          if (src[i] === '"') { i++; break; }
          s += src[i++];
        }
        return s;
      }
      if (/[0-9.]/.test(src[i])) {
        let s = '';
        while (i < src.length && /[0-9.eE]/.test(src[i])) s += src[i++];
        return Number(s);
      }
      // Fonksiyon veya başvuru
      let jeton = '';
      while (i < src.length && /[A-Za-zÀ-ÿıİğĞüÜşŞöÖçÇ0-9_.$!':]/.test(src[i])) jeton += src[i++];
      bosla();
      if (src[i] === '(') {   // fonksiyon
        i++;
        const args = [];
        bosla();
        if (src[i] === ')') { i++; }
        else {
          for (;;) {
            args.push(argOku());
            bosla();
            if (src[i] === ',') { i++; continue; }
            if (src[i] === ')') { i++; break; }
            throw new Error('argüman ayracı beklendi: ' + src.slice(i, i + 20));
          }
        }
        return fonksiyon(jeton.toUpperCase(), args);
      }
      return basvuru(jeton, sayfa);
    }

    // Argüman: aralık olabilir → ham metni de saklamak gerekiyor
    function argOku() {
      const bas = i;
      let derinlik = 0;
      while (i < src.length) {
        const c = src[i];
        if (c === '(') derinlik++;
        else if (c === ')') { if (derinlik === 0) break; derinlik--; }
        else if (c === ',' && derinlik === 0) break;
        else if (c === '"') { i++; while (i < src.length && !(src[i] === '"' && src[i + 1] !== '"')) i += (src[i] === '"' ? 2 : 1); }
        i++;
      }
      return src.slice(bas, i).trim();
    }

    function altHesapla(metin) {
      const eskiI = i, eskiSrc = src;
      const sonuc = hesapla(metin, sayfa);
      i = eskiI; void eskiSrc;
      return sonuc;
    }

    function aralikMi(metin) {
      return /^(?:'[^']+'!|[A-Za-zÀ-ÿıİğĞüÜşŞöÖçÇ]+!)?\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+$/.test(metin.trim());
    }

    function aralikCoz(metin) {
      let m = metin.trim();
      let sf = sayfa;
      const sm = /^(?:'([^']+)'|([A-Za-zÀ-ÿıİğĞüÜşŞöÖçÇ]+))!(.+)$/.exec(m);
      if (sm) { sf = sm[1] || sm[2]; m = sm[3]; }
      const [b, s] = m.split(':').map(x => x.replace(/\$/g, ''));
      return aralikDegerleri(sf, b, s);
    }

    function fonksiyon(ad, args) {
      if (ad === 'IF') {
        const k = altHesapla(args[0]);
        return k ? altHesapla(args[1]) : (args.length > 2 ? altHesapla(args[2]) : false);
      }
      if (ad === 'OR')  return args.some(a => !!altHesapla(a));
      if (ad === 'AND') return args.every(a => !!altHesapla(a));
      if (ad === 'SUM') {
        let t = 0;
        for (const a of args) {
          const v = aralikMi(a) ? aralikCoz(a) : [altHesapla(a)];
          for (const x of v) if (typeof x === 'number' && isFinite(x)) t += x;
        }
        return t;
      }
      if (ad === 'SUMIF') {
        const alan = aralikCoz(args[0]);
        const olcut = altHesapla(args[1]);
        const toplananAlan = args[2] ? aralikCoz(args[2]) : alan;
        let t = 0;
        for (let k = 0; k < alan.length; k++) {
          if (String(alan[k]) === String(olcut)) {
            const x = toplananAlan[k];
            if (typeof x === 'number' && isFinite(x)) t += x;
          }
        }
        return t;
      }
      throw new Error('desteklenmeyen fonksiyon: ' + ad);
    }

    function basvuru(jeton, sf) {
      let m = jeton;
      let hedefSayfa = sf;
      const sm = /^(?:'([^']+)'|([A-Za-zÀ-ÿıİğĞüÜşŞöÖçÇ]+))!(.+)$/.exec(m);
      if (sm) { hedefSayfa = sm[1] || sm[2]; m = sm[3]; }
      const temiz = m.replace(/\$/g, '');
      if (!/^[A-Z]+\d+$/.test(temiz)) throw new Error('çözülemeyen başvuru: ' + jeton);
      return hucreDeger(hedefSayfa, temiz);
    }

    const sonuc = ifade();
    return sonuc;
  }

  return { hucreDeger, sayfaAdi: varsayilanSayfa };
}

// ── ÇALIŞTIR ──────────────────────────────────────────────────────────────
const zip = zipOku(readFileSync(DOSYA));

cikti.push('PAKET YAPISI');
const beklenenParcalar = ['[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml',
  'xl/_rels/workbook.xml.rels', 'xl/styles.xml',
  'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml', 'xl/worksheets/sheet3.xml'];
for (const p of beklenenParcalar) ok(`parça var: ${p}`, p in zip);

const wb = zip['xl/workbook.xml'].toString('utf8');
ok('üç sayfa tanımlı (Portföy / Özet / Kullanım)',
   /name="Portföy"/.test(wb) && /name="Özet"/.test(wb) && /name="Kullanım"/.test(wb));
ok('açılışta tam yeniden hesaplama açık (fullCalcOnLoad)', /fullCalcOnLoad="1"/.test(wb));

const sayfalar = {
  'Portföy':  sayfaOku(zip['xl/worksheets/sheet1.xml'].toString('utf8')),
  'Özet':     sayfaOku(zip['xl/worksheets/sheet2.xml'].toString('utf8')),
  'Kullanım': sayfaOku(zip['xl/worksheets/sheet3.xml'].toString('utf8')),
};
const { hucreDeger } = motorKur(sayfalar, 'Portföy');
const P = (ref) => hucreDeger('Portföy', ref);
const O = (ref) => hucreDeger('Özet', ref);

// ── ELLE HESAPLANMIŞ BEKLENTİLER ──────────────────────────────────────────
// Kaynak: tools/excel-uret.mjs içindeki ORNEK dizisi. Aşağıdaki sayılar
// script'ten TÜRETİLMEDİ, elle çarpılıp toplandı — testin bağımsızlığı için.
const B = [
  { r: 4, ad: 'ASELS', tur: 'BIST Hisse',    maliyet: 25860,  deger: 27180,     kz: 1320,      kzp: 1320 / 25860 },
  { r: 5, ad: 'TGE',   tur: 'TEFAS Fonu',    maliyet: 3258,   deger: 3524.208,  kz: 266.208,   kzp: 266.208 / 3258 },
  { r: 6, ad: 'AAPL',  tur: 'ABD Hisse/ETF', maliyet: 59600,  deger: 62560,     kz: 2960,      kzp: 2960 / 59600 },
  { r: 7, ad: 'XAU',   tur: 'Altın/Emtia',   maliyet: 104500, deger: 108875,    kz: 4375,      kzp: 4375 / 104500 },
  { r: 8, ad: 'BTC',   tur: 'Kripto',        maliyet: 192500, deger: 180500,    kz: -12000,    kzp: -12000 / 192500 },
];
const TOP_MALIYET = 385718;
const TOP_DEGER   = 382639.208;
const TOP_KZ      = -3078.792;
const T = 44;   // toplam satırı

cikti.push('\nSATIR HESAPLARI (örnek portföy)');
for (const b of B) {
  ok(`${b.ad} · maliyet = adet × ort.alış = ${b.maliyet}`, yakin(P(`F${b.r}`), b.maliyet, 1e-6), `bulunan ${P(`F${b.r}`)}`);
  ok(`${b.ad} · güncel değer = adet × güncel fiyat = ${b.deger}`, yakin(P(`H${b.r}`), b.deger, 1e-6), `bulunan ${P(`H${b.r}`)}`);
  ok(`${b.ad} · K/Z = değer − maliyet = ${b.kz}`, yakin(P(`I${b.r}`), b.kz, 1e-6), `bulunan ${P(`I${b.r}`)}`);
  ok(`${b.ad} · K/Z % = K/Z ÷ maliyet = %${(b.kzp * 100).toFixed(2)}`, yakin(P(`J${b.r}`), b.kzp, 1e-9), `bulunan ${P(`J${b.r}`)}`);
}

cikti.push('\nNEGATİF K/Z');
ok('BTC satırı zararda (K/Z < 0)', P('I8') < 0, `bulunan ${P('I8')}`);
ok('BTC K/Z yüzdesi negatif', P('J8') < 0, `bulunan ${P('J8')}`);

cikti.push('\nTOPLAMLAR');
ok(`toplam maliyet = ${TOP_MALIYET}`, yakin(P(`F${T}`), TOP_MALIYET, 1e-6), `bulunan ${P(`F${T}`)}`);
ok(`toplam değer = ${TOP_DEGER}`,     yakin(P(`H${T}`), TOP_DEGER, 1e-6), `bulunan ${P(`H${T}`)}`);
ok(`toplam K/Z = ${TOP_KZ}`,          yakin(P(`I${T}`), TOP_KZ, 1e-6), `bulunan ${P(`I${T}`)}`);
ok('toplam K/Z, değer − maliyet ile tutarlı',
   yakin(P(`I${T}`), P(`H${T}`) - P(`F${T}`), 1e-6));
ok('toplam K/Z % = toplam K/Z ÷ toplam maliyet',
   yakin(P(`J${T}`), TOP_KZ / TOP_MALIYET, 1e-9), `bulunan ${P(`J${T}`)}`);

cikti.push('\nAĞIRLIKLAR');
let agirlikToplam = 0;
for (const b of B) {
  const bekl = b.deger / TOP_DEGER;
  agirlikToplam += P(`K${b.r}`);
  ok(`${b.ad} · ağırlık = %${(bekl * 100).toFixed(2)}`, yakin(P(`K${b.r}`), bekl, 1e-9), `bulunan ${P(`K${b.r}`)}`);
}
ok('ağırlıklar toplamı tam %100', yakin(agirlikToplam, 1, 1e-9), `bulunan ${agirlikToplam}`);
ok('toplam satırındaki ağırlık da %100', yakin(P(`K${T}`), 1, 1e-9), `bulunan ${P(`K${T}`)}`);

cikti.push('\nBOŞ SATIR DAVRANIŞI');
for (const ref of ['F20', 'H20', 'I20', 'J20', 'K20']) {
  ok(`${ref} boş satırda "" döndürüyor (0 veya hata değil)`, P(ref) === '', `bulunan ${JSON.stringify(P(ref))}`);
}
ok('boş satırlar toplamı bozmuyor', yakin(P(`F${T}`), TOP_MALIYET, 1e-6));

cikti.push('\nSIFIR DEĞER DAVRANIŞI');
// Maliyeti 0 olan bir satır simüle et: E9=0 iken J9 bölme hatası vermemeli.
sayfalar['Portföy']['D9'] = { deger: 10 };
sayfalar['Portföy']['E9'] = { deger: 0 };
sayfalar['Portföy']['G9'] = { deger: 5 };
{
  const { hucreDeger: hd2 } = motorKur(sayfalar, 'Portföy');
  const f9 = hd2('Portföy', 'F9'), j9 = hd2('Portföy', 'J9');
  ok('maliyet 0 iken F sütunu 0 üretiyor', f9 === 0, `bulunan ${f9}`);
  ok('maliyet 0 iken K/Z % "" döndürüyor (#DIV/0! yok)', j9 === '', `bulunan ${JSON.stringify(j9)}`);
}
delete sayfalar['Portföy']['D9']; delete sayfalar['Portföy']['E9']; delete sayfalar['Portföy']['G9'];
sayfalar['Portföy']['D9'] = { deger: '' }; sayfalar['Portföy']['E9'] = { deger: '' }; sayfalar['Portföy']['G9'] = { deger: '' };

cikti.push('\nÖZET SAYFASI');
const { hucreDeger: hd3 } = motorKur(sayfalar, 'Portföy');
const O2 = (ref) => hd3('Özet', ref);
ok('Özet · toplam maliyet Portföy ile aynı', yakin(O2('B5'), TOP_MALIYET, 1e-6), `bulunan ${O2('B5')}`);
ok('Özet · toplam değer Portföy ile aynı',   yakin(O2('B6'), TOP_DEGER, 1e-6), `bulunan ${O2('B6')}`);
ok('Özet · toplam K/Z Portföy ile aynı',     yakin(O2('B7'), TOP_KZ, 1e-6), `bulunan ${O2('B7')}`);
ok('Özet · toplam K/Z % Portföy ile aynı',   yakin(O2('B8'), TOP_KZ / TOP_MALIYET, 1e-9), `bulunan ${O2('B8')}`);

cikti.push('\nTÜRE GÖRE DAĞILIM (SUMIF)');
const turBekleneni = {
  'BIST Hisse': 27180, 'TEFAS Fonu': 3524.208, 'ABD Hisse/ETF': 62560,
  'Kripto': 180500, 'Altın/Emtia': 108875, 'Döviz': 0, 'Nakit': 0,
};
let turToplam = 0;
Object.entries(turBekleneni).forEach(([tur, bekl], i) => {
  const satir = 11 + i;
  const v = O2(`B${satir}`);
  turToplam += v;
  ok(`${tur} = ${bekl}`, yakin(v, bekl, 1e-6), `bulunan ${v}`);
});
ok('tür dağılımı toplamı = portföy toplam değeri', yakin(turToplam, TOP_DEGER, 1e-6), `bulunan ${turToplam}`);
ok('Özet TOPLAM satırı da aynı', yakin(O2('B18'), TOP_DEGER, 1e-6), `bulunan ${O2('B18')}`);
ok('tür ağırlıkları toplamı %100', yakin(O2('C18'), 1, 1e-9), `bulunan ${O2('C18')}`);

cikti.push('\nDÜRÜSTLÜK / İÇERİK');
const k3 = zip['xl/worksheets/sheet3.xml'].toString('utf8');
ok('Kullanım sayfası "internetten fiyat çekmez" diyor', /internetten fiyat çekmez/.test(cozXml(k3)));
ok('Kullanım sayfası örnek satırların silinebileceğini söylüyor', /silebilirsiniz/.test(cozXml(k3)));
ok('şablonda makro/VBA parçası YOK', !Object.keys(zip).some(k => /vbaProject|\.bin$/i.test(k)));
ok('şablonda dış veri bağlantısı YOK',
   !Object.keys(zip).some(k => /connections|queryTable|externalLink/i.test(k)));

console.log(cikti.join('\n'));
console.log(`\n${'─'.repeat(58)}\nGEÇEN: ${gecen}   KALAN: ${kalan}\n`);
process.exitCode = kalan > 0 ? 1 : 0;
