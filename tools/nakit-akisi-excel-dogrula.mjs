// ═══════════════════════════════════════════════════════════════════════════
// FÖY — NAKİT AKIŞI EXCEL ŞABLONU DOĞRULAMA
//
// Çalıştır:  node tools/nakit-akisi-excel-dogrula.mjs
//
// Portföy şablonunun doğrulayıcısıyla (tools/excel-dogrula.mjs) aynı fikir:
// üretilen .xlsx AÇILIR (ZIP → XML), her formül OKUNUR ve küçük bir hesap
// motoruyla ÇALIŞTIRILIR; çıkan sayılar elle hesaplanmış beklentilerle
// karşılaştırılır. Doğrulanan şey "kod ne yazdı" değil, "dosyadaki formüller
// ne üretiyor".
//
// Motor, portföy doğrulayıcısınınkinden İKİ noktada daha Excel'e sadık:
//   1. GERÇEK BOŞ hücre (null) ile formülün döndürdüğü "" ayrı tutulur.
//      Excel'de boş hücre aritmetikte 0'dır, ama ""+1 → #DEĞER!. Bir formül
//      başka bir formülün "" sonucunu toplamaya kalkarsa test bunu yakalar.
//   2. Metin–sayı karşılaştırması Excel kuralıyla: metin HER sayıdan büyüktür
//      (""  > 0 → DOĞRU). Pay/karşılama formüllerindeki `E=""` korumasının
//      gerçekten gerekli olduğunu ve yerinde durduğunu ancak böyle sınayabiliriz.
// ⚠️ Sınır: bu Excel'in kendisi değildir; Excel/Sheets'in dosyayı biçimsel
// olarak kusursuz açtığını gerçek açılış gösterir.
// ═══════════════════════════════════════════════════════════════════════════

import { inflateRawSync } from 'node:zlib';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const KOK = join(dirname(fileURLToPath(import.meta.url)), '..');
const DOSYA = join(KOK, 'foy-nakit-akisi-takip-excel.xlsx');

let gecen = 0, kalan = 0; const cikti = [];
const ok = (ad, kosul, detay = '') => {
  if (kosul) { gecen++; cikti.push(`  ✓ ${ad}`); }
  else { kalan++; cikti.push(`  ✗ ${ad}   ${detay}`); }
};
const yakin = (a, b, tol = 1e-6) => typeof a === 'number' && Math.abs(a - b) < tol;

// ── ZIP OKUYUCU (excel-dogrula.mjs'den aynen) ─────────────────────────────
function zipOku(buf) {
  const dosyalar = {};
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

// Portföy doğrulayıcısından tek fark: stil numarası (s) da okunur, çünkü
// "girdi krem / hesap beyaz" ayrımı ve eksi-işaretli biçim stile bağlı.
function sayfaOku(xml) {
  const hucreler = {};
  // ⚠️ Kendinden kapanan hücreler (`<c r="B10" s="5"/>`) ayrı ele alınır —
  // tek desen bir sonraki hücrenin gövdesini yutuyordu (portföy doğrulayıcısı dersi).
  for (const m of xml.matchAll(/<c r="([A-Z]+\d+)"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
    const ref = m[1], oz = m[2], govde = m[3];
    const sm = /\bs="(\d+)"/.exec(oz);
    const stil = sm ? Number(sm[1]) : 0;
    if (govde === undefined) { hucreler[ref] = { deger: null, stil }; continue; }
    const f = /<f>([\s\S]*?)<\/f>/.exec(govde);
    const v = /<v>([\s\S]*?)<\/v>/.exec(govde);
    const t = /<is><t[^>]*>([\s\S]*?)<\/t><\/is>/.exec(govde);
    if (f) hucreler[ref] = { formul: cozXml(f[1]), stil };
    else if (t) hucreler[ref] = { deger: cozXml(t[1]), stil };
    else if (v) hucreler[ref] = { deger: Number(v[1]), stil };
    else hucreler[ref] = { deger: null, stil };
  }
  return hucreler;
}

// ── FORMÜL MOTORU ─────────────────────────────────────────────────────────
// Bu şablonun kullandığı alt küme:
//   IF · AND · OR · SUM · COUNT · COUNTIF · AVERAGEIF · + - * / · = <> < > <= >=
// Değer türleri: sayı · metin ("" dahil) · mantıksal · null (GERÇEK boş hücre)
const DIV0 = new Error('#DIV/0!');

function motorKur(sayfalar) {
  const onbellek = new Map();
  const zincir = new Set();
  const sutunNo = (s) => { let n = 0; for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64); return n; };
  const sutunAd = (n) => { let ad = ''; while (n > 0) { const r = (n - 1) % 26; ad = String.fromCharCode(65 + r) + ad; n = Math.floor((n - 1) / 26); } return ad; };

  function hucreDeger(sayfa, ref) {
    const anahtar = `${sayfa}!${ref}`;
    if (onbellek.has(anahtar)) return onbellek.get(anahtar);
    if (zincir.has(anahtar)) throw new Error('döngüsel başvuru: ' + anahtar);
    const h = (sayfalar[sayfa] || {})[ref];
    if (!h) { onbellek.set(anahtar, null); return null; }
    if ('deger' in h) { onbellek.set(anahtar, h.deger); return h.deger; }
    zincir.add(anahtar);
    let sonuc;
    try { sonuc = hesapla(h.formul, sayfa); } finally { zincir.delete(anahtar); }
    onbellek.set(anahtar, sonuc);
    return sonuc;
  }

  function aralikDegerleri(sayfa, bas, son) {
    const [, bs, br] = /^([A-Z]+)(\d+)$/.exec(bas);
    const [, ss, sr] = /^([A-Z]+)(\d+)$/.exec(son);
    const s1 = sutunNo(bs), s2 = sutunNo(ss), r1 = Number(br), r2 = Number(sr);
    const out = [];
    for (let s = Math.min(s1, s2); s <= Math.max(s1, s2); s++)
      for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++) out.push(hucreDeger(sayfa, sutunAd(s) + r));
    return out;
  }

  // Aritmetik: boş hücre 0, mantıksal 1/0, METİN (""" dahil) #DEĞER!
  const sayiyaCevir = (x) => {
    if (x === null) return 0;
    if (typeof x === 'boolean') return x ? 1 : 0;
    if (typeof x === 'number') return x;
    throw new Error(`#VALUE! — metin aritmetikte: ${JSON.stringify(x)}`);
  };

  // Excel karşılaştırması: boş hücre karşı tarafın türüne bürünür;
  // türler farklıysa sıra sayı < metin < mantıksal.
  function karsilastir(a, b) {
    if (a === null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
    if (b === null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
    const sira = (x) => typeof x === 'number' ? 0 : typeof x === 'string' ? 1 : 2;
    if (sira(a) !== sira(b)) return sira(a) - sira(b);
    if (typeof a === 'string') { const x = a.toLowerCase(), y = b.toLowerCase(); return x < y ? -1 : x > y ? 1 : 0; }
    return a < b ? -1 : a > b ? 1 : 0;
  }

  // COUNTIF / AVERAGEIF ölçütü: ">0", "<=5", "<>x", "abc", 7 ...
  // Karşılaştırma işleçli SAYISAL ölçütte yalnız sayılar eşleşir (Excel davranışı).
  function olcutEslesir(deger, olcut) {
    if (typeof olcut === 'number') return typeof deger === 'number' && deger === olcut;
    const m = /^(<=|>=|<>|<|>|=)?(.*)$/.exec(String(olcut));
    const op = m[1] || '=', ham = m[2];
    const sayi = ham !== '' && !isNaN(Number(ham)) ? Number(ham) : null;
    if (sayi !== null) {
      if (typeof deger !== 'number') return op === '<>';
      return { '=': deger === sayi, '<>': deger !== sayi, '<': deger < sayi, '>': deger > sayi, '<=': deger <= sayi, '>=': deger >= sayi }[op];
    }
    const d = deger === null ? '' : String(deger).toLowerCase(), h = ham.toLowerCase();
    return op === '<>' ? d !== h : op === '=' ? d === h : false;
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
        const k = karsilastir(sol, sag);
        sol = { '=': k === 0, '<>': k !== 0, '<': k < 0, '>': k > 0, '<=': k <= 0, '>=': k >= 0 }[op];
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
        const a = sayiyaCevir(sol), b = sayiyaCevir(sag);
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
        const a = sayiyaCevir(sol), b = sayiyaCevir(sag);
        if (op === '/' && b === 0) throw DIV0;
        sol = op === '*' ? a * b : a / b;
        bosla();
      }
      return sol;
    }

    function tekli() {
      bosla();
      if (src[i] === '-') { i++; return -sayiyaCevir(tekli()); }
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
      let jeton = '';
      while (i < src.length && /[A-Za-zÀ-ÿıİğĞüÜşŞöÖçÇ0-9_.$!':]/.test(src[i])) jeton += src[i++];
      bosla();
      if (src[i] === '(') {
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

    const altHesapla = (metin) => hesapla(metin, sayfa);
    const aralikMi = (metin) =>
      /^(?:'[^']+'!|[A-Za-zÀ-ÿıİğĞüÜşŞöÖçÇ]+!)?\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+$/.test(metin.trim());

    function aralikCoz(metin) {
      let m = metin.trim();
      let sf = sayfa;
      const sm = /^(?:'([^']+)'|([A-Za-zÀ-ÿıİğĞüÜşŞöÖçÇ]+))!(.+)$/.exec(m);
      if (sm) { sf = sm[1] || sm[2]; m = sm[3]; }
      const [b, s] = m.split(':').map(x => x.replace(/\$/g, ''));
      return aralikDegerleri(sf, b, s);
    }
    const degerler = (a) => aralikMi(a) ? aralikCoz(a) : [altHesapla(a)];

    function fonksiyon(ad, args) {
      if (ad === 'IF') {
        const k = altHesapla(args[0]);
        return k ? altHesapla(args[1]) : (args.length > 2 ? altHesapla(args[2]) : false);
      }
      // Excel'de OR/AND tüm argümanları değerlendirir (kısa devre yok) — aynısı.
      if (ad === 'OR')  { const v = args.map(a => !!altHesapla(a)); return v.some(Boolean); }
      if (ad === 'AND') { const v = args.map(a => !!altHesapla(a)); return v.every(Boolean); }
      if (ad === 'SUM') {
        let t = 0;
        for (const a of args) for (const x of degerler(a)) if (typeof x === 'number' && isFinite(x)) t += x;
        return t;
      }
      if (ad === 'COUNT') {
        let n = 0;
        for (const a of args) for (const x of degerler(a)) if (typeof x === 'number' && isFinite(x)) n++;
        return n;
      }
      if (ad === 'COUNTIF') {
        const alan = aralikCoz(args[0]), olcut = altHesapla(args[1]);
        return alan.filter(x => olcutEslesir(x, olcut)).length;
      }
      if (ad === 'AVERAGEIF') {
        const alan = aralikCoz(args[0]), olcut = altHesapla(args[1]);
        const ort = args[2] ? aralikCoz(args[2]) : alan;
        let t = 0, n = 0;
        for (let k = 0; k < alan.length; k++)
          if (olcutEslesir(alan[k], olcut) && typeof ort[k] === 'number') { t += ort[k]; n++; }
        if (n === 0) throw DIV0;
        return t / n;
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

    return ifade();
  }

  // Hata fırlatan hücre testte "HATA:<mesaj>" olarak görünür, koşu durmaz.
  const guvenli = (sayfa, ref) => { try { return hucreDeger(sayfa, ref); } catch (e) { return 'HATA:' + e.message; } };
  return { hucreDeger: guvenli };
}

// ── ÇALIŞTIR ──────────────────────────────────────────────────────────────
const zip = zipOku(readFileSync(DOSYA));

cikti.push('PAKET YAPISI');
const beklenenParcalar = ['[Content_Types].xml', '_rels/.rels', 'docProps/core.xml', 'docProps/app.xml',
  'xl/workbook.xml', 'xl/_rels/workbook.xml.rels', 'xl/styles.xml',
  'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml', 'xl/worksheets/sheet3.xml'];
for (const p of beklenenParcalar) ok(`parça var: ${p}`, p in zip);

const wb = zip['xl/workbook.xml'].toString('utf8');
const sayfaSirasi = [...wb.matchAll(/<sheet name="([^"]+)"/g)].map(m => m[1]);
ok('üç sayfa bu sırayla: Aylık / Özet / Kullanım', sayfaSirasi.join('/') === 'Aylık/Özet/Kullanım', `bulunan ${sayfaSirasi.join('/')}`);
ok('açılışta tam yeniden hesaplama açık (fullCalcOnLoad)', /fullCalcOnLoad="1"/.test(wb));
const s1xml = zip['xl/worksheets/sheet1.xml'].toString('utf8');
ok('Aylık sayfası açılışta seçili sekme (tabSelected)', /tabSelected="1"/.test(s1xml));
ok('Özet ve Kullanım seçili DEĞİL',
   !/tabSelected/.test(zip['xl/worksheets/sheet2.xml'].toString('utf8')) &&
   !/tabSelected/.test(zip['xl/worksheets/sheet3.xml'].toString('utf8')));
ok('başlık satırı donuk (ySplit=3)', /<pane ySplit="3"[^>]*state="frozen"/.test(s1xml));

const sayfalar = {
  'Aylık':    sayfaOku(s1xml),
  'Özet':     sayfaOku(zip['xl/worksheets/sheet2.xml'].toString('utf8')),
  'Kullanım': sayfaOku(zip['xl/worksheets/sheet3.xml'].toString('utf8')),
};
const A = sayfalar['Aylık'];

// ── STİLLER: girdi/hesap ayrımı ve eksi işareti ───────────────────────────
cikti.push('\nSTİL VE SAYI BİÇİMİ');
const stilXml = zip['xl/styles.xml'].toString('utf8');
const numFmt = Object.fromEntries([...stilXml.matchAll(/<numFmt numFmtId="(\d+)" formatCode="([^"]+)"/g)]
  .map(m => [Number(m[1]), cozXml(m[2])]));
const xfler = [...(/<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/.exec(stilXml)[1]).matchAll(/<xf ([^>]*?)\/?>/g)]
  .map(m => ({ numFmt: Number(/numFmtId="(\d+)"/.exec(m[1])[1]), fill: Number(/fillId="(\d+)"/.exec(m[1])[1]) }));
ok('cellXfs count özniteliği gerçek sayıyla aynı',
   Number(/<cellXfs count="(\d+)"/.exec(stilXml)[1]) === xfler.length, `xf sayısı ${xfler.length}`);
const bicim = (ref) => numFmt[xfler[A[ref].stil].numFmt] || '';
const zemin = (ref) => xfler[A[ref].stil].fill;
const KREM = 3, BEYAZ = 5;

// Eksi bölüm: biçimin ';' sonrası kısmı \- ile başlamalı (renk etiketi hariç).
const eksiIsaretli = (kod) => { const p = kod.split(';'); return p.length >= 2 && /^(\[[^\]]+\])*\\-/.test(p[1]); };
ok('₺ biçimi eksi sayıyı EKSİ İŞARETİYLE gösteriyor (yalnız renk değil)', eksiIsaretli(bicim('D5')), `biçim ${bicim('D5')}`);
ok('₺ biçimi eksi sayıyı ayrıca kırmızı gösteriyor', /;\[Red\]/.test(bicim('D5')));
ok('₺ biçiminde ₺ işareti var', /₺/.test(bicim('B4')));
ok('yüzde biçimi de eksi işaretli', eksiIsaretli(bicim('F4')) && /%/.test(bicim('F4')), `biçim ${bicim('F4')}`);
ok('Toplam (E) sütunu da eksi işaretli ₺ biçiminde', eksiIsaretli(bicim('E5')) && /₺/.test(bicim('E5')));
ok('TOPLAM satırı ₺ hücreleri eksi işaretli', ['B28', 'C28', 'D28', 'E28', 'G28'].every(r => eksiIsaretli(bicim(r)) && /₺/.test(bicim(r))));

let stilHatasi = [];
for (let r = 4; r <= 27; r++) {
  for (const s of ['A', 'B', 'C', 'D', 'G']) if (zemin(`${s}${r}`) !== KREM || 'formul' in A[`${s}${r}`]) stilHatasi.push(`${s}${r}`);
  for (const s of ['E', 'F', 'H'])           if (zemin(`${s}${r}`) !== BEYAZ || !('formul' in A[`${s}${r}`])) stilHatasi.push(`${s}${r}`);
  for (const s of ['B', 'C', 'D', 'E', 'G']) if (!/₺/.test(bicim(`${s}${r}`))) stilHatasi.push(`${s}${r}(₺)`);
  for (const s of ['F', 'H'])                if (!/%/.test(bicim(`${s}${r}`))) stilHatasi.push(`${s}${r}(%)`);
}
ok('24 satırın tamamında: A–D, G krem girdi · E, F, H beyaz formül · ₺/% biçimleri yerinde',
   stilHatasi.length === 0, `sorunlu: ${stilHatasi.slice(0, 8).join(', ')}`);

// ── YAPI VE FORMÜL BAĞLANTILARI ───────────────────────────────────────────
cikti.push('\nYAPI VE FORMÜL BAĞLANTILARI');
const basliklar = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'].map(s => A[`${s}3`]?.deger);
ok('başlıklar: Ay · Aktif · Pasif · Portföy · Toplam · Pasif payı · Gider · Karşılama',
   /^Ay$/.test(basliklar[0]) && /Aktif gelir/.test(basliklar[1]) && /Pasif gelir/.test(basliklar[2]) &&
   /Portföy geliri/.test(basliklar[3]) && /Toplam/.test(basliklar[4]) && /Pasif payı/.test(basliklar[5]) &&
   /gider/i.test(basliklar[6]) && /karşılama/i.test(basliklar[7]), `bulunan ${JSON.stringify(basliklar)}`);

let baglantiHatasi = [];
for (let r = 4; r <= 27; r++) {
  const bekl = {
    E: `IF(AND(B${r}="",C${r}="",D${r}=""),"",B${r}+C${r}+D${r})`,
    F: `IF(OR(E${r}="",E${r}<=0),"",C${r}/E${r})`,
    H: `IF(OR(G${r}="",G${r}<=0),"",C${r}/G${r})`,
  };
  for (const [s, f] of Object.entries(bekl)) if (A[`${s}${r}`]?.formul !== f) baglantiHatasi.push(`${s}${r}: ${A[`${s}${r}`]?.formul}`);
}
ok('24 satırın her birinde E/F/H formülü KENDİ satırına bağlı (kayma yok)', baglantiHatasi.length === 0, baglantiHatasi.slice(0, 3).join(' | '));
ok('satır 28 TOPLAM etiketi', A['A28']?.deger === 'TOPLAM');
ok('satır 29 ve sonrası boş (çalışma alanı tam 24 ay)', !Object.keys(A).some(k => Number(/\d+/.exec(k)[0]) > 28));
const topBekl = {
  B28: 'SUM(B4:B27)', C28: 'SUM(C4:C27)', D28: 'SUM(D4:D27)', E28: 'SUM(E4:E27)', G28: 'SUM(G4:G27)',
  F28: 'IF(E28>0,C28/E28,"")',
  H28: 'IF(COUNTIF(G4:G27,">0")=0,"",AVERAGEIF(G4:G27,">0",H4:H27))',
};
for (const [ref, f] of Object.entries(topBekl)) ok(`${ref} = ${f}`, A[ref]?.formul === f, `bulunan ${A[ref]?.formul}`);

// Yalnız iki programda da çalışan temel işlevler
const tumFormuller = Object.values(sayfalar).flatMap(s => Object.values(s)).filter(h => h.formul).map(h => h.formul);
const izinli = new Set(['IF', 'AND', 'OR', 'SUM', 'COUNT', 'COUNTIF', 'AVERAGEIF']);
const kullanilan = new Set(tumFormuller.flatMap(f => [...f.replace(/"[^"]*"/g, '').matchAll(/([A-Z]+)\(/g)].map(m => m[1])));
ok(`yalnız temel işlevler (${[...kullanilan].sort().join(', ')})`, [...kullanilan].every(f => izinli.has(f)),
   `izinsiz: ${[...kullanilan].filter(f => !izinli.has(f)).join(', ')}`);
ok('formüllerde noktalı virgül yok (XML\'de İngilizce yazım, virgül ayraç)', tumFormuller.every(f => !/;/.test(f.replace(/"[^"]*"/g, ''))));

// ── ELLE HESAPLANMIŞ BEKLENTİLER ──────────────────────────────────────────
// Kaynak: tools/nakit-akisi-excel-uret.mjs içindeki ORNEK dizisi. Aşağıdaki
// sayılar script'ten TÜRETİLMEDİ, elle toplanıp bölündü — testin bağımsızlığı için.
//   Ocak : 85000 + 12500 + 18400 = 115900   gider 62000
//   Şubat: 85000 + 13250 −  9600 =  88650   gider 58500
//   Mart : 85000 + 14100 + 22750 = 121850   gider yok
const { hucreDeger } = motorKur(sayfalar);
const AY = (ref) => hucreDeger('Aylık', ref);
const OZ = (ref) => hucreDeger('Özet', ref);

const ORN = [
  { r: 4, ay: 'Ocak 2026',  aktif: 85000, pasif: 12500, portfoy: 18400, gider: 62000, toplam: 115900 },
  { r: 5, ay: 'Şubat 2026', aktif: 85000, pasif: 13250, portfoy: -9600, gider: 58500, toplam: 88650 },
  { r: 6, ay: 'Mart 2026',  aktif: 85000, pasif: 14100, portfoy: 22750, gider: null,  toplam: 121850 },
];
cikti.push('\nÖRNEK AYLAR (girdi hücreleri)');
for (const o of ORN) {
  ok(`satır ${o.r} · ${o.ay} · aktif ${o.aktif} · pasif ${o.pasif} · portföy ${o.portfoy} · gider ${o.gider ?? 'boş'}`,
     AY(`A${o.r}`) === o.ay && AY(`B${o.r}`) === o.aktif && AY(`C${o.r}`) === o.pasif &&
     AY(`D${o.r}`) === o.portfoy && AY(`G${o.r}`) === o.gider);
}
ok('örnek dışı ilk satır (7) tamamen boş', ['A7', 'B7', 'C7', 'D7', 'G7'].every(r => AY(r) === null));

cikti.push('\nSATIR HESAPLARI');
for (const o of ORN) {
  ok(`${o.ay} · toplam = ${o.toplam}`, yakin(AY(`E${o.r}`), o.toplam), `bulunan ${AY(`E${o.r}`)}`);
  ok(`${o.ay} · pasif payı = ${o.pasif}/${o.toplam} = %${(o.pasif / o.toplam * 100).toFixed(2)}`,
     yakin(AY(`F${o.r}`), o.pasif / o.toplam, 1e-12), `bulunan ${AY(`F${o.r}`)}`);
  if (o.gider) ok(`${o.ay} · karşılama = ${o.pasif}/${o.gider} = %${(o.pasif / o.gider * 100).toFixed(2)}`,
     yakin(AY(`H${o.r}`), o.pasif / o.gider, 1e-12), `bulunan ${AY(`H${o.r}`)}`);
  else ok(`${o.ay} · gider boş → karşılama "" (0 ya da hata değil)`, AY(`H${o.r}`) === '', `bulunan ${JSON.stringify(AY(`H${o.r}`))}`);
}

cikti.push('\nEKSİ PORTFÖY GELİRİ');
ok('Şubat portföy geliri eksi (-9600)', AY('D5') === -9600);
ok('Şubat toplamı eksiyi düşüyor: 98250 − 9600 = 88650', yakin(AY('E5'), 88650));

cikti.push('\nTOPLAM SATIRI (28)');
ok('aktif toplam = 255000',  yakin(AY('B28'), 255000), `bulunan ${AY('B28')}`);
ok('pasif toplam = 39850',   yakin(AY('C28'), 39850),  `bulunan ${AY('C28')}`);
ok('portföy toplam = 31550', yakin(AY('D28'), 31550),  `bulunan ${AY('D28')}`);
ok('toplam gelir = 326400',  yakin(AY('E28'), 326400), `bulunan ${AY('E28')}`);
ok('toplam gelir, üç kaynağın toplamıyla tutarlı', yakin(AY('E28'), AY('B28') + AY('C28') + AY('D28')));
ok('gider toplamı = 120500 (Mart boş)', yakin(AY('G28'), 120500), `bulunan ${AY('G28')}`);
ok('toplam pasif payı = 39850/326400', yakin(AY('F28'), 39850 / 326400, 1e-12), `bulunan ${AY('F28')}`);
const ORT_KARSILAMA = (12500 / 62000 + 13250 / 58500) / 2;
ok(`TOPLAM satırı karşılama = gideri girilmiş 2 ayın ortalaması = %${(ORT_KARSILAMA * 100).toFixed(2)}`,
   yakin(AY('H28'), ORT_KARSILAMA, 1e-12), `bulunan ${AY('H28')}`);

cikti.push('\nBOŞ SATIR DAVRANIŞI');
for (const ref of ['E10', 'F10', 'H10', 'E27', 'F27', 'H27']) {
  ok(`${ref} boş satırda "" döndürüyor (0 veya hata değil)`, AY(ref) === '', `bulunan ${JSON.stringify(AY(ref))}`);
}

cikti.push('\nKENAR DURUMLAR (geçici satırlarla)');
// Bu durumlar dosyada yok; sayfaya geçici değer yazıp motoru yeniden kuruyoruz.
const dene = (yamalar, okunacak) => {
  const kopya = { ...sayfalar, 'Aylık': { ...A } };
  for (const [ref, deger] of Object.entries(yamalar)) kopya['Aylık'][ref] = { ...A[ref], deger };
  const { hucreDeger: hd } = motorKur(kopya);
  return okunacak.map(r => hd(okunacak.sayfa || 'Aylık', r));
};
{
  const [e, f] = dene({ B9: 0, C9: 0, D9: 0 }, ['E9', 'F9']);
  ok('gelirler 0 girildiyse toplam 0, pasif payı "" (#DIV/0! yok)', e === 0 && f === '', `bulunan ${e}, ${JSON.stringify(f)}`);
}
{
  const [e, f] = dene({ B9: 1000, D9: -5000 }, ['E9', 'F9']);
  ok('toplam eksiyse (-4000) pasif payı "" (anlamsız yüzde gösterilmez)', e === -4000 && f === '', `bulunan ${e}, ${JSON.stringify(f)}`);
}
{
  const [e, f] = dene({ B9: 50000 }, ['E9', 'F9']);
  ok('yalnız aktif gelir girildiyse toplam 50000, pasif payı %0', e === 50000 && f === 0, `bulunan ${e}, ${JSON.stringify(f)}`);
}
{
  const [h] = dene({ C9: 5000, G9: 0 }, ['H9']);
  ok('gider 0 girildiyse karşılama "" (#DIV/0! yok)', h === '', `bulunan ${JSON.stringify(h)}`);
}
{
  const [h] = dene({ G9: 40000 }, ['H9']);
  ok('gider girilip pasif gelir boşsa karşılama %0', h === 0, `bulunan ${JSON.stringify(h)}`);
}
{
  // Kanıt: `E=""` koruması kaldırılsaydı Excel'de "" > 0 DOĞRU olur ve C/"" #DEĞER! verirdi.
  const kopya = { ...sayfalar, 'Aylık': { ...A, F10: { ...A.F10, formul: 'IF(E10<=0,"",C10/E10)' } } };
  const { hucreDeger: hd } = motorKur(kopya);
  ok('kontrol: E="" koruması OLMADAN boş satır #DEĞER! verirdi (motor Excel gibi davranıyor)',
     String(hd('Aylık', 'F10')).startsWith('HATA:#VALUE!'), `bulunan ${JSON.stringify(hd('Aylık', 'F10'))}`);
}
{
  // Hiç gider girilmemiş dosya: ortalama karşılama hata değil "".
  const k = { ...sayfalar, 'Aylık': { ...A, G4: { ...A.G4, deger: null }, G5: { ...A.G5, deger: null } } };
  const { hucreDeger: hd } = motorKur(k);
  ok('hiç gider yoksa TOPLAM karşılama "" ve Özet ortalama karşılama "" (#DIV/0! yok)',
     hd('Aylık', 'H28') === '' && hd('Özet', 'B9') === '', `bulunan ${JSON.stringify(hd('Aylık', 'H28'))}, ${JSON.stringify(hd('Özet', 'B9'))}`);
  ok('hiç gider yoksa gideri girilmiş ay sayısı 0', hd('Özet', 'B8') === 0);
}
{
  // Ay adı yazılmış ama gelir girilmemiş satır "dolu ay" sayılmaz.
  const k = { ...sayfalar, 'Aylık': { ...A, A7: { ...A.A7, deger: 'Nisan 2026' } } };
  const { hucreDeger: hd } = motorKur(k);
  ok('yalnız ay adı yazılmış satır dolu ay SAYILMAZ (ortalama bozulmaz)', hd('Özet', 'B5') === 3 && yakin(hd('Özet', 'B7'), 108800));
}

cikti.push('\nÖZET SAYFASI');
const O = sayfalar['Özet'];
ok('Özet formülleri Aylık sayfasına tırnaklı adla bağlı (\'Aylık\'!)',
   ['B5', 'B6', 'B8', 'B9', 'B12', 'B13', 'B14'].every(r => /'Aylık'!/.test(O[r]?.formul || '')));
ok('dolu ay sayısı = 3', OZ('B5') === 3, `bulunan ${OZ('B5')}`);
ok('toplam gelir = 326400', yakin(OZ('B6'), 326400), `bulunan ${OZ('B6')}`);
ok('aylık ortalama gelir = 326400/3 = 108800', yakin(OZ('B7'), 108800), `bulunan ${OZ('B7')}`);
ok('gideri girilmiş ay sayısı = 2', OZ('B8') === 2, `bulunan ${OZ('B8')}`);
ok(`ortalama karşılama = (12500/62000 + 13250/58500)/2 = %${(ORT_KARSILAMA * 100).toFixed(2)}`,
   yakin(OZ('B9'), ORT_KARSILAMA, 1e-12), `bulunan ${OZ('B9')}`);
ok('Özet ortalama karşılama = Aylık TOPLAM satırı karşılama (tek tanım)', yakin(OZ('B9'), AY('H28'), 1e-15));

const kaynak = [['B12', 'C12', 'Aktif', 255000], ['B13', 'C13', 'Pasif', 39850], ['B14', 'C14', 'Portföy', 31550]];
for (const [t, p, ad, bekl] of kaynak) {
  ok(`${ad} toplam = ${bekl}`, yakin(OZ(t), bekl), `bulunan ${OZ(t)}`);
  ok(`${ad} payı = ${bekl}/326400 = %${(bekl / 326400 * 100).toFixed(2)}`, yakin(OZ(p), bekl / 326400, 1e-12), `bulunan ${OZ(p)}`);
}
ok('Özet TOPLAM = 326400', yakin(OZ('B15'), 326400), `bulunan ${OZ('B15')}`);
ok('paylar toplamı tam %100', yakin(OZ('C15'), 1, 1e-12), `bulunan ${OZ('C15')}`);
ok('Özet pasif payı = Aylık TOPLAM satırı pasif payı', yakin(OZ('C13'), AY('F28'), 1e-15));
{
  // Eksi portföy geliri yılı: pay eksi görünür, diğerleri %100'ü aşar, toplam yine %100.
  const k = { ...sayfalar, 'Aylık': { ...A, D4: { ...A.D4, deger: -50000 } } };
  const { hucreDeger: hd } = motorKur(k);
  ok('portföy toplamı eksiyse payı eksi, paylar toplamı yine %100',
     hd('Özet', 'C14') < 0 && yakin(hd('Özet', 'C15'), 1, 1e-12), `bulunan ${hd('Özet', 'C14')}, ${hd('Özet', 'C15')}`);
}

cikti.push('\nDÜRÜSTLÜK / İÇERİK');
const k3 = cozXml(zip['xl/worksheets/sheet3.xml'].toString('utf8'));
const adimlar = [...k3.matchAll(/<t xml:space="preserve">([1-9])\.\s/g)].map(m => m[1]);
ok('Kullanım sayfasında 5 numaralı adım var', adimlar.join('') === '12345', `bulunan ${adimlar.join(',')}`);
ok('Kullanım: makro yok diyor', /Makro yoktur/.test(k3));
ok('Kullanım: dış veri bağlantısı yok diyor', /dış veri bağlantısı yoktur/.test(k3));
ok('Kullanım: Google E-Tablolar\'da açılır diyor', /Google E-Tablolar/.test(k3));
ok('Kullanım: işlem defteri tutmaz diyor', /İşlem defteri tutmaz/.test(k3));
ok('Kullanım: örnek satırların silinebileceğini söylüyor', /silebilirsiniz/.test(k3));
ok('Kullanım: eksi sayının işaretle gösterildiğini söylüyor', /eksi işaretiyle/.test(k3));
ok('şablonda makro/VBA parçası YOK', !Object.keys(zip).some(k => /vbaProject|\.bin$/i.test(k)));
ok('şablonda dış veri bağlantısı YOK',
   !Object.keys(zip).some(k => /connections|queryTable|externalLink/i.test(k)));
ok('hiçbir formül başka dosyaya başvurmuyor ([...] yok)', tumFormuller.every(f => !/\[/.test(f)));

console.log(cikti.join('\n'));
console.log(`\n${'─'.repeat(58)}\nGEÇEN: ${gecen}   KALAN: ${kalan}\n`);
process.exitCode = kalan > 0 ? 1 : 0;
