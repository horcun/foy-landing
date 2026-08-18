// appfoy.com — fiyat iddiası tutarlılık kontrolü
//
//   node check-pricing.mjs
//
// Sayfadaki aylık/yıllık fiyatı okur, "X AY BEDAVA" rozetini ve "%X tasarruf"
// notunu bu fiyatlardan yeniden hesaplar, tutmuyorsa hata verir.
//
// Neden: bu iddialar elle yazılıyor ve fiyat değişince sessizce yanlış kalıyor.
// Gerçekten oldu — 49 / 349 fiyatındayken sayfa "2 AY BEDAVA" diyordu, doğrusu
// 4 aydı; sonra fiyat 29,99 / 149,99 olunca "%58 / 7 ay" olması gerekti.
//
// Hesap mantığı uygulamadaki constants/pricing.js ile birebir aynı tutulmalı.

import { readFileSync } from 'node:fs';

const html = readFileSync(new URL('index.html', import.meta.url), 'utf8');
const sayi = (s) => Number(String(s).replace(/\./g, '').replace(',', '.'));

const fiyatlar = [...html.matchAll(/class="pprice">([\d.,]+)<small>\s*₺\/(ay|yıl)</g)]
  .map((m) => ({ tutar: sayi(m[1]), birim: m[2] }));

const aylik = fiyatlar.find((f) => f.birim === 'ay')?.tutar;
const yillik = fiyatlar.find((f) => f.birim === 'yıl')?.tutar;

const hatalar = [];
const bilgi = [];

if (!aylik || !yillik) {
  hatalar.push(`Fiyatlar okunamadı (aylık: ${aylik}, yıllık: ${yillik})`);
} else {
  bilgi.push(`aylık ${aylik.toFixed(2)} ₺ · yıllık ${yillik.toFixed(2)} ₺`);

  const tam = aylik * 12;
  const beklenenYuzde = Math.round(((tam - yillik) / tam) * 100);
  const beklenenAy = Math.floor(12 - yillik / aylik + 0.05);

  const yuzdeM = html.match(/%(\d+)\s*tasarruf/);
  if (!yuzdeM) {
    hatalar.push('"%X tasarruf" ifadesi sayfada bulunamadı');
  } else if (Number(yuzdeM[1]) !== beklenenYuzde) {
    hatalar.push(`Tasarruf yüzdesi yanlış: sayfada %${yuzdeM[1]}, hesaplanan %${beklenenYuzde}`);
  } else {
    bilgi.push(`%${beklenenYuzde} tasarruf — doğru`);
  }

  const ayM = html.match(/class="tag">(\d+)\s*AY BEDAVA</);
  if (!ayM) {
    hatalar.push('"X AY BEDAVA" rozeti sayfada bulunamadı');
  } else if (Number(ayM[1]) !== beklenenAy) {
    hatalar.push(`Bedava ay yanlış: rozette ${ayM[1]}, hesaplanan ${beklenenAy}`);
  } else {
    bilgi.push(`${beklenenAy} ay bedava — doğru`);
  }
}

// Sahip olunmayan özellik iddiaları geri sızmasın
const yasakli = ['bildirim', 'Bildirim', 'notification', 'push notification', 'gerçek zamanlı'];
for (const k of yasakli) {
  if (html.includes(k)) hatalar.push(`Kaldırılmış iddia geri gelmiş: "${k}"`);
}

console.log('appfoy.com fiyat ve iddia kontrolü');
bilgi.forEach((b) => console.log('  ✓ ' + b));
if (!hatalar.length) {
  console.log('  ✓ kaldırılmış özellik iddiaları geri gelmemiş');
  console.log('\nTüm kontroller geçti.');
} else {
  hatalar.forEach((h) => console.log('  ✗ ' + h));
  console.log(`\n${hatalar.length} sorun bulundu.`);
  process.exit(1);
}
