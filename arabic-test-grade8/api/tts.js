// Returns spoken audio (MP3) for a short text in Arabic or Hebrew.
// Usage: /api/tts?lang=ar&q=مرحبا
const LANGS = { ar: 'ar', he: 'iw' };

export default async function handler(req, res) {
  const lang = LANGS[req.query.lang] || 'ar';
  const q = String(req.query.q || '').trim().slice(0, 200);
  if (!q) {
    res.status(400).send('missing q');
    return;
  }
  const url =
    'https://translate.google.com/translate_tts?ie=UTF-8&client=tw-ob' +
    '&tl=' + lang + '&q=' + encodeURIComponent(q);
  try {
    const r = await fetch(url, {
      headers: {
        'User-Agent': 'Mozilla/5.0',
        Referer: 'https://translate.google.com/',
      },
    });
    if (!r.ok) {
      res.status(502).send('tts upstream ' + r.status);
      return;
    }
    const buf = Buffer.from(await r.arrayBuffer());
    res.setHeader('Content-Type', 'audio/mpeg');
    // same word = same audio, so let the browser and Vercel's CDN keep it
    res.setHeader('Cache-Control', 'public, max-age=86400, s-maxage=31536000, immutable');
    res.status(200).send(buf);
  } catch (e) {
    res.status(502).send('tts error');
  }
}
