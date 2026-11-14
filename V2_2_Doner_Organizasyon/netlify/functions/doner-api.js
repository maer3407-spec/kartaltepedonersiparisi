// Netlify Function proxy: public browser -> Netlify -> Google Apps Script.
// Set APPS_SCRIPT_URL in Netlify environment variables to the deployed /exec URL.
const APPS_SCRIPT_URL = process.env.APPS_SCRIPT_URL || 'https://script.google.com/macros/s/AKfycbzdm48QZ6Zljaup3v32I1oQH9V4yDCdkbMjIqy3SgIUB1dvtOzf-12xMxBFdbf5E1zo/exec';

exports.handler = async function (event) {
  const headers = {
    'Content-Type': 'application/json; charset=utf-8',
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Allow-Methods': 'POST, OPTIONS'
  };
  if (event.httpMethod === 'OPTIONS') return { statusCode: 204, headers, body: '' };
  if (event.httpMethod !== 'POST') return { statusCode: 405, headers, body: JSON.stringify({ status: 'error', message: 'Yalnızca POST desteklenir.' }) };
  try {
    const input = JSON.parse(event.body || '{}');
    if (!['create', 'trackOrder'].includes(input.action)) throw new Error('Geçersiz işlem.');
    if (input.action === 'create') {
      const d = input.data || {};
      if (!['satis', 'ikram'].includes(d.kategori)) throw new Error('Kayıt türü geçersiz.');
    }
    if (input.action === 'trackOrder' && !String(input.id || '').trim()) throw new Error('Sipariş numarası gerekli.');
    const upstream = await fetch(APPS_SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(input),
      redirect: 'follow'
    });
    const raw = await upstream.text();
    let data;
    try { data = JSON.parse(raw); }
    catch (_) { throw new Error('Apps Script geçerli JSON yanıtı döndürmedi. Web uygulaması dağıtımını ve APPS_SCRIPT_URL ayarını kontrol edin.'); }
    if (data.status === 'error' || data.ok === false) return { statusCode: 400, headers, body: JSON.stringify(data) };
    return { statusCode: 200, headers, body: JSON.stringify(data) };
  } catch (err) {
    return { statusCode: 400, headers, body: JSON.stringify({ status: 'error', message: err.message || 'İstek tamamlanamadı.' }) };
  }
};
