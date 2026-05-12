const GAS_URL = 'https://script.google.com/macros/s/AKfycbwYbeHqRdtS_pSpT2E5B886bsOU_Hxd9pb21ZZDzwA556OdFbdHG7edzeltMnuCagsg/exec';

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  try {
    const r = await fetch(GAS_URL + '?t=' + Date.now());
    const data = await r.json();
    res.status(200).json(data);
  } catch(e) {
    res.status(500).json({status:'error', message: e.toString()});
  }
}
