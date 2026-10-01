export default {
  async fetch(request) {
    const SHEET_ID = 'SHEET ID';

    const corsHeaders = {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type',
      'Content-Type': 'application/json'
    };

    if (request.method === 'OPTIONS') {
      return new Response(null, { headers: corsHeaders });
    }

    const url = new URL(request.url);
    const action = url.searchParams.get('action');

    // ── TIMESTAMP ──
    if (action === 'timestamp') {
      try {
        const sheetUrl = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&sheet=Productos&t=${Date.now()}`;
        const res = await fetch(sheetUrl, { cache: 'no-store' });
        const text = await res.text();
        const json = JSON.parse(text.substring(47, text.length - 2));
        const rows = json.table.rows;
        let hashStr = '';
        for (const row of rows) {
          const c = row.c;
          for (let i = 0; i < c.length; i++) hashStr += (c[i]?.v ?? '') + '|';
        }
        let hash = 0;
        for (let i = 0; i < hashStr.length; i++) {
          hash = ((hash << 5) - hash) + hashStr.charCodeAt(i);
          hash |= 0;
        }
        return new Response(JSON.stringify({ ts: Math.abs(hash) }), { headers: corsHeaders });
      } catch (e) {
        return new Response(JSON.stringify({ ts: 0 }), { headers: corsHeaders });
      }
    }

    // ── DELIVERY RATES ──
    if (action === 'delivery') {
      try {
        const sheetUrl = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&sheet=Delivery&t=${Date.now()}`;
        const res = await fetch(sheetUrl, { cache: 'no-store' });
        const text = await res.text();
        const json = JSON.parse(text.substring(47, text.length - 2));
        const rows = json.table.rows;
        const rates = rows.map(row => ({
          zone:  row.c[0]?.v ?? '',
          price: row.c[1]?.v ?? 0,
          notes: row.c[2]?.v ?? ''
        }));
        return new Response(JSON.stringify(rates), { headers: corsHeaders });
      } catch (e) {
        return new Response(JSON.stringify([]), { headers: corsHeaders });
      }
    }

    // ── GET PRODUCTS ──
    if (request.method === 'GET') {
      const sheetUrl = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&sheet=Productos&t=${Date.now()}`;
      const res = await fetch(sheetUrl, { cache: 'no-store' });
      const text = await res.text();
      const json = JSON.parse(text.substring(47, text.length - 2));
      const rows = json.table.rows;
      const productos = [];
      for (const row of rows) {
        const c = row.c;
        if (c[8] && c[8].v === 'SI') {
          productos.push({
            id:          c[0]?.v ?? '',
            nombre:      c[1]?.v ?? '',
            categoria:   c[2]?.v ?? '',
            precio:      c[3]?.v ?? 0,
            unidad:      c[4]?.v ?? '',
            stock:       c[5]?.v ?? 0,
            imagen:      c[6]?.v ?? '',
            descripcion: c[7]?.v ?? ''
          });
        }
      }
      return new Response(JSON.stringify(productos), { headers: corsHeaders });
    }

    // ── POST ORDER ──
    if (request.method === 'POST') {
      const GAS_URL = 'GOOGLE SHEET STORE URL IT USUALLY ENDS WITH /EXEC';
      const body = await request.text();
      const res = await fetch(GAS_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain' },
        body: body
      });
      const data = await res.text();
      return new Response(data, { headers: corsHeaders });
    }

    return new Response(JSON.stringify({ error: 'Method not allowed' }), { status: 405, headers: corsHeaders });
  }
};
