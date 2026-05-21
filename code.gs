// ========================================
// ✏️ VENDOR CONFIG — ONLY EDIT THIS SECTION
// ========================================
const SHEET_ID        = 'EDIT THE SHEET ID';
const STORE_NAME      = 'YOUR STORE NAME';
const TAGLINE         = 'YOUR STORE TAGLINE';
const CITY            = 'LOCATION OF YOUR STORE E.G PORT HARCOURT';
const ADDRESS         = 'SHOP ADDRESS OF YOUR STORE OR YOU CAN PUT LOCATION AND ADD CITY TO IT E.G PORT HARCOURT CITY';
const PHONE           = 'YOUR PHONE NUMBER';
const VENDOR_EMAIL    = 'STORE OWNER EMAIL ADDRESS';
const CSR_EMAIL       = 'STORE OWNER SALES REP EMAIL ADDRESS, OR LEAVE BLANK IF YOU DONT HAVE A SALES REP';
const LOGO_URL        = 'STORE LOGO URL';
const BANNER_TITLE    = 'Welcome to STORE NAME';
const BANNER_SUB      = 'The best perfume in Port Harcourt';
const MON_SAT         = '8:00 AM - 6:00 PM';
const SUN             = '12:00 PM - 6:00 PM';
const YEAR            = '2026';
const COLOR_PRIMARY   = '#667eea';
const COLOR_SECONDARY = '#764ba2';
const COLOR_ACCENT    = '#ff9800';
const DEVELOPER_NAME  = 'THE PERSON THAT CREATE THE STORE NAME';
const DEVELOPER_WA    = '2349065147362';
// ========================================
// ✅ DO NOT TOUCH ANYTHING BELOW THIS LINE
// ========================================

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'timestamp') {
    const val = PropertiesService.getScriptProperties().getProperty('last_update');
    return ContentService.createTextOutput(JSON.stringify({ ts: Number(val) || 0 }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle(STORE_NAME + ' - Online Store')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === 'pedido') {
      return ContentService.createTextOutput(
        JSON.stringify(procesarPedido(body.datosCliente, body.carrito, body.total))
      ).setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(
      JSON.stringify({ exito: false, mensaje: 'Unknown action' })
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(
      JSON.stringify({ exito: false, mensaje: error.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function onEdit(e) {
  PropertiesService.getScriptProperties().setProperty('last_update', String(Date.now()));
}

function obtenerUltimaActualizacion() {
  const val = PropertiesService.getScriptProperties().getProperty('last_update');
  return val ? Number(val) : 0;
}

// ========================================
// INITIALIZE SHEETS — Run this once!
// ========================================
function inicializarHojas() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // ── PRODUCTOS ──────────────────────────────
  let productos = ss.getSheetByName('Productos');
  if (!productos) productos = ss.insertSheet('Productos');
  productos.clearContents();
  productos.clearFormats();

  const pHeaders = [['ID', 'Nombre', 'Categoría', 'Precio', 'Unidad', 'Stock', 'Imagen', 'Descripción', 'Activo']];
  productos.getRange(1, 1, 1, 9).setValues(pHeaders);
  productos.getRange(1, 1, 1, 9)
    .setBackground('#667eea')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(11);

  // Demo watch products — edit or delete these and add your real products
  const demoProducts = [
    [1, 'Classic Gold Watch', 'Analog', 45000, 'piece', 20, 'https://placehold.co/300x300?text=Classic+Gold+Watch', 'Elegant classic gold wristwatch with leather strap', 'SI'],
    [2, 'Stainless Steel Chrono', 'Chronograph', 75000, 'piece', 15, 'https://placehold.co/300x300?text=Chrono+Watch', 'Premium stainless steel chronograph with date display', 'SI'],
    [3, 'Sport Digital Watch', 'Digital', 25000, 'piece', 30, 'https://placehold.co/300x300?text=Sport+Digital', 'Durable sport digital watch with stopwatch function', 'SI'],
  ];
  productos.getRange(2, 1, demoProducts.length, 9).setValues(demoProducts);

  productos.setColumnWidth(1, 50);
  productos.setColumnWidth(2, 250);
  productos.setColumnWidth(3, 120);
  productos.setColumnWidth(4, 80);
  productos.setColumnWidth(5, 80);
  productos.setColumnWidth(6, 60);
  productos.setColumnWidth(7, 350);
  productos.setColumnWidth(8, 250);
  productos.setColumnWidth(9, 70);
  productos.setFrozenRows(1);

  // ── PEDIDOS ────────────────────────────────
  let pedidos = ss.getSheetByName('Pedidos');
  if (!pedidos) pedidos = ss.insertSheet('Pedidos');
  pedidos.clearContents();
  pedidos.clearFormats();

  const pedHeaders = [['ID Pedido', 'Fecha', 'Nombre Cliente', 'Email', 'Teléfono', 'Dirección', 'Productos', 'Total', 'Estado']];
  pedidos.getRange(1, 1, 1, 9).setValues(pedHeaders);
  pedidos.getRange(1, 1, 1, 9)
    .setBackground('#2c3e50')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(11);

  pedidos.setColumnWidth(1, 160);
  pedidos.setColumnWidth(2, 140);
  pedidos.setColumnWidth(3, 180);
  pedidos.setColumnWidth(4, 200);
  pedidos.setColumnWidth(5, 130);
  pedidos.setColumnWidth(6, 250);
  pedidos.setColumnWidth(7, 300);
  pedidos.setColumnWidth(8, 100);
  pedidos.setColumnWidth(9, 100);
  pedidos.setFrozenRows(1);

  // ── CLIENTES ───────────────────────────────
  let clientes = ss.getSheetByName('Clientes');
  if (!clientes) clientes = ss.insertSheet('Clientes');
  clientes.clearContents();
  clientes.clearFormats();

  const cHeaders = [['#', 'Nombre', 'Email', 'Teléfono', 'Dirección', 'Fecha Registro', 'Total Compras']];
  clientes.getRange(1, 1, 1, 7).setValues(cHeaders);
  clientes.getRange(1, 1, 1, 7)
    .setBackground('#ff9800')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(11);

  clientes.setColumnWidth(1, 50);
  clientes.setColumnWidth(2, 180);
  clientes.setColumnWidth(3, 220);
  clientes.setColumnWidth(4, 130);
  clientes.setColumnWidth(5, 250);
  clientes.setColumnWidth(6, 140);
  clientes.setColumnWidth(7, 120);
  clientes.setFrozenRows(1);

  Logger.log('✅ Done! All 3 sheets created successfully.');
}

// ========================================
// GET PRODUCTS
// ========================================
function obtenerProductos() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Productos');
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    const productos = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][8] === 'SI') {
        productos.push({
          id: data[i][0], nombre: data[i][1], categoria: data[i][2] || '',
          precio: data[i][3], unidad: data[i][4], stock: data[i][5],
          imagen: data[i][6] || '', descripcion: data[i][7] || ''
        });
      }
    }
    return productos;
  } catch (error) {
    Logger.log('Error: ' + error);
    return [];
  }
}

// ========================================
// PROCESS ORDER
// ========================================
function procesarPedido(datosCliente, carrito, total) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheetPedidos   = ss.getSheetByName('Pedidos');
    const sheetClientes  = ss.getSheetByName('Clientes');
    const sheetProductos = ss.getSheetByName('Productos');
    const idPedido = 'ORD-' + new Date().getTime();
    const productosTexto = carrito.map(item => `${item.nombre} x${item.cantidad} (${item.unidad}) - ₦${fmt(item.subtotal)}`).join('\n');
    reducirStock(sheetProductos, carrito);
    PropertiesService.getScriptProperties().setProperty('last_update', String(Date.now()));
    sheetPedidos.appendRow([idPedido, new Date(), datosCliente.nombre, datosCliente.email, datosCliente.telefono, datosCliente.direccion, productosTexto, total, 'Pending']);
    registrarCliente(sheetClientes, datosCliente, total);
    enviarConfirmacionPedido(idPedido, datosCliente, carrito, total);
    return { exito: true, mensaje: 'Order processed successfully', idPedido };
  } catch (error) {
    Logger.log('Error: ' + error);
    return { exito: false, mensaje: error.toString() };
  }
}

function registrarCliente(sheet, datosCliente, totalCompra) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === datosCliente.email) {
      sheet.getRange(i + 1, 7).setValue((data[i][6] || 0) + totalCompra);
      return;
    }
  }
  sheet.appendRow([data.length, datosCliente.nombre, datosCliente.email, datosCliente.telefono, datosCliente.direccion, new Date(), totalCompra]);
}

function reducirStock(sheet, carrito) {
  const data = sheet.getDataRange().getValues();
  carrito.forEach(item => {
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === item.id) {
        sheet.getRange(i + 1, 6).setValue(data[i][5] - item.cantidad);
        break;
      }
    }
  });
}

function enviarConfirmacionPedido(idPedido, datosCliente, carrito, total) {
  try {
    const pdf = Utilities.newBlob(generarHTMLPDF(idPedido, datosCliente, carrito, total), 'text/html', 'order.html')
      .getAs('application/pdf').setName(`Order_${idPedido}.pdf`);
    MailApp.sendEmail({ to: datosCliente.email, subject: `Order Confirmation ${idPedido} - ${STORE_NAME}`, htmlBody: generarHTMLEmail(idPedido, datosCliente, carrito, total), attachments: [pdf], name: STORE_NAME });
    MailApp.sendEmail({ to: VENDOR_EMAIL, subject: `🔔 NEW ORDER ${idPedido} - ${STORE_NAME}`, htmlBody: generarHTMLEmailVendedor(idPedido, datosCliente, carrito, total), attachments: [pdf], name: STORE_NAME + ' - Orders' });
    // 👇 extra copy to customer support
MailApp.sendEmail({
  to: CSR_EMAIL,
  subject: `📦 CUSTOMER SUPPORT COPY ${idPedido} - ${STORE_NAME}`,
  htmlBody: generarHTMLEmailVendedor(idPedido, datosCliente, carrito, total),
  attachments: [pdf],
  name: STORE_NAME + ' - Support' });
    return true;
  } catch (error) {
    Logger.log('Error sending email: ' + error);
    return false;
  }
}

function generarHTMLPDF(idPedido, datosCliente, carrito, total) {
  const fecha = Utilities.formatDate(new Date(), 'GMT+1', 'dd/MM/yyyy HH:mm');
  const rows = carrito.map(item => `<tr><td style="padding:10px;border-bottom:1px solid #e0e0e0;">${item.nombre}</td><td style="padding:10px;border-bottom:1px solid #e0e0e0;text-align:center;">${item.cantidad}</td><td style="padding:10px;border-bottom:1px solid #e0e0e0;text-align:center;">${item.unidad}</td><td style="padding:10px;border-bottom:1px solid #e0e0e0;text-align:right;">₦${fmt(item.precio)}</td><td style="padding:10px;border-bottom:1px solid #e0e0e0;text-align:right;font-weight:bold;">₦${fmt(item.subtotal)}</td></tr>`).join('');
  return `<!DOCTYPE html><html><head><meta charset="UTF-8"><style>body{font-family:Arial,sans-serif;margin:40px;color:#333;}.header{text-align:center;margin-bottom:30px;border-bottom:3px solid ${COLOR_PRIMARY};padding-bottom:20px;}.empresa{color:${COLOR_PRIMARY};font-size:24px;font-weight:bold;}table{width:100%;border-collapse:collapse;margin:20px 0;}th{background:${COLOR_PRIMARY};color:white;padding:12px;text-align:left;}.total{text-align:right;font-size:24px;font-weight:bold;color:${COLOR_PRIMARY};}</style></head><body><div class="header"><img src="${LOGO_URL}" style="max-width:150px;" alt="Logo"><div class="empresa">${STORE_NAME}</div><div style="color:#666;font-size:12px;">${ADDRESS}, ${CITY}<br>Phone: ${PHONE}</div></div><p><strong>ORDER:</strong> ${idPedido} &nbsp; <strong>DATE:</strong> ${fecha} &nbsp; <strong>STATUS:</strong> Pending</p><p><strong>Name:</strong> ${datosCliente.nombre} &nbsp; <strong>Phone:</strong> ${datosCliente.telefono}</p><p><strong>Address:</strong> ${datosCliente.direccion}</p><table><thead><tr><th>Product</th><th style="text-align:center;">Qty</th><th style="text-align:center;">Unit</th><th style="text-align:right;">Price</th><th style="text-align:right;">Subtotal</th></tr></thead><tbody>${rows}</tbody></table><div class="total">TOTAL: ₦${fmt(total)}</div><p style="text-align:center;margin-top:40px;color:${COLOR_PRIMARY};font-weight:bold;">Thank you! ${TAGLINE}</p></body></html>`;
}

function generarHTMLEmail(idPedido, datosCliente, carrito, total) {
  const fecha = Utilities.formatDate(new Date(), 'GMT+1', 'dd/MM/yyyy HH:mm');
  const rows = carrito.map(item => `<tr><td style="padding:12px;border-bottom:1px solid #e8e8ff;">${item.nombre}</td><td style="padding:12px;border-bottom:1px solid #e8e8ff;text-align:center;">${item.cantidad} ${item.unidad}</td><td style="padding:12px;border-bottom:1px solid #e8e8ff;text-align:right;font-weight:bold;color:${COLOR_PRIMARY};">₦${fmt(item.subtotal)}</td></tr>`).join('');
  return `<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body style="font-family:Segoe UI,sans-serif;margin:0;padding:0;background:#f5f5f5;"><div style="max-width:600px;margin:20px auto;background:white;border-radius:10px;overflow:hidden;box-shadow:0 4px 6px rgba(0,0,0,0.1);"><div style="background:linear-gradient(135deg,${COLOR_PRIMARY},${COLOR_SECONDARY});color:white;padding:30px;text-align:center;"><img src="${LOGO_URL}" style="max-width:120px;background:white;padding:10px;border-radius:10px;" alt="Logo"><h1 style="margin:10px 0;">Order Confirmed!</h1><p>Thank you for shopping at ${STORE_NAME}</p></div><div style="padding:30px;"><div style="background:#e8e8ff;border-left:4px solid ${COLOR_PRIMARY};padding:15px;border-radius:5px;margin-bottom:20px;"><strong>Order ${idPedido}</strong><br>Date: ${fecha}<br>Status: <span style="color:${COLOR_ACCENT};font-weight:bold;">Pending</span></div><p><strong>Name:</strong> ${datosCliente.nombre}</p><p><strong>Phone:</strong> ${datosCliente.telefono}</p><p><strong>Address:</strong> ${datosCliente.direccion}</p><table style="width:100%;border-collapse:collapse;margin:20px 0;"><thead><tr><th style="background:${COLOR_PRIMARY};color:white;padding:12px;text-align:left;">Product</th><th style="background:${COLOR_PRIMARY};color:white;padding:12px;text-align:center;">Qty</th><th style="background:${COLOR_PRIMARY};color:white;padding:12px;text-align:right;">Subtotal</th></tr></thead><tbody>${rows}</tbody></table><div style="background:${COLOR_PRIMARY};color:white;padding:20px;text-align:center;border-radius:5px;"><p style="margin:0;font-size:16px;opacity:0.9;">TOTAL TO PAY</p><p style="margin:10px 0 0;font-size:36px;font-weight:bold;">₦${fmt(total)}</p></div></div><div style="background:#f5f5f5;padding:20px;text-align:center;font-size:12px;color:#666;"><strong>${STORE_NAME}</strong> | ${ADDRESS}, ${CITY} | 📞 ${PHONE}<p style="margin-top:10px;color:${COLOR_PRIMARY};font-weight:bold;">${TAGLINE}</p></div></div></body></html>`;
}

function generarHTMLEmailVendedor(idPedido, datosCliente, carrito, total) {
  const fecha = Utilities.formatDate(new Date(), 'GMT+1', 'dd/MM/yyyy HH:mm');
  const rows = carrito.map(item => `<tr><td style="padding:12px;border-bottom:1px solid #e8e8ff;">${item.nombre}</td><td style="padding:12px;border-bottom:1px solid #e8e8ff;text-align:center;">${item.cantidad} ${item.unidad}</td><td style="padding:12px;border-bottom:1px solid #e8e8ff;text-align:right;font-weight:bold;color:${COLOR_PRIMARY};">₦${fmt(item.subtotal)}</td></tr>`).join('');
  return `<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body style="font-family:Segoe UI,sans-serif;margin:0;padding:0;background:#f5f5f5;"><div style="max-width:600px;margin:20px auto;background:white;border-radius:10px;overflow:hidden;"><div style="background:linear-gradient(135deg,${COLOR_ACCENT},#f57c00);color:white;padding:30px;text-align:center;"><h1 style="margin:0;font-size:32px;">🔔 NEW ORDER RECEIVED!</h1><p style="opacity:0.9;font-size:18px;">${STORE_NAME}</p></div><div style="padding:30px;"><div style="background:#fff3e0;border-left:4px solid ${COLOR_ACCENT};padding:15px;border-radius:5px;margin-bottom:20px;"><strong>Order ${idPedido}</strong><br>Date: ${fecha}<br>Status: <span style="color:${COLOR_ACCENT};font-weight:bold;">⏳ PENDING - ACTION REQUIRED</span></div><div style="background:#e8f5e9;padding:15px;border-radius:5px;margin-bottom:20px;"><strong>👤 Customer</strong><br>Name: ${datosCliente.nombre}<br>Email: ${datosCliente.email}<br>Phone: ${datosCliente.telefono}<br>Address: ${datosCliente.direccion}${datosCliente.notas?'<br>Notes: '+datosCliente.notas:''}</div><table style="width:100%;border-collapse:collapse;margin:20px 0;"><thead><tr><th style="background:${COLOR_PRIMARY};color:white;padding:12px;text-align:left;">Product</th><th style="background:${COLOR_PRIMARY};color:white;padding:12px;text-align:center;">Qty</th><th style="background:${COLOR_PRIMARY};color:white;padding:12px;text-align:right;">Subtotal</th></tr></thead><tbody>${rows}</tbody></table><div style="background:${COLOR_PRIMARY};color:white;padding:20px;text-align:center;border-radius:5px;"><p style="margin:0;font-size:16px;opacity:0.9;">TOTAL ORDER VALUE</p><p style="margin:10px 0 0;font-size:36px;font-weight:bold;">₦${fmt(total)}</p></div><div style="text-align:center;margin-top:20px;padding:20px;background:#ffebee;border-radius:5px;border:2px solid #f44336;"><p style="color:#c62828;font-weight:bold;font-size:16px;">⚠️ Please contact the customer to confirm and arrange delivery</p></div></div></div></body></html>`;
}

function fmt(n) { return Number(n).toLocaleString('en-NG'); }
