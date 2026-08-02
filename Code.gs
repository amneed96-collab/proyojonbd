/**
 * ProyojonBD — Apps Script Backend (v5)
 * শীট ও হেডার অটো তৈরি হবে — আলাদাভাবে কিছু করতে হবে না।
 */

/* ======== SHEET & HEADER AUTO-SETUP ======== */

var SHEET_SCHEMAS = {
  'Products': ['id','name','category','desc','price','discount','stock','image','status'],
  'Orders':   ['orderId','timestamp','customerName','phone','productName','category','deliveryArea','address','quantity','unitPrice','subtotal','deliveryCharge','total','status'],
  'Expenses': ['expenseId','timestamp','category','description','amount']
};

function ensureSheet(name) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(SHEET_SCHEMAS[name]);
    sheet.getRange(1, 1, 1, SHEET_SCHEMAS[name].length)
         .setFontWeight('bold')
         .setBackground('#0E6E55')
         .setFontColor('#FFFFFF');
  } else {
    // হেডার না থাকলে বা ভুল থাকলে ঠিক করা
    var firstRow = sheet.getRange(1, 1, 1, SHEET_SCHEMAS[name].length).getValues()[0];
    if (!firstRow[0] || String(firstRow[0]).trim() === '') {
      sheet.insertRowBefore(1);
      sheet.getRange(1, 1, 1, SHEET_SCHEMAS[name].length).setValues([SHEET_SCHEMAS[name]]);
      sheet.getRange(1, 1, 1, SHEET_SCHEMAS[name].length)
           .setFontWeight('bold').setBackground('#0E6E55').setFontColor('#FFFFFF');
    }
  }
  return sheet;
}

function ensureAllSheets() {
  Object.keys(SHEET_SCHEMAS).forEach(function(name){ ensureSheet(name); });
}

/* ======== UTILITIES ======== */

function getHeaders(sheet) {
  var vals = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return vals.map(function(h){ return String(h).trim().toLowerCase().replace(/\s+/g,''); });
}

function rowsToObjects(rows) {
  if (!rows || rows.length < 2) return [];
  var headers = rows[0].map(function(h){ return String(h).trim().toLowerCase().replace(/\s+/g,''); });
  var out = [];
  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    if (!row[0] && !row[1]) continue;
    var obj = { _row: i + 1 };
    headers.forEach(function(h, idx){ obj[h] = row[idx]; });
    out.push(obj);
  }
  return out;
}

function buildRow(headers, map) {
  var row = new Array(headers.length).fill('');
  headers.forEach(function(h, i){ if (map.hasOwnProperty(h)) row[i] = map[h]; });
  return row;
}

/* ======== HTTP HANDLERS ======== */

function doGet(e) {
  ensureAllSheets();
  var action = e.parameter.action;
  if (action === 'products') return jsonResponse(getProducts());
  if (action === 'orders')   return jsonResponse(getOrders(e.parameter.status));
  if (action === 'expenses') return jsonResponse(getExpenses());
  return jsonResponse({ error: 'Unknown action' });
}

function doPost(e) {
  ensureAllSheets();
  var body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return jsonResponse({ error: 'Invalid JSON' }); }
  switch (body.action) {
    case 'order':             return jsonResponse(saveOrder(body.order));
    case 'updateOrderStatus': return jsonResponse(updateOrderStatus(body.orderId, body.status));
    case 'addProduct':        return jsonResponse(addProduct(body.product));
    case 'updateProduct':     return jsonResponse(updateProduct(body.product));
    case 'addExpense':        return jsonResponse(addExpense(body.expense));
    default:                  return jsonResponse({ error: 'Unknown action' });
  }
}

/* ======== PRODUCTS ======== */

function getProducts() {
  var sheet = ensureSheet('Products');
  var rows  = sheet.getDataRange().getValues();
  return {
    products: rowsToObjects(rows).map(function(o){
      return {
        id:       String(o.id),
        name:     o.name       || '',
        category: o.category   || 'Others',
        desc:     o.desc       || '',
        price:    Number(o.price)    || 0,
        discount: Number(o.discount) || 0,
        stock:    (o.stock===''||o.stock===undefined) ? null : Number(o.stock),
        image:    o.image      || '',
        status:   o.status     ? String(o.status).trim() : 'Active'
      };
    })
  };
}

function addProduct(p) {
  if (!p || !p.name || p.price===undefined) return { error: 'নাম ও দাম আবশ্যক' };
  var sheet   = ensureSheet('Products');
  var headers = getHeaders(sheet);
  var id      = 'P' + new Date().getTime();
  sheet.appendRow(buildRow(headers, {
    'id':       id,
    'name':     p.name,
    'category': p.category || 'Others',
    'desc':     p.desc     || '',
    'price':    Number(p.price)    || 0,
    'discount': Number(p.discount) || 0,
    'stock':    (p.stock===undefined||p.stock==='') ? '' : Number(p.stock),
    'image':    p.image    || '',
    'status':   'Active'
  }));
  return { success: true, id: id };
}

function updateProduct(p) {
  if (!p || !p.id) return { error: 'id প্রয়োজন' };
  var sheet   = ensureSheet('Products');
  var rows    = sheet.getDataRange().getValues();
  var headers = rows[0].map(function(h){ return String(h).trim().toLowerCase(); });
  var idCol   = headers.indexOf('id');
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][idCol]) !== String(p.id)) continue;
    var r   = i + 1;
    var set = function(col, val){
      if (val === undefined) return;
      var c = headers.indexOf(col);
      if (c >= 0) sheet.getRange(r, c+1).setValue(val);
    };
    set('name',     p.name);
    set('category', p.category);
    set('desc',     p.desc);
    if (p.price    !== undefined) set('price',    Number(p.price));
    if (p.discount !== undefined) set('discount', Number(p.discount));
    if (p.stock    !== undefined) set('stock',    p.stock==='' ? '' : Number(p.stock));
    set('image',    p.image);
    set('status',   p.status);
    return { success: true };
  }
  return { error: 'পণ্য পাওয়া যায়নি' };
}

/* ======== ORDERS ======== */

function getOrders(status) {
  var sheet  = ensureSheet('Orders');
  var rows   = sheet.getDataRange().getValues();
  var orders = rowsToObjects(rows)
    .filter(function(o){
      if (!status) return true;
      return (o.status||'Pending').toString().trim().toLowerCase() === status.toLowerCase();
    })
    .map(function(o){
      return {
        orderId:       o.orderid,
        timestamp:     o.timestamp,
        customerName:  o.customername,
        phone:         o.phone,
        productName:   o.productname,
        category:      o.category || '',
        deliveryArea:  o.deliveryarea,
        address:       o.address,
        quantity:      o.quantity,
        unitPrice:     o.unitprice,
        subtotal:      o.subtotal,
        deliveryCharge: o.deliverycharge,
        total:         o.total,
        status:        (o.status||'Pending').toString().trim()
      };
    });
  return { orders: orders.reverse() };
}

function saveOrder(order) {
  if (!order || !order.productName || !order.customerName) return { error: 'তথ্য অসম্পূর্ণ' };
  var sheet   = ensureSheet('Orders');
  var headers = getHeaders(sheet);
  var orderId = 'PRY-' + new Date().getTime();
  var ts      = new Date();
  var qty     = Number(order.quantity)       || 1;
  var unit    = Number(order.unitPrice)      || 0;
  var sub     = Number(order.subtotal)       || (unit * qty);
  var dc      = Number(order.deliveryCharge) || 0;
  var total   = Number(order.total)          || (sub + dc);
  var area    = order.deliveryArea           || 'Inside Feni';

  sheet.appendRow(buildRow(headers, {
    'orderid':        orderId,
    'timestamp':      ts,
    'customername':   order.customerName || '',
    'phone':          order.phone        || '',
    'productname':    order.productName  || '',
    'category':       order.category     || '',
    'deliveryarea':   area,
    'address':        order.address      || '',
    'quantity':       qty,
    'unitprice':      unit,
    'subtotal':       sub,
    'deliverycharge': dc,
    'total':          total,
    'status':         'Pending'
  }));

  return { success:true, orderId:orderId, timestamp:ts.toISOString(),
           subtotal:sub, deliveryArea:area, deliveryCharge:dc, total:total };
}

function updateOrderStatus(orderId, status) {
  if (!orderId || !status) return { error: 'orderId ও status দরকার' };
  var sheet   = ensureSheet('Orders');
  var rows    = sheet.getDataRange().getValues();
  var headers = rows[0].map(function(h){ return String(h).trim().toLowerCase().replace(/\s+/g,''); });
  var idCol   = headers.indexOf('orderid');
  var stCol   = headers.indexOf('status');
  if (idCol < 0 || stCol < 0) return { error: 'কলাম পাওয়া যায়নি' };
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][idCol]) === String(orderId)) {
      sheet.getRange(i+1, stCol+1).setValue(status);
      return { success: true };
    }
  }
  return { error: 'অর্ডার পাওয়া যায়নি' };
}

/* ======== EXPENSES ======== */

function getExpenses() {
  var sheet = ensureSheet('Expenses');
  var rows  = sheet.getDataRange().getValues();
  return {
    expenses: rowsToObjects(rows).map(function(o){
      return { expenseId:o.expenseid, timestamp:o.timestamp,
               category:o.category||'', description:o.description||'', amount:Number(o.amount)||0 };
    }).reverse()
  };
}

function addExpense(exp) {
  if (!exp || exp.amount===undefined) return { error: 'amount দরকার' };
  var sheet   = ensureSheet('Expenses');
  var headers = getHeaders(sheet);
  var id      = 'EXP-' + new Date().getTime();
  sheet.appendRow(buildRow(headers, {
    'expenseid':   id,
    'timestamp':   new Date(),
    'category':    exp.category    || 'অন্যান্য',
    'description': exp.description || '',
    'amount':      Number(exp.amount) || 0
  }));
  return { success:true, expenseId:id };
}

/* ======== JSON ======== */
function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
