/**
 * ProyojonBD — Apps Script Backend (v3)
 *
 * "Orders" শীটের হেডার (এই ক্রমে, এই নামে):
 *   orderId | timestamp | customerName | phone | productName | deliveryArea | address | quantity | unitPrice | subtotal | deliveryCharge | total | status
 *
 * "Products" শীটের হেডার:
 *   id | name | desc | price | discount | stock | image | status
 *   (discount কলাম নতুন — না থাকলে খালি রাখুন, ০ ধরা হবে)
 *
 * "Expenses" শীটের হেডার:
 *   expenseId | timestamp | category | description | amount
 */

function doGet(e) {
  var action = e.parameter.action;
  if (action === 'products') return jsonResponse(getProducts());
  if (action === 'orders') return jsonResponse(getOrders(e.parameter.status));
  if (action === 'expenses') return jsonResponse(getExpenses());
  return jsonResponse({ error: 'Unknown action' });
}

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return jsonResponse({ error: 'Invalid JSON' }); }
  switch (body.action) {
    case 'order': return jsonResponse(saveOrder(body.order));
    case 'updateOrderStatus': return jsonResponse(updateOrderStatus(body.orderId, body.status));
    case 'addProduct': return jsonResponse(addProduct(body.product));
    case 'updateProduct': return jsonResponse(updateProduct(body.product));
    case 'addExpense': return jsonResponse(addExpense(body.expense));
    default: return jsonResponse({ error: 'Unknown action' });
  }
}

function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
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

/* -------- Products -------- */
function getProducts() {
  var sheet = getSheet('Products');
  var rows = sheet.getDataRange().getValues();
  var all = rowsToObjects(rows);
  return {
    products: all.map(function(o){
      return {
        id: String(o.id),
        name: o.name || '',
        desc: o.desc || '',
        price: Number(o.price) || 0,
        discount: Number(o.discount) || 0,
        stock: (o.stock === '' || o.stock === undefined) ? null : Number(o.stock),
        image: o.image || '',
        status: o.status ? String(o.status).trim() : 'Active'
      };
    })
  };
}

function addProduct(p) {
  if (!p || !p.name || p.price === undefined) return { error: 'নাম ও দাম আবশ্যক' };
  var sheet = getSheet('Products');
  var id = 'P' + new Date().getTime();
  sheet.appendRow([id, p.name, p.desc||'', Number(p.price)||0, Number(p.discount)||0,
    (p.stock===undefined||p.stock==='')?'':Number(p.stock), p.image||'', 'Active']);
  return { success: true, id: id };
}

function updateProduct(p) {
  if (!p || !p.id) return { error: 'id প্রয়োজন' };
  var sheet = getSheet('Products');
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0].map(function(h){ return String(h).trim().toLowerCase(); });
  var idCol = headers.indexOf('id');
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][idCol]) === String(p.id)) {
      var r = i + 1;
      var set = function(col, val){ if(val!==undefined){ var c=headers.indexOf(col); if(c>=0) sheet.getRange(r,c+1).setValue(val); } };
      set('name', p.name); set('desc', p.desc);
      set('price', p.price!==undefined?Number(p.price):undefined);
      set('discount', p.discount!==undefined?Number(p.discount):undefined);
      set('stock', p.stock!==undefined?(p.stock===''?'':Number(p.stock)):undefined);
      set('image', p.image); set('status', p.status);
      return { success: true };
    }
  }
  return { error: 'পণ্য পাওয়া যায়নি' };
}

/* -------- Orders -------- */
function getOrders(status) {
  var sheet = getSheet('Orders');
  var rows = sheet.getDataRange().getValues();
  var all = rowsToObjects(rows);
  var orders = all.filter(function(o){
    if (!status) return true;
    return (o.status||'Pending').toString().trim().toLowerCase() === status.toLowerCase();
  }).map(function(o){
    return {
      orderId:      o.orderid,
      timestamp:    o.timestamp,
      customerName: o.customername,
      phone:        o.phone,
      productName:  o.productname,
      deliveryArea: o.deliveryarea,
      address:      o.address,
      quantity:     o.quantity,
      unitPrice:    o.unitprice,
      subtotal:     o.subtotal,
      deliveryCharge: o.deliverycharge,
      total:        o.total,
      status:       (o.status||'Pending').toString().trim()
    };
  });
  return { orders: orders.reverse() };
}

function saveOrder(order) {
  if (!order || !order.productName || !order.customerName) return { error: 'তথ্য অসম্পূর্ণ' };
  var sheet = getSheet('Orders');
  var orderId  = 'PRY-' + new Date().getTime();
  var timestamp = new Date();
  var qty      = Number(order.quantity)     || 1;
  var unit     = Number(order.unitPrice)    || 0;
  var sub      = Number(order.subtotal)     || (unit * qty);
  var dc       = Number(order.deliveryCharge) || 0;
  var total    = Number(order.total)        || (sub + dc);
  var area     = order.deliveryArea         || 'Inside Feni';

  // হেডার অনুযায়ী সঠিক ক্রম:
  // orderId | timestamp | customerName | phone | productName | deliveryArea | address | quantity | unitPrice | subtotal | deliveryCharge | total | status
  sheet.appendRow([
    orderId, timestamp,
    order.customerName || '',
    order.phone        || '',
    order.productName  || '',
    area,
    order.address      || '',
    qty, unit, sub, dc, total, 'Pending'
  ]);

  return { success:true, orderId:orderId, timestamp:timestamp.toISOString(),
           subtotal:sub, deliveryArea:area, deliveryCharge:dc, total:total };
}

function updateOrderStatus(orderId, status) {
  if (!orderId || !status) return { error: 'orderId ও status দরকার' };
  var sheet = getSheet('Orders');
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0].map(function(h){ return String(h).trim().toLowerCase().replace(/\s+/g,''); });
  var idCol  = headers.indexOf('orderid');
  var stCol  = headers.indexOf('status');
  if (idCol < 0 || stCol < 0) return { error: 'কলাম পাওয়া যায়নি' };
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][idCol]) === String(orderId)) {
      sheet.getRange(i + 1, stCol + 1).setValue(status);
      return { success: true };
    }
  }
  return { error: 'অর্ডার পাওয়া যায়নি' };
}

/* -------- Expenses -------- */
function getExpenses() {
  var sheet = getSheet('Expenses');
  var rows = sheet.getDataRange().getValues();
  return { expenses: rowsToObjects(rows).map(function(o){
    return { expenseId:o.expenseid, timestamp:o.timestamp, category:o.category||'', description:o.description||'', amount:Number(o.amount)||0 };
  }).reverse() };
}

function addExpense(exp) {
  if (!exp || exp.amount===undefined) return { error: 'amount দরকার' };
  var sheet = getSheet('Expenses');
  var id = 'EXP-' + new Date().getTime();
  sheet.appendRow([id, new Date(), exp.category||'অন্যান্য', exp.description||'', Number(exp.amount)||0]);
  return { success:true, expenseId:id };
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
