/**
 * ProyojonBD — Apps Script Backend (v2)
 * ---------------------------------------------------
 * সাপোর্টেড অ্যাকশন:
 *   GET  ?action=products              -> সব পণ্য (Active ও Inactive সবগুলো; ফ্রন্টএন্ড Active শো করে)
 *   GET  ?action=orders&status=X       -> নির্দিষ্ট status-এর অর্ডার (status না দিলে সব)
 *   GET  ?action=expenses              -> সব ব্যয়ের তালিকা
 *   POST { action:"order" }            -> নতুন অর্ডার যুক্ত (status=Pending)
 *   POST { action:"addProduct" }       -> নতুন পণ্য যুক্ত
 *   POST { action:"updateProduct" }    -> পণ্য তথ্য আপডেট (পার্শিয়াল — যেকোনো ফিল্ড, status টগলও এই দিয়ে হয়)
 *   POST { action:"addExpense" }       -> নতুন ব্যয় এন্ট্রি
 *
 * SHEET STRUCTURE
 * ---------------
 * "Products" ট্যাব হেডার (এই অর্ডারেই, ঠিক এই নামে):
 *   id | name | desc | price | stock | image | status
 *   status কলামে লিখবেন: Active বা Inactive
 *
 * "Orders" ট্যাব হেডার:
 *   orderId | timestamp | customerName | phone | productName | deliveryArea | address | quantity | unitPrice | subtotal | deliveryCharge | total | status
 *   deliveryArea: "Inside Feni" (চার্জ ৬০) বা "Outside Feni" (চার্জ ১৫০)
 *   status: Pending / Delivery / Cancel (নতুন অর্ডার সবসময় Pending দিয়ে শুরু হয়, ম্যানুয়ালি বদলান)
 *
 * "Expenses" ট্যাব হেডার:
 *   expenseId | timestamp | category | description | amount
 *
 * DEPLOY
 * 1. Sheet-এ ৩টা ট্যাব বানান: Products, Orders, Expenses — উপরের হেডার বসান।
 * 2. Extensions > Apps Script এ এই পুরো কোড পেস্ট করে Save করুন।
 * 3. Deploy > New deployment > Web app
 *      Execute as: Me | Who has access: Anyone
 * 4. যে URL পাবেন, index.html-এর শুরুতে API_URL ভ্যারিয়েবলে পেস্ট করুন।
 * 5. কোড পরিবর্তন করলে অবশ্যই নতুন deployment version বানাতে হবে (Manage deployments > Edit > New version)।
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
  catch (err) { return jsonResponse({ error: 'Invalid JSON body' }); }

  switch (body.action) {
    case 'order': return jsonResponse(saveOrder(body.order));
    case 'addProduct': return jsonResponse(addProduct(body.product));
    case 'updateProduct': return jsonResponse(updateProduct(body.product));
    case 'addExpense': return jsonResponse(addExpense(body.expense));
    default: return jsonResponse({ error: 'Unknown action' });
  }
}

/* ---------------- ইউটিলিটি ---------------- */

function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function rowsToObjects(rows) {
  if (!rows || rows.length < 1) return [];
  var headers = rows[0].map(function (h) { return String(h).trim().toLowerCase(); });
  var out = [];
  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    if (!row[0]) continue;
    var obj = { _row: i + 1 };
    headers.forEach(function (h, idx) { obj[h] = row[idx]; });
    out.push(obj);
  }
  return out;
}

/* ---------------- Products ---------------- */

function getProducts() {
  var sheet = getSheet('Products');
  var rows = sheet.getDataRange().getValues();
  var all = rowsToObjects(rows);

  var products = all.map(function (o) {
    return {
      id: String(o.id),
      name: o.name || '',
      desc: o.desc || '',
      price: Number(o.price) || 0,
      stock: (o.stock === '' || o.stock === undefined) ? null : Number(o.stock),
      image: o.image || '',
      status: o.status ? String(o.status).trim() : 'Active'
    };
  });
  return { products: products };
}

function addProduct(p) {
  if (!p || !p.name || p.price === undefined || p.price === '') {
    return { error: 'নাম ও দাম আবশ্যক' };
  }
  var sheet = getSheet('Products');
  var id = 'P' + new Date().getTime();
  sheet.appendRow([
    id, p.name, p.desc || '', Number(p.price) || 0,
    p.stock === undefined || p.stock === '' ? '' : Number(p.stock),
    p.image || '', 'Active'
  ]);
  return { success: true, id: id };
}

function updateProduct(p) {
  if (!p || !p.id) return { error: 'পণ্যের id প্রয়োজন' };
  var sheet = getSheet('Products');
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0].map(function (h) { return String(h).trim().toLowerCase(); });
  var idCol = headers.indexOf('id');

  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][idCol]) === String(p.id)) {
      var rowNum = i + 1;
      setIfDefined(sheet, rowNum, headers, 'name', p.name);
      setIfDefined(sheet, rowNum, headers, 'desc', p.desc);
      setIfDefined(sheet, rowNum, headers, 'price', p.price !== undefined ? Number(p.price) : undefined);
      setIfDefined(sheet, rowNum, headers, 'stock', p.stock !== undefined ? (p.stock === '' ? '' : Number(p.stock)) : undefined);
      setIfDefined(sheet, rowNum, headers, 'image', p.image);
      setIfDefined(sheet, rowNum, headers, 'status', p.status);
      return { success: true };
    }
  }
  return { error: 'পণ্য পাওয়া যায়নি' };
}

function setIfDefined(sheet, rowNum, headers, colName, value) {
  if (value === undefined) return;
  var col = headers.indexOf(colName);
  if (col === -1) return;
  sheet.getRange(rowNum, col + 1).setValue(value);
}

/* ---------------- Orders ---------------- */

function getOrders(status) {
  var sheet = getSheet('Orders');
  var rows = sheet.getDataRange().getValues();
  var all = rowsToObjects(rows);

  var orders = all
    .filter(function (o) {
      if (!status) return true;
      var s = (o.status || 'Pending').toString().trim().toLowerCase();
      return s === status.toLowerCase();
    })
    .map(function (o) {
      return {
        orderId: o.orderid,
        timestamp: o.timestamp,
        customerName: o.customername,
        phone: o.phone,
        productName: o.productname,
        deliveryArea: o.deliveryarea,
        address: o.address,
        quantity: o.quantity,
        unitPrice: o.unitprice,
        subtotal: o.subtotal,
        deliveryCharge: o.deliverycharge,
        total: o.total,
        status: (o.status || 'Pending').toString().trim()
      };
    });

  return { orders: orders.reverse() };
}

function saveOrder(order) {
  if (!order || !order.productName || !order.customerName) {
    return { error: 'প্রয়োজনীয় তথ্য অনুপস্থিত' };
  }
  var sheet = getSheet('Orders');
  var orderId = 'PRY-' + new Date().getTime();
  var timestamp = new Date();

  var quantity = Number(order.quantity) || 1;
  var unitPrice = Number(order.unitPrice) || 0;
  var subtotal = order.subtotal !== undefined ? Number(order.subtotal) : unitPrice * quantity;
  var deliveryArea = order.deliveryArea || 'Inside Feni';
  var deliveryCharge = Number(order.deliveryCharge) || 0;
  var total = order.total !== undefined ? Number(order.total) : subtotal + deliveryCharge;

  sheet.appendRow([
    orderId, timestamp, order.customerName || '', order.phone || '',
    order.productName || '', deliveryArea, order.address || '',
    quantity, unitPrice, subtotal, deliveryCharge, total, 'Pending'
  ]);

  return {
    success: true,
    orderId: orderId,
    timestamp: timestamp.toISOString(),
    subtotal: subtotal,
    deliveryArea: deliveryArea,
    deliveryCharge: deliveryCharge,
    total: total
  };
}

/* ---------------- Expenses ---------------- */

function getExpenses() {
  var sheet = getSheet('Expenses');
  var rows = sheet.getDataRange().getValues();
  var all = rowsToObjects(rows);
  var expenses = all.map(function (o) {
    return {
      expenseId: o.expenseid,
      timestamp: o.timestamp,
      category: o.category || 'অন্যান্য',
      description: o.description || '',
      amount: Number(o.amount) || 0
    };
  });
  return { expenses: expenses.reverse() };
}

function addExpense(exp) {
  if (!exp || exp.amount === undefined || exp.amount === '') return { error: 'amount আবশ্যক' };
  var sheet = getSheet('Expenses');
  var expenseId = 'EXP-' + new Date().getTime();
  var timestamp = new Date();
  sheet.appendRow([expenseId, timestamp, exp.category || 'অন্যান্য', exp.description || '', Number(exp.amount) || 0]);
  return { success: true, expenseId: expenseId };
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
