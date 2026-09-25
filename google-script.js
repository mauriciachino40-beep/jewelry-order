var PASSWORD = 'jewelry2026';

// Keep column references stable even when an existing Orders sheet was created
// before a new order field was introduced.
function ensureOrderColumns(sheet, names) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var columns = {};
  headers.forEach(function(header, index) {
    if (header) columns[String(header)] = index + 1;
  });
  names.forEach(function(name) {
    if (!columns[name]) {
      var col = sheet.getLastColumn() + 1;
      sheet.getRange(1, col).setValue(name).setFontWeight('bold').setBackground('#f0f0f0');
      columns[name] = col;
    }
  });
  return columns;
}

function getPaymentsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Payments');
  if (!sheet) {
    sheet = ss.insertSheet('Payments');
    sheet.appendRow(['Date', 'Amount']);
    sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#f0f0f0');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function doGet(e) {
  if (!e || !e.parameter) {
    return json({ status: 'ok', message: 'Jewelry Order API' });
  }
  var password = e.parameter.password;
  if (password !== PASSWORD) {
    return json({ error: 'Wrong password' });
  }
  var action = e.parameter.action;
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ===== Lisa 转账记录（存 Payments 表）=====
  if (action === 'payments_list') {
    var ps = getPaymentsSheet();
    var pd = ps.getDataRange().getValues();
    var list = [];
    for (var i = 1; i < pd.length; i++) {
      list.push({ row: i + 1, date: pd[i][0] || '', amount: Number(pd[i][1]) || 0 });
    }
    return json({ payments: list });
  }
  if (action === 'payments_add') {
    var pDate = e.parameter.date || '';
    var pAmount = parseFloat(e.parameter.amount);
    if (!pDate || !pAmount || pAmount <= 0) return json({ error: 'Missing date or amount' });
    getPaymentsSheet().appendRow([pDate, pAmount]);
    return json({ success: true });
  }
  if (action === 'payments_delete') {
    var pRow = parseInt(e.parameter.row);
    if (pRow && pRow > 1) {
      getPaymentsSheet().deleteRow(pRow);
      return json({ success: true });
    }
    return json({ error: 'Missing or invalid row' });
  }

  var sheet = ss.getSheetByName('Orders');
  if (!sheet) {
    return json({ orders: [] });
  }

  if (action === 'list') {
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var orders = [];
    for (var i = 1; i < data.length; i++) {
      var row = {};
      for (var j = 0; j < headers.length; j++) {
        row[headers[j]] = data[i][j] || '';
      }
      row._row = i + 1;
      orders.push(row);
    }
    orders.reverse();
    return json({ orders: orders });
  }

  if (action === 'status') {
    var row = parseInt(e.parameter.row);
    var status = e.parameter.status;
    if (row && status) {
      var statusCol = ensureOrderColumns(sheet, ['Status'])['Status'];
      sheet.getRange(row, statusCol).setValue(status);
      return json({ success: true });
    }
    return json({ error: 'Missing row or status' });
  }

  if (action === 'delete') {
    var row = parseInt(e.parameter.row);
    if (row && row > 1) {
      // 已发货订单不允许删除（已发出即需付款，删除会丢账）
      var delColumns = ensureOrderColumns(sheet, ['Status', 'Shipments']);
      var delStatus = sheet.getRange(row, delColumns['Status']).getValue();
      var delShipCell = sheet.getRange(row, delColumns['Shipments']).getValue();
      if (delStatus === 'Shipped') {
        return json({ success: false, error: 'Shipped orders cannot be deleted' });
      }
      if (delShipCell) {
        var delShips = [];
        try { delShips = JSON.parse(delShipCell); } catch(e) { delShips = []; }
        if (Array.isArray(delShips) && delShips.length > 0) {
          return json({ success: false, error: 'Order has shipped items and cannot be deleted' });
        }
      }
      sheet.deleteRow(row);
      return json({ success: true });
    }
    return json({ error: 'Missing or invalid row' });
  }

  if (action === 'note') {
    var row = parseInt(e.parameter.row);
    var note = e.parameter.note || '';
    if (row && row > 0) {
      var noteCol = ensureOrderColumns(sheet, ['Admin Note'])['Admin Note'];
      sheet.getRange(row, noteCol).setValue(note);
      return json({ success: true });
    }
    return json({ error: 'Missing or invalid row' });
  }

  if (action === 'ship') {
    var row = parseInt(e.parameter.row);
    var itemsParam = e.parameter.items || '';
    var tracking = e.parameter.tracking || '';
    var note = e.parameter.note || '';
    if (!row || !itemsParam) return json({ error: 'Missing row or items' });
    var items = itemsParam.split(',').map(function(s) { return parseInt(s, 10); }).filter(function(n) { return !isNaN(n); });
    if (items.length === 0) return json({ error: 'No valid items' });

    var shipColumns = ensureOrderColumns(sheet, ['Shipments', 'Status', 'Items Data']);
    var shipmentsCol = shipColumns['Shipments'];

    // 读已有发货记录
    var cellVal = sheet.getRange(row, shipmentsCol).getValue();
    var shipments = [];
    if (cellVal) {
      try { shipments = JSON.parse(cellVal); } catch(e) { shipments = []; }
    }
    if (!Array.isArray(shipments)) shipments = [];

    // 排除已发过的下标
    var already = {};
    shipments.forEach(function(s) { (s.items || []).forEach(function(x) { already[x] = true; }); });
    var newItems = items.filter(function(x) { return !already[x]; });
    if (newItems.length === 0) return json({ success: false, error: 'All selected items already shipped' });

    var date = Utilities.formatDate(new Date(), 'America/New_York', 'yyyy-MM-dd');
    shipments.push({ id: 'S' + Date.now(), items: newItems, tracking: tracking, note: note, date: date });
    sheet.getRange(row, shipmentsCol).setValue(JSON.stringify(shipments));

    // 自动推进整单状态：全部商品发货完 → Shipped
    try {
      var itemsData = sheet.getRange(row, shipColumns['Items Data']).getValue();
      var allItems = itemsData ? JSON.parse(itemsData) : [];
      var shippedSet = {};
      shipments.forEach(function(s) { (s.items || []).forEach(function(x) { shippedSet[x] = true; }); });
      if (Array.isArray(allItems) && allItems.length > 0 && Object.keys(shippedSet).length >= allItems.length) {
        sheet.getRange(row, shipColumns['Status']).setValue('Shipped');
      }
    } catch(e) {}

    return json({ success: true, shipments: shipments, status: sheet.getRange(row, shipColumns['Status']).getValue() });
  }

  if (action === 'unship') {
    var row = parseInt(e.parameter.row);
    var itemsParam = e.parameter.items || '';
    if (!row || !itemsParam) return json({ error: 'Missing row or items' });
    var items = itemsParam.split(',').map(function(s) { return parseInt(s, 10); }).filter(function(n) { return !isNaN(n); });
    var unshipCol = ensureOrderColumns(sheet, ['Shipments'])['Shipments'];
    var cellVal = sheet.getRange(row, unshipCol).getValue();
    var shipments = [];
    if (cellVal) {
      try { shipments = JSON.parse(cellVal); } catch(e) { shipments = []; }
    }
    if (!Array.isArray(shipments)) shipments = [];
    var removeSet = {};
    items.forEach(function(x) { removeSet[x] = true; });
    shipments = shipments.map(function(s) {
      s.items = (s.items || []).filter(function(x) { return !removeSet[x]; });
      return s;
    }).filter(function(s) { return s.items && s.items.length > 0; });
    sheet.getRange(row, unshipCol).setValue(JSON.stringify(shipments));
    return json({ success: true, shipments: shipments });
  }

  return json({ error: 'Unknown action' });
}

function doPost(e) {
  try {
    var raw = e.postData ? e.postData.contents : (e.parameter && e.parameter.data ? e.parameter.data : null);
    if (!raw) {
      return json({ success: false, error: 'No data received' });
    }
    var data = JSON.parse(raw);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Orders');

    if (!sheet) {
      sheet = ss.insertSheet('Orders');
      sheet.appendRow(['Order ID', 'Date (ET)', 'Customer Name', 'Email', 'Phone', 'Shipping Address', 'Items', 'Total', 'Notes', 'Status', 'Items Data', 'Admin Note', 'Shipments', 'Order Type', 'Related Order ID', 'Compensation Reason', 'Reference Total', 'Complimentary Total']);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, 18).setFontWeight('bold').setBackground('#f0f0f0');
    }

    var columns = ensureOrderColumns(sheet, [
      'Order ID', 'Date (ET)', 'Customer Name', 'Email', 'Phone', 'Shipping Address',
      'Items', 'Total', 'Notes', 'Status', 'Items Data', 'Admin Note', 'Shipments',
      'Order Type', 'Related Order ID', 'Compensation Reason', 'Reference Total', 'Complimentary Total'
    ]);

    var itemsStr = data.items.map(function(i) {
      var v = Object.values(i.variants || {}).join(', ');
      var lineTotal = Number(i.subtotal) || 0;
      var billingType = i.billingType || 'Regular';
      if (billingType === 'Compensation') return i.name + ' (' + v + ') x' + i.quantity + ' - Compensation / No charge (ref. $' + lineTotal.toFixed(2) + ')';
      if (billingType === 'Gift') return i.name + ' (' + v + ') x' + i.quantity + ' - Gift / No charge (ref. $' + lineTotal.toFixed(2) + ')';
      return i.name + ' (' + v + ') x' + i.quantity + ' - $' + lineTotal.toFixed(2);
    }).join('\n');

    var addrParts = [data.customer.address, data.customer.apt, data.customer.city + ', ' + data.customer.state + ' ' + data.customer.zip].filter(Boolean);
    var addr = addrParts.join(' ');

    // 完整 items JSON 存入第 11 列
    var itemsData = JSON.stringify(data.items);

    var row = new Array(sheet.getLastColumn()).fill('');
    row[columns['Order ID'] - 1] = data.orderId;
    row[columns['Date (ET)'] - 1] = data.date;
    row[columns['Customer Name'] - 1] = data.customer.name;
    row[columns['Email'] - 1] = data.customer.email || '';
    row[columns['Phone'] - 1] = data.customer.phone || '';
    row[columns['Shipping Address'] - 1] = addr;
    row[columns['Items'] - 1] = itemsStr;
    row[columns['Total'] - 1] = (Number(data.total) || 0).toFixed(2);
    row[columns['Notes'] - 1] = data.notes || '';
    row[columns['Status'] - 1] = 'New';
    row[columns['Items Data'] - 1] = itemsData;
    row[columns['Order Type'] - 1] = data.orderType || 'Standard';
    row[columns['Related Order ID'] - 1] = data.relatedOrderId || '';
    row[columns['Compensation Reason'] - 1] = data.compensationReason || '';
    row[columns['Reference Total'] - 1] = (Number(data.referenceTotal) || Number(data.total) || 0).toFixed(2);
    row[columns['Complimentary Total'] - 1] = (Number(data.complimentaryTotal) || 0).toFixed(2);
    sheet.appendRow(row);

    sheet.autoResizeColumns(1, 12);
    return json({ success: true, orderId: data.orderId });
  } catch (error) {
    return json({ success: false, error: error.toString() });
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
