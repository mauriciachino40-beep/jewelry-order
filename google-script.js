var PASSWORD = 'LISA';

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
      sheet.getRange(row, 10).setValue(status);
      return json({ success: true });
    }
    return json({ error: 'Missing row or status' });
  }

  if (action === 'delete') {
    var row = parseInt(e.parameter.row);
    if (row && row > 1) {
      sheet.deleteRow(row);
      return json({ success: true });
    }
    return json({ error: 'Missing or invalid row' });
  }

  if (action === 'note') {
    var row = parseInt(e.parameter.row);
    var note = e.parameter.note || '';
    if (row && row > 0) {
      // 确保第 12 列有表头
      var lastCol = sheet.getLastColumn();
      if (lastCol < 12) {
        sheet.getRange(1, 12).setValue('Admin Note');
        sheet.getRange(1, 12).setFontWeight('bold').setBackground('#f0f0f0');
      }
      sheet.getRange(row, 12).setValue(note);
      return json({ success: true });
    }
    return json({ error: 'Missing or invalid row' });
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
      sheet.appendRow(['Order ID', 'Date (ET)', 'Customer Name', 'Email', 'Phone', 'Shipping Address', 'Items', 'Total', 'Notes', 'Status', 'Items Data', 'Admin Note']);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, 12).setFontWeight('bold').setBackground('#f0f0f0');
    }

    // 确保有 Items Data 和 Admin Note 列
    var lastCol = sheet.getLastColumn();
    if (lastCol < 11) {
      sheet.getRange(1, 11).setValue('Items Data');
      sheet.getRange(1, 11).setFontWeight('bold').setBackground('#f0f0f0');
    }
    if (lastCol < 12) {
      sheet.getRange(1, 12).setValue('Admin Note');
      sheet.getRange(1, 12).setFontWeight('bold').setBackground('#f0f0f0');
    }

    var itemsStr = data.items.map(function(i) {
      var v = Object.values(i.variants || {}).join(', ');
      return i.name + ' (' + v + ') x' + i.quantity + ' - $' + i.subtotal.toFixed(2);
    }).join('\n');

    var addrParts = [data.customer.address, data.customer.apt, data.customer.city + ', ' + data.customer.state + ' ' + data.customer.zip].filter(Boolean);
    var addr = addrParts.join(' ');

    // 完整 items JSON 存入第 11 列
    var itemsData = JSON.stringify(data.items);

    sheet.appendRow([
      data.orderId,
      data.date,
      data.customer.name,
      data.customer.email || '',
      data.customer.phone || '',
      addr,
      itemsStr,
      data.total.toFixed(2),
      data.notes || '',
      'New',
      itemsData
    ]);

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
