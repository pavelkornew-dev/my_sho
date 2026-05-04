// ============================================================
// SNEAKER SHOP — Google Apps Script Backend
// Вставте цей код у Apps Script → Code.gs
// ============================================================

const SPREADSHEET_ID = ''; // ← Вставте ID вашої Google Таблиці
const SHEET_ORDERS   = 'Замовлення';
const SHEET_PRODUCTS = 'Товари';
const ADMIN_PASSWORD = 'admin123'; // ← Змініть пароль

// ---------- ENTRY POINTS ----------

function doGet(e) {
  const page = e.parameter.page || 'shop';
  if (page === 'admin') {
    return HtmlService.createTemplateFromFile('admin')
      .evaluate()
      .setTitle('Адмін-панель | SneakerHub')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('SneakerHub — Чоловічі кросівки')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ---------- PRODUCTS ----------

function getProducts() {
  const ss    = getSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_PRODUCTS, ['ID','Назва','Бренд','Ціна','Розміри','Опис','Зображення (URL)','Наявність','Дата додавання']);
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(r => ({
    id:          r[0],
    name:        r[1],
    brand:       r[2],
    price:       r[3],
    sizes:       r[4],
    description: r[5],
    image:       r[6],
    inStock:     r[7],
    date:        r[8]
  })).filter(p => p.id);
}

function addProduct(product) {
  const ss    = getSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_PRODUCTS, ['ID','Назва','Бренд','Ціна','Розміри','Опис','Зображення (URL)','Наявність','Дата додавання']);
  const id    = 'P' + Date.now();
  sheet.appendRow([id, product.name, product.brand, product.price, product.sizes, product.description, product.image, product.inStock ? 'Так' : 'Ні', new Date().toLocaleString('uk-UA')]);
  return { success: true, id };
}

function deleteProduct(id) {
  const ss    = getSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_PRODUCTS, []);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) { sheet.deleteRow(i + 1); return { success: true }; }
  }
  return { success: false };
}

function updateProductStock(id, inStock) {
  const ss    = getSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_PRODUCTS, []);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.getRange(i + 1, 8).setValue(inStock ? 'Так' : 'Ні');
      return { success: true };
    }
  }
  return { success: false };
}

// ---------- ORDERS ----------

function getOrders() {
  const ss    = getSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_ORDERS, ['ID','Ім\'я','Телефон','Email','Місто','Адреса','Товар','Розмір','Кількість','Сума','Коментар','Статус','Дата']);
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(r => ({
    id:       r[0],  name:     r[1],  phone:    r[2],
    email:    r[3],  city:     r[4],  address:  r[5],
    product:  r[6],  size:     r[7],  qty:      r[8],
    total:    r[9],  comment:  r[10], status:   r[11], date: r[12]
  })).filter(o => o.id);
}

function createOrder(order) {
  const ss    = getSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_ORDERS, ['ID','Ім\'я','Телефон','Email','Місто','Адреса','Товар','Розмір','Кількість','Сума','Коментар','Статус','Дата']);
  const id    = 'ORD-' + Date.now();
  sheet.appendRow([id, order.name, order.phone, order.email, order.city, order.address,
                   order.product, order.size, order.qty, order.total, order.comment || '', 'Нове', new Date().toLocaleString('uk-UA')]);
  return { success: true, orderId: id };
}

function updateOrderStatus(id, status) {
  const ss    = getSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_ORDERS, []);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.getRange(i + 1, 12).setValue(status);
      return { success: true };
    }
  }
  return { success: false };
}

// ---------- AUTH ----------

function checkAdminPassword(password) {
  return password === ADMIN_PASSWORD;
}

// ---------- HELPERS ----------

function getSpreadsheet() {
  if (SPREADSHEET_ID) return SpreadsheetApp.openById(SPREADSHEET_ID);
  // Якщо ID не вказано — використовуємо активну таблицю (якщо скрипт прив'язаний до неї)
  try { return SpreadsheetApp.getActiveSpreadsheet(); }
  catch(e) { throw new Error('Вкажіть SPREADSHEET_ID у коді!'); }
}

function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers.length) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
