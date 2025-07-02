/**
 * Hyggelyカンパーニュ専門店 予約管理システム - メール送信機能完全修正版
 * v5.1.0 - メール送信問題解決版
 */

// ===== システム設定 =====
const SYSTEM_CONFIG = {
  spreadsheetId: '18Wdo9hYY8KBF7KULuD8qAODDd5z4O_WvkMCekQpptJ8',
  sheets: {
    ORDER: '予約管理票',
    INVENTORY: '在庫管理票',
    PRODUCT_MASTER: '商品マスタ',
    EMAIL_SETTINGS: 'メール設定',
    SYSTEM_LOG: 'システムログ'
  },
  adminPassword: 'hyggelyAdmin2024',
  version: '5.1.0'
};

// ===== メインエントリーポイント =====
function doGet(e) {
  try {
    console.log('🍞 Hyggelyシステム起動 v' + SYSTEM_CONFIG.version);
    
    const params = e?.parameter || {};
    const action = params.action || '';
    const password = params.password || '';
    
    // 初期化チェック
    checkAndInitializeSystem();
    
    // ルーティング
    switch (action) {
      case 'dashboard':
        return handleDashboard(password);
      case 'email':
        return handleEmailSettings(password);
      case 'health':
        return handleHealthCheck();
      default:
        return handleOrderForm();
    }
  } catch (error) {
    console.error('❌ システムエラー:', error);
    logSystemEvent('ERROR', 'システム起動エラー', error.toString());
    return createErrorPage('システムエラーが発生しました', error.toString());
  }
}

// ===== HTMLファイル読み込み関数 =====
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (error) {
    console.error('❌ HTMLファイル読み込みエラー:', filename, error);
    return '<div>HTMLファイルの読み込みに失敗しました: ' + filename + '</div>';
  }
}

// ===== ルーティング処理 =====
function handleOrderForm() {
  try {
    return HtmlService.createHtmlOutputFromFile('OrderForm')
      .setTitle('Hyggelyカンパーニュ専門店 ご予約フォーム')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    console.error('❌ OrderForm読み込みエラー:', error);
    return createErrorPage('予約フォームの読み込みに失敗しました',
      'OrderForm.htmlファイルが見つかりません。');
  }
}

function handleDashboard(password) {
  if (password !== SYSTEM_CONFIG.adminPassword) {
    return createRedirectPage('認証失敗', '?');
  }
  
  try {
    return HtmlService.createHtmlOutputFromFile('Dashboard')
      .setTitle('Hyggelyカンパーニュ専門店 管理者ダッシュボード')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    return createErrorPage('ダッシュボードの読み込みに失敗しました', error.toString());
  }
}

function handleEmailSettings(password) {
  if (password !== SYSTEM_CONFIG.adminPassword) {
    return createRedirectPage('認証失敗', '?');
  }
  
  try {
    return HtmlService.createHtmlOutputFromFile('EmailSettings')
      .setTitle('Hyggelyカンパーニュ専門店 メール設定')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    return createErrorPage('メール設定の読み込みに失敗しました', error.toString());
  }
}

function handleHealthCheck() {
  const health = {
    status: 'healthy',
    version: SYSTEM_CONFIG.version,
    timestamp: new Date().toISOString(),
    spreadsheetId: SYSTEM_CONFIG.spreadsheetId
  };
  
  const html = `
    <!DOCTYPE html>
    <html><head><title>システム状態</title></head>
    <body style="font-family: Arial; padding: 20px;">
      <h1>✅ システム正常稼働中</h1>
      <pre>${JSON.stringify(health, null, 2)}</pre>
    </body></html>
  `;
  
  return HtmlService.createHtmlOutput(html);
}

// ===== エラーページ作成 =====
function createErrorPage(title, message) {
  const html = `
    <!DOCTYPE html>
    <html lang="ja">
    <head>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <title>エラー - Hyggelyカンパーニュ専門店</title>
      <style>
        body { font-family: Arial, sans-serif; text-align: center; padding: 50px;
               background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%); }
        .container { max-width: 600px; margin: 0 auto; background: white;
                     padding: 40px; border-radius: 15px; box-shadow: 0 8px 30px rgba(0,0,0,0.1); }
        .error-icon { font-size: 4rem; color: #dc3545; margin-bottom: 20px; }
        .btn { display: inline-block; padding: 12px 24px; background: #8B4513;
               color: white; text-decoration: none; border-radius: 8px; margin: 10px; }
      </style>
    </head>
    <body>
      <div class="container">
        <div class="error-icon">⚠️</div>
        <h1>${title}</h1>
        <p>${message}</p>
        <p>
          <a href="?" class="btn">🏠 予約フォームに戻る</a>
          <a href="javascript:location.reload()" class="btn">🔄 再読み込み</a>
        </p>
      </div>
    </body>
    </html>
  `;
  
  return HtmlService.createHtmlOutput(html);
}

function createRedirectPage(message, url) {
  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <title>リダイレクト中...</title>
    </head>
    <body style="text-align: center; padding: 100px; font-family: Arial;">
      <h2>${message}</h2>
      <p>3秒後にリダイレクトします...</p>
      <script>setTimeout(() => window.location.href = '${url}', 3000);</script>
    </body>
    </html>
  `;
  
  return HtmlService.createHtmlOutput(html);
}

// ===== システム初期化 =====
function checkAndInitializeSystem() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const sheets = spreadsheet.getSheets().map(s => s.getName());
    
    // 必要なシートが存在するかチェック
    Object.values(SYSTEM_CONFIG.sheets).forEach(sheetName => {
      if (!sheets.includes(sheetName)) {
        initializeSheet(spreadsheet, sheetName);
      }
    });
    
    console.log('✅ システム初期化完了');
  } catch (error) {
    console.error('❌ システム初期化エラー:', error);
    throw new Error('スプレッドシートにアクセスできません: ' + error.message);
  }
}

function initializeSheet(spreadsheet, sheetName) {
  try {
    let sheet = spreadsheet.insertSheet(sheetName);
    
    switch (sheetName) {
      case SYSTEM_CONFIG.sheets.PRODUCT_MASTER:
        initProductMaster(sheet);
        break;
      case SYSTEM_CONFIG.sheets.INVENTORY:
        initInventory(sheet);
        break;
      case SYSTEM_CONFIG.sheets.ORDER:
        initOrderSheet(sheet);
        break;
      case SYSTEM_CONFIG.sheets.EMAIL_SETTINGS:
        initEmailSettings(sheet);
        break;
      case SYSTEM_CONFIG.sheets.SYSTEM_LOG:
        initSystemLog(sheet);
        break;
    }
    
    console.log('✅ シート初期化完了:', sheetName);
  } catch (error) {
    console.error('❌ シート初期化エラー:', sheetName, error);
  }
}

function initProductMaster(sheet) {
  const headers = ['商品ID', '商品名', '価格', '表示順', '有効フラグ', '作成日', '更新日'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const products = getDefaultProducts();
  const data = products.map(p => [
    p.id, p.name, p.price, p.order, true, 
    new Date(), new Date()
  ]);
  
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  }
}

function initInventory(sheet) {
  const headers = ['商品ID', '商品名', '単価', '在庫数', '予約数', '残数', '最低在庫数', '更新日'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const products = getDefaultProducts();
  const data = products.map(p => [
    p.id, p.name, p.price, 10, 0, 10, 3, new Date()
  ]);
  
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  }
}

function initOrderSheet(sheet) {
  const basicHeaders = ['タイムスタンプ', '姓', '名', 'メール', '受取日', '受取時間'];
  const products = getDefaultProducts();
  const productHeaders = products.map(p => p.name);
  const finalHeaders = ['備考', '合計金額', '引渡済', '予約ID'];
  const allHeaders = [...basicHeaders, ...productHeaders, ...finalHeaders];
  
  sheet.getRange(1, 1, 1, allHeaders.length).setValues([allHeaders]);
}

function initEmailSettings(sheet) {
  const headers = ['設定項目', '設定値'];
  const data = [
    ['admin_email', 'hyggely2021@gmail.com'],
    ['email_enabled', 'TRUE'],
    ['customer_subject', 'Hyggelyカンパーニュ専門店 ご予約完了確認'],
    ['customer_body', '{lastName} {firstName} 様\n\nご予約ありがとうございます。\n{orderItems}\n\n合計: ¥{totalPrice}\n受取日時: {pickupDateTime}'],
    ['admin_subject', '【新規予約】{lastName} {firstName}様'],
    ['admin_body', '新規予約\n\nお客様: {lastName} {firstName}\nメール: {email}\n受取: {pickupDateTime}\n{orderItems}\n合計: ¥{totalPrice}']
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, data.length, headers.length).setValues(data);
}

function initSystemLog(sheet) {
  const headers = ['タイムスタンプ', 'レベル', 'イベント', '詳細', 'ユーザー'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}

// ===== デフォルト商品データ取得 =====
function getDefaultProducts() {
  return [
    {id: 'PRD001', name: 'プレミアムカンパーニュ', price: 1000, order: 1},
    {id: 'PRD002', name: 'プレミアムカンパーニュ 1/2', price: 600, order: 2},
    {id: 'PRD003', name: 'レーズン&クルミ', price: 1200, order: 3},
    {id: 'PRD004', name: 'レーズン&クルミ 1/2', price: 600, order: 4},
    {id: 'PRD005', name: 'いちじく&クルミ', price: 400, order: 5},
    {id: 'PRD006', name: '4種のMIXナッツ', price: 400, order: 6},
    {id: 'PRD007', name: 'MIXドライフルーツ', price: 400, order: 7},
    {id: 'PRD008', name: 'アールグレイ', price: 350, order: 8},
    {id: 'PRD009', name: 'チョコレート', price: 450, order: 9},
    {id: 'PRD010', name: 'チーズ', price: 450, order: 10},
    {id: 'PRD011', name: 'ひまわりの種', price: 400, order: 11},
    {id: 'PRD012', name: 'デーツ', price: 400, order: 12},
    {id: 'PRD013', name: 'カレーパン', price: 450, order: 13},
    {id: 'PRD014', name: 'バターロール', price: 230, order: 14},
    {id: 'PRD015', name: 'ショコラロール', price: 280, order: 15},
    {id: 'PRD016', name: '自家製クリームパン', price: 350, order: 16},
    {id: 'PRD017', name: '自家製あんバター', price: 380, order: 17},
    {id: 'PRD018', name: '抹茶&ホワイトチョコ', price: 400, order: 18},
    {id: 'PRD019', name: '黒ごまパン', price: 200, order: 19},
    {id: 'PRD020', name: 'レーズンジャムとクリームチーズのパン', price: 350, order: 20},
    {id: 'PRD021', name: 'ピーナッツクリームパン', price: 350, order: 21},
    {id: 'PRD022', name: 'あん食パン', price: 400, order: 22},
    {id: 'PRD023', name: 'コーンパン', price: 400, order: 23},
    {id: 'PRD024', name: 'レモンとクリームチーズのミニ食パン', price: 450, order: 24},
    {id: 'PRD025', name: 'ピザ マルゲリータ', price: 1100, order: 25},
    {id: 'PRD026', name: 'ピタパンサンド', price: 800, order: 26},
    {id: 'PRD027', name: 'フォカッチャ', price: 300, order: 27}
  ];
}

// ===== 商品マスタ管理 =====
function getProductMaster() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const masterSheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.PRODUCT_MASTER);
    
    if (!masterSheet || masterSheet.getLastRow() <= 1) {
      return getDefaultProducts().map(p => ({ ...p, enabled: true }));
    }
    
    const data = masterSheet.getDataRange().getValues();
    const products = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      products.push({
        id: row[0],
        name: row[1],
        price: row[2],
        order: row[3],
        enabled: row[4] !== false
      });
    }
    
    return products.sort((a, b) => a.order - b.order);
  } catch (error) {
    console.error('❌ 商品マスタ取得エラー:', error);
    return getDefaultProducts().map(p => ({ ...p, enabled: true }));
  }
}

function addProduct(productData) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const masterSheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.PRODUCT_MASTER);
    const inventorySheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.INVENTORY);
    
    // 新しい商品IDを生成
    const products = getProductMaster();
    const maxId = Math.max(...products.map(p => parseInt(p.id.substring(3)) || 0));
    const newId = 'PRD' + String(maxId + 1).padStart(3, '0');
    
    const now = new Date();
    
    // 商品マスタに追加
    const masterRow = [
      newId,
      productData.name,
      productData.price,
      productData.order || (maxId + 1),
      true,
      now,
      now
    ];
    
    masterSheet.appendRow(masterRow);
    
    // 在庫シートに追加
    const inventoryRow = [
      newId,
      productData.name,
      productData.price,
      productData.initialStock || 10,
      0,
      productData.initialStock || 10,
      productData.minStock || 3,
      now
    ];
    
    inventorySheet.appendRow(inventoryRow);
    
    // 注文シートのヘッダーを更新
    updateOrderSheetHeaders();
    
    logSystemEvent('INFO', '商品追加', `新商品追加: ${productData.name} (${newId})`);
    
    return {
      success: true,
      message: '商品を追加しました',
      productId: newId
    };
    
  } catch (error) {
    console.error('❌ 商品追加エラー:', error);
    logSystemEvent('ERROR', '商品追加エラー', error.toString());
    return {
      success: false,
      message: '商品の追加に失敗しました: ' + error.message
    };
  }
}

function updateProduct(productId, productData) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const masterSheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.PRODUCT_MASTER);
    const inventorySheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.INVENTORY);
    
    const now = new Date();
    
    // 商品マスタを更新
    const masterData = masterSheet.getDataRange().getValues();
    for (let i = 1; i < masterData.length; i++) {
      if (masterData[i][0] === productId) {
        masterSheet.getRange(i + 1, 2).setValue(productData.name);
        masterSheet.getRange(i + 1, 3).setValue(productData.price);
        masterSheet.getRange(i + 1, 4).setValue(productData.order);
        masterSheet.getRange(i + 1, 5).setValue(productData.enabled);
        masterSheet.getRange(i + 1, 7).setValue(now);
        break;
      }
    }
    
    // 在庫シートを更新
    const inventoryData = inventorySheet.getDataRange().getValues();
    for (let i = 1; i < inventoryData.length; i++) {
      if (inventoryData[i][0] === productId) {
        inventorySheet.getRange(i + 1, 2).setValue(productData.name);
        inventorySheet.getRange(i + 1, 3).setValue(productData.price);
        inventorySheet.getRange(i + 1, 8).setValue(now);
        break;
      }
    }
    
    // 注文シートのヘッダーを更新
    updateOrderSheetHeaders();
    
    logSystemEvent('INFO', '商品更新', `商品更新: ${productData.name} (${productId})`);
    
    return {
      success: true,
      message: '商品を更新しました'
    };
    
  } catch (error) {
    console.error('❌ 商品更新エラー:', error);
    logSystemEvent('ERROR', '商品更新エラー', error.toString());
    return {
      success: false,
      message: '商品の更新に失敗しました: ' + error.message
    };
  }
}

function deleteProduct(productId) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const masterSheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.PRODUCT_MASTER);
    const inventorySheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.INVENTORY);
    
    // 商品マスタから削除
    const masterData = masterSheet.getDataRange().getValues();
    for (let i = 1; i < masterData.length; i++) {
      if (masterData[i][0] === productId) {
        masterSheet.deleteRow(i + 1);
        break;
      }
    }
    
    // 在庫シートから削除
    const inventoryData = inventorySheet.getDataRange().getValues();
    for (let i = 1; i < inventoryData.length; i++) {
      if (inventoryData[i][0] === productId) {
        inventorySheet.deleteRow(i + 1);
        break;
      }
    }
    
    // 注文シートのヘッダーを更新
    updateOrderSheetHeaders();
    
    logSystemEvent('INFO', '商品削除', `商品削除: ${productId}`);
    
    return {
      success: true,
      message: '商品を削除しました'
    };
    
  } catch (error) {
    console.error('❌ 商品削除エラー:', error);
    logSystemEvent('ERROR', '商品削除エラー', error.toString());
    return {
      success: false,
      message: '商品の削除に失敗しました: ' + error.message
    };
  }
}

function updateOrderSheetHeaders() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const orderSheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.ORDER);
    const products = getProductMaster().filter(p => p.enabled);
    
    const basicHeaders = ['タイムスタンプ', '姓', '名', 'メール', '受取日', '受取時間'];
    const productHeaders = products.map(p => p.name);
    const finalHeaders = ['備考', '合計金額', '引渡済', '予約ID'];
    const allHeaders = [...basicHeaders, ...productHeaders, ...finalHeaders];
    
    // 既存データを保護しながらヘッダーのみ更新
    const currentData = orderSheet.getDataRange().getValues();
    if (currentData.length > 0) {
      orderSheet.getRange(1, 1, 1, allHeaders.length).setValues([allHeaders]);
    }
    
    console.log('✅ 注文シートヘッダー更新完了');
  } catch (error) {
    console.error('❌ 注文シートヘッダー更新エラー:', error);
  }
}

// ===== データ取得関数 =====
function getInventoryDataForForm() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const inventorySheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.INVENTORY);
    const products = getProductMaster().filter(p => p.enabled);
    
    if (!inventorySheet || inventorySheet.getLastRow() <= 1) {
      return products.map(p => ({
        id: p.id,
        name: p.name,
        price: p.price,
        stock: 10,
        reserved: 0,
        remaining: 10,
        minStock: 3
      }));
    }
    
    updateInventoryFromOrders();
    
    const data = inventorySheet.getDataRange().getValues();
    const inventory = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const product = products.find(p => p.id === row[0]);
      
      if (product) {
        inventory.push({
          id: row[0],
          name: row[1],
          price: row[2],
          stock: row[3] || 0,
          reserved: row[4] || 0,
          remaining: Math.max(0, (row[3] || 0) - (row[4] || 0)),
          minStock: row[6] || 3
        });
      }
    }
    
    return inventory.sort((a, b) => {
      const aProduct = products.find(p => p.id === a.id);
      const bProduct = products.find(p => p.id === b.id);
      return (aProduct?.order || 0) - (bProduct?.order || 0);
    });
    
  } catch (error) {
    console.error('❌ 在庫データ取得エラー:', error);
    const products = getProductMaster().filter(p => p.enabled);
    return products.map(p => ({
      id: p.id,
      name: p.name,
      price: p.price,
      stock: 10,
      reserved: 0,
      remaining: 10,
      minStock: 3
    }));
  }
}

function getOrderList() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const orderSheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.ORDER);
    
    if (!orderSheet || orderSheet.getLastRow() <= 1) {
      return [];
    }
    
    const data = orderSheet.getDataRange().getValues();
    const products = getProductMaster();
    const orders = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // pickupDateを必ずyyyy-MM-dd形式に整形
      let pickupDate = row[4];
      if (pickupDate instanceof Date) {
        pickupDate = Utilities.formatDate(pickupDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else if (typeof pickupDate === 'string' && pickupDate.match(/^\d{4}\/\d{1,2}\/\d{1,2}$/)) {
        // 例: 2024/06/01 → 2024-06-01
        const parts = pickupDate.split('/');
        pickupDate = `${parts[0]}-${('0'+parts[1]).slice(-2)}-${('0'+parts[2]).slice(-2)}`;
      }

      const order = {
        rowIndex: i + 1,
        timestamp: row[0],
        lastName: row[1],
        firstName: row[2],
        email: row[3],
        pickupDate: pickupDate,
        pickupTime: row[5],
        items: [],
        note: row[row.length - 4] || '',
        totalPrice: row[row.length - 3] || 0,
        isDelivered: row[row.length - 2] === '引渡済',
        orderId: row[row.length - 1] || generateOrderId()
      };
      
      // 商品データを解析
      for (let j = 6; j < 6 + products.length; j++) {
        if (j < row.length && row[j] > 0) {
          const productIndex = j - 6;
          if (productIndex < products.length) {
            order.items.push({
              productId: products[productIndex].id,
              name: products[productIndex].name,
              quantity: row[j],
              price: products[productIndex].price,
              subtotal: row[j] * products[productIndex].price
            });
          }
        }
      }
      
      orders.push(order);
    }
    
    return orders.reverse(); // 新しい順
  } catch (error) {
    console.error('❌ 予約一覧取得エラー:', error);
    return [];
  }
}

function getOrderDetails(orderId) {
  try {
    const orders = getOrderList();
    const order = orders.find(o => o.orderId === orderId);
    
    if (!order) {
      return { success: false, message: '予約が見つかりません' };
    }
    
    return { success: true, order: order };
  } catch (error) {
    console.error('❌ 予約詳細取得エラー:', error);
    return { success: false, message: '予約詳細の取得に失敗しました' };
  }
}

function getDashboardStats() {
  try {
    const orders = getOrderList();
    const inventory = getInventoryDataForForm();
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    const todayOrders = orders.filter(order => {
      try {
        const pickupDate = new Date(order.pickupDate);
        pickupDate.setHours(0, 0, 0, 0);
        return pickupDate.getTime() === today.getTime();
      } catch (e) {
        return false;
      }
    });
    
    const pendingOrders = orders.filter(order => !order.isDelivered);
    const outOfStock = inventory.filter(p => p.remaining <= 0);
    const lowStock = inventory.filter(p => p.remaining > 0 && p.remaining <= p.minStock);
    
    return {
      todayOrdersCount: todayOrders.length,
      pendingOrdersCount: pendingOrders.length,
      outOfStockCount: outOfStock.length,
      lowStockCount: lowStock.length,
      totalProducts: inventory.length,
      todayRevenue: todayOrders.reduce((sum, order) => sum + (order.totalPrice || 0), 0),
      systemVersion: SYSTEM_CONFIG.version,
      lastUpdate: new Date().toISOString()
    };
  } catch (error) {
    console.error('❌ 統計データ取得エラー:', error);
    return {
      todayOrdersCount: 0,
      pendingOrdersCount: 0,
      outOfStockCount: 0,
      lowStockCount: 0,
      totalProducts: 0,
      todayRevenue: 0,
      systemVersion: SYSTEM_CONFIG.version,
      lastUpdate: new Date().toISOString()
    };
  }
}

// ===== 予約処理 =====
function processOrder(formData) {
  try {
    console.log('🔄 予約処理開始:', JSON.stringify(formData, null, 2));
    
    // バリデーション
    if (!formData.lastName || !formData.firstName || !formData.email ||
        !formData.pickupDate || !formData.pickupTime) {
      throw new Error('必須項目が入力されていません');
    }
    
    // 在庫チェック
    const inventory = getInventoryDataForForm();
    const products = getProductMaster().filter(p => p.enabled);
    const orderedItems = [];
    let totalPrice = 0;
    
    for (let i = 0; i < products.length; i++) {
      const quantity = parseInt(formData[`product_${i}`]) || 0;
      if (quantity > 0) {
        const product = products[i];
        const inventoryItem = inventory.find(inv => inv.id === product.id);
        
        if (!inventoryItem || inventoryItem.remaining < quantity) {
          throw new Error(`${product.name}の在庫が不足しています`);
        }
        
        orderedItems.push({
          productId: product.id,
          name: product.name,
          quantity: quantity,
          price: product.price,
          subtotal: quantity * product.price
        });
        totalPrice += quantity * product.price;
      }
    }
    
    if (orderedItems.length === 0) {
      throw new Error('商品を1つ以上選択してください');
    }
    
    // 予約IDを生成
    const orderId = generateOrderId();
    
    // スプレッドシートに記録
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const orderSheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.ORDER);
    const lastRow = orderSheet.getLastRow() + 1;
    
    // 基本情報を記録
    const currentDate = new Date();
    orderSheet.getRange(lastRow, 1).setValue(currentDate);
    orderSheet.getRange(lastRow, 2).setValue(formData.lastName);
    orderSheet.getRange(lastRow, 3).setValue(formData.firstName);
    orderSheet.getRange(lastRow, 4).setValue(formData.email);
    orderSheet.getRange(lastRow, 5).setValue(formData.pickupDate);
    orderSheet.getRange(lastRow, 6).setValue(formData.pickupTime);
    
    // 商品数量を記録
    for (let i = 0; i < products.length; i++) {
      const quantity = parseInt(formData[`product_${i}`]) || 0;
      orderSheet.getRange(lastRow, 7 + i).setValue(quantity);
    }
    
    // 備考、合計金額、予約IDを記録
    const noteCol = 7 + products.length;
    const totalCol = noteCol + 1;
    const deliveredCol = totalCol + 1;
    const orderIdCol = deliveredCol + 1;
    
    orderSheet.getRange(lastRow, noteCol).setValue(formData.note || '');
    orderSheet.getRange(lastRow, totalCol).setValue(totalPrice);
    orderSheet.getRange(lastRow, orderIdCol).setValue(orderId);
    
    // 在庫を更新
    updateInventoryFromOrders();
    
    // メール送信（修正版）
    console.log('📧 メール送信処理開始');
    const emailResults = sendOrderEmails(formData, orderedItems, totalPrice, orderId);
    console.log('📧 メール送信結果:', emailResults);
    
    // ログ記録
    logSystemEvent('INFO', '新規予約',
      `顧客: ${formData.lastName} ${formData.firstName}, 金額: ¥${totalPrice}, 予約ID: ${orderId}, メール送信: ${emailResults.success ? '成功' : '失敗'}`);
    
    return {
      success: true,
      message: '予約が完了しました',
      orderDetails: {
        orderId: orderId,
        name: `${formData.lastName} ${formData.firstName}`,
        email: formData.email,
        pickupDate: formData.pickupDate,
        pickupTime: formData.pickupTime,
        items: orderedItems,
        totalPrice: totalPrice
      },
      emailResult: emailResults
    };
    
  } catch (error) {
    console.error('❌ 予約処理エラー:', error);
    logSystemEvent('ERROR', '予約処理エラー', error.toString());
    return {
      success: false,
      message: error.message
    };
  }
}

// ===== 在庫更新 =====
function updateInventoryFromOrders() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const orderSheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.ORDER);
    const inventorySheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.INVENTORY);
    
    if (!orderSheet || !inventorySheet || orderSheet.getLastRow() <= 1) {
      return;
    }
    
    const products = getProductMaster();
    
    // 予約数を集計
    const reservations = {};
    products.forEach(p => reservations[p.id] = 0);
    
    const orderData = orderSheet.getDataRange().getValues();
    for (let i = 1; i < orderData.length; i++) {
      const row = orderData[i];
      const isDelivered = row[row.length - 2] === '引渡済';
      
      if (!isDelivered) {
        for (let j = 0; j < products.length; j++) {
          const quantity = parseInt(row[6 + j]) || 0;
          if (quantity > 0 && products[j]) {
            reservations[products[j].id] += quantity;
          }
        }
      }
    }
    
    // 在庫シートを更新
    const inventoryData = inventorySheet.getDataRange().getValues();
    for (let i = 1; i < inventoryData.length; i++) {
      const productId = inventoryData[i][0];
      if (reservations[productId] !== undefined) {
        inventorySheet.getRange(i + 1, 5).setValue(reservations[productId]); // 予約数
        
        const stock = inventoryData[i][3] || 0;
        const remaining = Math.max(0, stock - reservations[productId]);
        inventorySheet.getRange(i + 1, 6).setValue(remaining); // 残数
        inventorySheet.getRange(i + 1, 8).setValue(new Date()); // 更新日
      }
    }
    
  } catch (error) {
    console.error('❌ 在庫更新エラー:', error);
  }
}

function updateInventoryStock(productId, newStock, updatedBy) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const inventorySheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.INVENTORY);
    const data = inventorySheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === productId) {
        inventorySheet.getRange(i + 1, 4).setValue(newStock);
        
        const reserved = data[i][4] || 0;
        const remaining = Math.max(0, newStock - reserved);
        inventorySheet.getRange(i + 1, 6).setValue(remaining);
        inventorySheet.getRange(i + 1, 8).setValue(new Date());
        
        logSystemEvent('INFO', '在庫更新',
          `商品: ${data[i][1]}, 新在庫: ${newStock}, 更新者: ${updatedBy}`);
        
        return { success: true, message: '在庫数を更新しました' };
      }
    }
    
    return { success: false, message: '商品が見つかりません' };
  } catch (error) {
    console.error('❌ 在庫数更新エラー:', error);
    logSystemEvent('ERROR', '在庫更新エラー', error.toString());
    return { success: false, message: '更新に失敗しました: ' + error.message };
  }
}

function bulkUpdateInventory(updates) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const inventorySheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.INVENTORY);
    const data = inventorySheet.getDataRange().getValues();
    
    let updateCount = 0;
    
    updates.forEach(update => {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === update.productId) {
          inventorySheet.getRange(i + 1, 4).setValue(update.newStock);
          
          const reserved = data[i][4] || 0;
          const remaining = Math.max(0, update.newStock - reserved);
          inventorySheet.getRange(i + 1, 6).setValue(remaining);
          inventorySheet.getRange(i + 1, 8).setValue(new Date());
          
          updateCount++;
          break;
        }
      }
    });
    
    logSystemEvent('INFO', '一括在庫更新', `${updateCount}商品の在庫を更新`);
    
    return { 
      success: true, 
      message: `${updateCount}商品の在庫を更新しました`,
      updateCount: updateCount
    };
    
  } catch (error) {
    console.error('❌ 一括在庫更新エラー:', error);
    logSystemEvent('ERROR', '一括在庫更新エラー', error.toString());
    return { success: false, message: '一括更新に失敗しました: ' + error.message };
  }
}

function updateDeliveryStatus(rowIndex, isDelivered) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const orderSheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.ORDER);
    const products = getProductMaster();
    
    const deliveredCol = 7 + products.length + 2; // 引渡済列
    const statusValue = isDelivered ? '引渡済' : '';
    
    orderSheet.getRange(rowIndex, deliveredCol).setValue(statusValue);
    
    updateInventoryFromOrders();
    
    // ログ記録
    const row = orderSheet.getRange(rowIndex, 1, 1, orderSheet.getLastColumn()).getValues()[0];
    const customerName = `${row[1]} ${row[2]}`;
    logSystemEvent('INFO', '引渡状態変更',
      `顧客: ${customerName}, 状態: ${isDelivered ? '引渡済' : '未引渡'}`);
    
    return {
      success: true,
      message: isDelivered ? '引渡完了にしました' : '引渡待ちに戻しました'
    };
  } catch (error) {
    console.error('❌ 引渡状態更新エラー:', error);
    logSystemEvent('ERROR', '引渡状態更新エラー', error.toString());
    return {
      success: false,
      message: '更新に失敗しました: ' + error.message
    };
  }
}

// ===== 予約編集・キャンセル =====
function updateOrder(orderId, updateData) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const orderSheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.ORDER);
    const data = orderSheet.getDataRange().getValues();
    
    // 予約を検索
    let orderRowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][data[i].length - 1] === orderId) {
        orderRowIndex = i + 1;
        break;
      }
    }
    
    if (orderRowIndex === -1) {
      return { success: false, message: '予約が見つかりません' };
    }
    
    // 基本情報を更新
    if (updateData.lastName) orderSheet.getRange(orderRowIndex, 2).setValue(updateData.lastName);
    if (updateData.firstName) orderSheet.getRange(orderRowIndex, 3).setValue(updateData.firstName);
    if (updateData.email) orderSheet.getRange(orderRowIndex, 4).setValue(updateData.email);
    if (updateData.pickupDate) orderSheet.getRange(orderRowIndex, 5).setValue(updateData.pickupDate);
    if (updateData.pickupTime) orderSheet.getRange(orderRowIndex, 6).setValue(updateData.pickupTime);
    if (updateData.note !== undefined) {
      const products = getProductMaster();
      const noteCol = 7 + products.length;
      orderSheet.getRange(orderRowIndex, noteCol).setValue(updateData.note);
    }
    
    updateInventoryFromOrders();
    
    logSystemEvent('INFO', '予約更新', `予約ID: ${orderId}, 更新内容: ${JSON.stringify(updateData)}`);
    
    return { success: true, message: '予約を更新しました' };
    
  } catch (error) {
    console.error('❌ 予約更新エラー:', error);
    logSystemEvent('ERROR', '予約更新エラー', error.toString());
    return { success: false, message: '予約の更新に失敗しました: ' + error.message };
  }
}

function cancelOrder(orderId) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const orderSheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.ORDER);
    const data = orderSheet.getDataRange().getValues();
    
    // 予約を検索して削除
    for (let i = 1; i < data.length; i++) {
      if (data[i][data[i].length - 1] === orderId) {
        const customerName = `${data[i][1]} ${data[i][2]}`;
        orderSheet.deleteRow(i + 1);
        
        updateInventoryFromOrders();
        
        logSystemEvent('INFO', '予約キャンセル', `顧客: ${customerName}, 予約ID: ${orderId}`);
        
        return { success: true, message: '予約をキャンセルしました' };
      }
    }
    
    return { success: false, message: '予約が見つかりません' };
    
  } catch (error) {
    console.error('❌ 予約キャンセルエラー:', error);
    logSystemEvent('ERROR', '予約キャンセルエラー', error.toString());
    return { success: false, message: '予約のキャンセルに失敗しました: ' + error.message };
  }
}

// ===== メール機能（完全修正版） =====

/**
 * メール設定を確実に取得する関数（修正版）
 */
function getEmailSettings() {
  try {
    console.log('📧 メール設定取得開始');
    
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const emailSheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.EMAIL_SETTINGS);
    
    // デフォルト設定
    const defaultSettings = {
      adminEmail: 'hyggely2021@gmail.com',
      emailEnabled: true,
      customerSubject: 'Hyggelyカンパーニュ専門店 ご予約完了確認',
      customerBody: '{lastName} {firstName} 様\n\nHyggely事前予約システムをご利用いただき誠にありがとうございます。\n以下の注文内容で承りましたのでお知らせします。\n\n予約ID: {orderId}\n\n{orderItems}\n\n・合計　　　¥{totalPrice}\n\n・受取日時：{pickupDateTime}\n\n受取日当日に現金またはPayPayでお支払いいただきます。\n当日は気をつけてお越しください。\n\n※このメールは自動送信されています。返信はできませんのでご了承ください。\nご不明な点がございましたら、Hyggely公式LINEまでお問い合わせください。',
      adminSubject: '【新規予約】{lastName} {firstName}様',
      adminBody: '【新規予約通知】\n\n予約ID: {orderId}\n\nお客様情報:\n・氏名: {lastName} {firstName} 様\n・メール: {email}\n・受取日時: {pickupDateTime}\n\n注文内容:\n{orderItems}\n\n合計金額: ¥{totalPrice}\n\n※予約管理システムより自動送信されています。'
    };
    
    if (!emailSheet || emailSheet.getLastRow() <= 1) {
      console.log('⚠️ メール設定シートが存在しないか空です。デフォルト設定を使用します。');
      return defaultSettings;
    }
    
    // スプレッドシートから設定を読み込み
    const data = emailSheet.getDataRange().getValues();
    const settings = { ...defaultSettings }; // デフォルト設定をベースにする
    
    console.log('📧 メール設定シートデータ:', data);
    
    for (let i = 1; i < data.length; i++) {
      const key = data[i][0];
      const value = data[i][1];
      
      if (!key) continue; // 空のキーはスキップ
      
      switch (key.toLowerCase()) {
        case 'admin_email':
          if (value && typeof value === 'string') {
            settings.adminEmail = value.trim();
          }
          break;
        case 'email_enabled':
          settings.emailEnabled = (value === 'TRUE' || value === true || value === 1);
          break;
        case 'customer_subject':
          if (value && typeof value === 'string') {
            settings.customerSubject = value;
          }
          break;
        case 'customer_body':
          if (value && typeof value === 'string') {
            settings.customerBody = value;
          }
          break;
        case 'admin_subject':
          if (value && typeof value === 'string') {
            settings.adminSubject = value;
          }
          break;
        case 'admin_body':
          if (value && typeof value === 'string') {
            settings.adminBody = value;
          }
          break;
      }
    }
    
    console.log('✅ メール設定取得完了:', settings);
    return settings;
    
  } catch (error) {
    console.error('❌ メール設定取得エラー:', error);
    logSystemEvent('ERROR', 'メール設定取得エラー', error.toString());
    
    // エラー時はデフォルト設定を返す
    return {
      adminEmail: 'hyggely2021@gmail.com',
      emailEnabled: true,
      customerSubject: 'Hyggelyカンパーニュ専門店 ご予約完了確認',
      customerBody: '{lastName} {firstName} 様\n\nご予約ありがとうございます。',
      adminSubject: '【新規予約】{lastName} {firstName}様',
      adminBody: '新規予約通知'
    };
  }
}

/**
 * メール送信の統合関数（修正版）
 */
function sendOrderEmails(formData, orderedItems, totalPrice, orderId) {
  console.log('📧 メール送信統合処理開始');
  
  const results = {
    success: false,
    customerEmailSent: false,
    adminEmailSent: false,
    errors: []
  };
  
  try {
    // メール設定を取得
    const settings = getEmailSettings();
    console.log('📧 取得した設定:', settings);
    
    if (!settings.emailEnabled) {
      console.log('⚠️ メール送信が無効に設定されています');
      results.errors.push('メール送信が無効に設定されています');
      return results;
    }
    
    // Gmail権限チェック
    if (!checkGmailPermission()) {
      const errorMsg = 'Gmail送信権限が不足しています。Google Apps Scriptでの権限承認が必要です。';
      console.error('❌ ' + errorMsg);
      results.errors.push(errorMsg);
      logSystemEvent('ERROR', 'Gmail権限エラー', errorMsg);
      return results;
    }
    
    // 顧客確認メール送信
    try {
      const customerResult = sendConfirmationEmail(formData, orderedItems, totalPrice, orderId, settings);
      results.customerEmailSent = customerResult.success;
      if (!customerResult.success) {
        results.errors.push('顧客メール送信失敗: ' + customerResult.error);
      }
    } catch (customerError) {
      console.error('❌ 顧客メール送信エラー:', customerError);
      results.errors.push('顧客メール送信エラー: ' + customerError.toString());
    }
    
    // 管理者通知メール送信
    try {
      const adminResult = sendAdminNotification(formData, orderedItems, totalPrice, orderId, settings);
      results.adminEmailSent = adminResult.success;
      if (!adminResult.success) {
        results.errors.push('管理者メール送信失敗: ' + adminResult.error);
      }
    } catch (adminError) {
      console.error('❌ 管理者メール送信エラー:', adminError);
      results.errors.push('管理者メール送信エラー: ' + adminError.toString());
    }
    
    // 成功判定
    results.success = results.customerEmailSent || results.adminEmailSent;
    
    console.log('📧 メール送信結果:', results);
    
    // ログ記録
    const logDetail = `顧客メール: ${results.customerEmailSent ? '成功' : '失敗'}, 管理者メール: ${results.adminEmailSent ? '成功' : '失敗'}`;
    if (results.errors.length > 0) {
      logDetail += `, エラー: ${results.errors.join(', ')}`;
    }
    logSystemEvent(results.success ? 'INFO' : 'ERROR', 'メール送信結果', logDetail);
    
    return results;
    
  } catch (error) {
    console.error('❌ メール送信統合処理エラー:', error);
    results.errors.push('統合処理エラー: ' + error.toString());
    logSystemEvent('ERROR', 'メール送信統合エラー', error.toString());
    return results;
  }
}

/**
 * Gmail権限チェック関数
 */
function checkGmailPermission() {
  try {
    // テスト用のメール送信（実際には送信しない）
    GmailApp.getInboxThreads(0, 1);
    return true;
  } catch (error) {
    console.error('❌ Gmail権限チェック失敗:', error);
    return false;
  }
}

/**
 * 顧客確認メール送信（修正版）
 */
function sendConfirmationEmail(formData, orderedItems, totalPrice, orderId, settings) {
  try {
    console.log('📧 顧客確認メール送信開始');
    
    if (!formData.email || !formData.email.includes('@')) {
      return { success: false, error: '有効なメールアドレスがありません' };
    }
    
    const pickupDateTime = `${formData.pickupDate} ${formData.pickupTime}`;
    const itemsText = orderedItems.map(item =>
      `・${item.name}　${item.quantity}個　¥${item.subtotal.toLocaleString()}`
    ).join('\n');
    
    // 変数置換
    const subject = replaceEmailVariables(settings.customerSubject, {
      lastName: formData.lastName,
      firstName: formData.firstName,
      orderId: orderId
    });
    
    const body = replaceEmailVariables(settings.customerBody, {
      lastName: formData.lastName,
      firstName: formData.firstName,
      orderItems: itemsText,
      totalPrice: totalPrice.toLocaleString(),
      pickupDateTime: pickupDateTime,
      orderId: orderId,
      email: formData.email
    });
    
    console.log('📧 顧客メール内容:', { to: formData.email, subject, body });
    
    // メール送信（リトライ機能付き）
    let attempt = 0;
    const maxAttempts = 3;
    
    while (attempt < maxAttempts) {
      try {
        attempt++;
        console.log(`📧 顧客メール送信試行 ${attempt}/${maxAttempts}`);
        
        GmailApp.sendEmail(formData.email, subject, body);
        
        console.log('✅ 顧客確認メール送信成功');
        logSystemEvent('INFO', '顧客メール送信成功', `宛先: ${formData.email}, 予約ID: ${orderId}`);
        return { success: true };
        
      } catch (sendError) {
        console.error(`❌ 顧客メール送信試行 ${attempt} 失敗:`, sendError);
        
        if (attempt >= maxAttempts) {
          logSystemEvent('ERROR', '顧客メール送信失敗', `宛先: ${formData.email}, エラー: ${sendError.toString()}`);
          return { success: false, error: sendError.toString() };
        }
        
        // 1秒待ってリトライ
        Utilities.sleep(1000);
      }
    }
    
  } catch (error) {
    console.error('❌ 顧客メール送信エラー:', error);
    logSystemEvent('ERROR', '顧客メール送信エラー', error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * 管理者通知メール送信（修正版）
 */
function sendAdminNotification(formData, orderedItems, totalPrice, orderId, settings) {
  try {
    console.log('📧 管理者通知メール送信開始');
    
    if (!settings.adminEmail || !settings.adminEmail.includes('@')) {
      return { success: false, error: '管理者メールアドレスが設定されていません' };
    }
    
    const pickupDateTime = `${formData.pickupDate} ${formData.pickupTime}`;
    const itemsText = orderedItems.map(item =>
      `・${item.name}　${item.quantity}個　¥${item.subtotal.toLocaleString()}`
    ).join('\n');
    
    // 変数置換
    const subject = replaceEmailVariables(settings.adminSubject, {
      lastName: formData.lastName,
      firstName: formData.firstName,
      orderId: orderId,
      totalPrice: totalPrice.toLocaleString()
    });
    
    const body = replaceEmailVariables(settings.adminBody, {
      lastName: formData.lastName,
      firstName: formData.firstName,
      email: formData.email,
      orderItems: itemsText,
      totalPrice: totalPrice.toLocaleString(),
      pickupDateTime: pickupDateTime,
      orderId: orderId
    });
    
    console.log('📧 管理者メール内容:', { to: settings.adminEmail, subject, body });
    
    // メール送信（リトライ機能付き）
    let attempt = 0;
    const maxAttempts = 3;
    
    while (attempt < maxAttempts) {
      try {
        attempt++;
        console.log(`📧 管理者メール送信試行 ${attempt}/${maxAttempts}`);
        
        GmailApp.sendEmail(settings.adminEmail, subject, body);
        
        console.log('✅ 管理者通知メール送信成功');
        logSystemEvent('INFO', '管理者メール送信成功', `宛先: ${settings.adminEmail}, 予約ID: ${orderId}`);
        return { success: true };
        
      } catch (sendError) {
        console.error(`❌ 管理者メール送信試行 ${attempt} 失敗:`, sendError);
        
        if (attempt >= maxAttempts) {
          logSystemEvent('ERROR', '管理者メール送信失敗', `宛先: ${settings.adminEmail}, エラー: ${sendError.toString()}`);
          return { success: false, error: sendError.toString() };
        }
        
        // 1秒待ってリトライ
        Utilities.sleep(1000);
      }
    }
    
  } catch (error) {
    console.error('❌ 管理者メール送信エラー:', error);
    logSystemEvent('ERROR', '管理者メール送信エラー', error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * メールテンプレートの変数置換関数
 */
function replaceEmailVariables(template, variables) {
  let result = template;
  
  Object.keys(variables).forEach(key => {
    const regex = new RegExp(`{${key}}`, 'g');
    result = result.replace(regex, variables[key] || '');
  });
  
  return result;
}

/**
 * メール設定更新関数（修正版）
 */
function updateEmailSettings(newSettings) {
  try {
    console.log('📧 メール設定更新開始:', newSettings);
    
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const emailSheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.EMAIL_SETTINGS);
    
    const data = [
      ['admin_email', newSettings.adminEmail || 'hyggely2021@gmail.com'],
      ['email_enabled', newSettings.emailEnabled ? 'TRUE' : 'FALSE'],
      ['customer_subject', newSettings.customerSubject || 'Hyggelyカンパーニュ専門店 ご予約完了確認'],
      ['customer_body', newSettings.customerBody || ''],
      ['admin_subject', newSettings.adminSubject || '【新規予約】{lastName} {firstName}様'],
      ['admin_body', newSettings.adminBody || '']
    ];
    
    emailSheet.clear();
    emailSheet.getRange(1, 1, 1, 2).setValues([['設定項目', '設定値']]);
    emailSheet.getRange(2, 1, data.length, 2).setValues(data);
    
    logSystemEvent('INFO', 'メール設定更新', 'メール設定が更新されました');
    console.log('✅ メール設定更新完了');
    
    return { success: true, message: 'メール設定を保存しました' };
  } catch (error) {
    console.error('❌ メール設定更新エラー:', error);
    logSystemEvent('ERROR', 'メール設定更新エラー', error.toString());
    return { success: false, message: '保存に失敗しました: ' + error.message };
  }
}

/**
 * メール送信テスト関数
 */
function testEmailSending() {
  try {
    console.log('📧 メール送信テスト開始');
    
    // テスト用データ
    const testFormData = {
      lastName: 'テスト',
      firstName: '太郎',
      email: 'test@example.com',
      pickupDate: '2024-12-25',
      pickupTime: '14:00',
      note: 'テスト注文です'
    };
    
    const testItems = [
      { name: 'プレミアムカンパーニュ', quantity: 1, price: 1000, subtotal: 1000 }
    ];
    
    const testOrderId = 'TEST' + Date.now();
    
    const result = sendOrderEmails(testFormData, testItems, 1000, testOrderId);
    
    console.log('📧 テスト結果:', result);
    return result;
    
  } catch (error) {
    console.error('❌ メール送信テストエラー:', error);
    return { success: false, error: error.toString() };
  }
}

// ===== ユーティリティ関数 =====
function generateOrderId() {
  const now = new Date();
  const dateStr = now.getFullYear().toString().slice(-2) +
                 String(now.getMonth() + 1).padStart(2, '0') +
                 String(now.getDate()).padStart(2, '0');
  const timeStr = String(now.getHours()).padStart(2, '0') +
                 String(now.getMinutes()).padStart(2, '0') +
                 String(now.getSeconds()).padStart(2, '0');
  return `ORD${dateStr}${timeStr}`;
}

function logSystemEvent(level, event, details, user = 'SYSTEM') {
  try {
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const logSheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.SYSTEM_LOG);
    
    if (logSheet) {
      logSheet.appendRow([
        new Date(),
        level,
        event,
        details,
        user
      ]);
      
      // ログが1000行を超えたら古いものを削除
      if (logSheet.getLastRow() > 1000) {
        logSheet.deleteRows(2, 100);
      }
    }
  } catch (error) {
    console.error('❌ ログ記録エラー:', error);
  }
}

function testConnection() {
  return {
    success: true,
    timestamp: new Date().toISOString(),
    version: SYSTEM_CONFIG.version,
    status: 'operational'
  };
}

// ===== 手動実行関数 =====
function forceInitializeSystem() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    
    // 既存のシートを削除（注意：データが失われます）
    Object.values(SYSTEM_CONFIG.sheets).forEach(sheetName => {
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (sheet) {
        spreadsheet.deleteSheet(sheet);
      }
    });
    
    // 再初期化
    checkAndInitializeSystem();
    
    console.log('✅ システム強制初期化完了');
    return { success: true, message: 'システムを初期化しました' };
  } catch (error) {
    console.error('❌ 強制初期化エラー:', error);
    return { success: false, message: error.toString() };
  }
}

/**
 * メール機能の手動テスト関数
 */
function manualEmailTest() {
  console.log('🧪 手動メールテスト実行');
  const result = testEmailSending();
  console.log('🧪 テスト完了:', result);
  return result;
}