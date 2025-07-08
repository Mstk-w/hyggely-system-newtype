/**
 * Hyggelyカンパーニュ専門店 予約管理システム - GAS環境最適化版
 * v5.3.4 - GAS環境最適化・デバッグ強化版
 * 
 * 🔧 主な修正内容：
 * - サンドボックス制限を考慮した設計
 * - getOrderList()と統計計算の統一
 * - デバッグ情報の大幅強化
 * - データ取得と表示処理の分離
 */

// ===== システム設定 =====
const SYSTEM_CONFIG = {
  spreadsheetId: '18Wdo9hYY8KBF7KULuD8qAODDd5z4O_WvkMCekQpptJ8',
  sheets: {
    ORDER: '予約管理票',
    INVENTORY: '在庫管理票',
    PRODUCT_MASTER: '商品マスタ',
    SYSTEM_LOG: 'システムログ'
  },
  adminPassword: 'hyggelyAdmin2024',
  version: '5.3.4',
  // メール設定（固定）
  email: {
    adminEmail: 'hyggely2021@gmail.com',
    enabled: true,
    customerSubject: 'Hyggelyカンパーニュ専門店 ご予約完了確認',
    customerBody: '{lastName} {firstName} 様\n\nHyggely事前予約システムをご利用いただき誠にありがとうございます。\n以下の注文内容で承りましたのでお知らせします。\n\n予約ID: {orderId}\n\n{orderItems}\n\n・合計　　　¥{totalPrice}\n\n・受取日時：{pickupDateTime}\n\n受取日当日に現金またはPayPayでお支払いいただきます。\n当日は気をつけてお越しください。\n\n※このメールは自動送信されています。返信はできませんのでご了承ください。\nご不明な点がございましたら、Hyggely公式LINEまでお問い合わせください。',
    adminSubject: '【新規予約】{lastName} {firstName}様 - {orderId}',
    adminBody: '【新規予約通知】\n\n予約ID: {orderId}\n\nお客様情報:\n・氏名: {lastName} {firstName} 様\n・メール: {email}\n・受取日時: {pickupDateTime}\n\n注文内容:\n{orderItems}\n\n合計金額: ¥{totalPrice}\n\n※予約管理システムより自動送信されています。'
  },
  // キャッシュ設定
  cache: {
    enabled: true,
    duration: 300, // 5分
    keys: {
      inventory: 'inventory_cache',
      orders: 'orders_cache',
      products: 'products_cache'
    }
  },
  // 受取時間の設定
  pickupTimes: {
    start: 11,    // 11時開始
    end: 17,      // 17時終了
    interval: 15  // 15分間隔
  }
};

// ===== 商品マスタ（固定順序） =====
function getDefaultProducts() {
  return [
    {id: 'PRD001', name: 'プレミアムカンパーニュ', price: 1000, order: 1, columnIndex: 6},        // G列
    {id: 'PRD002', name: 'プレミアムカンパーニュ 1/2', price: 600, order: 2, columnIndex: 7},     // H列
    {id: 'PRD003', name: 'レーズン&クルミ', price: 1200, order: 3, columnIndex: 8},              // I列
    {id: 'PRD004', name: 'レーズン&クルミ 1/2', price: 600, order: 4, columnIndex: 9},           // J列
    {id: 'PRD005', name: 'いちじく&クルミ', price: 400, order: 5, columnIndex: 10},               // K列
    {id: 'PRD006', name: '4種のMIXナッツ', price: 400, order: 6, columnIndex: 11},               // L列
    {id: 'PRD007', name: 'MIXドライフルーツ', price: 400, order: 7, columnIndex: 12},            // M列
    {id: 'PRD008', name: 'アールグレイ', price: 350, order: 8, columnIndex: 13},                 // N列
    {id: 'PRD009', name: 'チョコレート', price: 450, order: 9, columnIndex: 14},                 // O列
    {id: 'PRD010', name: 'チーズ', price: 450, order: 10, columnIndex: 15},                     // P列
    {id: 'PRD011', name: 'ひまわりの種', price: 400, order: 11, columnIndex: 16},               // Q列
    {id: 'PRD012', name: 'デーツ', price: 400, order: 12, columnIndex: 17},                    // R列
    {id: 'PRD013', name: 'カレーパン', price: 450, order: 13, columnIndex: 18},                // S列
    {id: 'PRD014', name: 'バターロール', price: 230, order: 14, columnIndex: 19},              // T列
    {id: 'PRD015', name: 'ショコラロール', price: 280, order: 15, columnIndex: 20},             // U列
    {id: 'PRD016', name: '自家製クリームパン', price: 350, order: 16, columnIndex: 21},         // V列
    {id: 'PRD017', name: '自家製あんバター', price: 380, order: 17, columnIndex: 22},           // W列
    {id: 'PRD018', name: '抹茶&ホワイトチョコ', price: 400, order: 18, columnIndex: 23},       // X列
    {id: 'PRD019', name: '黒ごまパン', price: 200, order: 19, columnIndex: 24},                // Y列
    {id: 'PRD020', name: 'レーズンジャムとクリームチーズのパン', price: 350, order: 20, columnIndex: 25}, // Z列
    {id: 'PRD021', name: 'ピーナッツクリームパン', price: 350, order: 21, columnIndex: 26},    // AA列
    {id: 'PRD022', name: 'あん食パン', price: 400, order: 22, columnIndex: 27},                // AB列
    {id: 'PRD023', name: 'コーンパン', price: 400, order: 23, columnIndex: 28},               // AC列
    {id: 'PRD024', name: 'レモンとクリームチーズのミニ食パン', price: 450, order: 24, columnIndex: 29}, // AD列
    {id: 'PRD025', name: 'ピザ マルゲリータ', price: 1100, order: 25, columnIndex: 30},        // AE列
    {id: 'PRD026', name: 'ピタパンサンド', price: 800, order: 26, columnIndex: 31},             // AF列
    {id: 'PRD027', name: 'フォカッチャ', price: 300, order: 27, columnIndex: 32}                // AG列
  ];
}

// ===== メインエントリーポイント =====
function doGet(e) {
  try {
    console.log('🍞 Hyggelyシステム起動 v' + SYSTEM_CONFIG.version);
    
    const params = e?.parameter || {};
    const action = params.action || '';
    const password = params.password || '';
    
    // システム初期化
    const initResult = checkAndInitializeSystem();
    if (!initResult.success) {
      console.error('❌ システム初期化失敗:', initResult.error);
      return createErrorPage('システム初期化エラー', initResult.error);
    }
    
    switch (action) {
      case 'dashboard':
        return handleDashboard(password);
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

// ===== Dashboard表示処理 =====
function handleDashboard(password) {
  try {
    // パスワード認証
    if (!password || password !== SYSTEM_CONFIG.adminPassword) {
      console.log('⚠️ 認証失敗 - 無効なパスワード');
      return createAuthenticationPage();
    }
    
    console.log('✅ Dashboard認証成功');
    
    // ダッシュボードHTML生成
    let htmlOutput;
    try {
      htmlOutput = HtmlService.createHtmlOutputFromFile('Dashboard');
    } catch (htmlError) {
      console.error('❌ Dashboard.html読み込みエラー:', htmlError);
      return createFallbackDashboard();
    }
    
    return htmlOutput
      .setTitle('Hyggelyカンパーニュ専門店 管理者ダッシュボード v' + SYSTEM_CONFIG.version)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    
  } catch (error) {
    console.error('❌ Dashboard処理エラー:', error);
    logSystemEvent('ERROR', 'Dashboard処理エラー', error.toString());
    return createFallbackDashboard();
  }
}

// ===== 🔧 統一されたデータ取得エンジン =====
function getRawOrderData() {
  try {
    console.log('📊 生データ取得開始（v5.3.4）');
    
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const orderSheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.ORDER);
    
    if (!orderSheet) {
      console.error('❌ 予約管理票シートが見つかりません');
      return { success: false, error: 'シートなし', data: [] };
    }
    
    const lastRow = orderSheet.getLastRow();
    if (lastRow <= 1) {
      console.log('⚠️ 予約データが存在しません（ヘッダーのみ）');
      return { success: true, data: [], totalRows: 1 };
    }
    
    // 全データ取得
    const allData = orderSheet.getDataRange().getValues();
    console.log(`📋 生データ取得完了: ${allData.length}行`);
    
    return { 
      success: true, 
      data: allData, 
      totalRows: allData.length,
      headerRow: allData[0] || []
    };
    
  } catch (error) {
    console.error('❌ 生データ取得エラー:', error);
    return { success: false, error: error.toString(), data: [] };
  }
}

function parseOrderFromRow(row, rowIndex, productMaster) {
  try {
    // 基本データ存在チェック
    if (!row[0] || (!row[1] && !row[2])) {
      return null;
    }
    
    // 引渡状況チェック（AJ列 = インデックス35）
    const deliveryStatus = row[35] || '未引渡';
    
    // 🔧 商品データ解析
    const selectedProducts = [];
    let totalCalculated = 0;
    
    productMaster.forEach((product) => {
      const quantity = parseInt(row[product.columnIndex]) || 0;
      
      if (quantity > 0) {
        const subtotal = product.price * quantity;
        totalCalculated += subtotal;
        
        selectedProducts.push({
          id: product.id,
          name: product.name,
          price: product.price,
          quantity: quantity,
          subtotal: subtotal,
          displayText: `・${product.name} ${quantity}個 ¥${subtotal.toLocaleString()}`
        });
      }
    });
    
    // 🔧 受取日時の安定化処理
    let pickupDate = '';
    let pickupTime = '';
    
    // 受取日処理（E列 = インデックス4）
    if (row[4]) {
      try {
        if (row[4] instanceof Date) {
          pickupDate = Utilities.formatDate(row[4], Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } else if (typeof row[4] === 'string' && row[4].trim()) {
          const parsedDate = new Date(row[4]);
          if (!isNaN(parsedDate.getTime())) {
            pickupDate = Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          } else {
            pickupDate = row[4].toString().trim();
          }
        } else if (typeof row[4] === 'number') {
          // Excelシリアル値の場合
          const date = new Date((row[4] - 25569) * 86400 * 1000);
          pickupDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
      } catch (e) {
        console.warn(`⚠️ 行${rowIndex}: 受取日変換エラー:`, e);
        pickupDate = row[4].toString();
      }
    }
    
    // 受取時間処理（F列 = インデックス5）
    if (row[5]) {
      pickupTime = row[5].toString().trim();
    }
    
    // 予約ID取得
    let orderId = row[36] || generateOrderId();
    
    // 完全な予約オブジェクト作成
    const order = {
      rowIndex: rowIndex,
      timestamp: row[0] || new Date(),
      lastName: (row[1] || '').toString().trim(),
      firstName: (row[2] || '').toString().trim(),
      email: (row[3] || '').toString().trim(),
      pickupDate: pickupDate,
      pickupTime: pickupTime,
      selectedProducts: selectedProducts,
      productsDisplay: selectedProducts.map(p => p.displayText).join('\n'),
      note: (row[33] || '').toString().trim(), // AH列
      totalPrice: parseFloat(row[34]) || totalCalculated, // AI列
      deliveryStatus: deliveryStatus,
      orderId: orderId,
      isDelivered: deliveryStatus === '引渡済',
      updatedAt: new Date(),
      // デバッグ用情報
      _debug: {
        originalRowIndex: rowIndex,
        hasProducts: selectedProducts.length > 0,
        calculatedTotal: totalCalculated,
        storedTotal: parseFloat(row[34]) || 0,
        deliveryStatus: deliveryStatus
      }
    };
    
    return order;
    
  } catch (error) {
    console.error(`❌ 行${rowIndex}解析エラー:`, error);
    return null;
  }
}

// ===== 🔧 統一された予約データ取得関数 =====
function getOrderList() {
  try {
    console.log('📊 予約一覧取得開始（v5.3.4 統一版）');
    
    // 生データ取得
    const rawResult = getRawOrderData();
    if (!rawResult.success) {
      console.error('❌ 生データ取得失敗:', rawResult.error);
      return [];
    }
    
    const allData = rawResult.data;
    if (allData.length <= 1) {
      console.log('⚠️ データなし');
      return [];
    }
    
    // 商品マスタ取得
    const productMaster = getDefaultProducts();
    console.log(`📦 商品マスタ: ${productMaster.length}商品`);
    
    const orders = [];
    
    // データ行を処理（2行目以降）
    for (let i = 1; i < allData.length; i++) {
      const row = allData[i];
      const rowNumber = i + 1;
      
      try {
        const order = parseOrderFromRow(row, rowNumber, productMaster);
        
        if (order && order.deliveryStatus === '未引渡') {
          console.log(`✅ 行${rowNumber}: 予約追加 - ${order.lastName} ${order.firstName} (${order.selectedProducts.length}商品)`);
          orders.push(order);
        } else if (order && order.deliveryStatus !== '未引渡') {
          console.log(`🔄 行${rowNumber}: 引渡済のためスキップ - ${order.lastName} ${order.firstName}`);
        }
        
      } catch (rowError) {
        console.error(`❌ 行${rowNumber}処理エラー:`, rowError);
        continue;
      }
    }
    
    // ソート処理
    orders.sort((a, b) => {
      try {
        // 受取日でソート
        const dateA = new Date(a.pickupDate + ' ' + (a.pickupTime || '00:00'));
        const dateB = new Date(b.pickupDate + ' ' + (b.pickupTime || '00:00'));
        
        const timeDiff = dateA.getTime() - dateB.getTime();
        if (timeDiff !== 0) return timeDiff;
        
        // 受取日が同じ場合はタイムスタンプでソート
        const timestampA = new Date(a.timestamp);
        const timestampB = new Date(b.timestamp);
        return timestampA.getTime() - timestampB.getTime();
      } catch (sortError) {
        console.warn('⚠️ ソート処理エラー:', sortError);
        return 0;
      }
    });
    
    console.log(`📊 予約一覧取得完了: ${orders.length}件（未引渡のみ）`);
    
    // デバッグ情報出力
    if (orders.length > 0) {
      console.log('🔍 最初の予約サンプル:', {
        orderId: orders[0].orderId,
        customer: `${orders[0].lastName} ${orders[0].firstName}`,
        pickupDateTime: `${orders[0].pickupDate} ${orders[0].pickupTime}`,
        productsCount: orders[0].selectedProducts.length,
        totalPrice: orders[0].totalPrice,
        deliveryStatus: orders[0].deliveryStatus
      });
    } else {
      console.log('🔍 未引渡の予約が見つかりませんでした');
    }
    
    return orders;
    
  } catch (error) {
    console.error('❌ 予約一覧取得エラー:', error);
    logSystemEvent('ERROR', '予約一覧取得エラー', error.toString());
    return [];
  }
}

// ===== 全予約データ取得関数（統一版） =====
function getAllOrders() {
  try {
    console.log('📊 全予約データ取得開始（v5.3.4 統一版）');
    
    // 生データ取得
    const rawResult = getRawOrderData();
    if (!rawResult.success) {
      console.error('❌ 生データ取得失敗:', rawResult.error);
      return [];
    }
    
    const allData = rawResult.data;
    if (allData.length <= 1) {
      return [];
    }
    
    const productMaster = getDefaultProducts();
    const orders = [];
    
    // データ行を処理（2行目以降）
    for (let i = 1; i < allData.length; i++) {
      const row = allData[i];
      const rowNumber = i + 1;
      
      const order = parseOrderFromRow(row, rowNumber, productMaster);
      if (order) {
        orders.push(order);
      }
    }
    
    console.log(`📊 全予約データ取得完了: ${orders.length}件`);
    return orders;
    
  } catch (error) {
    console.error('❌ 全予約データ取得エラー:', error);
    return [];
  }
}

// ===== 🔧 統計データ取得の改善（統一版） =====
function getDashboardStats() {
  try {
    console.log('📊 統計データ取得開始（v5.3.4 統一版）');
    
    if (SYSTEM_CONFIG.cache.enabled) {
      const cache = CacheService.getScriptCache();
      const cachedStats = cache.get('dashboard_stats');
      
      if (cachedStats) {
        console.log('📋 キャッシュからデータ取得');
        return JSON.parse(cachedStats);
      }
    }
    
    // 統一されたデータ取得エンジンを使用
    const allOrders = getAllOrders();
    const pendingOrders = allOrders.filter(order => order.deliveryStatus === '未引渡');
    const deliveredOrders = allOrders.filter(order => order.deliveryStatus === '引渡済');
    
    console.log(`📊 統計ベースデータ: 全${allOrders.length}件, 未引渡${pendingOrders.length}件, 引渡済${deliveredOrders.length}件`);
    
    const inventory = getInventoryDataForForm();
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    // 今日の予約（未引渡）を正確に計算
    const todayPendingOrders = pendingOrders.filter(order => {
      if (!order.pickupDate) return false;
      try {
        const pickupDate = new Date(order.pickupDate);
        pickupDate.setHours(0, 0, 0, 0);
        const isToday = pickupDate.getTime() === today.getTime();
        if (isToday) {
          console.log(`📅 今日の予約発見: ${order.lastName} ${order.firstName} - ${order.pickupDate} ${order.pickupTime}`);
        }
        return isToday;
      } catch (e) {
        console.warn('⚠️ 日付解析エラー:', order.pickupDate, e);
        return false;
      }
    });
    
    console.log(`📅 今日の予約（未引渡）: ${todayPendingOrders.length}件`);
    
    // 在庫アラート
    const outOfStock = inventory.filter(p => p.remaining <= 0);
    const lowStock = inventory.filter(p => p.remaining > 0 && p.remaining <= (p.minStock || 3));
    
    // 月売上計算
    const currentMonth = today.getMonth();
    const currentYear = today.getFullYear();
    
    const monthDeliveredOrders = deliveredOrders.filter(order => {
      if (!order.timestamp) return false;
      try {
        const orderDate = new Date(order.timestamp);
        return orderDate.getMonth() === currentMonth && 
               orderDate.getFullYear() === currentYear;
      } catch (e) {
        return false;
      }
    });
    
    const monthRevenue = monthDeliveredOrders.reduce((sum, order) => sum + (order.totalPrice || 0), 0);
    
    // 今日の売上計算
    const todayDeliveredOrders = deliveredOrders.filter(order => {
      if (!order.timestamp) return false;
      try {
        const orderDate = new Date(order.timestamp);
        orderDate.setHours(0, 0, 0, 0);
        return orderDate.getTime() === today.getTime();
      } catch (e) {
        return false;
      }
    });
    
    const todayRevenue = todayDeliveredOrders.reduce((sum, order) => sum + (order.totalPrice || 0), 0);
    
    const stats = {
      todayOrdersCount: todayPendingOrders.length,
      pendingOrdersCount: pendingOrders.length,
      deliveredOrdersCount: deliveredOrders.length,
      outOfStockCount: outOfStock.length,
      lowStockCount: lowStock.length,
      totalProducts: inventory.length,
      todayRevenue: todayRevenue,
      monthRevenue: monthRevenue,
      systemVersion: SYSTEM_CONFIG.version,
      lastUpdate: new Date().toISOString(),
      systemHealth: {
        totalOrdersCount: allOrders.length,
        pendingOrdersCount: pendingOrders.length,
        deliveredOrdersCount: deliveredOrders.length,
        inventoryItems: inventory.length,
        cacheEnabled: SYSTEM_CONFIG.cache.enabled,
        lastDataUpdate: new Date().toISOString()
      },
      // デバッグ情報
      _debug: {
        todayPendingDetails: todayPendingOrders.map(o => ({
          customer: `${o.lastName} ${o.firstName}`,
          pickupDate: o.pickupDate,
          pickupTime: o.pickupTime,
          orderId: o.orderId
        }))
      }
    };
    
    // キャッシュに保存
    if (SYSTEM_CONFIG.cache.enabled) {
      const cache = CacheService.getScriptCache();
      cache.put('dashboard_stats', JSON.stringify(stats), SYSTEM_CONFIG.cache.duration);
    }
    
    console.log('📊 統計データ取得完了（v5.3.4）');
    console.log(`📊 結果: 今日の予約${stats.todayOrdersCount}件, 未引渡${stats.pendingOrdersCount}件, 引渡済${stats.deliveredOrdersCount}件`);
    
    return stats;
    
  } catch (error) {
    console.error('❌ 統計データ取得エラー:', error);
    logSystemEvent('ERROR', '統計データ取得エラー', error.toString());
    return {
      todayOrdersCount: 0,
      pendingOrdersCount: 0,
      deliveredOrdersCount: 0,
      outOfStockCount: 0,
      lowStockCount: 0,
      totalProducts: 0,
      todayRevenue: 0,
      monthRevenue: 0,
      systemVersion: SYSTEM_CONFIG.version,
      lastUpdate: new Date().toISOString(),
      error: error.toString()
    };
  }
}

// ===== 予約フィルター関数 =====
function getFilteredOrders(filterType = 'pending') {
  try {
    const allOrders = getAllOrders();
    
    switch (filterType) {
      case 'pending':
        return allOrders.filter(order => order.deliveryStatus === '未引渡');
      case 'delivered':
        return allOrders.filter(order => order.deliveryStatus === '引渡済');
      case 'all':
      default:
        return allOrders;
    }
  } catch (error) {
    console.error('❌ 予約フィルター取得エラー:', error);
    return [];
  }
}

// ===== 引渡状況変更関数 =====
function updateDeliveryStatus(orderId, newStatus, updatedBy = 'ADMIN') {
  try {
    console.log(`🔄 引渡状況変更開始: ${orderId} -> ${newStatus}`);
    
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const orderSheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.ORDER);
    
    const allData = orderSheet.getDataRange().getValues();
    let targetRowIndex = -1;
    let orderData = null;
    
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][36] === orderId) {
        targetRowIndex = i + 1;
        orderData = allData[i];
        break;
      }
    }
    
    if (targetRowIndex === -1) {
      throw new Error(`予約ID ${orderId} が見つかりません`);
    }
    
    const deliveryStatusCol = 36; // AJ列
    orderSheet.getRange(targetRowIndex, deliveryStatusCol).setValue(newStatus);
    
    let revenueUpdate = null;
    if (newStatus === '引渡済') {
      const totalPrice = parseFloat(orderData[34]) || 0;
      revenueUpdate = {
        amount: totalPrice,
        date: new Date(),
        orderId: orderId,
        customer: `${orderData[1]} ${orderData[2]}`
      };
    }
    
    clearCache();
    updateInventoryFromOrders();
    
    const customerName = `${orderData[1]} ${orderData[2]}`;
    logSystemEvent('INFO', '引渡状況変更', 
      `予約ID: ${orderId}, 顧客: ${customerName}, 新状況: ${newStatus}, 更新者: ${updatedBy}`);
    
    console.log(`✅ 引渡状況変更完了: ${orderId} -> ${newStatus}`);
    
    return {
      success: true,
      message: `引渡状況を「${newStatus}」に変更しました`,
      orderId: orderId,
      newStatus: newStatus,
      customerName: customerName,
      revenueUpdate: revenueUpdate
    };
    
  } catch (error) {
    console.error('❌ 引渡状況変更エラー:', error);
    logSystemEvent('ERROR', '引渡状況変更エラー', `予約ID: ${orderId}, エラー: ${error.toString()}`);
    
    return {
      success: false,
      message: `引渡状況の変更に失敗しました: ${error.message}`,
      orderId: orderId
    };
  }
}

// ===== 一括引渡状況変更関数 =====
function bulkUpdateDeliveryStatus(orderIds, newStatus, updatedBy = 'ADMIN') {
  try {
    console.log(`🔄 一括引渡状況変更開始: ${orderIds.length}件 -> ${newStatus}`);
    
    const results = [];
    let successCount = 0;
    let errorCount = 0;
    let totalRevenue = 0;
    
    orderIds.forEach(orderId => {
      const result = updateDeliveryStatus(orderId, newStatus, updatedBy);
      results.push(result);
      
      if (result.success) {
        successCount++;
        if (result.revenueUpdate) {
          totalRevenue += result.revenueUpdate.amount;
        }
      } else {
        errorCount++;
      }
    });
    
    logSystemEvent('INFO', '一括引渡状況変更', 
      `対象: ${orderIds.length}件, 成功: ${successCount}件, 失敗: ${errorCount}件, 新状況: ${newStatus}`);
    
    console.log(`✅ 一括引渡状況変更完了: 成功${successCount}件, 失敗${errorCount}件`);
    
    return {
      success: errorCount === 0,
      message: `一括変更完了: 成功${successCount}件, 失敗${errorCount}件`,
      successCount: successCount,
      errorCount: errorCount,
      totalRevenue: totalRevenue,
      results: results
    };
    
  } catch (error) {
    console.error('❌ 一括引渡状況変更エラー:', error);
    logSystemEvent('ERROR', '一括引渡状況変更エラー', error.toString());
    
    return {
      success: false,
      message: `一括変更に失敗しました: ${error.message}`,
      successCount: 0,
      errorCount: orderIds.length
    };
  }
}

// ===== 🔧 強化されたデバッグ関数 =====
function debugOrderData() {
  try {
    console.log('🔍 デバッグモード開始（v5.3.4 強化版）');
    
    // 生データ取得
    const rawResult = getRawOrderData();
    
    if (!rawResult.success) {
      return {
        error: rawResult.error,
        totalRows: 0,
        totalColumns: 0,
        headers: [],
        sampleData: []
      };
    }
    
    const data = rawResult.data;
    const headers = rawResult.headerRow;
    const sampleData = data.length > 1 ? data[1] : [];
    
    console.log('📋 基本情報:');
    console.log(`- 総行数: ${data.length}`);
    console.log(`- 総列数: ${headers.length}`);
    
    // 🔧 スプレッドシート列構造の詳細確認
    const columnMapping = {
      'A列（タイムスタンプ）': data.length > 1 ? data[1][0] : null,
      'B列（姓）': data.length > 1 ? data[1][1] : null,
      'C列（名）': data.length > 1 ? data[1][2] : null,
      'D列（メール）': data.length > 1 ? data[1][3] : null,
      'E列（受取日）': data.length > 1 ? data[1][4] : null,
      'F列（受取時間）': data.length > 1 ? data[1][5] : null,
      'AH列（要望）': data.length > 1 ? data[1][33] : null,
      'AI列（合計金額）': data.length > 1 ? data[1][34] : null,
      'AJ列（引渡済）': data.length > 1 ? data[1][35] : null,
      'AK列（予約ID）': data.length > 1 ? data[1][36] : null
    };
    
    // 🔧 商品列の詳細デバッグ
    const productMaster = getDefaultProducts();
    const productColumns = {};
    
    productMaster.forEach((product, index) => {
      const columnLetter = String.fromCharCode(71 + index); // G=71
      if (data.length > 1) {
        productColumns[`${product.name} (${columnLetter}列-${product.columnIndex})`] = data[1][product.columnIndex];
      }
    });
    
    // 実際のデータ処理テスト
    console.log('🔍 データ処理テスト開始...');
    const testOrders = getOrderList();
    const testStats = getDashboardStats();
    
    console.log(`🔍 処理結果: getOrderList=${testOrders.length}件, 統計今日の予約=${testStats.todayOrdersCount}件`);
    
    // 引渡状況の分布確認
    const deliveryStatusDistribution = {};
    for (let i = 1; i < data.length; i++) {
      const status = data[i][35] || '未設定';
      deliveryStatusDistribution[status] = (deliveryStatusDistribution[status] || 0) + 1;
    }
    
    console.log('🔍 引渡状況分布:', deliveryStatusDistribution);
    
    return {
      systemVersion: SYSTEM_CONFIG.version,
      totalRows: data.length,
      totalColumns: headers.length,
      headers: headers,
      sampleData: sampleData,
      columnMapping: columnMapping,
      productColumns: productColumns,
      deliveryStatusDistribution: deliveryStatusDistribution,
      processedOrders: testOrders.length,
      statisticsData: {
        todayOrdersCount: testStats.todayOrdersCount,
        pendingOrdersCount: testStats.pendingOrdersCount,
        deliveredOrdersCount: testStats.deliveredOrdersCount
      },
      spreadsheetId: SYSTEM_CONFIG.spreadsheetId,
      sheetName: SYSTEM_CONFIG.sheets.ORDER,
      cacheEnabled: SYSTEM_CONFIG.cache.enabled,
      productMaster: productMaster.map(p => ({
        name: p.name,
        columnIndex: p.columnIndex,
        column: String.fromCharCode(71 + p.columnIndex - 6) // G=6なので71+0=71=G
      })),
      debugTimestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('❌ デバッグ確認エラー:', error);
    return {
      error: error.toString(),
      totalRows: 0,
      totalColumns: 0,
      headers: [],
      sampleData: [],
      debugTimestamp: new Date().toISOString()
    };
  }
}

// ===== システム初期化 =====
function checkAndInitializeSystem() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const sheets = spreadsheet.getSheets().map(s => s.getName());
    
    let missingSheets = [];
    Object.values(SYSTEM_CONFIG.sheets).forEach(sheetName => {
      if (!sheets.includes(sheetName)) {
        try {
          initializeSheet(spreadsheet, sheetName);
          console.log('✅ シート作成完了:', sheetName);
        } catch (error) {
          console.error('❌ シート作成エラー:', sheetName, error);
          missingSheets.push(sheetName);
        }
      }
    });
    
    if (missingSheets.length > 0) {
      return {
        success: false,
        error: `シート初期化失敗: ${missingSheets.join(', ')}`
      };
    }
    
    console.log('✅ システム初期化完了');
    return { success: true };
    
  } catch (error) {
    console.error('❌ システム初期化エラー:', error);
    return {
      success: false,
      error: 'スプレッドシートにアクセスできません: ' + error.message
    };
  }
}

// ===== 在庫データ取得 =====
function getInventoryDataForForm() {
  try {
    if (SYSTEM_CONFIG.cache.enabled) {
      const cache = CacheService.getScriptCache();
      const cachedInventory = cache.get(SYSTEM_CONFIG.cache.keys.inventory);
      
      if (cachedInventory) {
        console.log('📋 在庫データをキャッシュから取得');
        return JSON.parse(cachedInventory);
      }
    }
    
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const inventorySheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.INVENTORY);
    const products = getProductMaster().filter(p => p.enabled);
    
    if (!inventorySheet || inventorySheet.getLastRow() <= 1) {
      const defaultInventory = products.map(p => ({
        id: p.id,
        name: p.name,
        price: p.price,
        stock: 10,
        reserved: 0,
        remaining: 10,
        minStock: 3
      }));
      
      if (SYSTEM_CONFIG.cache.enabled) {
        const cache = CacheService.getScriptCache();
        cache.put(SYSTEM_CONFIG.cache.keys.inventory, JSON.stringify(defaultInventory), SYSTEM_CONFIG.cache.duration);
      }
      
      return defaultInventory;
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
    
    const sortedInventory = inventory.sort((a, b) => {
      const aProduct = products.find(p => p.id === a.id);
      const bProduct = products.find(p => p.id === b.id);
      return (aProduct?.order || 0) - (bProduct?.order || 0);
    });
    
    if (SYSTEM_CONFIG.cache.enabled) {
      const cache = CacheService.getScriptCache();
      cache.put(SYSTEM_CONFIG.cache.keys.inventory, JSON.stringify(sortedInventory), SYSTEM_CONFIG.cache.duration);
    }
    
    return sortedInventory;
    
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

// ===== 商品マスタ取得 =====
function getProductMaster() {
  try {
    if (SYSTEM_CONFIG.cache.enabled) {
      const cache = CacheService.getScriptCache();
      const cachedProducts = cache.get(SYSTEM_CONFIG.cache.keys.products);
      
      if (cachedProducts) {
        return JSON.parse(cachedProducts);
      }
    }
    
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const masterSheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.PRODUCT_MASTER);
    
    if (!masterSheet || masterSheet.getLastRow() <= 1) {
      const defaultProducts = getDefaultProducts().map(p => ({ ...p, enabled: true }));
      
      if (SYSTEM_CONFIG.cache.enabled) {
        const cache = CacheService.getScriptCache();
        cache.put(SYSTEM_CONFIG.cache.keys.products, JSON.stringify(defaultProducts), SYSTEM_CONFIG.cache.duration);
      }
      
      return defaultProducts;
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
    
    const sortedProducts = products.sort((a, b) => a.order - b.order);
    
    if (SYSTEM_CONFIG.cache.enabled) {
      const cache = CacheService.getScriptCache();
      cache.put(SYSTEM_CONFIG.cache.keys.products, JSON.stringify(sortedProducts), SYSTEM_CONFIG.cache.duration);
    }
    
    return sortedProducts;
  } catch (error) {
    console.error('❌ 商品マスタ取得エラー:', error);
    return getDefaultProducts().map(p => ({ ...p, enabled: true }));
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
    
    const productMaster = getDefaultProducts();
    
    const reservations = {};
    productMaster.forEach(p => reservations[p.id] = 0);
    
    const orderData = orderSheet.getDataRange().getValues();
    for (let i = 1; i < orderData.length; i++) {
      const row = orderData[i];
      const isDelivered = row[35] === '引渡済';
      
      if (!isDelivered) {
        productMaster.forEach(product => {
          const quantity = parseInt(row[product.columnIndex]) || 0;
          if (quantity > 0) {
            reservations[product.id] += quantity;
          }
        });
      }
    }
    
    const inventoryData = inventorySheet.getDataRange().getValues();
    for (let i = 1; i < inventoryData.length; i++) {
      const productId = inventoryData[i][0];
      if (reservations[productId] !== undefined) {
        inventorySheet.getRange(i + 1, 5).setValue(reservations[productId]);
        
        const stock = inventoryData[i][3] || 0;
        const remaining = Math.max(0, stock - reservations[productId]);
        inventorySheet.getRange(i + 1, 6).setValue(remaining);
        inventorySheet.getRange(i + 1, 8).setValue(new Date());
      }
    }
    
    clearCache();
    
  } catch (error) {
    console.error('❌ 在庫更新エラー:', error);
  }
}

// ===== 🔧 Dashboard用の未実装関数を追加 =====

// 在庫更新モーダル用
function showBulkUpdateModal() {
  console.log('一括更新モーダル表示');
  return { success: true, message: '一括更新モーダルを表示します' };
}

// 在庫リセット
function resetAllInventory() {
  try {
    const products = getProductMaster().filter(p => p.enabled);
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const inventorySheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.INVENTORY);
    
    inventorySheet.clear();
    initInventory(inventorySheet);
    
    clearCache();
    logSystemEvent('INFO', '在庫リセット', '全商品の在庫を初期値にリセットしました');
    
    return { success: true, message: '在庫をリセットしました' };
  } catch (error) {
    console.error('❌ 在庫リセットエラー:', error);
    return { success: false, message: '在庫リセットに失敗しました: ' + error.message };
  }
}

// 在庫エクスポート
function exportInventory() {
  try {
    const inventory = getInventoryDataForForm();
    const csvData = [
      ['商品ID', '商品名', '単価', '在庫数', '予約数', '残数', '最低在庫数', '更新日']
    ];
    
    inventory.forEach(item => {
      csvData.push([
        item.id,
        item.name,
        item.price,
        item.stock,
        item.reserved,
        item.remaining,
        item.minStock,
        new Date().toISOString()
      ]);
    });
    
    return { success: true, data: csvData, message: '在庫データをエクスポートしました' };
  } catch (error) {
    console.error('❌ 在庫エクスポートエラー:', error);
    return { success: false, message: '在庫エクスポートに失敗しました: ' + error.message };
  }
}

// 新商品追加
function addNewProduct(productData) {
  try {
    const { name, price, order, stock } = productData;
    
    if (!name || !price) {
      throw new Error('商品名と価格は必須です');
    }
    
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const masterSheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.PRODUCT_MASTER);
    const inventorySheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.INVENTORY);
    
    const existingProducts = getProductMaster();
    const newId = `PRD${String(existingProducts.length + 1).padStart(3, '0')}`;
    
    const newProductRow = [
      newId,
      name,
      parseInt(price),
      parseInt(order) || existingProducts.length + 1,
      true,
      new Date(),
      new Date()
    ];
    masterSheet.appendRow(newProductRow);
    
    const newInventoryRow = [
      newId,
      name,
      parseInt(price),
      parseInt(stock) || 10,
      0,
      parseInt(stock) || 10,
      3,
      new Date()
    ];
    inventorySheet.appendRow(newInventoryRow);
    
    clearCache();
    logSystemEvent('INFO', '新商品追加', `商品名: ${name}, ID: ${newId}`);
    
    return { success: true, message: `商品「${name}」を追加しました`, productId: newId };
  } catch (error) {
    console.error('❌ 新商品追加エラー:', error);
    logSystemEvent('ERROR', '新商品追加エラー', error.toString());
    return { success: false, message: '商品追加に失敗しました: ' + error.message };
  }
}

// 予約データエクスポート
function exportOrders() {
  try {
    const orders = getAllOrders();
    const csvData = [
      ['予約ID', 'タイムスタンプ', '姓', '名', 'メール', '受取日', '受取時間', 
       '商品詳細', '合計金額', '引渡状況', 'その他要望']
    ];
    
    orders.forEach(order => {
      csvData.push([
        order.orderId,
        order.timestamp,
        order.lastName,
        order.firstName,
        order.email,
        order.pickupDate,
        order.pickupTime,
        order.productsDisplay,
        order.totalPrice,
        order.deliveryStatus,
        order.note
      ]);
    });
    
    return { success: true, data: csvData, message: '予約データをエクスポートしました' };
  } catch (error) {
    console.error('❌ 予約エクスポートエラー:', error);
    return { success: false, message: '予約エクスポートに失敗しました: ' + error.message };
  }
}

// 在庫数更新
function updateStock(productId, newStock, minStock) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const inventorySheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.INVENTORY);
    
    const data = inventorySheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === productId) {
        inventorySheet.getRange(i + 1, 4).setValue(parseInt(newStock));
        if (minStock !== undefined) {
          inventorySheet.getRange(i + 1, 7).setValue(parseInt(minStock));
        }
        inventorySheet.getRange(i + 1, 8).setValue(new Date());
        break;
      }
    }
    
    updateInventoryFromOrders();
    clearCache();
    
    logSystemEvent('INFO', '在庫更新', `商品ID: ${productId}, 新在庫: ${newStock}`);
    
    return { success: true, message: '在庫を更新しました' };
  } catch (error) {
    console.error('❌ 在庫更新エラー:', error);
    return { success: false, message: '在庫更新に失敗しました: ' + error.message };
  }
}

// ===== 受取時間生成関数 =====
function generatePickupTimeOptions() {
  const times = [];
  const config = SYSTEM_CONFIG.pickupTimes;
  
  for (let hour = config.start; hour <= config.end; hour++) {
    for (let minute = 0; minute < 60; minute += config.interval) {
      const timeString = `${hour.toString().padStart(2, '0')}:${minute.toString().padStart(2, '0')}`;
      times.push({
        value: timeString,
        label: timeString,
        available: true
      });
    }
  }
  
  return times;
}

// ===== キャッシュクリア関数 =====
function clearCache() {
  try {
    const cache = CacheService.getScriptCache();
    cache.removeAll([
      'dashboard_stats',
      SYSTEM_CONFIG.cache.keys.inventory,
      SYSTEM_CONFIG.cache.keys.orders,
      SYSTEM_CONFIG.cache.keys.products
    ]);
    console.log('✅ キャッシュクリア完了');
    return { success: true, message: 'キャッシュをクリアしました' };
  } catch (error) {
    console.error('❌ キャッシュクリアエラー:', error);
    return { success: false, message: 'キャッシュクリアに失敗しました' };
  }
}

// ===== 認証ページ生成 =====
function createAuthenticationPage() {
  const html = `
    <!DOCTYPE html>
    <html lang="ja">
    <head>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <title>管理者認証 - Hyggelyカンパーニュ専門店</title>
      <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
      <style>
        body { 
          background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
          min-height: 100vh;
          display: flex;
          align-items: center;
          justify-content: center;
        }
        .auth-card { 
          max-width: 400px; 
          background: white; 
          border-radius: 15px; 
          box-shadow: 0 8px 30px rgba(139, 69, 19, 0.15);
        }
        .btn-primary {
          background: linear-gradient(135deg, #8B4513 0%, #A0522D 100%);
          border: none;
        }
      </style>
    </head>
    <body>
      <div class="auth-card p-4">
        <div class="text-center mb-4">
          <h3 class="text-primary">🔐 管理者認証</h3>
          <p class="text-muted">ダッシュボードにアクセスするには認証が必要です</p>
        </div>
        <form id="auth-form">
          <div class="mb-3">
            <label for="password" class="form-label">パスワード</label>
            <input type="password" class="form-control" id="password" required>
          </div>
          <button type="submit" class="btn btn-primary w-100">ログイン</button>
        </form>
      </div>
      <script>
        document.getElementById('auth-form').addEventListener('submit', function(e) {
          e.preventDefault();
          const password = document.getElementById('password').value;
          window.location.href = '?action=dashboard&password=' + encodeURIComponent(password);
        });
      </script>
    </body>
    </html>
  `;
  
  return HtmlService.createHtmlOutput(html)
    .setTitle('管理者認証')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ===== フォールバックダッシュボード =====
function createFallbackDashboard() {
  const html = `
    <!DOCTYPE html>
    <html lang="ja">
    <head>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <title>管理者ダッシュボード（簡易版）</title>
      <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
      <style>
        body { background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%); }
        .card { border-radius: 15px; box-shadow: 0 4px 20px rgba(0,0,0,0.08); }
        .btn-primary { background: linear-gradient(135deg, #8B4513 0%, #A0522D 100%); border: none; }
      </style>
    </head>
    <body>
      <div class="container mt-4">
        <div class="card">
          <div class="card-header bg-primary text-white">
            <h3>🛠️ 管理者ダッシュボード（簡易版）</h3>
          </div>
          <div class="card-body">
            <div class="alert alert-info">
              <h5>システム状態</h5>
              <p>メインダッシュボードが利用できないため、簡易版で表示しています。</p>
            </div>
            
            <div class="row">
              <div class="col-md-6 mb-3">
                <div class="card">
                  <div class="card-body text-center">
                    <h5>📊 データ確認</h5>
                    <button class="btn btn-primary" onclick="checkData()">データ状態確認</button>
                  </div>
                </div>
              </div>
              <div class="col-md-6 mb-3">
                <div class="card">
                  <div class="card-body text-center">
                    <h5>🏠 予約フォーム</h5>
                    <a href="?" class="btn btn-primary">予約フォームへ</a>
                  </div>
                </div>
              </div>
            </div>
            
            <div id="data-result" class="mt-3"></div>
          </div>
        </div>
      </div>
      
      <script>
        function checkData() {
          const resultDiv = document.getElementById('data-result');
          resultDiv.innerHTML = '<div class="spinner-border text-primary" role="status"></div> データ確認中...';
          
          google.script.run
            .withSuccessHandler(function(result) {
              resultDiv.innerHTML = '<div class="alert alert-success"><h6>システム状態</h6><pre>' + JSON.stringify(result, null, 2) + '</pre></div>';
            })
            .withFailureHandler(function(error) {
              resultDiv.innerHTML = '<div class="alert alert-danger">エラー: ' + error + '</div>';
            })
            .getDashboardStats();
        }
      </script>
    </body>
    </html>
  `;
  
  return HtmlService.createHtmlOutput(html)
    .setTitle('管理者ダッシュボード（簡易版）')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ===== その他のヘルパー関数 =====
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (error) {
    console.error('❌ HTMLファイル読み込みエラー:', filename, error);
    return '<div>HTMLファイルの読み込みに失敗しました: ' + filename + '</div>';
  }
}

function handleOrderForm() {
  try {
    return HtmlService.createHtmlOutputFromFile('OrderForm')
      .setTitle('Hyggelyカンパーニュ専門店 ご予約フォーム')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    console.error('❌ OrderForm読み込みエラー:', error);
    return createErrorPage('予約フォームの読み込みに失敗しました', 'OrderForm.htmlファイルが見つかりません。');
  }
}

function handleHealthCheck() {
  const health = {
    status: 'healthy',
    version: SYSTEM_CONFIG.version,
    timestamp: new Date().toISOString(),
    spreadsheetId: SYSTEM_CONFIG.spreadsheetId,
    cacheEnabled: SYSTEM_CONFIG.cache.enabled,
    pickupTimes: SYSTEM_CONFIG.pickupTimes
  };
  
  const html = `
    <!DOCTYPE html>
    <html><head><title>システム状態</title></head>
    <body style="font-family: Arial; padding: 20px;">
      <h1>✅ システム正常稼働中 v${SYSTEM_CONFIG.version}</h1>
      <pre>${JSON.stringify(health, null, 2)}</pre>
    </body></html>
  `;
  
  return HtmlService.createHtmlOutput(html);
}

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

// ===== シート初期化関数 =====
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
      case SYSTEM_CONFIG.sheets.SYSTEM_LOG:
        initSystemLog(sheet);
        break;
    }
    
    console.log('✅ シート初期化完了:', sheetName);
  } catch (error) {
    console.error('❌ シート初期化エラー:', sheetName, error);
  }
}

function initOrderSheet(sheet) {
  const basicHeaders = ['タイムスタンプ', '姓', '名', 'メール', '受取日', '受取時間'];
  const products = getDefaultProducts();
  const productHeaders = products.map(p => p.name);
  const finalHeaders = ['その他ご要望', '合計金額', '引渡済', '予約ID'];
  const allHeaders = [...basicHeaders, ...productHeaders, ...finalHeaders];
  
  sheet.getRange(1, 1, 1, allHeaders.length).setValues([allHeaders]);
  
  sheet.getRange(1, 34).setNote('AH列：その他のご要望');
  sheet.getRange(1, 35).setNote('AI列：合計金額');
  sheet.getRange(1, 36).setNote('AJ列：引渡済');
  sheet.getRange(1, 37).setNote('AK列：予約ID');
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

function initSystemLog(sheet) {
  const headers = ['タイムスタンプ', 'レベル', 'イベント', '詳細', 'ユーザー'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}

// ===== 予約処理関数 =====
function processOrder(formData) {
  try {
    console.log('🔄 予約処理開始（v5.3.4）:', JSON.stringify(formData, null, 2));
    
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
    
    const outOfStockItems = [];
    for (let i = 0; i < products.length; i++) {
      const quantity = parseInt(formData[`product_${i}`]) || 0;
      if (quantity > 0) {
        const product = products[i];
        const inventoryItem = inventory.find(inv => inv.id === product.id);
        
        if (!inventoryItem || inventoryItem.remaining < quantity) {
          outOfStockItems.push(`${product.name}（要求: ${quantity}個, 在庫: ${inventoryItem?.remaining || 0}個）`);
          continue;
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
    
    if (outOfStockItems.length > 0) {
      throw new Error(`在庫不足: ${outOfStockItems.join(', ')}`);
    }
    
    if (orderedItems.length === 0) {
      throw new Error('商品を1つ以上選択してください');
    }
    
    // 予約ID生成
    const orderId = generateOrderId();
    
    // スプレッドシート記録
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const orderSheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.ORDER);
    const lastRow = orderSheet.getLastRow() + 1;
    
    const rowData = new Array(37).fill('');
    const currentDate = new Date();
    
    rowData[0] = currentDate;
    rowData[1] = formData.lastName;
    rowData[2] = formData.firstName;
    rowData[3] = formData.email;
    rowData[4] = formData.pickupDate;
    rowData[5] = formData.pickupTime;
    
    // 商品数量（G~AG列：インデックス6~32）
    for (let i = 0; i < products.length; i++) {
      const quantity = parseInt(formData[`product_${i}`]) || 0;
      rowData[6 + i] = quantity;
    }
    
    rowData[33] = formData.note || '';
    rowData[34] = totalPrice;
    rowData[35] = '未引渡';
    rowData[36] = orderId;
    
    // 一括書き込み
    orderSheet.getRange(lastRow, 1, 1, rowData.length).setValues([rowData]);
    
    // キャッシュクリア
    clearCache();
    
    // 在庫更新
    updateInventoryFromOrders();
    
    // メール送信
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
      
      if (logSheet.getLastRow() > 1000) {
        logSheet.deleteRows(2, 100);
      }
    }
  } catch (error) {
    console.error('❌ ログ記録エラー:', error);
  }
}

// ===== 🔧 簡略化されたメール関連関数 =====
function sendOrderEmails(formData, orderedItems, totalPrice, orderId) {
  console.log('📧 メール送信統合処理開始');
  
  const results = {
    success: false,
    customerEmailSent: false,
    adminEmailSent: false,
    errors: []
  };
  
  try {
    const settings = SYSTEM_CONFIG.email;
    console.log('📧 固定設定使用:', settings.adminEmail);
    
    if (!settings.enabled) {
      console.log('⚠️ メール送信が無効に設定されています');
      results.errors.push('メール送信が無効に設定されています');
      return results;
    }
    
    if (!checkGmailPermission()) {
      const errorMsg = 'Gmail送信権限が不足しています。Google Apps Scriptでの権限承認が必要です。';
      console.error('❌ ' + errorMsg);
      results.errors.push(errorMsg);
      logSystemEvent('ERROR', 'Gmail権限エラー', errorMsg);
      return results;
    }
    
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
    
    results.success = results.customerEmailSent || results.adminEmailSent;
    
    console.log('📧 メール送信結果:', results);
    
    const logDetail = `顧客メール: ${results.customerEmailSent ? '成功' : '失敗'}, 管理者メール: ${results.adminEmailSent ? '成功' : '失敗'}`;
    logSystemEvent(results.success ? 'INFO' : 'ERROR', 'メール送信結果', logDetail);
    
    return results;
    
  } catch (error) {
    console.error('❌ メール送信統合処理エラー:', error);
    results.errors.push('統合処理エラー: ' + error.toString());
    logSystemEvent('ERROR', 'メール送信統合エラー', error.toString());
    return results;
  }
}

function checkGmailPermission() {
  try {
    GmailApp.getInboxThreads(0, 1);
    return true;
  } catch (error) {
    console.error('❌ Gmail権限チェック失敗:', error);
    return false;
  }
}

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
    
    GmailApp.sendEmail(formData.email, subject, body);
    
    console.log('✅ 顧客確認メール送信成功');
    logSystemEvent('INFO', '顧客メール送信成功', `宛先: ${formData.email}, 予約ID: ${orderId}`);
    return { success: true };
    
  } catch (error) {
    console.error('❌ 顧客メール送信エラー:', error);
    logSystemEvent('ERROR', '顧客メール送信エラー', error.toString());
    return { success: false, error: error.toString() };
  }
}

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
    
    GmailApp.sendEmail(settings.adminEmail, subject, body);
    
    console.log('✅ 管理者通知メール送信成功');
    logSystemEvent('INFO', '管理者メール送信成功', `宛先: ${settings.adminEmail}, 予約ID: ${orderId}`);
    return { success: true };
    
  } catch (error) {
    console.error('❌ 管理者メール送信エラー:', error);
    logSystemEvent('ERROR', '管理者メール送信エラー', error.toString());
    return { success: false, error: error.toString() };
  }
}

function replaceEmailVariables(template, variables) {
  let result = template;
  
  Object.keys(variables).forEach(key => {
    const regex = new RegExp(`{${key}}`, 'g');
    result = result.replace(regex, variables[key] || '');
  });
  
  return result;
}

// ===== その他のユーティリティ関数 =====
function testConnection() {
  return {
    success: true,
    timestamp: new Date().toISOString(),
    version: SYSTEM_CONFIG.version,
    status: 'operational'
  };
}

console.log('✅ Hyggelyカンパーニュ専門店 予約管理システム v5.3.4 - GAS環境最適化・デバッグ強化版読み込み完了');