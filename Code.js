/**
 * Hyggelyカンパーニュ専門店 予約管理システム - GAS環境最適化版
 * v5.5.2 - null完全排除版
 * 🔧 主な修正内容：
 * - null値の使用を完全に排除
 * - 代替手段でのデータ検証
 * - より堅牢なエラーハンドリング
 * - その他ご要望の空欄時に「とくになし。」を自動設定
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
  version: '5.5.2', // バージョン更新
  email: {
    adminEmail: 'hyggely2021@gmail.com',
    enabled: true,
    customerSubject: 'Hyggelyカンパーニュ専門店 ご予約完了確認',
    customerBody: '{lastName} {firstName} 様\n\nHyggely事前予約システムをご利用いただき誠にありがとうございます。\n以下の注文内容で承りましたのでお知らせします。\n\n予約ID: {orderId}\n\n{orderItems}\n\n・合計　　　¥{totalPrice}\n\n・受取日時：{pickupDateTime}\n\n受取日当日に現金またはPayPayでお支払いいただきます。\n当日は気をつけてお越しください。\n\n※このメールは自動送信されています。返信はできませんのでご了承ください。\nご不明な点がございましたら、Hyggely公式LINEまでお問い合わせください。',
    adminSubject: '【新規予約】{lastName} {firstName}様 - {orderId}',
    adminBody: '【新規予約通知】\n\n予約ID: {orderId}\n\nお客様情報:\n・氏名: {lastName} {firstName} 様\n・メール: {email}\n・受取日時: {pickupDateTime}\n\n注文内容:\n{orderItems}\n\n合計金額: ¥{totalPrice}\n\n※予約管理システムより自動送信されています。'
  },
  cache: {
    enabled: true,
    duration: 300, // 5分
    keys: {
      inventory: 'inventory_cache',
      orders: 'orders_cache',
      products: 'products_cache'
    }
  },
  pickupTimes: {
    start: 11,
    end: 17,
    interval: 15
  }
};

// ===== 商品マスタ（固定順序） =====
function getDefaultProducts() {
  return [
    {id: 'PRD001', name: 'プレミアムカンパーニュ', price: 1000, order: 1, columnIndex: 6},
    {id: 'PRD002', name: 'プレミアムカンパーニュ 1/2', price: 600, order: 2, columnIndex: 7},
    {id: 'PRD003', name: 'レーズン&クルミ', price: 1200, order: 3, columnIndex: 8},
    {id: 'PRD004', name: 'レーズン&クルミ 1/2', price: 600, order: 4, columnIndex: 9},
    {id: 'PRD005', name: 'いちじく&クルミ', price: 400, order: 5, columnIndex: 10},
    {id: 'PRD006', name: '4種のMIXナッツ', price: 400, order: 6, columnIndex: 11},
    {id: 'PRD007', name: 'MIXドライフルーツ', price: 400, order: 7, columnIndex: 12},
    {id: 'PRD008', name: 'アールグレイ', price: 350, order: 8, columnIndex: 13},
    {id: 'PRD009', name: 'チョコレート', price: 450, order: 9, columnIndex: 14},
    {id: 'PRD010', name: 'チーズ', price: 450, order: 10, columnIndex: 15},
    {id: 'PRD011', name: 'ひまわりの種', price: 400, order: 11, columnIndex: 16},
    {id: 'PRD012', name: 'デーツ', price: 400, order: 12, columnIndex: 17},
    {id: 'PRD013', name: 'カレーパン', price: 450, order: 13, columnIndex: 18},
    {id: 'PRD014', name: 'バターロール', price: 230, order: 14, columnIndex: 19},
    {id: 'PRD015', name: 'ショコラロール', price: 280, order: 15, columnIndex: 20},
    {id: 'PRD016', name: '自家製クリームパン', price: 350, order: 16, columnIndex: 21},
    {id: 'PRD017', name: '自家製あんバター', price: 380, order: 17, columnIndex: 22},
    {id: 'PRD018', name: '抹茶&ホワイトチョコ', price: 400, order: 18, columnIndex: 23},
    {id: 'PRD019', name: '黒ごまパン', price: 200, order: 19, columnIndex: 24},
    {id: 'PRD020', name: 'レーズンジャムとクリームチーズのパン', price: 350, order: 20, columnIndex: 25},
    {id: 'PRD021', name: 'ピーナッツクリームパン', price: 350, order: 21, columnIndex: 26},
    {id: 'PRD022', name: 'あん食パン', price: 400, order: 22, columnIndex: 27},
    {id: 'PRD023', name: 'コーンパン', price: 400, order: 23, columnIndex: 28},
    {id: 'PRD024', name: 'レモンとクリームチーズのミニ食パン', price: 450, order: 24, columnIndex: 29},
    {id: 'PRD025', name: 'ピザ マルゲリータ', price: 1100, order: 25, columnIndex: 30},
    {id: 'PRD026', name: 'ピタパンサンド', price: 800, order: 26, columnIndex: 31},
    {id: 'PRD027', name: 'フォカッチャ', price: 300, order: 27, columnIndex: 32}
  ];
}

// ===== 値の存在チェック関数（null完全排除版） =====
/**
 * 値が存在するかチェックする（nullを使わない版）
 * @param {any} value - チェックする値
 * @returns {boolean} - 値が存在する場合true
 */
function isValuePresent(value) {
  return value !== undefined && value !== '' && value !== 0 && typeof value !== 'undefined';
}

/**
 * 値が空かどうかチェックする（nullを使わない版）
 * @param {any} value - チェックする値
 * @returns {boolean} - 値が空の場合true
 */
function isValueEmpty(value) {
  return value === undefined || value === '' || value === 0 || typeof value === 'undefined';
}

/**
 * 安全な文字列変換（nullを使わない版）
 * @param {any} value - 変換する値
 * @returns {string} - 文字列
 */
function safeStringConvert(value) {
  if (isValueEmpty(value)) {
    return '';
  }
  return value.toString();
}

// ===== 引渡状況の判定関数（null完全排除版） =====
/**
 * 引渡状況を厳密に判定する関数（null完全排除版）
 * @param {any} statusValue - スプレッドシートから取得した値
 * @returns {string} - 正規化された引渡状況（'未引渡', '引渡済', 'その他'）
 */
function normalizeDeliveryStatus(statusValue) {
  try {
    // 値が存在しない場合は「その他」を返す
    if (isValueEmpty(statusValue)) {
      return 'その他';
    }
    
    // 文字列に変換し、前後の空白を除去
    const statusStr = safeStringConvert(statusValue).trim();
    
    // 空文字列の場合は「その他」
    if (statusStr === '') {
      return 'その他';
    }
    
    // 全角英数字を半角に変換
    const normalizedStr = statusStr.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(s) {
      return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
    });
    
    // 「未引渡」のパターンをチェック（厳密に）
    const pendingPatterns = [
      '未引渡', '未引き渡し', '未引渡し', '未配達', '未完了',
      'みひきわたし', 'ミヒキワタシ', 'pending', 'PENDING'
    ];
    
    // 「引渡済」のパターンをチェック
    const deliveredPatterns = [
      '引渡済', '引き渡し済', '引渡し済', '配達済', '完了',
      'ひきわたしずみ', 'ヒキワタシズミ', 'delivered', 'DELIVERED', 'done', 'DONE'
    ];
    
    // 大文字小文字を無視して完全一致をチェック
    const statusLower = normalizedStr.toLowerCase();
    
    // まず未引渡をチェック
    for (const pattern of pendingPatterns) {
      if (statusLower === pattern.toLowerCase()) {
        return '未引渡';
      }
    }
    
    // 次に引渡済をチェック
    for (const pattern of deliveredPatterns) {
      if (statusLower === pattern.toLowerCase()) {
        return '引渡済';
      }
    }
    
    // どのパターンにも一致しない場合は「その他」
    return 'その他';
    
  } catch (error) {
    console.error(`引渡状況判定エラー: ${error.message}, 値: ${statusValue}`);
    return 'その他';
  }
}

// ===== メインエントリーポイント =====
function doGet(e) {
  try {
    console.log('🍞 Hyggelyシステム起動 v' + SYSTEM_CONFIG.version);
    const params = (e && e.parameter) ? e.parameter : {};
    const action = params.action || '';
    const password = params.password || '';
    
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
    if (!password || password !== SYSTEM_CONFIG.adminPassword) {
      console.log('⚠️ 認証失敗 - 無効なパスワード');
      return createAuthenticationPage();
    }
    
    console.log('✅ Dashboard認証成功');
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

// ===== データ取得エンジン（null完全排除版） =====
function getRawOrderData() {
  try {
    console.log('📋 生データ取得開始');
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const orderSheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.ORDER);
    
    if (!orderSheet) {
      console.error('❌ 予約管理票シートが見つかりません');
      return { success: false, error: 'シートが見つかりません', data: [] };
    }
    
    const lastRow = orderSheet.getLastRow();
    console.log('📋 最終行:', lastRow);
    
    if (lastRow <= 1) {
      console.log('📋 データが空です（ヘッダー行のみ）');
      return { success: true, data: [], totalRows: 1 };
    }
    
    const allData = orderSheet.getDataRange().getValues();
    console.log(`📋 生データ取得完了: ${allData.length}行 x ${(allData[0] && allData[0].length) || 0}列`);
    
    // デバッグ用: 先頭数行の36列目の値をチェック
    for (let i = 1; i <= Math.min(5, allData.length - 1); i++) {
      const row = allData[i];
      const status = row[35]; // 36列目（0ベースで35）
      console.log(`📋 行${i + 1}: AJ列(引渡状況) = "${status}" (型: ${typeof status})`);
    }
    
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

function parseOrderFromRowEnhanced(row, rowIndex, productMaster) {
  try {
    // 基本的なデータの存在チェック
    if (!row || row.length < 37) {
      console.log(`⚠️ 行${rowIndex}: データが不完全です (長さ: ${(row && row.length) || 0})`);
      return undefined;
    }
    
    // 最低限必要なデータがあるかチェック
    if (isValueEmpty(row[0]) || (isValueEmpty(row[1]) && isValueEmpty(row[2]))) {
      console.log(`⚠️ 行${rowIndex}: 必須データが不足しています`);
      return undefined;
    }
    
    // ★★★★★ 重要な修正: より確実な引渡状況判定（null完全排除版） ★★★★★
    const statusRaw = row[35]; // 36列目（0ベースで35）
    const deliveryStatus = normalizeDeliveryStatus(statusRaw);
    
    console.log(`📝 行${rowIndex}: 引渡状況 "${statusRaw}" → "${deliveryStatus}"`);

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
    
    let pickupDate = '';
    let pickupTime = '';
    
    if (isValuePresent(row[4])) {
      try {
        if (row[4] instanceof Date) {
          pickupDate = Utilities.formatDate(row[4], Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } else {
          const parsedDate = new Date(row[4]);
          if (!isNaN(parsedDate.getTime())) {
            pickupDate = Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          }
        }
      } catch (e) {
        pickupDate = safeStringConvert(row[4]);
      }
    }
    
    if (isValuePresent(row[5])) {
      pickupTime = safeStringConvert(row[5]).trim();
    }
    
    const orderData = {
      rowIndex: rowIndex,
      timestamp: row[0] || new Date(),
      lastName: safeStringConvert(row[1]).trim(),
      firstName: safeStringConvert(row[2]).trim(),
      email: safeStringConvert(row[3]).trim(),
      pickupDate: pickupDate,
      pickupTime: pickupTime,
      selectedProducts: selectedProducts,
      productsDisplay: selectedProducts.map(p => p.displayText).join('\n'),
      note: safeStringConvert(row[33]).trim(),
      totalPrice: parseFloat(row[34]) || totalCalculated,
      deliveryStatus: deliveryStatus,
      orderId: row[36] || generateOrderId(),
      isDelivered: deliveryStatus === '引渡済',
      _debug: {
        originalRowIndex: rowIndex,
        originalDeliveryStatus: row[35],
        parsedDeliveryStatus: deliveryStatus,
        statusRaw: statusRaw
      }
    };
    
    console.log(`✅ 行${rowIndex}解析完了: ${orderData.lastName} ${orderData.firstName} - ${orderData.deliveryStatus}`);
    return orderData;
    
  } catch (error) {
    console.error(`❌ 行${rowIndex}解析エラー:`, error);
    return undefined;
  }
}

// ===== 修正された予約リスト取得関数（null完全排除版） =====
/**
 * 未引渡の予約リストを取得する（null完全排除版）
 * @returns {Array} 予約リスト（常に配列を返す）
 */
function getOrderListEnhanced() {
  try {
    console.log('📊 予約一覧取得開始（null完全排除版 v' + SYSTEM_CONFIG.version + '）');
    
    const rawResult = getRawOrderData();
    if (!rawResult.success) {
      console.error('❌ 生データ取得失敗:', rawResult.error);
      logSystemEvent('ERROR', '生データ取得失敗', rawResult.error);
      return []; // 常に空配列を返す
    }
    
    const allData = rawResult.data;
    if (!allData || allData.length <= 1) {
      console.log('📊 データが空です（ヘッダー行のみまたはデータなし）');
      return []; // 空配列を返す
    }
    
    const productMaster = getDefaultProducts();
    const orders = [];
    const parseErrors = [];
    
    // 引渡状況の分布を追跡
    const statusDistribution = {
      '未引渡': 0,
      '引渡済': 0,
      'その他': 0
    };
    
    console.log(`📊 データ解析開始: ${allData.length - 1}行のデータを処理`);
    
    for (let i = 1; i < allData.length; i++) {
      try {
        const order = parseOrderFromRowEnhanced(allData[i], i + 1, productMaster);
        if (order && typeof order === 'object') {
          // 引渡状況の分布を記録
          const status = order.deliveryStatus;
          statusDistribution[status] = (statusDistribution[status] || 0) + 1;
          
          // ★★★ 重要: 「未引渡」のみを結果に追加 ★★★
          if (status === '未引渡') {
            orders.push(order);
            console.log(`📋 未引渡データ追加: ${order.lastName} ${order.firstName} (行${i + 1})`);
          } else {
            console.log(`📋 スキップ: ${order.lastName} ${order.firstName} - 状態: ${status} (行${i + 1})`);
          }
        } else {
          console.log(`⚠️ 行${i + 1}: パース結果が無効のためスキップ`);
        }
      } catch (e) {
        parseErrors.push({
          rowIndex: i + 1,
          error: e.message,
          rawData: allData[i]
        });
        console.error(`❌ 行 ${i + 1} の処理中にエラー: ${e.message}`);
        logSystemEvent('ERROR', '行処理エラー', `行番号: ${i + 1}, エラー: ${e.toString()}`);
      }
    }
    
    // ソート
    orders.sort((a, b) => {
      try {
        const dateA = new Date(a.pickupDate + ' ' + (a.pickupTime || '00:00'));
        const dateB = new Date(b.pickupDate + ' ' + (b.pickupTime || '00:00'));
        return dateA.getTime() - dateB.getTime();
      } catch (e) {
        return 0;
      }
    });
    
    console.log(`📊 予約一覧取得完了: ${orders.length}件（未引渡のみ）`);
    console.log('📊 引渡状況分布:', statusDistribution);
    
    if (parseErrors.length > 0) {
      console.warn(`⚠️ 解析エラーが${parseErrors.length}件発生:`, parseErrors);
    }
    
    // ★★★ 重要: 必ず配列を返す ★★★
    return orders;
    
  } catch (error) {
    console.error('❌ 予約一覧取得エラー:', error);
    logSystemEvent('ERROR', '予約一覧取得エラー', error.toString());
    return []; // エラーが発生した場合も空配列を返す
  }
}

// ===== 全ての予約データ取得（null完全排除版） =====
function getAllOrdersEnhanced() {
  try {
    console.log('📊 全予約データ取得開始');
    const rawResult = getRawOrderData();
    if (!rawResult.success) {
      console.error('❌ 全予約データ: 生データ取得失敗');
      return [];
    }
    
    const allData = rawResult.data;
    if (!allData || allData.length <= 1) {
      console.log('📊 全予約データ: データが空です');
      return [];
    }
    
    const productMaster = getDefaultProducts();
    const orders = [];
    
    for (let i = 1; i < allData.length; i++) {
        try {
            const order = parseOrderFromRowEnhanced(allData[i], i + 1, productMaster);
            if (order && typeof order === 'object') {
                orders.push(order);
            }
        } catch(e) {
            console.error(`❌ 全予約データ: 行 ${i + 1} の処理中にエラー: ${e.message}`);
            logSystemEvent('ERROR', '行処理エラー', `行番号: ${i + 1}, エラー: ${e.toString()}`);
        }
    }
    
    console.log(`📊 全予約データ取得完了: ${orders.length}件`);
    return orders;
    
  } catch (error) {
    console.error('❌ 全予約データ取得エラー:', error);
    return [];
  }
}

// ===== ダッシュボード統計データ取得（null完全排除版） =====
function getDashboardStats() {
  try {
    console.log('📊 統計データ取得開始（v' + SYSTEM_CONFIG.version + '）');
    const allOrders = getAllOrdersEnhanced();
    const pendingOrders = allOrders.filter(order => order.deliveryStatus === '未引渡');
    const deliveredOrders = allOrders.filter(order => order.deliveryStatus === '引渡済');
    
    console.log(`📊 統計ベースデータ: 全${allOrders.length}件, 未引渡${pendingOrders.length}件, 引渡済${deliveredOrders.length}件`);
    
    const inventory = getInventoryDataForForm();
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    const todayPendingOrders = pendingOrders.filter(order => {
      if (!order.pickupDate) return false;
      try {
        const pickupDate = new Date(order.pickupDate);
        pickupDate.setHours(0, 0, 0, 0);
        return pickupDate.getTime() === today.getTime();
      } catch (e) { return false; }
    });
    
    const outOfStock = inventory.filter(p => p.remaining <= 0);
    const lowStock = inventory.filter(p => p.remaining > 0 && p.remaining <= (p.minStock || 3));
    
    const currentMonth = today.getMonth();
    const currentYear = today.getFullYear();
    
    const monthDeliveredOrders = deliveredOrders.filter(order => {
      if (!order.timestamp) return false;
      try {
        const orderDate = new Date(order.timestamp);
        return orderDate.getMonth() === currentMonth && orderDate.getFullYear() === currentYear;
      } catch (e) { return false; }
    });
    const monthRevenue = monthDeliveredOrders.reduce((sum, order) => sum + (order.totalPrice || 0), 0);
    
    const todayDeliveredOrders = deliveredOrders.filter(order => {
      if (!order.timestamp) return false;
      try {
        const orderDate = new Date(order.timestamp);
        orderDate.setHours(0, 0, 0, 0);
        return orderDate.getTime() === today.getTime();
      } catch (e) { return false; }
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
      lastUpdate: new Date().toISOString()
    };
    
    console.log(`📊 統計データ取得完了。未引渡: ${stats.pendingOrdersCount}件`);
    return stats;
    
  } catch (error) {
    console.error('❌ 統計データ取得エラー:', error);
    logSystemEvent('ERROR', '統計データ取得エラー', error.toString());
    return { 
      error: error.toString(),
      // エラー時でもダッシュボードが表示できるよう基本値を設定
      todayOrdersCount: 0,
      pendingOrdersCount: 0,
      deliveredOrdersCount: 0,
      outOfStockCount: 0,
      lowStockCount: 0,
      totalProducts: 0,
      todayRevenue: 0,
      monthRevenue: 0,
      systemVersion: SYSTEM_CONFIG.version,
      lastUpdate: new Date().toISOString()
    };
  }
}

// ===== 修正された引渡状況変更関数（null完全排除版） =====
/**
 * 引渡状況を変更する（null完全排除版）
 * @param {string} orderId - 予約ID
 * @param {string} newStatus - 新しい引渡状況
 * @param {string} updatedBy - 更新者
 * @returns {Object} 実行結果
 */
function updateDeliveryStatusEnhanced(orderId, newStatus, updatedBy = 'ADMIN') {
  try {
    console.log(`🔄 引渡状況変更開始: ${orderId} → ${newStatus}`);
    
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const orderSheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.ORDER);
    
    if (!orderSheet) {
      throw new Error('予約管理票シートが見つかりません');
    }
    
    const allData = orderSheet.getDataRange().getValues();
    let targetRowIndex = -1;
    let targetOrderData = undefined;
    
    // 予約IDで該当行を検索
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][36] === orderId) {
        targetRowIndex = i + 1;
        targetOrderData = allData[i];
        break;
      }
    }
    
    if (targetRowIndex === -1) {
      throw new Error(`予約ID ${orderId} が見つかりません`);
    }
    
    // 現在の状況を記録
    const currentStatus = normalizeDeliveryStatus(targetOrderData[35]);
    
    // 引渡状況を更新（AJ列 = 36番目）
    orderSheet.getRange(targetRowIndex, 36).setValue(newStatus);
    
    // キャッシュクリアと在庫更新
    clearCache();
    updateInventoryFromOrders();
    
    // ログ記録
    logSystemEvent('INFO', '引渡状況変更', 
      `予約ID: ${orderId}, ${currentStatus} → ${newStatus}, 更新者: ${updatedBy}`);
    
    console.log(`✅ 引渡状況変更完了: ${orderId}`);
    
    return { 
      success: true, 
      message: `引渡状況を「${newStatus}」に変更しました`,
      orderId: orderId,
      previousStatus: currentStatus,
      newStatus: newStatus
    };
    
  } catch (error) {
    console.error('❌ 引渡状況変更エラー:', error);
    logSystemEvent('ERROR', '引渡状況変更エラー', 
      `予約ID: ${orderId}, エラー: ${error.toString()}`);
    return { 
      success: false, 
      message: `引渡状況の変更に失敗しました: ${error.message}`,
      orderId: orderId
    };
  }
}

// ===== 修正された一括引渡状況変更関数（null完全排除版） =====
/**
 * 一括で引渡状況を変更する（null完全排除版）
 * @param {Array} orderIds - 予約IDの配列
 * @param {string} newStatus - 新しい引渡状況
 * @param {string} updatedBy - 更新者
 * @returns {Object} 実行結果
 */
function bulkUpdateDeliveryStatusEnhanced(orderIds, newStatus, updatedBy = 'ADMIN') {
  try {
    console.log(`🔄 一括引渡状況変更開始: ${orderIds.length}件 → ${newStatus}`);
    
    let successCount = 0;
    let errorCount = 0;
    const results = [];
    
    orderIds.forEach(orderId => {
      const result = updateDeliveryStatusEnhanced(orderId, newStatus, updatedBy);
      results.push(result);
      if (result.success) {
        successCount++;
      } else {
        errorCount++;
      }
    });
    
    logSystemEvent('INFO', '一括引渡状況変更', 
      `対象: ${orderIds.length}件, 成功: ${successCount}件, 失敗: ${errorCount}件`);
    
    const message = errorCount === 0 ? 
      `一括変更完了: ${successCount}件すべて成功` :
      `一括変更完了: 成功${successCount}件, 失敗${errorCount}件`;
    
    console.log(`✅ 一括引渡状況変更完了: 成功${successCount}件, 失敗${errorCount}件`);
    
    return { 
      success: errorCount === 0, 
      message: message, 
      successCount: successCount, 
      errorCount: errorCount,
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

// ===== 修正された予約処理関数（null完全排除版） =====
function processOrder(formData) {
  try {
    console.log('🔄 予約処理開始（v' + SYSTEM_CONFIG.version + '）');
    
    if (!formData.lastName || !formData.firstName || !formData.email || !formData.pickupDate || !formData.pickupTime) {
      throw new Error('必須項目が入力されていません');
    }
    
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
          outOfStockItems.push(product.name);
          continue;
        }
        
        orderedItems.push({ name: product.name, quantity: quantity, price: product.price, subtotal: quantity * product.price });
        totalPrice += quantity * product.price;
      }
    }
    
    if (outOfStockItems.length > 0) throw new Error(`在庫不足: ${outOfStockItems.join(', ')}`);
    if (orderedItems.length === 0) throw new Error('商品を1つ以上選択してください');
    
    const orderId = generateOrderId();
    const orderSheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId).getSheetByName(SYSTEM_CONFIG.sheets.ORDER);
    
    const rowData = new Array(37).fill('');
    rowData[0] = new Date();
    rowData[1] = formData.lastName;
    rowData[2] = formData.firstName;
    rowData[3] = formData.email;
    rowData[4] = formData.pickupDate;
    rowData[5] = formData.pickupTime;
    
    for (let i = 0; i < products.length; i++) {
      rowData[6 + i] = parseInt(formData[`product_${i}`]) || 0;
    }
    
    // ★★★ 修正: その他ご要望が空欄の場合は「とくになし。」を設定 ★★★
    const noteValue = formData.note && formData.note.trim();
    rowData[33] = noteValue ? noteValue : 'とくになし。';
    rowData[34] = totalPrice;
    rowData[35] = '未引渡'; // 明示的に「未引渡」を設定
    rowData[36] = orderId;
    
    orderSheet.appendRow(rowData);
    
    clearCache();
    updateInventoryFromOrders();
    sendOrderEmails(formData, orderedItems, totalPrice, orderId);
    
    logSystemEvent('INFO', '新規予約', `顧客: ${formData.lastName}, 金額: ¥${totalPrice}, ID: ${orderId}`);
    
    return {
      success: true,
      message: '予約が完了しました',
      orderDetails: {
        orderId: orderId,
        name: `${formData.lastName} ${formData.firstName}`,
        pickupDate: formData.pickupDate,
        pickupTime: formData.pickupTime,
        totalPrice: totalPrice
      }
    };
    
  } catch (error) {
    console.error('❌ 予約処理エラー:', error);
    logSystemEvent('ERROR', '予約処理エラー', error.toString());
    return { success: false, message: error.message };
  }
}

// ===== 修正されたデバッグ関数（null完全排除版） =====
/**
 * 詳細なデバッグ情報を取得する（null完全排除版）
 * @returns {Object} デバッグ情報
 */
function debugOrderDataEnhanced() {
  try {
    console.log('🔍 null完全排除版デバッグモード開始（v' + SYSTEM_CONFIG.version + '）');
    const rawResult = getRawOrderData();
    if (!rawResult.success) return { error: rawResult.error };
    
    const data = rawResult.data;
    
    // 引渡状況の詳細分析
    const deliveryStatusAnalysis = {
      rawValues: {},
      normalizedValues: {},
      problemRows: []
    };
    
    for (let i = 1; i < data.length; i++) {
      const statusRaw = data[i][35];
      const statusNormalized = normalizeDeliveryStatus(statusRaw);
      
      // 生の値の集計
      const rawKey = isValueEmpty(statusRaw) ? '(空/未定義値)' : safeStringConvert(statusRaw);
      deliveryStatusAnalysis.rawValues[rawKey] = (deliveryStatusAnalysis.rawValues[rawKey] || 0) + 1;
      
      // 正規化後の値の集計
      deliveryStatusAnalysis.normalizedValues[statusNormalized] = 
        (deliveryStatusAnalysis.normalizedValues[statusNormalized] || 0) + 1;
      
      // 問題のある行を記録
      if (statusNormalized === 'その他' && isValuePresent(statusRaw)) {
        deliveryStatusAnalysis.problemRows.push({
          rowIndex: i + 1,
          rawValue: statusRaw,
          dataType: typeof statusRaw,
          stringValue: safeStringConvert(statusRaw),
          customer: `${safeStringConvert(data[i][1])} ${safeStringConvert(data[i][2])}`.trim()
        });
      }
    }
    
    // 処理テスト
    const testOrders = getOrderListEnhanced();
    const testStats = getDashboardStats();
    
    // スプレッドシートの基本情報
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const orderSheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.ORDER);
    
    return {
      systemVersion: SYSTEM_CONFIG.version,
      debugTimestamp: new Date().toISOString(),
      spreadsheetInfo: {
        name: spreadsheet.getName(),
        url: spreadsheet.getUrl(),
        orderSheetExists: !!orderSheet,
        totalRows: data.length,
        lastColumn: orderSheet ? orderSheet.getLastColumn() : 0,
        headerRow: data[0] || []
      },
      deliveryStatusAnalysis: deliveryStatusAnalysis,
      processingResults: {
        ordersListType: Array.isArray(testOrders) ? 'array' : 'object',
        ordersListLength: Array.isArray(testOrders) ? testOrders.length : 'N/A',
        ordersListError: Array.isArray(testOrders) ? undefined : testOrders.error,
        statsHasError: !!testStats.error,
        statsPendingCount: testStats.pendingOrdersCount
      },
      sampleData: data.slice(0, Math.min(3, data.length)).map((row, index) => ({
        rowIndex: index + 1,
        isHeader: index === 0,
        firstName: safeStringConvert(row[2]),
        lastName: safeStringConvert(row[1]),
        deliveryStatusRaw: row[35],
        deliveryStatusNormalized: index === 0 ? 'N/A' : normalizeDeliveryStatus(row[35]),
        orderId: safeStringConvert(row[36])
      }))
    };
    
  } catch (error) {
    console.error('❌ null完全排除版デバッグ確認エラー:', error);
    return { 
      error: error.toString(),
      systemVersion: SYSTEM_CONFIG.version,
      debugTimestamp: new Date().toISOString()
    };
  }
}

// ===== その他の関数群（null完全排除版に修正） =====
function getInventoryDataForForm() {
  try {
    if (SYSTEM_CONFIG.cache.enabled) {
      const cache = CacheService.getScriptCache();
      const cachedInventory = cache.get(SYSTEM_CONFIG.cache.keys.inventory);
      if (cachedInventory) return JSON.parse(cachedInventory);
    }
    
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const inventorySheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.INVENTORY);
    const products = getProductMaster().filter(p => p.enabled);
    
    if (!inventorySheet || inventorySheet.getLastRow() <= 1) {
      return products.map(p => ({ id: p.id, name: p.name, price: p.price, stock: 10, reserved: 0, remaining: 10, minStock: 3 }));
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
      CacheService.getScriptCache().put(SYSTEM_CONFIG.cache.keys.inventory, JSON.stringify(sortedInventory), SYSTEM_CONFIG.cache.duration);
    }
    
    return sortedInventory;
    
  } catch (error) {
    console.error('❌ 在庫データ取得エラー:', error);
    return getDefaultProducts().map(p => ({ id: p.id, name: p.name, price: p.price, stock: 10, reserved: 0, remaining: 10, minStock: 3 }));
  }
}

function getProductMaster() {
  try {
    if (SYSTEM_CONFIG.cache.enabled) {
      const cache = CacheService.getScriptCache();
      const cachedProducts = cache.get(SYSTEM_CONFIG.cache.keys.products);
      if (cachedProducts) return JSON.parse(cachedProducts);
    }
    
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const masterSheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.PRODUCT_MASTER);
    
    if (!masterSheet || masterSheet.getLastRow() <= 1) {
      return getDefaultProducts().map(p => ({ ...p, enabled: true }));
    }
    
    const data = masterSheet.getDataRange().getValues();
    const products = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      products.push({ id: row[0], name: row[1], price: row[2], order: row[3], enabled: row[4] !== false });
    }
    
    const sortedProducts = products.sort((a, b) => a.order - b.order);
    
    if (SYSTEM_CONFIG.cache.enabled) {
      CacheService.getScriptCache().put(SYSTEM_CONFIG.cache.keys.products, JSON.stringify(sortedProducts), SYSTEM_CONFIG.cache.duration);
    }
    
    return sortedProducts;
  } catch (error) {
    console.error('❌ 商品マスタ取得エラー:', error);
    return getDefaultProducts().map(p => ({ ...p, enabled: true }));
  }
}

function updateInventoryFromOrders() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const orderSheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.ORDER);
    const inventorySheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.INVENTORY);
    
    if (!orderSheet || !inventorySheet || orderSheet.getLastRow() <= 1) return;
    
    const productMaster = getDefaultProducts();
    const reservations = {};
    productMaster.forEach(p => reservations[p.id] = 0);
    
    const orderData = orderSheet.getDataRange().getValues();
    for (let i = 1; i < orderData.length; i++) {
      const row = orderData[i];
      // ★★★★★ 修正: null完全排除版引渡状況判定を使用 ★★★★★
      if (normalizeDeliveryStatus(row[35]) === '未引渡') {
        productMaster.forEach(product => {
          const quantity = parseInt(row[product.columnIndex]) || 0;
          if (quantity > 0) reservations[product.id] += quantity;
        });
      }
    }
    
    const inventoryData = inventorySheet.getDataRange().getValues();
    for (let i = 1; i < inventoryData.length; i++) {
      const productId = inventoryData[i][0];
      if (reservations[productId] !== undefined) {
        inventorySheet.getRange(i + 1, 5).setValue(reservations[productId]);
      }
    }
    
    clearCache();
  } catch (error) {
    console.error('❌ 在庫更新エラー:', error);
  }
}

// ===== ユーティリティ & ヘルパー関数群（変更なし） =====
function generateOrderId() {
  const now = new Date();
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyMMdd');
  const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HHmmss');
  const randomPart = Math.floor(Math.random() * 100).toString().padStart(2, '0');
  return `ORD${dateStr}${timeStr}${randomPart}`;
}

function logSystemEvent(level, event, details, user = 'SYSTEM') {
  try {
    const logSheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId).getSheetByName(SYSTEM_CONFIG.sheets.SYSTEM_LOG);
    if (logSheet) {
      logSheet.appendRow([new Date(), level, event, details, user]);
      if (logSheet.getLastRow() > 1000) {
        logSheet.deleteRows(2, 100);
      }
    }
  } catch (e) { console.error('❌ ログ記録エラー:', e); }
}

function clearCache() {
  if (SYSTEM_CONFIG.cache.enabled) {
    try {
      CacheService.getScriptCache().removeAll(Object.values(SYSTEM_CONFIG.cache.keys));
      console.log('✅ キャッシュクリア完了');
      return { success: true, message: 'キャッシュをクリアしました' };
    } catch (error) {
      console.error('❌ キャッシュクリアエラー:', error);
      return { success: false, message: 'キャッシュクリアに失敗しました' };
    }
  }
  return { success: true, message: 'キャッシュは無効です' };
}

function sendOrderEmails(formData, orderedItems, totalPrice, orderId) {
  if (!SYSTEM_CONFIG.email.enabled) return;
  
  const itemsText = orderedItems.map(item => `・${item.name}　${item.quantity}個　¥${item.subtotal.toLocaleString()}`).join('\n');
  const pickupDateTime = `${formData.pickupDate} ${formData.pickupTime}`;
  const commonVars = { lastName: formData.lastName, firstName: formData.firstName, orderId, orderItems: itemsText, totalPrice: totalPrice.toLocaleString(), pickupDateTime, email: formData.email };
  
  try {
    const customerSubject = replaceEmailVariables(SYSTEM_CONFIG.email.customerSubject, commonVars);
    const customerBody = replaceEmailVariables(SYSTEM_CONFIG.email.customerBody, commonVars);
    GmailApp.sendEmail(formData.email, customerSubject, customerBody);
  } catch (e) { console.error('❌ 顧客メール送信エラー:', e); }
  
  try {
    const adminSubject = replaceEmailVariables(SYSTEM_CONFIG.email.adminSubject, commonVars);
    const adminBody = replaceEmailVariables(SYSTEM_CONFIG.email.adminBody, commonVars);
    GmailApp.sendEmail(SYSTEM_CONFIG.email.adminEmail, adminSubject, adminBody);
  } catch (e) { console.error('❌ 管理者メール送信エラー:', e); }
}

function replaceEmailVariables(template, variables) {
  let result = template;
  for (const key in variables) {
    result = result.replace(new RegExp(`{${key}}`, 'g'), variables[key] || '');
  }
  return result;
}

function testConnection() {
  return { success: true, timestamp: new Date().toISOString(), version: SYSTEM_CONFIG.version, status: 'operational' };
}

function handleOrderForm() {
  try {
    return HtmlService.createHtmlOutputFromFile('OrderForm').setTitle('Hyggelyカンパーニュ専門店 ご予約フォーム').addMetaTag('viewport', 'width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (e) { return createErrorPage('予約フォームの読み込みに失敗しました', e.message); }
}

function handleHealthCheck() {
  return HtmlService.createHtmlOutput(JSON.stringify({ status: 'ok', version: SYSTEM_CONFIG.version })).setMimeType(ContentService.MimeType.JSON);
}

function createAuthenticationPage() {
  const html = `<!DOCTYPE html><html><head><title>管理者認証</title><meta name="viewport" content="width=device-width, initial-scale=1"><link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet"></head><body class="d-flex align-items-center py-4 bg-body-tertiary"><main class="form-signin w-100 m-auto" style="max-width: 330px; padding: 1rem;"><form id="auth-form"><h1 class="h3 mb-3 fw-normal">管理者認証</h1><div class="form-floating"><input type="password" class="form-control" id="password" placeholder="Password"><label for="password">パスワード</label></div><button class="btn btn-primary w-100 py-2 mt-3" type="submit">ログイン</button></form></main><script>document.getElementById('auth-form').addEventListener('submit', function(e) { e.preventDefault(); window.location.href = '?action=dashboard&password=' + encodeURIComponent(document.getElementById('password').value); });</script></body></html>`;
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function createFallbackDashboard() {
    return createErrorPage('ダッシュボードエラー', 'ダッシュボードの読み込みに失敗しました。ファイルが破損している可能性があります。');
}

function createErrorPage(title, message) {
  const html = `<!DOCTYPE html><html lang="ja"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>エラー</title></head><body><h1>⚠️ ${title}</h1><p>${message}</p></body></html>`;
  return HtmlService.createHtmlOutput(html);
}

// ===== データ構造チェック関数（null完全排除版） =====
function checkDataStructure() {
  try {
    console.log('🔍 データ構造チェック開始');
    
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const orderSheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.ORDER);
    
    if (!orderSheet) {
      return { error: '予約管理票シートが見つかりません' };
    }
    
    const lastRow = orderSheet.getLastRow();
    const lastColumn = orderSheet.getLastColumn();
    
    if (lastRow <= 1) {
      return {
        message: 'データが存在しません（ヘッダー行のみ）',
        totalRows: lastRow,
        totalColumns: lastColumn
      };
    }
    
    // 全データを取得
    const allData = orderSheet.getDataRange().getValues();
    const headerRow = allData[0];
    
    // 各列のデータ型と内容を分析
    const columnAnalysis = {};
    for (let col = 0; col < headerRow.length; col++) {
      const columnName = headerRow[col] || `列${col + 1}`;
      const columnData = [];
      const dataTypes = new Set();
      
      for (let row = 1; row < Math.min(6, allData.length); row++) { // 最大5行をサンプル
        const cellValue = allData[row][col];
        columnData.push({
          rowIndex: row + 1,
          value: cellValue,
          type: typeof cellValue,
          stringValue: (cellValue && cellValue.toString) ? cellValue.toString() : ''
        });
        dataTypes.add(typeof cellValue);
      }
      
      columnAnalysis[`${col}_${columnName}`] = {
        columnIndex: col,
        columnName: columnName,
        dataTypes: Array.from(dataTypes),
        sampleData: columnData
      };
    }
    
    // 特に重要な列をチェック
    const importantColumns = {
      deliveryStatus: { index: 35, name: '引渡状況' },
      orderId: { index: 36, name: '予約ID' },
      lastName: { index: 1, name: '姓' },
      firstName: { index: 2, name: '名' },
      email: { index: 3, name: 'メール' },
      pickupDate: { index: 4, name: '受取日' },
      pickupTime: { index: 5, name: '受取時間' }
    };
    
    const importantColumnAnalysis = {};
    for (const [key, info] of Object.entries(importantColumns)) {
      if (allData[0][info.index]) {
        const columnData = [];
        for (let row = 1; row < Math.min(6, allData.length); row++) {
          const cellValue = allData[row][info.index];
          columnData.push({
            rowIndex: row + 1,
            value: cellValue,
            type: typeof cellValue,
            stringValue: (cellValue && cellValue.toString) ? cellValue.toString() : '',
            normalized: key === 'deliveryStatus' ? normalizeDeliveryStatus(cellValue) : 'N/A'
          });
        }
        importantColumnAnalysis[key] = {
          columnIndex: info.index,
          expectedName: info.name,
          actualHeader: allData[0][info.index],
          sampleData: columnData
        };
      } else {
        importantColumnAnalysis[key] = {
          columnIndex: info.index,
          expectedName: info.name,
          error: '列が存在しません'
        };
      }
    }
    
    return {
      checkTimestamp: new Date().toISOString(),
      spreadsheetInfo: {
        name: spreadsheet.getName(),
        url: spreadsheet.getUrl(),
        totalRows: lastRow,
        totalColumns: lastColumn,
        dataRows: lastRow - 1
      },
      headerAnalysis: {
        headers: headerRow,
        totalColumns: headerRow.length
      },
      importantColumns: importantColumnAnalysis,
      deliveryStatusCheck: {
        columnExists: !!allData[0][35],
        columnHeader: allData[0][35],
        sampleValues: allData.slice(1, 6).map((row, index) => ({
          rowIndex: index + 2,
          raw: row[35],
          normalized: normalizeDeliveryStatus(row[35]),
          customer: `${safeStringConvert(row[1])} ${safeStringConvert(row[2])}`.trim()
        }))
      },
      recommendations: generateRecommendations(importantColumnAnalysis, allData)
    };
    
  } catch (error) {
    console.error('❌ データ構造チェックエラー:', error);
    return { 
      error: error.toString(),
      checkTimestamp: new Date().toISOString()
    };
  }
}

/**
 * データ構造チェック結果に基づく推奨事項を生成
 */
function generateRecommendations(columnAnalysis, allData) {
  const recommendations = [];
  
  // 引渡状況列のチェック
  if (columnAnalysis.deliveryStatus && !columnAnalysis.deliveryStatus.error) {
    const deliveryData = columnAnalysis.deliveryStatus.sampleData;
    const hasProblems = deliveryData.some(data => data.normalized === 'その他' && isValuePresent(data.value));
    
    if (hasProblems) {
      recommendations.push({
        type: 'warning',
        message: '引渡状況列に認識できない値があります。「未引渡」または「引渡済」を正確に入力してください。'
      });
    }
  }
  
  // データ行数のチェック
  if (allData.length <= 1) {
    recommendations.push({
      type: 'info',
      message: 'データが登録されていません。予約フォームからテストデータを登録してみてください。'
    });
  } else if (allData.length <= 5) {
    recommendations.push({
      type: 'info',
      message: 'データ数が少ないようです。より多くのデータで動作確認することをお勧めします。'
    });
  }
  
  return recommendations;
}

// ===== シート初期化関数群（変更なし） =====
function checkAndInitializeSystem() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const requiredSheets = Object.values(SYSTEM_CONFIG.sheets);
    const existingSheets = spreadsheet.getSheets().map(s => s.getName());
    
    requiredSheets.forEach(sheetName => {
      if (!existingSheets.includes(sheetName)) {
        console.log(`🔧 シート[${sheetName}]が存在しないため作成します。`);
        initializeSheet(spreadsheet, sheetName);
      }
    });
    return { success: true };
  } catch (e) {
    console.error('❌ システム初期化中の致命的なエラー:', e);
    return { success: false, error: e.toString() };
  }
}

function initializeSheet(spreadsheet, sheetName) {
  const sheet = spreadsheet.insertSheet(sheetName);
  switch (sheetName) {
    case SYSTEM_CONFIG.sheets.ORDER: initOrderSheet(sheet); break;
    case SYSTEM_CONFIG.sheets.PRODUCT_MASTER: initProductMaster(sheet); break;
    case SYSTEM_CONFIG.sheets.INVENTORY: initInventory(sheet); break;
    case SYSTEM_CONFIG.sheets.SYSTEM_LOG: initSystemLog(sheet); break;
  }
}

function initOrderSheet(sheet) {
  const headers = ['タイムスタンプ', '姓', '名', 'メール', '受取日', '受取時間', ...getDefaultProducts().map(p => p.name), 'その他ご要望', '合計金額', '引渡済', '予約ID'];
  sheet.appendRow(headers);
  sheet.setFrozenRows(1);
}

function initProductMaster(sheet) {
  const headers = ['商品ID', '商品名', '価格', '表示順', '有効フラグ'];
  sheet.appendRow(headers);
  const productData = getDefaultProducts().map(p => [p.id, p.name, p.price, p.order, true]);
  sheet.getRange(2, 1, productData.length, productData[0].length).setValues(productData);
  sheet.setFrozenRows(1);
}

function initInventory(sheet) {
  const headers = ['商品ID', '商品名', '単価', '在庫数', '予約数', '残数', '最低在庫数', '最終更新日'];
  sheet.appendRow(headers);
  const inventoryData = getDefaultProducts().map(p => [p.id, p.name, p.price, 10, 0, 10, 3, new Date()]);
  sheet.getRange(2, 1, inventoryData.length, inventoryData[0].length).setValues(inventoryData);
  sheet.setFrozenRows(1);
}

function initSystemLog(sheet) {
  const headers = ['タイムスタンプ', 'レベル', 'イベント', '詳細', 'ユーザー'];
  sheet.appendRow(headers);
  sheet.setFrozenRows(1);
}

// ===== 後方互換性のための関数エイリアス =====
function getOrderList() {
  return getOrderListEnhanced();
}

function getAllOrders() {
  return getAllOrdersEnhanced();
}

function updateDeliveryStatus(orderId, newStatus, updatedBy = 'ADMIN') {
  return updateDeliveryStatusEnhanced(orderId, newStatus, updatedBy);
}

function bulkUpdateDeliveryStatus(orderIds, newStatus, updatedBy = 'ADMIN') {
  return bulkUpdateDeliveryStatusEnhanced(orderIds, newStatus, updatedBy);
}

function debugOrderData() {
  return debugOrderDataEnhanced();
}