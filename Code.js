/**
 * Hyggelyカンパーニュ専門店 予約管理システム - 修正版
 * v5.2.3 - エラー解決・安定動作版
 * 
 * 修正内容:
 * - OrderForm.html依存を削除し、Code.js内で直接HTML生成
 * - CSP/セキュリティヘッダーをGAS環境に最適化
 * - エラーハンドリングを強化
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
  version: '5.2.3'
};

// ===== メインエントリーポイント =====
function doGet(e) {
  try {
    console.log('🍞 Hyggelyシステム起動 v' + SYSTEM_CONFIG.version);
    
    const params = e?.parameter || {};
    const action = params.action || '';
    const password = params.password || '';
    
    checkAndInitializeSystem();
    
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

// 🔧 修正版：OrderForm.htmlに依存しない予約フォーム生成
function handleOrderForm() {
  try {
    console.log('📝 予約フォーム生成開始');
    
    // 商品データと在庫データを取得
    const products = getProductMaster().filter(p => p.enabled);
    const inventory = getInventoryDataForForm();
    
    // 商品選択HTML生成
    let productsHtml = '';
    products.forEach((product, index) => {
      const inventoryItem = inventory.find(inv => inv.id === product.id) || 
                           { remaining: 0, stock: 0, reserved: 0 };
      
      const isAvailable = inventoryItem.remaining > 0;
      const stockStatus = isAvailable ? 
        `在庫：${inventoryItem.remaining}個` : 
        '在庫切れ';
      
      productsHtml += `
        <div class="col-md-6 mb-3">
          <div class="card product-card ${!isAvailable ? 'out-of-stock' : ''}">
            <div class="card-body">
              <h6 class="product-name">${product.name}</h6>
              <div class="d-flex justify-content-between align-items-center mb-2">
                <span class="price">¥${product.price.toLocaleString()}</span>
                <span class="stock-status ${isAvailable ? 'text-success' : 'text-danger'}">
                  ${stockStatus}
                </span>
              </div>
              <div class="quantity-selector">
                <label class="form-label">数量</label>
                <select class="form-select product-quantity" 
                        data-product-index="${index}"
                        data-price="${product.price}"
                        ${!isAvailable ? 'disabled' : ''}>
                  <option value="0">0個</option>
                  ${Array.from({length: Math.min(10, inventoryItem.remaining)}, (_, i) => 
                    `<option value="${i + 1}">${i + 1}個</option>`
                  ).join('')}
                </select>
              </div>
            </div>
          </div>
        </div>
      `;
    });

    const orderFormHtml = `
<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Hyggelyカンパーニュ専門店 ご予約フォーム</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.0/font/bootstrap-icons.css">
  <style>
    body {
      background: linear-gradient(135deg, #f8f6f0 0%, #f0ede5 100%);
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif;
      min-height: 100vh;
    }
    
    .container {
      max-width: 800px;
      margin: 0 auto;
      padding: 20px;
    }
    
    .header {
      text-align: center;
      margin-bottom: 2rem;
      padding: 2rem;
      background: linear-gradient(135deg, #8B4513 0%, #A0522D 100%);
      color: white;
      border-radius: 15px;
      box-shadow: 0 8px 30px rgba(139, 69, 19, 0.3);
    }
    
    .card {
      border: none;
      border-radius: 12px;
      box-shadow: 0 4px 20px rgba(0, 0, 0, 0.08);
      margin-bottom: 1.5rem;
      background: rgba(255, 255, 255, 0.95);
    }
    
    .card-header {
      background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
      border-bottom: 2px solid rgba(139, 69, 19, 0.1);
      font-weight: 700;
      padding: 1rem 1.5rem;
      border-radius: 12px 12px 0 0;
    }
    
    .product-card {
      transition: all 0.3s ease;
      height: 100%;
    }
    
    .product-card:hover:not(.out-of-stock) {
      transform: translateY(-2px);
      box-shadow: 0 8px 25px rgba(0, 0, 0, 0.15);
    }
    
    .product-card.out-of-stock {
      opacity: 0.6;
      background: #f8f9fa;
    }
    
    .product-name {
      color: #8B4513;
      font-weight: 600;
      margin-bottom: 0.5rem;
    }
    
    .price {
      font-size: 1.1rem;
      font-weight: bold;
      color: #d4a574;
    }
    
    .stock-status {
      font-size: 0.9rem;
      font-weight: 600;
    }
    
    .form-control, .form-select {
      border-radius: 8px;
      border: 2px solid #e9ecef;
      transition: all 0.3s ease;
    }
    
    .form-control:focus, .form-select:focus {
      border-color: #8B4513;
      box-shadow: 0 0 0 0.2rem rgba(139, 69, 19, 0.15);
    }
    
    .btn-primary {
      background: linear-gradient(135deg, #8B4513 0%, #A0522D 100%);
      border: none;
      border-radius: 8px;
      font-weight: 600;
      padding: 12px 30px;
      transition: all 0.3s ease;
    }
    
    .btn-primary:hover {
      transform: translateY(-2px);
      box-shadow: 0 8px 25px rgba(139, 69, 19, 0.3);
    }
    
    .total-section {
      background: linear-gradient(135deg, #fff3cd 0%, #ffeaa7 100%);
      border: 2px solid #d4a574;
      border-radius: 12px;
      padding: 1.5rem;
      margin: 2rem 0;
    }
    
    .total-price {
      font-size: 1.5rem;
      font-weight: bold;
      color: #8B4513;
    }
    
    .loading {
      display: none;
      text-align: center;
      padding: 2rem;
    }
    
    .loading.show { display: block; }
    
    .alert {
      border-radius: 8px;
      margin-bottom: 1rem;
    }
    
    @media (max-width: 768px) {
      .container { padding: 10px; }
      .header { padding: 1rem; }
      .card-body { padding: 1rem; }
    }
  </style>
</head>
<body>
  <div class="container">
    <!-- ヘッダー -->
    <div class="header">
      <h1><i class="bi bi-shop me-2"></i>Hyggelyカンパーニュ専門店</h1>
      <p class="mb-0">美味しいパンのご予約はこちらから</p>
    </div>

    <!-- アラート表示エリア -->
    <div id="alert-container"></div>

    <!-- 予約フォーム -->
    <form id="order-form" onsubmit="submitOrder(event)">
      <!-- 基本情報 -->
      <div class="card">
        <div class="card-header">
          <i class="bi bi-person me-2"></i>お客様情報
        </div>
        <div class="card-body">
          <div class="row">
            <div class="col-md-6 mb-3">
              <label class="form-label">姓 <span class="text-danger">*</span></label>
              <input type="text" class="form-control" id="lastName" required>
            </div>
            <div class="col-md-6 mb-3">
              <label class="form-label">名 <span class="text-danger">*</span></label>
              <input type="text" class="form-control" id="firstName" required>
            </div>
            <div class="col-12 mb-3">
              <label class="form-label">メールアドレス <span class="text-danger">*</span></label>
              <input type="email" class="form-control" id="email" required>
            </div>
          </div>
        </div>
      </div>

      <!-- 受取情報 -->
      <div class="card">
        <div class="card-header">
          <i class="bi bi-calendar-check me-2"></i>受取情報
        </div>
        <div class="card-body">
          <div class="row">
            <div class="col-md-6 mb-3">
              <label class="form-label">受取日 <span class="text-danger">*</span></label>
              <input type="date" class="form-control" id="pickupDate" required>
            </div>
            <div class="col-md-6 mb-3">
              <label class="form-label">受取時間 <span class="text-danger">*</span></label>
              <select class="form-select" id="pickupTime" required>
                <option value="">時間を選択してください</option>
                <option value="09:00">09:00</option>
                <option value="09:30">09:30</option>
                <option value="10:00">10:00</option>
                <option value="10:30">10:30</option>
                <option value="11:00">11:00</option>
                <option value="11:30">11:30</option>
                <option value="12:00">12:00</option>
                <option value="12:30">12:30</option>
                <option value="13:00">13:00</option>
                <option value="13:30">13:30</option>
                <option value="14:00">14:00</option>
                <option value="14:30">14:30</option>
                <option value="15:00">15:00</option>
                <option value="15:30">15:30</option>
                <option value="16:00">16:00</option>
                <option value="16:30">16:30</option>
                <option value="17:00">17:00</option>
              </select>
            </div>
          </div>
        </div>
      </div>

      <!-- 商品選択 -->
      <div class="card">
        <div class="card-header">
          <i class="bi bi-basket me-2"></i>商品選択
        </div>
        <div class="card-body">
          <div class="row" id="products-container">
            ${productsHtml}
          </div>
        </div>
      </div>

      <!-- 合計金額 -->
      <div class="total-section">
        <div class="d-flex justify-content-between align-items-center">
          <span class="fs-5 fw-semibold">合計金額</span>
          <span class="total-price" id="total-price">¥0</span>
        </div>
      </div>

      <!-- その他ご要望 -->
      <div class="card">
        <div class="card-header">
          <i class="bi bi-chat-text me-2"></i>その他ご要望
        </div>
        <div class="card-body">
          <textarea class="form-control" id="note" rows="3" 
                    placeholder="アレルギーや特別なご要望がございましたらご記入ください"></textarea>
        </div>
      </div>

      <!-- 送信ボタン -->
      <div class="text-center mt-4">
        <button type="submit" class="btn btn-primary btn-lg" id="submit-btn">
          <i class="bi bi-check-circle me-2"></i>予約を確定する
        </button>
      </div>

      <!-- ローディング -->
      <div class="loading" id="loading">
        <div class="spinner-border" style="color: #8B4513;"></div>
        <p class="mt-3">予約を処理しています...</p>
      </div>
    </form>
  </div>

  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>
  <script>
    let currentProducts = ${JSON.stringify(products)};
    let currentInventory = ${JSON.stringify(inventory)};

    // 初期化
    document.addEventListener('DOMContentLoaded', function() {
      console.log('🚀 予約フォーム初期化');
      
      // 最小受取日を今日+1日に設定
      const tomorrow = new Date();
      tomorrow.setDate(tomorrow.getDate() + 1);
      document.getElementById('pickupDate').min = tomorrow.toISOString().split('T')[0];
      
      // 商品数量変更イベント
      document.querySelectorAll('.product-quantity').forEach(select => {
        select.addEventListener('change', updateTotal);
      });
      
      updateTotal();
    });

    // 合計金額計算
    function updateTotal() {
      let total = 0;
      
      document.querySelectorAll('.product-quantity').forEach(select => {
        const quantity = parseInt(select.value) || 0;
        const price = parseFloat(select.dataset.price) || 0;
        total += quantity * price;
      });
      
      document.getElementById('total-price').textContent = '¥' + total.toLocaleString();
    }

    // 予約送信
    function submitOrder(event) {
      event.preventDefault();
      
      const formData = {
        lastName: document.getElementById('lastName').value.trim(),
        firstName: document.getElementById('firstName').value.trim(),
        email: document.getElementById('email').value.trim(),
        pickupDate: document.getElementById('pickupDate').value,
        pickupTime: document.getElementById('pickupTime').value,
        note: document.getElementById('note').value.trim()
      };

      // 商品データ追加
      document.querySelectorAll('.product-quantity').forEach((select, index) => {
        formData[\`product_\${index}\`] = select.value;
      });

      // バリデーション
      if (!formData.lastName || !formData.firstName || !formData.email || 
          !formData.pickupDate || !formData.pickupTime) {
        showAlert('必須項目をすべて入力してください', 'danger');
        return;
      }

      // 商品選択チェック
      const hasProducts = Object.keys(formData)
        .filter(key => key.startsWith('product_'))
        .some(key => parseInt(formData[key]) > 0);

      if (!hasProducts) {
        showAlert('商品を1つ以上選択してください', 'warning');
        return;
      }

      // 送信処理
      showLoading(true);
      
      google.script.run
        .withSuccessHandler(handleOrderSuccess)
        .withFailureHandler(handleOrderError)
        .processOrder(formData);
    }

    // 成功時の処理
    function handleOrderSuccess(result) {
      showLoading(false);
      
      if (result.success) {
        showAlert(\`予約が完了しました！<br>予約ID: <strong>\${result.orderDetails.orderId}</strong><br>確認メールを送信いたします。\`, 'success');
        document.getElementById('order-form').reset();
        updateTotal();
        
        // 5秒後にページをリロード
        setTimeout(() => {
          location.reload();
        }, 5000);
      } else {
        showAlert(result.message, 'danger');
      }
    }

    // エラー時の処理
    function handleOrderError(error) {
      showLoading(false);
      console.error('予約エラー:', error);
      showAlert('予約の処理中にエラーが発生しました。もう一度お試しください。', 'danger');
    }

    // アラート表示
    function showAlert(message, type) {
      const container = document.getElementById('alert-container');
      const alertClass = type === 'danger' ? 'alert-danger' : 
                        type === 'warning' ? 'alert-warning' : 
                        type === 'success' ? 'alert-success' : 'alert-info';
      
      container.innerHTML = \`
        <div class="alert \${alertClass} alert-dismissible fade show">
          \${message}
          <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
        </div>
      \`;
      
      // 画面上部にスクロール
      container.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }

    // ローディング表示制御
    function showLoading(show) {
      const loading = document.getElementById('loading');
      const submitBtn = document.getElementById('submit-btn');
      
      if (show) {
        loading.classList.add('show');
        submitBtn.disabled = true;
        submitBtn.innerHTML = '<span class="spinner-border spinner-border-sm me-2"></span>処理中...';
      } else {
        loading.classList.remove('show');
        submitBtn.disabled = false;
        submitBtn.innerHTML = '<i class="bi bi-check-circle me-2"></i>予約を確定する';
      }
    }
  </script>
</body>
</html>
    `;

    console.log('✅ 予約フォーム生成完了');
    
    return HtmlService.createHtmlOutput(orderFormHtml)
      .setTitle('Hyggelyカンパーニュ専門店 ご予約フォーム')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
      
  } catch (error) {
    console.error('❌ 予約フォーム生成エラー:', error);
    return createErrorPage('予約フォームの生成に失敗しました', error.toString());
  }
}

// 🔧 修正版：Dashboard（パスワード認証対応）
function handleDashboard(password) {
  if (password !== SYSTEM_CONFIG.adminPassword) {
    return createRedirectPage('認証失敗', '?');
  }
  
  try {
    const htmlOutput = HtmlService.createHtmlOutputFromFile('Dashboard')
      .setTitle('Hyggelyカンパーニュ専門店 管理者ダッシュボード')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
      
    return htmlOutput;
  } catch (error) {
    return createErrorPage('ダッシュボードの読み込みに失敗しました', error.toString());
  }
}

// 🔧 修正版：メール設定
function handleEmailSettings(password) {
  if (password !== SYSTEM_CONFIG.adminPassword) {
    return createRedirectPage('認証失敗', '?');
  }
  
  try {
    // EmailSettings.htmlが存在しない場合は簡易版を生成
    const emailSettingsHtml = `
<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>メール設定 - Hyggelyカンパーニュ専門店</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body>
  <div class="container py-4">
    <h1>メール設定</h1>
    <div class="alert alert-info">
      メール設定機能は開発中です。<br>
      現在はデフォルト設定（hyggely2021@gmail.com）が使用されています。
    </div>
    <a href="?action=dashboard&password=hyggelyAdmin2024" class="btn btn-primary">
      ダッシュボードに戻る
    </a>
  </div>
</body>
</html>
    `;
    
    return HtmlService.createHtmlOutput(emailSettingsHtml)
      .setTitle('メール設定')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
      
  } catch (error) {
    return createErrorPage('メール設定の読み込みに失敗しました', error.toString());
  }
}

// 🔧 修正版：ヘルスチェック
function handleHealthCheck() {
  const health = {
    status: 'healthy',
    version: SYSTEM_CONFIG.version,
    timestamp: new Date().toISOString(),
    spreadsheetId: SYSTEM_CONFIG.spreadsheetId,
    environment: 'Google Apps Script'
  };
  
  const html = `
    <!DOCTYPE html>
    <html lang="ja">
    <head>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <title>システム状態</title>
      <style>
        body { 
          font-family: Arial, sans-serif; 
          padding: 20px; 
          background: #f8f9fa;
          margin: 0;
        }
        .container {
          max-width: 600px;
          margin: 0 auto;
          background: white;
          padding: 20px;
          border-radius: 8px;
          box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        pre {
          background: #f8f9fa;
          padding: 15px;
          border-radius: 4px;
          overflow-x: auto;
        }
      </style>
    </head>
    <body>
      <div class="container">
        <h1>✅ システム正常稼働中</h1>
        <pre>${JSON.stringify(health, null, 2)}</pre>
      </div>
    </body>
    </html>
  `;
  
  return HtmlService.createHtmlOutput(html)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

// 🔧 修正版：エラーページ
function createErrorPage(title, message) {
  const html = `
    <!DOCTYPE html>
    <html lang="ja">
    <head>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <title>エラー - Hyggelyカンパーニュ専門店</title>
      <style>
        body { 
          font-family: Arial, sans-serif; 
          text-align: center; 
          padding: 50px;
          background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%); 
          margin: 0;
          min-height: 100vh;
          display: flex;
          align-items: center;
          justify-content: center;
        }
        .container { 
          max-width: 600px; 
          margin: 0 auto; 
          background: white;
          padding: 40px; 
          border-radius: 15px; 
          box-shadow: 0 8px 30px rgba(0,0,0,0.1); 
        }
        .error-icon { 
          font-size: 4rem; 
          color: #dc3545; 
          margin-bottom: 20px; 
        }
        .btn { 
          display: inline-block; 
          padding: 12px 24px; 
          background: #8B4513;
          color: white; 
          text-decoration: none; 
          border-radius: 8px; 
          margin: 10px;
          border: none;
          cursor: pointer;
          font-size: 16px;
        }
        .btn:hover {
          background: #a0522d;
        }
      </style>
    </head>
    <body>
      <div class="container">
        <div class="error-icon">⚠️</div>
        <h1>${title}</h1>
        <p>${message}</p>
        <p>
          <a href="?" class="btn">🏠 予約フォームに戻る</a>
          <button onclick="location.reload()" class="btn">🔄 再読み込み</button>
        </p>
      </div>
    </body>
    </html>
  `;
  
  return HtmlService.createHtmlOutput(html)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

// 🔧 修正版：リダイレクトページ
function createRedirectPage(message, url) {
  const html = `
    <!DOCTYPE html>
    <html lang="ja">
    <head>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <title>リダイレクト中...</title>
      <style>
        body {
          text-align: center; 
          padding: 100px; 
          font-family: Arial;
          background: #f8f9fa;
          margin: 0;
        }
        .container {
          max-width: 400px;
          margin: 0 auto;
          background: white;
          padding: 30px;
          border-radius: 8px;
          box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
      </style>
    </head>
    <body>
      <div class="container">
        <h2>${message}</h2>
        <p>3秒後にリダイレクトします...</p>
        <script>
          setTimeout(function() { 
            window.location.href = '${url}'; 
          }, 3000);
        </script>
      </div>
    </body>
    </html>
  `;
  
  return HtmlService.createHtmlOutput(html)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

// ===== 既存の関数群（変更なし） =====
function checkAndInitializeSystem() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const sheets = spreadsheet.getSheets().map(s => s.getName());
    
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

function initOrderSheet(sheet) {
  const basicHeaders = ['タイムスタンプ', '姓', '名', 'メール', '受取日', '受取時間'];
  const products = getDefaultProducts();
  const productHeaders = products.map(p => p.name);
  const finalHeaders = ['その他ご要望', '合計金額', '引渡済', '予約ID'];
  const allHeaders = [...basicHeaders, ...productHeaders, ...finalHeaders];
  
  sheet.getRange(1, 1, 1, allHeaders.length).setValues([allHeaders]);
  
  // 列の説明をコメントとして追加
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

// 🔧 修正版：商品マスタ
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

// 🔧 修正版：予約一覧取得関数
function getOrderList() {
  try {
    console.log('📊 予約一覧取得開始');
    
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const orderSheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.ORDER);
    
    if (!orderSheet || orderSheet.getLastRow() <= 1) {
      console.log('⚠️ 予約データが存在しません');
      return [];
    }
    
    const data = orderSheet.getDataRange().getValues();
    const products = getProductMaster();
    
    console.log(`📋 データ行数: ${data.length - 1}, 商品数: ${products.length}`);
    
    const orders = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // 基本データのチェック
      if (!row[0] || (!row[1] && !row[2])) {
        console.log(`⚠️ 行 ${i + 1}: 必須データが不足しています`);
        continue;
      }
      
      // 受取日を yyyy-MM-dd 形式に統一
      let pickupDate = '';
      if (row[4]) {
        if (row[4] instanceof Date) {
          pickupDate = Utilities.formatDate(row[4], Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } else if (typeof row[4] === 'string') {
          const dateStr = row[4].toString().trim();
          if (dateStr.match(/^\d{4}[-\/]\d{1,2}[-\/]\d{1,2}$/)) {
            const parts = dateStr.split(/[-\/]/);
            pickupDate = `${parts[0]}-${('0' + parts[1]).slice(-2)}-${('0' + parts[2]).slice(-2)}`;
          } else {
            console.log(`⚠️ 行 ${i + 1}: 受取日の形式が不正です: ${dateStr}`);
            pickupDate = dateStr;
          }
        }
      }
      
      // 予約IDの処理
      let orderId = row[36]; // AK列
      if (!orderId) {
        orderId = generateOrderId();
        console.log(`⚠️ 行 ${i + 1}: 予約IDが空のため生成しました: ${orderId}`);
        try {
          orderSheet.getRange(i + 1, 37).setValue(orderId);
        } catch (e) {
          console.error('予約ID記録エラー:', e);
        }
      }
      
      // 基本オーダー情報
      const order = {
        rowIndex: i + 1,
        timestamp: row[0],
        lastName: row[1] || '',
        firstName: row[2] || '',
        email: row[3] || '',
        pickupDate: pickupDate,
        pickupTime: row[5] || '',
        items: [],
        note: row[33] || '',
        totalPrice: parseFloat(row[34]) || 0,
        isDelivered: row[35] === '引渡済',
        orderId: orderId,
        updatedAt: row[0] || new Date()
      };
      
      // 商品データを解析（G列～AG列：6～32のインデックス）
      let totalCalculatedPrice = 0;
      for (let j = 6; j <= 32; j++) {
        const quantity = parseInt(row[j]) || 0;
        if (quantity > 0) {
          const productIndex = j - 6;
          
          if (productIndex < products.length) {
            const product = products[productIndex];
            const subtotal = quantity * product.price;
            
            order.items.push({
              productId: product.id,
              name: product.name,
              quantity: quantity,
              price: product.price,
              subtotal: subtotal
            });
            
            totalCalculatedPrice += subtotal;
          }
        }
      }
      
      // 計算された合計金額の補完
      if (order.totalPrice === 0 && totalCalculatedPrice > 0) {
        order.totalPrice = totalCalculatedPrice;
      }
      
      orders.push(order);
    }
    
    // 受取日時昇順でソート
    orders.sort((a, b) => {
      const dateA = new Date(a.pickupDate + ' ' + (a.pickupTime || '00:00'));
      const dateB = new Date(b.pickupDate + ' ' + (b.pickupTime || '00:00'));
      
      if (dateA.getTime() !== dateB.getTime()) {
        return dateA.getTime() - dateB.getTime();
      }
      
      const timestampA = new Date(a.timestamp);
      const timestampB = new Date(b.timestamp);
      return timestampA.getTime() - timestampB.getTime();
    });
    
    console.log(`📊 予約一覧取得完了: ${orders.length}件`);
    return orders;
    
  } catch (error) {
    console.error('❌ 予約一覧取得エラー:', error);
    logSystemEvent('ERROR', '予約一覧取得エラー', error.toString());
    return [];
  }
}

// 🔧 修正版：統計データ取得関数
function getDashboardStats() {
  try {
    console.log('📊 統計データ取得開始');
    
    const orders = getOrderList();
    const inventory = getInventoryDataForForm();
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    // 今日の予約
    const todayOrders = orders.filter(order => {
      if (!order.pickupDate) return false;
      try {
        const pickupDate = new Date(order.pickupDate);
        pickupDate.setHours(0, 0, 0, 0);
        return pickupDate.getTime() === today.getTime();
      } catch (e) {
        return false;
      }
    });
    
    // 未引渡し予約
    const pendingOrders = orders.filter(order => !order.isDelivered);
    
    // 在庫アラート
    const outOfStock = inventory.filter(p => p.remaining <= 0);
    const lowStock = inventory.filter(p => p.remaining > 0 && p.remaining <= (p.minStock || 3));
    
    // 今月の売上
    const currentMonth = today.getMonth();
    const currentYear = today.getFullYear();
    const monthOrders = orders.filter(order => {
      if (!order.timestamp) return false;
      try {
        const orderDate = new Date(order.timestamp);
        return orderDate.getMonth() === currentMonth && 
               orderDate.getFullYear() === currentYear;
      } catch (e) {
        return false;
      }
    });
    const monthRevenue = monthOrders.reduce((sum, order) => sum + (order.totalPrice || 0), 0);
    
    const stats = {
      todayOrdersCount: todayOrders.length,
      pendingOrdersCount: pendingOrders.length,
      outOfStockCount: outOfStock.length,
      lowStockCount: lowStock.length,
      totalProducts: inventory.length,
      todayRevenue: todayOrders.reduce((sum, order) => sum + (order.totalPrice || 0), 0),
      monthRevenue: monthRevenue,
      systemVersion: SYSTEM_CONFIG.version,
      lastUpdate: new Date().toISOString()
    };
    
    console.log('📊 統計データ:', stats);
    return stats;
    
  } catch (error) {
    console.error('❌ 統計データ取得エラー:', error);
    return {
      todayOrdersCount: 0,
      pendingOrdersCount: 0,
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

// 🔧 新規追加：デバッグ用関数
function debugOrderData() {
  try {
    console.log('🔍 デバッグモード開始');
    
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const orderSheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.ORDER);
    
    if (!orderSheet) {
      return {
        error: '予約管理票シートが見つかりません',
        totalRows: 0,
        totalColumns: 0,
        headers: [],
        sampleData: []
      };
    }
    
    const data = orderSheet.getDataRange().getValues();
    const headers = data[0] || [];
    const sampleData = data[1] || [];
    
    const importantColumns = {
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
    
    return {
      totalRows: data.length,
      totalColumns: data[0] ? data[0].length : 0,
      headers: headers,
      sampleData: sampleData,
      importantColumns: importantColumns,
      spreadsheetId: SYSTEM_CONFIG.spreadsheetId,
      sheetName: SYSTEM_CONFIG.sheets.ORDER
    };
    
  } catch (error) {
    console.error('❌ デバッグ確認エラー:', error);
    return {
      error: error.toString(),
      totalRows: 0,
      totalColumns: 0,
      headers: [],
      sampleData: []
    };
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
    
    // 商品数量を記録（G~AG列：7~33列目）
    for (let i = 0; i < products.length; i++) {
      const quantity = parseInt(formData[`product_${i}`]) || 0;
      orderSheet.getRange(lastRow, 7 + i).setValue(quantity);
    }
    
    // 最終項目を記録
    orderSheet.getRange(lastRow, 34).setValue(formData.note || '');
    orderSheet.getRange(lastRow, 35).setValue(totalPrice);
    orderSheet.getRange(lastRow, 36).setValue('未引渡');
    orderSheet.getRange(lastRow, 37).setValue(orderId);
    
    // 在庫を更新
    updateInventoryFromOrders();
    
    // メール送信
    console.log('📧 メール送信処理開始');
    const emailResults = sendOrderEmails(formData, orderedItems, totalPrice, orderId);
    console.log('📧 メール送信結果:', emailResults);
    
    // ログ記録
    logSystemEvent('INFO', '新規予約',
      `顧客: ${formData.lastName} ${formData.firstName}, 金額: ¥${totalPrice}, 予約ID: ${orderId}`);
    
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
      const isDelivered = row[35] === '引渡済';
      
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
        inventorySheet.getRange(i + 1, 5).setValue(reservations[productId]);
        
        const stock = inventoryData[i][3] || 0;
        const remaining = Math.max(0, stock - reservations[productId]);
        inventorySheet.getRange(i + 1, 6).setValue(remaining);
        inventorySheet.getRange(i + 1, 8).setValue(new Date());
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

// 🔧 修正版：引渡状態更新関数
function updateDeliveryStatus(rowIndex, isDelivered) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const orderSheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.ORDER);
    
    const statusValue = isDelivered ? '引渡済' : '未引渡';
    orderSheet.getRange(rowIndex, 36).setValue(statusValue);
    
    updateInventoryFromOrders();
    
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

// ===== メール機能 =====
function getEmailSettings() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SYSTEM_CONFIG.spreadsheetId);
    const emailSheet = spreadsheet.getSheetByName(SYSTEM_CONFIG.sheets.EMAIL_SETTINGS);
    
    const defaultSettings = {
      adminEmail: 'hyggely2021@gmail.com',
      emailEnabled: true,
      customerSubject: 'Hyggelyカンパーニュ専門店 ご予約完了確認',
      customerBody: '{lastName} {firstName} 様\n\nご予約ありがとうございます。\n{orderItems}\n\n合計: ¥{totalPrice}\n受取日時: {pickupDateTime}',
      adminSubject: '【新規予約】{lastName} {firstName}様',
      adminBody: '新規予約\n\nお客様: {lastName} {firstName}\nメール: {email}\n受取: {pickupDateTime}\n{orderItems}\n合計: ¥{totalPrice}'
    };
    
    if (!emailSheet || emailSheet.getLastRow() <= 1) {
      return defaultSettings;
    }
    
    const data = emailSheet.getDataRange().getValues();
    const settings = { ...defaultSettings };
    
    for (let i = 1; i < data.length; i++) {
      const key = data[i][0];
      const value = data[i][1];
      
      if (!key) continue;
      
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
    
    return settings;
    
  } catch (error) {
    console.error('❌ メール設定取得エラー:', error);
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

function sendOrderEmails(formData, orderedItems, totalPrice, orderId) {
  console.log('📧 メール送信統合処理開始');
  
  const results = {
    success: false,
    customerEmailSent: false,
    adminEmailSent: false,
    errors: []
  };
  
  try {
    const settings = getEmailSettings();
    
    if (!settings.emailEnabled) {
      console.log('⚠️ メール送信が無効に設定されています');
      results.errors.push('メール送信が無効に設定されています');
      return results;
    }
    
    if (!checkGmailPermission()) {
      const errorMsg = 'Gmail送信権限が不足しています';
      console.error('❌ ' + errorMsg);
      results.errors.push(errorMsg);
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
    
    return results;
    
  } catch (error) {
    console.error('❌ メール送信統合処理エラー:', error);
    results.errors.push('統合処理エラー: ' + error.toString());
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
    
    GmailApp.sendEmail(formData.email, subject, body);
    
    console.log('✅ 顧客確認メール送信成功');
    return { success: true };
    
  } catch (error) {
    console.error('❌ 顧客メール送信エラー:', error);
    return { success: false, error: error.toString() };
  }
}

function sendAdminNotification(formData, orderedItems, totalPrice, orderId, settings) {
  try {
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
    
    GmailApp.sendEmail(settings.adminEmail, subject, body);
    
    console.log('✅ 管理者通知メール送信成功');
    return { success: true };
    
  } catch (error) {
    console.error('❌ 管理者メール送信エラー:', error);
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

function testConnection() {
  return {
    success: true,
    timestamp: new Date().toISOString(),
    version: SYSTEM_CONFIG.version,
    status: 'operational'
  };
}