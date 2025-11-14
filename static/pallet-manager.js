// ========================================
// パレット管理モジュール
// index.htmlの2761行目～3048行目付近から抽出
// ========================================

// ========================================
// パレット一覧を読み込み
// ========================================
// index.htmlの2761行目から抽出
// switchTab('pallets')時に呼び出される
async function loadPallets() {
    try {
        const response = await fetch('/api/pallets/list');
        const data = await response.json();
        
        if (data.success) {
            // displayPallets()とdisplayPalletStats()を呼び出してデータを表示
            displayPallets(data.pallets);
            displayPalletStats(data.pallets);
        } else {
            showToast('パレット一覧の読み込みに失敗しました', 'error');
        }
    } catch (error) {
        console.error('Error loading pallets:', error);
        showToast('エラー: ' + error, 'error');
    }
}

// ========================================
// パレット一覧を表示
// ========================================
// index.htmlの2779行目から抽出
// loadPallets()から呼び出される
function displayPallets(pallets) {
    const palletList = document.getElementById('palletList');
    
    if (pallets.length === 0) {
        palletList.innerHTML = '<p>パレットがありません</p>';
        return;
    }
    
    let html = '';
    pallets.forEach(pallet => {
        // フロアバッジを生成（1F=緑、2F=青）
        const floorColor = pallet.floor === '1F' ? '#28a745' : '#17a2b8';
        const floorBadge = pallet.floor ? 
            `<span style="background: ${floorColor}; color: white; padding: 3px 8px; border-radius: 12px; font-size: 0.85em; margin-left: 5px;">${pallet.floor}</span>` : '';
        
        // パレットカードの開始
        html += `
            <div class="order-card" style="cursor: default;">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
                    <h3 style="margin: 0; color: #667eea;">📦 ${pallet.pallet_number}</h3>
                    ${floorBadge}
                </div>
                <div style="margin-bottom: 15px; padding: 10px; background: #f8f9fa; border-radius: 5px;">
                    <strong>格納製番 (${pallet.order_count}件)</strong>
                    <ul style="margin: 10px 0 0 0; padding-left: 20px; list-style: none;">
        `;
        
        // パレット内の各注文を表示
        pallet.orders.forEach(order => {
            // ステータスに応じた色を設定
            const statusColor = order.status === '納品完了' ? '#2ef148e0' : 
                            order.status === '納品中' ? '#1758fdff' : '#f5d800ff';
            
            // 品名を表示（長い場合は省略）
            const productName = order.product_name ? 
                (order.product_name.length > 20 ? order.product_name.substring(0, 20) + '...' : order.product_name) : 
                '';

            // 得意先略称を表示
            const customerAbbr = order.customer_abbr ? order.customer_abbr : '';

            // 各注文の行を生成（クリックでshowOrderDetails()を呼び出し）
            html += `
                <li style="margin: 8px 0; cursor: pointer; padding: 5px; background: white; border-radius: 3px;" onclick="showOrderDetails(${order.id})">
                    <div style="display: flex; justify-content: space-between; align-items: center;">
                        <div>
                            <span style="font-weight: bold;">${order.seiban}</span>
                            ${order.unit ? `<span style="color: #6c757d; margin-left: 3px;">(${order.unit})</span>` : ''}
                            ${productName ? `<br><span style="font-size: 0.85em; color: #6c757d;">${productName}</span>` : ''}
                            ${customerAbbr ? `<br><span style="font-size: 0.8em; color: #17a2b8;">🏢 ${customerAbbr}</span>` : ''}
                        </div>
                        <span style="color: ${statusColor}; font-size: 1.2em;">●</span>
                    </div>
                </li>
            `;
        });
        
        // パレットカードの終了とラベル印刷ボタン（printPalletLabel()を呼び出し）
        html += `
                    </ul>
                </div>
                <button class="btn btn-info btn-sm" onclick="printPalletLabel('${pallet.pallet_number}')">
                    🖨️ ラベル印刷
                </button>
            </div>
        `;
    });
    
    // 生成したHTMLをpalletListに表示
    palletList.innerHTML = html;
}


// ========================================
// パレット統計を表示
// ========================================
// index.htmlの2845行目から抽出
// loadPallets()から呼び出される
function displayPalletStats(pallets) {
    const statsDiv = document.getElementById('palletStats');
    
    // フロア別にパレット数を集計
    const floor1F = pallets.filter(p => p.floor === '1F').length;
    const floor2F = pallets.filter(p => p.floor === '2F').length;
    const noFloor = pallets.filter(p => !p.floor || p.floor === '').length;
    const totalOrders = pallets.reduce((sum, p) => sum + p.order_count, 0);
    
    // 統計カードを生成
    statsDiv.innerHTML = `
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px;">
            <div style="background: #667eea; color: white; padding: 20px; border-radius: 10px; text-align: center;">
                <div style="font-size: 2em; font-weight: bold;">${pallets.length}</div>
                <div>総保管場所数</div>
            </div>
            <div style="background: #28a745; color: white; padding: 20px; border-radius: 10px; text-align: center;">
                <div style="font-size: 2em; font-weight: bold;">${floor1F}</div>
                <div>1F 保管場所</div>
            </div>
            <div style="background: #17a2b8; color: white; padding: 20px; border-radius: 10px; text-align: center;">
                <div style="font-size: 2em; font-weight: bold;">${floor2F}</div>
                <div>2F 保管場所</div>
            </div>
            <div style="background: #ffc107; color: #212529; padding: 20px; border-radius: 10px; text-align: center;">
                <div style="font-size: 2em; font-weight: bold;">${totalOrders}</div>
                <div>総製番数</div>
            </div>
        </div>
    `;
}

// ========================================
// 製番検索
// ========================================
// index.htmlの2876行目から抽出
// ユーザーが検索ボタンを押したときに呼び出される
async function searchSeiban() {
    const searchQuery = document.getElementById('seibanSearchInput').value.trim();
    
    if (!searchQuery) {
        showToast('製番または品名を入力してください', 'warning');
        return;
    }
    
    try {
        // APIで検索を実行
        const response = await fetch(`/api/pallets/search?query=${encodeURIComponent(searchQuery)}`);
        const data = await response.json();
        
        const resultsDiv = document.getElementById('searchResults');
        
        if (data.success) {
            // 検索結果を表示
            let html = `<div style="margin-top: 20px;"><h4>検索結果 (${data.results.length}件)</h4>`;
            
            data.results.forEach(result => {
                // パレット情報を生成
                const palletInfo = result.pallet_number && result.pallet_number !== '未設定' ? 
                    `<strong>📦 パレット: ${result.pallet_number}</strong> (${result.floor || '未設定'})` : 
                    '<span style="color: #dc3545;">パレット未設定</span>';
                
                // ステータスに応じた色を設定
                const statusColor = result.status === '納品完了' ? '#28a745' : 
                                result.status === '納品中' ? '#17a2b8' : '#ffc107';
                
                // 品名を生成
                const productName = result.product_name ? 
                    `<div style="font-size: 0.95em; color: #495057; margin-top: 3px;">品名: ${result.product_name}</div>` : '';
                
                // 各検索結果のカードを生成（クリックでshowOrderDetails()を呼び出し）
                html += `
                    <div style="background: white; border: 1px solid #dee2e6; border-radius: 8px; padding: 15px; margin-bottom: 10px; cursor: pointer;" 
                        onclick="showOrderDetails(${result.id})">
                        <div style="display: flex; justify-content: space-between; align-items: center;">
                            <div style="flex: 1;">
                                <div style="font-size: 1.2em; font-weight: bold; margin-bottom: 5px;">
                                    ${result.seiban} ${result.unit ? `(${result.unit})` : ''}
                                </div>
                                ${productName}
                                <div style="margin-top: 8px; margin-bottom: 5px;">
                                    ${palletInfo}
                                </div>
                                <div style="font-size: 0.9em; color: #6c757d;">
                                    ${result.customer_abbr ? `得意先: ${result.customer_abbr}` : ''}
                                </div>
                            </div>
                            <div style="background: ${statusColor}; color: white; padding: 5px 15px; border-radius: 20px; font-size: 0.9em; white-space: nowrap;">
                                ${result.status}
                            </div>
                        </div>
                    </div>
                `;
            });
            
            html += '</div>';
            resultsDiv.innerHTML = html;
        } else {
            resultsDiv.innerHTML = `<div class="alert alert-warning" style="margin-top: 20px;">${data.error}</div>`;
        }
    } catch (error) {
        console.error('Search error:', error);
        showToast('検索エラー: ' + error, 'error');
    }
}

// ========================================
// パレットラベル印刷
// ========================================
// index.htmlの2939行目から抽出
// displayPallets()内のボタンから呼び出される
async function printPalletLabel(palletNumber) {
    try {
        // APIでラベル情報を取得
        const response = await fetch(`/api/pallets/${palletNumber}/label`);
        const data = await response.json();
        
        if (data.success) {
            // 印刷用ウィンドウを開く
            const printWindow = window.open('', '_blank', 'width=800,height=600');
            
            // 注文リストのHTMLを生成
            const ordersHtml = data.orders.map(order => 
                `<li>${order.seiban} ${order.unit ? `(${order.unit})` : ''}</li>`
            ).join('');
            
            // ラベルのHTMLを生成して印刷ウィンドウに書き込み
            printWindow.document.write(`
                <!DOCTYPE html>
                <html>
                <head>
                    <title>パレットラベル - ${palletNumber}</title>
                    <style>
                        body {
                            font-family: Arial, sans-serif;
                            padding: 20px;
                            text-align: center;
                        }
                        .label-container {
                            border: 2px solid #000;
                            padding: 30px;
                            max-width: 600px;
                            margin: 0 auto;
                        }
                        h1 {
                            font-size: 2.5em;
                            margin: 10px 0;
                        }
                        .floor-badge {
                            display: inline-block;
                            background: #17a2b8;
                            color: white;
                            padding: 10px 20px;
                            border-radius: 20px;
                            font-size: 1.5em;
                            margin: 10px 0;
                        }
                        img {
                            max-width: 300px;
                            margin: 20px 0;
                        }
                        .orders-list {
                            text-align: left;
                            margin: 20px 0;
                            padding: 20px;
                            background: #f8f9fa;
                            border-radius: 10px;
                        }
                        .orders-list h3 {
                            margin-top: 0;
                        }
                        ul {
                            list-style-position: inside;
                        }
                        li {
                            margin: 8px 0;
                            font-size: 1.1em;
                        }
                        @media print {
                            .no-print {
                                display: none;
                            }
                        }
                    </style>
                </head>
                <body>
                    <div class="label-container">
                        <h1>📦 パレット ${data.pallet_number}</h1>
                        ${data.floor ? `<div class="floor-badge">${data.floor}</div>` : ''}
                        <img src="data:image/png;base64,${data.qr_code}" alt="QR Code">
                        <div class="orders-list">
                            <h3>格納製番 (${data.order_count}件)</h3>
                            <ul>
                                ${ordersHtml}
                            </ul>
                        </div>
                    </div>
                    <button class="no-print" onclick="window.print()" 
                            style="margin-top: 20px; padding: 10px 20px; font-size: 1em; cursor: pointer;">
                        🖨️ 印刷
                    </button>
                </body>
                </html>
            `);
            
            printWindow.document.close();
        } else {
            showToast('ラベル生成に失敗しました: ' + data.error, 'error');
        }
    } catch (error) {
        console.error('Print label error:', error);
        showToast('エラー: ' + error, 'error');
    }
}

// ========================================
// タブ切り替え時の処理（初期化）
// ========================================
// index.htmlの3042行目から抽出
// DOMContentLoaded後にswitchTabを拡張
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initPalletManager);
} else {
    initPalletManager();
}

function initPalletManager() {
    // switchTab関数が定義されているか確認してから拡張
    if (typeof switchTab === 'function') {
        const originalSwitchTab = switchTab;
        window.switchTab = function(tabName) {
            originalSwitchTab(tabName);
            // パレットタブに切り替えたときにloadPallets()を呼び出し
            if (tabName === 'pallets') {
                loadPallets();
            }
        };
        console.log('パレット管理モジュールが初期化されました');
    } else {
        console.warn('switchTab関数が見つかりません。パレット管理の自動読み込みが無効化されています。');
    }
}