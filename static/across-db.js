/**
 * Across DB テストツール
 * V_D系ビューへの直接クエリUI
 */

// ========== 接続テスト ==========
async function testAcrossConnection() {
    const resultDiv = document.getElementById('acrossTestResult');
    resultDiv.innerHTML = '<span style="color:#666;">接続テスト中...</span>';

    try {
        const res = await fetch('/api/across-db/test');
        const data = await res.json();
        if (data.success) {
            resultDiv.innerHTML = '<span style="color:green;">✓ ' + data.message + '</span>';
        } else {
            resultDiv.innerHTML = '<span style="color:red;">✗ ' + data.message + '</span>';
        }
    } catch (e) {
        resultDiv.innerHTML = '<span style="color:red;">✗ 接続失敗: ' + e + '</span>';
    }
}

// ========== ビュー選択変更 ==========
function onAcrossViewChange() {
    const view = document.getElementById('acrossViewSelect').value;
    const searchType = document.getElementById('acrossSearchType');
    const searchInput = document.getElementById('acrossSearchInput');
    const placeholder = document.getElementById('acrossSearchPlaceholder');

    // 検索タイプの選択肢をビューに応じて変更
    searchType.innerHTML = '';

    if (view === 'V_D発注' || view === 'V_D発注残' || view === 'V_D仕入') {
        searchType.innerHTML += '<option value="発注番号">発注番号</option>';
        searchType.innerHTML += '<option value="製番">製番</option>';
        searchInput.placeholder = '例: 89074 または MHT0620';
    } else if (view === 'V_D手配リスト') {
        searchType.innerHTML += '<option value="製番">製番</option>';
        searchInput.placeholder = '例: MHT0620';
    }

    searchType.innerHTML += '<option value="">全件（上位100件）</option>';

    // カラム情報を取得
    loadAcrossColumns(view);
}

// ========== カラム一覧取得 ==========
async function loadAcrossColumns(viewName) {
    const colDiv = document.getElementById('acrossColumns');
    colDiv.innerHTML = '<small style="color:#999;">カラム読み込み中...</small>';

    try {
        const res = await fetch('/api/across-db/columns?view=' + encodeURIComponent(viewName));
        const data = await res.json();
        if (data.columns) {
            colDiv.innerHTML = '<small style="color:#888;">カラム: ' +
                data.columns.map(c => '<code>' + c + '</code>').join(', ') + '</small>';
        }
    } catch (e) {
        colDiv.innerHTML = '';
    }
}

// ========== 検索実行 ==========
async function executeAcrossQuery() {
    const view = document.getElementById('acrossViewSelect').value;
    const searchType = document.getElementById('acrossSearchType').value;
    const searchValue = document.getElementById('acrossSearchInput').value.trim();
    const limitVal = document.getElementById('acrossLimit').value || 100;
    const resultDiv = document.getElementById('acrossQueryResult');

    resultDiv.innerHTML = '<div style="text-align:center; padding:20px; color:#666;">検索中...</div>';

    const params = new URLSearchParams({
        view: view,
        limit: limitVal
    });

    if (searchType && searchValue) {
        params.set('search_type', searchType);
        params.set('search_value', searchValue);
    }

    try {
        const res = await fetch('/api/across-db/query?' + params.toString());
        const data = await res.json();

        if (data.error) {
            resultDiv.innerHTML = '<div class="alert alert-danger">' + data.error + '</div>';
            return;
        }

        renderAcrossResult(data, resultDiv);

    } catch (e) {
        resultDiv.innerHTML = '<div class="alert alert-danger">クエリ失敗: ' + e + '</div>';
    }
}

// ========== 結果テーブル描画 ==========
function renderAcrossResult(data, container) {
    if (!data.rows || data.rows.length === 0) {
        container.innerHTML = '<div class="alert alert-warning">該当データなし（0件）</div>';
        return;
    }

    const columns = data.columns;
    const rows = data.rows;

    let html = '<div style="margin-bottom:8px;">';
    html += '<strong>' + data.count + '件</strong> 取得';
    html += '</div>';

    html += '<div style="overflow-x:auto; max-height:500px; overflow-y:auto; border:1px solid #ddd; border-radius:5px;">';
    html += '<table style="width:100%; border-collapse:collapse; font-size:0.85em; white-space:nowrap;">';

    // ヘッダー
    html += '<thead><tr style="background:#f0f0f0; position:sticky; top:0; z-index:1;">';
    html += '<th style="padding:6px 10px; border:1px solid #ddd; background:#e0e0e0; color:#666;">#</th>';
    for (const col of columns) {
        html += '<th style="padding:6px 10px; border:1px solid #ddd; background:#e0e0e0;">' + escapeForTable(col) + '</th>';
    }
    html += '</tr></thead>';

    // データ行
    html += '<tbody>';
    for (let i = 0; i < rows.length; i++) {
        const bgColor = i % 2 === 0 ? '#fff' : '#f9f9f9';
        html += '<tr style="background:' + bgColor + ';">';
        html += '<td style="padding:4px 8px; border:1px solid #eee; color:#999; text-align:center;">' + (i + 1) + '</td>';
        for (let j = 0; j < rows[i].length; j++) {
            const val = rows[i][j];
            const display = val === null ? '<span style="color:#ccc;">(NULL)</span>' : escapeForTable(String(val));
            html += '<td style="padding:4px 8px; border:1px solid #eee;">' + display + '</td>';
        }
        html += '</tr>';
    }
    html += '</tbody></table></div>';

    container.innerHTML = html;
}

// ========== HTMLエスケープ ==========
function escapeForTable(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
}

// ========== 発注番号クイック検索 ==========
async function quickSearchOrder() {
    const input = document.getElementById('acrossQuickOrderInput').value.trim();
    if (!input) return;

    const resultDiv = document.getElementById('acrossQuickResult');
    resultDiv.innerHTML = '<div style="padding:10px; color:#666;">検索中...</div>';

    try {
        const res = await fetch('/api/across-db/order-detail?order_number=' + encodeURIComponent(input));
        const data = await res.json();

        if (data.error) {
            resultDiv.innerHTML = '<div class="alert alert-danger">' + data.error + '</div>';
            return;
        }

        renderOrderDetail(data, resultDiv);

    } catch (e) {
        resultDiv.innerHTML = '<div class="alert alert-danger">検索失敗: ' + e + '</div>';
    }
}

// ========== 発注詳細表示 ==========
function renderOrderDetail(data, container) {
    const order = data.order;
    const remaining = data.remaining;
    const receipts = data.receipts;

    if (!order || order.count === 0) {
        container.innerHTML = '<div class="alert alert-warning">該当する発注番号が見つかりません</div>';
        return;
    }

    let html = '';

    // 発注基本情報
    const row = order.rows[0];
    const cols = order.columns;
    const rec = {};
    for (let i = 0; i < cols.length; i++) {
        rec[cols[i]] = row[i];
    }

    html += '<div style="background:#e8f5e9; padding:12px; border-radius:6px; margin-bottom:10px;">';
    html += '<h4 style="margin:0 0 8px;">発注番号: ' + escapeForTable(rec['発注番号'] || '') + '</h4>';
    html += '<div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(200px, 1fr)); gap:4px 16px; font-size:0.9em;">';

    const keyFields = ['製番', '品名', '仕様１', '仕様２', '仕入先名', '発注数', '単位', '発注単価', '発注金額', '発注日', '納期', '回答納期', '手配区分', '材質', '備考'];
    for (const key of keyFields) {
        if (rec[key] !== undefined && rec[key] !== null) {
            html += '<div><strong>' + escapeForTable(key) + ':</strong> ' + escapeForTable(String(rec[key])) + '</div>';
        }
    }
    html += '</div></div>';

    // 発注残（納入状況）
    if (remaining && remaining.count > 0) {
        const remRow = remaining.rows[0];
        const remCols = remaining.columns;
        const remRec = {};
        for (let i = 0; i < remCols.length; i++) {
            remRec[remCols[i]] = remRow[i];
        }

        const ordered = parseFloat(remRec['発注数']) || 0;
        const delivered = parseFloat(remRec['納入済数']) || 0;
        const pct = ordered > 0 ? Math.round((delivered / ordered) * 100) : 0;
        const barColor = pct >= 100 ? '#28a745' : pct > 0 ? '#ffc107' : '#dc3545';

        html += '<div style="background:#fff3cd; padding:10px; border-radius:6px; margin-bottom:10px;">';
        html += '<strong>納入状況:</strong> ' + delivered + ' / ' + ordered + ' (' + pct + '%)';
        html += '<div style="background:#e9ecef; border-radius:10px; overflow:hidden; height:16px; margin-top:4px;">';
        html += '<div style="height:100%; width:' + pct + '%; background:' + barColor + '; transition:width 0.3s;"></div>';
        html += '</div></div>';
    }

    // 仕入（納入実績）
    if (receipts && receipts.count > 0) {
        html += '<div style="background:#e3f2fd; padding:10px; border-radius:6px;">';
        html += '<strong>納入実績 (' + receipts.count + '件):</strong>';
        html += '<table style="width:100%; border-collapse:collapse; font-size:0.85em; margin-top:6px;">';
        html += '<tr style="background:#bbdefb;">';
        for (const col of receipts.columns) {
            html += '<th style="padding:4px 6px; border:1px solid #90caf9;">' + escapeForTable(col) + '</th>';
        }
        html += '</tr>';
        for (const rRow of receipts.rows) {
            html += '<tr>';
            for (const val of rRow) {
                const display = val === null ? '' : escapeForTable(String(val));
                html += '<td style="padding:3px 6px; border:1px solid #e0e0e0;">' + display + '</td>';
            }
            html += '</tr>';
        }
        html += '</table></div>';
    }

    container.innerHTML = html;
}

// ========== Enter キー対応 ==========
document.addEventListener('DOMContentLoaded', function() {
    const searchInput = document.getElementById('acrossSearchInput');
    if (searchInput) {
        searchInput.addEventListener('keypress', function(e) {
            if (e.key === 'Enter') executeAcrossQuery();
        });
    }
    const quickInput = document.getElementById('acrossQuickOrderInput');
    if (quickInput) {
        quickInput.addEventListener('keypress', function(e) {
            if (e.key === 'Enter') quickSearchOrder();
        });
    }
});
