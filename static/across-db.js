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
        searchType.innerHTML += '<option value="品目CD">品目CD</option>';
        searchType.innerHTML += '<option value="品名">品名（部分一致）</option>';
        searchType.innerHTML += '<option value="仕様１">仕様１（部分一致）</option>';
        searchInput.placeholder = '例: 89074, MHT0620, 品名の一部';
    } else if (view === 'V_D手配リスト' || view === 'V_D未発注') {
        searchType.innerHTML += '<option value="製番">製番</option>';
        searchType.innerHTML += '<option value="品目CD">品目CD</option>';
        searchType.innerHTML += '<option value="品名">品名（部分一致）</option>';
        searchType.innerHTML += '<option value="仕様１">仕様１（部分一致）</option>';
        searchInput.placeholder = '例: MHT0620, 品名の一部';
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

// ========== 製番マージテスト ==========
async function executeMergeTest() {
    const seiban = document.getElementById('acrossMergeSeiban').value.trim();
    if (!seiban) return;

    const resultDiv = document.getElementById('acrossMergeResult');
    resultDiv.innerHTML = '<div style="text-align:center; padding:20px; color:#666;">マージ処理中...</div>';

    try {
        const res = await fetch('/api/across-db/merge-test?seiban=' + encodeURIComponent(seiban));
        const data = await res.json();

        if (data.error) {
            resultDiv.innerHTML = '<div class="alert alert-danger">' + escapeForTable(data.error) + '</div>';
            return;
        }

        if (!data.success) {
            resultDiv.innerHTML = '<div class="alert alert-danger">' + escapeForTable(data.error || '不明なエラー') + '</div>';
            return;
        }

        renderMergeResult(data, resultDiv);

    } catch (e) {
        resultDiv.innerHTML = '<div class="alert alert-danger">マージテスト失敗: ' + e + '</div>';
    }
}

function renderMergeResult(data, container) {
    const stats = data.stats;
    let html = '';

    // 統計サマリー
    const matchColor = stats.match_rate >= 80 ? '#28a745' : stats.match_rate >= 50 ? '#ffc107' : '#dc3545';
    html += '<div style="background:#e8f5e9; padding:12px; border-radius:6px; margin-bottom:10px;">';
    html += '<h4 style="margin:0 0 8px;">製番: ' + escapeForTable(data.seiban) + ' マージ結果</h4>';
    html += '<div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(150px, 1fr)); gap:6px; font-size:0.9em;">';
    html += '<div style="background:#fff; padding:8px; border-radius:4px; text-align:center;">';
    html += '<div style="font-size:1.5em; font-weight:bold;">' + stats.tehai_count + '</div>';
    html += '<div style="color:#666;">手配リスト件数</div></div>';
    html += '<div style="background:#fff; padding:8px; border-radius:4px; text-align:center;">';
    html += '<div style="font-size:1.5em; font-weight:bold;">' + stats.hatchu_count + '</div>';
    html += '<div style="color:#666;">発注データ件数</div></div>';
    html += '<div style="background:#fff; padding:8px; border-radius:4px; text-align:center;">';
    html += '<div style="font-size:1.5em; font-weight:bold; color:' + matchColor + ';">' + stats.match_rate + '%</div>';
    html += '<div style="color:#666;">マッチ率 (' + stats.match_count + '/' + stats.tehai_count + ')</div></div>';
    html += '<div style="background:#fff; padding:8px; border-radius:4px; text-align:center;">';
    html += '<div style="font-size:1.5em; font-weight:bold;">' + stats.unit_count + '</div>';
    html += '<div style="color:#666;">ユニット数</div></div>';
    html += '</div>';

    // ユニット一覧
    if (stats.units && stats.units.length > 0) {
        html += '<div style="margin-top:8px; font-size:0.85em;">';
        html += '<strong>ユニット:</strong> ';
        html += stats.units.map(u => '<span style="background:#e0e0e0; padding:2px 8px; border-radius:10px; margin:2px; display:inline-block;">' + escapeForTable(u || '(空)') + '</span>').join('');
        html += '</div>';
    }
    html += '</div>';

    // マージ結果テーブル
    if (data.rows && data.rows.length > 0) {
        html += '<div style="overflow-x:auto; max-height:500px; overflow-y:auto; border:1px solid #ddd; border-radius:5px;">';
        html += '<table style="width:100%; border-collapse:collapse; font-size:0.8em; white-space:nowrap;">';

        // ヘッダー
        html += '<thead><tr style="position:sticky; top:0; z-index:1;">';
        html += '<th style="padding:5px 8px; border:1px solid #ddd; background:#e0e0e0; color:#666;">#</th>';
        for (const col of data.columns) {
            let bg = '#e0e0e0';
            if (['発注番号', '仕入先略称', '納期', '納入済数'].includes(col)) bg = '#c8e6c9';
            if (col === 'match_type') bg = '#fff9c4';
            html += '<th style="padding:5px 8px; border:1px solid #ddd; background:' + bg + ';">' + escapeForTable(col) + '</th>';
        }
        html += '</tr></thead><tbody>';

        // データ
        for (let i = 0; i < data.rows.length; i++) {
            const row = data.rows[i];
            const matchType = row[data.columns.indexOf('match_type')] || '';
            const rowBg = matchType ? (i % 2 === 0 ? '#fff' : '#f9f9f9') : '#fff3e0';

            html += '<tr style="background:' + rowBg + ';">';
            html += '<td style="padding:3px 6px; border:1px solid #eee; color:#999; text-align:center;">' + (i + 1) + '</td>';
            for (let j = 0; j < row.length; j++) {
                const val = row[j];
                const colName = data.columns[j];
                let cellStyle = 'padding:3px 6px; border:1px solid #eee;';

                // マージ由来カラムをハイライト
                if (['発注番号', '仕入先略称', '納期', '納入済数'].includes(colName) && val) {
                    cellStyle += ' background:#e8f5e9; font-weight:bold;';
                }
                if (colName === 'match_type') {
                    if (val) {
                        cellStyle += ' background:#c8e6c9; color:#2e7d32; font-size:0.85em;';
                    } else {
                        cellStyle += ' background:#ffcdd2; color:#c62828; font-size:0.85em;';
                    }
                }

                const display = (val === null || val === '' || val === undefined)
                    ? '<span style="color:#ccc;">-</span>'
                    : escapeForTable(String(val));
                html += '<td style="' + cellStyle + '">' + display + '</td>';
            }
            html += '</tr>';
        }
        html += '</tbody></table></div>';
    }

    container.innerHTML = html;
}

// ========== 未発注検索 (社内加工品) ==========
async function searchMihatchu() {
    const seiban = document.getElementById('acrossMihatchuSeiban').value.trim();
    if (!seiban) return;

    const resultDiv = document.getElementById('acrossMihatchuResult');
    resultDiv.innerHTML = '<div style="text-align:center; padding:20px; color:#666;">検索中...</div>';

    const params = new URLSearchParams({
        seiban: seiban,
        supplier_cd: 'MHT',
        order_type_cd: '11'
    });

    try {
        const res = await fetch('/api/across-db/mihatchu?' + params.toString());
        const data = await res.json();

        if (data.error) {
            resultDiv.innerHTML = '<div class="alert alert-danger">' + escapeForTable(data.error) + '</div>';
            return;
        }

        if (!data.rows || data.rows.length === 0) {
            resultDiv.innerHTML = '<div class="alert alert-info">社内加工品(MHT+11)の未発注はありません</div>';
            return;
        }

        // 結果表示
        let html = '<div style="margin-bottom:8px;">';
        html += '<strong style="color:#9c27b0;">' + data.count + '件</strong> の社内加工品（未発注）があります';
        html += '</div>';

        renderAcrossResult(data, resultDiv);
        resultDiv.insertAdjacentHTML('afterbegin', html);

    } catch (e) {
        resultDiv.innerHTML = '<div class="alert alert-danger">検索失敗: ' + e + '</div>';
    }
}

// ========== DB更新チェック ==========
async function checkDbUpdates() {
    const resultDiv = document.getElementById('dbUpdateCheckResult');
    resultDiv.innerHTML = '<div style="padding:10px; color:#666;">更新確認中...</div>';

    // 現在登録済みの製番リストを取得
    let seibans = [];
    const seibanSet = new Set();
    document.querySelectorAll('.seiban-cb').forEach(cb => {
        const s = cb.value.trim();
        if (s && !seibanSet.has(s)) {
            seibanSet.add(s);
            seibans.push(s);
        }
    });

    if (seibans.length === 0) {
        resultDiv.innerHTML = '<div class="alert alert-warning">製番リストが空です</div>';
        return;
    }

    try {
        const res = await fetch('/api/across-db/check-updates', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ seibans: seibans })
        });
        const data = await res.json();

        if (data.error) {
            resultDiv.innerHTML = '<div class="alert alert-danger">' + escapeForTable(data.error) + '</div>';
            return;
        }

        if (!data.success || !data.results) {
            resultDiv.innerHTML = '<div class="alert alert-danger">更新チェック失敗</div>';
            return;
        }

        renderDbUpdateResult(data.results, resultDiv);

    } catch (e) {
        resultDiv.innerHTML = '<div class="alert alert-danger">更新チェック失敗: ' + e + '</div>';
    }
}

function renderDbUpdateResult(results, container) {
    let html = '<div style="margin-bottom:8px;"><strong>DB更新状況</strong></div>';

    html += '<div style="overflow-x:auto; max-height:300px; overflow-y:auto; border:1px solid #ddd; border-radius:5px;">';
    html += '<table style="width:100%; border-collapse:collapse; font-size:0.85em;">';
    html += '<thead><tr style="background:#e0e0e0; position:sticky; top:0;">';
    html += '<th style="padding:6px 10px; border:1px solid #ddd;">製番</th>';
    html += '<th style="padding:6px 10px; border:1px solid #ddd;">手配リスト</th>';
    html += '<th style="padding:6px 10px; border:1px solid #ddd;">発注</th>';
    html += '<th style="padding:6px 10px; border:1px solid #ddd; background:#f3e5f5;">社内加工(未発注)</th>';
    html += '</tr></thead><tbody>';

    let totalMihatchu = 0;
    for (const seiban of Object.keys(results).sort()) {
        const r = results[seiban];
        const hasMihatchu = r.mihatchu_count > 0;
        totalMihatchu += r.mihatchu_count;

        const rowBg = hasMihatchu ? '#fce4ec' : '#fff';
        html += '<tr style="background:' + rowBg + ';">';
        html += '<td style="padding:4px 8px; border:1px solid #eee; font-weight:bold;">' + escapeForTable(seiban) + '</td>';
        html += '<td style="padding:4px 8px; border:1px solid #eee; text-align:right;">' + r.tehai_count + '</td>';
        html += '<td style="padding:4px 8px; border:1px solid #eee; text-align:right;">' + r.hatchu_count + '</td>';
        html += '<td style="padding:4px 8px; border:1px solid #eee; text-align:right;' + (hasMihatchu ? ' color:#c2185b; font-weight:bold;' : '') + '">' + r.mihatchu_count + '</td>';
        html += '</tr>';
    }
    html += '</tbody></table></div>';

    // 通知バッジ
    if (totalMihatchu > 0) {
        html += '<div class="alert alert-warning" style="margin-top:10px;">';
        html += '<strong>⚠️ ' + totalMihatchu + '件</strong> の社内加工品（未発注）があります。DB直接処理で取り込み可能です。';
        html += '</div>';
        updateMihatchuBadge(totalMihatchu);
    } else {
        html += '<div class="alert alert-success" style="margin-top:10px;">';
        html += '✓ 社内加工品の未発注はありません';
        html += '</div>';
        updateMihatchuBadge(0);
    }

    container.innerHTML = html;
}

function updateMihatchuBadge(count) {
    const badge = document.getElementById('mihatchuBadge');
    if (badge) {
        if (count > 0) {
            badge.textContent = count;
            badge.style.display = 'inline-block';
        } else {
            badge.style.display = 'none';
        }
    }
}

// 自動更新チェック（ページ読み込み時）
function autoCheckDbUpdates() {
    setTimeout(() => {
        const resultDiv = document.getElementById('dbUpdateCheckResult');
        if (resultDiv && document.querySelectorAll('.seiban-cb').length > 0) {
            checkDbUpdates();
        }
    }, 2000);
}

// ========== 在庫部品検索 (手配区分CD=15) ==========
async function searchZaikoBuhin() {
    const resultDiv = document.getElementById('zaikoBuhinResult');
    resultDiv.innerHTML = '<div style="text-align:center; padding:20px; color:#666;">在庫部品を検索中...</div>';

    // 現在登録済みの製番リストを取得
    let seibans = [];
    const seibanSet = new Set();
    document.querySelectorAll('.seiban-cb').forEach(cb => {
        const s = cb.value.trim();
        if (s && !seibanSet.has(s)) {
            seibanSet.add(s);
            seibans.push(s);
        }
    });

    try {
        const res = await fetch('/api/across-db/zaiko-buhin', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ seibans: seibans.length > 0 ? seibans : null })
        });
        const data = await res.json();

        if (data.error) {
            resultDiv.innerHTML = '<div class="alert alert-danger">' + escapeForTable(data.error) + '</div>';
            return;
        }

        if (!data.success) {
            resultDiv.innerHTML = '<div class="alert alert-danger">検索失敗</div>';
            return;
        }

        renderZaikoBuhinResult(data, resultDiv);

    } catch (e) {
        resultDiv.innerHTML = '<div class="alert alert-danger">検索失敗: ' + e + '</div>';
    }
}

function renderZaikoBuhinResult(data, container) {
    if (!data.rows || data.rows.length === 0) {
        container.innerHTML = '<div class="alert alert-info">在庫部品（手配区分CD=15）はありません</div>';
        return;
    }

    let html = '<div style="margin-bottom:10px;">';
    html += '<strong style="color:#2196f3;">' + data.count + '件</strong> の在庫部品があります';
    html += ' (<strong>' + data.seiban_count + '</strong> 製番)';
    html += '</div>';

    // 製番ごとにグループ表示
    const bySeiban = data.by_seiban || {};
    const columns = data.columns;

    for (const seiban of Object.keys(bySeiban).sort()) {
        const items = bySeiban[seiban];
        html += '<div style="margin-bottom:15px; border:1px solid #90caf9; border-radius:6px; overflow:hidden;">';
        html += '<div style="background:#e3f2fd; padding:8px 12px; font-weight:bold;">';
        html += '製番: ' + escapeForTable(seiban) + ' (' + items.length + '件)';
        html += '</div>';

        html += '<table style="width:100%; border-collapse:collapse; font-size:0.85em;">';
        html += '<thead><tr style="background:#f5f5f5;">';
        // columns[0]は製番なのでスキップ
        for (let i = 1; i < columns.length; i++) {
            html += '<th style="padding:6px 8px; border:1px solid #ddd; text-align:left;">' + escapeForTable(columns[i]) + '</th>';
        }
        html += '</tr></thead><tbody>';

        for (const row of items) {
            html += '<tr>';
            for (let i = 1; i < row.length; i++) {
                const val = row[i];
                const display = (val === null || val === '') ? '-' : escapeForTable(String(val));
                html += '<td style="padding:4px 8px; border:1px solid #eee;">' + display + '</td>';
            }
            html += '</tr>';
        }
        html += '</tbody></table></div>';
    }

    container.innerHTML = html;
}

// ========== 0ZAIKO 手配リスト検索 ==========
async function search0ZaikoTehai() {
    const resultDiv = document.getElementById('zaikoTehaiResult');
    resultDiv.innerHTML = '<div style="text-align:center; padding:20px; color:#666;">0ZAIKO手配リストを検索中...</div>';

    try {
        const res = await fetch('/api/across-db/0zaiko');
        const data = await res.json();

        if (data.error) {
            resultDiv.innerHTML = '<div class="alert alert-danger">' + escapeForTable(data.error) + '</div>';
            return;
        }

        if (!data.success) {
            resultDiv.innerHTML = '<div class="alert alert-danger">検索失敗</div>';
            return;
        }

        render0ZaikoResult(data, resultDiv);

    } catch (e) {
        resultDiv.innerHTML = '<div class="alert alert-danger">検索失敗: ' + e + '</div>';
    }
}

function render0ZaikoResult(data, container) {
    if (!data.rows || data.rows.length === 0) {
        container.innerHTML = '<div class="alert alert-info">0ZAIKO（在庫発注用製番）の手配データはありません</div>';
        return;
    }

    let html = '<div style="margin-bottom:10px;">';
    html += '<strong style="color:#ff9800;">' + data.count + '件</strong> の在庫発注データがあります';
    html += '</div>';

    html += '<div style="overflow-x:auto; max-height:400px; overflow-y:auto; border:1px solid #ddd; border-radius:5px;">';
    html += '<table style="width:100%; border-collapse:collapse; font-size:0.85em; white-space:nowrap;">';

    // ヘッダー
    html += '<thead><tr style="background:#fff3e0; position:sticky; top:0;">';
    html += '<th style="padding:6px 8px; border:1px solid #ddd; background:#ffe0b2;">#</th>';
    for (const col of data.columns) {
        if (col === '製番') continue; // 0ZAIKOは固定なのでスキップ
        html += '<th style="padding:6px 8px; border:1px solid #ddd; background:#ffe0b2;">' + escapeForTable(col) + '</th>';
    }
    html += '</tr></thead><tbody>';

    // データ
    for (let i = 0; i < data.rows.length; i++) {
        const row = data.rows[i];
        const bgColor = i % 2 === 0 ? '#fff' : '#fff8e1';
        html += '<tr style="background:' + bgColor + ';">';
        html += '<td style="padding:4px 8px; border:1px solid #eee; color:#999; text-align:center;">' + (i + 1) + '</td>';
        for (let j = 0; j < row.length; j++) {
            if (data.columns[j] === '製番') continue;
            const val = row[j];
            const display = (val === null || val === '') ? '-' : escapeForTable(String(val));
            html += '<td style="padding:4px 8px; border:1px solid #eee;">' + display + '</td>';
        }
        html += '</tr>';
    }
    html += '</tbody></table></div>';

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
    const mergeInput = document.getElementById('acrossMergeSeiban');
    if (mergeInput) {
        mergeInput.addEventListener('keypress', function(e) {
            if (e.key === 'Enter') executeMergeTest();
        });
    }
    const mihatchuInput = document.getElementById('acrossMihatchuSeiban');
    if (mihatchuInput) {
        mihatchuInput.addEventListener('keypress', function(e) {
            if (e.key === 'Enter') searchMihatchu();
        });
    }

    // 自動更新チェック
    autoCheckDbUpdates();
});
