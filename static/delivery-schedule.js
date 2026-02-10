// ========================================
// 納品予定表（任意開始日 + 1週間）
// ========================================

let deliveryScheduleData = null;

// 現在のフィルター状態
let dsFilters = {
    seiban: '',
    unit: '',
    supplier: '',
    date: ''
};

// 納品予定を読み込み
async function loadDeliverySchedule() {
    const container = document.getElementById('deliveryScheduleContent');
    if (!container) return;

    container.innerHTML = '<p style="text-align: center; padding: 20px; color: #6c757d;">読み込み中...</p>';

    try {
        let url = '/api/delivery-schedule';
        const startDateInput = document.getElementById('deliveryStartDate');
        if (startDateInput && startDateInput.value) {
            url += '?start_date=' + startDateInput.value;
            console.log('納品予定: 開始日指定 =', startDateInput.value);
        } else {
            console.log('納品予定: 開始日なし（今日から）');
        }

        const response = await fetch(url);
        const data = await response.json();

        if (!data.success) {
            container.innerHTML = `<p style="color: #dc3545;">エラー: ${data.error}</p>`;
            return;
        }

        deliveryScheduleData = data;
        renderDeliverySchedule(data);
    } catch (error) {
        container.innerHTML = `<p style="color: #dc3545;">読み込みエラー: ${error}</p>`;
    }
}

// 開始日を今日にリセット
function resetDeliveryStartDate() {
    const input = document.getElementById('deliveryStartDate');
    if (input) {
        input.value = '';
    }
    loadDeliverySchedule();
}

// 納品予定の受入トグル（行単位でインライン更新）
async function toggleDeliveryReceive(detailId, btnElement) {
    if (!btnElement) return;
    const origText = btnElement.textContent;
    btnElement.disabled = true;
    btnElement.textContent = '...';

    try {
        const response = await fetch(`/api/detail/${detailId}/toggle-receive`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' }
        });
        const result = await response.json();

        if (result.success) {
            // 行をインライン更新（全体リロードしない）
            const row = btnElement.closest('tr');
            if (!row) { loadDeliverySchedule(); return; }

            const isNowReceived = result.is_received;
            const nextRow = row.nextElementSibling;

            // キャッシュデータも更新
            if (deliveryScheduleData) {
                deliveryScheduleData.days.forEach(day => {
                    day.items.forEach(item => {
                        if (item.detail_id === detailId) {
                            item.is_received = isNowReceived;
                        }
                    });
                });
            }

            // 行の背景色を更新
            row.style.background = isNowReceived ? '#d4edda' : '';

            // ボタンを切替
            if (isNowReceived) {
                btnElement.style.background = '#fff';
                btnElement.style.color = '#dc3545';
                btnElement.style.border = '1px solid #dc3545';
                btnElement.textContent = '取消';
            } else {
                btnElement.style.background = '#28a745';
                btnElement.style.color = '#fff';
                btnElement.style.border = '1px solid #28a745';
                btnElement.textContent = '受入';
            }
            btnElement.disabled = false;

            // 日ヘッダーの受入カウントを更新
            updateDayHeaderCounts();
        } else {
            alert('エラー: ' + (result.error || '受入処理に失敗'));
            btnElement.disabled = false;
            btnElement.textContent = origText;
        }
    } catch (error) {
        alert('通信エラー: ' + error);
        btnElement.disabled = false;
        btnElement.textContent = origText;
    }
}

// 日ヘッダーの受入カウントを更新
function updateDayHeaderCounts() {
    if (!deliveryScheduleData) return;
    deliveryScheduleData.days.forEach(day => {
        const received = day.items.filter(i => i.is_received).length;
        day.received = received;
        const countEl = document.getElementById('deliveryDayCount_' + day.date);
        if (countEl) {
            countEl.textContent = `${received}/${day.total} 受入済`;
        }
    });
}

// NEXT処理ステップを生成
function buildNextStepsHtml(item) {
    if (!item.next_steps || item.next_steps.length === 0) return '';

    let steps = '';
    item.next_steps.forEach((step, idx) => {
        if (idx > 0) steps += ' → ';
        steps += step.supplier;
        if (step.is_mekki) {
            steps += ' → <span style="color: #dc3545; font-weight: bold;">⚠️メッキ出</span>';
        }
    });
    steps += ' → 仕分 → 完了';

    return `<tr class="ds-next-row" style="background: #f0f8ff; border-bottom: 1px solid #eee;">
        <td colspan="9" style="padding: 3px 10px 3px 30px; font-size: 0.82em; color: #555;">
            <span style="background: #17a2b8; color: white; padding: 1px 6px; border-radius: 3px; font-size: 0.85em; margin-right: 5px;">NEXT</span>
            ${steps}
        </td>
    </tr>`;
}

// フィルター変更時
function applyDeliveryFilters() {
    dsFilters.seiban = (document.getElementById('dsFilterSeiban')?.value || '').toLowerCase();
    dsFilters.unit = (document.getElementById('dsFilterUnit')?.value || '').toLowerCase();
    dsFilters.supplier = (document.getElementById('dsFilterSupplier')?.value || '').toLowerCase();
    dsFilters.date = document.getElementById('dsFilterDate')?.value || '';
    filterDeliveryRows();
}

// フィルタークリア
function clearDeliveryFilters() {
    ['dsFilterSeiban', 'dsFilterUnit', 'dsFilterSupplier', 'dsFilterDate'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.value = '';
    });
    dsFilters = { seiban: '', unit: '', supplier: '', date: '' };
    filterDeliveryRows();
}

// 行レベルのフィルタリング
function filterDeliveryRows() {
    const rows = document.querySelectorAll('tr.ds-item-row');
    rows.forEach(row => {
        const seiban = (row.dataset.seiban || '').toLowerCase();
        const unit = (row.dataset.unit || '').toLowerCase();
        const supplier = (row.dataset.supplier || '').toLowerCase();
        const rowDate = row.dataset.date || '';

        const match =
            (!dsFilters.seiban || seiban.includes(dsFilters.seiban)) &&
            (!dsFilters.unit || unit.includes(dsFilters.unit)) &&
            (!dsFilters.supplier || supplier.includes(dsFilters.supplier)) &&
            (!dsFilters.date || rowDate === dsFilters.date);

        row.style.display = match ? '' : 'none';

        // NEXT行も連動
        const nextRow = row.nextElementSibling;
        if (nextRow && nextRow.classList.contains('ds-next-row')) {
            nextRow.style.display = match ? '' : 'none';
        }
    });

    // 日付フィルター時: フィルター対象外の日セクションを非表示
    if (dsFilters.date && deliveryScheduleData) {
        deliveryScheduleData.days.forEach(day => {
            const daySection = document.getElementById('deliveryDay_' + day.date);
            const dayHeader = daySection?.parentElement;
            if (dayHeader) {
                if (day.date === dsFilters.date) {
                    dayHeader.style.display = '';
                    daySection.style.display = 'block';
                } else {
                    dayHeader.style.display = 'none';
                }
            }
        });
    } else if (deliveryScheduleData) {
        // 日付フィルターなし→全日セクション表示
        deliveryScheduleData.days.forEach(day => {
            const daySection = document.getElementById('deliveryDay_' + day.date);
            const dayHeader = daySection?.parentElement;
            if (dayHeader) {
                dayHeader.style.display = '';
            }
        });
    }
}

// 表示中の未受入アイテムを一括受入
async function batchReceiveFiltered() {
    // 表示中で未受入の行を収集
    const rows = document.querySelectorAll('tr.ds-item-row');
    const targets = [];
    rows.forEach(row => {
        if (row.style.display === 'none') return;
        const btn = row.querySelector('button');
        if (!btn) return;
        // 受入ボタン（緑）のみ対象、取消ボタン（白/赤）はスキップ
        if (btn.textContent.trim() === '受入') {
            const detailId = parseInt(row.dataset.detailId);
            if (detailId) targets.push({ detailId, btn, row });
        }
    });

    if (targets.length === 0) {
        alert('受入対象のアイテムがありません');
        return;
    }

    const hasFilter = dsFilters.seiban || dsFilters.unit || dsFilters.supplier;
    const msg = hasFilter
        ? `フィルター中の未受入 ${targets.length}件 を一括受入しますか？`
        : `表示中の未受入 ${targets.length}件 を一括受入しますか？\n（フィルターで絞り込むことを推奨します）`;

    if (!confirm(msg)) return;

    const batchBtn = document.getElementById('dsBatchReceiveBtn');
    if (batchBtn) {
        batchBtn.disabled = true;
        batchBtn.textContent = `処理中... 0/${targets.length}`;
    }

    let successCount = 0;
    let errorCount = 0;

    for (let i = 0; i < targets.length; i++) {
        const { detailId, btn, row } = targets[i];
        try {
            const response = await fetch(`/api/detail/${detailId}/toggle-receive`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' }
            });
            const result = await response.json();

            if (result.success && result.is_received) {
                // インライン更新
                row.style.background = '#d4edda';
                btn.style.background = '#fff';
                btn.style.color = '#dc3545';
                btn.style.border = '1px solid #dc3545';
                btn.textContent = '取消';

                // キャッシュ更新
                if (deliveryScheduleData) {
                    deliveryScheduleData.days.forEach(day => {
                        day.items.forEach(item => {
                            if (item.detail_id === detailId) {
                                item.is_received = true;
                            }
                        });
                    });
                }
                successCount++;
            } else {
                errorCount++;
            }
        } catch (e) {
            errorCount++;
        }

        if (batchBtn) {
            batchBtn.textContent = `処理中... ${i + 1}/${targets.length}`;
        }
    }

    // 日ヘッダーカウント更新
    updateDayHeaderCounts();

    if (batchBtn) {
        batchBtn.disabled = false;
        batchBtn.textContent = '一括受入';
    }

    if (errorCount > 0) {
        alert(`完了: ${successCount}件 受入成功 / ${errorCount}件 エラー`);
    }
}

// フィルター用セレクトボックスの選択肢を構築
function buildFilterOptions(data) {
    const seibans = new Set();
    const units = new Set();
    const suppliers = new Set();
    const dates = [];

    data.days.forEach(day => {
        dates.push({ value: day.date, label: day.display_date, isToday: day.is_today });
        day.items.forEach(item => {
            if (item.seiban) seibans.add(item.seiban);
            if (item.unit) units.add(item.unit);
            if (item.supplier) suppliers.add(item.supplier);
        });
    });

    return { seibans: [...seibans].sort(), units: [...units].sort(), suppliers: [...suppliers].sort(), dates };
}

// 納品予定を描画
function renderDeliverySchedule(data) {
    const container = document.getElementById('deliveryScheduleContent');
    if (!container) return;

    if (data.days.length === 0) {
        container.innerHTML = '<p style="text-align: center; padding: 20px; color: #6c757d;">指定期間の納品予定はありません</p>';
        return;
    }

    const opts = buildFilterOptions(data);
    let html = '';

    // サマリーバー
    const todayData = data.days.find(d => d.is_today);
    const todayCount = todayData ? todayData.total : 0;
    const todayReceived = todayData ? todayData.received : 0;
    const summary = data.summary || {};
    html += `<div style="display: flex; gap: 15px; margin-bottom: 15px; flex-wrap: wrap;">
        <div style="flex: 1; min-width: 150px; background: ${todayCount > 0 ? '#fff3cd' : '#d4edda'}; padding: 12px 18px; border-radius: 8px; border-left: 4px solid ${todayCount > 0 ? '#ffc107' : '#28a745'};">
            <div style="font-size: 0.85em; color: #666;">今日の納品</div>
            <div style="font-size: 1.8em; font-weight: bold; color: ${todayCount > 0 ? '#856404' : '#155724'};">${todayCount}件</div>
            <div style="font-size: 0.8em; color: #888;">受入済: ${todayReceived}件</div>
        </div>
        <div style="flex: 1; min-width: 150px; background: #e8f4ff; padding: 12px 18px; border-radius: 8px; border-left: 4px solid #007bff;">
            <div style="font-size: 0.85em; color: #666;">期間合計</div>
            <div style="font-size: 1.8em; font-weight: bold; color: #004085;">${data.total_items}件</div>
            <div style="font-size: 0.8em; color: #888;">${data.days.length}日間</div>
        </div>
        <div style="flex: 1; min-width: 150px; background: #f8f9fa; padding: 12px 18px; border-radius: 8px; border-left: 4px solid #6c757d;">
            <div style="font-size: 0.85em; color: #666;">製番</div>
            <div style="font-size: 1.2em; font-weight: bold; color: #333;">${summary.seiban_count || 0}件</div>
            <div style="font-size: 0.75em; color: #888; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${(summary.seibans || []).join(', ')}">${(summary.seibans || []).join(', ')}</div>
        </div>
        <div style="flex: 1; min-width: 150px; background: #f8f9fa; padding: 12px 18px; border-radius: 8px; border-left: 4px solid #6c757d;">
            <div style="font-size: 0.85em; color: #666;">ユニット</div>
            <div style="font-size: 1.2em; font-weight: bold; color: #333;">${summary.unit_count || 0}種</div>
            <div style="font-size: 0.75em; color: #888; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${(summary.units || []).join(', ')}">${(summary.units || []).join(', ')}</div>
        </div>
        <div style="flex: 1; min-width: 150px; background: #f8f9fa; padding: 12px 18px; border-radius: 8px; border-left: 4px solid #6c757d;">
            <div style="font-size: 0.85em; color: #666;">仕入先</div>
            <div style="font-size: 1.2em; font-weight: bold; color: #333;">${summary.supplier_count || 0}社</div>
            <div style="font-size: 0.75em; color: #888; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${(summary.suppliers || []).join(', ')}">${(summary.suppliers || []).join(', ')}</div>
        </div>
    </div>`;

    // フィルターバー
    html += `<div style="display:flex; gap:8px; align-items:center; flex-wrap:wrap; margin-bottom:12px; padding:8px 12px; background:#f8f9fa; border-radius:8px;">
        <span style="font-weight:bold; font-size:0.9em;">絞り込み:</span>
        <select id="dsFilterDate" onchange="applyDeliveryFilters()" style="padding:4px 8px; border:1px solid #ccc; border-radius:4px; font-size:0.88em;">
            <option value="">全日付</option>
            ${opts.dates.map(d => `<option value="${d.value}">${d.label}${d.isToday ? ' [今日]' : ''}</option>`).join('')}
        </select>
        <select id="dsFilterSeiban" onchange="applyDeliveryFilters()" style="padding:4px 8px; border:1px solid #ccc; border-radius:4px; font-size:0.88em;">
            <option value="">全製番</option>
            ${opts.seibans.map(s => `<option value="${s}">${s}</option>`).join('')}
        </select>
        <select id="dsFilterUnit" onchange="applyDeliveryFilters()" style="padding:4px 8px; border:1px solid #ccc; border-radius:4px; font-size:0.88em;">
            <option value="">全ユニット</option>
            ${opts.units.map(u => `<option value="${u}">${u}</option>`).join('')}
        </select>
        <select id="dsFilterSupplier" onchange="applyDeliveryFilters()" style="padding:4px 8px; border:1px solid #ccc; border-radius:4px; font-size:0.88em;">
            <option value="">全仕入先</option>
            ${opts.suppliers.map(s => `<option value="${s}">${s}</option>`).join('')}
        </select>
        <button onclick="clearDeliveryFilters()" style="padding:4px 10px; border:1px solid #ccc; border-radius:4px; background:#fff; cursor:pointer; font-size:0.88em;">クリア</button>
        <div style="margin-left:auto;">
            <button id="dsBatchReceiveBtn" onclick="batchReceiveFiltered()" style="padding:5px 14px; border:1px solid #28a745; border-radius:4px; background:#28a745; color:#fff; cursor:pointer; font-size:0.88em; font-weight:bold;">一括受入</button>
        </div>
    </div>`;

    // 日ごとのセクション
    data.days.forEach(day => {
        const bgColor = day.is_today ? '#fffbea' : day.is_weekend ? '#fff0f0' : '#ffffff';
        const borderColor = day.is_today ? '#ffc107' : day.is_weekend ? '#ffcccc' : '#dee2e6';
        const todayBadge = day.is_today ? '<span style="background: #ffc107; color: #333; padding: 2px 8px; border-radius: 10px; font-size: 0.75em; font-weight: bold; margin-left: 8px;">TODAY</span>' : '';
        const weekendBadge = day.is_weekend ? '<span style="background: #ffcccc; color: #cc0000; padding: 2px 8px; border-radius: 10px; font-size: 0.75em; margin-left: 8px;">休日</span>' : '';

        html += `<div style="background: ${bgColor}; border: 1px solid ${borderColor}; border-radius: 8px; margin-bottom: 10px; overflow: hidden;">`;
        html += `<div style="padding: 10px 15px; background: ${day.is_today ? '#fff3cd' : '#f8f9fa'}; border-bottom: 1px solid ${borderColor}; display: flex; align-items: center; justify-content: space-between; cursor: pointer;" onclick="toggleDeliveryDay('${day.date}')">`;
        html += `<div><strong style="font-size: 1.1em;">${day.display_date}</strong>${todayBadge}${weekendBadge}</div>`;
        html += `<div style="display: flex; align-items: center; gap: 10px;">
            <span id="deliveryDayCount_${day.date}" style="font-size: 0.9em; color: #666;">${day.received}/${day.total} 受入済</span>
            <span style="background: ${day.received === day.total ? '#28a745' : '#007bff'}; color: white; padding: 2px 10px; border-radius: 12px; font-weight: bold;">${day.total}件</span>
            <span id="deliveryDayArrow_${day.date}" style="transition: transform 0.2s;">&#9660;</span>
        </div>`;
        html += `</div>`;

        // 明細テーブル（今日はデフォルト展開、今日がない場合は最初の日を展開）
        const hasToday = data.days.some(d => d.is_today);
        const isFirstDay = (data.days.indexOf(day) === 0);
        const display = (day.is_today || (!hasToday && isFirstDay)) ? 'block' : 'none';
        html += `<div id="deliveryDay_${day.date}" style="display: ${display}; padding: 0;">`;
        html += `<table style="width: 100%; border-collapse: collapse; font-size: 0.88em;">`;
        html += `<thead><tr style="background: #f1f3f5;">
            <th style="padding: 6px 10px; text-align: center; border-bottom: 1px solid #ddd; width: 70px;">受入</th>
            <th style="padding: 6px 10px; text-align: left; border-bottom: 1px solid #ddd;">製番</th>
            <th style="padding: 6px 10px; text-align: left; border-bottom: 1px solid #ddd;">ユニット</th>
            <th style="padding: 6px 10px; text-align: left; border-bottom: 1px solid #ddd;">仕入先</th>
            <th style="padding: 6px 10px; text-align: left; border-bottom: 1px solid #ddd;">品名</th>
            <th style="padding: 6px 10px; text-align: left; border-bottom: 1px solid #ddd;">仕様１</th>
            <th style="padding: 6px 10px; text-align: left; border-bottom: 1px solid #ddd;">手配区分</th>
            <th style="padding: 6px 10px; text-align: right; border-bottom: 1px solid #ddd;">数量</th>
            <th style="padding: 6px 10px; text-align: left; border-bottom: 1px solid #ddd;">発注番号</th>
        </tr></thead><tbody>`;

        // 製番でグループ化して表示
        const grouped = {};
        day.items.forEach(item => {
            const key = item.seiban;
            if (!grouped[key]) grouped[key] = [];
            grouped[key].push(item);
        });

        Object.keys(grouped).sort().forEach(seiban => {
            grouped[seiban].forEach(item => {
                const receivedStyle = item.is_received
                    ? 'background: #d4edda;'
                    : '';
                const btn = item.is_received
                    ? `<button onclick="toggleDeliveryReceive(${item.detail_id}, this)" style="padding: 2px 8px; border: 1px solid #dc3545; background: #fff; color: #dc3545; border-radius: 4px; cursor: pointer; font-size: 0.82em; white-space: nowrap;">取消</button>`
                    : `<button onclick="toggleDeliveryReceive(${item.detail_id}, this)" style="padding: 2px 8px; border: 1px solid #28a745; background: #28a745; color: #fff; border-radius: 4px; cursor: pointer; font-size: 0.82em; white-space: nowrap;">受入</button>`;

                // /api/open-cad/ エンドポイントでCADファイルを表示
                let spec1Cell = item.spec1;
                if (item.cad_link) {
                    spec1Cell = `<a href="/api/open-cad/${item.detail_id}" target="_blank" style="color: #0000FF; text-decoration: underline;" title="${item.spec1}">${item.spec1}</a>`;
                }

                // 加工用ブランクの場合、手配区分にバッジ表示
                const orderTypeCell = item.is_blank
                    ? `<span style="background: #17a2b8; color: white; padding: 1px 5px; border-radius: 3px; font-size: 0.9em;">${item.order_type}</span>`
                    : item.order_type;

                html += `<tr class="ds-item-row" data-detail-id="${item.detail_id}" data-seiban="${item.seiban}" data-unit="${item.unit}" data-supplier="${item.supplier}" data-date="${day.date}" style="${receivedStyle} border-bottom: 1px solid #eee;">
                    <td style="padding: 5px 10px; text-align: center;">${btn}</td>
                    <td style="padding: 5px 10px; white-space: nowrap;">${item.seiban}</td>
                    <td style="padding: 5px 10px; font-size: 0.9em;">${item.unit}</td>
                    <td style="padding: 5px 10px;">${item.supplier}</td>
                    <td style="padding: 5px 10px;">${item.item_name}</td>
                    <td style="padding: 5px 10px; font-size: 0.85em;">${spec1Cell}</td>
                    <td style="padding: 5px 10px; font-size: 0.85em;">${orderTypeCell}</td>
                    <td style="padding: 5px 10px; text-align: right;">${item.quantity} ${item.unit_measure}</td>
                    <td style="padding: 5px 10px;">${item.order_number}</td>
                </tr>`;

                // 加工用ブランクの場合、NEXT処理ステップを表示
                if (item.is_blank) {
                    if (item.next_steps && item.next_steps.length > 0) {
                        html += buildNextStepsHtml(item);
                    } else {
                        html += `<tr class="ds-next-row" style="background: #fff8e1; border-bottom: 1px solid #eee;">
                            <td colspan="9" style="padding: 3px 10px 3px 30px; font-size: 0.82em; color: #856404;">
                                <span style="background: #ffc107; color: #333; padding: 1px 6px; border-radius: 3px; font-size: 0.85em; margin-right: 5px;">NEXT</span>
                                → 追加工待ち → 仕分 → 完了
                            </td>
                        </tr>`;
                    }
                }
            });
        });

        html += `</tbody></table></div></div>`;
    });

    container.innerHTML = html;
}

// 日ごとの折りたたみ切替
function toggleDeliveryDay(dateKey) {
    const el = document.getElementById('deliveryDay_' + dateKey);
    const arrow = document.getElementById('deliveryDayArrow_' + dateKey);
    if (!el) return;

    if (el.style.display === 'none') {
        el.style.display = 'block';
        if (arrow) arrow.style.transform = 'rotate(0deg)';
    } else {
        el.style.display = 'none';
        if (arrow) arrow.style.transform = 'rotate(-90deg)';
    }
}

// 納品予定を印刷
function printDeliverySchedule() {
    if (!deliveryScheduleData || deliveryScheduleData.days.length === 0) {
        alert('納品予定データがありません');
        return;
    }

    const data = deliveryScheduleData;
    const now = new Date().toLocaleString('ja-JP');

    let tableRows = '';
    data.days.forEach(day => {
        // 日付ヘッダー行
        const todayMark = day.is_today ? ' [TODAY]' : '';
        tableRows += `<tr style="background: ${day.is_today ? '#fff3cd' : '#e9ecef'};">
            <td colspan="9" style="padding: 8px; font-weight: bold; font-size: 1.1em; border: 1px solid #ccc;">
                ${day.display_date}${todayMark} - ${day.total}件 (受入済: ${day.received})
            </td></tr>`;

        day.items.forEach(item => {
            const mark = item.is_received ? '&#10003;' : '';
            tableRows += `<tr style="${item.is_received ? 'background: #e8f5e9;' : ''}">
                <td style="padding: 4px 8px; border: 1px solid #ccc; text-align: center;">${mark}</td>
                <td style="padding: 4px 8px; border: 1px solid #ccc;">${item.seiban}</td>
                <td style="padding: 4px 8px; border: 1px solid #ccc;">${item.unit}</td>
                <td style="padding: 4px 8px; border: 1px solid #ccc;">${item.supplier}</td>
                <td style="padding: 4px 8px; border: 1px solid #ccc;">${item.item_name}</td>
                <td style="padding: 4px 8px; border: 1px solid #ccc;">${item.spec1}</td>
                <td style="padding: 4px 8px; border: 1px solid #ccc;">${item.order_type}</td>
                <td style="padding: 4px 8px; border: 1px solid #ccc; text-align: right;">${item.quantity} ${item.unit_measure}</td>
                <td style="padding: 4px 8px; border: 1px solid #ccc;">${item.order_number}</td>
            </tr>`;

            // 印刷時もNEXTステップを表示
            if (item.is_blank && item.next_steps && item.next_steps.length > 0) {
                let steps = '';
                item.next_steps.forEach((step, idx) => {
                    if (idx > 0) steps += ' → ';
                    steps += step.supplier;
                    if (step.is_mekki) steps += ' → ⚠メッキ出';
                });
                steps += ' → 仕分 → 完了';
                tableRows += `<tr style="background: #f0f8ff;">
                    <td colspan="9" style="padding: 3px 8px 3px 25px; border: 1px solid #ccc; font-size: 0.9em; color: #555;">
                        NEXT: ${steps}
                    </td>
                </tr>`;
            } else if (item.is_blank) {
                tableRows += `<tr style="background: #fff8e1;">
                    <td colspan="9" style="padding: 3px 8px 3px 25px; border: 1px solid #ccc; font-size: 0.9em; color: #856404;">
                        NEXT: → 追加工待ち → 仕分 → 完了
                    </td>
                </tr>`;
            }
        });
    });

    const printWindow = window.open('', '_blank');
    printWindow.document.write(`<!DOCTYPE html><html><head><title>納品予定表</title>
        <style>
            @media print { @page { size: landscape; margin: 8mm; } body { margin: 0; } }
            body { font-family: 'Meiryo', sans-serif; font-size: 11px; }
            h2 { margin: 0 0 5px 0; }
            .info { font-size: 0.85em; color: #666; margin-bottom: 10px; }
            table { width: 100%; border-collapse: collapse; }
            th { background: #343a40; color: white; padding: 6px 8px; border: 1px solid #ccc; text-align: left; }
        </style></head><body>
        <h2>納品予定表</h2>
        <div class="info">印刷日時: ${now} ／ 合計: ${data.total_items}件</div>
        <table>
            <thead><tr>
                <th style="width:30px;">済</th><th>製番</th><th>ユニット</th><th>仕入先</th><th>品名</th><th>仕様１</th><th>手配区分</th><th>数量</th><th>発注番号</th>
            </tr></thead>
            <tbody>${tableRows}</tbody>
        </table>
        <script>window.onload = function() { window.print(); window.close(); };</script>
    </body></html>`);
    printWindow.document.close();
}


// ========================================
// 発注DB直接取得 - 納品予定
// ========================================

let dbDeliveryData = null;

// DB納品予定のフィルター状態
let dbDsFilters = {
    seiban: '',
    supplier: '',
    orderType: '',
    date: ''
};

// DB開始日リセット
function resetDbDeliveryDate() {
    const input = document.getElementById('dbDeliveryStartDate');
    if (input) {
        input.value = '';
    }
    loadDbDeliverySchedule();
}

// 発注DBから納品予定を読み込み
async function loadDbDeliverySchedule() {
    const container = document.getElementById('dbDeliveryScheduleContent');
    if (!container) return;

    container.innerHTML = '<p style="text-align: center; padding: 20px; color: #6c757d;">読み込み中...</p>';

    try {
        let url = '/api/across-db/delivery-schedule';
        const startDateInput = document.getElementById('dbDeliveryStartDate');
        const params = new URLSearchParams();

        if (startDateInput && startDateInput.value) {
            params.append('start_date', startDateInput.value);
        }

        if (params.toString()) {
            url += '?' + params.toString();
        }

        const response = await fetch(url);
        const data = await response.json();

        if (!data.success) {
            container.innerHTML = `<p style="color: #dc3545;">エラー: ${data.error}</p>`;
            return;
        }

        dbDeliveryData = data;
        renderDbDeliverySchedule(data);
    } catch (error) {
        container.innerHTML = `<p style="color: #dc3545;">読み込みエラー: ${error}</p>`;
    }
}

// DB納品予定用フィルターオプションを構築
function buildDbFilterOptions(data) {
    const seibans = new Set();
    const suppliers = new Set();
    const orderTypes = new Set();
    const dates = [];

    const sortedDates = Object.keys(data.days).sort();
    const today = new Date().toISOString().split('T')[0];

    sortedDates.forEach(dateKey => {
        const items = data.days[dateKey];
        const isToday = dateKey === today;
        // 日付表示用フォーマット
        const d = new Date(dateKey);
        const weekdays = ['日', '月', '火', '水', '木', '金', '土'];
        const displayDate = `${d.getMonth() + 1}/${d.getDate()}(${weekdays[d.getDay()]})`;
        dates.push({ value: dateKey, label: displayDate, isToday, count: items.length });

        items.forEach(item => {
            if (item['製番']) seibans.add(item['製番']);
            if (item['仕入先']) suppliers.add(item['仕入先']);
            if (item['手配区分']) orderTypes.add(item['手配区分']);
        });
    });

    return {
        seibans: [...seibans].sort(),
        suppliers: [...suppliers].sort(),
        orderTypes: [...orderTypes].sort(),
        dates
    };
}

// DB納品予定フィルター適用
function applyDbDeliveryFilters() {
    dbDsFilters.seiban = (document.getElementById('dbDsFilterSeiban')?.value || '').toLowerCase();
    dbDsFilters.supplier = (document.getElementById('dbDsFilterSupplier')?.value || '').toLowerCase();
    dbDsFilters.orderType = (document.getElementById('dbDsFilterOrderType')?.value || '').toLowerCase();
    dbDsFilters.date = document.getElementById('dbDsFilterDate')?.value || '';
    filterDbDeliveryRows();
}

// DB納品予定フィルタークリア
function clearDbDeliveryFilters() {
    ['dbDsFilterSeiban', 'dbDsFilterSupplier', 'dbDsFilterOrderType', 'dbDsFilterDate'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.value = '';
    });
    dbDsFilters = { seiban: '', supplier: '', orderType: '', date: '' };
    filterDbDeliveryRows();
}

// DB納品予定の行フィルタリング
function filterDbDeliveryRows() {
    const rows = document.querySelectorAll('tr.dbds-item-row');
    rows.forEach(row => {
        const seiban = (row.dataset.seiban || '').toLowerCase();
        const supplier = (row.dataset.supplier || '').toLowerCase();
        const orderType = (row.dataset.ordertype || '').toLowerCase();
        const rowDate = row.dataset.date || '';

        const match =
            (!dbDsFilters.seiban || seiban.includes(dbDsFilters.seiban)) &&
            (!dbDsFilters.supplier || supplier.includes(dbDsFilters.supplier)) &&
            (!dbDsFilters.orderType || orderType.includes(dbDsFilters.orderType)) &&
            (!dbDsFilters.date || rowDate === dbDsFilters.date);

        row.style.display = match ? '' : 'none';
    });

    // 日付セクションの表示制御
    if (dbDeliveryData) {
        const sortedDates = Object.keys(dbDeliveryData.days).sort();
        sortedDates.forEach(dateKey => {
            const daySection = document.getElementById('dbDeliveryDay_' + dateKey);
            const dayContainer = daySection?.closest('.dbds-day-container');
            if (dayContainer) {
                if (dbDsFilters.date && dateKey !== dbDsFilters.date) {
                    dayContainer.style.display = 'none';
                } else {
                    dayContainer.style.display = '';
                    // 該当日の表示件数を更新
                    const visibleRows = daySection.querySelectorAll('tr.dbds-item-row:not([style*="display: none"])');
                    const countEl = document.getElementById('dbDeliveryDayCount_' + dateKey);
                    if (countEl) {
                        const totalRows = daySection.querySelectorAll('tr.dbds-item-row').length;
                        if (visibleRows.length === totalRows) {
                            countEl.textContent = `${totalRows}件`;
                        } else {
                            countEl.textContent = `${visibleRows.length}/${totalRows}件`;
                        }
                    }
                }
            }
        });
    }
}

// DB納品予定の日ごと折りたたみ切替
function toggleDbDeliveryDay(dateKey) {
    const el = document.getElementById('dbDeliveryDay_' + dateKey);
    const arrow = document.getElementById('dbDeliveryDayArrow_' + dateKey);
    if (!el) return;

    if (el.style.display === 'none') {
        el.style.display = 'block';
        if (arrow) arrow.style.transform = 'rotate(0deg)';
    } else {
        el.style.display = 'none';
        if (arrow) arrow.style.transform = 'rotate(-90deg)';
    }
}

// 発注DB納品予定を描画
function renderDbDeliverySchedule(data) {
    const container = document.getElementById('dbDeliveryScheduleContent');
    if (!container) return;

    if (data.total === 0) {
        container.innerHTML = '<p style="text-align: center; padding: 20px; color: #6c757d;">指定期間の納品予定はありません</p>';
        return;
    }

    const opts = buildDbFilterOptions(data);
    const today = new Date().toISOString().split('T')[0];
    let html = '';

    // サマリー
    const todayData = opts.dates.find(d => d.isToday);
    const todayCount = todayData ? todayData.count : 0;

    html += `<div style="display: flex; gap: 15px; margin-bottom: 15px; flex-wrap: wrap;">
        <div style="flex: 1; min-width: 150px; background: ${todayCount > 0 ? '#fff3cd' : '#d4edda'}; padding: 12px 18px; border-radius: 8px; border-left: 4px solid ${todayCount > 0 ? '#ffc107' : '#28a745'};">
            <div style="font-size: 0.85em; color: #666;">今日の納品</div>
            <div style="font-size: 1.8em; font-weight: bold; color: ${todayCount > 0 ? '#856404' : '#155724'};">${todayCount}件</div>
        </div>
        <div style="flex: 1; min-width: 150px; background: #e8f5e9; padding: 12px 18px; border-radius: 8px; border-left: 4px solid #28a745;">
            <div style="font-size: 0.85em; color: #666;">期間合計</div>
            <div style="font-size: 1.8em; font-weight: bold; color: #155724;">${data.total}件</div>
            <div style="font-size: 0.8em; color: #888;">${data.start_date} ～ ${data.end_date}</div>
        </div>
        <div style="flex: 1; min-width: 150px; background: #f8f9fa; padding: 12px 18px; border-radius: 8px; border-left: 4px solid #6c757d;">
            <div style="font-size: 0.85em; color: #666;">日数</div>
            <div style="font-size: 1.2em; font-weight: bold; color: #333;">${Object.keys(data.days).length}日</div>
        </div>
    </div>`;

    // フィルターバー
    html += `<div style="display:flex; gap:8px; align-items:center; flex-wrap:wrap; margin-bottom:12px; padding:8px 12px; background:#f8f9fa; border-radius:8px;">
        <span style="font-weight:bold; font-size:0.9em;">絞り込み:</span>
        <select id="dbDsFilterDate" onchange="applyDbDeliveryFilters()" style="padding:4px 8px; border:1px solid #ccc; border-radius:4px; font-size:0.88em;">
            <option value="">全日付</option>
            ${opts.dates.map(d => `<option value="${d.value}">${d.label}${d.isToday ? ' [今日]' : ''} (${d.count}件)</option>`).join('')}
        </select>
        <select id="dbDsFilterSeiban" onchange="applyDbDeliveryFilters()" style="padding:4px 8px; border:1px solid #ccc; border-radius:4px; font-size:0.88em;">
            <option value="">全製番</option>
            ${opts.seibans.map(s => `<option value="${s}">${s}</option>`).join('')}
        </select>
        <select id="dbDsFilterSupplier" onchange="applyDbDeliveryFilters()" style="padding:4px 8px; border:1px solid #ccc; border-radius:4px; font-size:0.88em;">
            <option value="">全仕入先</option>
            ${opts.suppliers.map(s => `<option value="${s}">${s}</option>`).join('')}
        </select>
        <select id="dbDsFilterOrderType" onchange="applyDbDeliveryFilters()" style="padding:4px 8px; border:1px solid #ccc; border-radius:4px; font-size:0.88em;">
            <option value="">全手配区分</option>
            ${opts.orderTypes.map(t => `<option value="${t}">${t}</option>`).join('')}
        </select>
        <button onclick="clearDbDeliveryFilters()" style="padding:4px 10px; border:1px solid #ccc; border-radius:4px; background:#fff; cursor:pointer; font-size:0.88em;">クリア</button>
    </div>`;

    // 日付順にソートして日ごとのセクションを作成
    const sortedDates = Object.keys(data.days).sort();

    sortedDates.forEach((dateKey, idx) => {
        const items = data.days[dateKey];
        const isToday = dateKey === today;

        // 日付表示用フォーマット
        const d = new Date(dateKey);
        const weekdays = ['日', '月', '火', '水', '木', '金', '土'];
        const dayOfWeek = d.getDay();
        const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
        const displayDate = `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}(${weekdays[dayOfWeek]})`;

        const bgColor = isToday ? '#fffbea' : isWeekend ? '#fff0f0' : '#ffffff';
        const borderColor = isToday ? '#ffc107' : isWeekend ? '#ffcccc' : '#dee2e6';
        const todayBadge = isToday ? '<span style="background: #ffc107; color: #333; padding: 2px 8px; border-radius: 10px; font-size: 0.75em; font-weight: bold; margin-left: 8px;">TODAY</span>' : '';
        const weekendBadge = isWeekend ? '<span style="background: #ffcccc; color: #cc0000; padding: 2px 8px; border-radius: 10px; font-size: 0.75em; margin-left: 8px;">休日</span>' : '';

        // 今日または最初の日を展開、他は折りたたみ
        const hasToday = sortedDates.includes(today);
        const display = (isToday || (!hasToday && idx === 0)) ? 'block' : 'none';
        const arrowRotate = display === 'block' ? '0deg' : '-90deg';

        html += `<div class="dbds-day-container" style="background: ${bgColor}; border: 1px solid ${borderColor}; border-radius: 8px; margin-bottom: 10px; overflow: hidden;">`;
        html += `<div style="padding: 10px 15px; background: ${isToday ? '#fff3cd' : '#f8f9fa'}; border-bottom: 1px solid ${borderColor}; display: flex; align-items: center; justify-content: space-between; cursor: pointer;" onclick="toggleDbDeliveryDay('${dateKey}')">`;
        html += `<div><strong style="font-size: 1.1em;">${displayDate}</strong>${todayBadge}${weekendBadge}</div>`;
        html += `<div style="display: flex; align-items: center; gap: 10px;">
            <span id="dbDeliveryDayCount_${dateKey}" style="background: #28a745; color: white; padding: 2px 10px; border-radius: 12px; font-weight: bold;">${items.length}件</span>
            <span id="dbDeliveryDayArrow_${dateKey}" style="transition: transform 0.2s; transform: rotate(${arrowRotate});">&#9660;</span>
        </div>`;
        html += `</div>`;

        // 明細テーブル
        html += `<div id="dbDeliveryDay_${dateKey}" style="display: ${display}; padding: 0;">`;
        html += `<table style="width: 100%; border-collapse: collapse; font-size: 0.88em;">
            <thead><tr style="background: #28a745; color: white;">
                <th style="padding: 6px 10px; text-align: left; border-bottom: 1px solid #1e7e34;">製番</th>
                <th style="padding: 6px 10px; text-align: left; border-bottom: 1px solid #1e7e34;">仕入先</th>
                <th style="padding: 6px 10px; text-align: left; border-bottom: 1px solid #1e7e34;">品名</th>
                <th style="padding: 6px 10px; text-align: left; border-bottom: 1px solid #1e7e34;">仕様１</th>
                <th style="padding: 6px 10px; text-align: left; border-bottom: 1px solid #1e7e34;">手配区分</th>
                <th style="padding: 6px 10px; text-align: right; border-bottom: 1px solid #1e7e34;">数量</th>
                <th style="padding: 6px 10px; text-align: left; border-bottom: 1px solid #1e7e34;">発注番号</th>
            </tr></thead><tbody>`;

        items.forEach((item, rowIdx) => {
            const rowBg = rowIdx % 2 === 0 ? '#ffffff' : '#f8f9fa';
            const spec1 = item['仕様１'] || '';

            // 仕様１がN**-で始まる場合はCADリンクを追加
            let spec1Cell = spec1 || '-';
            if (spec1 && /^N[A-Z]{2}-/.test(spec1)) {
                spec1Cell = `<a href="/api/open-cad-by-spec/${encodeURIComponent(spec1)}" target="_blank" style="color: #0000FF; text-decoration: underline;" title="${spec1}">${spec1}</a>`;
            }

            html += `<tr class="dbds-item-row" data-date="${dateKey}" data-seiban="${item['製番'] || ''}" data-supplier="${item['仕入先'] || ''}" data-ordertype="${item['手配区分'] || ''}" style="background: ${rowBg}; border-bottom: 1px solid #eee;">
                <td style="padding: 6px 10px; font-weight: bold;">${item['製番'] || '-'}</td>
                <td style="padding: 6px 10px;">${item['仕入先'] || '-'}</td>
                <td style="padding: 6px 10px;">${item['品名'] || '-'}</td>
                <td style="padding: 6px 10px; font-size: 0.9em;">${spec1Cell}</td>
                <td style="padding: 6px 10px; font-size: 0.9em;">${item['手配区分'] || '-'}</td>
                <td style="padding: 6px 10px; text-align: right;">${item['発注数'] || '-'} ${item['単位'] || ''}</td>
                <td style="padding: 6px 10px;">${item['発注番号'] || '-'}</td>
            </tr>`;
        });

        html += '</tbody></table></div></div>';
    });

    container.innerHTML = html;
}

// 発注DB納品予定を印刷
function printDbDeliverySchedule() {
    if (!dbDeliveryData || dbDeliveryData.total === 0) {
        alert('納品予定データがありません');
        return;
    }

    const data = dbDeliveryData;
    const now = new Date().toLocaleString('ja-JP');
    const today = new Date().toISOString().split('T')[0];
    const weekdays = ['日', '月', '火', '水', '木', '金', '土'];

    let tableRows = '';
    const sortedDates = Object.keys(data.days).sort();

    sortedDates.forEach(dateKey => {
        const items = data.days[dateKey];
        const d = new Date(dateKey);
        const displayDate = `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}(${weekdays[d.getDay()]})`;
        const isToday = dateKey === today;
        const todayMark = isToday ? ' [TODAY]' : '';

        // 日付ヘッダー行
        tableRows += `<tr style="background: ${isToday ? '#fff3cd' : '#e9ecef'};">
            <td colspan="7" style="padding: 8px; font-weight: bold; font-size: 1.1em; border: 1px solid #ccc;">
                ${displayDate}${todayMark} - ${items.length}件
            </td></tr>`;

        items.forEach(item => {
            tableRows += `<tr>
                <td style="padding: 4px 8px; border: 1px solid #ccc;">${item['製番'] || '-'}</td>
                <td style="padding: 4px 8px; border: 1px solid #ccc;">${item['仕入先'] || '-'}</td>
                <td style="padding: 4px 8px; border: 1px solid #ccc;">${item['品名'] || '-'}</td>
                <td style="padding: 4px 8px; border: 1px solid #ccc;">${item['仕様１'] || '-'}</td>
                <td style="padding: 4px 8px; border: 1px solid #ccc;">${item['手配区分'] || '-'}</td>
                <td style="padding: 4px 8px; border: 1px solid #ccc; text-align: right;">${item['発注数'] || '-'} ${item['単位'] || ''}</td>
                <td style="padding: 4px 8px; border: 1px solid #ccc;">${item['発注番号'] || '-'}</td>
            </tr>`;
        });
    });

    const printWindow = window.open('', '_blank');
    printWindow.document.write(`<!DOCTYPE html><html><head><title>納品予定表（発注DB）</title>
        <style>
            @media print { @page { size: landscape; margin: 8mm; } body { margin: 0; } }
            body { font-family: 'Meiryo', sans-serif; font-size: 11px; }
            h2 { margin: 0 0 5px 0; }
            .info { font-size: 0.85em; color: #666; margin-bottom: 10px; }
            table { width: 100%; border-collapse: collapse; }
            th { background: #28a745; color: white; padding: 6px 8px; border: 1px solid #ccc; text-align: left; }
        </style></head><body>
        <h2>納品予定表（発注DB）</h2>
        <div class="info">印刷日時: ${now} ／ 期間: ${data.start_date} ～ ${data.end_date} ／ 合計: ${data.total}件</div>
        <table>
            <thead><tr>
                <th>製番</th><th>仕入先</th><th>品名</th><th>仕様１</th><th>手配区分</th><th>数量</th><th>発注番号</th>
            </tr></thead>
            <tbody>${tableRows}</tbody>
        </table>
        <script>window.onload = function() { window.print(); window.close(); };</script>
    </body></html>`);
    printWindow.document.close();
}
