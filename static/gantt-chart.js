// ========================================
// ガントチャート専用スクリプト
// ========================================

// ========================================
// グローバル変数
// ========================================
let ganttChartInstance = null;
let allGanttData = [];
let selectedSeibansForGantt = new Set();

// ========================================
// グループ1: 基本機能（依存なし）
// ========================================

// ローディング表示制御
function showGanttLoading(show) {
    const overlay = document.getElementById('ganttLoadingOverlay');
    if (overlay) {
        overlay.style.display = show ? 'block' : 'none';
    }
}

// ローディング進捗更新
function updateGanttLoadingProgress(percent, message, detail) {
    const bar = document.getElementById('ganttLoadingBar');
    const percentText = document.getElementById('ganttLoadingPercent');
    const messageDiv = document.getElementById('ganttLoadingMessage');
    const detailDiv = document.getElementById('ganttLoadingDetail');
    
    if (bar) {
        bar.style.width = percent + '%';
    }
    if (percentText) {
        percentText.textContent = Math.round(percent) + '%';
    }
    if (messageDiv && message) {
        messageDiv.textContent = message;
    }
    if (detailDiv && detail) {
        detailDiv.textContent = detail;
    }
}

// 日付パース関数（YY/MM/DD、YYYY-MM-DD対応）
function parseDeliveryDate(dateStr) {
    if (!dateStr || dateStr === '-') return null;
    
    // YY/MM/DD形式（例: "25/10/14"）
    const yymmdd = dateStr.match(/^(\d{2})\/(\d{1,2})\/(\d{1,2})$/);
    if (yymmdd) {
        const year = 2000 + parseInt(yymmdd[1]);
        const month = parseInt(yymmdd[2]) - 1; // 0-indexed
        const day = parseInt(yymmdd[3]);
        return new Date(year, month, day);
    }
    
    // YYYY/MM/DD形式
    const yyyymmdd = dateStr.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
    if (yyyymmdd) {
        const year = parseInt(yyyymmdd[1]);
        const month = parseInt(yyyymmdd[2]) - 1;
        const day = parseInt(yyyymmdd[3]);
        return new Date(year, month, day);
    }
    
    // ISO形式やその他
    const date = new Date(dateStr);
    return isNaN(date) ? null : date;
}

// フィルタ件数を更新
function updateGanttFilterCount() {
    const total = document.querySelectorAll('.gantt-seiban-checkbox').length;
    const selected = document.querySelectorAll('.gantt-seiban-checkbox:checked').length;
    document.getElementById('ganttFilterCount').textContent = `${selected} / ${total} 製番選択中`;
}

// 表示切替
function toggleGanttChart() {
    const container = document.getElementById('ganttChartContainer');
    const icon = document.getElementById('ganttToggleIcon');
    const text = document.getElementById('ganttToggleText');
    
    if (container.style.display === 'none') {
        container.style.display = 'block';
        icon.textContent = '📉';
        text.textContent = '非表示';
    } else {
        container.style.display = 'none';
        icon.textContent = '📈';
        text.textContent = '表示';
    }
}

// ========================================
// グループ2: フィルタ機能
// ========================================

// フィルタUI初期化
function initializeGanttFilter(data) {
    // ユニークな製番を抽出
    const seibans = [...new Set(data.map(d => d.seiban))].sort();
    
    const listDiv = document.getElementById('ganttSeibanList');
    listDiv.innerHTML = '';
    
    seibans.forEach(seiban => {
        const checkbox = document.createElement('label');
        checkbox.style.cursor = 'pointer';
        checkbox.style.display = 'flex';
        checkbox.style.alignItems = 'center';
        checkbox.style.padding = '5px';
        checkbox.innerHTML = `
            <input type="checkbox" class="gantt-seiban-checkbox" value="${seiban}" 
                   ${selectedSeibansForGantt.has(seiban) ? 'checked' : ''}
                   style="margin-right: 5px;">
            <span>${seiban}</span>
        `;
        listDiv.appendChild(checkbox);
    });
    
    updateGanttFilterCount();
}

// 全て選択
function selectAllSeibansForGantt() {
    document.querySelectorAll('.gantt-seiban-checkbox').forEach(cb => {
        cb.checked = true;
        selectedSeibansForGantt.add(cb.value);
    });
    updateGanttFilterCount();
}

// 全て解除
function deselectAllSeibansForGantt() {
    document.querySelectorAll('.gantt-seiban-checkbox').forEach(cb => {
        cb.checked = false;
        selectedSeibansForGantt.delete(cb.value);
    });
    updateGanttFilterCount();
}

// フィルタ適用
function applyGanttFilter() {
    // 選択された製番を更新
    selectedSeibansForGantt.clear();
    document.querySelectorAll('.gantt-seiban-checkbox:checked').forEach(cb => {
        selectedSeibansForGantt.add(cb.value);
    });
    
    if (selectedSeibansForGantt.size === 0) {
        showToast('少なくとも1つの製番を選択してください', 'warning');
        return;
    }
    
    // フィルタされたデータ
    const filteredData = allGanttData.filter(d => selectedSeibansForGantt.has(d.seiban));
    
    console.log(`🔍 フィルタ適用: ${allGanttData.length}件 → ${filteredData.length}件`);
    
    if (filteredData.length > 0) {
        renderGanttChart(filteredData);
        showToast(`${filteredData.length}件のデータを表示中`, 'success', 2000);
    } else {
        document.getElementById('ganttChartContainer').innerHTML = 
            '<p style="text-align: center; padding: 50px; color: #6c757d;">選択した製番にデータがありません</p>';
    }
    
    updateGanttFilterCount();
}

// ========================================
// グループ3: チャート描画
// ========================================

// ガントチャートのクリック処理
async function handleGanttLabelClick(item) {
    console.log('🔍 クリックされたアイテム:', item);
    console.log('  - ラベル:', item.label);
    console.log('  - 製番:', item.seiban);
    
    try {
        const response = await fetch('/api/orders');
        const orders = await response.json();
        
        console.log('📦 全注文数:', orders.length);
        
        // ラベルから製番とユニット名を分解
        const [clickedSeiban, ...clickedUnitParts] = item.label.split('_');
        const clickedUnit = clickedUnitParts.join('_'); // アンダースコアが複数ある場合に対応
        
        console.log('  - 分解した製番:', clickedSeiban);
        console.log('  - 分解したユニット:', clickedUnit);
        
        // 該当する注文を検索
        const targetOrder = orders.find(order => {
            const orderSeiban = order.seiban;
            const orderUnit = order.unit || 'ユニット名無し';
            
            console.log(`    比較: ${orderSeiban}_${orderUnit} === ${item.label}`);
            
            return orderSeiban === clickedSeiban && orderUnit === clickedUnit;
        });
        
        if (targetOrder) {
            console.log('✅ 注文が見つかりました:', targetOrder);
            showOrderDetails(targetOrder.id);
        } else {
            console.error('❌ 注文が見つかりません');
            console.log('検索対象:', orders.map(o => ({
                seiban: o.seiban,
                unit: o.unit || 'ユニット名無し',
                label: `${o.seiban}_${o.unit || 'ユニット名無し'}`
            })));
            showToast(`⚠️ ${item.label} の詳細が見つかりません`, 'warning');
        }
    } catch (error) {
        console.error('❌ エラー:', error);
        showToast(`エラー: ${error.message}`, 'error');
    }
}

// 🔥【修正】updateGanttChart関数内の進捗更新を追加
async function updateGanttChart(orders) {
    // ローディング表示を開始
    showGanttLoading(true);
    updateGanttLoadingProgress(0, 'データを準備中...', '注文情報を取得しています...');
    
    // 製番/ユニットごとの納期情報を集計
    const ganttData = [];
    
    console.log('📊 ガントチャート更新開始:', orders.length, '件');
    
    // 🔥【追加】進捗カウント用変数
    const totalOrders = orders.length;
    let completedOrders = 0;
    
    // 全注文の詳細を並行取得
    const promises = orders.map(async (order) => {
        try {
            // ユニット名を先に確認
            const unitName = order.unit || 'ユニット名無し';
            console.log(`  処理中: ${order.seiban}_${unitName}`);
            
            // 🔥【追加】進捗を更新（データ取得フェーズは0-70%）
            completedOrders++;
            const progress = (completedOrders / totalOrders) * 70;
            updateGanttLoadingProgress(
                progress,
                'データ取得中...',
                `${completedOrders} / ${totalOrders} 件処理済み`
            );
            
            const res = await fetch(`/api/order/${order.id}`);
            const data = await res.json();
            
            console.log(`    詳細取得: ${data.details.length}件`);
            
            // 納期を抽出してパース
            const dates = data.details
                .map(d => {
                    // 納期の前処理
                    const dateStr = (d.delivery_date || '').trim();
                    if (!dateStr || dateStr === '-') {
                        return null;
                    }
                    return dateStr;
                })
                .filter(d => d !== null)
                .map(d => parseDeliveryDate(d))
                .filter(d => d && !isNaN(d));
            
            console.log(`    ${order.seiban}_${unitName}: 有効な納期${dates.length}件`);
            
            if (dates.length > 0) {
                const minDate = new Date(Math.min(...dates));
                const maxDate = new Date(Math.max(...dates));
                
                console.log(`    ✅ 追加: ${minDate.toLocaleDateString()} ～ ${maxDate.toLocaleDateString()}`);
                
                return {
                    seiban: order.seiban,
                    label: `${order.seiban}_${unitName}`,
                    start: minDate,
                    end: maxDate,
                    status: order.status,
                    progress: order.detail_count > 0 ? (order.received_count / order.detail_count) * 100 : 0
                };
            } else {
                console.warn(`    ⚠️ スキップ（納期なし）: ${order.seiban}_${unitName}`);
            }
        } catch (error) {
            console.error(`エラー (${order.seiban}):`, error);
        }
        return null;
    });
    
    // 🔥【追加】データ集計フェーズの進捗表示（70-85%）
    updateGanttLoadingProgress(70, 'データを集計中...', 'すべてのデータを結合しています...');
    
    // 全データ取得を待機
    const results = await Promise.all(promises);
    const validData = results.filter(d => d !== null);
    
    // 🔥【追加】データ整理フェーズの進捗表示（85%）
    updateGanttLoadingProgress(85, 'データを整理中...', `${validData.length}件のデータを整理しています...`);
    
    console.log('✅ ガントチャートデータ:', validData.length, '件');
    console.table(validData.map(d => ({
        ユニット: d.label,
        最早納期: d.start.toLocaleDateString(),
        最遅納期: d.end.toLocaleDateString()
    })));
    
    if (validData.length > 0) {
        allGanttData = validData;
        
        // 🔥【追加】製番フィルタUIを初期化（90%）
        updateGanttLoadingProgress(90, 'フィルタを準備中...', 'フィルタUIを初期化しています...');
        initializeGanttFilter(validData);
        
        // 初回は全て表示
        if (selectedSeibansForGantt.size === 0) {
            validData.forEach(d => selectedSeibansForGantt.add(d.seiban));
        }
        
        // 🔥【追加】チャート描画フェーズ（95%）
        updateGanttLoadingProgress(95, 'チャートを描画中...', 'ガントチャートを生成しています...');
        renderGanttChart(validData);
        
        // 🔥【追加】完了（100%）
        updateGanttLoadingProgress(100, '完了！', `${validData.length}件のデータを表示しました`);
        
        // 🔥【追加】0.5秒後にローディング画面を非表示
        setTimeout(() => {
            showGanttLoading(false);
        }, 500);
    } else {
        console.warn('⚠️ 表示可能な納期データがありません');
        document.getElementById('ganttChartContainer').innerHTML = 
            '<p style="text-align: center; padding: 50px; color: #6c757d;">納期データがありません</p>';
        // 🔥【追加】データがない場合もローディングを非表示
        showGanttLoading(false);
    }
}

// 🔥【修正なし】チャート描画関数（関数名のみ表記）
// 注意: この関数は長いため、元のindex.htmlから完全にコピーしてください
function renderGanttChart(data) {
    const ctx = document.getElementById('ganttChart');
    if (!ctx) return;
    
    // 既存チャートを破棄
    if (ganttChartInstance) {
        ganttChartInstance.destroy();
    }
    
    // 今日の日付を取得
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    console.log('📅 今日の日付:', today.toLocaleDateString('ja-JP'), '(timestamp:', today.getTime(), ')');
    
    const threeMonthsAgo = new Date(today);
    threeMonthsAgo.setMonth(today.getMonth() - 3);
    const threeMonthsLater = new Date(today);
    threeMonthsLater.setMonth(today.getMonth() + 3);
    
    console.log('📅 表示範囲:', threeMonthsAgo.toLocaleDateString('ja-JP'), '～', threeMonthsLater.toLocaleDateString('ja-JP'));
    
    // 範囲外のデータをフィルタ
    const filteredData = data.filter(item => {
        return item.start <= threeMonthsLater && item.end >= threeMonthsAgo;
    });
    
    console.log(`📅 フィルタ: ${data.length}件 → ${filteredData.length}件`);
    
    // 1日のみの納期をデバッグ
    const oneDayItems = filteredData.filter(item => {
        const diff = (item.end - item.start) / (1000 * 60 * 60 * 24);
        return diff === 0;
    });
    console.log(`📅 1日のみの納期: ${oneDayItems.length}件`);
    if (oneDayItems.length > 0) {
        console.table(oneDayItems.map(d => ({
            label: d.label,
            date: d.start.toLocaleDateString('ja-JP')
        })));
    }
    
    if (filteredData.length === 0) {
        document.getElementById('ganttChartContainer').innerHTML = 
            '<p style="text-align: center; padding: 50px; color: #6c757d;">今後3ヶ月以内の納期データがありません</p>';
        return;
    }
    
    // 納期でソート
    filteredData.sort((a, b) => a.start - b.start);
    
    // 製番ごとに色を生成する関数
    function getSeibanColor(seiban, alpha = 0.8) {
        let hash = 0;
        for (let i = 0; i < seiban.length; i++) {
            hash = seiban.charCodeAt(i) + ((hash << 5) - hash);
        }
        const hue = Math.abs(hash % 360);
        return `hsla(${hue}, 70%, 55%, ${alpha})`;
    }
    
    // 製番の変わり目を検出
    const seibanBoundaries = [];
    let currentSeiban = null;
    filteredData.forEach((d, index) => {
        if (d.seiban !== currentSeiban) {
            if (index > 0) {
                seibanBoundaries.push(index - 0.5);
            }
            currentSeiban = d.seiban;
        }
    });
    
    // 製番ごとに色を割り当て
    const colors = filteredData.map(d => {
        const baseColor = getSeibanColor(d.seiban, 0.8);
        if (d.status === '納品完了') {
            return getSeibanColor(d.seiban, 0.5);
        } else if (d.status === '納品中') {
            return getSeibanColor(d.seiban, 0.7);
        }
        return baseColor;
    });
    
    const borderColors = filteredData.map(d => getSeibanColor(d.seiban, 1.0));
    
    ganttChartInstance = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: filteredData.map(d => d.label),
            datasets: [{
                label: '納期範囲',
                data: filteredData.map((d, index) => {
                    const start = d.start.getTime();
                    const end = d.end.getTime();
                    const oneDayMs = 24 * 60 * 60 * 1000;
                    const actualEnd = (end === start) ? (start + oneDayMs) : end;
                    
                    return {
                        x: [start, actualEnd],
                        y: index
                    };
                }),
                backgroundColor: colors,
                borderColor: borderColors,
                borderWidth: 2,
                barThickness: 30,
                borderRadius: 5
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                title: {
                    display: true,
                    text: `納期ガントチャート（${filteredData.length}件）`,
                    font: { size: 16 }
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            const item = filteredData[context.dataIndex];
                            const start = item.start.toLocaleDateString('ja-JP');
                            const end = item.end.toLocaleDateString('ja-JP');
                            const days = Math.max(1, Math.ceil((item.end - item.start) / (1000 * 60 * 60 * 24)) + 1);
                            return [
                                `📦 ${item.label}`,
                                `📅 ${start} ～ ${end}`,
                                `⏱️ ${days}日間`,
                                `📊 進捗: ${item.progress.toFixed(1)}%`,
                                `📖 ${item.status}`
                            ];
                        }
                    }
                },
                annotation: {
                    annotations: {
                        todayLine: {
                            type: 'line',
                            xMin: today.getTime(),
                            xMax: today.getTime(),
                            borderColor: 'red',
                            borderWidth: 2,
                            borderDash: [5, 5],
                            label: {
                                display: true,
                                content: '今日',
                                position: 'start',
                                backgroundColor: 'rgba(255, 0, 0, 0.8)',
                                color: 'white',
                                font: { size: 12 }
                            }
                        }
                    }
                }
            },
            scales: {
                x: {
                    type: 'time',
                    time: {
                        unit: 'day',
                        displayFormats: { day: 'M/d' },
                        tooltipFormat: 'yyyy/MM/dd'
                    },
                    min: threeMonthsAgo.getTime(),
                    max: threeMonthsLater.getTime(),
                    title: { display: true, text: '納期' },
                    grid: {
                        color: function(context) {
                            const date = new Date(context.tick.value);
                            if (date.getDay() === 0 || date.getDay() === 6) {
                                return 'rgba(255, 0, 0, 0.1)';
                            }
                            return 'rgba(0, 0, 0, 0.1)';
                        },
                        lineWidth: function(context) {
                            const date = new Date(context.tick.value);
                            if (date.getDate() === 1) {
                                return 2;
                            }
                            return 1;
                        },
                        drawOnChartArea: true
                    }
                },
                y: {
                    title: { display: true, text: 'ユニット（クリックで詳細）' },
                    ticks: {
                        autoSkip: false,
                        font: { size: 11 }
                    }
                }
            },

            onClick: (event, activeElements, chart) => {
                console.log('🖱️ チャートクリックイベント発火');
                console.log('  - activeElements:', activeElements);
                
                // バー要素のクリック検出
                if (activeElements && activeElements.length > 0) {
                    const dataIndex = activeElements[0].index;
                    console.log('✅ バーがクリックされました (index:', dataIndex, ')');
                    
                    if (dataIndex >= 0 && dataIndex < filteredData.length) {
                        const clickedItem = filteredData[dataIndex];
                        console.log('📌 クリックされたユニット:', clickedItem.label);
                        handleGanttLabelClick(clickedItem);
                        return;
                    }
                }
                
                // ラベル領域のクリック検出
                try {
                    const rect = chart.canvas.getBoundingClientRect();
                    const clientX = event.native ? event.native.clientX : event.x;
                    const clientY = event.native ? event.native.clientY : event.y;
                    const canvasX = clientX - rect.left;
                    const canvasY = clientY - rect.top;
                    
                    const yAxis = chart.scales.y;
                    const labelWidth = yAxis.width;
                    
                    console.log('🔍 座標:', {canvasX, labelWidth});
                    
                    if (canvasX >= 0 && canvasX <= labelWidth) {
                        console.log('✅ ラベル領域内のクリック');
                        const value = yAxis.getValueForPixel(canvasY);
                        
                        if (value !== null && value >= 0 && value < filteredData.length) {
                            const dataIndex = Math.floor(value);
                            const clickedItem = filteredData[dataIndex];
                            console.log('📌 クリックされたユニット:', clickedItem.label);
                            handleGanttLabelClick(clickedItem);
                            return;
                        }
                    }
                } catch (error) {
                    console.error('❌ エラー:', error);
                }
                
                console.log('❌ クリック可能領域外');
            }
        },
        plugins: [
            {
                id: 'seibanDividers',
                afterDraw: (chart) => {
                    const ctx = chart.ctx;
                    const yAxis = chart.scales.y;
                    const xAxis = chart.scales.x;
                    
                    ctx.save();
                    ctx.strokeStyle = 'rgba(200, 0, 0, 0.4)';
                    ctx.lineWidth = 2;
                    ctx.setLineDash([8, 4]);
                    
                    seibanBoundaries.forEach(boundary => {
                        const y = yAxis.getPixelForValue(boundary);
                        ctx.beginPath();
                        ctx.moveTo(xAxis.left, y);
                        ctx.lineTo(xAxis.right, y);
                        ctx.stroke();
                    });
                    
                    ctx.restore();
                }
            },
            {
                id: 'yAxisHover',
                afterEvent: (chart, args) => {
                    const event = args.event;
                    const yAxis = chart.scales.y;
                    const labelWidth = yAxis.width;
                    const labelLeft = yAxis.left - labelWidth;
                    
                    if (event.x >= labelLeft && event.x <= yAxis.left &&
                        event.y >= yAxis.top && event.y <= yAxis.bottom) {
                        chart.canvas.style.cursor = 'pointer';
                    } else {
                        chart.canvas.style.cursor = 'default';
                    }
                }
            }
        ]
    });

    // チャートコンテナの高さを動的調整
    const container = document.getElementById('ganttChartContainer');
    container.style.height = Math.max(400, filteredData.length * 50) + 'px';

    console.log('✅ ガントチャート描画完了:', filteredData.length, '件');
}

// ========================================
// Chart.jsプラグイン登録（初回のみ実行）
// ========================================
if (typeof Chart !== 'undefined' && Chart.registry) {
    if (typeof chartjsPluginAnnotation !== 'undefined') {
        Chart.register(chartjsPluginAnnotation);
        console.log('✅ Annotation plugin registered');
    } else {
        console.warn('⚠️ Annotation plugin not loaded');
    }
}