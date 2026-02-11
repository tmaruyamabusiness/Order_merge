// ========================================
// QRコード・バーコードスキャナーモジュール
// Html5Qrcode版（バーコード読み取り強化）
// ========================================

// ========================================
// グローバル変数
// ========================================
let html5QrCode = null;
let lastScannedCode = null;
let scanCooldown = false;
let scanHistory = new Set();
let isScannerPaused = false;
let isModalOpen = false;

// 設定値
const CONFIG = {
    SCAN_COOLDOWN_MS: 2000,       // 連続読み取り防止時間（ミリ秒）
    AUTO_RESUME_MS: 1500,         // 受入後の自動再開待ち時間（ミリ秒）
    BEEP_FREQUENCY_SUCCESS: 1200, // 成功時のビープ音周波数
    BEEP_FREQUENCY_ERROR: 300,    // エラー時のビープ音周波数
    BEEP_GAIN_SUCCESS: 0.5,       // 成功時の音量
    BEEP_GAIN_ERROR: 0.3,         // エラー時の音量
    BEEP_DURATION_SUCCESS: 0.2,   // 成功時の音の長さ
    BEEP_DURATION_ERROR: 0.15,    // エラー時の音の長さ
    VIBRATION_PATTERN: [100, 50, 100]  // バイブレーションパターン
};

// スキャン処理済み発注番号（重複受入防止）
let processedOrderNumbers = new Set();

// スキャンモード: 'wide'(デフォルト), 'narrow'(ピンポイント), 'barcode'(横長)
let currentScanMode = 'wide';

// ========================================
// QRスキャナーを開始
// ========================================
function startQRScanner() {
    if (isModalOpen) {
        console.log('Scanner modal already open');
        return;
    }
    isModalOpen = true;

    // 前回の読み取り結果をクリア
    const qrResult = document.getElementById('qrResult');
    if (qrResult) {
        qrResult.innerHTML = '';
    }

    // スキャン履歴をクリア
    scanHistory.clear();
    processedOrderNumbers.clear();
    lastScannedCode = null;
    isScannerPaused = false;

    // モーダルを表示
    document.getElementById('qrScannerModal').classList.add('show');

    // スキャンログをクリア
    const scanLog = document.getElementById('scanLog');
    if (scanLog) scanLog.innerHTML = '';

    // スキャナー初期化
    setTimeout(() => {
        initializeScanner();
    }, 300);
}

// ========================================
// スキャナーの初期化（Html5Qrcode版）
// ========================================
function initializeScanner() {
    console.log('スキャナー初期化開始');

    const readerElement = document.getElementById("qr-reader");
    if (!readerElement) {
        console.error('qr-readerエレメントが見つかりません');
        alert('QRリーダーの初期化に失敗しました');
        return;
    }

    // 既存のスキャナーがあれば停止
    if (html5QrCode) {
        try {
            if (html5QrCode.isScanning) {
                html5QrCode.stop().then(() => {
                    html5QrCode.clear();
                    createNewScanner();
                }).catch((err) => {
                    console.error("既存スキャナー停止エラー:", err);
                    createNewScanner();
                });
            } else {
                html5QrCode.clear();
                createNewScanner();
            }
        } catch (e) {
            console.error('スキャナークリアエラー:', e);
            createNewScanner();
        }
    } else {
        createNewScanner();
    }
}

// ========================================
// 新しいスキャナーを作成
// ========================================
// ========================================
// スキャンモード切替（カメラ再起動）
// ========================================
function setScanMode(mode) {
    currentScanMode = mode;

    // ボタンのアクティブ状態を更新
    ['Wide', 'Narrow', 'Barcode'].forEach(m => {
        const btn = document.getElementById('scanMode' + m);
        if (btn) {
            if (m.toLowerCase() === mode) {
                btn.style.background = '#0d6efd';
                btn.style.color = '#fff';
            } else {
                btn.style.background = '';
                btn.style.color = '';
            }
        }
    });

    // スキャナーを再起動して新しい枠サイズを適用
    if (html5QrCode && html5QrCode.isScanning) {
        html5QrCode.stop().then(() => {
            html5QrCode.clear();
            createNewScanner();
        }).catch(() => {
            createNewScanner();
        });
    }
}

function createNewScanner() {
    html5QrCode = new Html5Qrcode("qr-reader");

    // バーコード・QRコード対応の設定（モードに応じて枠サイズ変更）
    const config = {
        fps: 15,
        qrbox: function(viewfinderWidth, viewfinderHeight) {
            let w, h;
            if (currentScanMode === 'narrow') {
                // ピンポイント: 小さい正方形（近接QR用）
                let size = Math.floor(Math.min(viewfinderWidth, viewfinderHeight) * 0.35);
                size = Math.max(100, Math.min(180, size));
                w = size;
                h = size;
            } else if (currentScanMode === 'barcode') {
                // バーコード: 横長の矩形
                w = Math.floor(viewfinderWidth * 0.8);
                w = Math.max(250, Math.min(400, w));
                h = Math.floor(w * 0.3);
            } else {
                // 広域: デフォルト
                let size = Math.floor(Math.min(viewfinderWidth, viewfinderHeight) * 0.7);
                size = Math.max(200, Math.min(350, size));
                w = size;
                h = size;
            }
            return { width: w, height: h };
        },
        aspectRatio: 1.0,
        videoConstraints: {
            width: { ideal: 1920, min: 1280 },
            height: { ideal: 1080, min: 720 },
            focusMode: "continuous",
            facingMode: "environment"
        },
        // 対応フォーマット（1次元バーコード・2次元コード）
        formatsToSupport: [
            // 2次元コード
            Html5QrcodeSupportedFormats.QR_CODE,
            Html5QrcodeSupportedFormats.DATA_MATRIX,
            Html5QrcodeSupportedFormats.AZTEC,
            Html5QrcodeSupportedFormats.PDF_417,
            // 1次元バーコード
            Html5QrcodeSupportedFormats.CODE_128,
            Html5QrcodeSupportedFormats.CODE_39,
            Html5QrcodeSupportedFormats.CODE_93,
            Html5QrcodeSupportedFormats.EAN_13,
            Html5QrcodeSupportedFormats.EAN_8,
            Html5QrcodeSupportedFormats.UPC_A,
            Html5QrcodeSupportedFormats.UPC_E,
            Html5QrcodeSupportedFormats.CODABAR,
            Html5QrcodeSupportedFormats.ITF
        ]
    };

    // カメラの取得
    Html5Qrcode.getCameras().then(devices => {
        console.log(`検出されたカメラ数: ${devices.length}`);

        if (devices && devices.length) {
            // 背面カメラを優先
            let cameraId = devices[0].id;
            for (let device of devices) {
                if (device.label && (
                    device.label.toLowerCase().includes('back') ||
                    device.label.toLowerCase().includes('rear') ||
                    device.label.toLowerCase().includes('environment') ||
                    device.label.includes('後'))) {
                    cameraId = device.id;
                    console.log(`背面カメラを選択: ${device.label}`);
                    break;
                }
            }

            // カメラを起動
            html5QrCode.start(
                cameraId,
                config,
                (decodedText, decodedResult) => {
                    console.log(`読み取り成功 - タイプ: ${decodedResult.result.format?.formatName || '不明'}, 値: ${decodedText}`);
                    onCodeScannedWithCooldown(decodedText);
                },
                (errorMessage) => {
                    // スキャン中のエラーは無視（NotFoundException等）
                }
            ).then(() => {
                console.log('カメラ起動成功');
                showScannerStatus('カメラ起動中 - QR/バーコードを枠内に収めてください', 'success');
            }).catch((err) => {
                console.error(`カメラ起動エラー: ${err}`);
                handleCameraError(err);
            });
        } else {
            console.error('利用可能なカメラが見つかりません');
            alert('利用可能なカメラが見つかりません');
        }
    }).catch((err) => {
        console.error(`カメラ一覧取得エラー: ${err}`);
        alert('カメラにアクセスできません。ブラウザの権限設定を確認してください。');
    });
}

// ========================================
// カメラエラーのハンドリング
// ========================================
function handleCameraError(err) {
    let errorMsg = 'カメラの起動に失敗しました。';
    if (err.toString().includes('NotAllowedError')) {
        errorMsg += 'カメラの使用が許可されていません。';
    } else if (err.toString().includes('NotFoundError')) {
        errorMsg += 'カメラが見つかりません。';
    } else if (err.toString().includes('NotReadableError')) {
        errorMsg += 'カメラが他のアプリで使用中の可能性があります。';
    }

    alert(errorMsg);

    // フォールバック
    const fallbackConfig = {
        fps: 10,
        qrbox: { width: 250, height: 250 }
    };

    html5QrCode.start(
        { facingMode: "environment" },
        fallbackConfig,
        (decodedText) => {
            onCodeScannedWithCooldown(decodedText);
        },
        () => {}
    ).then(() => {
        console.log('カメラ起動成功（フォールバック）');
    }).catch((fallbackErr) => {
        console.error('フォールバックも失敗:', fallbackErr);
    });
}

// ========================================
// スキャナーステータス表示
// ========================================
function showScannerStatus(message, type) {
    const qrResult = document.getElementById('qrResult');
    if (qrResult) {
        const color = type === 'success' ? '#28a745' : type === 'error' ? '#dc3545' : '#17a2b8';
        qrResult.innerHTML = `<div style="color: ${color}; font-size: 0.9em;">${message}</div>`;
    }
}

// ========================================
// 連続読み取り防止機能付きコード読み取り
// ========================================
function onCodeScannedWithCooldown(scannedText) {
    if (isScannerPaused) {
        console.log('スキャナー一時停止中');
        return;
    }

    scannedText = scannedText.trim();

    // 連続読み取り防止
    if (scanCooldown && lastScannedCode === scannedText) {
        console.log(`同じコードの連続読み取りをブロック: ${scannedText}`);
        return;
    }

    // 既に読み取り済みかチェック
    if (scanHistory.has(scannedText)) {
        console.log(`既に読み取り済み: ${scannedText}`);
        playBeep(false);
        showScannerStatus(`既に読み取り済み: ${scannedText}`, 'warning');
        return;
    }

    console.log(`コード読み取り成功: ${scannedText}`);

    scanHistory.add(scannedText);
    lastScannedCode = scannedText;

    // スキャナーを一時停止
    isScannerPaused = true;

    // 連続読み取り防止
    scanCooldown = true;
    setTimeout(() => {
        scanCooldown = false;
        lastScannedCode = null;
    }, CONFIG.SCAN_COOLDOWN_MS);

    // 成功音とバイブレーション
    playBeep(true);
    if (navigator.vibrate) {
        navigator.vibrate(CONFIG.VIBRATION_PATTERN);
    }

    // 読み取り結果表示
    showScannerStatus(`✅ 読み取り成功: ${scannedText}`, 'success');

    // コード処理
    processScannedCode(scannedText);
}

// ========================================
// ビープ音再生
// ========================================
function playBeep(success = true) {
    try {
        const audioContext = new (window.AudioContext || window.webkitAudioContext)();
        const oscillator = audioContext.createOscillator();
        const gainNode = audioContext.createGain();

        oscillator.connect(gainNode);
        gainNode.connect(audioContext.destination);

        if (success) {
            oscillator.frequency.value = CONFIG.BEEP_FREQUENCY_SUCCESS;
            gainNode.gain.value = CONFIG.BEEP_GAIN_SUCCESS;
            oscillator.type = 'sine';

            const now = audioContext.currentTime;
            gainNode.gain.setValueAtTime(CONFIG.BEEP_GAIN_SUCCESS, now);
            gainNode.gain.exponentialRampToValueAtTime(0.01, now + CONFIG.BEEP_DURATION_SUCCESS);

            oscillator.start(now);
            oscillator.stop(now + CONFIG.BEEP_DURATION_SUCCESS);
        } else {
            oscillator.frequency.value = CONFIG.BEEP_FREQUENCY_ERROR;
            gainNode.gain.value = CONFIG.BEEP_GAIN_ERROR;
            oscillator.type = 'square';

            oscillator.start(audioContext.currentTime);
            oscillator.stop(audioContext.currentTime + CONFIG.BEEP_DURATION_ERROR);
        }
    } catch (e) {
        console.log('音声再生エラー:', e);
    }
}

// ========================================
// スキャンしたコードを処理（連続スキャン対応）
// ========================================
function processScannedCode(data) {
    const purchaseOrderInput = document.getElementById('purchaseOrderInput');

    // データの前処理
    data = data.trim();
    let orderNumber = null;

    // 🔥 発注番号バーコードパターン: 8桁の数字 + アルファベット1文字 (例: 00088333P)
    const purchaseOrderBarcodePattern = /^(\d{8})[A-Za-z]$/;
    const barcodeMatch = data.match(purchaseOrderBarcodePattern);

    if (barcodeMatch) {
        const numericPart = barcodeMatch[1];
        orderNumber = String(parseInt(numericPart, 10));
        console.log(`発注番号バーコード検出: ${data} → ${orderNumber}`);
    }

    // 🔥 誤読み取り防止: 数字+アルファベットの混在パターンを検出
    if (!orderNumber) {
        const hasEightDigits = /\d{8}/.test(data);
        if (!hasEightDigits) {
            const invalidBarcodePattern = /^\d*[A-Za-z]+\d*[A-Za-z]*$/;
            if (invalidBarcodePattern.test(data) && data.length >= 8) {
                console.log(`無効なバーコード形式を検出: ${data}`);
                showScannerStatus(`⚠️ 読み取りエラー: ${data}`, 'error');
                playBeep(false);
                isScannerPaused = false;
                scanHistory.delete(data);
                return;
            }
        }
    }

    // QRコードのフォーマットをチェックして発注番号を抽出
    if (!orderNumber) {
        if (data.toUpperCase().startsWith('PO:')) {
            orderNumber = data.substring(3);
        } else if (data.toUpperCase().startsWith('ORDER:')) {
            // 注文IDのQRコード - 別処理
            const orderId = data.substring(6);
            stopQRScanner();
            if (typeof showOrderDetails === 'function') {
                showOrderDetails(parseInt(orderId));
            }
            return;
        } else if (/^\d{4,6}$/.test(data)) {
            orderNumber = data;
        } else {
            // テキスト内から8桁以上の連続数字を抽出
            const allDigitMatches = data.match(/\d{8,}/g);
            if (allDigitMatches) {
                const longestMatch = allDigitMatches[allDigitMatches.length - 1];
                const numericPart = longestMatch.slice(-8);
                orderNumber = String(parseInt(numericPart, 10));
                console.log(`QRテキストから発注番号抽出: ${data} → ${orderNumber}`);
            } else if (data.toUpperCase().startsWith('MHT')) {
                showScannerToast(`⚠️ 製番のみ: ${data}`, 'warning');
                resumeScanning();
                return;
            } else {
                showScannerToast(`❓ 不明: ${data}`, 'warning');
                resumeScanning();
                return;
            }
        }
    }

    if (!orderNumber) {
        resumeScanning();
        return;
    }

    // 既に処理済みの発注番号はスキップ
    if (processedOrderNumbers.has(orderNumber)) {
        console.log(`既に処理済みの発注番号: ${orderNumber}`);
        showScannerToast(`⏭️ 処理済み: ${orderNumber}`, 'info');
        resumeScanning();
        return;
    }

    if (purchaseOrderInput) purchaseOrderInput.value = orderNumber;

    // 連続スキャンモード: スキャナーを止めずに受入処理
    autoReceiveByOrderNumber(orderNumber);
}

// ========================================
// スキャナーを止めずに自動受入処理
// ========================================
async function autoReceiveByOrderNumber(orderNumber) {
    showScannerToast(`🔍 検索中: ${orderNumber}`, 'info');

    try {
        const response = await fetch(`/api/search-by-purchase-order/${orderNumber}`);
        const data = await response.json();

        if (!data.found || data.details.length === 0) {
            showScannerToast(`❌ 未登録: ${orderNumber}`, 'error');
            playBeep(false);
            resumeScanning();
            return;
        }

        const mergedDetails = data.details.filter(d => d.source === 'merged');

        if (mergedDetails.length === 0) {
            showScannerToast(`⚠️ 未マージ: ${orderNumber}`, 'warning');
            playBeep(false);
            resumeScanning();
            return;
        }

        // 未受入のものだけ抽出
        const unreceived = mergedDetails.filter(d => !d.is_received);

        if (unreceived.length === 0) {
            // 全て受入済み
            showScannerToast(`✅ 受入済み: ${orderNumber}（${mergedDetails[0].item_name || ''})`, 'success');
            processedOrderNumbers.add(orderNumber);
            addScanLogEntry(orderNumber, mergedDetails[0], 'already');
            resumeScanning();
            return;
        }

        if (unreceived.length === 1) {
            // 単一アイテム → 確認ポップアップを表示
            const detail = unreceived[0];
            showReceiveConfirmPopup(orderNumber, detail);
            return;
        }

        // 複数アイテム → 確認ポップアップ（スキャナー一時停止のまま）
        showBarcodeReceivePopup(orderNumber);

    } catch (error) {
        console.error('自動受入エラー:', error);
        showScannerToast(`❌ エラー: ${orderNumber}`, 'error');
        playBeep(false);
        resumeScanning();
    }
}

// ========================================
// スキャナー上にトースト通知を表示
// ========================================
function showScannerToast(message, type) {
    const qrResult = document.getElementById('qrResult');
    if (!qrResult) return;

    const colors = {
        success: { bg: '#d4edda', border: '#28a745', text: '#155724' },
        error:   { bg: '#f8d7da', border: '#dc3545', text: '#721c24' },
        warning: { bg: '#fff3cd', border: '#ffc107', text: '#856404' },
        info:    { bg: '#d1ecf1', border: '#17a2b8', text: '#0c5460' }
    };
    const c = colors[type] || colors.info;

    qrResult.innerHTML = `<div style="
        background:${c.bg}; border:2px solid ${c.border}; color:${c.text};
        border-radius:8px; padding:8px 14px; font-size:1em; font-weight:bold;
        animation: toastFadeIn 0.2s ease-out;
    ">${message}</div>`;
}

// ========================================
// 受入確認ポップアップ（単一アイテム用）
// ========================================
function showReceiveConfirmPopup(orderNumber, detail) {
    const modalBody = document.getElementById('barcodeReceiveModalBody');

    const html = `
        <div style="text-align: center; margin-bottom: 15px;">
            <h3 style="margin: 0; color: #333;">受入確認</h3>
            <p style="color: #6c757d; margin: 5px 0; font-size:0.9em;">以下の内容でよろしいですか？</p>
        </div>
        <div style="background: #f8f9fa; border: 2px solid #007bff; border-radius: 10px; padding: 15px; margin-bottom: 15px;">
            <table style="width: 100%; font-size: 0.95em; border-collapse: collapse;">
                <tr style="border-bottom: 1px solid #dee2e6;">
                    <td style="font-weight:bold; padding: 8px 0; width: 90px; color: #495057;">発注番号:</td>
                    <td style="padding: 8px 0; font-size: 1.1em;"><strong>${orderNumber}</strong></td>
                </tr>
                <tr style="border-bottom: 1px solid #dee2e6;">
                    <td style="font-weight:bold; padding: 8px 0; color: #495057;">製番:</td>
                    <td style="padding: 8px 0;">${detail.seiban || '-'}</td>
                </tr>
                <tr style="border-bottom: 1px solid #dee2e6;">
                    <td style="font-weight:bold; padding: 8px 0; color: #495057;">ユニット名:</td>
                    <td style="padding: 8px 0;">${detail.unit || '-'}</td>
                </tr>
                <tr style="border-bottom: 1px solid #dee2e6;">
                    <td style="font-weight:bold; padding: 8px 0; color: #495057;">品名:</td>
                    <td style="padding: 8px 0;">${detail.item_name || '-'}</td>
                </tr>
                <tr style="border-bottom: 1px solid #dee2e6;">
                    <td style="font-weight:bold; padding: 8px 0; color: #495057;">仕様1:</td>
                    <td style="padding: 8px 0;">${detail.spec1 || '-'}</td>
                </tr>
                <tr>
                    <td style="font-weight:bold; padding: 8px 0; color: #495057;">個数:</td>
                    <td style="padding: 8px 0; font-size: 1.1em;"><strong>${detail.quantity || '-'} ${detail.unit_measure || ''}</strong></td>
                </tr>
            </table>
        </div>
        <div style="display: flex; gap: 10px;">
            <button class="btn btn-success" id="confirmReceiveBtn" onclick="executeConfirmedReceive(${detail.id}, '${orderNumber}')"
                    style="flex: 1; padding: 12px; font-size: 1.1em; font-weight: bold;">
                ✅ 受入する
            </button>
            <button class="btn btn-secondary" onclick="cancelReceiveAndResume()"
                    style="flex: 1; padding: 12px; font-size: 1em;">
                キャンセル
            </button>
        </div>
    `;

    modalBody.innerHTML = html;
    document.getElementById('barcodeReceiveModal').classList.add('show');
}

// ========================================
// 確認済み受入処理を実行
// ========================================
async function executeConfirmedReceive(detailId, orderNumber) {
    const btn = document.getElementById('confirmReceiveBtn');
    if (btn) { btn.disabled = true; btn.textContent = '処理中...'; }

    try {
        const response = await fetch(`/api/detail/${detailId}/toggle-receive`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ is_received: true })
        });
        const data = await response.json();

        if (data.success) {
            playBeep(true);
            if (navigator.vibrate) navigator.vibrate(CONFIG.VIBRATION_PATTERN);
            showScannerToast(`✅ 受入完了: ${orderNumber}`, 'success');
            processedOrderNumbers.add(orderNumber);

            // ポップアップを更新（カウントダウン付き）
            const modalBody = document.getElementById('barcodeReceiveModalBody');
            let countdown = 3;
            modalBody.innerHTML = `
                <div style="text-align: center; padding: 30px;">
                    <div style="font-size: 3em; margin-bottom: 15px;">✅</div>
                    <h3 style="color: #28a745; margin-bottom: 10px;">受入完了</h3>
                    <p style="color: #6c757d; font-size: 0.95em;">発注番号: ${orderNumber}</p>
                    <p id="countdownText" style="color: #6c757d; font-size: 0.9em; margin-top: 15px;">${countdown}秒後に自動でスキャン画面に戻ります</p>
                    <button class="btn" onclick="cancelAutoResume()" style="background: #ffc107; color: #000; padding: 10px 30px; font-size: 1em; margin-top: 10px; font-weight: bold;">
                        🟨 キャンセル
                    </button>
                </div>
            `;

            // カウントダウンタイマー
            window.autoResumeTimer = setInterval(() => {
                countdown--;
                const countdownEl = document.getElementById('countdownText');
                if (countdownEl) {
                    countdownEl.textContent = `${countdown}秒後に自動でスキャン画面に戻ります`;
                }
                if (countdown <= 0) {
                    clearInterval(window.autoResumeTimer);
                    closeBarcodeReceiveAndResume();
                }
            }, 1000);
        } else {
            if (btn) { btn.disabled = false; btn.textContent = '✅ 受入する'; }
            showScannerToast(`❌ 受入失敗: ${orderNumber}`, 'error');
            playBeep(false);
        }
    } catch (error) {
        console.error('受入エラー:', error);
        if (btn) { btn.disabled = false; btn.textContent = '✅ 受入する'; }
        showScannerToast(`❌ エラー: ${orderNumber}`, 'error');
        playBeep(false);
    }
}

// ========================================
// 受入キャンセルしてスキャン再開
// ========================================
function cancelReceiveAndResume() {
    document.getElementById('barcodeReceiveModal').classList.remove('show');
    showScannerToast('キャンセルしました', 'info');
    resumeScanning();
}

// ========================================
// スキャンログにエントリ追加
// ========================================
function addScanLogEntry(orderNumber, detail, status) {
    const scanLog = document.getElementById('scanLog');
    if (!scanLog) return;

    const icon = status === 'received' ? '✅' : status === 'already' ? '🔄' : '❌';
    const statusText = status === 'received' ? '受入完了' : status === 'already' ? '受入済み' : 'エラー';

    const entry = document.createElement('div');
    entry.style.cssText = 'display:flex; justify-content:space-between; align-items:center; padding:4px 8px; font-size:0.82em; border-bottom:1px solid #eee;';
    entry.innerHTML = `
        <span>${icon} <strong>${orderNumber}</strong></span>
        <span style="color:#6c757d;">${detail.item_name || ''}</span>
        <span style="color:#6c757d; font-size:0.9em;">${statusText}</span>
    `;

    // 最新を上に追加
    scanLog.insertBefore(entry, scanLog.firstChild);
}

// ========================================
// スキャン再開（連続スキャン用）
// ========================================
function resumeScanning() {
    setTimeout(() => {
        isScannerPaused = false;
    }, CONFIG.AUTO_RESUME_MS);
}

// ========================================
// QRスキャナーを停止
// ========================================
function stopQRScanner() {
    console.log('スキャナー停止処理開始');

    isModalOpen = false;
    isScannerPaused = false;
    scanHistory.clear();
    processedOrderNumbers.clear();
    lastScannedCode = null;
    scanCooldown = false;

    // モーダルを閉じる
    document.getElementById('qrScannerModal').classList.remove('show');
    document.getElementById('barcodeReceiveModal').classList.remove('show');

    // スキャナー停止
    if (html5QrCode) {
        try {
            if (html5QrCode.isScanning) {
                html5QrCode.stop().then(() => {
                    console.log('スキャナー停止成功');
                    html5QrCode.clear();
                    html5QrCode = null;
                }).catch((err) => {
                    console.error(`スキャナー停止エラー: ${err}`);
                    try { html5QrCode.clear(); } catch (e) {}
                    html5QrCode = null;
                });
            } else {
                try { html5QrCode.clear(); } catch (e) {}
                html5QrCode = null;
            }
        } catch (e) {
            console.error('スキャナー停止処理エラー:', e);
            html5QrCode = null;
        }
    }
}

// ========================================
// バーコード受入確認ポップアップ（複数アイテム時のみ使用）
// ========================================
async function showBarcodeReceivePopup(orderNumber) {
    try {
        const response = await fetch(`/api/search-by-purchase-order/${orderNumber}`);
        const data = await response.json();

        const modalBody = document.getElementById('barcodeReceiveModalBody');

        if (!data.found || data.details.length === 0) {
            showScannerToast(`❌ 未登録: ${orderNumber}`, 'error');
            resumeScanning();
            return;
        }

        const mergedDetails = data.details.filter(d => d.source === 'merged');

        if (mergedDetails.length === 0) {
            showScannerToast(`⚠️ 未マージ: ${orderNumber}`, 'warning');
            resumeScanning();
            return;
        }

        // 複数アイテムの受入画面を生成
        const unreceived = mergedDetails.filter(d => !d.is_received);
        let html = `
            <div style="text-align: center; margin-bottom: 15px;">
                <h3 style="margin: 0;">複数アイテム: ${orderNumber}</h3>
                <p style="color: #6c757d; margin: 5px 0; font-size:0.9em;">${unreceived.length}件 未受入</p>
            </div>
        `;

        mergedDetails.forEach((detail) => {
            const isReceived = detail.is_received;
            const bgColor = isReceived ? '#d4edda' : '#fff3cd';
            const borderColor = isReceived ? '#28a745' : '#ffc107';

            html += `
                <div id="barcodeItem_${detail.id}" style="background: ${bgColor}; border: 2px solid ${borderColor}; border-radius: 10px; padding: 12px; margin-bottom: 8px;">
                    ${isReceived ? '<div style="color: #28a745; font-weight: bold; margin-bottom: 5px; font-size:0.9em;">✅ 受入済み</div>' : ''}
                    <table style="width: 100%; font-size: 0.88em;">
                        <tr><td style="font-weight:bold; width:70px;">製番:</td><td>${detail.seiban || '-'}</td></tr>
                        <tr><td style="font-weight:bold;">品名:</td><td>${detail.item_name || '-'}</td></tr>
                        <tr><td style="font-weight:bold;">個数:</td><td><strong>${detail.quantity || '-'} ${detail.unit_measure || ''}</strong></td></tr>
                    </table>
                    ${!isReceived ? `
                        <button class="btn btn-success" id="receiveBtn_${detail.id}" onclick="executeBarcodeReceive(${detail.id}, '${orderNumber}')"
                                style="width: 100%; margin-top: 8px; padding: 10px; font-size: 1em;">
                            受入する
                        </button>
                    ` : ''}
                </div>
            `;
        });

        // 一括受入ボタン（未受入が2件以上の場合）
        if (unreceived.length >= 2) {
            html += `
                <button class="btn btn-success" id="barcodeReceiveAllBtn" onclick="executeAllBarcodeReceive('${orderNumber}', [${unreceived.map(d => d.id).join(',')}])"
                        style="width: 100%; margin-top: 5px; padding: 12px; font-size: 1.05em; font-weight:bold;">
                    全て受入する（${unreceived.length}件）
                </button>
            `;
        }

        html += `
            <button class="btn btn-secondary" onclick="closeBarcodeReceiveAndResume()" style="width: 100%; margin-top: 8px;">
                スキャンに戻る
            </button>
        `;

        modalBody.innerHTML = html;
        document.getElementById('barcodeReceiveModal').classList.add('show');

    } catch (error) {
        console.error('検索エラー:', error);
        showScannerToast(`❌ エラー: ${orderNumber}`, 'error');
        resumeScanning();
    }
}

// ========================================
// バーコード受入を実行（複数アイテムモーダル内の個別ボタン用）
// ========================================
async function executeBarcodeReceive(detailId, orderNumber) {
    const btn = document.getElementById(`receiveBtn_${detailId}`);
    if (btn) { btn.disabled = true; btn.textContent = '処理中...'; }

    try {
        const response = await fetch(`/api/detail/${detailId}/toggle-receive`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ is_received: true })
        });
        const data = await response.json();

        if (data.success) {
            playBeep(true);
            if (navigator.vibrate) navigator.vibrate(CONFIG.VIBRATION_PATTERN);

            // アイテム表示をインライン更新
            const itemDiv = document.getElementById(`barcodeItem_${detailId}`);
            if (itemDiv) {
                itemDiv.style.background = '#d4edda';
                itemDiv.style.borderColor = '#28a745';
                if (btn) btn.remove();
                const doneLabel = document.createElement('div');
                doneLabel.style.cssText = 'color:#28a745; font-weight:bold; margin-top:5px; font-size:0.9em;';
                doneLabel.textContent = '✅ 受入完了';
                itemDiv.appendChild(doneLabel);
            }

            processedOrderNumbers.add(orderNumber);
        } else {
            if (btn) { btn.disabled = false; btn.textContent = '受入する'; }
            showScannerToast(`❌ 受入失敗: ${orderNumber}`, 'error');
        }
    } catch (error) {
        console.error('受入エラー:', error);
        if (btn) { btn.disabled = false; btn.textContent = '受入する'; }
    }
}

// ========================================
// 複数アイテム一括受入
// ========================================
async function executeAllBarcodeReceive(orderNumber, detailIds) {
    const allBtn = document.getElementById('barcodeReceiveAllBtn');
    if (allBtn) { allBtn.disabled = true; allBtn.textContent = '処理中...'; }

    let successCount = 0;
    for (const detailId of detailIds) {
        try {
            const response = await fetch(`/api/detail/${detailId}/toggle-receive`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ is_received: true })
            });
            const data = await response.json();
            if (data.success) {
                successCount++;
                // アイテム表示を更新
                const itemDiv = document.getElementById(`barcodeItem_${detailId}`);
                if (itemDiv) {
                    itemDiv.style.background = '#d4edda';
                    itemDiv.style.borderColor = '#28a745';
                    const btn = document.getElementById(`receiveBtn_${detailId}`);
                    if (btn) btn.remove();
                    const doneLabel = document.createElement('div');
                    doneLabel.style.cssText = 'color:#28a745; font-weight:bold; margin-top:5px; font-size:0.9em;';
                    doneLabel.textContent = '✅ 受入完了';
                    itemDiv.appendChild(doneLabel);
                }
            }
        } catch (e) { console.error('一括受入エラー:', e); }
    }

    if (successCount > 0) {
        playBeep(true);
        if (navigator.vibrate) navigator.vibrate(CONFIG.VIBRATION_PATTERN);
        processedOrderNumbers.add(orderNumber);
    }

    // ポップアップを更新（カウントダウン付き）
    const modalBody = document.getElementById('barcodeReceiveModalBody');
    let countdown = 3;
    modalBody.innerHTML = `
        <div style="text-align: center; padding: 30px;">
            <div style="font-size: 3em; margin-bottom: 15px;">✅</div>
            <h3 style="color: #28a745; margin-bottom: 10px;">受入完了</h3>
            <p style="color: #6c757d; font-size: 0.95em;">${successCount}/${detailIds.length}件 受入完了</p>
            <p style="color: #6c757d; font-size: 0.95em;">発注番号: ${orderNumber}</p>
            <p id="countdownText" style="color: #6c757d; font-size: 0.9em; margin-top: 15px;">${countdown}秒後に自動でスキャン画面に戻ります</p>
            <button class="btn" onclick="cancelAutoResume()" style="background: #ffc107; color: #000; padding: 10px 30px; font-size: 1em; margin-top: 10px; font-weight: bold;">
                🟨 キャンセル
            </button>
        </div>
    `;

    // カウントダウンタイマー
    window.autoResumeTimer = setInterval(() => {
        countdown--;
        const countdownEl = document.getElementById('countdownText');
        if (countdownEl) {
            countdownEl.textContent = `${countdown}秒後に自動でスキャン画面に戻ります`;
        }
        if (countdown <= 0) {
            clearInterval(window.autoResumeTimer);
            closeBarcodeReceiveAndResume();
        }
    }, 1000);
}

// ========================================
// 自動復帰をキャンセル
// ========================================
function cancelAutoResume() {
    if (window.autoResumeTimer) {
        clearInterval(window.autoResumeTimer);
        window.autoResumeTimer = null;
    }
    const countdownEl = document.getElementById('countdownText');
    if (countdownEl) {
        countdownEl.textContent = '自動復帰をキャンセルしました';
    }
}

// ========================================
// バーコード受入モーダルを閉じてスキャン再開
// ========================================
function closeBarcodeReceiveAndResume() {
    if (window.autoResumeTimer) {
        clearInterval(window.autoResumeTimer);
        window.autoResumeTimer = null;
    }
    document.getElementById('barcodeReceiveModal').classList.remove('show');
    resumeScanning();
}

// ========================================
// バーコード受入モーダルを閉じる（後方互換）
// ========================================
function closeBarcodeReceiveModal() {
    document.getElementById('barcodeReceiveModal').classList.remove('show');
}
