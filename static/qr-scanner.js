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
    SCAN_COOLDOWN_MS: 3000,       // 連続読み取り防止時間（ミリ秒）
    BEEP_FREQUENCY_SUCCESS: 1200, // 成功時のビープ音周波数
    BEEP_FREQUENCY_ERROR: 300,    // エラー時のビープ音周波数
    BEEP_GAIN_SUCCESS: 0.5,       // 成功時の音量
    BEEP_GAIN_ERROR: 0.3,         // エラー時の音量
    BEEP_DURATION_SUCCESS: 0.2,   // 成功時の音の長さ
    BEEP_DURATION_ERROR: 0.15,    // エラー時の音の長さ
    VIBRATION_PATTERN: [100, 50, 100]  // バイブレーションパターン
};

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
    lastScannedCode = null;
    isScannerPaused = false;

    // モーダルを表示
    document.getElementById('qrScannerModal').classList.add('show');

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
function createNewScanner() {
    html5QrCode = new Html5Qrcode("qr-reader");

    // バーコード・QRコード対応の設定
    const config = {
        fps: 15,
        qrbox: function(viewfinderWidth, viewfinderHeight) {
            let minEdgePercentage = 0.7;
            let minEdgeSize = Math.min(viewfinderWidth, viewfinderHeight);
            let qrboxSize = Math.floor(minEdgeSize * minEdgePercentage);
            qrboxSize = Math.max(200, Math.min(350, qrboxSize));
            return {
                width: qrboxSize,
                height: qrboxSize
            };
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
// スキャンしたコードを処理
// ========================================
function processScannedCode(data) {
    const purchaseOrderInput = document.getElementById('purchaseOrderInput');

    // データの前処理
    data = data.trim();

    // 🔥 発注番号バーコードパターン: 8桁の数字 + アルファベット1文字 (例: 00088333P)
    const purchaseOrderBarcodePattern = /^(\d{8})[A-Za-z]$/;
    const barcodeMatch = data.match(purchaseOrderBarcodePattern);

    if (barcodeMatch) {
        // 発注番号バーコード形式を検出
        const numericPart = barcodeMatch[1];  // 8桁の数字部分

        // 先頭のゼロを除去して発注番号を取得
        const orderNumber = String(parseInt(numericPart, 10));

        console.log(`発注番号バーコード検出: ${data} → ${orderNumber}`);

        if (purchaseOrderInput) purchaseOrderInput.value = orderNumber;
        stopQRScanner();
        showBarcodeReceivePopup(orderNumber);
        return;
    }

    // 🔥 誤読み取り防止: 数字+アルファベットの混在パターンを検出（無効なバーコード）
    // ただし8桁以上の連続数字を含む場合は発注番号として処理するため除外
    const hasEightDigits = /\d{8}/.test(data);
    if (!hasEightDigits) {
        const invalidBarcodePattern = /^\d*[A-Za-z]+\d*[A-Za-z]*$/;
        if (invalidBarcodePattern.test(data) && data.length >= 8) {
            console.log(`無効なバーコード形式を検出: ${data}`);
            showScannerStatus(`⚠️ 読み取りエラー: ${data}（再スキャンしてください）`, 'error');
            playBeep(false);
            // スキャナーを再開
            isScannerPaused = false;
            scanHistory.delete(data);  // 履歴から削除して再スキャン可能に
            return;
        }
    }

    // QRコードのフォーマットをチェック
    if (data.toUpperCase().startsWith('PO:')) {
        // 発注番号のQRコード
        const orderNumber = data.substring(3);
        if (purchaseOrderInput) purchaseOrderInput.value = orderNumber;
        stopQRScanner();
        showBarcodeReceivePopup(orderNumber);
    } else if (data.toUpperCase().startsWith('ORDER:')) {
        // 注文IDのQRコード
        const orderId = data.substring(6);
        stopQRScanner();
        if (typeof showOrderDetails === 'function') {
            showOrderDetails(parseInt(orderId));
        }
    } else if (/^\d{4,6}$/.test(data)) {
        // 4-6桁の数字は発注番号として扱う
        if (purchaseOrderInput) purchaseOrderInput.value = data;
        stopQRScanner();
        showBarcodeReceivePopup(data);
    } else {
        // 🔥 テキスト内から8桁以上の連続数字を抽出（例: "MHT0620エキシボリ00088066" → "00088066"）
        // 全マッチから最も長い（最後の）8桁以上の数字列を使う
        const allDigitMatches = data.match(/\d{8,}/g);
        if (allDigitMatches) {
            const longestMatch = allDigitMatches[allDigitMatches.length - 1];
            const numericPart = longestMatch.slice(-8);
            const orderNumber = String(parseInt(numericPart, 10));
            console.log(`QRテキストから発注番号抽出: ${data} → ${orderNumber}`);
            if (purchaseOrderInput) purchaseOrderInput.value = orderNumber;
            stopQRScanner();
            showBarcodeReceivePopup(orderNumber);
        } else if (data.toUpperCase().startsWith('MHT')) {
            // 製番のみで発注番号が含まれない場合
            stopQRScanner();
            alert(`製番 ${data} を検出しました。\n発注番号が含まれていません。`);
        } else {
            // その他のフォーマット
            if (purchaseOrderInput) purchaseOrderInput.value = data;
            stopQRScanner();

            if (confirm(`読み取った値: ${data}\n\nこれを発注番号として検索しますか？`)) {
                showBarcodeReceivePopup(data);
            }
        }
    }
}

// ========================================
// QRスキャナーを停止
// ========================================
function stopQRScanner() {
    console.log('スキャナー停止処理開始');

    isModalOpen = false;
    isScannerPaused = false;
    scanHistory.clear();
    lastScannedCode = null;
    scanCooldown = false;

    // モーダルを閉じる
    document.getElementById('qrScannerModal').classList.remove('show');

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
// バーコード受入確認ポップアップを表示
// ========================================
async function showBarcodeReceivePopup(orderNumber) {
    try {
        // 発注番号でDBを検索
        const response = await fetch(`/api/search-by-purchase-order/${orderNumber}`);
        const data = await response.json();

        const modalBody = document.getElementById('barcodeReceiveModalBody');

        if (!data.found || data.details.length === 0) {
            // 見つからない場合
            modalBody.innerHTML = `
                <div style="text-align: center; padding: 20px;">
                    <div style="font-size: 3em; margin-bottom: 20px;">❌</div>
                    <h3 style="color: #dc3545;">発注番号が見つかりません</h3>
                    <p style="color: #6c757d;">発注番号: <strong>${orderNumber}</strong></p>
                    <p style="font-size: 0.9em; color: #6c757d;">
                        この発注番号はデータベースに登録されていません。<br>
                        マージ処理が必要な可能性があります。
                    </p>
                    <button class="btn btn-secondary" onclick="closeBarcodeReceiveModal()" style="margin-top: 20px;">
                        閉じる
                    </button>
                </div>
            `;
            document.getElementById('barcodeReceiveModal').classList.add('show');
            return;
        }

        // マージ済みデータのみ抽出
        const mergedDetails = data.details.filter(d => d.source === 'merged');

        if (mergedDetails.length === 0) {
            // マージ済みデータがない場合
            modalBody.innerHTML = `
                <div style="text-align: center; padding: 20px;">
                    <div style="font-size: 3em; margin-bottom: 20px;">⚠️</div>
                    <h3 style="color: #ffc107;">未マージデータです</h3>
                    <p style="color: #6c757d;">発注番号: <strong>${orderNumber}</strong></p>
                    <p style="font-size: 0.9em; color: #6c757d;">
                        このデータは未マージのため受入処理ができません。<br>
                        先にマージ処理を実行してください。
                    </p>
                    <button class="btn btn-secondary" onclick="closeBarcodeReceiveModal()" style="margin-top: 20px;">
                        閉じる
                    </button>
                </div>
            `;
            document.getElementById('barcodeReceiveModal').classList.add('show');
            return;
        }

        // 受入確認画面を生成
        let html = `
            <div style="text-align: center; margin-bottom: 20px;">
                <div style="font-size: 2.5em; margin-bottom: 10px;">📦</div>
                <h3 style="margin: 0;">受入しますか？</h3>
                <p style="color: #6c757d; margin: 5px 0;">発注番号: <strong style="font-size: 1.2em;">${orderNumber}</strong></p>
            </div>
        `;

        mergedDetails.forEach((detail, index) => {
            const isReceived = detail.is_received;
            const bgColor = isReceived ? '#d4edda' : '#fff3cd';
            const borderColor = isReceived ? '#28a745' : '#ffc107';

            html += `
                <div style="background: ${bgColor}; border: 2px solid ${borderColor}; border-radius: 10px; padding: 15px; margin-bottom: 10px;">
                    ${isReceived ? '<div style="color: #28a745; font-weight: bold; margin-bottom: 10px;">✅ 受入済み</div>' : ''}
                    <table style="width: 100%; font-size: 0.95em;">
                        <tr>
                            <td style="font-weight: bold; width: 80px; padding: 3px 0;">製番:</td>
                            <td style="padding: 3px 0;">${detail.seiban || '-'}</td>
                        </tr>
                        <tr>
                            <td style="font-weight: bold; padding: 3px 0;">ユニット:</td>
                            <td style="padding: 3px 0;">${detail.unit || '-'}</td>
                        </tr>
                        <tr>
                            <td style="font-weight: bold; padding: 3px 0;">発注番号:</td>
                            <td style="padding: 3px 0;">${orderNumber}</td>
                        </tr>
                        <tr>
                            <td style="font-weight: bold; padding: 3px 0;">品名:</td>
                            <td style="padding: 3px 0;">${detail.item_name || '-'}</td>
                        </tr>
                        <tr>
                            <td style="font-weight: bold; padding: 3px 0;">仕様1:</td>
                            <td style="padding: 3px 0;">${detail.spec1 || '-'}</td>
                        </tr>
                        <tr>
                            <td style="font-weight: bold; padding: 3px 0;">個数:</td>
                            <td style="padding: 3px 0;"><strong style="font-size: 1.1em;">${detail.quantity || '-'} ${detail.unit_measure || ''}</strong></td>
                        </tr>
                    </table>
                    ${!isReceived ? `
                        <button class="btn btn-success" onclick="executeBarcodeReceive(${detail.id}, '${orderNumber}')"
                                style="width: 100%; margin-top: 15px; padding: 12px; font-size: 1.1em;">
                            ✅ この品目を受入する
                        </button>
                    ` : ''}
                </div>
            `;
        });

        html += `
            <button class="btn btn-secondary" onclick="closeBarcodeReceiveModal()" style="width: 100%; margin-top: 10px;">
                キャンセル
            </button>
        `;

        modalBody.innerHTML = html;
        document.getElementById('barcodeReceiveModal').classList.add('show');

    } catch (error) {
        console.error('検索エラー:', error);
        alert('検索中にエラーが発生しました: ' + error.message);
    }
}

// ========================================
// バーコード受入を実行
// ========================================
async function executeBarcodeReceive(detailId, orderNumber) {
    try {
        const response = await fetch(`/api/detail/${detailId}/toggle-receive`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ is_received: true })
        });

        const data = await response.json();

        if (data.success) {
            // 成功メッセージを表示
            const modalBody = document.getElementById('barcodeReceiveModalBody');
            modalBody.innerHTML = `
                <div style="text-align: center; padding: 30px;">
                    <div style="font-size: 4em; margin-bottom: 20px;">✅</div>
                    <h2 style="color: #28a745; margin-bottom: 10px;">受入完了！</h2>
                    <p style="font-size: 1.1em;">発注番号: <strong>${orderNumber}</strong></p>
                    <button class="btn btn-primary" onclick="closeBarcodeReceiveModal(); if(typeof loadOrders === 'function') loadOrders();" style="margin-top: 20px; padding: 12px 30px;">
                        OK
                    </button>
                </div>
            `;

            // ビープ音
            playBeep(true);

            // バイブレーション
            if (navigator.vibrate) {
                navigator.vibrate([100, 50, 100]);
            }
        } else {
            alert('受入処理に失敗しました: ' + (data.error || '不明なエラー'));
        }
    } catch (error) {
        console.error('受入エラー:', error);
        alert('受入処理中にエラーが発生しました: ' + error.message);
    }
}

// ========================================
// バーコード受入モーダルを閉じる
// ========================================
function closeBarcodeReceiveModal() {
    document.getElementById('barcodeReceiveModal').classList.remove('show');
}
