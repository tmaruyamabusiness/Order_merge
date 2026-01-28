// ========================================
// QRコード・バーコードスキャナーモジュール
// index.htmlの886行目～1210行目付近から抽出
// ========================================

// ========================================
// グローバル変数
// ========================================
// index.htmlの875行目～883行目から抽出
let qrScanning = false;
let qrVideo = null;
let qrCanvas = null;
let qrContext = null;
let qrStream = null;
let scanInterval = null;
let lastScannedCode = '';
let scanCount = 0;
let isModalOpen = false;

// ========================================
// メイン関数: QRスキャナーを開始
// ========================================
// index.htmlの886行目から抽出
async function startQRScanner() {
    // すでにスキャン中またはモーダルが開いている場合は終了
    if (isModalOpen || qrScanning) {
        console.log('Scanner already running or modal already open');
        return;
    }
    isModalOpen = true;

    try {
        // モーダルを表示
        document.getElementById('qrScannerModal').classList.add('show');
        
        // ビデオとキャンバス要素を取得
        qrVideo = document.getElementById('qrVideo');
        qrCanvas = document.getElementById('qrCanvas');
        qrContext = qrCanvas.getContext('2d', { willReadFrequently: true });
        
        // カメラの制約を設定（オートフォーカス有効化）
        const constraints = {
            video: {
                facingMode: 'environment',  // 背面カメラを優先
                width: { ideal: 1280 },     // 解像度を上げる
                height: { ideal: 720 },
                focusMode: { ideal: 'continuous' },  // 連続オートフォーカス
                focusDistance: { ideal: 0 },
                exposureMode: { ideal: 'continuous' },
                whiteBalanceMode: { ideal: 'continuous' }
            }
        };
        
        // カメラにアクセス
        try {
            qrStream = await navigator.mediaDevices.getUserMedia(constraints);
        } catch (error) {
            // フォールバック：基本的な制約で再試行
            console.log('高度な制約でカメラアクセス失敗、基本制約で再試行');
            qrStream = await navigator.mediaDevices.getUserMedia({
                video: { facingMode: 'environment' }
            });
        }
        
        qrVideo.srcObject = qrStream;
        await qrVideo.play();
        qrScanning = true;
        scanCount = 0;
        
        // startCombinedScanning()を呼び出し（次の関数）
        startCombinedScanning();
        
    } catch (error) {
        console.error('カメラアクセスエラー:', error);
        alert('カメラにアクセスできません。\nカメラの使用を許可してください。\n\nエラー: ' + error.message);
        stopQRScanner();
    }
}

// ========================================
// 複合スキャン: QRコード + バーコード
// ========================================
// index.htmlの943行目から抽出
// startQRScanner()から呼び出される
function startCombinedScanning() {
    if (!qrScanning) return;
    
    // 100ミリ秒ごとにscanFrame()を実行（フレームレートを上げる）
    scanInterval = setInterval(() => {
        if (qrScanning && qrVideo.readyState === qrVideo.HAVE_ENOUGH_DATA) {
            scanFrame();
        }
    }, 100);
}

// ========================================
// フレームをスキャン
// ========================================
// index.htmlの955行目から抽出
// startCombinedScanning()から100msごとに呼び出される
function scanFrame() {
    if (!qrScanning || !qrVideo) return;
    
    // キャンバスのサイズを設定
    const videoWidth = qrVideo.videoWidth;
    const videoHeight = qrVideo.videoHeight;
    
    if (videoWidth === 0 || videoHeight === 0) return;
    
    qrCanvas.width = videoWidth;
    qrCanvas.height = videoHeight;
    
    // ビデオフレームをキャンバスに描画
    qrContext.drawImage(qrVideo, 0, 0, videoWidth, videoHeight);
    
    // 画像データを取得
    const imageData = qrContext.getImageData(0, 0, videoWidth, videoHeight);
    
    // QRコードをスキャン（改善された設定）
    const qrCode = jsQR(imageData.data, imageData.width, imageData.height, {
        inversionAttempts: 'attemptBoth'  // 白黒反転も試す
    });
    
    // QRコードが検出された場合、handleCodeDetected()を呼び出して終了
    if (qrCode && qrCode.data) {
        handleCodeDetected(qrCode.data, 'QR');
        return;
    }
    
    // QRコードが見つからない場合はバーコードをスキャン（scanBarcode()を呼び出し）
    scanBarcode(qrCanvas);
    
    scanCount++;
    
    // スキャン状態を表示（updateScanStatus()を呼び出し）
    updateScanStatus();
}

// ========================================
// バーコードスキャン
// ========================================
// index.htmlの993行目から抽出
// scanFrame()から呼び出される
function scanBarcode(canvas) {
    if (!window.Quagga) return;
    
    // Canvasから画像データを取得
    const dataURL = canvas.toDataURL('image/png');
    
    Quagga.decodeSingle({
        decoder: {
            readers: [
                'code_128_reader',
                'ean_reader',
                'ean_8_reader',
                'code_39_reader',
                'code_39_vin_reader',
                'codabar_reader',
                'upc_reader',
                'upc_e_reader',
                'i2of5_reader'
            ]
        },
        locate: true,
        src: dataURL
    }, function(result) {
        // バーコードが検出された場合、handleCodeDetected()を呼び出し
        if (result && result.codeResult) {
            handleCodeDetected(result.codeResult.code, 'Barcode');
        }
    });
}

// ========================================
// コードが検出された時の処理
// ========================================
// index.htmlの1023行目から抽出
// scanFrame()またはscanBarcode()から呼び出される
function handleCodeDetected(data, type) {
    // 同じコードの連続読み取りを防ぐ
    if (data === lastScannedCode) {
        return;
    }
    
    lastScannedCode = data;
    console.log(`${type}コード検出:`, data);
    
    // 結果をqrResult要素に表示
    document.getElementById('qrResult').innerHTML = `
        <div style="color: #28a745; font-weight: bold; font-size: 1.2em;">
            ✅ ${type}コード検出成功！
        </div>
        <div style="margin-top: 10px; font-size: 1.1em;">
            読み取り値: <strong>${data}</strong>
        </div>
        <div style="margin-top: 10px; color: #6c757d;">
            処理中...
        </div>
    `;
    
    // ビープ音を鳴らす（playBeep()を呼び出し）
    playBeep();
    
    // バイブレーション（対応デバイスのみ）
    if (navigator.vibrate) {
        navigator.vibrate(200);
    }
    
    // スキャンを一時停止
    qrScanning = false;
    if (scanInterval) {
        clearInterval(scanInterval);
        scanInterval = null;
    }
    
    // 1秒後にprocessScannedCode()を呼び出し
    setTimeout(() => {
        processScannedCode(data);
    }, 1000);
}

// ========================================
// スキャンしたコードを処理
// ========================================
// index.htmlの1067行目から抽出
// handleCodeDetected()から1秒後に呼び出される
function processScannedCode(data) {
    const purchaseOrderInput = document.getElementById('purchaseOrderInput');

    // データの前処理（空白除去、大文字変換など）
    data = data.trim().toUpperCase();

    // QRコードのフォーマットをチェック
    if (data.startsWith('PO:')) {
        // 発注番号のQRコード
        const orderNumber = data.replace('PO:', '');
        purchaseOrderInput.value = orderNumber;
        stopQRScanner();
        showBarcodeReceivePopup(orderNumber);  // 受入確認ポップアップを表示
    } else if (data.startsWith('ORDER:')) {
        // 注文IDのQRコード
        const orderId = data.replace('ORDER:', '');
        stopQRScanner();
        showOrderDetails(parseInt(orderId));  // index.htmlの関数を呼び出し
    } else if (/^\d{5,6}$/.test(data)) {
        // 5-6桁の数字は発注番号として扱う
        purchaseOrderInput.value = data;
        stopQRScanner();
        showBarcodeReceivePopup(data);  // 受入確認ポップアップを表示
    } else if (data.startsWith('MHT')) {
        // 製番の場合
        stopQRScanner();
        alert(`製番 ${data} を検出しました。\n製番での検索機能を実装予定です。`);
    } else {
        // その他のフォーマット
        purchaseOrderInput.value = data;
        stopQRScanner();

        // ユーザーに選択させる
        if (confirm(`読み取った値: ${data}\n\nこれを発注番号として検索しますか？`)) {
            showBarcodeReceivePopup(data);  // 受入確認ポップアップを表示
        }
    }
}

// ========================================
// 🔥 バーコード受入確認ポップアップを表示
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
// 🔥 バーコード受入を実行
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
                    <button class="btn btn-primary" onclick="closeBarcodeReceiveModal(); loadOrders();" style="margin-top: 20px; padding: 12px 30px;">
                        OK
                    </button>
                </div>
            `;

            // ビープ音
            playBeep();

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
// 🔥 バーコード受入モーダルを閉じる
// ========================================
function closeBarcodeReceiveModal() {
    document.getElementById('barcodeReceiveModal').classList.remove('show');
}

// ========================================
// スキャン状態を更新
// ========================================
// index.htmlの1106行目から抽出
// scanFrame()から呼び出される
function updateScanStatus() {
    const statusDiv = document.getElementById('qrResult');
    if (!statusDiv || lastScannedCode) return;
    
    const messages = [
        'QRコード/バーコードを探しています...',
        'カメラにコードを向けてください',
        'もう少し近づけてみてください',
        'コードが画面中央にくるように調整してください'
    ];
    
    const messageIndex = Math.floor(scanCount / 10) % messages.length;
    
    statusDiv.innerHTML = `
        <div style="color: #17a2b8;">
            <div class="spinner" style="margin: 0 auto 10px;"></div>
            ${messages[messageIndex]}
        </div>
        <div style="margin-top: 10px; font-size: 0.9em; color: #6c757d;">
            スキャン回数: ${scanCount}
        </div>
    `;
}

// ========================================
// QRコードスキャナーを停止
// ========================================
// index.htmlの1131行目から抽出
// processScannedCode()や外部から呼び出される
function stopQRScanner() {
    console.log('stopQRScanner called');
    cleanup();  // cleanup()を呼び出してリソースを解放
    const modal = document.getElementById('qrScannerModal');
    modal.classList.remove('show');
    document.getElementById('qrResult').innerHTML = '';
    console.log('QR Scanner stopped');

    // インターバルをクリア
    if (scanInterval) {
        clearInterval(scanInterval);
        scanInterval = null;
    }
    
    // カメラストリームを停止
    if (qrStream) {
        qrStream.getTracks().forEach(track => {
            track.stop();
            console.log('Track stopped:', track.label);
        });
        qrStream = null;
    }
    
    // ビデオを停止
    if (qrVideo) {
        qrVideo.pause();
        qrVideo.srcObject = null;
        qrVideo = null;
    }
    
    // Quaggaを停止
    if (window.Quagga && Quagga.stop) {
        Quagga.stop();
    }
    
    // モーダルを閉じる
    document.getElementById('qrScannerModal').classList.remove('show');
    document.getElementById('qrResult').innerHTML = '';
}

// ========================================
// クリーンアップ処理
// ========================================
// index.htmlの1171行目から抽出
// stopQRScanner()から呼び出される
function cleanup() {
    qrScanning = false;
    isModalOpen = false;
    lastScannedCode = '';
    scanCount = 0;
    
    if (scanInterval) {
        clearInterval(scanInterval);
        scanInterval = null;
    }
    
    if (qrStream) {
        qrStream.getTracks().forEach(track => track.stop());
        qrStream = null;
    }
    
    if (qrVideo) {
        qrVideo.pause();
        qrVideo.srcObject = null;
        qrVideo = null;
    }
    
    qrContext = null;
}

// ========================================
// ビープ音を鳴らす
// ========================================
// index.htmlの1197行目から抽出
// handleCodeDetected()から呼び出される
function playBeep() {
    try {
        const audioContext = new (window.AudioContext || window.webkitAudioContext)();
        const oscillator = audioContext.createOscillator();
        const gainNode = audioContext.createGain();
        
        oscillator.connect(gainNode);
        gainNode.connect(audioContext.destination);
        
        // より聞き取りやすい音に調整
        oscillator.frequency.value = 1000;  // 周波数を上げる
        oscillator.type = 'sine';
        gainNode.gain.value = 0.3;  // 音量を上げる
        
        oscillator.start(audioContext.currentTime);
        oscillator.stop(audioContext.currentTime + 0.2);  // 少し長めに
    } catch (e) {
        console.log('ビープ音の再生に失敗しました:', e);
    }
}

// ========================================
// 初期化処理
// ========================================
// DOMContentLoaded後に実行
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initQRScanner);
} else {
    initQRScanner();
}

function initQRScanner() {
    // スキャンアニメーション用のCSSを追加
    const style = document.createElement('style');
    style.textContent = `
        @keyframes scanning {
            0% { border-color: #667eea; }
            50% { border-color: #28a745; }
            100% { border-color: #667eea; }
        }
        
        #qrVideo.scanning {
            animation: scanning 2s infinite;
            border-width: 3px;
        }
        
        .scan-guide {
            position: absolute;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            width: 200px;
            height: 200px;
            border: 2px dashed rgba(255, 255, 255, 0.5);
            pointer-events: none;
        }
    `;
    document.head.appendChild(style);

    // Escキーでスキャナーを閉じる
    document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape' && qrScanning) {
            stopQRScanner();
        }
    });
}