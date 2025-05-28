// static/js/main.js

document.addEventListener('DOMContentLoaded', () => {
    const dropAreaExcel = document.getElementById('drop-area-excel');
    const dropAreaPdf = document.getElementById('drop-area-pdf');
    const excelInput = document.getElementById('excel-input');
    const pdfInput = document.getElementById('pdf-input');
    const fileList = document.getElementById('file-list');
    const processBtn = document.getElementById('process-btn');
    const loadingOverlay = document.getElementById('loading-overlay');
    const progressBar = document.getElementById('progress'); // ProgressBarは今回使わないが変数として残す

    let selectedExcelFile = null;
    let selectedPdfFiles = [];

    // --- ファイル選択/ドラッグ＆ドロップ処理 ---

    function handleFiles(files, type) {
        if (type === 'excel' && files.length > 0) {
            selectedExcelFile = files[0];
        } else if (type === 'pdf') {
            // 複数ファイルを扱うため、既存の配列に追加ではなく、置き換える
            selectedPdfFiles = Array.from(files).filter(file => file.type === 'application/pdf');
        }
        updateFileList();
    }

    function updateFileList() {
        fileList.innerHTML = ''; // リストをクリア
        if (selectedExcelFile) {
            const li = document.createElement('li');
            li.innerHTML = `
                <span class="info">📊 Excel: ${selectedExcelFile.name}</span>
                <button type="button" class="btn-remove" data-type="excel">✕</button>
            `;
            fileList.appendChild(li);
        }
        selectedPdfFiles.forEach(file => {
            const li = document.createElement('li');
            li.innerHTML = `
                <span class="info">📄 PDF: ${file.name}</span>
                <button type="button" class="btn-remove" data-type="pdf" data-name="${file.name}">✕</button>
            `;
            fileList.appendChild(li);
        });

        // ファイル削除ボタンのイベントリスナー
        fileList.querySelectorAll('.btn-remove').forEach(button => {
            button.addEventListener('click', () => {
                const type = button.dataset.type;
                if (type === 'excel') {
                    selectedExcelFile = null;
                    excelInput.value = ''; // input要素のリセット
                } else if (type === 'pdf') {
                    const nameToRemove = button.dataset.name;
                    selectedPdfFiles = selectedPdfFiles.filter(file => file.name !== nameToRemove);
                    // PDF inputは複数ファイル選択なので、個別にリセットは難しい。
                    // 必要であれば、個別のファイル選択ボタンを設けるか、複雑なロジックが必要。
                    // 今回は個別のファイル削除はリストから削除するのみとし、inputはそのまま。
                    // もし再度選択し直すなら、ユーザーはinputをもう一度クリックすることになる。
                }
                updateFileList();
            });
        });

        // 処理開始ボタンの有効/無効
        processBtn.disabled = !(selectedExcelFile && selectedPdfFiles.length > 0);
        processBtn.style.opacity = processBtn.disabled ? 0.5 : 1;
        processBtn.style.cursor = processBtn.disabled ? 'not-allowed' : 'pointer';
    }

    // ドラッグ＆ドロップ時のデフォルト挙動抑制
    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    // ドロップエリアのハイライト処理
    function highlightDropArea(element) {
        element.classList.add('highlight');
    }

    function unhighlightDropArea(element) {
        element.classList.remove('highlight');
    }

    // イベントリスナーの設定 (Excel)
    excelInput.addEventListener('change', (e) => handleFiles(e.target.files, 'excel'));
    ['dragenter', 'dragover'].forEach(eventName => {
        dropAreaExcel.addEventListener(eventName, () => highlightDropArea(dropAreaExcel), false);
    });
    ['dragleave', 'drop'].forEach(eventName => {
        dropAreaExcel.addEventListener(eventName, () => unhighlightDropArea(dropAreaExcel), false);
    });
    dropAreaExcel.addEventListener('drop', (e) => {
        preventDefaults(e);
        handleFiles(e.dataTransfer.files, 'excel');
    }, false);
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(ev => {
        dropAreaExcel.addEventListener(ev, preventDefaults, false);
    });


    // イベントリスナーの設定 (PDF)
    pdfInput.addEventListener('change', (e) => handleFiles(e.target.files, 'pdf'));
    ['dragenter', 'dragover'].forEach(eventName => {
        dropAreaPdf.addEventListener(eventName, () => highlightDropArea(dropAreaPdf), false);
    });
    ['dragleave', 'drop'].forEach(eventName => {
        dropAreaPdf.addEventListener(eventName, () => unhighlightDropArea(dropAreaPdf), false);
    });
    dropAreaPdf.addEventListener('drop', (e) => {
        preventDefaults(e);
        handleFiles(e.dataTransfer.files, 'pdf');
    }, false);
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(ev => {
        dropAreaPdf.addEventListener(ev, preventDefaults, false);
    });


    // --- 処理開始ボタンイベント ---

    processBtn.addEventListener('click', async () => {
        if (!selectedExcelFile || selectedPdfFiles.length === 0) {
            alert('ExcelファイルとPDFファイルを両方選択してください。');
            return;
        }

        loadingOverlay.classList.remove('hidden');
        progressBar.textContent = '処理中...'; // 進捗バーはシンプルに「処理中」表示

        const formData = new FormData();
        formData.append('excel_file', selectedExcelFile);
        selectedPdfFiles.forEach(file => {
            formData.append('pdf_files', file); // バックエンドで複数ファイルとして受け取るための名前
        });

        try {
            const response = await fetch('/upload_and_process', { // Flaskのエンドポイント
                method: 'POST',
                body: formData,
            });

            if (response.ok) {
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = 'processed_data.xlsx'; // ダウンロードファイル名
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a); // ダウンロード後、要素を削除
                window.URL.revokeObjectURL(url); // URLを解放

                alert('処理が完了し、ファイルがダウンロードされました。');
                // ファイル選択状態をリセット
                selectedExcelFile = null;
                selectedPdfFiles = [];
                excelInput.value = '';
                pdfInput.value = '';
                updateFileList();
            } else {
                const errorText = await response.text();
                alert(`エラーが発生しました: ${errorText}`);
                console.error('Server error:', errorText);
            }
        } catch (error) {
            console.error('Network or processing error:', error);
            alert('ファイルのアップロードまたは処理中にネットワークエラーが発生しました。');
        } finally {
            loadingOverlay.classList.add('hidden'); // 処理終了後はローディング表示を非表示に
        }
    });

    updateFileList(); // 初期状態を反映
});

// PWA関連 (現状維持でOK)
if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
        navigator.serviceWorker.register('/static/service-worker.js')
            .then(registration => {
                console.log('ServiceWorker登録成功:', registration.scope);
            })
            .catch(err => {
                console.log('ServiceWorker登録失敗:', err);
            });
    });
}

const manifestLink = document.createElement('link');
manifestLink.rel = 'manifest';
manifestLink.href = '/static/manifest.json';
document.head.appendChild(manifestLink);
