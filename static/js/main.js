document.addEventListener('DOMContentLoaded', () => {
    const dropAreaExcel = document.getElementById('drop-area-excel');
    const dropAreaPdf = document.getElementById('drop-area-pdf');
    const excelInput = document.getElementById('excel-input');
    const pdfInput = document.getElementById('pdf-input');
    const fileList = document.getElementById('file-list');
    const processBtn = document.getElementById('process-btn');
    // const loadingOverlay = document.getElementById('loading-overlay'); // 削除
    // const progressText = document.getElementById('progress'); // 削除

    let excelFile = null;
    let pdfFiles = [];

    // --- ファイル追加・削除共通関数 ---
    function addFileToList(file, type) {
        const li = document.createElement('li');
        const infoDiv = document.createElement('div');
        infoDiv.classList.add('info');
        
        let iconClass = '';
        if (type === 'excel') {
            iconClass = 'fas fa-file-excel';
        } else if (type === 'pdf') {
            iconClass = 'fas fa-file-pdf';
        }
        
        infoDiv.innerHTML = `<i class="${iconClass}"></i> <span class="name">${file.name}</span>`;
        li.appendChild(infoDiv);

        const removeBtn = document.createElement('button');
        removeBtn.classList.add('btn-remove');
        removeBtn.innerHTML = '<i class="fas fa-times"></i>'; // ✕アイコン
        removeBtn.addEventListener('click', () => {
            if (type === 'excel') {
                excelFile = null;
                dropAreaExcel.classList.remove('highlight');
            } else if (type === 'pdf') {
                pdfFiles = pdfFiles.filter(f => f !== file);
                if (pdfFiles.length === 0) {
                    dropAreaPdf.classList.remove('highlight');
                }
            }
            li.remove();
            updateProcessButtonState();
        });
        li.appendChild(removeBtn);
        fileList.appendChild(li);
        
        updateProcessButtonState();
    }

    function updateFileList() {
        fileList.innerHTML = ''; // リストをクリア
        if (excelFile) {
            addFileToList(excelFile, 'excel');
        }
        pdfFiles.forEach(file => addFileToList(file, 'pdf'));
        updateProcessButtonState();
    }

    function updateProcessButtonState() {
        if (excelFile && pdfFiles.length > 0) {
            processBtn.disabled = false;
        } else {
            processBtn.disabled = true;
        }
    }

    // --- ドラッグ＆ドロップ処理 ---
    function setupDropArea(dropArea, fileInput, type) {
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            dropArea.addEventListener(eventName, preventDefaults, false);
        });

        function preventDefaults(e) {
            e.preventDefault();
            e.stopPropagation();
        }

        ['dragenter', 'dragover'].forEach(eventName => {
            dropArea.addEventListener(eventName, () => dropArea.classList.add('highlight'), false);
        });

        ['dragleave', 'drop'].forEach(eventName => {
            dropArea.addEventListener(eventName, () => dropArea.classList.remove('highlight'), false);
        });

        dropArea.addEventListener('drop', (e) => {
            const dt = e.dataTransfer;
            const files = dt.files;

            if (type === 'excel') {
                const newFile = files[0];
                if (newFile && (newFile.name.endsWith('.xls') || newFile.name.endsWith('.xlsx'))) {
                    excelFile = newFile;
                    dropArea.classList.add('highlight');
                } else {
                    alert('Excelファイルのみをドロップしてください。');
                }
            } else if (type === 'pdf') {
                const newPdfs = Array.from(files).filter(file => file.name.endsWith('.pdf'));
                if (newPdfs.length > 0) {
                    pdfFiles = [...pdfFiles, ...newPdfs];
                    dropArea.classList.add('highlight');
                } else {
                    alert('PDFファイルのみをドロップしてください。');
                }
            }
            updateFileList();
        }, false);

        fileInput.addEventListener('change', (e) => {
            const files = e.target.files;
            if (type === 'excel') {
                if (files.length > 0) {
                    excelFile = files[0];
                    dropArea.classList.add('highlight');
                }
            } else if (type === 'pdf') {
                if (files.length > 0) {
                    pdfFiles = [...pdfFiles, ...Array.from(files)];
                    dropArea.classList.add('highlight');
                }
            }
            updateFileList();
        });
    }

    setupDropArea(dropAreaExcel, excelInput, 'excel');
    setupDropArea(dropAreaPdf, pdfInput, 'pdf');

    // --- 処理開始ボタン ---
    processBtn.addEventListener('click', async () => {
        if (!excelFile || pdfFiles.length === 0) {
            alert('ExcelファイルとPDFファイルの両方をアップロードしてください。');
            return;
        }

        const formData = new FormData();
        formData.append('excel_file', excelFile);
        pdfFiles.forEach(file => {
            formData.append('pdf_files', file);
        });

        // loadingOverlay.classList.remove('hidden'); // 削除
        // progressText.textContent = '0%'; // 削除

        try {
            const response = await fetch('/upload_and_process', {
                method: 'POST',
                body: formData,
            });

            // loadingOverlay.classList.add('hidden'); // 削除

            if (response.ok) {
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const disposition = response.headers.get('Content-Disposition');
                let filename = 'processed_data.xlsx';
                if (disposition && disposition.indexOf('attachment') !== -1) {
                    const filenameRegex = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/;
                    const matches = filenameRegex.exec(disposition);
                    if (matches != null && matches[1]) {
                        filename = decodeURIComponent(matches[1].replace(/['"]/g, ''));
                    }
                }

                const a = document.createElement('a');
                a.href = url;
                a.download = filename;
                document.body.appendChild(a);
                a.click();
                a.remove();
                window.URL.revokeObjectURL(url);

                // 成功したらファイルをクリア
                excelFile = null;
                pdfFiles = [];
                updateFileList();
                dropAreaExcel.classList.remove('highlight');
                dropAreaPdf.classList.remove('highlight');
                alert('処理が完了し、ファイルがダウンロードされました。');

            } else {
                const errorText = await response.text();
                // エラー表示のため、error.htmlにリダイレクト
                window.location.href = `/error?message=${encodeURIComponent(errorText)}`;
            }
        } catch (error) {
            // loadingOverlay.classList.add('hidden'); // 削除
            console.error('Fetch error:', error);
            // エラー表示のため、error.htmlにリダイレクト
            window.location.href = `/error?message=${encodeURIComponent('ファイルのアップロードまたは処理中にネットワークエラーが発生しました。')}`;
        }
    });

    updateFileList(); // 初期表示でボタンの状態を更新
});
