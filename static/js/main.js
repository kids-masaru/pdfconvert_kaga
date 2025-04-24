// static/js/main.js

document.addEventListener('DOMContentLoaded', () => {
  const dropArea = document.getElementById('drop-area');
  const pdfInput = document.getElementById('pdf-input');
  const excelInput = document.getElementById('excel-input');
  const fileList = document.getElementById('file-list');
  const processBtn = document.getElementById('process-btn');
  const loading = document.getElementById('loading-overlay');
  const progressText = document.getElementById('progress');

  let pdfFile = null;
  let excelFile = null;

  // Prevent default drag behaviors
  ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(ev => {
    dropArea.addEventListener(ev, e => { e.preventDefault(); e.stopPropagation(); });
  });
  dropArea.addEventListener('dragover', () => dropArea.classList.add('highlight'));
  dropArea.addEventListener('dragleave', () => dropArea.classList.remove('highlight'));

  // Handle drop
  dropArea.addEventListener('drop', e => {
    dropArea.classList.remove('highlight');
    handleFiles(e.dataTransfer.files);
  });

  // Handle manual select
  pdfInput.addEventListener('change', () => {
    pdfFile = pdfInput.files[0] || null;
    updateFileList();
  });
  excelInput.addEventListener('change', () => {
    excelFile = excelInput.files[0] || null;
    updateFileList();
  });

  function handleFiles(files) {
    Array.from(files).forEach(f => {
      if (f.name.toLowerCase().endsWith('.pdf')) pdfFile = f;
      else if (f.name.toLowerCase().match(/\\.xls/)) excelFile = f;
    });
    updateFileList();
  }

  function updateFileList() {
    fileList.innerHTML = '';
    [ ['PDF', pdfFile], ['Excel', excelFile] ].forEach(([label, file]) => {
      if (file) {
        const li = document.createElement('li');
        li.innerHTML = `
          <span class="info">${label}: ${file.name}</span>
          <button type="button" class="btn-remove">✕</button>
        `;
        li.querySelector('.btn-remove').addEventListener('click', () => {
          if (label === 'PDF') {
            pdfFile = null; pdfInput.value = '';
          } else {
            excelFile = null; excelInput.value = '';
          }
          updateFileList();
        });
        fileList.appendChild(li);
      }
    });
  }

  // Process
  processBtn.addEventListener('click', () => {
    if (!pdfFile || !excelFile) {
      alert('PDFとExcelの両方をアップロードしてください。');
      return;
    }
    // Show loading
    loading.classList.remove('hidden');
    progressText.textContent = '0%';

    const formData = new FormData();
    formData.append('pdf_file', pdfFile);
    formData.append('excel_file', excelFile);

    const xhr = new XMLHttpRequest();
    xhr.open('POST', '/process');
    xhr.responseType = 'blob';

    xhr.upload.onprogress = e => {
      if (e.lengthComputable) {
        const percent = Math.round((e.loaded / e.total) * 100);
        progressText.textContent = percent + '%';
      }
    };

    xhr.onload = () => {
      loading.classList.add('hidden');
      if (xhr.status === 200) {
        // ダウンロード
        const blob = xhr.response;
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        // 固定ファイル名、またはヘッダーを解析してもOK
        a.download = 'Combined_Result.xlsm';
        a.href = url;
        a.click();
        window.URL.revokeObjectURL(url);
      } else {
        alert('処理中にエラーが発生しました。');
      }
    };

    xhr.onerror = () => {
      loading.classList.add('hidden');
      alert('通信エラーが発生しました。');
    };

    xhr.send(formData);
  });
});
