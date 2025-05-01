document.addEventListener('DOMContentLoaded', () => {
  const dropArea   = document.getElementById('drop-area');
  const excelInput = document.getElementById('excel-input');
  const fileList   = document.getElementById('file-list');
  const processBtn = document.getElementById('process-btn');
  const loading    = document.getElementById('loading-overlay');
  const progressEl = document.getElementById('progress');

  let excelFile = null;
  let fakeProgressTimer;

  // ドラッグ＆ドロップ時のデフォルト挙動抑制
  ['dragenter','dragover','dragleave','drop'].forEach(ev => {
    dropArea.addEventListener(ev, e => { e.preventDefault(); e.stopPropagation(); });
  });
  dropArea.addEventListener('dragover', () => dropArea.classList.add('highlight'));
  dropArea.addEventListener('dragleave', () => dropArea.classList.remove('highlight'));

  // drop された時
  dropArea.addEventListener('drop', e => {
    dropArea.classList.remove('highlight');
    handleFiles(e.dataTransfer.files);
  });

  // ファイル選択ボタンから
  excelInput.addEventListener('change', () => {
    excelFile = excelInput.files[0] || null;
    updateFileList();
  });

  // ファイル判定
  function handleFiles(files) {
    Array.from(files).forEach(f => {
      const name = f.name.toLowerCase();
      if (name.endsWith('.xls') || name.endsWith('.xlsx')) {
        excelFile = f;
      }
    });
    updateFileList();
  }

  // ファイル一覧描画
  function updateFileList() {
    fileList.innerHTML = '';
    if (excelFile) {
      const li = document.createElement('li');
      li.innerHTML = `
        <span class="info">Excel: ${excelFile.name}</span>
        <button type="button" class="btn-remove">✕</button>
      `;
      li.querySelector('.btn-remove').addEventListener('click', () => {
        excelFile = null;
        excelInput.value = '';
        updateFileList();
      });
      fileList.appendChild(li);
    }
  }

  // 送信ボタン
  processBtn.addEventListener('click', () => {
    if (!excelFile) {
      alert('Excelファイルをアップロードしてください。');
      return;
    }

    // ローディング表示＆プログレス初期化
    loading.classList.remove('hidden');
    progressEl.textContent = '0%';

    // 疑似プログレスバー開始
    let fakePercent = 0;
    clearInterval(fakeProgressTimer);
    fakeProgressTimer = setInterval(() => {
      if (fakePercent < 90) {
        fakePercent += 5;
        progressEl.textContent = fakePercent + '%';
      } else {
        clearInterval(fakeProgressTimer);
      }
    }, 500);

    // 実ファイル送信
    const formData = new FormData();
    formData.append('excel_file', excelFile);

    const xhr = new XMLHttpRequest();
    xhr.open('POST', '/process');
    xhr.responseType = 'blob';

    xhr.upload.onprogress = e => {
      if (e.lengthComputable) {
        const realPercent = Math.round((e.loaded / e.total) * 100);
        progressEl.textContent = realPercent + '%';
      }
    };

    xhr.onload = () => {
      clearInterval(fakeProgressTimer);
      progressEl.textContent = '100%';
      loading.classList.add('hidden');

      if (xhr.status === 200) {
        // ダウンロード処理
        const blob = xhr.response;
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'Processed_Result.xlsm';
        a.click();
        window.URL.revokeObjectURL(url);
      } else {
        alert('処理中にエラーが発生しました。');
      }
    };

    xhr.onerror = () => {
      clearInterval(fakeProgressTimer);
      loading.classList.add('hidden');
      alert('通信エラーが発生しました。');
    };

    xhr.send(formData);
  });
});
if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('/static/service-worker.js')
      .then(registration => {
        console.log('SW registered:', registration);
      }).catch(err => {
        console.log('SW registration failed:', err);
      });
  });
}
