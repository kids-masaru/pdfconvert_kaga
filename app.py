<!-- index.html -->
<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Excelマージツール</title>
  <link rel="manifest" href="{{ url_for('static', filename='manifest.json') }}">
  <link rel="icon" href="{{ favicon_url }}">
  <meta name="theme-color" content="#4a90e2">
  <link rel="stylesheet" href="{{ url_for('static', filename='css/style.css') }}">
</head>
<body>
  <!-- ローディングオーバーレイ -->
  <div id="loading-overlay" class="hidden">
    <div class="spinner"></div>
    <p id="loading-text">処理中です。しばらくお待ちください… (<span id="progress">0%</span>)</p>
  </div>

  <header class="header">
    <img src="{{ url_for('static', filename='icons/icon-192.png') }}" alt="App Icon">
    <div>
      <h1 class="title">Excelマージツール</h1>
      <p class="subtitle">Excelファイルをアップロードしてテンプレートと結合</p>
    </div>
  </header>

  <div class="card upload-area" id="drop-area">
    <div class="upload-icon">⇪</div>
    <h2>Upload your Excel</h2>
    <p>Excelファイルをここにドラッグ＆ドロップ、または以下で選択</p>
    <div>
      <label class="btn btn-excel">Excelを選択<input type="file" id="excel-input" accept=".xls,.xlsx,.xlsm" style="display:none;"></label>
    </div>
    <ul id="file-list" class="file-list"></ul>
    <button id="process-btn" class="btn-process">処理開始</button>
  </div>

  <script src="{{ url_for('static', filename='js/main.js') }}"></script>
</body>
</html>

<!-- main.js -->
// static/js/main.js

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
        if (name.endsWith('.xls') || name.endsWith('.xlsx') || name.endsWith('.xlsm')) {
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
          const blob = xhr.response;
          const url = window.URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          // ダウンロードファイル名を元ファイル名_Merged.xlsm に設定
          const baseName = excelFile.name.replace(/\.(xlsx?|xlsm)$/, '');
          a.download = `${baseName}_Merged.xlsm`;
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
