// static/js/main.js

document.addEventListener('DOMContentLoaded', () => {
  const dropArea = document.getElementById('drop-area');
  const pdfInput = document.getElementById('pdf-input');
  const excelInput = document.getElementById('excel-input');
  const fileList = document.getElementById('file-list');
  const form = document.getElementById('upload-form');
  const loading = document.getElementById('loading-overlay');

  // Prevent default drag behaviors
  ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
    dropArea.addEventListener(eventName, e => {
      e.preventDefault();
      e.stopPropagation();
    });
  });

  // Highlight on dragover
  dropArea.addEventListener('dragover', () => {
    dropArea.classList.add('highlight');
  });
  dropArea.addEventListener('dragleave', () => {
    dropArea.classList.remove('highlight');
  });

  // Handle drop
  dropArea.addEventListener('drop', e => {
    dropArea.classList.remove('highlight');
    const dt = e.dataTransfer;
    handleFiles(dt.files);
  });

  // Handle file selection
  pdfInput.addEventListener('change', () => handleFiles(pdfInput.files));
  excelInput.addEventListener('change', () => handleFiles(excelInput.files));

  function handleFiles(files) {
    // Clear existing list
    fileList.innerHTML = '';
    Array.from(files).forEach(file => {
      const li = document.createElement('li');
      li.textContent = file.name;
      fileList.appendChild(li);
    });
  }

  // On form submit, show loader
  form.addEventListener('submit', () => {
    loading.classList.remove('hidden');
  });
});
