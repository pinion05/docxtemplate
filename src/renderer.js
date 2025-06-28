const { ipcRenderer } = require('electron');

document.addEventListener('DOMContentLoaded', () => {
  const templateFileInput = document.getElementById('templateFile');
  const placeholderContainer = document.getElementById('placeholderContainer');
  const generateBtn = document.getElementById('generateBtn');
  const resultDiv = document.getElementById('result');
  const previewFrame = document.getElementById('docxPreview');
  const printBtn = document.getElementById('printBtn');

  // 템플릿 선택 시 플레이스홀더 추출 및 동적 입력 생성
  templateFileInput.addEventListener('change', () => {
    const file = templateFileInput.files[0];
    if (!file) return;
    placeholderContainer.innerHTML = '플레이스홀더 분석 중...';
    generateBtn.disabled = true;

    const reader = new FileReader();
    reader.onload = (e) => {
      const buffer = e.target.result;
      ipcRenderer.invoke('extract-placeholders', Buffer.from(buffer)).then((placeholders) => {
        placeholderContainer.innerHTML = '';
        if (placeholders.length === 0) {
          placeholderContainer.textContent = '플레이스홀더가 없습니다.';
        }
        placeholders.forEach((name) => {
          const wrapper = document.createElement('div');
          wrapper.innerHTML = `<label>${name}</label><br><input type="text" class="ph-input" data-key="${name}" placeholder="${name}">`;
          placeholderContainer.appendChild(wrapper);
        });
        generateBtn.disabled = false;
      }).catch(err => {
        placeholderContainer.textContent = '추출 실패: ' + err;
      });
    };
    reader.readAsArrayBuffer(file);
  });

  generateBtn.addEventListener('click', () => {
    const file = templateFileInput.files[0];
    if (!file) {
      resultDiv.textContent = '템플릿 파일을 업로드하세요.';
      return;
    }
    // 동적으로 생성된 입력값 수집
    const data = {};
    placeholderContainer.querySelectorAll('input.ph-input').forEach((inp) => {
      data[inp.dataset.key] = inp.value || '';
    });
    // 파일을 ArrayBuffer로 읽어서 main 프로세스에 전달
    const reader = new FileReader();
    reader.onload = function(e) {
      const arrayBuffer = e.target.result;
      ipcRenderer.invoke('generate-docx', {
        templateBuffer: Buffer.from(arrayBuffer),
        data
      }).then((outputPath) => {
        resultDiv.innerHTML = `<a href="${outputPath}" download>생성된 파일 다운로드</a>`;
        // docx 파일을 HTML로 변환 요청
        ipcRenderer.invoke('convert-docx-to-html', outputPath).then(html => {
          previewFrame.srcdoc = html;
        }).catch(err => {
          resultDiv.textContent = '미리보기 변환 실패: ' + err;
        });
      }).catch(err => {
        resultDiv.textContent = '문서 생성 실패: ' + err;
      });
    };
    reader.readAsArrayBuffer(file);
  });

  printBtn.addEventListener('click', () => {
    previewFrame.contentWindow.print();
  });
});
