const { ipcRenderer } = require('electron');

document.addEventListener('DOMContentLoaded', () => {
  const templateFileInput = document.getElementById('templateFile');
  const inputName1 = document.getElementById('inputName1');
  const inputName2 = document.getElementById('inputName2');
  const generateBtn = document.getElementById('generateBtn');
  const resultDiv = document.getElementById('result');
  const previewFrame = document.getElementById('docxPreview');
  const printBtn = document.getElementById('printBtn');

  generateBtn.addEventListener('click', () => {
    const file = templateFileInput.files[0];
    if (!file) {
      resultDiv.textContent = '템플릿 파일을 업로드하세요.';
      return;
    }
    const name1 = inputName1.value;
    const name2 = inputName2.value;
    if (!name1 || !name2) {
      resultDiv.textContent = 'name1, name2 값을 모두 입력하세요.';
      return;
    }
    // 파일을 ArrayBuffer로 읽어서 main 프로세스에 전달
    const reader = new FileReader();
    reader.onload = function(e) {
      const arrayBuffer = e.target.result;
      ipcRenderer.invoke('generate-docx', {
        templateBuffer: Buffer.from(arrayBuffer),
        data: { name1, name2 }
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
