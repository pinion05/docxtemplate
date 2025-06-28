// 핫리로드 적용
try {
  require('electron-reload')(__dirname, {
    electron: require(`${__dirname}/../node_modules/electron`)
  });
} catch (e) {
  console.log('electron-reload를 불러올 수 없습니다:', e);
}


const { app, BrowserWindow, ipcMain } = require('electron');
const fs = require('fs');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const path = require('node:path');

// Windows에서 설치/제거 시 바로가기를 생성/제거하는 작업 처리
if (require('electron-squirrel-startup')) {
  app.quit();
}

const createWindow = () => {
  // 브라우저 윈도우 생성
  const mainWindow = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      nodeIntegration: true,
      contextIsolation: false,
    },
  });

  // 앱의 index.html 파일을 로드
  mainWindow.loadFile(path.join(__dirname, 'index.html'));

  // 개발자 도구 열기
  mainWindow.webContents.openDevTools();
};

// 이 메서드는 Electron이 초기화를 마치고 브라우저 윈도우를
// 생성할 준비가 되었을 때 호출됩니다.
// 일부 API는 이 이벤트가 발생한 이후에만 사용할 수 있습니다.
const mammoth = require('mammoth');

app.whenReady().then(() => {
  createWindow();

  ipcMain.handle('generate-docx', async (event, { templateBuffer, data }) => {
    try {
      const zip = new PizZip(templateBuffer);
      const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
      doc.setData(data);
      doc.render();
      const buf = doc.getZip().generate({ type: 'nodebuffer' });
      // 데스크탑에 저장
      const outputPath = path.join(app.getPath('desktop'), `output_${Date.now()}.docx`);
      fs.writeFileSync(outputPath, buf);
      return outputPath;
    } catch (e) {
      throw e.message || e;
    }
  });

  ipcMain.handle('extract-placeholders', async (event, buffer) => {
    try {
      const zip = new PizZip(buffer);
      const names = new Set();
      Object.keys(zip.files).forEach((relPath) => {
        if (/word\/.*\.xml$/.test(relPath)) {
          const xml = zip.files[relPath].asText();
          // XML 태그 제거 후 순수 텍스트만 추출 (runs 분할 문제 해결)
          const text = xml.replace(/<[^>]+>/g, '');
          const regex = /\{[#/\\^]?([^{}]+?)\}/g;
          let m;
          while ((m = regex.exec(text)) !== null) {
            names.add(m[1]);
          }
        }
      });
      return Array.from(names);
    } catch (e) {
      throw e.message || e;
    }
  });

  ipcMain.handle('convert-docx-to-html', async (event, docxPath) => {
    try {
      const result = await mammoth.convertToHtml({ path: docxPath });
      return result.value;
    } catch (e) {
      throw e.message || e;
    }
  });

  // macOS에서는 독 아이콘을 클릭하고 열려 있는 윈도우가 없을 때
  // 윈도우를 다시 생성하는 것이 일반적입니다.
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

// 모든 윈도우가 닫히면 종료합니다. 단, macOS에서는 예외입니다.
// macOS에서는 사용자가 Cmd + Q로 명시적으로 종료할 때까지
// 애플리케이션과 메뉴 바가 활성 상태로 남아 있는 것이 일반적입니다.
app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

// 이 파일에는 앱의 메인 프로세스에 필요한 나머지 코드를 포함할 수 있습니다.
// 또한 별도의 파일로 분리하여 이곳에서 import할 수도 있습니다.
