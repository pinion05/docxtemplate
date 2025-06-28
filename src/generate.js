const fs = require('fs');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

// 템플릿 파일 경로
const TEMPLATE_PATH = 'template.docx';

// 반복적으로 생성할 데이터 예시
const dataList = [
  { name1: '홍길동', name2: '서울' },
  { name1: '김철수', name2: '부산' },
  { name1: '이영희', name2: '대구' },
];

// 템플릿 파일 읽기
const content = fs.readFileSync(TEMPLATE_PATH, 'binary');

// 데이터별로 반복 생성
for (let i = 0; i < dataList.length; i++) {
  const zip = new PizZip(content);
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
  });
  doc.setData(dataList[i]);

  try {
    doc.render();
    const buf = doc.getZip().generate({ type: 'nodebuffer' });
    fs.writeFileSync(`output_${i + 1}.docx`, buf);
    console.log(`output_${i + 1}.docx 파일 생성 완료`);
  } catch (error) {
    console.error('문서 생성 오류:', error);
  }
}
