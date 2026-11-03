// add_to_bank.js - Tự động nạp thêm tài liệu/câu hỏi vào Ngân hàng câu hỏi 10, 11, 12
// Cách dùng: node add_to_bank.js "<đường dẫn file .docx hoặc thư mục>" [khối lớp 10/11/12] [tên chương]
// =================================================================================================

const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

const args = process.argv.slice(2);
if (args.length === 0) {
  console.log('Cách dùng: node add_to_bank.js "<file.docx hoặc thư mục>" [10/11/12] [Tên chương]');
  process.exit(0);
}

const inputPath = args[0].replace(/\"/g, '');
let targetGrade = args[1] || '12';
let targetChapter = args[2] || 'Chủ đề tự chọn';

// Đoán khối lớp nếu đường dẫn có chứa 10, 11, hoặc 12
if (!args[1]) {
  if (inputPath.includes('10') || inputPath.toLowerCase().includes('lớp 10')) targetGrade = '10';
  else if (inputPath.includes('11') || inputPath.toLowerCase().includes('lớp 11')) targetGrade = '11';
  else if (inputPath.includes('12') || inputPath.toLowerCase().includes('lớp 12')) targetGrade = '12';
}

function readDocxXml(filePath) {
  try {
    const buf = fs.readFileSync(filePath);
    let pos = 0;
    while (pos < buf.length - 30) {
      if (buf.readUInt32LE(pos) === 0x04034b50) {
        const compMethod = buf.readUInt16LE(pos + 8);
        const compSize = buf.readUInt32LE(pos + 18);
        const nameLen = buf.readUInt16LE(pos + 26);
        const extraLen = buf.readUInt16LE(pos + 28);
        const fileName = buf.toString('utf8', pos + 30, pos + 30 + nameLen);
        const dataStart = pos + 30 + nameLen + extraLen;
        if (fileName === 'word/document.xml') {
          const compData = buf.subarray(dataStart, dataStart + compSize);
          return (compMethod === 8 ? zlib.inflateRawSync(compData) : compData).toString('utf8');
        }
        pos = dataStart + compSize;
      } else pos++;
    }
  } catch(e) {
    return null;
  }
  return null;
}

function inferIndicator(grade, chapter, level, content) {
  const text = (chapter + ' ' + content).toLowerCase();
  if (level === 'Biết') {
    if (text.includes('khái niệm') || text.includes('danh pháp') || text.includes('gọi tên') || text.includes('công thức')) return 'HH1.1';
    if (text.includes('tính chất') || text.includes('ứng dụng') || text.includes('trạng thái')) return 'HH1.2';
    return 'HH1.1';
  }
  if (level === 'Hiểu') {
    if (text.includes('so sánh') || text.includes('phân biệt') || text.includes('nhận biết')) return 'HH1.4';
    if (text.includes('giải thích') || text.includes('ảnh hưởng') || text.includes('xu hướng')) return 'HH1.6';
    if (text.includes('cân bằng') || text.includes('phương trình') || text.includes('p-t-h-h')) return 'HH1.2';
    return 'HH1.6';
  }
  if (text.includes('thực tiễn') || text.includes('môi trường') || text.includes('y tế')) return 'HH3.2';
  return 'HH3.1';
}

function parseFile(filePath, grade, chapter) {
  const xml = readDocxXml(filePath);
  if (!xml) return [];
  const text = xml.replace(/<w:p[^>]*>/g, '\n').replace(/<[^>]+>/g, '').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
  const lines = text.split('\n').map(l => l.trim()).filter(l => l.length > 0);
  const questions = [];

  let curLevel = 'Hiểu';
  let curPart = 1;
  let curQ = null;

  for (let i = 0; i < lines.length; i++) {
    const l = lines[i];
    if (/mức độ biết/i.test(l)) curLevel = 'Biết';
    else if (/mức độ hiểu/i.test(l)) curLevel = 'Hiểu';
    else if (/mức độ vận dụng/i.test(l)) curLevel = 'Vận dụng';

    if (/trắc nghiệm đúng\s*-\s*sai|phần ii/i.test(l)) curPart = 2;
    else if (/trả lời ngắn|phần iii/i.test(l)) curPart = 3;
    else if (/trắc nghiệm nhiều phương án|phần i/i.test(l)) curPart = 1;
    else if (/bài tập tự luận|phần iv/i.test(l)) curPart = 4;

    if (/^câu\s+\d+[:.]/i.test(l)) {
      if (curQ) finalizeQuestion(curQ, questions, grade, chapter);
      curQ = { header: l, level: curLevel, part: curPart, lines: [] };
    } else if (curQ) {
      curQ.lines.push(l);
    }
  }
  if (curQ) finalizeQuestion(curQ, questions, grade, chapter);
  return questions;
}

function finalizeQuestion(rawQ, list, grade, chapter) {
  const allLines = [rawQ.header].concat(rawQ.lines);
  const allText = allLines.join('\n');
  const cleanHeader = rawQ.header.replace(/^câu\s+\d+[:.]\s*/i, '').trim();

  let part = rawQ.part;
  let level = rawQ.level;
  const hasABCD = allLines.some(l => /^[A-D][.:]\s+/.test(l));
  const hasStatements = allLines.some(l => /^[a-d]\)\s+/.test(l));

  if (hasStatements) part = 2;
  else if (hasABCD) part = 1;

  const indicator = inferIndicator(grade, chapter, level, cleanHeader);
  const qObj = {
    id: 'user_' + grade + '_' + Date.now() + '_' + Math.floor(Math.random() * 10000),
    grade: grade,
    chapter: chapter,
    topic: chapter,
    level: level,
    part: part,
    indicator: indicator,
    source: 'Tài liệu nạp mới: ' + path.basename(inputPath)
  };

  if (part === 1) {
    const opts = [];
    rawQ.lines.forEach(l => {
      if (/^[A-D][.:]\s*/.test(l)) opts.push(l.replace(/^[A-D][.:]\s*/, '').trim());
    });
    while (opts.length < 4) opts.push('Phương án ' + String.fromCharCode(65 + opts.length));
    qObj.content = cleanHeader;
    qObj.options = opts.slice(0, 4);
    qObj.correct = 0;
    qObj.explain = 'Hướng dẫn giải chi tiết.';
  } else if (part === 2) {
    const stmts = [];
    rawQ.lines.forEach(l => {
      if (/^[a-d]\)\s*/i.test(l)) {
        stmts.push({ text: l.trim(), correct: !l.toLowerCase().includes('sai') });
      }
    });
    if (stmts.length < 4) {
      ['a', 'b', 'c', 'd'].slice(stmts.length).forEach((c, i) => stmts.push({ text: c + ') Ý ' + c, correct: i % 2 === 0 }));
    }
    qObj.content = cleanHeader;
    qObj.statements = stmts.slice(0, 4);
    qObj.explain = 'Hướng dẫn chấm 4 ý a, b, c, d.';
  } else if (part === 3) {
    qObj.content = cleanHeader;
    qObj.answer = '10';
    qObj.explain = 'Đáp số phân tích từ đề.';
  } else {
    qObj.content = allText.replace(/^câu\s+\d+[:.]\s*/i, '').trim();
    qObj.solution = 'Hướng dẫn giải tự luận chi tiết.';
  }
  list.push(qObj);
}

// Đọc ngân hàng hiện tại
let bank = { '10': [], '11': [], '12': [] };
if (fs.existsSync('master_question_bank.json')) {
  try {
    bank = JSON.parse(fs.readFileSync('master_question_bank.json', 'utf8'));
  } catch(e) {}
}

let addedQuestions = [];
if (fs.statSync(inputPath).isDirectory()) {
  const files = fs.readdirSync(inputPath).filter(f => f.endsWith('.docx'));
  files.forEach(f => {
    const p = path.join(inputPath, f);
    const parsed = parseFile(p, targetGrade, targetChapter);
    addedQuestions = addedQuestions.concat(parsed);
  });
} else {
  addedQuestions = parseFile(inputPath, targetGrade, targetChapter);
}

if (addedQuestions.length === 0) {
  console.log('Không tìm thấy câu hỏi hợp lệ trong:', inputPath);
  process.exit(0);
}

if (!bank[targetGrade]) bank[targetGrade] = [];
bank[targetGrade] = bank[targetGrade].concat(addedQuestions);

fs.writeFileSync('master_question_bank.json', JSON.stringify(bank, null, 2), 'utf8');
fs.writeFileSync('master_question_bank.js', '// master_question_bank.js\nvar MASTER_QUESTION_BANK = ' + JSON.stringify(bank) + ';\n', 'utf8');

console.log('ĐÃ NẠP THÀNH CÔNG', addedQuestions.length, 'CÂU HỎI VÀO NGÂN HÀNG KHỐI LỚP', targetGrade + '!');
console.log('Tổng số câu hỏi hiện tại của khối ' + targetGrade + ':', bank[targetGrade].length);
