// build_master_question_bank.js - Trích xuất & Phân loại Ngân hàng Câu hỏi Hóa học 10, 11, 12 (Thầy Trần Thanh Bình)
// Phân loại: Dạng thức (P.I, P.II, P.III, P.IV) + Mức độ (Biết, Hiểu, Vận dụng) + Mã chỉ báo năng lực (HH1.1 - HH3.3)
// ===================================================================================================================

const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

// Đọc word/document.xml từ file .docx không cần thư viện bên ngoài
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
      } else {
        pos++;
      }
    }
  } catch (e) {
    console.error('Lỗi đọc docx:', filePath, e.message);
  }
  return null;
}

// Danh mục 21 chuyên đề chuẩn của 3 khối lớp
const CHAPTERS_CONFIG = [
  // LỚP 10
  { g: '10', p: 'd:/01. CONG VIEC GIANG DAY/1.3. TAI NGUYEN THAM KHAO/01. Tai lieu day them GV khac/HOA 10/TRAN THANH BINH/HOA 10 CA NAM - DEN 2027. TTB11/2026 - 2027. 10.1 Cau tao nguyen tu. TTB - DA chi tiet.docx', chap: 'Cấu tạo nguyên tử' },
  { g: '10', p: 'd:/01. CONG VIEC GIANG DAY/1.3. TAI NGUYEN THAM KHAO/01. Tai lieu day them GV khac/HOA 10/TRAN THANH BINH/HOA 10 CA NAM - DEN 2027. TTB11/2026 - 2027. 10.2 Bang tuan hoan cac nguyen to Hoa hoc. TTB - DA chi tiet.docx', chap: 'Bảng tuần hoàn các nguyên tố hóa học' },
  { g: '10', p: 'd:/01. CONG VIEC GIANG DAY/1.3. TAI NGUYEN THAM KHAO/01. Tai lieu day them GV khac/HOA 10/TRAN THANH BINH/HOA 10 CA NAM - DEN 2027. TTB11/2026 - 2027. 10.3 Lien ket hoa hoc. TTB - DA chi tiet.docx', chap: 'Liên kết hóa học' },
  { g: '10', p: 'd:/01. CONG VIEC GIANG DAY/1.3. TAI NGUYEN THAM KHAO/01. Tai lieu day them GV khac/HOA 10/TRAN THANH BINH/HOA 10 CA NAM - DEN 2027. TTB11/2026 - 2027. 10.4 Phan ung oxi hoa - khu. TTB - Dap an chi tiet.docx', chap: 'Phản ứng oxi hóa – khử' },
  { g: '10', p: 'd:/01. CONG VIEC GIANG DAY/1.3. TAI NGUYEN THAM KHAO/01. Tai lieu day them GV khac/HOA 10/TRAN THANH BINH/HOA 10 CA NAM - DEN 2027. TTB11/2026 - 2027. 10.5 Nang luong hoa hoc. TTB - DA chi tiet.docx', chap: 'Năng lượng hóa học' },
  { g: '10', p: 'd:/01. CONG VIEC GIANG DAY/1.3. TAI NGUYEN THAM KHAO/01. Tai lieu day them GV khac/HOA 10/TRAN THANH BINH/HOA 10 CA NAM - DEN 2027. TTB11/2026 - 2027. 10.6 Toc do phan ung. TTB - DA chi tiet.docx', chap: 'Tốc độ phản ứng hóa học' },
  { g: '10', p: 'd:/01. CONG VIEC GIANG DAY/1.3. TAI NGUYEN THAM KHAO/01. Tai lieu day them GV khac/HOA 10/TRAN THANH BINH/HOA 10 CA NAM - DEN 2027. TTB11/2026 - 2027. 10.7 Nguyen to nhom halogen. TTB - DA chi tiet.docx', chap: 'Nguyên tố nhóm VIIA (Halogen)' },

  // LỚP 11
  { g: '11', p: 'd:/01. CONG VIEC GIANG DAY/1.3. TAI NGUYEN THAM KHAO/01. Tai lieu day them GV khac/HOA 11/TRAN THANH BINH/2026-2027/2026 - 2027. 11.1 Can bang hoa hoc. TTB  - Dap an.docx', chap: 'Cân bằng hóa học' },
  { g: '11', p: 'd:/01. CONG VIEC GIANG DAY/1.3. TAI NGUYEN THAM KHAO/01. Tai lieu day them GV khac/HOA 11/TRAN THANH BINH/2026-2027/2026 - 2027. 11.2 Nitrogen - Sulfur. TTB - Dap an.docx', chap: 'Nitrogen và Sulfur' },
  { g: '11', p: 'd:/01. CONG VIEC GIANG DAY/1.3. TAI NGUYEN THAM KHAO/01. Tai lieu day them GV khac/HOA 11/TRAN THANH BINH/2026-2027/2026 - 2027. 11.3 Dai cuong ve hoa hoc huu co. TTB  - Dap an.docx', chap: 'Đại cương về hóa học hữu cơ' },
  { g: '11', p: 'd:/01. CONG VIEC GIANG DAY/1.3. TAI NGUYEN THAM KHAO/01. Tai lieu day them GV khac/HOA 11/TRAN THANH BINH/2026-2027/2026 - 2027. 11.4 Hydrocarbon. TTB  - Dap an.docx', chap: 'Hydrocarbon (Alkane, Alkene, Alkyne, Arene)' },
  { g: '11', p: 'd:/01. CONG VIEC GIANG DAY/1.3. TAI NGUYEN THAM KHAO/01. Tai lieu day them GV khac/HOA 11/TRAN THANH BINH/2026-2027/2026 - 2027. 11.5 Dan xuat halogen - alcohol - phenol. TTB  - Dap an.docx', chap: 'Dẫn xuất halogen – Alcohol – Phenol' },
  { g: '11', p: 'd:/01. CONG VIEC GIANG DAY/1.3. TAI NGUYEN THAM KHAO/01. Tai lieu day them GV khac/HOA 11/TRAN THANH BINH/2026-2027/2026 - 2027. 11.6 Hop chat carbonyl - carboxylic acid. TTB  - Dap an.docx', chap: 'Hợp chất carbonyl – Carboxylic acid' },

  // LỚP 12
  { g: '12', p: 'd:/01. CONG VIEC GIANG DAY/1.3. TAI NGUYEN THAM KHAO/01. Tai lieu day them GV khac/HOA 12/TRAN THANH BINH/HOÁ 12 MỚI SGK - TTB1/2026 - 2027. 12.1 Ester - Lipid. TTB  - Dap an.docx', chap: 'Ester – Lipid' },
  { g: '12', p: 'd:/01. CONG VIEC GIANG DAY/1.3. TAI NGUYEN THAM KHAO/01. Tai lieu day them GV khac/HOA 12/TRAN THANH BINH/HOÁ 12 MỚI SGK - TTB1/2026 - 2027. 12.2 Carbohydrate. TTB  - Dap an.docx', chap: 'Carbohydrate' },
  { g: '12', p: 'd:/01. CONG VIEC GIANG DAY/1.3. TAI NGUYEN THAM KHAO/01. Tai lieu day them GV khac/HOA 12/TRAN THANH BINH/HOÁ 12 MỚI SGK - TTB1/2026 - 2027. 12.3 Hop chat chua nitrogen. TTB  - Dap an.docx', chap: 'Hợp chất chứa nitrogen (Amine, Amino acid, Protein)' },
  { g: '12', p: 'd:/01. CONG VIEC GIANG DAY/1.3. TAI NGUYEN THAM KHAO/01. Tai lieu day them GV khac/HOA 12/TRAN THANH BINH/HOÁ 12 MỚI SGK - TTB1/2026 - 2027. 12.4 Polymer. TTB  - Dap an.docx', chap: 'Polymer và vật liệu polymer' },
  { g: '12', p: 'd:/01. CONG VIEC GIANG DAY/1.3. TAI NGUYEN THAM KHAO/01. Tai lieu day them GV khac/HOA 12/TRAN THANH BINH/HOÁ 12 MỚI SGK - TTB1/2026 - 2027. 12.5 Pin dien va dien phan. TTB  - Dap an.docx', chap: 'Pin điện và điện phân' },
  { g: '12', p: 'd:/01. CONG VIEC GIANG DAY/1.3. TAI NGUYEN THAM KHAO/01. Tai lieu day them GV khac/HOA 12/TRAN THANH BINH/HOÁ 12 MỚI SGK - TTB1/2026 - 2027. 12.6 Dai cuong ve kim loai. TTB  - Dap an.docx', chap: 'Đại cương về kim loại' },
  { g: '12', p: 'd:/01. CONG VIEC GIANG DAY/1.3. TAI NGUYEN THAM KHAO/01. Tai lieu day them GV khac/HOA 12/TRAN THANH BINH/HOÁ 12 MỚI SGK - TTB1/2026 - 2027. 12.7 Nguyen to nhom IA va IIA. TTB  - Dap an.docx', chap: 'Nguyên tố nhóm IA và nhóm IIA' },
  { g: '12', p: 'd:/01. CONG VIEC GIANG DAY/1.3. TAI NGUYEN THAM KHAO/01. Tai lieu day them GV khac/HOA 12/TRAN THANH BINH/HOÁ 12 MỚI SGK - TTB1/2026 - 2027. 12.8 So luoc ve day kim loai chuyen tiep thu nhat va phuc chat. TTB  - Dap an.docx', chap: 'Kim loại chuyển tiếp dãy thứ nhất & Phức chất' }
];

// Hàm suy luận Mã chỉ báo năng lực (HH1.1 - HH3.3) dựa trên mức độ, chuyên đề và nội dung câu hỏi
function inferIndicatorCode(grade, chapter, level, content) {
  const text = (chapter + ' ' + content).toLowerCase();

  if (level === 'Biết') {
    if (text.includes('khái niệm') || text.includes('định nghĩa') || text.includes('phát biểu') || text.includes('danh pháp') || text.includes('gọi tên') || text.includes('công thức phân tử')) {
      return 'HH1.1';
    }
    if (text.includes('thành phần') || text.includes('tính chất vật lí') || text.includes('ứng dụng') || text.includes('trạng thái tự nhiên')) {
      return 'HH1.2';
    }
    if (text.includes('cấu tạo') || text.includes('cấu hình') || text.includes('orbital') || text.includes('bảng tuần hoàn') || text.includes('liên kết')) {
      return 'HH1.3';
    }
    return 'HH1.1';
  }

  if (level === 'Hiểu') {
    if (text.includes('so sánh') || text.includes('phân biệt') || text.includes('đồng phân') || text.includes('nhận biết')) {
      return 'HH1.4';
    }
    if (text.includes('phân loại') || text.includes('nguyên tố s') || text.includes('bậc của')) {
      return 'HH1.5';
    }
    if (text.includes('giải thích') || text.includes('quy luật') || text.includes('xu hướng') || text.includes('chuyển dịch') || text.includes('ảnh hưởng')) {
      return 'HH1.6';
    }
    if (text.includes('cân bằng') || text.includes('thăng bằng electron') || text.includes('số oxi hóa') || text.includes('phương trình hóa học')) {
      return 'HH1.7';
    }
    if (text.includes('thí nghiệm') || text.includes('hiện tượng') || text.includes('ống nghiệm') || text.includes('quan sát')) {
      return 'HH2.4';
    }
    return 'HH1.6';
  }

  // Mức độ Vận dụng
  if (text.includes('thực tiễn') || text.includes('đời sống') || text.includes('môi trường') || text.includes('y tế') || text.includes('nước mắm') || text.includes('xăng')) {
    return 'HH3.2';
  }
  if (text.includes('thí nghiệm') || text.includes('chuẩn độ') || text.includes('chiết') || text.includes('chưng cất')) {
    return 'HH2.4';
  }
  if (text.includes('tổng hợp') || text.includes('sản xuất') || text.includes('quy trình') || text.includes('mắt xích')) {
    return 'HH3.3';
  }
  return 'HH3.1'; // Tính toán định lượng, biến thiên enthalpy, hằng số Kc, pH, hiệu suất...
}

// Phân tích văn bản 1 file docx thành danh sách câu hỏi có cấu trúc
function parseDocxQuestions(filePath, grade, defaultChapter) {
  const xml = readDocxXml(filePath);
  if (!xml) return [];

  const text = xml.replace(/<w:p[^>]*>/g, '\n')
                  .replace(/<[^>]+>/g, '')
                  .replace(/&lt;/g, '<')
                  .replace(/&gt;/g, '>')
                  .replace(/&amp;/g, '&')
                  .replace(/&quot;/g, '"');

  const lines = text.split('\n').map(l => l.trim()).filter(l => l.length > 0);
  const questions = [];

  let curLevel = 'Biết';
  let curPart = 1;
  let curSubTopic = defaultChapter;
  let curQ = null;

  for (let i = 0; i < lines.length; i++) {
    const l = lines[i];

    // Nhận diện mức độ
    if (/mức độ biết/i.test(l)) { curLevel = 'Biết'; }
    else if (/mức độ hiểu/i.test(l)) { curLevel = 'Hiểu'; }
    else if (/mức độ vận dụng/i.test(l)) { curLevel = 'Vận dụng'; }

    // Nhận diện phần câu hỏi
    if (/trắc nghiệm đúng\s*-\s*sai/i.test(l) || /phần ii/i.test(l)) { curPart = 2; }
    else if (/trả lời ngắn/i.test(l) || /phần iii/i.test(l)) { curPart = 3; }
    else if (/trắc nghiệm nhiều phương án/i.test(l) || /phần i/i.test(l)) { curPart = 1; }
    else if (/bài tập tự luận/i.test(l) || /phần iv/i.test(l)) { curPart = 4; }

    // Nhận diện tiêu đề chuyên đề nhỏ / dạng bài
    if (/^(cđ\d|dạng \d|phần \d:)/i.test(l) && l.length < 80) {
      curSubTopic = l.replace(/^(cđ\d[:.]?|dạng \d[:.]?|phần \d:?)\s*/i, '').trim() || defaultChapter;
    }

    // Nhận diện bắt đầu một câu hỏi
    const isQStart = /^câu\s+\d+[:.]/i.test(l);
    if (isQStart) {
      if (curQ) {
        finalizeQuestionItem(curQ, questions, grade, defaultChapter);
      }
      curQ = {
        header: l,
        level: curLevel,
        part: curPart,
        subTopic: curSubTopic,
        lines: []
      };
    } else if (curQ) {
      curQ.lines.push(l);
    }
  }

  if (curQ) {
    finalizeQuestionItem(curQ, questions, grade, defaultChapter);
  }

  return questions;
}

// Chuẩn hóa câu hỏi hoàn chỉnh
function finalizeQuestionItem(rawQ, list, grade, chapter) {
  const allLines = [rawQ.header].concat(rawQ.lines);
  const allText = allLines.join('\n');

  // Làm sạch tiêu đề câu
  const cleanHeader = rawQ.header.replace(/^câu\s+\d+[:.]\s*/i, '').trim();
  let part = rawQ.part;
  let level = rawQ.level;

  // Tự động kiểm tra dạng thức thực tế của câu hỏi nếu văn bản có dấu hiệu đặc thù
  const hasABCD = allLines.some(l => /^[A-D][.:]\s+/.test(l));
  const hasABCDStatements = allLines.some(l => /^[a-d]\)\s+/.test(l));
  const hasDapSo = /đáp số|đ\/s|kết quả[:=]/i.test(allText);

  if (hasABCDStatements && (allText.includes('đúng') || allText.includes('sai') || rawQ.part === 2)) {
    part = 2;
  } else if (hasABCD && rawQ.part !== 2) {
    part = 1;
  } else if (hasDapSo && rawQ.part !== 1 && rawQ.part !== 2) {
    part = 3;
  }

  const topicName = chapter + (rawQ.subTopic && rawQ.subTopic !== chapter ? ' (' + rawQ.subTopic + ')' : '');
  const indicator = inferIndicatorCode(grade, chapter, level, cleanHeader + ' ' + allText.slice(0, 200));

  const qObj = {
    id: 'ttb_' + grade + '_' + (list.length + 1) + '_' + Math.random().toString(36).substring(2, 7),
    grade: grade,
    chapter: chapter,
    topic: topicName,
    level: level,
    part: part,
    indicator: indicator,
    source: 'Thầy Trần Thanh Bình (2026 - 2027)'
  };

  if (part === 1) {
    // Trắc nghiệm 4 lựa chọn
    const options = [];
    let correct = 0;
    let mainContent = cleanHeader;
    const bodyTextLines = [];
    let inSolution = false;
    let solutionLines = [];

    rawQ.lines.forEach(l => {
      if (/hướng dẫn giải|lời giải chi tiết|chọn đáp án/i.test(l)) {
        inSolution = true;
      }
      if (inSolution) {
        solutionLines.push(l);
      } else if (/^[A-D][.:]\s*/.test(l)) {
        options.push(l.replace(/^[A-D][.:]\s*/, '').trim());
      } else {
        bodyTextLines.push(l);
      }
    });

    if (bodyTextLines.length > 0 && options.length === 0) {
      mainContent += ' ' + bodyTextLines.join(' ');
    }

    // Nếu không tìm thấy dòng A. B. C. D. riêng biệt, tìm trong cùng 1 dòng
    if (options.length < 4) {
      const matchOpts = allText.match(/[A-D][.:]\s+([^\nABCD]+)/g);
      if (matchOpts && matchOpts.length >= 4) {
        options.length = 0;
        matchOpts.slice(0, 4).forEach(o => options.push(o.replace(/^[A-D][.:]\s*/, '').trim()));
      }
    }

    while (options.length < 4) {
      options.push('Phương án ' + String.fromCharCode(65 + options.length));
    }

    // Xác định đáp án đúng nếu có ghi chú
    const solText = solutionLines.join(' ');
    const chooseMatch = solText.match(/(?:chọn|đáp án|chọn phương án)\s+([A-D])/i);
    if (chooseMatch) {
      correct = chooseMatch[1].toUpperCase().charCodeAt(0) - 65;
    }

    qObj.content = mainContent.replace(/\s+/g, ' ').trim();
    qObj.options = options.slice(0, 4);
    qObj.correct = (correct >= 0 && correct <= 3) ? correct : 0;
    qObj.explain = solText ? solText.trim() : 'Lời giải chi tiết theo tài liệu giảng dạy.';
    list.push(qObj);

  } else if (part === 2) {
    // Trắc nghiệm Đúng / Sai (4 ý a, b, c, d)
    const statements = [];
    let inSolution = false;
    let solutionLines = [];

    rawQ.lines.forEach(l => {
      if (/hướng dẫn giải|lời giải chi tiết/i.test(l)) inSolution = true;
      if (inSolution) {
        solutionLines.push(l);
      } else if (/^[a-d]\)\s*/i.test(l)) {
        const textStmt = l.trim();
        const isTrue = textStmt.toLowerCase().includes('(đ)') || textStmt.toLowerCase().includes('đúng') || !textStmt.toLowerCase().includes('(s)');
        statements.push({ text: textStmt, correct: isTrue });
      }
    });

    if (statements.length < 4) {
      ['a', 'b', 'c', 'd'].slice(statements.length).forEach((char, idx) => {
        statements.push({ text: char + ') Nhận định ' + char, correct: idx % 2 === 0 });
      });
    }

    qObj.content = cleanHeader.replace(/\s+/g, ' ').trim();
    qObj.statements = statements.slice(0, 4);
    qObj.explain = solutionLines.join(' ').trim() || 'Hướng dẫn chấm chi tiết 4 ý a, b, c, d.';
    list.push(qObj);

  } else if (part === 3) {
    // Trắc nghiệm trả lời ngắn
    let ans = '10';
    let solText = '';
    const matchAns = allText.match(/(?:đáp số|đ\/s|kết quả)[:=\s]+([0-9.,+-]+)/i);
    if (matchAns) {
      ans = matchAns[1].trim();
    }
    const idxSol = allText.search(/hướng dẫn giải|lời giải/i);
    if (idxSol !== -1) {
      solText = allText.slice(idxSol).trim();
    }

    qObj.content = cleanHeader.replace(/\s+/g, ' ').trim();
    qObj.answer = ans;
    qObj.explain = solText || ('Đáp số: ' + ans);
    list.push(qObj);

  } else {
    // Tự luận (Part 4)
    let solText = 'Hướng dẫn giải chi tiết.';
    const idxSol = allText.search(/hướng dẫn giải|lời giải chi tiết/i);
    let qContent = cleanHeader;
    if (idxSol !== -1) {
      solText = allText.slice(idxSol).trim();
      qContent = allText.slice(0, idxSol).replace(/^câu\s+\d+[:.]\s*/i, '').trim();
    }

    qObj.content = qContent.replace(/\s+/g, ' ').trim();
    qObj.solution = solText;
    list.push(qObj);
  }
}

// Tiến hành trích xuất toàn bộ 21 chuyên đề
console.log('=== BẮT ĐẦU TRÍCH XUẤT VÀ PHÂN LOẠI NGÂN HÀNG CÂU HỎI THẦY TRẦN THANH BÌNH ===');
const MASTER_BANK = {
  '10': [],
  '11': [],
  '12': []
};

CHAPTERS_CONFIG.forEach((item, idx) => {
  if (fs.existsSync(item.p)) {
    console.log(`[${idx + 1}/21] Đang xử lý: Lớp ${item.g} - ${item.chap}...`);
    const qList = parseDocxQuestions(item.p, item.g, item.chap);
    console.log(`  -> Trích xuất thành công: ${qList.length} câu hỏi`);
    
    // Chọn lọc cân bằng theo từng Part & Mức độ để ngân hàng hoạt động mượt mà
    // Mỗi chuyên đề lấy ~60-80 câu bao quát đầy đủ 4 phần (Part I, II, III, IV) và 3 mức độ (Biết, Hiểu, Vận dụng)
    const p1_b = qList.filter(q => q.part === 1 && q.level === 'Biết').slice(0, 15);
    const p1_h = qList.filter(q => q.part === 1 && q.level === 'Hiểu').slice(0, 15);
    const p1_v = qList.filter(q => q.part === 1 && q.level === 'Vận dụng').slice(0, 15);
    const p2_all = qList.filter(q => q.part === 2).slice(0, 10);
    const p3_all = qList.filter(q => q.part === 3).slice(0, 12);
    const p4_all = qList.filter(q => q.part === 4).slice(0, 10);

    const curated = [].concat(p1_b, p1_h, p1_v, p2_all, p3_all, p4_all);
    MASTER_BANK[item.g] = MASTER_BANK[item.g].concat(curated);
  } else {
    console.warn(`[!] Không tìm thấy file: ${item.p}`);
  }
});

// Thống kê tổng thể
console.log('\n=== TỔNG KẾT NGÂN HÀNG CÂU HỎI ĐÃ PHÂN LOẠI ===');
['10', '11', '12'].forEach(g => {
  const total = MASTER_BANK[g].length;
  const p1 = MASTER_BANK[g].filter(q => q.part === 1).length;
  const p2 = MASTER_BANK[g].filter(q => q.part === 2).length;
  const p3 = MASTER_BANK[g].filter(q => q.part === 3).length;
  const p4 = MASTER_BANK[g].filter(q => q.part === 4).length;
  const b = MASTER_BANK[g].filter(q => q.level === 'Biết').length;
  const h = MASTER_BANK[g].filter(q => q.level === 'Hiểu').length;
  const v = MASTER_BANK[g].filter(q => q.level === 'Vận dụng').length;
  console.log(`Lớp ${g}: Tổng ${total} câu (P.I: ${p1}, P.II: ${p2}, P.III: ${p3}, P.IV: ${p4}) - (Biết: ${b}, Hiểu: ${h}, Vận dụng: ${v})`);
});

// Lưu thành file JSON và file JS tải trực tiếp trên trình duyệt
const jsonOut = JSON.stringify(MASTER_BANK, null, 2);
fs.writeFileSync('master_question_bank.json', jsonOut, 'utf8');

const jsOut = '// master_question_bank.js - Ngân Hàng Câu Hỏi Phân Loại Chuẩn Hóa 3 Khối Lớp 10, 11, 12\n' +
  '// Phân loại: Dạng thức (P.I, P.II, P.III, P.IV) + Mức độ (Biết, Hiểu, Vận dụng) + Mã chỉ báo năng lực (HH1.1 - HH3.3)\n' +
  'var MASTER_QUESTION_BANK = ' + JSON.stringify(MASTER_BANK) + ';\n';
fs.writeFileSync('master_question_bank.js', jsOut, 'utf8');

console.log('\nĐã lưu thành công:');
console.log('- master_question_bank.json (' + (fs.statSync('master_question_bank.json').size / 1024 / 1024).toFixed(2) + ' MB)');
console.log('- master_question_bank.js (' + (fs.statSync('master_question_bank.js').size / 1024 / 1024).toFixed(2) + ' MB)');
