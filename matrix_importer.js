/* =========================================================================================
   HỆ THỐNG SỐ HÓA MA TRẬN & TẠO ĐỀ THI HÓA HỌC GDPT 2018 (CV 7991)
   MODULE: NHẬP MA TRẬN & BẢN ĐẶC TẢ TỪ MÁY TÍNH (WORD / EXCEL / ẢNH / VĂN BẢN)
   KÈM TÍNH NĂNG: RA ĐỀ VÀ CHỈ LẤY CÂU HỎI TỪ ĐỀ CƯƠNG ÔN TẬP TRÊN TRƯỜNG
   Tác giả: GV. NGUYỄN VỸ - Trường THPT Bùi Thị Xuân (Mũi Né)
   ========================================================================================= */

(function(window) {
    'use strict';

    // State lưu trữ dữ liệu ma trận và đề cương trường nhập khẩu
    var importedState = {
        fileName: "",
        fileType: "",
        detectedGrade: "12",
        detectedTerm: "GK1",
        detectedFormat: "MIXED",
        examTitle: "ĐỀ KIỂM TRA ĐỊNH KỲ MÔN HÓA HỌC",
        examTime: "50",
        rows: [],
        rawGrid: [],
        imagePreviewUrl: "",
        ocrText: "",
        activeTab: "file", // "file", "paste"

        // Nguồn câu hỏi: "bank" (ngân hàng chuẩn 10.094 câu) hoặc "syllabus" (chỉ lấy từ đề cương trường)
        questionSource: "bank",
        syllabusFileName: "",
        syllabusQuestions: [],
        strictSyllabusOnly: true
    };

    /* --------------------------------------------------------------------------
       1. SMART TOPIC NORMALIZATION & CANONICAL MAPPING
       -------------------------------------------------------------------------- */
    function detectGradeFromText(text) {
        if (!text) return null;
        var t = String(text).toLowerCase();
        if (/hóa\s*(học)?\s*10|khối\s*10|lớp\s*10|\b(g10|10a|10b)\b/i.test(t)) return "10";
        if (/hóa\s*(học)?\s*11|khối\s*11|lớp\s*11|\b(g11|11a|11b)\b/i.test(t)) return "11";
        if (/hóa\s*(học)?\s*12|khối\s*12|lớp\s*12|\b(g12|12a|12b)\b/i.test(t)) return "12";
        return null;
    }

    function detectGradeFromTopics(topics) {
        var score10 = 0, score11 = 0, score12 = 0;
        (topics || []).forEach(function(top) {
            var t = String(top).toLowerCase();
            // Grade 10
            if (/nguyên tử|bảng tuần hoàn|liên kết|oxi hóa|năng lượng hóa học|enthalpy|tốc độ phản ứng|halogen|clorua|bromine|iodine/i.test(t)) score10 += 3;
            // Grade 11
            if (/cân bằng|điện li|nitrogen|sulfur|ammonia|hydrocarbon|alkane|alkene|alkyne|arene|alcohol|ancol|phenol|carbonyl|carboxylic/i.test(t)) score11 += 3;
            // Grade 12
            if (/ester|este|lipid|xà phòng|carbohydrate|cacbohidrat|glucose|fructose|amine|amin|amino acid|protein|polymer|polime|pin điện|điện phân|kim loại|phức chất/i.test(t)) score12 += 3;
        });
        if (score12 >= score11 && score12 >= score10 && score12 > 0) return "12";
        if (score11 >= score10 && score11 > 0) return "11";
        if (score10 > 0) return "10";
        return (typeof currentGrade !== "undefined" && currentGrade) ? currentGrade : "12";
    }

    // Chuẩn hóa tên chủ đề thành tên chương/bài chuẩn trong hệ thống để matching câu hỏi chính xác 100%
    function mapToCanonicalTopic(rawTopic, grade) {
        if (!rawTopic) return "Chủ đề Hóa học";
        var t = rawTopic.trim();
        var g = String(grade || "12");

        if (g === "10") {
            if (/nguyên tử|hạt nhân|proton|neutron|electron|đồng vị|vỏ electron/i.test(t)) return "Chương 1: Cấu tạo nguyên tử - Bài 1, 2: Thành phần của nguyên tử";
            if (/bảng tuần hoàn|chu kì|nhóm|xu hướng biến đổi/i.test(t)) return "Chương 2: Bảng tuần hoàn các nguyên tố hóa học - Bài 6, 7: Cấu tạo & Xu hướng biến đổi";
            if (/liên kết|cộng hóa trị|ion|hydrogen|van der waals/i.test(t)) return "Chương 3: Liên kết hóa học - Bài 9, 10: Liên kết ion & Cộng hóa trị";
            if (/oxi hóa|chất khử|chất oxi hóa/i.test(t)) return "Chương 4: Phản ứng oxi hóa - khử - Bài 12: Phản ứng oxi hóa - khử";
            if (/năng lượng|enthalpy|nhiệt phản ứng|tỏa nhiệt|thu nhiệt/i.test(t)) return "Chương 5: Năng lượng hóa học - Bài 13: Enthalpy và biến thiên enthalpy";
            if (/tốc độ|hằng số tốc độ|xúc tác/i.test(t)) return "Chương 6: Tốc độ phản ứng hóa học - Bài 15: Tốc độ phản ứng hóa học";
            if (/halogen|viia|chlorine|bromine|iodine|muối halide/i.test(t)) return "Chương 7: Nhóm Halogen - Bài 18, 19: Nguyên tố nhóm VIIA (Halogen)";
        } else if (g === "11") {
            if (/cân bằng|điện li|ph\b|acid|base/i.test(t)) return "Chương 1: Cân bằng hóa học - Bài 1, 2: Cân bằng hóa học & Sự điện li";
            if (/nitrogen|sulfur|ammonia|nitric|sulfuric/i.test(t)) return "Chương 2: Nitrogen và sulfur - Bài 3, 4, 5: Nitrogen, Ammonia & Sulfur";
            if (/đại cương.*hữu cơ|hóa học hữu cơ|công thức phân tử/i.test(t)) return "Chương 3: Đại cương về hóa học hữu cơ - Bài 8: Đại cương hóa hữu cơ";
            if (/hydrocarbon|alkane|alkene|alkyne|arene/i.test(t)) return "Chương 4: Hydrocarbon - Bài 12, 13: Alkane, Alkene & Alkyne";
            if (/dẫn xuất halogen|alcohol|ancol|phenol/i.test(t)) return "Chương 5: Dẫn xuất halogen - Alcohol - Phenol - Bài 15, 16: Dẫn xuất halogen, Alcohol & Phenol";
            if (/carbonyl|aldehyde|ketone|carboxylic/i.test(t)) return "Chương 6: Hợp chất carbonyl - Carboxylic acid - Bài 18, 19: Carbonyl & Carboxylic acid";
        } else {
            // Grade 12
            if (/ester|este|lipid|xà phòng|chất béo/i.test(t)) return "Chương 1: Ester - Lipid - Bài 1: Ester – Lipid";
            if (/carbohydrate|cacbohidrat|glucose|fructose|saccharose|tinh bột|cellulose/i.test(t)) return "Chương 2: Carbohydrate - Bài 3, 4, 5: Carbohydrate (Glucose, Fructose, Tinh bột)";
            if (/amine|amin|amino acid|axit amin|peptide|protein/i.test(t)) return "Chương 3: Hợp chất chứa nitrogen - Bài 6, 7: Amine & Amino acid";
            if (/polymer|polime|vật liệu polime|chất dẻo|cao su/i.test(t)) return "Chương 4: Polymer - Bài 12: Đại cương về polymer";
            if (/pin điện|điện phân|thế điện cực/i.test(t)) return "Chương 5: Pin điện và điện phân - Bài 15, 16: Pin điện và điện phân";
            if (/đại cương.*kim loại|tính chất kim loại|ăn mòn/i.test(t)) return "Chương 6: Đại cương về kim loại - Bài 18: Đại cương về kim loại";
            if (/nhóm ia|nhóm iia|kim loại kiềm|kiềm thổ/i.test(t)) return "Chương 7: Nguyên tố nhóm IA và nhóm IIA - Bài 22: Kim loại kiềm & kiềm thổ";
            if (/phức chất|chuyển tiếp|dãy thứ nhất/i.test(t)) return "Chương 8: Phức chất - Bài 25: Phức chất";
        }

        return rawTopic;
    }

    /* --------------------------------------------------------------------------
       2. CORE TABLE GRID PARSER (UNIFIED FOR DOCX, EXCEL, TEXT)
       -------------------------------------------------------------------------- */
    function parseTableGrid(grid, gradeHint) {
        if (!grid || !Array.isArray(grid) || grid.length === 0) {
            return { rows: [], grade: gradeHint || "12" };
        }

        var cleanGrid = grid.map(function(row) {
            return (row || []).map(function(cell) {
                return (cell !== null && cell !== undefined) ? String(cell).trim() : "";
            });
        });

        var headerRowIdx = -1;
        var topicColIdx = -1;
        var sttColIdx = -1;

        // Quét 6 dòng đầu để tìm tiêu đề bảng ma trận
        for (var r = 0; r < Math.min(cleanGrid.length, 6); r++) {
            var row = cleanGrid[r];
            for (var c = 0; c < row.length; c++) {
                var cell = row[c].toLowerCase();
                if (/^(tt|stt|số tt|số thứ tự)$/i.test(cell) && sttColIdx === -1) {
                    sttColIdx = c;
                }
                if (/chủ đề|mạch kiến thức|nội dung|tên bài|bài học|chương|kiến thức/i.test(cell) && topicColIdx === -1) {
                    topicColIdx = c;
                    headerRowIdx = r;
                }
            }
            if (topicColIdx !== -1) break;
        }

        if (sttColIdx === -1 && topicColIdx === 1) sttColIdx = 0;
        if (topicColIdx === -1) {
            topicColIdx = (sttColIdx === 0) ? 1 : 0;
            headerRowIdx = 0;
        }

        var h2 = cleanGrid[headerRowIdx + 1] || [];
        var startDataRow = headerRowIdx + 1;
        var h2NumCount = h2.filter(function(cell) { return /^\d+$/.test(cell.trim()); }).length;
        if (h2.length > 0 && h2NumCount < 2 && (/nhận biết|thông hiểu|vận dụng|\bvd\b|biết|hiểu/i.test(h2.join(" ")))) {
            startDataRow = headerRowIdx + 2;
        }

        var dataRows = [];
        for (var i = startDataRow; i < cleanGrid.length; i++) {
            var rData = cleanGrid[i];
            if (!rData || rData.length === 0) continue;
            var rowStr = rData.join(" ").trim();
            if (!rowStr) continue;

            // Bỏ qua dòng tổng kết
            if (/^(tổng|tổng cộng|tỉ lệ|tỷ lệ|điểm số|tổng điểm|ghi chú|xác nhận)/i.test(rowStr)) continue;

            var topicCell = (rData[topicColIdx] || "").trim();
            if (!topicCell || /^\d+$/.test(topicCell)) {
                var foundTextCol = rData.findIndex(function(cVal, idx) {
                    return idx !== sttColIdx && cVal && !/^\d+$/.test(cVal.trim()) && cVal.trim().length > 3;
                });
                if (foundTextCol !== -1) {
                    topicCell = rData[foundTextCol].trim();
                } else {
                    continue;
                }
            }

            var cleanTopic = topicCell.replace(/^\d+[\.\/\-\s]+/, "").trim();
            if (cleanTopic.length < 2) continue;

            // Thu thập các ô số
            var numCells = [];
            for (var cIdx = 0; cIdx < rData.length; cIdx++) {
                if (cIdx === topicColIdx || cIdx === sttColIdx) continue;
                var mM = rData[cIdx].match(/\d+/);
                if (mM) {
                    numCells.push({ col: cIdx, val: parseInt(mM[0], 10), raw: rData[cIdx] });
                }
            }

            var p1_k = 0, p1_h = 0, p1_v = 0;
            var p2_k = 0, p2_h = 0, p2_v = 0;
            var p3_k = 0, p3_h = 0, p3_v = 0;
            var p4_k = 0, p4_h = 0, p4_v = 0;

            var fullRowText = rData.join(" ");
            // 1. Nếu là dạng Bản đặc tả có mô tả số câu dạng: - Phần I (MCQ): 2 câu (B: 1, H: 1)...
            if (/phần\s*(i|ii|iii|iv)|mcq|đúng\/sai|trả lời ngắn/i.test(fullRowText) && /\(\s*[bhv]/i.test(fullRowText)) {
                var parseSubLevels = function(str) {
                    var k = 0, h = 0, v = 0;
                    var bM = str.match(/b[^\d]*(\d+)/i) || str.match(/biết[^\d]*(\d+)/i); if (bM) k = parseInt(bM[1], 10);
                    var hM = str.match(/h[^\d]*(\d+)/i) || str.match(/hiểu[^\d]*(\d+)/i); if (hM) h = parseInt(hM[1], 10);
                    var vM = str.match(/v[^\d]*(\d+)/i) || str.match(/vận dụng[^\d]*(\d+)/i); if (vM) v = parseInt(vM[1], 10);
                    return { k: k, h: h, v: v };
                };

                var lines = fullRowText.split(/[\r\n]+/);
                lines.forEach(function(l) {
                    var lvlM = l.match(/\(([^)]*(?:b|biết|h|hiểu|v|vd)[^)]*)\)/i);
                    var sub = lvlM ? lvlM[1] : "";
                    if (!sub) return;
                    var res = parseSubLevels(sub);

                    if (/phần\s*i\b|mcq/i.test(l)) {
                        p1_k = res.k; p1_h = res.h; p1_v = res.v;
                    } else if (/phần\s*ii\b|đúng\/sai|đ\/s/i.test(l)) {
                        p2_k = res.k; p2_h = res.h; p2_v = res.v;
                    } else if (/phần\s*iii\b|trả lời ngắn|tln/i.test(l)) {
                        p3_k = res.k; p3_h = res.h; p3_v = res.v;
                    } else if (/phần\s*iv\b|tự luận/i.test(l)) {
                        p4_k = res.k; p4_h = res.h; p4_v = res.v;
                    }
                });
            } else if (numCells.length >= 9) {
                // 2. Chuẩn 12 cột CV 7991:
                p1_k = numCells[0].val; p1_h = numCells[1].val; p1_v = numCells[2].val;
                p2_k = numCells[3].val; p2_h = numCells[4].val; p2_v = numCells[5].val;
                p3_k = numCells[6].val; p3_h = numCells[7].val; p3_v = numCells[8].val;
                if (numCells.length >= 12) {
                    p4_k = numCells[9].val; p4_h = numCells[10].val; p4_v = numCells[11].val;
                }
            } else if (numCells.length >= 4) {
                // 3. Ma trận 4 mức độ: Nhận biết, Thông hiểu, Vận dụng, Vận dụng cao
                p1_k = numCells[0].val;
                p1_h = numCells[1].val;
                p1_v = numCells[2].val;
                p3_v = numCells[3].val;
            } else if (numCells.length === 3) {
                p1_k = numCells[0].val;
                p1_h = numCells[1].val;
                p1_v = numCells[2].val;
            } else if (numCells.length > 0) {
                p1_k = numCells[0].val;
                if (numCells[1]) p1_h = numCells[1].val;
            }

            dataRows.push({
                rawTopic: cleanTopic,
                topic: cleanTopic,
                p1_k: p1_k, p1_h: p1_h, p1_v: p1_v,
                p2_k: p2_k, p2_h: p2_h, p2_v: p2_v,
                p3_k: p3_k, p3_h: p3_h, p3_v: p3_v,
                p4_k: p4_k, p4_h: p4_h, p4_v: p4_v,
                codes: ["HH1.1", "HH1.2"]
            });
        }

        var detectedGrade = gradeHint || detectGradeFromTopics(dataRows.map(function(r) { return r.rawTopic; }));

        // Chuẩn hóa topic theo khối lớp đã nhận diện
        dataRows.forEach(function(r) {
            r.canonicalTopic = mapToCanonicalTopic(r.rawTopic, detectedGrade);
            r.topic = r.canonicalTopic;
        });

        return {
            rows: dataRows,
            grade: detectedGrade
        };
    }

    /* --------------------------------------------------------------------------
       3. DOCX (WORD) PARSER USING JSZIP + DOMPARSER
       -------------------------------------------------------------------------- */
    function parseDocxFile(arrayBuffer) {
        return new Promise(function(resolve, reject) {
            if (!window.JSZip) {
                return reject(new Error("Thư viện JSZip chưa được tải."));
            }

            JSZip.loadAsync(arrayBuffer).then(function(zip) {
                var docXmlFile = zip.file("word/document.xml");
                if (!docXmlFile) {
                    return reject(new Error("Không tìm thấy nội dung văn bản (word/document.xml) trong tệp docx."));
                }

                docXmlFile.async("string").then(function(xmlText) {
                    var parser = new DOMParser();
                    var xmlDoc = parser.parseFromString(xmlText, "application/xml");

                    // 1. Quét tiêu đề và khối lớp từ toàn bộ văn bản
                    var fullDocText = xmlDoc.body ? xmlDoc.body.textContent : "";
                    var detectedGrade = detectGradeFromText(fullDocText);

                    var examTitle = "ĐỀ KIỂM TRA ĐỊNH KỲ HÓA HỌC";
                    var titleMatch = fullDocText.match(/đề kiểm tra[^\r\n]+|ma trận[^\r\n]+kiểm tra[^\r\n]+/i);
                    if (titleMatch) {
                        examTitle = titleMatch[0].trim().replace(/\s+/g, " ").slice(0, 80);
                    }

                    // 2. Trích xuất các bảng trong văn bản Word
                    var tblElements = xmlDoc.getElementsByTagName("w:tbl");
                    if (tblElements.length === 0) {
                        return reject(new Error("Không tìm thấy bảng ma trận nào trong tệp Word đã chọn."));
                    }

                    var bestParsed = null;
                    var maxRows = 0;

                    for (var tIdx = 0; tIdx < tblElements.length; tIdx++) {
                        var tbl = tblElements[tIdx];
                        var trs = tbl.getElementsByTagName("w:tr");
                        var grid = [];

                        for (var rIdx = 0; rIdx < trs.length; rIdx++) {
                            var tr = trs[rIdx];
                            var tcs = tr.getElementsByTagName("w:tc");
                            var rowCells = [];

                            for (var cIdx = 0; cIdx < tcs.length; cIdx++) {
                                var tc = tcs[cIdx];
                                var textContent = "";
                                var textNodes = tc.getElementsByTagName("w:t");
                                for (var tnIdx = 0; tnIdx < textNodes.length; tnIdx++) {
                                    textContent += textNodes[tnIdx].textContent;
                                }

                                var gridSpanEl = tc.getElementsByTagName("w:gridSpan")[0];
                                var span = gridSpanEl ? parseInt(gridSpanEl.getAttribute("w:val") || "1", 10) : 1;

                                rowCells.push(textContent.trim());
                                for (var s = 1; s < span; s++) {
                                    rowCells.push("");
                                }
                            }
                            grid.push(rowCells);
                        }

                        var parsed = parseTableGrid(grid, detectedGrade);
                        if (parsed.rows.length > maxRows) {
                            maxRows = parsed.rows.length;
                            bestParsed = parsed;
                            bestParsed.grid = grid;
                        }
                    }

                    if (!bestParsed || bestParsed.rows.length === 0) {
                        return reject(new Error("Không thể phân tích dữ liệu các dòng trong bảng Word."));
                    }

                    resolve({
                        grade: bestParsed.grade || detectedGrade || "12",
                        title: examTitle,
                        rows: bestParsed.rows,
                        grid: bestParsed.grid
                    });
                }).catch(reject);
            }).catch(reject);
        });
    }

    /* --------------------------------------------------------------------------
       4. EXCEL (.XLSX / .XLS) PARSER USING SHEETJS
       -------------------------------------------------------------------------- */
    function parseExcelFile(arrayBuffer) {
        return new Promise(function(resolve, reject) {
            if (!window.XLSX) {
                return reject(new Error("Thư viện SheetJS (XLSX) chưa được nạp."));
            }

            try {
                var workbook = XLSX.read(arrayBuffer, { type: 'array' });
                if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
                    return reject(new Error("Tệp Excel không chứa trang tính (sheet) nào."));
                }

                var targetSheetName = workbook.SheetNames.find(function(name) {
                    return /ma\s*tr[ậa]n|b[ảa]n\s*đ[ặa]c\s*t[ảa]|đ[ặa]c\s*t[ảa]/i.test(name);
                }) || workbook.SheetNames[0];

                var sheet = workbook.Sheets[targetSheetName];
                var grid = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

                if (!grid || grid.length === 0) {
                    return reject(new Error("Trang tính " + targetSheetName + " trống rỗng."));
                }

                var allText = grid.map(function(r) { return r.join(" "); }).join("\n");
                var detectedGrade = detectGradeFromText(allText);

                var parsed = parseTableGrid(grid, detectedGrade);
                if (parsed.rows.length === 0) {
                    return reject(new Error("Không tìm thấy dòng dữ liệu chủ đề nào trong bảng tính Excel."));
                }

                resolve({
                    grade: parsed.grade || detectedGrade || "12",
                    title: "ĐỀ KIỂM TRA HÓA HỌC (TỪ FILE EXCEL)",
                    rows: parsed.rows,
                    grid: grid
                });
            } catch (err) {
                reject(err);
            }
        });
    }

    /* --------------------------------------------------------------------------
       5. IMAGE PARSER USING TESSERACT.JS OCR
       -------------------------------------------------------------------------- */
    function parseImageFile(imageFile, progressCallback) {
        return new Promise(function(resolve, reject) {
            if (!window.Tesseract) {
                return reject(new Error("Thư viện nhận dạng ảnh Tesseract.js chưa sẵn sàng."));
            }

            var reader = new FileReader();
            reader.onload = function(e) {
                var dataUrl = e.target.result;
                importedState.imagePreviewUrl = dataUrl;

                if (typeof progressCallback === "function") {
                    progressCallback(10, "Đang khởi tạo bộ máy OCR tiếng Việt...");
                }

                Tesseract.recognize(dataUrl, 'vie+eng', {
                    logger: function(m) {
                        if (m.status === 'recognizing text' && typeof progressCallback === "function") {
                            var pct = Math.round((m.progress || 0) * 80) + 10;
                            progressCallback(pct, "Đang nhận diện ký tự & cấu trúc bảng (" + pct + "%)...");
                        }
                    }
                }).then(function(result) {
                    var text = (result && result.data && result.data.text) ? result.data.text : "";
                    importedState.ocrText = text;

                    if (typeof progressCallback === "function") {
                        progressCallback(95, "Đang trích xuất các dòng ma trận...");
                    }

                    var lines = text.split(/[\r\n]+/).map(function(l) { return l.trim(); }).filter(Boolean);
                    var grid = [];

                    lines.forEach(function(line) {
                        var parts = line.split(/[\|\t]+/).map(function(p) { return p.trim(); }).filter(Boolean);
                        if (parts.length > 1) {
                            grid.push(parts);
                        } else {
                            var numMatches = line.match(/\b\d+\b/g);
                            var words = line.replace(/\b\d+\b/g, "").replace(/[\:\-\–\.\/]+/g, " ").trim();
                            if (words.length > 2 && numMatches && numMatches.length >= 2) {
                                var row = [words].concat(numMatches);
                                grid.push(row);
                            }
                        }
                    });

                    var detectedGrade = detectGradeFromText(text);
                    var parsed = parseTableGrid(grid, detectedGrade);

                    if (parsed.rows.length === 0) {
                        var detectedTopics = [];
                        var lower = text.toLowerCase();
                        if (/nguyên tử/i.test(lower)) detectedTopics.push("Chương 1: Cấu tạo nguyên tử");
                        if (/tuần hoàn/i.test(lower)) detectedTopics.push("Chương 2: Bảng tuần hoàn các nguyên tố hóa học");
                        if (/liên kết/i.test(lower)) detectedTopics.push("Chương 3: Liên kết hóa học");
                        if (/cân bằng/i.test(lower)) detectedTopics.push("Chương 1: Cân bằng hóa học");
                        if (/nitrogen|sulfur/i.test(lower)) detectedTopics.push("Chương 2: Nitrogen và sulfur");
                        if (/hydrocarbon/i.test(lower)) detectedTopics.push("Chương 4: Hydrocarbon");
                        if (/este|lipid/i.test(lower)) detectedTopics.push("Chương 1: Ester - Lipid");
                        if (/carbohydrate|cacbohidrat/i.test(lower)) detectedTopics.push("Chương 2: Carbohydrate");
                        if (/amin|amino/i.test(lower)) detectedTopics.push("Chương 3: Hợp chất chứa nitrogen");
                        if (/polymer|polime/i.test(lower)) detectedTopics.push("Chương 4: Polymer");
                        if (/pin điện|điện phân/i.test(lower)) detectedTopics.push("Chương 5: Pin điện và điện phân");

                        if (detectedTopics.length > 0) {
                            parsed.grade = detectedGrade || detectGradeFromTopics(detectedTopics);
                            parsed.rows = detectedTopics.map(function(top) {
                                return {
                                    rawTopic: top,
                                    topic: mapToCanonicalTopic(top, parsed.grade),
                                    p1_k: 2, p1_h: 1, p1_v: 0,
                                    p2_k: 0, p2_h: 4, p2_v: 0,
                                    p3_k: 0, p3_h: 1, p3_v: 0,
                                    p4_k: 0, p4_h: 0, p4_v: 0,
                                    codes: ["HH1.1", "HH1.2"]
                                };
                            });
                        }
                    }

                    resolve({
                        grade: parsed.grade || detectedGrade || "12",
                        title: "ĐỀ KIỂM TRA HÓA HỌC (QUÉT TỪ ẢNH)",
                        rows: parsed.rows,
                        ocrText: text,
                        grid: grid
                    });
                }).catch(reject);
            };
            reader.readAsDataURL(imageFile);
        });
    }

    /* --------------------------------------------------------------------------
       6. TEXT / CLIPBOARD PASTE PARSER
       -------------------------------------------------------------------------- */
    function parseClipboardText(rawText) {
        if (!rawText || !rawText.trim()) return null;
        var lines = rawText.trim().split(/[\r\n]+/);
        var grid = lines.map(function(line) {
            if (line.indexOf("\t") !== -1) return line.split("\t");
            if (line.indexOf("|") !== -1) return line.split("|").map(function(s) { return s.trim(); });
            if (line.indexOf(";") !== -1) return line.split(";");
            return line.split(/\s{2,}/);
        });

        var detectedGrade = detectGradeFromText(rawText);
        return parseTableGrid(grid, detectedGrade);
    }

    /* --------------------------------------------------------------------------
       7. SCHOOL SYLLABUS (ĐỀ CƯƠNG TRÊN TRƯỜNG) PARSER
       Tự động bóc tách từng câu hỏi (MCQ, Đúng/Sai, Trả lời ngắn, Tự luận)
       và gán nhãn chương/bài/mức độ để phục vụ sinh đề ĐỘC QUYỀN từ đề cương.
       -------------------------------------------------------------------------- */
    function parseSchoolSyllabusText(text, defaultGrade, sourceName) {
        if (!text || !text.trim()) return [];
        var lines = text.split(/\r?\n/).map(function(l) { return l.trim(); }).filter(Boolean);
        var grade = defaultGrade || importedState.detectedGrade || (typeof currentGrade !== "undefined" ? currentGrade : "12");
        var questions = [];
        var currentChapter = "";
        var currentTopic = "";
        var currentPart = 1;

        var currentQ = null;

        function finalizeCurrentQ() {
            if (!currentQ) return;
            var allText = [currentQ.header].concat(currentQ.bodyLines).join('\n');

            // 1. Nhận diện mức độ nhận thức nếu có gắn thẻ: [Biết], [Hiểu], [Vận dụng], [NB], [TH], [VD]
            var lvlMatch = allText.match(/\[(biết|nhận biết|nb|hiểu|thông hiểu|th|vận dụng|vd|vận dụng cao|vdc)\]|\((biết|nhận biết|nb|hiểu|thông hiểu|th|vận dụng|vd|vận dụng cao|vdc)\)/i);
            if (lvlMatch) {
                var rawLvl = (lvlMatch[1] || lvlMatch[2]).toLowerCase();
                if (/biết|nb/i.test(rawLvl)) currentQ.level = "Biết";
                else if (/hiểu|th\b/i.test(rawLvl)) currentQ.level = "Hiểu";
                else currentQ.level = "Vận dụng";
            }

            var cleanHeader = currentQ.header
                .replace(/^câu\s+\d+[:.]\s*/i, '')
                .replace(/^bài\s+\d+[:.]\s*/i, '')
                .replace(/\[(biết|nhận biết|nb|hiểu|thông hiểu|th|vận dụng|vd|vận dụng cao|vdc)\]\s*/i, '')
                .replace(/\((biết|nhận biết|nb|hiểu|thông hiểu|th|vận dụng|vd|vận dụng cao|vdc)\)\s*/i, '')
                .trim();

            // 2. Nhận diện dạng Phần II: Đúng / Sai
            var stmtLines = currentQ.bodyLines.filter(function(l) { return /^[a-d]\)/i.test(l) || /^\([1-4a-d]\)/i.test(l); });
            var hasPart2Keywords = /đúng.*sai|phát biểu sau|mệnh đề sau|mỗi ý a|ý a, b, c, d/i.test(allText);

            if (currentQ.part === 2 || (stmtLines.length >= 2 && (hasPart2Keywords || currentPart === 2))) {
                currentQ.part = 2;
                currentQ.content = cleanHeader;
                currentQ.statements = [];
                currentQ.bodyLines.forEach(function(l) {
                    if (/^[a-d]\)/i.test(l) || /^\([1-4a-d]\)/i.test(l)) {
                        var sText = l.replace(/^[a-d]\)\s*/i, '').replace(/^\([1-4a-d]\)\s*/i, '');
                        var sLvlM = sText.match(/\[(biết|hiểu|vận dụng|nb|th|vd)\]|\((biết|hiểu|vận dụng|nb|th|vd)\)/i);
                        var sLvl = "Hiểu";
                        if (sLvlM) {
                            var sl = (sLvlM[1] || sLvlM[2]).toLowerCase();
                            if (/biết|nb/i.test(sl)) sLvl = "Biết";
                            else if (/hiểu|th/i.test(sl)) sLvl = "Hiểu";
                            else sLvl = "Vận dụng";
                            sText = sText.replace(/\[(biết|hiểu|vận dụng|nb|th|vd)\]\s*/i, '').replace(/\((biết|hiểu|vận dụng|nb|th|vd)\)\s*/i, '');
                        }
                        currentQ.statements.push({ text: sText.trim(), correct: true, level: sLvl });
                    }
                });
                if (currentQ.statements.length > 0) {
                    questions.push(currentQ);
                }
            } else {
                // 3. Nhận diện dạng Phần I: Trắc nghiệm 4 lựa chọn (MCQ)
                var optLines = [];
                currentQ.bodyLines.forEach(function(l) {
                    if (/^[A-D][.:]\s*/i.test(l)) {
                        optLines.push(l.replace(/^[A-D][.:]\s*/i, '').trim());
                    }
                });

                if (optLines.length < 4) {
                    var singleLineOpts = allText.match(/(?:^|[\s\n])A[.:\s]+(.*?)(?:[\s\n]+)B[.:\s]+(.*?)(?:[\s\n]+)C[.:\s]+(.*?)(?:[\s\n]+)D[.:\s]+(.*?)$/is);
                    if (singleLineOpts && singleLineOpts[1] && singleLineOpts[2] && singleLineOpts[3] && singleLineOpts[4]) {
                        optLines = [singleLineOpts[1].trim(), singleLineOpts[2].trim(), singleLineOpts[3].trim(), singleLineOpts[4].trim()];
                    }
                }

                if (currentQ.part === 1 || optLines.length === 4) {
                    currentQ.part = 1;
                    currentQ.content = cleanHeader;
                    currentQ.options = optLines.slice(0, 4);
                    currentQ.correct = 0;
                    questions.push(currentQ);
                } else if (currentQ.part === 3 || currentPart === 3 || /tính.*bao nhiêu|kết quả.*bằng|đáp số|hãy cho biết giá trị/i.test(allText)) {
                    currentQ.part = 3;
                    currentQ.content = cleanHeader + " " + currentQ.bodyLines.join(' ');
                    currentQ.answer = "10";
                    questions.push(currentQ);
                } else {
                    currentQ.part = 4;
                    currentQ.content = cleanHeader + " " + currentQ.bodyLines.join(' ');
                    questions.push(currentQ);
                }
            }

            currentQ = null;
        }

        for (var i = 0; i < lines.length; i++) {
            var line = lines[i];

            // Quét tiêu đề chương / bài học trong đề cương
            if (/^chương\s+\d+|chủ\s*đề\s+\d+|bài\s+\d+[:.]/i.test(line)) {
                currentChapter = line.replace(/[:\-–].*$/, '').trim();
                currentTopic = line.trim();
                continue;
            }

            // Quét tiêu đề từng phần (Phần I, II, III, IV)
            if (/phần\s*(i|1)\b|trắc nghiệm.*4.*lựa chọn|mcq/i.test(line)) {
                currentPart = 1;
                continue;
            }
            if (/phần\s*(ii|2)\b|đúng\s*\/\s*sai|đúng\s*sai/i.test(line)) {
                currentPart = 2;
                continue;
            }
            if (/phần\s*(iii|3)\b|trả lời ngắn/i.test(line)) {
                currentPart = 3;
                continue;
            }
            if (/phần\s*(iv|4)\b|tự luận/i.test(line)) {
                currentPart = 4;
                continue;
            }

            // Bắt đầu một câu hỏi mới
            var isQStart = /^câu\s+\d+[:.]/i.test(line) || /^bài\s+\d+[:.]/i.test(line);

            if (isQStart) {
                finalizeCurrentQ();
                var topicTitle = currentTopic || currentChapter || "Chủ đề Đề cương trường";
                var canonicalChapter = (typeof mapToCanonicalTopic === "function") ? mapToCanonicalTopic(topicTitle, grade) : topicTitle;

                currentQ = {
                    id: "school_syl_" + Date.now() + "_" + Math.floor(Math.random() * 100000) + "_" + (questions.length + 1),
                    grade: grade,
                    source: sourceName || "Đề cương trường",
                    isFromSchoolSyllabus: true,
                    chapter: canonicalChapter,
                    topic: canonicalChapter,
                    level: (currentPart === 3) ? "Vận dụng" : "Biết",
                    part: currentPart,
                    header: line,
                    bodyLines: []
                };
            } else if (currentQ) {
                currentQ.bodyLines.push(line);
            }
        }
        finalizeCurrentQ();

        return questions;
    }

    function setQuestionSource(src) {
        importedState.questionSource = src;
        var box = document.getElementById("import-syllabus-box");
        if (box) {
            if (src === "syllabus") {
                box.classList.remove("hidden");
            } else {
                box.classList.add("hidden");
            }
        }
    }

    function setStrictSyllabus(val) {
        importedState.strictSyllabusOnly = Boolean(val);
    }

    function handleSyllabusFileSelection(e) {
        var file = e.target.files ? e.target.files[0] : null;
        if (!file) return;

        importedState.syllabusFileName = file.name;
        var nameLabel = document.getElementById("syllabus-file-name-label");
        if (nameLabel) nameLabel.innerText = file.name;

        var statusEl = document.getElementById("import-syllabus-status");
        if (statusEl) {
            statusEl.classList.remove("hidden");
            statusEl.innerHTML = "<div class='flex items-center gap-2 text-indigo-700 bg-indigo-50 p-2 rounded border border-indigo-200'>" +
                "<i class='fa-solid fa-spinner fa-spin'></i> Đang đọc và bóc tách câu hỏi từ đề cương <strong>" + file.name + "</strong>...</div>";
        }

        var ext = file.name.split('.').pop().toLowerCase();
        var reader = new FileReader();

        if (ext === 'docx') {
            reader.onload = function(evt) {
                var arrayBuffer = evt.target.result;
                if (window.mammoth) {
                    window.mammoth.extractRawText({ arrayBuffer: arrayBuffer }).then(function(result) {
                        onSyllabusTextLoaded(result.value, file.name);
                    }).catch(function(err) {
                        if (statusEl) statusEl.innerHTML = "<div class='text-rose-600'>Lỗi đọc docx: " + err.message + "</div>";
                    });
                } else if (window.JSZip) {
                    JSZip.loadAsync(arrayBuffer).then(function(zip) {
                        var docXml = zip.file("word/document.xml");
                        if (!docXml) throw new Error("Không tìm thấy document.xml");
                        return docXml.async("string");
                    }).then(function(xmlText) {
                        var parser = new DOMParser();
                        var xmlDoc = parser.parseFromString(xmlText, "application/xml");
                        var rawText = xmlDoc.body ? xmlDoc.body.textContent : "";
                        onSyllabusTextLoaded(rawText, file.name);
                    }).catch(function(err) {
                        if (statusEl) statusEl.innerHTML = "<div class='text-rose-600'>Lỗi đọc file đề cương: " + err.message + "</div>";
                    });
                } else {
                    if (statusEl) statusEl.innerHTML = "<div class='text-rose-600'>Thư viện đọc file chưa sẵn sàng.</div>";
                }
            };
            reader.readAsArrayBuffer(file);
        } else {
            // Text file
            reader.onload = function(evt) {
                onSyllabusTextLoaded(evt.target.result, file.name);
            };
            reader.readAsText(file, "UTF-8");
        }
    }

    function onSyllabusTextLoaded(text, sourceName) {
        var g = importedState.detectedGrade || (typeof currentGrade !== "undefined" ? currentGrade : "12");
        var questions = parseSchoolSyllabusText(text, g, sourceName);
        importedState.syllabusQuestions = questions;

        var p1 = questions.filter(function(q) { return q.part === 1; }).length;
        var p2 = questions.filter(function(q) { return q.part === 2; }).length;
        var p3 = questions.filter(function(q) { return q.part === 3; }).length;
        var p4 = questions.filter(function(q) { return q.part === 4; }).length;

        var statusEl = document.getElementById("import-syllabus-status");
        if (statusEl) {
            statusEl.classList.remove("hidden");
            statusEl.innerHTML = "<div class='text-emerald-800 bg-emerald-50 border border-emerald-200 p-2.5 rounded-lg flex items-center justify-between gap-2'>" +
                "<div><i class='fa-solid fa-circle-check text-emerald-600 mr-1.5'></i> Đã bóc tách thành công <strong>" + questions.length + " câu hỏi</strong> từ Đề cương <em>" + sourceName + "</em>! " +
                "<span class='text-[11px] text-emerald-700 block mt-0.5'>(P.I Trắc nghiệm: <strong>" + p1 + "</strong> | P.II Đúng/Sai: <strong>" + p2 + "</strong> | P.III Trả lời ngắn: <strong>" + p3 + "</strong> | P.IV Tự luận: <strong>" + p4 + "</strong>)</span></div>" +
                "</div>";
        }

        var badge = document.getElementById("import-syllabus-badge");
        if (badge) {
            badge.className = "font-bold text-emerald-800 bg-emerald-100 px-2.5 py-0.5 rounded border border-emerald-300";
            badge.innerHTML = "<i class='fa-solid fa-check mr-1'></i> Đã sẵn sàng " + questions.length + " câu đề cương";
        }
    }

    function openSyllabusPasteDialog() {
        var pasteBox = document.getElementById("import-syllabus-paste-dialog");
        if (pasteBox) {
            pasteBox.classList.remove("hidden");
        }
    }

    function closeSyllabusPasteDialog() {
        var pasteBox = document.getElementById("import-syllabus-paste-dialog");
        if (pasteBox) {
            pasteBox.classList.add("hidden");
        }
    }

    function processSyllabusPastedText() {
        var area = document.getElementById("import-syllabus-paste-area");
        if (!area || !area.value.trim()) {
            alert("Vui lòng dán văn bản câu hỏi trong đề cương!");
            return;
        }
        onSyllabusTextLoaded(area.value, "Đề cương trường (Dán văn bản)");
        closeSyllabusPasteDialog();
    }

    /* --------------------------------------------------------------------------
       8. AUTOMATIC QUESTION ASSIGNMENT (SYSTEM BANK OR EXCLUSIVELY SYLLABUS)
       -------------------------------------------------------------------------- */
    function autoAssignQuestionsForImportedMatrix() {
        var grade = String(importedState.detectedGrade || (typeof currentGrade !== "undefined" ? currentGrade : "12"));
        var isSyllabus = (importedState.questionSource === "syllabus" && importedState.syllabusQuestions && importedState.syllabusQuestions.length > 0);
        var pool = isSyllabus
            ? importedState.syllabusQuestions
            : ((typeof QUESTION_BANK !== "undefined" && QUESTION_BANK[grade]) ? QUESTION_BANK[grade] : []);

        var usedIds = new Set();
        var usedContentKeys = new Set();
        var totalAssigned = 0;
        var totalAssignedFromSyllabus = 0;

        function getNormKey(q) {
            if (!q || !q.content) return "";
            return q.content.toLowerCase().replace(/[^a-z0-9àáảãạăằắẳẵặâầấẩẫậèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđ]/g, "").slice(0, 60);
        }

        function isAvailable(q) {
            if (usedIds.has(q.id)) return false;
            var nk = getNormKey(q);
            if (nk && usedContentKeys.has(nk)) return false;
            return true;
        }

        function markUsed(q) {
            usedIds.add(q.id);
            var nk = getNormKey(q);
            if (nk) usedContentKeys.add(nk);
            totalAssigned++;
            if (q.isFromSchoolSyllabus) totalAssignedFromSyllabus++;
        }

        var formatEl = document.getElementById("exam-format-select");
        var format = formatEl ? formatEl.value : importedState.detectedFormat;

        matrixRows.forEach(function(row) {
            var req = (typeof getRequiredCountsForRow === "function")
                ? getRequiredCountsForRow(row, format)
                : {
                    p1: (row.p1_k || 0) + (row.p1_h || 0) + (row.p1_v || 0),
                    p2: ((row.p2_k || 0) + (row.p2_h || 0) + (row.p2_v || 0) >= 4 ? 1 : 0),
                    p3: (row.p3_k || 0) + (row.p3_h || 0) + (row.p3_v || 0),
                    p4: ((row.p4_k || 0) + (row.p4_h || 0) + (row.p4_v || 0))
                };

            function pickQuestions(part, level, count) {
                if (count <= 0) return [];
                var res = [];
                // 1. Lọc ứng viên từ pool (nếu là syllabus thì chỉ tìm trong syllabus)
                var candidates = pool.filter(function(q) {
                    if (q.part !== part || !isAvailable(q)) return false;
                    return (typeof isQuestionInTopic === "function")
                        ? isQuestionInTopic(q, row.topic, grade)
                        : true;
                });

                // Khớp chính xác mức độ trước
                var matched = candidates.filter(function(q) {
                    return level ? (q.level === level) : true;
                });
                res = res.concat(matched.slice(0, count));
                res.forEach(markUsed);

                // Nếu chưa đủ, bù bằng các câu cùng chủ đề ở mức độ lân cận trong cùng pool
                if (res.length < count) {
                    var need = count - res.length;
                    var sameTopicOthers = candidates.filter(isAvailable);
                    var extra = sameTopicOthers.slice(0, need);
                    extra.forEach(markUsed);
                    res = res.concat(extra);
                }

                // Nếu là nguồn syllabus và bật chế độ Nghiêm ngặt (strictSyllabusOnly):
                // TUYỆT ĐỐI CHỈ LẤY CÂU HỎI TỪ ĐỀ CƯƠNG TRÊN TRƯỜNG!
                if (isSyllabus && importedState.strictSyllabusOnly) {
                    // Nếu vẫn còn thiếu trong chủ đề này, thử tìm câu cùng dạng thức trong syllabus
                    if (res.length < count) {
                        var needSyl = count - res.length;
                        var anySylCandidates = pool.filter(function(q) {
                            return q.part === part && isAvailable(q);
                        }).slice(0, needSyl);
                        anySylCandidates.forEach(markUsed);
                        res = res.concat(anySylCandidates);
                    }
                    return res;
                }

                // Nếu không phải strict syllabus, fallback sang ngân hàng chung nếu thiếu
                if (res.length < count) {
                    var needGeneral = count - res.length;
                    var generalCandidates = ((typeof QUESTION_BANK !== "undefined" && QUESTION_BANK[grade]) ? QUESTION_BANK[grade] : []).filter(function(q) {
                        return q.part === part && isAvailable(q) && (!level || q.level === level);
                    }).slice(0, needGeneral);
                    generalCandidates.forEach(markUsed);
                    res = res.concat(generalCandidates);
                }

                return res;
            }

            var assigned = [];
            if (row.p1_k > 0) assigned = assigned.concat(pickQuestions(1, 'Biết', row.p1_k));
            if (row.p1_h > 0) assigned = assigned.concat(pickQuestions(1, 'Hiểu', row.p1_h));
            if (row.p1_v > 0) assigned = assigned.concat(pickQuestions(1, 'Vận dụng', row.p1_v));

            if (req.p2 > 0) assigned = assigned.concat(pickQuestions(2, '', req.p2));

            if (row.p3_k > 0) assigned = assigned.concat(pickQuestions(3, 'Biết', row.p3_k));
            if (row.p3_h > 0) assigned = assigned.concat(pickQuestions(3, 'Hiểu', row.p3_h));
            if (row.p3_v > 0) assigned = assigned.concat(pickQuestions(3, 'Vận dụng', row.p3_v));

            if (format !== "TN_ONLY") {
                if (row.p4_k > 0) assigned = assigned.concat(pickQuestions(4, 'Biết', row.p4_k));
                if (row.p4_h > 0) assigned = assigned.concat(pickQuestions(4, 'Hiểu', row.p4_h));
                if (row.p4_v > 0) assigned = assigned.concat(pickQuestions(4, 'Vận dụng', row.p4_v));
            }

            row.assignedQuestions = assigned;
        });

        return {
            totalAssigned: totalAssigned,
            totalTopics: matrixRows.length,
            isSyllabusSource: isSyllabus,
            totalFromSyllabus: totalAssignedFromSyllabus,
            syllabusPoolSize: isSyllabus ? importedState.syllabusQuestions.length : 0
        };
    }

    /* --------------------------------------------------------------------------
       9. MODAL CONTROLLER & INTERACTIVE VIEW
       -------------------------------------------------------------------------- */
    function openImportMatrixModal() {
        var modal = document.getElementById("modal-import-matrix");
        if (!modal) {
            createModalElement();
            modal = document.getElementById("modal-import-matrix");
        }
        if (modal) {
            modal.classList.remove("hidden");
            renderImportPreviewTable();
        }
    }

    function closeImportMatrixModal() {
        var modal = document.getElementById("modal-import-matrix");
        if (modal) modal.classList.add("hidden");
    }

    function switchImportTab(tabKey) {
        importedState.activeTab = tabKey;
        var tabBtns = document.querySelectorAll(".import-tab-btn");
        var tabPanes = document.querySelectorAll(".import-tab-pane");

        tabBtns.forEach(function(b) {
            b.classList.remove("border-brand-600", "text-brand-600", "font-bold");
            b.classList.add("border-transparent", "text-gray-500");
        });
        tabPanes.forEach(function(p) { p.classList.add("hidden"); });

        var activeBtn = document.getElementById("btn-import-tab-" + tabKey);
        var activePane = document.getElementById("pane-import-tab-" + tabKey);
        if (activeBtn) {
            activeBtn.classList.add("border-brand-600", "text-brand-600", "font-bold");
            activeBtn.classList.remove("border-transparent", "text-gray-500");
        }
        if (activePane) activePane.classList.remove("hidden");
    }

    function handleFileSelection(e) {
        var file = e.target.files ? e.target.files[0] : (e.dataTransfer ? e.dataTransfer.files[0] : null);
        if (!file) return;

        importedState.fileName = file.name;
        var ext = file.name.split('.').pop().toLowerCase();
        importedState.fileType = ext;

        var statusEl = document.getElementById("import-process-status");
        if (statusEl) {
            statusEl.classList.remove("hidden");
            statusEl.innerHTML = "<div class='flex items-center gap-2'><i class='fa-solid fa-spinner fa-spin text-brand-500'></i> <span>Đang đọc và phân tích bảng ma trận <strong>" + file.name + "</strong>...</span></div>";
        }

        var reader = new FileReader();

        if (ext === 'docx') {
            reader.onload = function(evt) {
                parseDocxFile(evt.target.result).then(function(res) {
                    onParseSuccess(res);
                }).catch(onParseError);
            };
            reader.readAsArrayBuffer(file);
        } else if (ext === 'xlsx' || ext === 'xls') {
            reader.onload = function(evt) {
                parseExcelFile(evt.target.result).then(function(res) {
                    onParseSuccess(res);
                }).catch(onParseError);
            };
            reader.readAsArrayBuffer(file);
        } else if (ext === 'png' || ext === 'jpg' || ext === 'jpeg' || ext === 'webp') {
            parseImageFile(file, function(pct, msg) {
                if (statusEl) {
                    statusEl.innerHTML = "<div class='space-y-1.5'>" +
                        "<div class='flex items-center justify-between text-xs font-semibold'><span>" + msg + "</span><span>" + pct + "%</span></div>" +
                        "<div class='w-full bg-gray-200 h-2 rounded-full overflow-hidden'><div class='bg-brand-500 h-2 rounded-full transition-all' style='width: " + pct + "%;'></div></div>" +
                        "</div>";
                }
            }).then(function(res) {
                onParseSuccess(res);
            }).catch(onParseError);
        } else {
            // Text fallback
            reader.onload = function(evt) {
                var text = evt.target.result;
                var res = parseClipboardText(text);
                if (res && res.rows.length > 0) {
                    onParseSuccess({
                        grade: res.grade,
                        title: "ĐỀ KIỂM TRA (TỪ TỆP VĂN BẢN)",
                        rows: res.rows
                    });
                } else {
                    onParseError(new Error("Không thể phân tích dữ liệu dạng bảng từ tệp văn bản này."));
                }
            };
            reader.readAsText(file, "UTF-8");
        }
    }

    function onParseSuccess(res) {
        var statusEl = document.getElementById("import-process-status");
        if (statusEl) {
            statusEl.innerHTML = "<div class='text-emerald-700 bg-emerald-50 border border-emerald-200 p-2.5 rounded-lg flex items-center justify-between gap-2'>" +
                "<div class='flex items-center gap-2'><i class='fa-solid fa-check-circle text-emerald-600 text-base'></i> <span>Phân tích thành công! Đã trích xuất <strong>" + res.rows.length + " chủ đề</strong> từ <em>" + importedState.fileName + "</em>.</span></div>" +
                "<span class='bg-emerald-200 text-emerald-900 text-xs px-2 py-0.5 rounded font-bold'>Khối " + res.grade + "</span>" +
                "</div>";
        }

        importedState.detectedGrade = res.grade || "12";
        importedState.examTitle = res.title || importedState.examTitle;
        importedState.rows = res.rows || [];

        var gradeSel = document.getElementById("import-grade-select");
        if (gradeSel) gradeSel.value = importedState.detectedGrade;
        var titleInput = document.getElementById("import-title-input");
        if (titleInput) titleInput.value = importedState.examTitle;

        renderImportPreviewTable();
    }

    function onParseError(err) {
        var statusEl = document.getElementById("import-process-status");
        if (statusEl) {
            statusEl.innerHTML = "<div class='text-rose-700 bg-rose-50 border border-rose-200 p-2.5 rounded-lg flex items-center gap-2'>" +
                "<i class='fa-solid fa-circle-exclamation text-rose-600 text-base'></i>" +
                "<span>Lỗi phân tích: " + (err.message || err) + ". Vui lòng thử dùng chức năng 'Dán Văn Bản' hoặc chọn tệp khác.</span>" +
                "</div>";
        }
    }

    function handlePasteAnalysis() {
        var area = document.getElementById("import-paste-textarea");
        if (!area || !area.value.trim()) {
            alert("Vui lòng dán văn bản bảng ma trận hoặc bản đặc tả vào ô trống.");
            return;
        }

        var res = parseClipboardText(area.value);
        if (res && res.rows.length > 0) {
            importedState.fileName = "Văn bản dán từ Clipboard";
            onParseSuccess({
                grade: res.grade,
                title: "ĐỀ KIỂM TRA (TỪ BẢNG DÁN)",
                rows: res.rows
            });
        } else {
            onParseError(new Error("Không tìm thấy cấu trúc bảng hợp lệ trong đoạn văn bản đã dán."));
        }
    }

    function renderImportPreviewTable() {
        var container = document.getElementById("import-preview-tbody");
        if (!container) return;

        if (!importedState.rows || importedState.rows.length === 0) {
            container.innerHTML = "<tr><td colspan='7' class='text-center py-8 text-gray-400 italic bg-gray-50/50'>" +
                "<i class='fa-solid fa-file-arrow-up text-2xl mb-1 text-gray-300 block'></i>" +
                "Chưa có dữ liệu ma trận. Hãy chọn tệp Word (.docx), Excel (.xlsx) hoặc Ảnh (.png/.jpg) ở phía trên." +
                "</td></tr>";
            updateImportSummaryStats();
            return;
        }

        container.innerHTML = importedState.rows.map(function(r, idx) {
            return "<tr class='hover:bg-amber-50/40 transition border-b border-gray-100'>" +
                "<td class='text-center font-bold text-gray-500'>" + (idx + 1) + "</td>" +
                "<td>" +
                "<input type='text' value='" + (r.topic || '').replace(/'/g, "&apos;") + "' onchange='window.matrixImporter.updateRowField(" + idx + ", \"topic\", this.value)' class='w-full text-xs p-1 border rounded font-semibold text-gray-800 focus:ring-1 focus:ring-brand-500'>" +
                "<div class='text-[10px] text-brand-700 truncate mt-0.5' title='Chuẩn hóa theo SGK: " + (r.canonicalTopic || r.topic) + "'>" +
                "<i class='fa-solid fa-bookmark text-amber-500 mr-1'></i>" + (r.canonicalTopic || r.topic) +
                "</div>" +
                "</td>" +
                "<td class='text-center bg-green-50/30'><div class='flex items-center justify-center gap-1'>" +
                "<input type='number' min='0' value='" + (r.p1_k || 0) + "' onchange='window.matrixImporter.updateRowField(" + idx + ", \"p1_k\", this.value)' class='w-8 text-center border rounded text-xs' title='Nhận biết'>" +
                "<input type='number' min='0' value='" + (r.p1_h || 0) + "' onchange='window.matrixImporter.updateRowField(" + idx + ", \"p1_h\", this.value)' class='w-8 text-center border rounded text-xs' title='Thông hiểu'>" +
                "<input type='number' min='0' value='" + (r.p1_v || 0) + "' onchange='window.matrixImporter.updateRowField(" + idx + ", \"p1_v\", this.value)' class='w-8 text-center border rounded text-xs' title='Vận dụng'>" +
                "</div></td>" +
                "<td class='text-center bg-blue-50/30'><div class='flex items-center justify-center gap-1'>" +
                "<input type='number' min='0' value='" + (r.p2_k || 0) + "' onchange='window.matrixImporter.updateRowField(" + idx + ", \"p2_k\", this.value)' class='w-8 text-center border rounded text-xs text-blue-900 font-bold' title='Biết (ý)'>" +
                "<input type='number' min='0' value='" + (r.p2_h || 0) + "' onchange='window.matrixImporter.updateRowField(" + idx + ", \"p2_h\", this.value)' class='w-8 text-center border rounded text-xs text-blue-900 font-bold' title='Hiểu (ý)'>" +
                "<input type='number' min='0' value='" + (r.p2_v || 0) + "' onchange='window.matrixImporter.updateRowField(" + idx + ", \"p2_v\", this.value)' class='w-8 text-center border rounded text-xs text-blue-900 font-bold' title='Vận dụng (ý)'>" +
                "</div></td>" +
                "<td class='text-center bg-purple-50/30'><div class='flex items-center justify-center gap-1'>" +
                "<input type='number' min='0' value='" + (r.p3_k || 0) + "' onchange='window.matrixImporter.updateRowField(" + idx + ", \"p3_k\", this.value)' class='w-8 text-center border rounded text-xs' title='Nhận biết'>" +
                "<input type='number' min='0' value='" + (r.p3_h || 0) + "' onchange='window.matrixImporter.updateRowField(" + idx + ", \"p3_h\", this.value)' class='w-8 text-center border rounded text-xs' title='Thông hiểu'>" +
                "<input type='number' min='0' value='" + (r.p3_v || 0) + "' onchange='window.matrixImporter.updateRowField(" + idx + ", \"p3_v\", this.value)' class='w-8 text-center border rounded text-xs' title='Vận dụng'>" +
                "</div></td>" +
                "<td class='text-center bg-rose-50/30'><div class='flex items-center justify-center gap-1'>" +
                "<input type='number' min='0' value='" + (r.p4_k || 0) + "' onchange='window.matrixImporter.updateRowField(" + idx + ", \"p4_k\", this.value)' class='w-8 text-center border rounded text-xs' title='Nhận biết'>" +
                "<input type='number' min='0' value='" + (r.p4_h || 0) + "' onchange='window.matrixImporter.updateRowField(" + idx + ", \"p4_h\", this.value)' class='w-8 text-center border rounded text-xs' title='Thông hiểu'>" +
                "<input type='number' min='0' value='" + (r.p4_v || 0) + "' onchange='window.matrixImporter.updateRowField(" + idx + ", \"p4_v\", this.value)' class='w-8 text-center border rounded text-xs' title='Vận dụng'>" +
                "</div></td>" +
                "<td class='text-center'>" +
                "<button type='button' onclick='window.matrixImporter.deleteRow(" + idx + ")' class='text-rose-500 hover:text-rose-700 p-1'><i class='fa-solid fa-trash-can'></i></button>" +
                "</td>" +
                "</tr>";
        }).join('');

        updateImportSummaryStats();
    }

    function updateRowField(idx, field, val) {
        if (!importedState.rows[idx]) return;
        if (field === "topic") {
            importedState.rows[idx].topic = val;
            importedState.rows[idx].canonicalTopic = mapToCanonicalTopic(val, importedState.detectedGrade);
        } else {
            importedState.rows[idx][field] = parseInt(val, 10) || 0;
        }
        updateImportSummaryStats();
    }

    function deleteRow(idx) {
        if (!importedState.rows[idx]) return;
        importedState.rows.splice(idx, 1);
        renderImportPreviewTable();
    }

    function addNewRow() {
        var g = importedState.detectedGrade || "12";
        importedState.rows.push({
            rawTopic: "Chủ đề mới",
            topic: (g === "12") ? "Chương 1: Ester - Lipid - Bài 1: Ester – Lipid" : (g === "11" ? "Chương 1: Cân bằng hóa học" : "Chương 1: Cấu tạo nguyên tử"),
            canonicalTopic: (g === "12") ? "Chương 1: Ester - Lipid - Bài 1: Ester – Lipid" : (g === "11" ? "Chương 1: Cân bằng hóa học" : "Chương 1: Cấu tạo nguyên tử"),
            p1_k: 1, p1_h: 1, p1_v: 0,
            p2_k: 0, p2_h: 0, p2_v: 0,
            p3_k: 0, p3_h: 0, p3_v: 0,
            p4_k: 0, p4_h: 0, p4_v: 0,
            codes: ["HH1.1", "HH1.2"]
        });
        renderImportPreviewTable();
    }

    function updateImportSummaryStats() {
        var p1 = 0, p2 = 0, p3 = 0, p4 = 0;
        (importedState.rows || []).forEach(function(r) {
            p1 += ((r.p1_k || 0) + (r.p1_h || 0) + (r.p1_v || 0));
            p2 += ((r.p2_k || 0) + (r.p2_h || 0) + (r.p2_v || 0));
            p3 += ((r.p3_k || 0) + (r.p3_h || 0) + (r.p3_v || 0));
            p4 += ((r.p4_k || 0) + (r.p4_h || 0) + (r.p4_v || 0));
        });

        var p2Questions = (p2 >= 4) ? Math.round(p2 / 4) : (p2 > 0 ? (p2 / 4).toFixed(1) : 0);
        var score1 = p1 * 0.25;
        var score2 = (p2 >= 4) ? (Math.round(p2 / 4) * 1.0) : (p2 * 0.25);
        var score3 = p3 * 0.5;
        var score4 = p4 * 1.0;
        var totalScore = score1 + score2 + score3 + score4;

        var statEl = document.getElementById("import-summary-bar");
        if (statEl) {
            statEl.innerHTML = "<span><strong>P.I (MCQ):</strong> " + p1 + " câu (" + score1.toFixed(2) + "đ)</span>" +
                "<span><strong>P.II (Đ/S):</strong> " + p2 + " ý (~" + p2Questions + " câu) (" + score2.toFixed(2) + "đ)</span>" +
                "<span><strong>P.III (TLN):</strong> " + p3 + " câu (" + score3.toFixed(2) + "đ)</span>" +
                "<span><strong>P.IV (TL):</strong> " + p4 + " câu (" + score4.toFixed(2) + "đ)</span>" +
                "<span class='text-brand-700 font-extrabold text-sm ml-auto'>Tổng: " + totalScore.toFixed(2) + "/10.0đ</span>";
        }
    }

    /* --------------------------------------------------------------------------
       10. APPLY & AUTO-GENERATE EXAM WORKFLOW
       -------------------------------------------------------------------------- */
    function applyImportedMatrix(autoGenerate) {
        if (!importedState.rows || importedState.rows.length === 0) {
            alert("Chưa có dòng ma trận nào để áp dụng. Vui lòng tải file ma trận hoặc dán bảng.");
            return;
        }

        // Kiểm tra điều kiện nguồn câu hỏi Đề Cương Trường
        if (autoGenerate && importedState.questionSource === "syllabus") {
            if (!importedState.syllabusQuestions || importedState.syllabusQuestions.length === 0) {
                alert("Thầy đã chọn chế độ 'Chỉ lấy câu hỏi từ đề cương trên trường' nhưng chưa chọn tệp hoặc chưa dán đề cương!\n\nVui lòng bấm 'Chọn File Đề Cương' (.docx / .txt) hoặc bấm 'Dán Văn Bản' để nạp câu hỏi đề cương trước khi sinh đề.");
                return;
            }
        }

        // 1. Đồng bộ Khối lớp
        var targetGrade = document.getElementById("import-grade-select") ? document.getElementById("import-grade-select").value : importedState.detectedGrade;
        if (targetGrade && typeof currentGrade !== "undefined") {
            currentGrade = targetGrade;
            var mGradeSel = document.getElementById("matrix-grade-select");
            if (mGradeSel) mGradeSel.value = currentGrade;
            var gGradeSel = document.getElementById("grade-select");
            if (gGradeSel) gGradeSel.value = currentGrade;
        }

        // 2. Đồng bộ Hình thức thi
        var formatSel = document.getElementById("import-format-select");
        if (formatSel) {
            var examFormatSel = document.getElementById("exam-format-select");
            if (examFormatSel) {
                examFormatSel.value = formatSel.value;
                if (typeof toggleEssayColumns === "function") toggleEssayColumns();
            }
        }

        // 3. Đồng bộ Tiêu đề đề thi
        var titleInput = document.getElementById("import-title-input");
        if (titleInput && titleInput.value.trim()) {
            var examTitleInput = document.getElementById("exam-title-input");
            if (examTitleInput) examTitleInput.value = titleInput.value.trim();
        }

        // 4. Nạp vào matrixRows toàn cục
        matrixRows = importedState.rows.map(function(r) {
            return {
                topic: r.topic || r.canonicalTopic,
                p1_k: r.p1_k || 0,
                p1_h: r.p1_h || 0,
                p1_v: r.p1_v || 0,
                p2_k: r.p2_k || 0,
                p2_h: r.p2_h || 0,
                p2_v: r.p2_v || 0,
                p3_k: r.p3_k || 0,
                p3_h: r.p3_h || 0,
                p3_v: r.p3_v || 0,
                p4_k: r.p4_k || 0,
                p4_h: r.p4_h || 0,
                p4_v: r.p4_v || 0,
                codes: r.codes || ["HH1.1", "HH1.2"],
                assignedQuestions: []
            };
        });

        // 5. Cập nhật giao diện Ma trận & Bản đặc tả
        if (typeof renderMatrixEditor === "function") renderMatrixEditor();
        if (typeof recalculateMatrixTotals === "function") recalculateMatrixTotals();

        // 6. Tự động sinh đề và gán câu hỏi nếu được yêu cầu
        if (autoGenerate) {
            var assignStats = autoAssignQuestionsForImportedMatrix();
            if (typeof renderSpecificationTable === "function") renderSpecificationTable();

            // Chuyển sang Tab 4 và hiển thị đề thi hoàn chỉnh
            if (typeof switchTab === "function") switchTab('generator');
            if (typeof renderGeneratedExam === "function") renderGeneratedExam();

            closeImportMatrixModal();

            if (assignStats.isSyllabusSource) {
                showSuccessBanner("Đã sinh đề thi hoàn chỉnh CHỈ TỪ ĐỀ CƯƠNG TRÊN TRƯỜNG (" + assignStats.totalAssigned + " câu hỏi) theo đúng ma trận & bản đặc tả!");
            } else {
                showSuccessBanner("Đã nhập thành công ma trận (" + assignStats.totalTopics + " chủ đề) và tự động sinh đề thi hoàn chỉnh với " + assignStats.totalAssigned + " câu hỏi chuẩn GDPT 2018!");
            }
        } else {
            if (typeof renderSpecificationTable === "function") renderSpecificationTable();
            closeImportMatrixModal();
            showSuccessBanner("Đã nạp thành công ma trận " + matrixRows.length + " chủ đề vào hệ thống! Bạn có thể xem bản đặc tả hoặc bấm 'Tự động lấy câu hỏi'.");
        }
    }

    function showSuccessBanner(message) {
        var banner = document.createElement("div");
        banner.className = "fixed bottom-5 right-5 z-50 bg-brand-900 text-white px-5 py-3.5 rounded-xl shadow-2xl border border-brand-500/30 flex items-center gap-3 transition-all transform translate-y-0 text-sm max-w-md animate-bounce";
        banner.innerHTML = "<div class='w-8 h-8 rounded-full bg-emerald-500 text-white flex items-center justify-center flex-shrink-0'><i class='fa-solid fa-check'></i></div>" +
            "<div><p class='font-bold text-amber-300'>Thành công!</p><p class='text-xs text-gray-200 mt-0.5'>" + message + "</p></div>";
        document.body.appendChild(banner);
        setTimeout(function() {
            banner.style.opacity = "0";
            setTimeout(function() { banner.remove(); }, 500);
        }, 5000);
    }

    /* --------------------------------------------------------------------------
       11. MODAL DOM CREATOR (INSERT HTML IF NOT PRESENT)
       -------------------------------------------------------------------------- */
    function createModalElement() {
        if (document.getElementById("modal-import-matrix")) return;

        var modalHtml = `
        <div id="modal-import-matrix" class="fixed inset-0 bg-black/60 z-50 flex items-center justify-center p-3 sm:p-4 hidden no-print backdrop-blur-sm">
            <div class="bg-white rounded-2xl shadow-2xl max-w-5xl w-full p-5 sm:p-6 space-y-4 max-h-[94vh] overflow-y-auto border border-gray-100 flex flex-col">
                <!-- Header -->
                <div class="flex items-center justify-between border-b border-gray-200 pb-3 flex-shrink-0">
                    <div class="flex items-center gap-3">
                        <div class="w-10 h-10 rounded-xl bg-gradient-to-tr from-amber-500 to-orange-500 text-white flex items-center justify-center text-lg font-bold shadow-md">
                            <i class="fa-solid fa-file-import"></i>
                        </div>
                        <div>
                            <h3 class="font-bold text-base text-gray-900 flex items-center gap-2">
                                Nhập Ma Trận & Bản Đặc Tả Đề Kiểm Tra Từ Máy Tính
                                <span class="bg-amber-100 text-amber-800 text-[11px] px-2 py-0.5 rounded-full font-bold">Word / Excel / Ảnh / Đề Cương</span>
                            </h3>
                            <p class="text-xs text-gray-500">Tự động nhận diện cấu trúc ma trận và hỗ trợ chỉ lấy câu hỏi từ đề cương ôn tập trên trường.</p>
                        </div>
                    </div>
                    <button type="button" onclick="window.matrixImporter.closeModal()" class="text-gray-400 hover:text-gray-700 p-1.5 rounded-lg hover:bg-gray-100 transition">
                        <i class="fa-solid fa-xmark text-xl"></i>
                    </button>
                </div>

                <!-- BƯỚC 1: NHẬP MA TRẬN & BẢN ĐẶC TẢ TRÊN TRƯỜNG -->
                <div class="space-y-2 flex-shrink-0">
                    <div class="flex items-center justify-between">
                        <span class="text-xs font-bold text-gray-800 flex items-center gap-1.5">
                            <span class="w-5 h-5 rounded-full bg-amber-600 text-white flex items-center justify-center text-[10px] font-extrabold">1</span>
                            Nhập Ma Trận & Bản Đặc Tả Trên Trường (File Ảnh / Word / Excel / Dán bảng):
                        </span>
                        <!-- Sub tabs -->
                        <div class="flex text-xs font-semibold gap-2">
                            <button type="button" id="btn-import-tab-file" onclick="window.matrixImporter.switchTab('file')" class="import-tab-btn py-1 px-2.5 border-b-2 border-brand-600 text-brand-600 font-bold flex items-center gap-1">
                                <i class="fa-solid fa-cloud-arrow-up"></i> Tải Tệp Ma Trận
                            </button>
                            <button type="button" id="btn-import-tab-paste" onclick="window.matrixImporter.switchTab('paste')" class="import-tab-btn py-1 px-2.5 border-b-2 border-transparent text-gray-500 flex items-center gap-1 hover:text-gray-700">
                                <i class="fa-solid fa-paste"></i> Dán Bảng Ma Trận
                            </button>
                        </div>
                    </div>

                    <!-- Pane 1: File Drop Zone -->
                    <div id="pane-import-tab-file" class="import-tab-pane">
                        <div id="import-drop-zone" class="border-2 border-dashed border-amber-300 hover:border-amber-500 bg-amber-50/30 hover:bg-amber-50/60 transition rounded-xl p-4 text-center cursor-pointer relative">
                            <input type="file" id="import-file-input" accept=".docx,.xlsx,.xls,.png,.jpg,.jpeg,.txt" onchange="window.matrixImporter.handleFile(event)" class="absolute inset-0 opacity-0 cursor-pointer w-full h-full">
                            <div class="space-y-1 pointer-events-none">
                                <i class="fa-solid fa-folder-open text-amber-600 text-2xl"></i>
                                <p class="text-xs font-bold text-gray-800">Chọn tệp Ma Trận & Bản Đặc Tả (Word .docx, Excel .xlsx, hoặc Ảnh .png/.jpg)</p>
                                <p class="text-[11px] text-gray-500">Hệ thống sẽ tự động quét các dòng chủ đề, số câu MCQ, Đúng/Sai, Trả lời ngắn và Tự luận.</p>
                            </div>
                        </div>
                    </div>

                    <!-- Pane 2: Clipboard Paste -->
                    <div id="pane-import-tab-paste" class="import-tab-pane hidden space-y-1.5">
                        <textarea id="import-paste-textarea" rows="3" class="w-full text-xs font-mono p-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-brand-500" placeholder="Dán bảng ma trận (copy từ Excel, Word hoặc web)..."></textarea>
                        <button type="button" onclick="window.matrixImporter.analyzePaste()" class="px-3.5 py-1.5 bg-brand-600 hover:bg-brand-700 text-white text-xs font-bold rounded-lg shadow-sm transition flex items-center gap-1.5">
                            <i class="fa-solid fa-wand-magic-sparkles text-amber-300"></i> Phân Tích Bảng Đã Dán
                        </button>
                    </div>

                    <!-- Status Notification Box -->
                    <div id="import-process-status" class="hidden text-xs"></div>
                </div>

                <!-- BƯỚC 2: CHỌN NGUỒN CÂU HỎI ĐỂ SINH ĐỀ (HỆ THỐNG VS ĐỀ CƯƠNG TRÊN TRƯỜNG) -->
                <div class="bg-gradient-to-r from-blue-50 via-indigo-50 to-purple-50 p-3.5 rounded-xl border border-indigo-200 space-y-2.5 text-xs flex-shrink-0">
                    <div class="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-2 border-b border-indigo-200/60 pb-2">
                        <span class="font-bold text-indigo-950 flex items-center gap-1.5">
                            <span class="w-5 h-5 rounded-full bg-indigo-600 text-white flex items-center justify-center text-[10px] font-extrabold">2</span>
                            Nguồn Câu Hỏi Khi Sinh Đề:
                        </span>
                        <div class="flex flex-wrap items-center gap-3">
                            <label class="flex items-center gap-1.5 cursor-pointer text-gray-700 font-medium">
                                <input type="radio" name="import-q-source" value="bank" checked onchange="window.matrixImporter.setQuestionSource('bank')" class="text-brand-600 focus:ring-brand-500">
                                <span>Ngân hàng chuẩn hệ thống (10.094 câu)</span>
                            </label>
                            <label class="flex items-center gap-1.5 cursor-pointer text-indigo-900 font-extrabold bg-white px-2.5 py-1 rounded-lg border border-indigo-300 shadow-xs">
                                <input type="radio" name="import-q-source" value="syllabus" onchange="window.matrixImporter.setQuestionSource('syllabus')" class="text-indigo-600 focus:ring-indigo-500">
                                <span>🏫 CHỈ LẤY CÂU HỎI TỪ ĐỀ CƯƠNG TRÊN TRƯỜNG</span>
                            </label>
                        </div>
                    </div>

                    <!-- Khung nạp Đề Cương Trên Trường (Hiển thị khi chọn syllabus) -->
                    <div id="import-syllabus-box" class="hidden space-y-2 pt-1">
                        <div class="grid grid-cols-1 sm:grid-cols-2 gap-2">
                            <div class="bg-white p-2.5 rounded-lg border border-indigo-200 flex items-center justify-between gap-2 shadow-xs">
                                <div class="truncate">
                                    <p class="font-bold text-gray-800 flex items-center gap-1">
                                        <i class="fa-solid fa-file-word text-blue-600"></i> Tải Tệp Đề Cương (.docx / .txt)
                                    </p>
                                    <p class="text-[10px] text-gray-500 truncate" id="syllabus-file-name-label">Chọn tệp đề cương ôn tập trên trường</p>
                                </div>
                                <label class="cursor-pointer bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-[11px] px-3 py-1.5 rounded-lg shadow-xs transition flex items-center gap-1 flex-shrink-0">
                                    <i class="fa-solid fa-arrow-up-from-bracket"></i> Chọn Tệp
                                    <input type="file" id="import-syllabus-file-input" accept=".docx,.txt" onchange="window.matrixImporter.handleSyllabusFile(event)" class="hidden">
                                </label>
                            </div>
                            <div class="bg-white p-2.5 rounded-lg border border-indigo-200 flex items-center justify-between gap-2 shadow-xs">
                                <div class="truncate">
                                    <p class="font-bold text-gray-800 flex items-center gap-1">
                                        <i class="fa-solid fa-paste text-teal-600"></i> Hoặc Dán Văn Bản Đề Cương
                                    </p>
                                    <p class="text-[10px] text-gray-500 truncate">Dán danh sách câu hỏi đề cương trường</p>
                                </div>
                                <button type="button" onclick="window.matrixImporter.openSyllabusPasteDialog()" class="bg-teal-600 hover:bg-teal-700 text-white font-bold text-[11px] px-3 py-1.5 rounded-lg shadow-xs transition flex items-center gap-1 flex-shrink-0">
                                    <i class="fa-solid fa-clipboard-list"></i> Dán Văn Bản
                                </button>
                            </div>
                        </div>

                        <!-- Trạng thái nạp Đề Cương -->
                        <div id="import-syllabus-status" class="hidden text-xs"></div>

                        <div class="flex flex-wrap items-center justify-between text-[11px] bg-indigo-100/70 p-2 rounded-lg border border-indigo-200 gap-2">
                            <label class="flex items-center gap-1.5 cursor-pointer font-bold text-indigo-950">
                                <input type="checkbox" id="import-syllabus-strict-checkbox" checked onchange="window.matrixImporter.setStrictSyllabus(this.checked)" class="rounded text-indigo-600 focus:ring-indigo-500">
                                <span>Chỉ được lấy câu hỏi từ đề cương trên trường (Không tự ý lấy thêm từ ngân hàng ngoài)</span>
                            </label>
                            <span id="import-syllabus-badge" class="font-bold text-indigo-800 bg-white px-2 py-0.5 rounded border border-indigo-200">
                                Chưa nạp câu hỏi đề cương
                            </span>
                        </div>
                    </div>
                </div>

                <!-- Detected Controls: Grade, Title, Format -->
                <div class="bg-gray-50 p-2.5 rounded-xl border border-gray-200 grid grid-cols-1 sm:grid-cols-3 gap-3 text-xs flex-shrink-0">
                    <div>
                        <label class="block font-bold text-gray-700 mb-1">Khối Lớp Kiểm Tra:</label>
                        <select id="import-grade-select" class="w-full p-1.5 border border-gray-300 rounded-lg font-bold bg-white text-brand-700">
                            <option value="10">Hóa Học Lớp 10</option>
                            <option value="11">Hóa Học Lớp 11</option>
                            <option value="12" selected>Hóa Học Lớp 12</option>
                        </select>
                    </div>
                    <div>
                        <label class="block font-bold text-gray-700 mb-1">Hình Thức Đề:</label>
                        <select id="import-format-select" class="w-full p-1.5 border border-gray-300 rounded-lg font-medium bg-white">
                            <option value="MIXED">Trắc Nghiệm + Tự Luận (CV 7991)</option>
                            <option value="TN_ONLY">100% Trắc Nghiệm (Chuẩn 2025)</option>
                        </select>
                    </div>
                    <div>
                        <label class="block font-bold text-gray-700 mb-1">Tiêu Đề Đề Thi:</label>
                        <input type="text" id="import-title-input" value="ĐỀ KIỂM TRA ĐỊNH KỲ HÓA HỌC" class="w-full p-1.5 border border-gray-300 rounded-lg font-semibold bg-white">
                    </div>
                </div>

                <!-- Preview Table -->
                <div class="space-y-2 flex-grow overflow-hidden flex flex-col">
                    <div class="flex items-center justify-between flex-shrink-0">
                        <span class="text-xs font-bold text-gray-800 flex items-center gap-1.5">
                            <span class="w-5 h-5 rounded-full bg-emerald-600 text-white flex items-center justify-center text-[10px] font-extrabold">3</span>
                            Xem & Hiệu Chỉnh Phân Bổ Câu Hỏi Theo Từng Chủ Đề:
                        </span>
                        <button type="button" onclick="window.matrixImporter.addRow()" class="px-2.5 py-1 bg-gray-100 hover:bg-gray-200 text-gray-700 font-semibold text-xs rounded border border-gray-300 transition flex items-center gap-1">
                            <i class="fa-solid fa-plus text-emerald-600"></i> Thêm Dòng
                        </button>
                    </div>

                    <div class="overflow-x-auto overflow-y-auto max-h-56 border border-gray-200 rounded-lg shadow-inner">
                        <table class="w-full text-left text-xs border-collapse">
                            <thead class="bg-gray-100 text-gray-700 sticky top-0 border-b border-gray-200">
                                <tr>
                                    <th class="p-1.5 w-8 text-center">TT</th>
                                    <th class="p-1.5 w-1/3">Chủ Đề Kiến Thức</th>
                                    <th class="p-1.5 text-center bg-green-100/70 text-green-950 font-bold" title="Phần I: 4 lựa chọn">P.I (B - H - VD)</th>
                                    <th class="p-1.5 text-center bg-blue-100/70 text-blue-950 font-bold" title="Phần II: Đúng/Sai (Ý)">P.II (B - H - VD)</th>
                                    <th class="p-1.5 text-center bg-purple-100/70 text-purple-950 font-bold" title="Phần III: Trả lời ngắn">P.III (B - H - VD)</th>
                                    <th class="p-1.5 text-center bg-rose-100/70 text-rose-950 font-bold" title="Phần IV: Tự luận">P.IV (B - H - VD)</th>
                                    <th class="p-1.5 w-8 text-center">Xóa</th>
                                </tr>
                            </thead>
                            <tbody id="import-preview-tbody">
                                <!-- Populated dynamically -->
                            </tbody>
                        </table>
                    </div>

                    <!-- Summary Stats Bar -->
                    <div id="import-summary-bar" class="p-2 bg-amber-50 rounded-lg border border-amber-200 text-xs font-semibold text-amber-900 flex flex-wrap items-center gap-3 flex-shrink-0">
                        <span>Chưa có dữ liệu</span>
                    </div>
                </div>

                <!-- Footer Buttons -->
                <div class="flex items-center justify-between border-t border-gray-200 pt-3 flex-shrink-0">
                    <button type="button" onclick="window.matrixImporter.closeModal()" class="px-4 py-2 border border-gray-300 rounded-lg text-xs font-semibold text-gray-700 hover:bg-gray-100 transition">
                        Đóng
                    </button>
                    <div class="flex items-center gap-2">
                        <button type="button" onclick="window.matrixImporter.apply(false)" class="px-4 py-2 bg-gray-100 hover:bg-gray-200 border border-gray-300 text-gray-800 text-xs font-bold rounded-lg transition">
                            <i class="fa-solid fa-table mr-1"></i> Áp Dụng Vào Bảng Ma Trận
                        </button>
                        <button type="button" onclick="window.matrixImporter.apply(true)" class="px-5 py-2 bg-gradient-to-r from-teal-600 via-emerald-600 to-green-600 hover:from-teal-700 hover:to-green-700 text-white text-xs font-extrabold rounded-lg shadow-md transition flex items-center gap-1.5">
                            <i class="fa-solid fa-wand-magic-sparkles text-amber-300"></i>
                            <span>⚡ ÁP DỤNG & TỰ ĐỘNG SINH ĐỀ THI</span>
                        </button>
                    </div>
                </div>
            </div>
        </div>

        <!-- DIALOG DÁN VĂN BẢN ĐỀ CƯƠNG TRƯỜNG -->
        <div id="import-syllabus-paste-dialog" class="fixed inset-0 bg-black/70 z-50 flex items-center justify-center p-4 hidden no-print backdrop-blur-sm">
            <div class="bg-white rounded-2xl shadow-2xl max-w-2xl w-full p-5 space-y-3 border border-gray-100">
                <div class="flex items-center justify-between border-b pb-2">
                    <h4 class="font-bold text-sm text-gray-800 flex items-center gap-2">
                        <i class="fa-solid fa-paste text-indigo-600"></i>
                        Dán Nội Dung Câu Hỏi Đề Cương Trên Trường
                    </h4>
                    <button type="button" onclick="window.matrixImporter.closeSyllabusPasteDialog()" class="text-gray-400 hover:text-gray-700">
                        <i class="fa-solid fa-xmark text-lg"></i>
                    </button>
                </div>
                <p class="text-xs text-gray-600">Sao chép văn bản câu hỏi từ file Word đề cương của trường (bao gồm Câu 1, Câu 2, các phương án A, B, C, D hoặc a, b, c, d) rồi dán vào đây:</p>
                <textarea id="import-syllabus-paste-area" rows="10" class="w-full text-xs font-mono p-3 border border-gray-300 rounded-xl focus:ring-2 focus:ring-indigo-500" placeholder="Câu 1. Este no, đơn chức, mạch hở có công thức chung là...&#10;A. CnH2nO2&#10;B. CnH2nO&#10;..."></textarea>
                <div class="flex items-center justify-end gap-2 pt-2 border-t">
                    <button type="button" onclick="window.matrixImporter.closeSyllabusPasteDialog()" class="px-3 py-1.5 border border-gray-300 rounded-lg text-xs font-medium text-gray-700 hover:bg-gray-100">Hủy</button>
                    <button type="button" onclick="window.matrixImporter.processSyllabusPaste()" class="px-4 py-1.5 bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs rounded-lg shadow-sm transition">
                        <i class="fa-solid fa-check mr-1"></i> Bóc Tách & Nạp Đề Cương
                    </button>
                </div>
            </div>
        </div>
        `;

        var wrapper = document.createElement("div");
        wrapper.innerHTML = modalHtml;
        document.body.appendChild(wrapper.firstElementChild);
        document.body.appendChild(wrapper.lastElementChild);
    }

    // Export module to global scope
    window.matrixImporter = {
        openModal: openImportMatrixModal,
        closeModal: closeImportMatrixModal,
        switchTab: switchImportTab,
        handleFile: handleFileSelection,
        analyzePaste: handlePasteAnalysis,
        updateRowField: updateRowField,
        deleteRow: deleteRow,
        addRow: addNewRow,
        apply: applyImportedMatrix,
        parseDocx: parseDocxFile,
        parseExcel: parseExcelFile,
        parseImage: parseImageFile,
        parseGrid: parseTableGrid,
        autoAssign: autoAssignQuestionsForImportedMatrix,

        // Syllabus additions
        setQuestionSource: setQuestionSource,
        setStrictSyllabus: setStrictSyllabus,
        handleSyllabusFile: handleSyllabusFileSelection,
        openSyllabusPasteDialog: openSyllabusPasteDialog,
        closeSyllabusPasteDialog: closeSyllabusPasteDialog,
        processSyllabusPaste: processSyllabusPastedText,
        parseSyllabus: parseSchoolSyllabusText,
        getState: function() { return importedState; }
    };

    // Global shortcut
    window.openImportMatrixModal = openImportMatrixModal;
    window.closeImportMatrixModal = closeImportMatrixModal;

    // Tự động khởi tạo khi DOM ready
    if (document.readyState === "loading") {
        document.addEventListener("DOMContentLoaded", createModalElement);
    } else {
        createModalElement();
    }

})(window);
