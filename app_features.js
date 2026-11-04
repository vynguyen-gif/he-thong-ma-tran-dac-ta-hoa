// app_features.js - Tích hợp Ma trận, Bản đặc tả & Ngân hàng Đề chính thức Tổ Hóa Học (THPT Bùi Thị Xuân - Mũi Né, 2025-2026)
// Chuẩn Công văn 7991/BGDĐT-GDTrH - Bộ sách Kết Nối Tri Thức Với Cuộc Sống (Lớp 10, 11, 12)
// =========================================================================================================================

var currentTerm = "GK1";

/* 1. Cấu hình Ma trận & Đặc tả 12 Kỳ kiểm tra (GK1, CK1, GK2, CK2 - Khối 10, 11, 12) */
const KNTT_TERMS_CONFIG = (typeof OFFICIAL_BTX_MATRICES !== "undefined") ? OFFICIAL_BTX_MATRICES : {};

/* Hàm nhận diện và tự động chuyển đổi bảng bị phẳng hóa (từ Word hoặc OCR) thành thẻ HTML bảng chuẩn */
function autoFormatFlattenedTables(text) {
    if (!text || typeof text !== "string" || text.includes("<table")) return text;
    var s = text;

    // 0. Nhận diện & chuyển đổi tự động Bảng dạng nhiều dòng (Multiline Table) từ Word/OCR
    // Khi các ô của bảng bị bóc tách thành các dòng đơn lẻ liên tiếp (thường gặp trong Hướng dẫn giải / Tự luận)
    var headerSignatures = [
        ["Nguyên tử", "Khối lượng (kg)", "Khối lượng (amu)"],
        ["Loại hạt", "Khối lượng (amu)", "Điện tích (e0)"],
        ["Loại hạt", "Khối lượng (amu)", "Điện tích (e₀)"],
        ["Loại hạt", "Khối lượng (amu)", "Điện tích"],
        ["Nguyên tố", "Z", "Cấu hình electron", "Vị trí trong BTH", "Loại nguyên tố"],
        ["Vị trí trong BTH", "Cấu hình electron nguyên tử", "Tên nguyên tố"],
        ["Hợp chất", "Nhiệt độ sôi (°C)", "Hợp chất", "Nhiệt độ sôi (°C)"],
        ["Hợp chất", "Nhiệt độ sôi (oC)", "Hợp chất", "Nhiệt độ sôi (oC)"],
        ["Hợp chất", "Công thức", "Dạng Cn(H2O)m", "Carbohydrate"],
        ["Hợp chất", "Công thức phân tử", "Dạng Cn(H2O)m", "Loại carbohydrate"],
        ["Nguyên tố", "Cấu hình electron", "Số e lớp ngoài cùng", "Vị trí trong BTH"],
        ["Tính chất vật lý", "Ví dụ ứng dụng tương ứng", "Ví dụ kí hiệu hóa học của kim loại phù hợp"],
        ["Tính chất vật lý", "Ví dụ ứng dụng tương ứng", "Kí hiệu hóa học của kim loại phù hợp"],
        ["Nguyên tố", "Số hiệu nguyên tử", "Cấu hình electron", "Số electron hóa trị"],
        ["Kim loại", "K", "Ca", "Fe", "Cu"],
        ["Thông số vật lí", "K", "Ca", "Fe", "Cu"],
        ["Công thức", "Tên gọi thay thế"],
        ["Mẫu thử", "Thuốc thử", "Hiện tượng"],
        ["Mẫu thử", "Thí nghiệm", "Hiện tượng"],
        ["Chất", "Thuốc thử", "Hiện tượng"],
        ["Thí nghiệm", "Nồng độ chất A (M)", "Nồng độ chất B (M)", "Tốc độ phản ứng (M/s)"],
        ["Hạt", "Số electron", "Số neutron", "Số proton"],
        ["Phần tử", "Số electron", "Số neutron", "Số proton"]
    ];

    if (s.includes("\n")) {
        var rawLines = s.split(/\r?\n/);
        var trimmedLines = rawLines.map(function(l) { return l.trim(); }).filter(Boolean);

        for (var idxSig = 0; idxSig < headerSignatures.length; idxSig++) {
            var sig = headerSignatures[idxSig];
            var sigLen = sig.length;
            for (var i = 0; i <= trimmedLines.length - sigLen; i++) {
                var isMatch = true;
                for (var k = 0; k < sigLen; k++) {
                    if (trimmedLines[i + k].toLowerCase() !== sig[k].toLowerCase()) {
                        isMatch = false;
                        break;
                    }
                }
                if (isMatch) {
                    var introLines = trimmedLines.slice(0, i);
                    var afterHeaders = trimmedLines.slice(i + sigLen);
                    var rowItems = [];
                    var outroLines = [];
                    for (var r = 0; r < afterHeaders.length; r++) {
                        var itm = afterHeaders[r];
                        if ((itm.startsWith("(") && itm.includes(")") && itm.length > 40) ||
                            itm.startsWith("+ ") || itm.startsWith("• ") ||
                            itm.startsWith("Nhận xét") || itm.startsWith("Giải thích") ||
                            itm.startsWith("Nhiệt độ sôi của một số")) {
                            outroLines = afterHeaders.slice(r);
                            break;
                        }
                        rowItems.push(itm);
                    }

                    var validCount = Math.floor(rowItems.length / sigLen) * sigLen;
                    if (validCount >= sigLen) {
                        var tableData = rowItems.slice(0, validCount);
                        var leftover = rowItems.slice(validCount).concat(outroLines);

                        var tableHtml = "<table class='chem-table' style='width: 100%; max-width: 680px; border-collapse: collapse; margin: 8px 0; font-size: 10.5pt; border: 1.5px solid #94a3b8;' border='1'><thead><tr>";
                        for (var h = 0; h < sigLen; h++) {
                            tableHtml += "<th style='background-color: #f1f5f9; color: #0f172a; font-weight: bold; padding: 6px 10px; border: 1px solid #94a3b8; text-align: center;'>" + sig[h] + "</th>";
                        }
                        tableHtml += "</tr></thead><tbody>";

                        for (var rowIdx = 0; rowIdx < tableData.length; rowIdx += sigLen) {
                            tableHtml += "<tr>";
                            for (var colIdx = 0; colIdx < sigLen; colIdx++) {
                                var cellVal = tableData[rowIdx + colIdx];
                                // Format lũy thừa cho khối lượng kg nếu có
                                if (/10[-–]\d+/.test(cellVal)) {
                                    cellVal = cellVal.replace(/\b(\d+(?:[.,]\d+)?)\s*[.·x×*]\s*10[-–](\d+)\b/g, "$1 · 10<sup>-$2</sup>")
                                                     .replace(/10[-–](\d+)/g, "10<sup>-$1</sup>");
                                }
                                var align = (colIdx === 0 && !["Kim loại", "Thông số vật lí"].includes(sig[0])) ? "left" : "center";
                                tableHtml += "<td style='padding: 6px 10px; border: 1px solid #94a3b8; text-align: " + align + ";'>" + cellVal + "</td>";
                            }
                            tableHtml += "</tr>";
                        }
                        tableHtml += "</tbody></table>";

                        var res = (introLines.length > 0 ? introLines.join("\n") + "\n" : "") +
                                  tableHtml +
                                  (leftover.length > 0 ? "\n" + leftover.join("\n") : "");
                        return res;
                    }
                }
            }
        }
    }

    // 1. Dạng Bảng Khối Lượng Nguyên Tử inline (Magnesium, Carbon, Oxygen, Beryllium, ...)
    if (s.includes("Nguyên tử") && s.includes("Khối lượng (kg)") && s.includes("Khối lượng (amu)")) {
        var introMatch = s.match(/^(.*?)(?:Nguyên tử\s+Khối lượng\s*\(kg\)\s+Khối lượng\s*\(amu\)\s*)(.*)$/s);
        if (introMatch) {
            var intro = introMatch[1].trim();
            var rest = introMatch[2].trim();
            var parts = rest.split(/(?=[A-Z][a-z]+\s*\([A-Z][a-z]?\))/);
            var rows = [];
            parts.forEach(function(p) {
                var atomMatch = p.trim().match(/^([A-Za-z]+\s*\([A-Za-z0-9]+\))\s*(.*)$/);
                if (atomMatch) {
                    var atomName = atomMatch[1].trim();
                    var vals = atomMatch[2].trim().split(/\s+/).filter(Boolean);
                    var kgVal = "<span style='color: #64748b; font-style: italic;'>...?...</span>";
                    var amuVal = "<span style='color: #64748b; font-style: italic;'>...?...</span>";
                    if (vals.length === 1) {
                        var v = vals[0];
                        if (v.includes("10-") || v.includes("10–") || v.includes("10^") || /[-–]\d+$/.test(v)) {
                            kgVal = v.replace(/\b(\d+(?:[.,]\d+)?)\s*[.·x×*]\s*10[-–](\d+)\b/g, "$1 · 10<sup>-$2</sup>")
                                     .replace(/10[-–](\d+)/g, "10<sup>-$1</sup>");
                        } else {
                            amuVal = v;
                        }
                    } else if (vals.length >= 2) {
                        kgVal = vals[0].replace(/\b(\d+(?:[.,]\d+)?)\s*[.·x×*]\s*10[-–](\d+)\b/g, "$1 · 10<sup>-$2</sup>")
                                       .replace(/10[-–](\d+)/g, "10<sup>-$1</sup>");
                        amuVal = vals[1];
                    }
                    rows.push([atomName, kgVal, amuVal]);
                }
            });
            if (rows.length >= 2) {
                var tbl = "<table class='chem-table' border='1' cellpadding='5' cellspacing='0' style='border-collapse: collapse; width: 100%; max-width: 600px; margin: 8px 0; text-align: center; font-size: 10pt; font-family: \"Times New Roman\", serif;'>" +
                    "<thead><tr style='background-color: #f1f5f9; font-weight: bold; color: #1e3a8a;'>" +
                    "<th style='padding: 6px 8px; border: 1px solid #94a3b8; width: 40%; text-align: left;'>Nguyên tử</th>" +
                    "<th style='padding: 6px 8px; border: 1px solid #94a3b8; width: 30%;'>Khối lượng (kg)</th>" +
                    "<th style='padding: 6px 8px; border: 1px solid #94a3b8; width: 30%;'>Khối lượng (amu)</th>" +
                    "</tr></thead><tbody>";
                rows.forEach(function(r) {
                    tbl += "<tr><td style='padding: 6px 8px; border: 1px solid #cbd5e1; text-align: left;'>" + r[0] + "</td>" +
                        "<td style='padding: 6px 8px; border: 1px solid #cbd5e1;'>" + r[1] + "</td>" +
                        "<td style='padding: 6px 8px; border: 1px solid #cbd5e1;'>" + r[2] + "</td></tr>";
                });
                tbl += "</tbody></table>";
                return (intro ? intro + "\n" : "") + tbl;
            }
        }
    }

    // 2. Dạng Bảng Thí Nghiệm Nhận Biết / Thuốc Thử (Mẫu thử | Thuốc thử | Hiện tượng)
    if ((s.includes("Mẫu thử") || s.includes("Chất")) && s.includes("Thuốc thử") && s.includes("Hiện tượng")) {
        var headerMatch = s.match(/^(.*?)(?:(?:Mẫu thử|Chất)\s+Thuốc thử\s+Hiện tượng\s+)(.*)$/s);
        if (headerMatch) {
            var intro2 = headerMatch[1].trim();
            var body2 = headerMatch[2].trim();
            var sampleRegex = /\b([XYZTW]|Mẫu\s+\d+|Ống\s+\d+)\s+([^XYZTW\n]+?)\s+((?:Tạo|Có|Chuyển|Không|Kết|Hợp|Mất|Quỳ)[^\n\r]+?)(?=\s+[XYZTW]\b|\s+Mẫu\s+\d+|\s+Ống\s+\d+|\s+Các chất|\s+Các dung dịch|\s+Dãy|\s*$)/g;
            var rows2 = [];
            var m2;
            var lastIdx2 = 0;
            while ((m2 = sampleRegex.exec(body2)) !== null) {
                rows2.push([m2[1].trim(), m2[2].trim(), m2[3].trim()]);
                lastIdx2 = sampleRegex.lastIndex;
            }
            if (rows2.length >= 2) {
                var ending2 = body2.slice(lastIdx2).trim();
                var tbl2 = "<table border='1' cellpadding='4' cellspacing='0' style='border-collapse: collapse; width: 100%; max-width: 580px; margin: 8px 0; text-align: center; font-size: 10pt; font-family: \"Times New Roman\", serif;'>" +
                    "<thead><tr style='background-color: #f1f5f9; font-weight: bold; color: #1e3a8a;'>" +
                    "<th style='padding: 6px 8px; border: 1px solid #94a3b8; width: 20%;'>Mẫu thử</th>" +
                    "<th style='padding: 6px 8px; border: 1px solid #94a3b8; width: 45%; text-align: left;'>Thuốc thử</th>" +
                    "<th style='padding: 6px 8px; border: 1px solid #94a3b8; width: 35%; text-align: left;'>Hiện tượng</th>" +
                    "</tr></thead><tbody>";
                rows2.forEach(function(r) {
                    tbl2 += "<tr><td style='padding: 5px 8px; border: 1px solid #cbd5e1; font-weight: bold;'>" + r[0] + "</td>" +
                        "<td style='padding: 5px 8px; border: 1px solid #cbd5e1; text-align: left;'>" + r[1] + "</td>" +
                        "<td style='padding: 5px 8px; border: 1px solid #cbd5e1; text-align: left;'>" + r[2] + "</td></tr>";
                });
                tbl2 += "</tbody></table>";
                return (intro2 ? intro2 + "\n" : "") + tbl2 + (ending2 ? "\n" + ending2 : "");
            }
        }
    }

    // 3. Dạng Bảng Pin Điện Hóa (Pin điện hóa | Sức điện động chuẩn)
    if (s.includes("Pin điện hóa") && s.includes("Sức điện động chuẩn")) {
        var m3 = s.match(/^(.*?)(?:Pin điện hóa\s+([^\n\r]+?)\s+Sức điện động chuẩn\s*(?:\(V\))?\s+([0-9.,\s]+))(.*)$/s);
        if (m3) {
            var intro3 = m3[1].trim();
            var cells3 = m3[2].trim().split(/\s{2,}|\t|\s+(?=[A-Za-z0-9]+\s*[–\-])/).map(function(c) { return c.trim(); }).filter(Boolean);
            var voltages3 = m3[3].trim().split(/\s+/).map(function(v) { return v.trim(); }).filter(Boolean);
            var ending3 = m3[4].trim();
            if (cells3.length > 0 && cells3.length === voltages3.length) {
                var tbl3 = "<table border='1' cellpadding='4' cellspacing='0' style='border-collapse: collapse; width: 100%; max-width: 500px; margin: 8px 0; text-align: center; font-size: 10pt; font-family: \"Times New Roman\", serif;'>" +
                    "<thead><tr style='background-color: #f1f5f9; font-weight: bold; color: #1e3a8a;'>" +
                    "<th style='padding: 6px 8px; border: 1px solid #94a3b8; background-color: #f8fafc;'>Pin điện hóa</th>" +
                    cells3.map(function(c) { return "<td style='padding: 6px 8px; border: 1px solid #cbd5e1; font-weight: bold;'>" + c + "</td>"; }).join("") +
                    "</tr></thead><tbody><tr>" +
                    "<th style='padding: 6px 8px; border: 1px solid #94a3b8; background-color: #f8fafc;'>Sức điện động chuẩn (V)</th>" +
                    voltages3.map(function(v) { return "<td style='padding: 6px 8px; border: 1px solid #cbd5e1;'>" + v + "</td>"; }).join("") +
                    "</tr></tbody></table>";
                return (intro3 ? intro3 + "\n" : "") + tbl3 + (ending3 ? "\n" + ending3 : "");
            }
        }
    }

    // 4. Dạng Bảng Hạt Cơ Bản (Loại hạt | Khối lượng (amu) | Điện tích (e0))
    if (s.includes("Loại hạt") && s.includes("Khối lượng (amu)") && s.includes("Điện tích")) {
        var m4 = s.match(/^(.*?)(?:Loại hạt\s+Khối lượng\s*\(amu\)\s+Điện tích\s*(?:\(e0\))?\s*)(?:(neutron|electron|proton)\s+([0-9.,]+)\s+([+\-–]?\d+)\s*)+(.*)$/si);
        if (m4) {
            var intro4 = m4[1].trim();
            var ending4 = m4[m4.length - 1].trim();
            var tbl4 = "<table border='1' cellpadding='4' cellspacing='0' style='border-collapse: collapse; width: 100%; max-width: 480px; margin: 8px 0; text-align: center; font-size: 10pt; font-family: \"Times New Roman\", serif;'>" +
                "<thead><tr style='background-color: #f1f5f9; font-weight: bold; color: #1e3a8a;'>" +
                "<th style='padding: 6px 8px; border: 1px solid #94a3b8;'>Loại hạt</th>" +
                "<th style='padding: 6px 8px; border: 1px solid #94a3b8;'>Khối lượng (amu)</th>" +
                "<th style='padding: 6px 8px; border: 1px solid #94a3b8;'>Điện tích (e<sub>0</sub>)</th>" +
                "</tr></thead><tbody>" +
                "<tr><td style='padding: 5px 8px; border: 1px solid #cbd5e1;'>neutron</td><td style='padding: 5px 8px; border: 1px solid #cbd5e1;'>1</td><td style='padding: 5px 8px; border: 1px solid #cbd5e1;'>0</td></tr>" +
                "<tr><td style='padding: 5px 8px; border: 1px solid #cbd5e1;'>electron</td><td style='padding: 5px 8px; border: 1px solid #cbd5e1;'>0,00055</td><td style='padding: 5px 8px; border: 1px solid #cbd5e1;'>–1</td></tr>" +
                "<tr><td style='padding: 5px 8px; border: 1px solid #cbd5e1;'>proton</td><td style='padding: 5px 8px; border: 1px solid #cbd5e1;'>1</td><td style='padding: 5px 8px; border: 1px solid #cbd5e1;'>+1</td></tr>" +
                "</tbody></table>";
            return (intro4 ? intro4 + "\n" : "") + tbl4 + (ending4 ? "\n" + ending4 : "");
        }
    }

    return s;
}

/* Hàm quy đổi chuẩn hóa tên Chương / Chuyên đề theo đúng chương trình GDPT 2018 (SGK Kết Nối Tri Thức) */
function getCanonicalChapter(str, grade) {
    if (!str) {
        if (grade === "11") return "Chương 1: Cân bằng hóa học";
        if (grade === "12") return "Chương 1: Ester - Lipid";
        return "Chương 1: Cấu tạo nguyên tử";
    }
    var s = String(str).toLowerCase().trim();

    // 1. Phân loại chuẩn theo Khối 10
    if (grade === "10" || (!grade && (s.includes("nguyên tử") || s.includes("hạt nhân") || s.includes("tuần hoàn") || s.includes("liên kết") || s.includes("oxi hóa") || s.includes("enthalpy") || s.includes("tốc độ") || (s.includes("halogen") && !s.includes("dẫn xuất"))))) {
        if (s.includes("nguyên tử") || s.includes("hạt nhân") || s.includes("lớp vỏ") || s.includes("cấu tạo nguyên tử")) return "Chương 1: Cấu tạo nguyên tử";
        if (s.includes("bảng tuần hoàn") || s.includes("tuần hoàn") || s.includes("định luật")) return "Chương 2: Bảng tuần hoàn các nguyên tố hóa học";
        if (s.includes("liên kết") || s.includes("octet") || s.includes("ionic")) return "Chương 3: Liên kết hóa học";
        if (s.includes("oxi hóa") || s.includes("oxit hóa") || s.includes("khử")) return "Chương 4: Phản ứng oxi hóa - khử";
        if (s.includes("năng lượng") || s.includes("enthalpy") || s.includes("nhiệt")) return "Chương 5: Năng lượng hóa học";
        if (s.includes("tốc độ") || s.includes("van't hoff")) return "Chương 6: Tốc độ phản ứng hóa học";
        if (s.includes("halogen") || s.includes("viia") || s.includes("halide")) return "Chương 7: Nhóm Halogen (Nhóm VIIA)";
        return "Chương 1: Cấu tạo nguyên tử";
    }

    // 2. Phân loại chuẩn theo Khối 11
    if (grade === "11" || (!grade && (s.includes("cân bằng") || (s.includes("nitrogen") && !s.includes("chứa nitrogen")) || (s.includes("hữu cơ") && s.includes("đại cương")) || s.includes("hydrocarbon") || s.includes("alcohol") || s.includes("phenol") || s.includes("dẫn xuất") || s.includes("carbonyl") || s.includes("carboxylic")))) {
        if (s.includes("alcohol") || s.includes("phenol") || s.includes("dẫn xuất halogen") || s.includes("ancol")) return "Chương 5: Dẫn xuất halogen - Alcohol - Phenol";
        if (s.includes("carbonyl") || s.includes("carboxylic") || s.includes("aldehyde") || s.includes("ketone")) return "Chương 6: Hợp chất carbonyl - Carboxylic acid";
        if (s.includes("hydrocarbon") || s.includes("alkane") || s.includes("alkene") || s.includes("alkyne") || s.includes("arene")) return "Chương 4: Hydrocarbon";
        if (s.includes("hữu cơ") && (s.includes("đại cương") || s.includes("tách biệt") || s.includes("nhóm chức") || s.includes("phổ"))) return "Chương 3: Đại cương về hóa học hữu cơ";
        if (s.includes("nitrogen") || s.includes("sulfur") || s.includes("ammonia") || s.includes("hno3") || s.includes("h2so4")) return "Chương 2: Nitrogen và sulfur";
        if (s.includes("cân bằng") || s.includes("điện li") || s.includes("chuẩn độ") || s.includes("brønsted")) return "Chương 1: Cân bằng hóa học";
        return "Chương 1: Cân bằng hóa học";
    }

    // 3. Phân loại chuẩn theo Khối 12
    if (grade === "12" || (!grade && (s.includes("ester") || s.includes("lipid") || s.includes("carbohydrate") || s.includes("chứa nitrogen") || s.includes("polymer") || s.includes("pin") || s.includes("điện phân") || (s.includes("đại cương") && s.includes("kim loại")) || s.includes("nhóm ia") || s.includes("nhóm iia") || s.includes("chuyển tiếp") || s.includes("phức chất")))) {
        if (s.includes("chuyển tiếp") || s.includes("phức chất") || s.includes("phối tử") || s.includes("aqua")) return "Chương 8: Kim loại chuyển tiếp dãy thứ nhất & Phức chất";
        if (s.includes("nhóm ia") || s.includes("nhóm iia") || s.includes("kim loại kiềm") || s.includes("nước cứng")) return "Chương 7: Nguyên tố nhóm IA và nhóm IIA";
        if (s.includes("đại cương") && s.includes("kim loại") || s.includes("tách kim loại") || s.includes("hợp kim") || s.includes("ăn mòn")) return "Chương 6: Đại cương về kim loại";
        if (s.includes("pin") || s.includes("điện phân") || s.includes("thế điện cực") || s.includes("galvani")) return "Chương 5: Pin điện và điện phân";
        if (s.includes("polymer") || s.includes("chất dẻo") || s.includes("cao su") || s.includes("tơ")) return "Chương 4: Polymer";
        if (s.includes("chứa nitrogen") || s.includes("nitrogen") || s.includes("amine") || s.includes("amino acid") || s.includes("peptide") || s.includes("protein")) return "Chương 3: Hợp chất chứa nitrogen";
        if (s.includes("carbohydrate") || s.includes("glucose") || s.includes("fructose") || s.includes("tinh bột") || s.includes("cellulose")) return "Chương 2: Carbohydrate";
        if (s.includes("ester") || s.includes("lipid") || s.includes("chất béo") || s.includes("xà phòng")) return "Chương 1: Ester - Lipid";
        return "Chương 1: Ester - Lipid";
    }

    return str;
}

/* Hàm chuyển đổi ký hiệu ô orbital dạng ngoặc vuông [↑], [↑↓], [ ] thành các ô vuông (AO) chuẩn SGK */
function formatOrbitalBoxes(text) {
    if (!text || typeof text !== "string") return text;

    // Khớp cụm ngoặc vuông orbital: ví dụ [↑↑][  ][  ], [↑ ][↑ ][↑ ], [↑↓], [↑], [  ], [ ]
    var aoGroupRegex = /(?:\[[↑↓⥮\s]{0,4}\]\s*)+/g;

    return text.replace(aoGroupRegex, function(match) {
        var singleBoxRegex = /\[([↑↓⥮\s]{0,4})\]/g;
        var boxes = [];
        var m;
        var hasArrow = false;
        while ((m = singleBoxRegex.exec(match)) !== null) {
            var raw = m[1].replace(/\s+/g, "");
            if (raw.length > 0) hasArrow = true;
            boxes.push(raw);
        }

        if (boxes.length === 0) {
            return match;
        }

        // Bỏ qua nếu là 1 ô đơn rỗng trong câu văn dài không liên quan tới orbital (tránh nhầm chỗ điền từ)
        if (!hasArrow && boxes.length === 1 && text.length > 30 && !/(?:orbital|ao|ô lượng tử|electron|phân lớp|lớp vỏ|trạng thái|cơ bản|pauli|hund)/i.test(text)) {
            return match;
        }

        var cells = boxes.map(function(b) {
            var content = "&nbsp;&nbsp;";
            if (b === "↑" || b === "↑ ") content = "&#8593;&nbsp;";
            else if (b === "↓" || b === "↓ ") content = "&#8595;&nbsp;";
            else if (b === "↑↓" || b === "⥮") content = "&#8593;&#8595;";
            else if (b === "↑↑") content = "&#8593;&#8593;";
            else if (b === "↓↓") content = "&#8595;&#8595;";
            else if (b) content = b;

            return "<td style='border: 1.5pt solid #000; width: 22px; height: 22px; min-width: 22px; max-width: 22px; text-align: center; vertical-align: middle; font-size: 11pt; font-family: \"Times New Roman\", serif; font-weight: bold; line-height: 1; padding: 0; background-color: #ffffff; color: #000000;'>" + content + "</td>";
        }).join("");

        var totalWidth = boxes.length * 24;
        return "<table class='ao-boxes' width='" + totalWidth + "' style='display: inline-table; border-collapse: collapse; vertical-align: middle; margin: 0 4px; border: 1.5pt solid #000; width: " + totalWidth + "px;'><tbody><tr>" + cells + "</tr></tbody></table>";
    });
}

/* Bảng số hiệu nguyên tử (Z) chuẩn IUPAC của các nguyên tố hóa học */
const ATOMIC_NUMBERS_MAP = {
    H: 1, He: 2, Li: 3, Be: 4, B: 5, C: 6, N: 7, O: 8, F: 9, Ne: 10,
    Na: 11, Mg: 12, Al: 13, Si: 14, P: 15, S: 16, Cl: 17, Ar: 18, K: 19, Ca: 20,
    Sc: 21, Ti: 22, V: 23, Cr: 24, Mn: 25, Fe: 26, Co: 27, Ni: 28, Cu: 29, Zn: 30,
    Ga: 31, Ge: 32, As: 33, Se: 34, Br: 35, Kr: 36, Rb: 37, Sr: 38, Y: 39, Zr: 40,
    Nb: 41, Mo: 42, Tc: 43, Ru: 44, Rh: 45, Pd: 46, Ag: 47, Cd: 48, In: 49, Sn: 50,
    Sb: 51, Te: 52, I: 53, Xe: 54, Cs: 55, Ba: 56, La: 57, Ce: 58, Pr: 59, Nd: 60,
    Pm: 61, Sm: 62, Eu: 63, Gd: 64, Tb: 65, Dy: 66, Ho: 67, Er: 68, Tm: 69, Yb: 70,
    Lu: 71, Hf: 72, Ta: 73, W: 74, Re: 75, Os: 76, Ir: 77, Pt: 78, Au: 79, Hg: 80,
    Tl: 81, Pb: 82, Bi: 83, Po: 84, At: 85, Rn: 86, Fr: 87, Ra: 88, Ac: 89, Th: 90,
    Pa: 91, U: 92, Np: 93, Pu: 94
};

/* Tự động nhận diện và khôi phục ký hiệu nguyên tử / đồng vị bị phẳng hóa (3919K, 6329Cu, 3581Br...) */
function autoFixFlattenedIsotopes(text) {
    if (!text || typeof text !== "string") return text;
    var s = text;

    // 1. Dạng dính liền 3 đến 5 chữ số trước ký hiệu nguyên tố: 3919K, 6329Cu, 6529Cu, 3581Br, 1656Fe...
    s = s.replace(/\b(\d{3,5})([A-Z][a-z]?)\b/g, function(match, digits, elem) {
        var zKnown = ATOMIC_NUMBERS_MAP[elem];
        if (!zKnown) return match;

        // Bỏ qua các đơn vị như 298K, 500K (nhiệt độ Kelvin), 5500C (nhiệt độ Celsius), 001M (nồng độ M)
        if ((elem === "K" || elem === "C") && (digits === "298" || digits === "273" || digits === "500" || digits === "5500" || digits === "4500" || digits === "100" || digits === "250" || digits === "600")) {
            return match;
        }
        if (elem === "M" || elem === "V" || elem === "L" || elem === "J") return match;

        var zStr = String(zKnown);

        // Trường hợp 1: Số hiệu Z nằm ở cuối chuỗi số: ví dụ 3919K (A=39, Z=19), 6329Cu (A=63, Z=29)
        if (digits.endsWith(zStr)) {
            var aStr = digits.substring(0, digits.length - zStr.length);
            var aNum = parseInt(aStr, 10);
            if (aNum > 0 && aNum >= zKnown && aNum <= zKnown * 3) {
                return "\\(^{" + aStr + "}_{" + zStr + "}" + elem + "\\)";
            }
        }

        // Trường hợp 2: Số hiệu Z nằm ở đầu chuỗi số: ví dụ 3581Br (Z=35, A=81)
        if (digits.startsWith(zStr)) {
            var aStr2 = digits.substring(zStr.length);
            var aNum2 = parseInt(aStr2, 10);
            if (aNum2 > 0 && aNum2 >= zKnown && aNum2 <= zKnown * 3) {
                return "\\(^{" + aStr2 + "}_{" + zStr + "}" + elem + "\\)";
            }
        }

        // Trường hợp đặc biệt OCR lỗi: 1656Fe (thực chất là 2656Fe, Z=26, A=56)
        if (elem === "Fe" && (digits === "1656" || digits === "2656")) {
            return "\\(^{56}_{26}Fe\\)";
        }

        return match;
    });

    // 2. Dạng số khối dính liền trước ký hiệu nguyên tố sau từ "đồng vị":
    // ví dụ: "đồng vị 63Cu và 65Cu" -> "đồng vị ^{63}Cu và ^{65}Cu"
    s = s.replace(/((?:đồng vị|dạng đồng vị|các đồng vị)\s+(?:bền\s+)?(?:của\s+)?(?:trong đó\s+)?(?:là\s+)?|,\s*|và\s+)(\d{1,3})([A-Z][a-z]?)(?=[\s,.:;)\]\}]|$)/g, function(match, prefix, aStr, elem) {
        var zKnown = ATOMIC_NUMBERS_MAP[elem];
        if (!zKnown) return match;
        if (elem === "K" && (aStr === "298" || aStr === "273" || aStr === "500")) return match;
        if (elem === "C" && (aStr === "100" || aStr === "250" || aStr === "500")) return match;
        var aNum = parseInt(aStr, 10);
        if (aNum >= 1 && aNum <= zKnown * 3 + 5) {
            return prefix + "\\(^{" + aStr + "}" + elem + "\\)";
        }
        return match;
    });

    return s;
}

/* Hàm chuẩn hóa điện tích ion (đa nguyên tử & đơn nguyên tử: Li+, Na+, N3-, O2-, Mg2+, Al3+, Fe2+, Cu2+, SO4^2-, PO4^3-...) */
function formatChemicalIons(text) {
    if (!text) return "";
    var s = String(text);

    // Chuẩn hóa dấu gạch ngang en-dash / em-dash thành dấu trừ chuẩn
    s = s.replace(/([0-9A-Za-z\)\]\>])[\u2013\u2014]/g, function(_, g) { return g + "-"; });

    // 1. Ion đa nguyên tử:
    var polyMap = [
        { re: /\bSO4\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "SO<sub>4</sub>" },
        { re: /\bSO3\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "SO<sub>3</sub>" },
        { re: /\bCO3\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "CO<sub>3</sub>" },
        { re: /\bNO3\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "NO<sub>3</sub>" },
        { re: /\bNO2\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "NO<sub>2</sub>" },
        { re: /\bPO4\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "PO<sub>4</sub>" },
        { re: /\bHPO4\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "HPO<sub>4</sub>" },
        { re: /\bH2PO4\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "H<sub>2</sub>PO<sub>4</sub>" },
        { re: /\bHSO4\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "HSO<sub>4</sub>" },
        { re: /\bHSO3\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "HSO<sub>3</sub>" },
        { re: /\bHCO3\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "HCO<sub>3</sub>" },
        { re: /\bOH\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "OH" },
        { re: /\bClO4\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "ClO<sub>4</sub>" },
        { re: /\bClO3\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "ClO<sub>3</sub>" },
        { re: /\bClO2\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "ClO<sub>2</sub>" },
        { re: /\bClO\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "ClO" },
        { re: /\bMnO4\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "MnO<sub>4</sub>" },
        { re: /\bCrO4\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "CrO<sub>4</sub>" },
        { re: /\bCr2O7\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "Cr<sub>2</sub>O<sub>7</sub>" },
        { re: /\bAlO2\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "AlO<sub>2</sub>" },
        { re: /\bZnO2\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "ZnO<sub>2</sub>" },
        { re: /\bCH3COO\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "CH<sub>3</sub>COO" },
        { re: /\bNH4\s*(?:\^\{?|\^)?([1-4]?[+]|[+][1-4]?)\}?/g, b: "NH<sub>4</sub>" },
        { re: /\bH3O\s*(?:\^\{?|\^)?([1-4]?[+]|[+][1-4]?)\}?/g, b: "H<sub>3</sub>O" }
    ];

    polyMap.forEach(function(item) {
        s = s.replace(item.re, function(match, rawCharge) {
            var ch = (rawCharge || "-").replace(/[\u2013\u2014]/g, "-").trim();
            if (/^[+\-]/.test(ch)) ch = ch.slice(1) + ch[0];
            if (ch === "1+" || ch === "+1") ch = "+";
            if (ch === "1-" || ch === "-1") ch = "-";
            return item.b + "<sup>" + ch + "</sup>";
        });
    });

    // 2. Ion đơn nguyên tử & ion kim loại / phi kim tổng quát (M, X, R):
    var allElemSyms = "He|Li|Be|Ne|Na|Mg|Al|Si|Cl|Ar|Ca|Sc|Ti|Cr|Mn|Fe|Co|Ni|Cu|Zn|Ga|Ge|As|Se|Br|Kr|Rb|Sr|Zr|Nb|Mo|Tc|Ru|Rh|Pd|Ag|Cd|In|Sn|Sb|Te|Xe|Cs|Ba|La|Ce|Pr|Nd|Sm|Eu|Gd|Tb|Dy|Ho|Er|Tm|Yb|Lu|Hf|Ta|Re|Os|Ir|Pt|Au|Hg|Tl|Pb|Bi|Po|At|Rn|Fr|Ra|Ac|Th|Pa|H|B|C|N|O|F|P|S|K|V|Y|W|U|I|M|X|R";
    var monoIonRegex = new RegExp("\\b(" + allElemSyms + ")(?:\\^\\{?([0-9]*[+\\-\\u2013][0-9]*)\\}?|\\^?([2-4]?[+\\-\\u2013]|[+\\-\\u2013][2-4]?))(?=[\\s,.:;)\\]\\/\\?\\<]|$)", "g");

    s = s.replace(monoIonRegex, function(match, elem, caretCharge, directCharge) {
        var ch = (caretCharge || directCharge || "").replace(/[\u2013\u2014]/g, "-").trim();
        if (!ch) return match;
        if (/^[+\-]/.test(ch)) {
            var sign = ch[0];
            var num = ch.slice(1);
            ch = num + sign;
        }
        if (ch === "1+" || ch === "+1") ch = "+";
        if (ch === "1-" || ch === "-1") ch = "-";
        return elem + "<sup>" + ch + "</sup>";
    });

    return s;
}

/* =========================================================================================
   HÀM CHUẨN HÓA LỜI GIẢI CHI TIẾT & LOẠI BỎ TOÀN BỘ TÊN TÁC GIẢ KHỎI HỆ THỐNG
   ========================================================================================= */
function stripAuthorAndPromotion(text) {
    if (!text || typeof text !== "string") return "";
    var res = text;
    res = res.replace(/Ths\.?\s*TRẦN THANH BÌNH[^\n]*/gi, "")
             .replace(/CHUYÊN VIẾT VÀ CHUYỂN GIAO TÀI LIỆU[^\n]*/gi, "")
             .replace(/Thầy cô có nhu cầu[^\n]*/gi, "")
             .replace(/zalo:?[^\n]*/gi, "")
             .replace(/Link đọc thử:?[^\n]*/gi, "")
             .replace(/https?:\/\/[^\s]+/gi, "")
             .replace(/_{3,}/g, "")
             .replace(/Lời giải chi tiết theo chuẩn\s*(?:Thầy|Cô)?[^\n\.\;]*/gi, "")
             .replace(/Hướng dẫn giải(?: 4 ý a, b, c, d)? chuẩn\s*(?:Thầy|Cô)?[^\n\.\;]*/gi, "")
             .replace(/Lời giải theo tài liệu\s*(?:Thầy|Cô)?[^\n\.\;]*/gi, "")
             .replace(/Lời giải chi tiết theo tài liệu giảng dạy\.?/gi, "")
             .replace(/Lời giải chi tiết theo chuẩn\.?/gi, "")
             .replace(/Lời giải chi tiết\.?/gi, "")
             .replace(/(?:Thầy|Cô)\s+(?:Cao Quốc Trung|Trần Thanh Bình|Nguyễn Hiếu|[A-ZÀ-Ỵ][a-zà-ỹ]+(?:\s+[A-ZÀ-Ỵ][a-zà-ỹ]+)*)/gu, "")
             .replace(/Cao Quốc Trung|Trần Thanh Bình|Nguyễn Hiếu/gi, "")
             .replace(/\([^\)]*(?:H10_|CHUONG|GIAO VIEN|BẮC GIANG|DE \d|TAP HUAN)[^\)]*\)/gi, "")
             .replace(/\s{2,}/g, " ")
             .trim();
    res = res.replace(/^[\s\:\-\.\,]+/, "").replace(/[\s\:\-\,]+$/, "").trim();
    return res;
}

function formatDetailedExplanation(exp, q) {
    var text = stripAuthorAndPromotion(exp);

    // Nếu văn bản sau khi lọc vẫn còn nội dung giải thích hóa học thực chất (> 20 ký tự)
    if (text.length > 20 && !/^(?:hướng dẫn giải|đáp án|đáp số|lời giải|phương án)[\s\:\.]*$/i.test(text)) {
        return text;
    }

    // 2. Tự động sinh lời giải chi tiết theo bản chất hóa học & nội dung câu hỏi
    if (q) {
        var c = (q.content || "").toLowerCase();
        if (c.includes("hạt không mang điện") && c.includes("nguyên tử")) {
            return "Trong cấu tạo nguyên tử: hạt proton mang điện tích dương (+1), hạt electron mang điện tích âm (-1), còn hạt neutron không mang điện. Do đó, hạt không mang điện là neutron.";
        }
        if (c.includes("hạt mang điện tích dương") || c.includes("hạt mang điện dương")) {
            return "Trong hạt nhân nguyên tử, hạt mang điện tích dương là proton (+1), hạt electron mang điện tích âm (-1) ở lớp vỏ nguyên tử, còn hạt neutron không mang điện.";
        }
        if (c.includes("hạt mang điện tích âm") || c.includes("hạt mang điện âm")) {
            return "Hạt mang điện tích âm trong nguyên tử là electron (-1), chuyển động rất nhanh xung quanh hạt nhân trong không gian rỗng của nguyên tử.";
        }
        if (c.includes("cấu tạo nguyên tử") || c.includes("thành phần nguyên tử")) {
            return "Nguyên tử cấu tạo gồm 2 phần: Hạt nhân ở trung tâm (chứa proton mang điện dương và neutron không mang điện) và lớp vỏ gồm các electron mang điện âm chuyển động xung quanh.";
        }
        if (c.includes("đồng vị") || c.includes("isotope")) {
            return "Đồng vị là những nguyên tử có cùng số proton (cùng số hiệu nguyên tử Z) nhưng khác nhau về số neutron (dẫn đến số khối A khác nhau).";
        }
        if (c.includes("cấu hình electron") || c.includes("electron lớp ngoài cùng")) {
            return "Cấu hình electron được viết theo thứ tự mức năng lượng tăng dần: 1s 2s 2p 3s 3p 4s 3d 4p... Số electron lớp ngoài cùng quyết định tính chất hóa học đặc trưng của nguyên tố.";
        }
        if (c.includes("quy tắc octet") || c.includes("bát tử")) {
            return "Quy tắc octet: Trong quá trình hình thành liên kết hóa học, nguyên tử các nguyên tố nhóm A có xu hướng nhường, nhận hoặc góp chung electron để đạt cấu hình electron bền vững với 8 electron ở lớp ngoài cùng.";
        }
        if (c.includes("bán kính nguyên tử")) {
            return "Quy luật biến đổi tuần hoàn: Trong cùng một chu kỳ từ trái sang phải bán kính nguyên tử giảm dần; trong cùng một nhóm A từ trên xuống dưới bán kính nguyên tử tăng dần.";
        }
        if (c.includes("độ âm điện")) {
            return "Trong một chu kỳ từ trái sang phải, độ âm điện tăng dần theo chiều tăng điện tích hạt nhân; trong cùng nhóm A từ trên xuống dưới, độ âm điện giảm dần.";
        }
        if (c.includes("liên kết ion") || c.includes("liên kết ionic")) {
            return "Liên kết ion là liên kết được hình thành bởi lực hút tĩnh điện giữa các ion mang điện tích trái dấu (cation kim loại và anion phi kim).";
        }
        if (c.includes("liên kết cộng hóa trị")) {
            return "Liên kết cộng hóa trị là liên kết được hình thành giữa hai nguyên tử bằng một hoặc nhiều cặp electron dùng chung.";
        }
        if (c.includes("liên kết hydrogen")) {
            return "Liên kết hydrogen là liên kết tĩnh điện yếu giữa nguyên tử H linh động với nguyên tử có độ âm điện lớn khác còn cặp electron tự do. Liên kết hydrogen làm tăng nhiệt độ sôi và độ tan.";
        }
        if (c.includes("số oxi hóa") || c.includes("oxi hóa - khử")) {
            return "Chất khử là chất nhường electron (số oxi hóa tăng sau phản ứng); chất oxi hóa là chất nhận electron (số oxi hóa giảm sau phản ứng).";
        }
        if (c.includes("enthalpy") || c.includes("nhiệt tạo thành")) {
            return "Biến thiên enthalpy chuẩn của phản ứng: Δr H°₂₉₈ = ∑Δf H°₂₉₈ (sản phẩm) - ∑Δf H°₂₉₈ (chất đầu). Nếu Δr H°₂₉₈ < 0 là phản ứng tỏa nhiệt; nếu Δr H°₂₉₈ > 0 là phản ứng thu nhiệt.";
        }
        if (c.includes("tốc độ phản ứng")) {
            return "Các yếu tố ảnh hưởng đến tốc độ phản ứng: nồng độ chất phản ứng, nhiệt độ, áp suất (chất khí), diện tích tiếp xúc và chất xúc tác.";
        }
        if (c.includes("halogen")) {
            return "Tính oxi hóa của các đơn chất halogen giảm dần theo thứ tự: F₂ > Cl₂ > Br₂ > I₂. Halogen mạnh hơn đẩy được halogen yếu hơn ra khỏi dung dịch muối halide.";
        }
        if (c.includes("cân bằng hóa học") || c.includes("le chatelier")) {
            return "Theo nguyên lý Le Chatelier: Khi thay đổi một trong các điều kiện (nhiệt độ, áp suất, nồng độ), cân bằng sẽ chuyển dịch theo chiều làm giảm tác động của sự thay đổi đó.";
        }
        if (c.includes("ester") || c.includes("xà phòng hóa")) {
            return "Ester no đơn chức mạch hở có CTPT CₙH₂ₙO₂ (n ≥ 2). Phản ứng xà phòng hóa trong dung dịch kiềm là phản ứng một chiều tạo muối carboxylate và alcohol.";
        }
        if (c.includes("carbohydrate") || c.includes("glucose")) {
            return "Phân loại carbohydrate: Monosaccharide (glucose, fructose); Disaccharide (saccharose); Polysaccharide (tinh bột, cellulose). Glucose có nhóm -CHO nên tráng bạc được.";
        }
        if (c.includes("amine") || c.includes("amino acid")) {
            return "Amine có tính base do nguyên tử N còn cặp electron tự do. Amino acid là hợp chất tạp chức chứa nhóm -NH₂ và -COOH, có tính chất lưỡng tính.";
        }
        if (c.includes("pin điện") || c.includes("điện phân")) {
            return "Trong pin điện hóa: Anode là cực âm (oxi hóa), Cathode là cực dương (khử). Trong bình điện phân: Anode là cực dương (oxi hóa), Cathode là cực âm (khử).";
        }

        if (q.part === 1) {
            var letter = String.fromCharCode(65 + (q.correct || 0));
            var optVal = (q.options && q.options[q.correct || 0]) ? (" (" + q.options[q.correct || 0] + ")") : "";
            return "Phương án đúng là " + letter + optVal + ". Dựa trên lý thuyết và tính chất hóa học chuẩn SGK Hóa học, nội dung phương án " + letter + " hoàn toàn chính xác.";
        }
        if (q.part === 2) {
            return "Hướng dẫn phân tích từng mệnh đề:\n- Đối chiếu dữ kiện đề bài với quy luật và bản chất hóa học của phản ứng.\n- Xét tính đúng/sai của từng nhận định a, b, c, d dựa trên cơ sở lý thuyết và phương trình phản ứng tương ứng.";
        }
        if (q.part === 3) {
            return "Đáp số chính xác: " + (q.answer || "") + ". Áp dụng định luật bảo toàn khối lượng, bảo toàn electron hoặc công thức hóa học tương ứng để tính toán ra kết quả.";
        }
        if (q.part === 4) {
            return "Các bước giải chi tiết:\n1. Phân tích hiện tượng và viết phương trình hóa học diễn ra.\n2. Lập các biểu thức tính toán hoặc hệ phương trình liên hệ số mol.\n3. Tính toán và kết luận đáp số chuẩn theo yêu cầu đề bài.";
        }
    }

    return "Phân tích chi tiết dựa trên lý thuyết và bản chất hóa học theo chương trình SGK Hóa học.";
}

/* Hàm tự động chuẩn hóa chất lượng câu hỏi khi nạp vào hệ thống */
function autoSanitizeQuestion(q) {
    if (!q) return q;

    function sanitizeText(str) {
        if (!str || typeof str !== "string") return str;
        var s = str;
        // 1. Sửa lỗi chính tả & phông chữ phổ biến
        s = s.replace(/\bnững\b/gi, "những")
             .replace(/\bfrucrose\b/gi, "fructose")
             .replace(/\bright\d+\b/gi, ""); // Xóa mã neo hình vẽ Word

        // 2. Tự động chuyển đổi bảng bị phẳng hóa thành HTML bảng
        s = autoFormatFlattenedTables(s);

        // 3. Chuẩn hóa ô orbital thành các ô vuông liên kết chuẩn SGK
        s = formatOrbitalBoxes(s);

        // 4. Chuẩn hóa ký hiệu nguyên tử / đồng vị bị phẳng hóa (3919K, 6329Cu, 3581Br...)
        s = autoFixFlattenedIsotopes(s);

        // 4.5 Chuẩn hóa điện tích ion (Li+, Na+, N3-, O2-, Mg2+, Al3+, Fe2+, Cu2+, SO4^2-, PO4^3-...)
        s = formatChemicalIons(s);

        // 5. Chuẩn hóa lũy thừa và ký hiệu khoa học
        if (!s.includes("<table") || s.includes("class='ao-boxes'") || s.includes("ao-group")) {
            s = s.replace(/\b(\d+(?:[.,]\d+)?)\s*[.·x×*]\s*10[-–](\d+)\b/g, "$1 · 10<sup>-$2</sup>")
                 .replace(/\b(\d+(?:[.,]\d+)?)\s*[.·x×*]\s*10\^([+-]?\d+)\b/g, "$1 · 10<sup>$2</sup>")
                 .replace(/\b10[-–](\d+)\b(?=\s*(?:kg|amu|C|mol|gam|g|M)\b)/g, "10<sup>-$1</sup>")
                 .replace(/1 tỉ\s*\(109\)/g, "1 tỉ (10<sup>9</sup>)")
                 .replace(/104\s*lần/g, "10<sup>4</sup> lần");
        }
        return s;
    }

    // Tự động sửa các câu hỏi nguyên tử bị lệch đáp án
    var cLower = (q.content || "").toLowerCase();
    if (cLower.includes("hạt không mang điện là") && cLower.includes("trong nguyên tử") && Array.isArray(q.options)) {
        var nIdx = q.options.findIndex(function(opt) { return opt.toLowerCase().includes("neutron"); });
        if (nIdx !== -1) {
            q.correct = nIdx;
        }
    }

    // Xóa tên tác giả khỏi trường nguồn
    if (q.source && /(?:Thầy|Cô|Cao Quốc Trung|Trần Thanh Bình|Nguyễn Hiếu)/i.test(q.source)) {
        q.source = "Ngân hàng câu hỏi SGK Kết Nối Tri Thức (KNTT) & CV 7991";
    }

    if (q.content) q.content = sanitizeText(q.content);
    q.explain = formatDetailedExplanation(q.explain, q);
    if (q.solution) q.solution = formatDetailedExplanation(q.solution, q);

    if (q.part === 1 && Array.isArray(q.options)) {
        q.options = q.options.map(function(opt) {
            return sanitizeText(String(opt).trim().replace(/^[A-D][\.\:\)]\s*/i, ""));
        });
    }

    if (q.part === 2 && Array.isArray(q.statements)) {
        q.statements = q.statements.map(function(st) {
            if (st && st.text) st.text = sanitizeText(st.text);
            return st;
        });
    }

    // 4. Chuẩn hóa và gộp thống nhất tên Chương / Chuyên đề
    var canonicalChapter = getCanonicalChapter(q.chapter || q.topic, q.grade);
    q.chapter = canonicalChapter;
    if (!q.topic || q.topic.includes("Chủ đề") || q.topic.includes("Phần") || q.topic.toLowerCase().includes("nguyên tử") || q.topic === canonicalChapter) {
        q.topic = canonicalChapter;
    }

    return q;
}

/* Đồng bộ toàn bộ câu hỏi từ Đề chính thức của Tổ vào QUESTION_BANK ngay khi tải trang */
function initOfficialBTXSystem() {
    if (typeof OFFICIAL_BTX_EXAMS !== "undefined") {
        ["10", "11", "12"].forEach(function(g) {
            if (!QUESTION_BANK[g]) QUESTION_BANK[g] = [];
            var existingIds = new Set(QUESTION_BANK[g].map(function(q) { return q.id; }));
            var terms = ["GK1", "CK1", "GK2", "CK2"];
            var officialList = [];
            terms.forEach(function(t) {
                if (OFFICIAL_BTX_EXAMS[g] && OFFICIAL_BTX_EXAMS[g][t]) {
                    OFFICIAL_BTX_EXAMS[g][t].forEach(function(q) {
                        var copy = JSON.parse(JSON.stringify(q));
                        copy.term = t;
                        if (!existingIds.has(copy.id)) {
                            officialList.push(copy);
                            existingIds.add(copy.id);
                        }
                    });
                }
            });
            // Đưa bộ câu hỏi chính thức của Tổ lên đầu ngân hàng và tự động chuẩn hóa
            QUESTION_BANK[g] = officialList.concat(QUESTION_BANK[g]).map(autoSanitizeQuestion);
        });
    }

    // Tích hợp toàn bộ Ngân hàng câu hỏi phân loại chuẩn SGK KNTT & CV 7991
    if (typeof MASTER_QUESTION_BANK !== "undefined") {
        ["10", "11", "12"].forEach(function(g) {
            if (MASTER_QUESTION_BANK[g]) {
                var existingIds = new Set(QUESTION_BANK[g].map(function(q) { return q.id; }));
                MASTER_QUESTION_BANK[g].forEach(function(q) {
                    if (!existingIds.has(q.id)) {
                        QUESTION_BANK[g].push(autoSanitizeQuestion(q));
                        existingIds.add(q.id);
                    }
                });
            }
        });
    }

    // Gộp Ngân hàng chỉ báo KNTT chính thức vào Tab 2
    if (typeof OFFICIAL_KNTT_INDICATORS !== "undefined" && typeof INDICATORS_DATABASE !== "undefined") {
        var existingKeys = new Set(INDICATORS_DATABASE.map(function(i) { return i.grade + "_" + i.code + "_" + i.outcome.slice(0, 30); }));
        OFFICIAL_KNTT_INDICATORS.forEach(function(item) {
            var key = item.grade + "_" + item.code + "_" + item.outcome.slice(0, 30);
            if (!existingKeys.has(key)) {
                INDICATORS_DATABASE.unshift(item);
                existingKeys.add(key);
            }
        });
    }

    // Khởi tạo mặc định theo Khối & Kỳ đang chọn trên giao diện
    applyKNTTTermPreset();
    if (typeof renderIndicators === "function") {
        renderIndicators();
    }
    if (typeof renderGeneratedExam === "function") {
        renderGeneratedExam();
    }
    if (typeof initBankExplorer === "function") {
        initBankExplorer();
    }
}

/* Chuyển đổi Khối lớp (Lớp 10, 11, 12) */
function onGradeSelectChange() {
    var gradeSelect = document.getElementById("exam-grade-select");
    if (gradeSelect) {
        currentGrade = gradeSelect.value;
    }
    applyKNTTTermPreset();
}

/* Áp dụng Ma trận & Bản đặc tả theo Kỳ kiểm tra (GK1, CK1, GK2, CK2) của Tổ Hóa Học */
function applyKNTTTermPreset() {
    var gradeSelect = document.getElementById("exam-grade-select");
    var termSelect = document.getElementById("kntt-term-select");
    if (!gradeSelect || !termSelect) return;

    currentGrade = gradeSelect.value;
    currentTerm = termSelect.value;

    var config = KNTT_TERMS_CONFIG[currentGrade] && KNTT_TERMS_CONFIG[currentGrade][currentTerm];
    if (config) {
        var titleInput = document.getElementById("exam-title-input");
        if (titleInput) titleInput.value = config.title;

        var timeSelect = document.getElementById("exam-time-select");
        if (timeSelect) timeSelect.value = "45";

        var formatSelect = document.getElementById("exam-format-select");
        if (formatSelect) formatSelect.value = "MIXED";

        matrixRows = JSON.parse(JSON.stringify(config.rows));
        toggleEssayColumns();
        renderMatrixEditor();
        renderSpecificationTable();
        if (typeof renderGeneratedExam === "function") {
            renderGeneratedExam();
        }
    }
}

/* Ghi đè hàm tính tổng điểm Ma trận để hỗ trợ cả cách nhập 8 lệnh hỏi Đ/S (0,25đ/lệnh) lẫn 2 câu lớn Đ/S (1,0đ/câu) */
/* Ghi đè hàm tính tổng điểm Ma trận để đảm bảo chính xác 100% theo Công văn 7991:
   Phần I (Nhiều phương án): 12 câu = 3.00đ (0.25đ/câu)
   Phần II (Đ/S): 8 ý = 2 câu lớn = 2.00đ (0.25đ/ý hoặc 1.00đ/câu lớn)
   Phần III (TLN): 4 câu = 2.00đ (0.50đ/câu)
   Phần IV (Tự luận): 3 câu = 3.00đ (1.00đ/câu) hoặc 2 câu = 3.00đ
   Tổng điểm: 10.00 / 10.0 */
function recalculateMatrixTotals() {
    var formatEl = document.getElementById("exam-format-select");
    var format = formatEl ? formatEl.value : "MIXED";
    var p1 = 0, p2 = 0, p3 = 0, p4 = 0;

    matrixRows.forEach(function(r) {
        p1 += (r.p1_k + r.p1_h + r.p1_v);
        p2 += (r.p2_k + r.p2_h + r.p2_v);
        p3 += (r.p3_k + r.p3_h + r.p3_v);
        p4 += ((r.p4_k || 0) + (r.p4_h || 0) + (r.p4_v || 0));
    });

    var cntP1 = document.getElementById("cnt-part1");
    var cntP2 = document.getElementById("cnt-part2");
    var cntP3 = document.getElementById("cnt-part3");
    var cntP4 = document.getElementById("cnt-part4");

    if (cntP1) cntP1.innerText = p1;
    if (cntP2) cntP2.innerText = (p2 > 4) ? (p2 + " ý (" + Math.round(p2 / 4) + " câu lớn)") : (p2 + " câu lớn");
    if (cntP3) cntP3.innerText = p3;
    if (cntP4) cntP4.innerText = p4;

    var score1 = p1 * 0.25;
    var score2 = (p2 > 4) ? (p2 * 0.25) : (p2 * 1.0);
    var score3 = (p3 === 6 && format === "TN_ONLY") ? 1.5 : (p3 * 0.5);
    var score4 = (format === "TN_ONLY") ? 0 : (p4 === 3 ? 3.0 : (p4 === 2 ? 3.0 : (p4 * 1.0)));

    var scP1 = document.getElementById("score-part1");
    var scP2 = document.getElementById("score-part2");
    var scP3 = document.getElementById("score-part3");
    var scP4 = document.getElementById("score-part4");

    if (scP1) scP1.innerText = score1.toFixed(2);
    if (scP2) scP2.innerText = score2.toFixed(2);
    if (scP3) scP3.innerText = score3.toFixed(2);
    if (scP4) scP4.innerText = score4.toFixed(2);

    var totalScore = score1 + score2 + score3 + score4;
    var summaryBar = document.getElementById("matrix-summary-bar");
    var totalDisplay = document.getElementById("cnt-total-score");

    if (totalDisplay) totalDisplay.innerText = totalScore.toFixed(2);

    if (summaryBar) {
        if (Math.abs(totalScore - 10.0) < 0.05) {
            summaryBar.className = "mt-4 p-3 rounded-lg border text-xs flex flex-wrap items-center justify-between gap-2 bg-emerald-50 border-emerald-200 text-emerald-900 font-medium";
        } else {
            summaryBar.className = "mt-4 p-3 rounded-lg border text-xs flex flex-wrap items-center justify-between gap-2 bg-amber-50 border-amber-300 text-amber-900 font-medium";
        }
    }
}

/* Làm nổi bật mã chỉ báo (HH1.1, HH1.2...) ngay trong câu mô tả YCCĐ trên giao diện Web */
function highlightIndicatorCodesInText(text, badgeClass) {
    if (!text) return "";
    return String(text).replace(/\(?(HH\d\.\d(?:\s*,\s*HH\d\.\d)*)\)?/g, function(_, group) {
        var items = group.split(/\s*,\s*/);
        return items.map(function(code) {
            return "<span class='inline-block " + badgeClass + " font-bold px-1.5 py-0.2 rounded text-[11px] border mx-0.5'>[" + code + "]</span>";
        }).join(" ");
    });
}

/* Tách các ý YCCĐ theo từng mã chỉ báo (tránh phân mảnh câu khi có dấu chấm phẩy giữa chừng) */
function splitSpecBullets(text) {
    if (!text || text === "—" || text === "-") return [];
    var rawParts = String(text).split(/[;\n]/).map(function(s) { return s.trim(); }).filter(Boolean);
    var bullets = [];
    var temp = "";
    rawParts.forEach(function(p) {
        if (!temp) {
            temp = p;
        } else {
            temp += "; " + p;
        }
        if (/HH\d\.\d/.test(p)) {
            bullets.push(temp);
            temp = "";
        }
    });
    if (temp) {
        if (bullets.length > 0) {
            bullets[bullets.length - 1] += "; " + temp;
        } else {
            bullets.push(temp);
        }
    }
    return bullets;
}

/* Định dạng YCCĐ thành các dòng gạch đầu dòng rõ ràng từng chỉ báo nếu có nhiều ý */
function formatSpecDescWeb(text, badgeClass) {
    if (!text || text === "—" || text === "-" || text.trim() === "") return "—";
    var items = splitSpecBullets(text);
    if (items.length <= 1) {
        return highlightIndicatorCodesInText(items[0] || "", badgeClass);
    }
    return "<div class='space-y-1 mt-0.5 ml-2'>" + items.map(function(item) {
        var clean = item.replace(/^[-•*]\s*/, "");
        return "<div class='flex items-start gap-1.5 text-xs text-gray-800 leading-relaxed'><span class='text-gray-400 font-bold select-none'>-</span><span>" + highlightIndicatorCodesInText(clean, badgeClass) + "</span></div>";
    }).join("") + "</div>";
}

/* Tra cứu thông tin Đặc tả & Chỉ báo cho từng hàng của Ma trận */
function resolveRowSpecification(row) {
    var knowDesc = row.knowDesc || "";
    var underDesc = row.underDesc || "";
    var applyDesc = row.applyDesc || "";

    if (!knowDesc && !underDesc && !applyDesc) {
        var matchedIndicators = [];
        if (typeof OFFICIAL_KNTT_INDICATORS !== "undefined") {
            matchedIndicators = OFFICIAL_KNTT_INDICATORS.filter(function(item) {
                return item.grade === currentGrade && (
                    row.topic.toLowerCase().indexOf(item.topic.toLowerCase().slice(0, 6)) !== -1 ||
                    item.topic.toLowerCase().indexOf(row.topic.toLowerCase().slice(0, 6)) !== -1
                );
            });
        }
        if (matchedIndicators.length > 0) {
            var kList = matchedIndicators.filter(function(i) { return i.level === "Biết"; }).map(function(i) { return i.outcome + " (" + i.code + ")"; });
            var uList = matchedIndicators.filter(function(i) { return i.level === "Hiểu"; }).map(function(i) { return i.outcome + " (" + i.code + ")"; });
            var aList = matchedIndicators.filter(function(i) { return i.level === "Vận dụng"; }).map(function(i) { return i.outcome + " (" + i.code + ")"; });
            knowDesc = kList.join("; ") || "Nhận biết khái niệm, công thức hóa học, danh pháp IUPAC (HH1.1); Trình bày tính chất vật lí đặc trưng (HH1.2).";
            underDesc = uList.join("; ") || "Trình bày tính chất vật lí, trạng thái tự nhiên (HH1.2); Trình bày tính chất hóa học, viết PTHH minh họa (HH1.3).";
            applyDesc = aList.join("; ") || "Vận dụng tính toán định lượng, hiệu suất phản ứng (HH1.6); Giải bài toán thực tiễn đời sống (HH3.1).";
        } else {
            knowDesc = "Nhận biết khái niệm, công thức hóa học, danh pháp IUPAC (HH1.1); Trình bày tính chất vật lí đặc trưng (HH1.2).";
            underDesc = "Trình bày tính chất vật lí, trạng thái tự nhiên (HH1.2); Trình bày tính chất hóa học, viết PTHH minh họa (HH1.3).";
            applyDesc = "Vận dụng tính toán định lượng, hiệu suất (HH1.6); Giải bài toán thực tiễn đời sống (HH3.1).";
        }
    }

    var kCodes = Array.from(new Set(knowDesc.match(/HH\d\.\d/g) || ["HH1.1"]));
    var uCodes = Array.from(new Set(underDesc.match(/HH\d\.\d/g) || ["HH1.2"]));
    var aCodes = Array.from(new Set(applyDesc.match(/HH\d\.\d/g) || ["HH1.4"]));

    var codesWebHtml = "<div class='space-y-1.5 text-left'>" +
        "<div><span class='inline-block bg-emerald-100 text-emerald-900 border border-emerald-300 text-[11px] px-2 py-0.5 rounded font-bold w-full'>• Biết: " + kCodes.join(", ") + "</span></div>" +
        "<div><span class='inline-block bg-blue-100 text-blue-900 border border-blue-300 text-[11px] px-2 py-0.5 rounded font-bold w-full'>• Hiểu: " + uCodes.join(", ") + "</span></div>" +
        "<div><span class='inline-block bg-purple-100 text-purple-900 border border-purple-300 text-[11px] px-2 py-0.5 rounded font-bold w-full'>• VD: " + aCodes.join(", ") + "</span></div>" +
        "</div>";

    var codesWordHtml = "<p><strong>• Biết:</strong> " + kCodes.join(", ") + "</p>" +
        "<p><strong>• Hiểu:</strong> " + uCodes.join(", ") + "</p>" +
        "<p><strong>• VD:</strong> " + aCodes.join(", ") + "</p>";

    return {
        knowDesc: knowDesc || "—",
        underDesc: underDesc || "—",
        applyDesc: applyDesc || "—",
        knowWeb: formatSpecDescWeb(knowDesc, "bg-emerald-100 text-emerald-900 border-emerald-300"),
        underWeb: formatSpecDescWeb(underDesc, "bg-blue-100 text-blue-900 border-blue-300"),
        applyWeb: formatSpecDescWeb(applyDesc, "bg-purple-100 text-purple-900 border-purple-300"),
        codes: codesWebHtml,
        codesWord: codesWordHtml,
        kCodes: kCodes,
        uCodes: uCodes,
        aCodes: aCodes
    };
}

/* =========================================================================================
   CHI TIẾT CÁC Ý (a, b, c, d) CHO CÂU TRẮC NGHIỆM ĐÚNG/SAI (PHẦN II - THEO CÔNG VĂN 7991)
   Hỗ trợ 2 mô hình linh hoạt:
   1. Cùng 1 bài học (trọn vẹn 1 câu lớn: 4 ý a, b, c, d)
   2. Liên bài (các ý a, b, c, d được phân bổ ở các bài học khác nhau trong chương trình)
   ========================================================================================= */
function resolveRowP2Statements(row, spec, allRows) {
    if (!row) return [];
    var p2Count = (row.p2_k || 0) + (row.p2_h || 0) + (row.p2_v || 0);
    if (p2Count === 0 && (!row.p2_statements || row.p2_statements.length === 0)) return [];

    if (row.p2_statements && Array.isArray(row.p2_statements) && row.p2_statements.length === p2Count) {
        return row.p2_statements.map(function(s, idx) {
            var qNum = s.questionIndex || 1;
            var ltr = s.letter || ["a", "b", "c", "d"][idx % 4];
            var isSingle = (row.p2_statements.length === 4 && (!s.isIntegrated));
            return {
                questionIndex: qNum,
                globalIndex: s.globalIndex !== undefined ? s.globalIndex : idx,
                letter: ltr,
                label: s.label || (isSingle ? ("Ý " + ltr) : ("C" + qNum + "." + ltr)),
                fullLabel: s.fullLabel || ("Câu " + qNum + "." + ltr),
                shortLabel: s.shortLabel || ("C" + qNum + "." + ltr),
                subLabel: s.subLabel || ("Ý " + ltr),
                isIntegrated: s.isIntegrated !== undefined ? s.isIntegrated : (!isSingle),
                level: s.level || "Hiểu",
                indicator: s.indicator || "HH1.1",
                desc: s.desc || ""
            };
        });
    }

    if (!spec) {
        spec = resolveRowSpecification(row);
    }

    var kCodes = (spec.kCodes && spec.kCodes.length > 0) ? spec.kCodes : ["HH1.1"];
    var uCodes = (spec.uCodes && spec.uCodes.length > 0) ? spec.uCodes : ["HH1.2"];
    var aCodes = (spec.aCodes && spec.aCodes.length > 0) ? spec.aCodes : ["HH1.4"];

    var rows = (Array.isArray(allRows) && allRows.length > 0)
        ? allRows
        : ((typeof matrixRows !== "undefined" && Array.isArray(matrixRows)) ? matrixRows : [row]);

    // Tính cursor toàn cục: tổng số ý Phần II đã được phân bổ trước dòng này
    var globalCursor = 0;
    for (var r = 0; r < rows.length; r++) {
        if (rows[r] === row) break;
        var rCnt = (rows[r].p2_k || 0) + (rows[r].p2_h || 0) + (rows[r].p2_v || 0);
        globalCursor += rCnt;
    }

    // Xác định mức độ nhận thức của từng ý trên dòng này
    var levels = [];
    for (var i = 0; i < (row.p2_k || 0); i++) levels.push("Biết");
    for (var i = 0; i < (row.p2_h || 0); i++) levels.push("Hiểu");
    for (var i = 0; i < (row.p2_v || 0); i++) levels.push("Vận dụng");

    // Nếu người dùng nhập p2Count nhưng các ô cụ thể chưa có, dùng thang chuẩn CV 7991
    if (levels.length === 0 && p2Count > 0) {
        var ladder = ["Biết", "Hiểu", "Hiểu", "Vận dụng"];
        for (var i = 0; i < p2Count; i++) {
            levels.push(ladder[(globalCursor + i) % 4]);
        }
    }

    var statements = [];
    var letters = ["a", "b", "c", "d"];
    var isSingleTopicFull = (levels.length === 4 && (globalCursor % 4 === 0));

    for (var s = 0; s < levels.length; s++) {
        var globalIdx = globalCursor + s;
        var qNum = Math.floor(globalIdx / 4) + 1;
        var letter = letters[globalIdx % 4];
        var lvl = levels[s] || "Hiểu";

        var code = (lvl === "Biết") ? (kCodes[s % kCodes.length] || "HH1.1")
                 : (lvl === "Hiểu") ? (uCodes[s % uCodes.length] || "HH1.2")
                 : (aCodes[s % aCodes.length] || "HH1.4");

        var desc = (lvl === "Biết") ? spec.knowDesc : (lvl === "Hiểu" ? spec.underDesc : spec.applyDesc);
        if (!desc || desc === "—") desc = "Đánh giá nhận định hóa học theo yêu cầu cần đạt.";

        statements.push({
            questionIndex: qNum,
            globalIndex: globalIdx,
            letter: letter,
            label: isSingleTopicFull ? ("Ý " + letter) : ("C" + qNum + "." + letter),
            fullLabel: "Câu " + qNum + "." + letter,
            shortLabel: "C" + qNum + "." + letter,
            subLabel: "Ý " + letter,
            isIntegrated: !isSingleTopicFull,
            level: lvl,
            indicator: code,
            desc: desc
        });
    }

    return statements;
}

function getStatementMeta(q, sIdx, row) {
    if (q && q.statements && q.statements[sIdx]) {
        var s = q.statements[sIdx];
        if (s.level && s.indicator) {
            return { level: s.level, indicator: s.indicator };
        }
    }
    if (row) {
        var rowStmts = resolveRowP2Statements(row, null, (typeof matrixRows !== "undefined" ? matrixRows : null));
        if (rowStmts && rowStmts[sIdx]) {
            return { level: rowStmts[sIdx].level, indicator: rowStmts[sIdx].indicator };
        }
    }
    // Mặc định theo thang bậc tiệm tiến CV 7991:
    // Ý a: Biết; Ý b: Hiểu (hoặc Biết); Ý c: Hiểu; Ý d: Vận dụng
    var defaultLadder = [
        { level: "Biết", indicator: (q && q.indicator) ? q.indicator : "HH1.1" },
        { level: "Hiểu", indicator: "HH1.2" },
        { level: "Hiểu", indicator: "HH1.4" },
        { level: "Vận dụng", indicator: "HH1.6" }
    ];
    return defaultLadder[sIdx % 4];
}

/* =========================================================================================
   HỆ THỐNG GÁN CÂU HỎI TỰ ĐỘNG & TỰ CHỌN TỪ NGÂN HÀNG VÀO MA TRẬN & BẢN ĐẶC TẢ
   ========================================================================================= */

// Chuẩn hóa mã nhận diện chương/chủ đề chính xác tuyệt đối, ngăn chặn triệt để lẫn lộn giữa các chương
function getCanonicalChapterKey(grade, text) {
    if (!text || typeof text !== "string") return "";
    var t = text.toLowerCase();
    var g = String(grade || currentGrade);

    if (g === "10") {
        if (/\b(?:chương|chủ đề)\s*1\b/i.test(t)) return "CH_10_1";
        if (/\b(?:chương|chủ đề)\s*2\b/i.test(t)) return "CH_10_2";
        if (/\b(?:chương|chủ đề)\s*3\b/i.test(t)) return "CH_10_3";
        if (/\b(?:chương|chủ đề)\s*4\b/i.test(t)) return "CH_10_4";
        if (/\b(?:chương|chủ đề)\s*5\b/i.test(t)) return "CH_10_5";
        if (/\b(?:chương|chủ đề)\s*6\b/i.test(t)) return "CH_10_6";
        if (/\b(?:chương|chủ đề)\s*7\b/i.test(t)) return "CH_10_7";

        if (/nhóm halogen|halogen|nhóm viia|\bviia\b|clorua vôi|nước javel|muối halide|\bhalide\b|hydrochloric|\bhcl\b|agcl|agbr|agi|bromine|chlorine|fluorine|iodine/i.test(t)) return "CH_10_7";
        if (/tốc độ phản ứng|hằng số tốc độ|chất xúc tác|van't hoff|diện tích tiếp xúc/i.test(t)) return "CH_10_6";
        if (/năng lượng hóa học|enthalpy|biến thiên enthalpy|nhiệt tạo thành|thu nhiệt|tỏa nhiệt|\bkj\b|\\delta_r|\\delta_f|nhiệt phản ứng/i.test(t)) return "CH_10_5";
        if (/oxi hóa – khử|oxi hóa - khử|số oxi hóa|chất khử|chất oxi hóa|thăng bằng electron|thằng bằng electron/i.test(t)) return "CH_10_4";
        if (/liên kết|cộng hóa trị|liên kết ion|liên kết hydrogen|van der waals|xen phủ|quy tắc octet|\boctet\b|công thức lewis|tạo thành naf/i.test(t)) return "CH_10_3";
        if (/bảng tuần hoàn|bảng hệ thống tuần hoàn|chu kì|chu kỳ|bán kính nguyên tử|độ âm điện|định luật tuần hoàn|ô nguyên tố/i.test(t)) return "CH_10_2";
        if (/cấu tạo nguyên tử|hạt nhân|proton|neutron|electron|đồng vị|số khối|nguyên tử khối|cấu hình electron|orbital|\bao\b/i.test(t)) return "CH_10_1";
    } else if (g === "11") {
        if (/\b(?:chương|chủ đề)\s*1\b/i.test(t)) return "CH_11_1";
        if (/\b(?:chương|chủ đề)\s*2\b/i.test(t)) return "CH_11_2";
        if (/\b(?:chương|chủ đề)\s*3\b/i.test(t)) return "CH_11_3";
        if (/\b(?:chương|chủ đề)\s*4\b/i.test(t)) return "CH_11_4";
        if (/\b(?:chương|chủ đề)\s*5\b/i.test(t)) return "CH_11_5";
        if (/\b(?:chương|chủ đề)\s*6\b/i.test(t)) return "CH_11_6";

        if (/carbonyl|carboxylic|aldehyde|ketone|acid carboxylic/i.test(t)) return "CH_11_6";
        if (/dẫn xuất halogen|alcohol|phenol/i.test(t)) return "CH_11_5";
        if (/hydrocarbon|alkane|alkene|alkyne|arene/i.test(t)) return "CH_11_4";
        if (/đại cương.*hữu cơ|hóa học hữu cơ|công thức phân tử hợp chất hữu cơ/i.test(t)) return "CH_11_3";
        if (/nitrogen|sulfur|ammonia|nitric acid|sulfuric acid|khí so2/i.test(t)) return "CH_11_2";
        if (/cân bằng hóa học|sự điện li|ph\b|dung dịch acid - base|dung dịch acid base/i.test(t)) return "CH_11_1";
    } else if (g === "12") {
        if (/\b(?:chương|chủ đề)\s*1\b/i.test(t)) return "CH_12_1";
        if (/\b(?:chương|chủ đề)\s*2\b/i.test(t)) return "CH_12_2";
        if (/\b(?:chương|chủ đề)\s*3\b/i.test(t)) return "CH_12_3";
        if (/\b(?:chương|chủ đề)\s*4\b/i.test(t)) return "CH_12_4";
        if (/\b(?:chương|chủ đề)\s*5\b/i.test(t)) return "CH_12_5";
        if (/\b(?:chương|chủ đề)\s*6\b/i.test(t)) return "CH_12_6";
        if (/\b(?:chương|chủ đề)\s*7\b/i.test(t)) return "CH_12_7";
        if (/\b(?:chương|chủ đề)\s*8\b/i.test(t)) return "CH_12_8";

        if (/phức chất|chuyển tiếp dãy thứ nhất/i.test(t)) return "CH_12_8";
        if (/nhóm ia|nhóm iia|kim loại kiềm|kim loại kiềm thổ/i.test(t)) return "CH_12_7";
        if (/đại cương.*kim loại|tính chất kim loại|ăn mòn kim loại|điều chế kim loại/i.test(t)) return "CH_12_6";
        if (/pin điện|điện phân|thế điện cực/i.test(t)) return "CH_12_5";
        if (/polymer|polime|chất dẻo|cao su|tơ sợi/i.test(t)) return "CH_12_4";
        if (/amine|amino acid|peptide|protein/i.test(t)) return "CH_12_3";
        if (/carbohydrate|glucose|fructose|saccharose|tinh bột|cellulose/i.test(t)) return "CH_12_2";
        if (/ester|lipid|chất béo|xà phòng/i.test(t)) return "CH_12_1";
    }
    return "";
}

function isQuestionInTopic(q, rowTopic, grade) {
    if (!rowTopic) return true;
    var g = String(grade || currentGrade);
    var rowKey = getCanonicalChapterKey(g, rowTopic);
    if (!rowKey) return true;

    // 1. Ưu tiên tuyệt đối mã chương chuẩn hóa đã lưu của câu hỏi
    var qChapKey = getCanonicalChapterKey(g, q.chapter || q.topic || "");
    if (qChapKey) {
        return rowKey === qChapKey;
    }

    // 2. Nếu câu hỏi chưa có nhãn chương chuẩn, suy luận từ nội dung
    var stmtsText = (q.statements && Array.isArray(q.statements)) ? q.statements.map(function(s){ return s.text || ""; }).join(" ") : "";
    var fullContent = (q.content || "") + " " + stmtsText;
    var contentKey = getCanonicalChapterKey(g, fullContent);
    if (contentKey) {
        return rowKey === contentKey;
    }
    return false;
}

// Tính số lượng câu hỏi cần thiết cho 1 dòng ma trận
function getRequiredCountsForRow(row, format) {
    var rawP2 = (row.p2_k || 0) + (row.p2_h || 0) + (row.p2_v || 0);
    var p2Count = 0;
    if (rawP2 >= 4) {
        p2Count = Math.max(1, Math.round(rawP2 / 4));
    } else if (rawP2 > 0) {
        var p2Stmts = resolveRowP2Statements(row, null, (typeof matrixRows !== "undefined" ? matrixRows : null));
        var hasLeadingStmt = p2Stmts.some(function(s) { return s.letter === "a"; });
        p2Count = hasLeadingStmt ? 1 : 0;
    }
    var p4Count = (format !== "TN_ONLY") ? ((row.p4_k || 0) + (row.p4_h || 0) + (row.p4_v || 0)) : 0;
    return {
        p1: row.p1_k + row.p1_h + row.p1_v,
        p1_k: row.p1_k,
        p1_h: row.p1_h,
        p1_v: row.p1_v,
        p2: p2Count,
        p3: row.p3_k + row.p3_h + row.p3_v,
        p3_k: row.p3_k,
        p3_h: row.p3_h,
        p3_v: row.p3_v,
        p4: p4Count,
        p4_k: row.p4_k || 0,
        p4_h: row.p4_h || 0,
        p4_v: row.p4_v || 0,
        total: (row.p1_k + row.p1_h + row.p1_v) + p2Count + (row.p3_k + row.p3_h + row.p3_v) + p4Count
    };
}

// Tự động gán câu hỏi cho 1 dòng ma trận cụ thể từ Ngân hàng
function autoAssignQuestionsToRow(rowIndex) {
    if (rowIndex < 0 || rowIndex >= matrixRows.length) return;
    var row = matrixRows[rowIndex];
    var formatEl = document.getElementById("exam-format-select");
    var format = formatEl ? formatEl.value : "MIXED";
    var req = getRequiredCountsForRow(row, format);

    var pool = QUESTION_BANK[currentGrade] || [];
    var usedIds = new Set();
    var usedContentKeys = new Set();

    function getNormalizedKey(q) {
        if (!q || !q.content) return "";
        return q.content.toLowerCase().replace(/[^a-z0-9àáảãạăằắẳẵặâầấẩẫậèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđ]/g, "").slice(0, 60);
    }

    matrixRows.forEach(function(r, idx) {
        if (idx !== rowIndex && r.assignedQuestions) {
            r.assignedQuestions.forEach(function(q) {
                usedIds.add(q.id);
                var nk = getNormalizedKey(q);
                if (nk) usedContentKeys.add(nk);
            });
        }
    });

    function isAvailable(q) {
        if (usedIds.has(q.id)) return false;
        var nk = getNormalizedKey(q);
        if (nk && usedContentKeys.has(nk)) return false;
        return true;
    }

    function markUsed(q) {
        usedIds.add(q.id);
        var nk = getNormalizedKey(q);
        if (nk) usedContentKeys.add(nk);
    }

    function pickQuestions(part, level, count) {
        if (count <= 0) return [];
        var res = [];
        var candidates = pool.filter(function(q) {
            if (q.part !== part || !isAvailable(q)) return false;
            if (q.term && q.term !== currentTerm) return false;
            return isQuestionInTopic(q, row.topic, currentGrade);
        });

        var matched = candidates.filter(function(q) {
            return level ? (q.level === level) : true;
        });
        res = res.concat(matched.slice(0, count));
        res.forEach(markUsed);

        if (res.length < count) {
            var need = count - res.length;
            var sameTopicOthers = candidates.filter(isAvailable);
            var extra = sameTopicOthers.slice(0, need);
            extra.forEach(markUsed);
            res = res.concat(extra);
        }

        return res;
    }

    var assigned = [];
    if (row.p1_k > 0) assigned = assigned.concat(pickQuestions(1, 'Biết', row.p1_k));
    if (row.p1_h > 0) assigned = assigned.concat(pickQuestions(1, 'Hiểu', row.p1_h));
    if (row.p1_v > 0) assigned = assigned.concat(pickQuestions(1, 'Vận dụng', row.p1_v));

    if (req.p2 > 0) {
        assigned = assigned.concat(pickQuestions(2, '', req.p2));
    }

    if (row.p3_k > 0) assigned = assigned.concat(pickQuestions(3, 'Biết', row.p3_k));
    if (row.p3_h > 0) assigned = assigned.concat(pickQuestions(3, 'Hiểu', row.p3_h));
    if (row.p3_v > 0) assigned = assigned.concat(pickQuestions(3, 'Vận dụng', row.p3_v));

    if (format !== "TN_ONLY") {
        if (row.p4_k > 0) assigned = assigned.concat(pickQuestions(4, 'Biết', row.p4_k));
        if (row.p4_h > 0) assigned = assigned.concat(pickQuestions(4, 'Hiểu', row.p4_h));
        if (row.p4_v > 0) assigned = assigned.concat(pickQuestions(4, 'Vận dụng', row.p4_v));
    }

    row.assignedQuestions = assigned;
    renderSpecificationTable();
    if (typeof renderGeneratedExam === "function") renderGeneratedExam();
}

// Tự động gán cho toàn bộ ma trận
function autoAssignAllMatrixQuestions() {
    var formatEl = document.getElementById("exam-format-select");
    var format = formatEl ? formatEl.value : "MIXED";
    var pool = QUESTION_BANK[currentGrade] || [];
    var usedIds = new Set();
    var usedContentKeys = new Set();

    function getNormalizedKey(q) {
        if (!q || !q.content) return "";
        return q.content.toLowerCase().replace(/[^a-z0-9àáảãạăằắẳẵặâầấẩẫậèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđ]/g, "").slice(0, 60);
    }

    function isAvailable(q) {
        if (usedIds.has(q.id)) return false;
        var nk = getNormalizedKey(q);
        if (nk && usedContentKeys.has(nk)) return false;
        return true;
    }

    function markUsed(q) {
        usedIds.add(q.id);
        var nk = getNormalizedKey(q);
        if (nk) usedContentKeys.add(nk);
    }

    matrixRows.forEach(function(row) {
        var req = getRequiredCountsForRow(row, format);
        function pickQuestions(part, level, count) {
            if (count <= 0) return [];
            var res = [];
            var candidates = pool.filter(function(q) {
                if (q.part !== part || !isAvailable(q)) return false;
                if (q.term && q.term !== currentTerm) return false;
                return isQuestionInTopic(q, row.topic, currentGrade);
            });

            var matched = candidates.filter(function(q) {
                return level ? (q.level === level) : true;
            });
            res = res.concat(matched.slice(0, count));
            res.forEach(markUsed);

            if (res.length < count) {
                var need = count - res.length;
                var sameTopicOthers = candidates.filter(isAvailable);
                var extra = sameTopicOthers.slice(0, need);
                extra.forEach(markUsed);
                res = res.concat(extra);
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

    renderSpecificationTable();
    if (typeof renderGeneratedExam === "function") renderGeneratedExam();
    alert("Đã tự động lấy và gán thành công " + usedIds.size + " câu hỏi từ Ngân hàng vào toàn bộ các bài học trong Ma trận & Bản đặc tả!");
}

// Xóa tất cả câu đã gán
function clearAllAssignedQuestions() {
    matrixRows.forEach(function(r) { r.assignedQuestions = []; });
    renderSpecificationTable();
    if (typeof renderGeneratedExam === "function") renderGeneratedExam();
}

function clearRowAssignedQuestions(rowIndex) {
    if (matrixRows[rowIndex]) {
        matrixRows[rowIndex].assignedQuestions = [];
        renderSpecificationTable();
        if (typeof renderGeneratedExam === "function") renderGeneratedExam();
    }
}

// Thay thế 1 câu cụ thể bằng câu khác cùng mức độ và dạng thức
function replaceAssignedQuestion(rowIndex, qId) {
    var row = matrixRows[rowIndex];
    if (!row || !row.assignedQuestions) return;
    var targetIdx = row.assignedQuestions.findIndex(function(q) { return q.id === qId; });
    if (targetIdx === -1) return;
    var oldQ = row.assignedQuestions[targetIdx];

    var pool = QUESTION_BANK[currentGrade] || [];
    var usedIds = new Set();
    matrixRows.forEach(function(r) {
        if (r.assignedQuestions) r.assignedQuestions.forEach(function(q) { usedIds.add(q.id); });
    });

    var candidates = pool.filter(function(q) {
        return q.part === oldQ.part && !usedIds.has(q.id) && (oldQ.level ? q.level === oldQ.level : true);
    });
    if (candidates.length === 0) {
        candidates = pool.filter(function(q) { return q.part === oldQ.part && !usedIds.has(q.id); });
    }

    if (candidates.length > 0) {
        var newQ = candidates[Math.floor(Math.random() * candidates.length)];
        row.assignedQuestions[targetIdx] = newQ;
        renderSpecificationTable();
        if (typeof renderGeneratedExam === "function") renderGeneratedExam();
    } else {
        alert("Không còn câu hỏi thay thế nào khả dụng trong Ngân hàng cho dạng thức này!");
    }
}

function removeAssignedQuestion(rowIndex, qId) {
    var row = matrixRows[rowIndex];
    if (!row || !row.assignedQuestions) return;
    row.assignedQuestions = row.assignedQuestions.filter(function(q) { return q.id !== qId; });
    renderSpecificationTable();
    if (typeof renderGeneratedExam === "function") renderGeneratedExam();
}

function toggleAssignedPreview(rowIndex) {
    var el = document.getElementById("assigned-preview-" + rowIndex);
    var btn = document.getElementById("btn-preview-label-" + rowIndex);
    if (!el) return;
    if (el.classList.contains("hidden")) {
        el.classList.remove("hidden");
        if (btn) btn.innerText = "Ẩn bớt";
    } else {
        el.classList.add("hidden");
        var count = matrixRows[rowIndex].assignedQuestions ? matrixRows[rowIndex].assignedQuestions.length : 0;
        if (btn) btn.innerText = "Xem " + count + " câu";
    }
    if (typeof triggerKaTeX === "function") triggerKaTeX();
}

function generateExamFromAssignedMatrix() {
    switchTab('generator');
    renderGeneratedExam();
}

/* =========================================================================================
   CONTROLLER MODAL TỰ CHỌN CÂU HỎI TỪ NGÂN HÀNG (ROW QUESTION PICKER)
   ========================================================================================= */
var currentRowPickerRowIndex = -1;
var tempRowSelectedQuestions = [];

function openRowQuestionPicker(rowIndex) {
    if (rowIndex < 0 || rowIndex >= matrixRows.length) return;
    currentRowPickerRowIndex = rowIndex;
    var row = matrixRows[rowIndex];

    var modal = document.getElementById("modal-row-question-picker");
    if (!modal) return;

    var topicBadge = document.getElementById("row-picker-topic-badge");
    if (topicBadge) topicBadge.innerText = row.topic;

    tempRowSelectedQuestions = row.assignedQuestions ? JSON.parse(JSON.stringify(row.assignedQuestions)) : [];

    updateRowPickerStats();
    renderRowPickerCandidates();

    modal.classList.remove("hidden");
    if (typeof triggerKaTeX === "function") triggerKaTeX();
}

function closeRowQuestionPickerModal() {
    var modal = document.getElementById("modal-row-question-picker");
    if (modal) modal.classList.add("hidden");
    currentRowPickerRowIndex = -1;
    tempRowSelectedQuestions = [];
}

function updateRowPickerStats() {
    if (currentRowPickerRowIndex < 0 || currentRowPickerRowIndex >= matrixRows.length) return;
    var row = matrixRows[currentRowPickerRowIndex];
    var formatEl = document.getElementById("exam-format-select");
    var format = formatEl ? formatEl.value : "MIXED";
    var req = getRequiredCountsForRow(row, format);

    var curP1 = tempRowSelectedQuestions.filter(function(q) { return q.part === 1; }).length;
    var curP2 = tempRowSelectedQuestions.filter(function(q) { return q.part === 2; }).length;
    var curP3 = tempRowSelectedQuestions.filter(function(q) { return q.part === 3; }).length;
    var curP4 = tempRowSelectedQuestions.filter(function(q) { return q.part === 4; }).length;

    var elP1 = document.getElementById("row-picker-p1-stat");
    var elP2 = document.getElementById("row-picker-p2-stat");
    var elP3 = document.getElementById("row-picker-p3-stat");
    var elP4 = document.getElementById("row-picker-p4-stat");
    var elTotal = document.getElementById("row-picker-total-selected-count");

    if (elP1) elP1.innerText = "P.I (Nhiều phương án): " + curP1 + "/" + req.p1;
    if (elP2) elP2.innerText = "P.II (Đ/S): " + curP2 + "/" + req.p2;
    if (elP3) elP3.innerText = "P.III (TLN): " + curP3 + "/" + req.p3;
    if (elP4) elP4.innerText = "P.IV (TL): " + curP4 + "/" + req.p4;
    if (elTotal) elTotal.innerText = tempRowSelectedQuestions.length;
}

function renderRowPickerCandidates() {
    if (currentRowPickerRowIndex < 0 || currentRowPickerRowIndex >= matrixRows.length) return;
    var row = matrixRows[currentRowPickerRowIndex];
    var container = document.getElementById("row-picker-candidates-container");
    if (!container) return;

    var filterPart = document.getElementById("row-picker-filter-part").value;
    var filterLevel = document.getElementById("row-picker-filter-level").value;
    var filterScope = document.getElementById("row-picker-filter-scope").value;
    var searchKeyword = document.getElementById("row-picker-filter-search").value.trim().toLowerCase();

    var pool = QUESTION_BANK[currentGrade] || [];
    var selectedIds = new Set(tempRowSelectedQuestions.map(function(q) { return q.id; }));

    var filtered = pool.filter(function(q) {
        if (filterPart !== "ALL" && String(q.part) !== filterPart) return false;
        if (filterLevel !== "ALL" && q.level !== filterLevel) return false;
        if (filterScope === "TOPIC_FIRST") {
            var tMatch = isQuestionInTopic(q, row.topic, currentGrade);
            if (!tMatch && !selectedIds.has(q.id)) return false;
        }
        if (searchKeyword) {
            var contentStr = (q.content + " " + (q.options ? q.options.join(" ") : "") + " " + (q.explain || "")).toLowerCase();
            if (contentStr.indexOf(searchKeyword) === -1) return false;
        }
        return true;
    });

    if (filtered.length === 0) {
        container.innerHTML = "<div class='text-center py-10 text-gray-400 text-xs'>" +
            "<i class='fa-solid fa-box-open text-3xl mb-2 text-gray-300'></i>" +
            "<p>Không tìm thấy câu hỏi phù hợp với bộ lọc hiện tại.</p>" +
            "<p class='mt-1 text-gray-500'>Hãy thử chọn 'Toàn bộ ngân hàng khối lớp này' hoặc xóa bớt từ khóa tìm kiếm.</p>" +
            "</div>";
        return;
    }

    container.innerHTML = filtered.map(function(q) {
        var isSelected = selectedIds.has(q.id);
        var partName = (q.part === 1) ? "Phần I: Nhiều phương án" : (q.part === 2) ? "Phần II: Đúng/Sai" : (q.part === 3) ? "Phần III: TL Ngắn" : "Phần IV: Tự Luận";
        var partBadgeColor = (q.part === 1) ? "bg-green-100 text-green-800" : (q.part === 2) ? "bg-blue-100 text-blue-800" : (q.part === 3) ? "bg-purple-100 text-purple-800" : "bg-rose-100 text-rose-800";
        var levelColor = (q.level === "Biết") ? "bg-emerald-50 text-emerald-700 border-emerald-200" : (q.level === "Hiểu") ? "bg-blue-50 text-blue-700 border-blue-200" : "bg-purple-50 text-purple-700 border-purple-200";

        var bodyHtml = "";
        if (q.part === 1 && q.options) {
            bodyHtml = "<div class='space-y-1 mt-2 text-[11px] pl-1'>" +
                q.options.map(function(opt, optIdx) {
                    var isCorrect = (optIdx === q.correct);
                    return "<div class='p-1.5 rounded border " + (isCorrect ? "bg-green-50 border-green-300 font-bold text-green-900" : "bg-gray-50 border-gray-200 text-gray-700") + "'>" +
                        "<strong>" + String.fromCharCode(65 + optIdx) + ".</strong> " + cleanLatexForWord(opt) +
                        "</div>";
                }).join("") +
                "</div>";
        } else if (q.part === 2 && q.statements) {
            bodyHtml = "<div class='space-y-1 mt-2 text-[11px]'>" +
                q.statements.map(function(st) {
                    return "<div class='flex items-center justify-between p-1 bg-gray-50 rounded border border-gray-200'>" +
                        "<span>" + cleanLatexForWord(st.text) + "</span>" +
                        "<span class='font-bold text-[10px] px-1.5 py-0.5 rounded " + (st.correct ? "bg-green-100 text-green-800" : "bg-rose-100 text-rose-800") + "'>" + (st.correct ? "ĐÚNG" : "SAI") + "</span>" +
                        "</div>";
                }).join("") +
                "</div>";
        } else if (q.part === 3) {
            bodyHtml = "<div class='mt-1.5 text-[11px] text-purple-800 font-semibold bg-purple-50 p-1.5 rounded border border-purple-200'>Đáp án ngắn: " + cleanLatexForWord(q.answer || "N/A") + "</div>";
        } else if (q.part === 4) {
            bodyHtml = "<div class='mt-1.5 text-[11px] text-rose-800 bg-rose-50/70 p-2 rounded border border-rose-200 leading-relaxed'>" + formatSolutionHtml(q.solution || q.explain || "") + "</div>";
        }

        return "<div class='p-3.5 rounded-xl border transition " + (isSelected ? "bg-brand-50/40 border-brand-500 shadow-sm ring-1 ring-brand-500" : "bg-white border-gray-200 hover:border-gray-300") + "'>" +
            "<div class='flex items-start justify-between gap-2'>" +
            "<div class='space-y-1 flex-grow'>" +
            "<div class='flex flex-wrap items-center gap-1.5 text-[11px]'>" +
            "<span class='px-2 py-0.5 rounded font-bold " + partBadgeColor + "'>" + partName + "</span>" +
            "<span class='px-2 py-0.5 rounded font-semibold border " + levelColor + "'>" + q.level + "</span>" +
            (q.indicator ? "<span class='px-2 py-0.5 rounded font-bold bg-amber-100 text-amber-800 text-[10px]'>[" + q.indicator + "]</span>" : "") +
            "<span class='text-gray-400 text-[10px]'>ID: " + q.id + "</span>" +
            "</div>" +
            "<p class='text-xs text-gray-800 font-medium leading-relaxed mt-1'>" + cleanLatexForWord(q.content) + "</p>" +
            bodyHtml +
            "</div>" +
            "<div class='flex-shrink-0 pt-1'>" +
            "<button type='button' onclick='toggleRowPickerCandidate(\"" + q.id + "\")' class='px-3 py-1.5 rounded-lg text-xs font-bold transition flex items-center gap-1 " + (isSelected ? "bg-brand-600 hover:bg-brand-700 text-white shadow-sm" : "bg-gray-100 hover:bg-brand-50 text-gray-700 border border-gray-300") + "'>" +
            (isSelected ? "<i class='fa-solid fa-check'></i> <span>Đã Chọn</span>" : "<i class='fa-solid fa-plus'></i> <span>Chọn Câu Này</span>") +
            "</button>" +
            "</div>" +
            "</div>" +
            "</div>";
    }).join("");

    if (typeof triggerKaTeX === "function") triggerKaTeX();
}

function toggleRowPickerCandidate(qId) {
    var pool = QUESTION_BANK[currentGrade] || [];
    var existingIdx = tempRowSelectedQuestions.findIndex(function(q) { return q.id === qId; });
    if (existingIdx !== -1) {
        tempRowSelectedQuestions.splice(existingIdx, 1);
    } else {
        var found = pool.find(function(q) { return q.id === qId; });
        if (found) {
            tempRowSelectedQuestions.push(found);
        }
    }
    updateRowPickerStats();
    renderRowPickerCandidates();
}

function autoFillCurrentRowPicker() {
    if (currentRowPickerRowIndex < 0 || currentRowPickerRowIndex >= matrixRows.length) return;
    autoAssignQuestionsToRow(currentRowPickerRowIndex);
    tempRowSelectedQuestions = matrixRows[currentRowPickerRowIndex].assignedQuestions ? JSON.parse(JSON.stringify(matrixRows[currentRowPickerRowIndex].assignedQuestions)) : [];
    updateRowPickerStats();
    renderRowPickerCandidates();
}

function saveRowQuestionPickerSelection() {
    if (currentRowPickerRowIndex < 0 || currentRowPickerRowIndex >= matrixRows.length) return;
    matrixRows[currentRowPickerRowIndex].assignedQuestions = tempRowSelectedQuestions;
    closeRowQuestionPickerModal();
    renderSpecificationTable();
    if (typeof renderGeneratedExam === "function") renderGeneratedExam();
}

/* =========================================================================================
   HIỂN THỊ BẢN ĐẶC TẢ TRÊN GIAO DIỆN WEB (TAB 3) VỚI CỘT GÁN CÂU HỎI TỪ NGÂN HÀNG
   ========================================================================================= */
function renderSpecificationTable() {
    var formatEl = document.getElementById("exam-format-select");
    var format = formatEl ? formatEl.value : "MIXED";
    var tbody = document.getElementById("specification-tbody");
    if (!tbody) return;

    var totalP2 = 0;
    var totalAssigned = 0;
    var totalNeededAll = 0;

    matrixRows.forEach(function(r) {
        totalP2 += (r.p2_k + r.p2_h + r.p2_v);
        var req = getRequiredCountsForRow(r, format);
        totalNeededAll += req.total;
        if (r.assignedQuestions) totalAssigned += r.assignedQuestions.length;
    });
    var p2Unit = (totalP2 > 4) ? "lệnh hỏi (ý)" : "câu lớn";

    // Cập nhật thống kê trên thanh thông tin
    var bankTotalBadge = document.getElementById("matrix-bank-total-badge");
    if (bankTotalBadge && QUESTION_BANK[currentGrade]) {
        bankTotalBadge.innerText = QUESTION_BANK[currentGrade].length + " câu (Lớp " + currentGrade + ")";
    }
    var assignedSummaryBadge = document.getElementById("matrix-assigned-summary-badge");
    if (assignedSummaryBadge) {
        if (totalAssigned >= totalNeededAll && totalNeededAll > 0) {
            assignedSummaryBadge.className = "bg-emerald-100 text-emerald-800 px-2.5 py-0.5 rounded-full font-bold";
            assignedSummaryBadge.innerHTML = "<i class='fa-solid fa-check mr-1'></i>Đã gán đủ: " + totalAssigned + "/" + totalNeededAll + " câu";
        } else if (totalAssigned > 0) {
            assignedSummaryBadge.className = "bg-amber-100 text-amber-800 px-2.5 py-0.5 rounded-full font-bold";
            assignedSummaryBadge.innerHTML = "<i class='fa-solid fa-clock mr-1'></i>Đã gán: " + totalAssigned + "/" + totalNeededAll + " câu";
        } else {
            assignedSummaryBadge.className = "bg-gray-100 text-gray-600 px-2.5 py-0.5 rounded-full font-bold";
            assignedSummaryBadge.innerText = "Chưa gán câu hỏi (" + totalNeededAll + " câu cần gán)";
        }
    }

    tbody.innerHTML = matrixRows.map(function(row, idx) {
        var p1_cnt = row.p1_k + row.p1_h + row.p1_v;
        var p2_cnt = row.p2_k + row.p2_h + row.p2_v;
        var p3_cnt = row.p3_k + row.p3_h + row.p3_v;
        var p4_cnt = (row.p4_k || 0) + (row.p4_h || 0) + (row.p4_v || 0);
        var spec = resolveRowSpecification(row);
        var req = getRequiredCountsForRow(row, format);
        var assigned = row.assignedQuestions || [];

        // Trạng thái badge cho cột gán câu hỏi
        var statusBadge = "";
        if (req.total === 0) {
            statusBadge = "<span class='bg-gray-100 text-gray-500 font-semibold px-2 py-0.5 rounded text-[10px]'>Không phân phối câu</span>";
        } else if (assigned.length >= req.total) {
            statusBadge = "<span class='bg-emerald-100 text-emerald-800 font-bold px-2 py-0.5 rounded text-[10px] inline-flex items-center gap-1'><i class='fa-solid fa-check'></i> Đã gán đủ: " + assigned.length + "/" + req.total + " câu</span>";
        } else if (assigned.length > 0) {
            statusBadge = "<span class='bg-amber-100 text-amber-800 font-bold px-2 py-0.5 rounded text-[10px] inline-flex items-center gap-1'><i class='fa-solid fa-clock'></i> Đã gán: " + assigned.length + "/" + req.total + " câu</span>";
        } else {
            statusBadge = "<span class='bg-rose-50 text-rose-700 font-semibold px-2 py-0.5 rounded text-[10px] border border-rose-200'>Chưa gán (" + req.total + " câu)</span>";
        }

        // Ngăn xếp hiển thị câu hỏi đã gán
        var assignedDrawerHtml = "";
        if (assigned.length > 0) {
            assignedDrawerHtml = "<div id='assigned-preview-" + idx + "' class='hidden mt-2 p-2.5 bg-white rounded-lg border border-gray-200 space-y-2 max-h-56 overflow-y-auto text-[11px] shadow-inner'>" +
                assigned.map(function(q, qIdx) {
                    var pBadge = (q.part === 1) ? "<span class='bg-green-100 text-green-800 font-bold px-1.5 py-0.2 rounded text-[10px]'>P.I</span>" : (q.part === 2) ? "<span class='bg-blue-100 text-blue-800 font-bold px-1.5 py-0.2 rounded text-[10px]'>P.II</span>" : (q.part === 3) ? "<span class='bg-purple-100 text-purple-800 font-bold px-1.5 py-0.2 rounded text-[10px]'>P.III</span>" : "<span class='bg-rose-100 text-rose-800 font-bold px-1.5 py-0.2 rounded text-[10px]'>P.IV</span>";
                    return "<div class='flex items-start justify-between gap-1.5 p-1.5 bg-gray-50 rounded border border-gray-100'>" +
                        "<div class='space-y-0.5 flex-grow truncate'>" +
                        "<div class='flex items-center gap-1 text-[10px] font-semibold'>" +
                        pBadge + "<span class='text-gray-600'>[" + q.level + (q.indicator ? " - " + q.indicator : "") + "]</span>" +
                        "</div>" +
                        "<p class='text-gray-800 truncate'>" + q.content + "</p>" +
                        "</div>" +
                        "<div class='flex items-center gap-1 flex-shrink-0'>" +
                        "<button type='button' onclick='replaceAssignedQuestion(" + idx + ", \"" + q.id + "\")' title='Đổi sang câu khác' class='p-1 text-teal-600 hover:bg-teal-50 rounded'><i class='fa-solid fa-rotate'></i></button>" +
                        "<button type='button' onclick='removeAssignedQuestion(" + idx + ", \"" + q.id + "\")' title='Xóa câu này' class='p-1 text-rose-500 hover:bg-rose-50 rounded'><i class='fa-solid fa-xmark'></i></button>" +
                        "</div>" +
                        "</div>";
                }).join("") +
                "</div>";
        }

        var p2Stmts = resolveRowP2Statements(row, spec, matrixRows);
        var p2CodesSummary = "";
        var p2DetailsHtml = "";
        var isSingleTopicFull = (p2_cnt === 4 && p2Stmts.length === 4 && !p2Stmts[0].isIntegrated);

        if (p2_cnt > 0 && p2Stmts.length > 0) {
            var p2Header = isSingleTopicFull ? "Chỉ báo 4 ý Đ/S (Cùng bài):" : ("Chỉ báo Đ/S Liên bài (" + p2_cnt + " ý):");
            p2CodesSummary = "<div class='mt-1.5 pt-1.5 border-t border-brand-200/80 text-[10px] text-left space-y-0.5'>" +
                "<span class='font-bold text-blue-950 block'><i class='fa-solid fa-list-check text-blue-600 mr-1'></i>" + p2Header + "</span>" +
                p2Stmts.map(function(st) {
                    return "<div class='flex items-center justify-between'><span>• " + st.label + ":</span><span class='font-mono font-bold text-brand-700 bg-white px-1 rounded border border-brand-200'>[" + st.indicator + "]</span></div>";
                }).join("") +
                "</div>";

            p2DetailsHtml = "<div class='mt-1.5 p-1.5 bg-blue-50/70 border border-blue-200 rounded text-[11px] leading-tight space-y-1 text-left'>" +
                "<div class='font-bold text-blue-900 border-b border-blue-200/60 pb-0.5 flex items-center justify-between'>" +
                "<span><i class='fa-solid fa-circle-question text-blue-600 mr-1'></i>" + (isSingleTopicFull ? "Chi tiết 4 ý Đ/S (Cùng 1 bài):" : "Chi tiết ý Đ/S (Liên bài):") + "</span>" +
                "</div>" +
                "<div class='space-y-0.5 pt-0.5'>" +
                p2Stmts.map(function(st) {
                    var lvlBadge = (st.level === "Biết")
                        ? "<span class='bg-emerald-100 text-emerald-800 px-1 py-0.2 rounded font-bold text-[10px]'>Biết</span>"
                        : (st.level === "Hiểu")
                        ? "<span class='bg-blue-100 text-blue-800 px-1 py-0.2 rounded font-bold text-[10px]'>Hiểu</span>"
                        : "<span class='bg-purple-100 text-purple-800 px-1 py-0.2 rounded font-bold text-[10px]'>Vận dụng</span>";
                    return "<div class='flex items-center justify-between text-[11px]'>" +
                        "<span class='font-bold text-gray-800'>" + st.label + ":</span>" +
                        "<span class='inline-flex items-center gap-1'>" + lvlBadge + " <span class='font-mono font-bold text-brand-700 text-[10px]'>[" + st.indicator + "]</span></span>" +
                        "</div>";
                }).join("") +
                "</div>" +
                "</div>";
        }

        var rowP2Label = (p2_cnt === 4 && isSingleTopicFull)
            ? "1 câu (4 ý a,b,c,d)"
            : (p2_cnt + " ý (" + p2Stmts.map(function(s) { return s.shortLabel; }).join(", ") + ")");

        return "<tr>" +
            "<td class='text-center font-bold'>" + (idx + 1) + "</td>" +
            "<td>" +
            "<p class='font-bold text-gray-800 text-xs'>" + row.topic + "</p>" +
            "<button type='button' onclick='openSpecPicker(" + idx + ")' class='mt-1.5 px-2 py-0.5 bg-amber-500 hover:bg-amber-600 text-white rounded text-[10px] font-bold inline-flex items-center gap-1 shadow-sm transition'>" +
            "<i class='fa-solid fa-pen-to-square'></i> <span>Sửa YCCĐ & Chỉ Báo</span>" +
            "</button>" +
            "</td>" +
            "<td class='space-y-1.5 leading-relaxed'>" +
            (spec.knowDesc && spec.knowDesc !== "—" && spec.knowDesc !== "-" ? "<div><strong class='text-emerald-800'>• Nhận biết:</strong> " + spec.knowWeb + "</div>" : "") +
            (spec.underDesc && spec.underDesc !== "—" && spec.underDesc !== "-" ? "<div><strong class='text-blue-800'>• Thông hiểu:</strong> " + spec.underWeb + "</div>" : "") +
            (spec.applyDesc && spec.applyDesc !== "—" && spec.applyDesc !== "-" ? "<div><strong class='text-purple-800'>• Vận dụng:</strong> " + spec.applyWeb + "</div>" : "") +
            "</td>" +
            "<td class='align-middle bg-brand-50/60 p-2 text-center'>" + spec.codes + p2CodesSummary + "</td>" +
            "<td class='space-y-1'>" +
            "<p>- Phần I (Nhiều phương án): <strong>" + p1_cnt + "</strong> (B: " + row.p1_k + ", H: " + row.p1_h + ", VD: " + row.p1_v + ")</p>" +
            "<p>- Phần II (Đ/S): <strong>" + p2_cnt + "</strong> ý — " + rowP2Label + " (B: " + row.p2_k + ", H: " + row.p2_h + ", VD: " + row.p2_v + ")</p>" +
            p2DetailsHtml +
            "<p>- Phần III (TLN): <strong>" + p3_cnt + "</strong> (B: " + row.p3_k + ", H: " + row.p3_h + ", VD: " + row.p3_v + ")</p>" +
            (format !== "TN_ONLY" ? "<p class='text-rose-700 font-semibold'>- Phần IV (TL): <strong>" + p4_cnt + "</strong> (B: " + (row.p4_k || 0) + ", H: " + (row.p4_h || 0) + ", VD: " + (row.p4_v || 0) + ")</p>" : "") +
            "</td>" +
            "<td class='space-y-1.5 align-top bg-teal-50/30 p-2'>" +
            "<div class='flex items-center justify-between gap-1'>" +
            statusBadge +
            (assigned.length > 0 ? "<button type='button' onclick='clearRowAssignedQuestions(" + idx + ")' title='Bỏ gán các câu của dòng này' class='text-[10px] text-rose-500 hover:underline'><i class='fa-solid fa-trash-can mr-0.5'></i>Xóa</button>" : "") +
            "</div>" +
            "<div class='flex flex-wrap items-center gap-1.5 mt-1'>" +
            "<button type='button' onclick='autoAssignQuestionsToRow(" + idx + ")' class='px-2 py-1 bg-teal-600 hover:bg-teal-700 text-white rounded text-[11px] font-bold inline-flex items-center gap-1 shadow-sm transition'>" +
            "<i class='fa-solid fa-bolt text-amber-300'></i> <span>⚡ Tự Động Lấy</span>" +
            "</button>" +
            "<button type='button' onclick='openRowQuestionPicker(" + idx + ")' class='px-2 py-1 bg-blue-600 hover:bg-blue-700 text-white rounded text-[11px] font-bold inline-flex items-center gap-1 shadow-sm transition'>" +
            "<i class='fa-solid fa-hand-pointer'></i> <span>🔍 Tự Chọn</span>" +
            "</button>" +
            (assigned.length > 0 ? "<button type='button' onclick='toggleAssignedPreview(" + idx + ")' class='px-2 py-1 bg-gray-100 hover:bg-gray-200 text-gray-800 rounded text-[11px] font-semibold border border-gray-300 inline-flex items-center gap-1 transition'><i class='fa-solid fa-eye text-brand-600'></i> <span id='btn-preview-label-" + idx + "'>Xem " + assigned.length + " câu</span></button>" : "") +
            "</div>" +
            assignedDrawerHtml +
            "</td>" +
            "</tr>";
    }).join("");

    if (typeof triggerKaTeX === "function") {
        triggerKaTeX();
    }
}

/* =========================================================================================
   BỘ ĐIỀU KHIỂN MODAL TỰ CHỌN MỨC ĐỘ ĐÁNH GIÁ, YCCĐ & CHỈ BÁO NĂNG LỰC (KNTT & CV 7991)
   ========================================================================================= */
var currentEditingSpecRowIndex = -1;

function openSpecPicker(rowIndex) {
    if (rowIndex < 0 || rowIndex >= matrixRows.length) return;
    currentEditingSpecRowIndex = rowIndex;

    var modal = document.getElementById("modal-spec-picker");
    if (!modal) return;

    var row = matrixRows[rowIndex];

    // Badge khối lớp
    var gradeBadge = document.getElementById("spec-picker-grade-badge");
    if (gradeBadge) gradeBadge.innerText = currentGrade;

    // Danh sách bài học theo SGK KNTT
    var lessonSelect = document.getElementById("spec-picker-lesson-select");
    var lessons = (typeof KNTT_LESSONS_STANDARDS !== "undefined" && KNTT_LESSONS_STANDARDS[currentGrade]) ? KNTT_LESSONS_STANDARDS[currentGrade] : [];

    var selectHtml = "<option value=''>-- Chọn bài học chuẩn SGK Kết Nối Tri Thức (Lớp " + currentGrade + ") --</option>";
    var currentChapter = "";
    lessons.forEach(function(l) {
        if (l.chapter !== currentChapter) {
            if (currentChapter !== "") selectHtml += "</optgroup>";
            currentChapter = l.chapter;
            selectHtml += "<optgroup label='" + currentChapter + "'>";
        }
        selectHtml += "<option value='" + l.id + "'>" + l.lesson + "</option>";
    });
    if (currentChapter !== "") selectHtml += "</optgroup>";
    selectHtml += "<option value='custom'>[+] Tự nhập chủ đề khác</option>";
    if (lessonSelect) lessonSelect.innerHTML = selectHtml;

    // Tìm bài học khớp với dòng hiện tại nếu có
    var matchedLessonId = "";
    lessons.forEach(function(l) {
        var cleanRow = row.topic.toLowerCase().replace(/chủ đề \d+:?/g, "").trim();
        var cleanLesson = l.lesson.toLowerCase();
        if (cleanRow.indexOf(cleanLesson.slice(0, 10)) !== -1 || cleanLesson.indexOf(cleanRow.slice(0, 10)) !== -1) {
            matchedLessonId = l.id;
        }
    });
    if (lessonSelect) {
        lessonSelect.value = matchedLessonId || "";
    }

    // Tên chủ đề
    var topicInput = document.getElementById("spec-picker-topic-input");
    if (topicInput) topicInput.value = row.topic || "";

    // Số câu hỏi
    if (document.getElementById("spec-p1-k")) document.getElementById("spec-p1-k").value = row.p1_k || 0;
    if (document.getElementById("spec-p1-h")) document.getElementById("spec-p1-h").value = row.p1_h || 0;
    if (document.getElementById("spec-p1-v")) document.getElementById("spec-p1-v").value = row.p1_v || 0;

    if (document.getElementById("spec-p2-k")) document.getElementById("spec-p2-k").value = row.p2_k || 0;
    if (document.getElementById("spec-p2-h")) document.getElementById("spec-p2-h").value = row.p2_h || 0;
    if (document.getElementById("spec-p2-v")) document.getElementById("spec-p2-v").value = row.p2_v || 0;

    if (document.getElementById("spec-p3-k")) document.getElementById("spec-p3-k").value = row.p3_k || 0;
    if (document.getElementById("spec-p3-h")) document.getElementById("spec-p3-h").value = row.p3_h || 0;
    if (document.getElementById("spec-p3-v")) document.getElementById("spec-p3-v").value = row.p3_v || 0;

    if (document.getElementById("spec-p4-k")) document.getElementById("spec-p4-k").value = row.p4_k || 0;
    if (document.getElementById("spec-p4-h")) document.getElementById("spec-p4-h").value = row.p4_h || 0;
    if (document.getElementById("spec-p4-v")) document.getElementById("spec-p4-v").value = row.p4_v || 0;

    // Nội dung YCCĐ hiện tại
    var spec = resolveRowSpecification(row);
    var knowText = row.knowDesc || spec.knowDesc;
    var underText = row.underDesc || spec.underDesc;
    var applyText = row.applyDesc || spec.applyDesc;

    if (document.getElementById("spec-picker-know-textarea")) document.getElementById("spec-picker-know-textarea").value = (knowText !== "—") ? knowText : "";
    if (document.getElementById("spec-picker-under-textarea")) document.getElementById("spec-picker-under-textarea").value = (underText !== "—") ? underText : "";
    if (document.getElementById("spec-picker-apply-textarea")) document.getElementById("spec-picker-apply-textarea").value = (applyText !== "—") ? applyText : "";

    // Render checklist YCCĐ
    renderSpecChecklists(matchedLessonId, knowText, underText, applyText);

    // Render danh sách động các ý Phần II (Đúng/Sai) - hỗ trợ cả Cùng 1 bài và Liên bài
    renderSpecP2DynamicSection(row, spec);

    modal.classList.remove("hidden");
}

function renderSpecP2DynamicSection(row, spec) {
    var container = document.getElementById("spec-picker-p2-dynamic-list");
    if (!container) return;

    var curK = parseInt(document.getElementById("spec-p2-k") ? document.getElementById("spec-p2-k").value : (row.p2_k || 0)) || 0;
    var curH = parseInt(document.getElementById("spec-p2-h") ? document.getElementById("spec-p2-h").value : (row.p2_h || 0)) || 0;
    var curV = parseInt(document.getElementById("spec-p2-v") ? document.getElementById("spec-p2-v").value : (row.p2_v || 0)) || 0;
    var totalP2 = curK + curH + curV;

    if (totalP2 === 0) {
        container.innerHTML = "<div class='text-center py-4 bg-white rounded-lg border border-dashed border-gray-300 text-gray-500 text-xs'>" +
            "<i class='fa-regular fa-circle-question text-gray-400 text-lg mb-1 block'></i>" +
            "Bài học này hiện không có câu hỏi Phần II (Đúng/Sai).<br/>" +
            "Bấm vào các nút phía trên để thêm: <strong>Trọn vẹn 4 ý (Cùng bài)</strong> hoặc <strong>1-2 ý liên bài</strong>." +
            "</div>";
        return;
    }

    var tempRow = Object.assign({}, row, { p2_k: curK, p2_h: curH, p2_v: curV });
    var stmts = (row.p2_statements && row.p2_statements.length === totalP2)
        ? row.p2_statements
        : resolveRowP2Statements(tempRow, spec, matrixRows);

    var html = stmts.map(function(st, sIdx) {
        var lvlClass = (st.level === "Biết") ? "bg-emerald-50 text-emerald-900 border-emerald-300"
                     : (st.level === "Hiểu") ? "bg-blue-50 text-blue-900 border-blue-300"
                     : "bg-purple-50 text-purple-900 border-purple-300";

        var badgeBg = (st.level === "Biết") ? "bg-emerald-600"
                    : (st.level === "Hiểu") ? "bg-blue-600"
                    : "bg-purple-600";

        var qNum = st.questionIndex || 1;
        var letter = st.letter || ["a", "b", "c", "d"][sIdx % 4];

        return "<div class='p-2.5 bg-white rounded-lg border border-blue-200 shadow-xs space-y-2 spec-p2-item-card' data-idx='" + sIdx + "'>" +
            "<div class='flex flex-wrap items-center justify-between gap-2 border-b border-gray-100 pb-1.5'>" +
            "<div class='flex items-center gap-1.5'>" +
            "<span class='w-5 h-5 rounded-full " + badgeBg + " text-white inline-flex items-center justify-center text-xs font-bold'>" + letter + "</span>" +
            "<span class='font-bold text-gray-800 text-xs'>Phần II:</span>" +
            "<select class='spec-p2-qnum text-xs border border-gray-300 rounded px-1.5 py-0.5 font-bold text-blue-900 bg-blue-50' title='Thuộc câu số mấy trong đề'>" +
            "<option value='1'" + (qNum === 1 ? " selected" : "") + ">Câu 1</option>" +
            "<option value='2'" + (qNum === 2 ? " selected" : "") + ">Câu 2</option>" +
            "<option value='3'" + (qNum === 3 ? " selected" : "") + ">Câu 3</option>" +
            "<option value='4'" + (qNum === 4 ? " selected" : "") + ">Câu 4</option>" +
            "</select>" +
            "<select class='spec-p2-letter text-xs border border-gray-300 rounded px-1.5 py-0.5 font-bold text-emerald-900 bg-emerald-50' title='Lệnh hỏi'>" +
            "<option value='a'" + (letter === "a" ? " selected" : "") + ">Ý a</option>" +
            "<option value='b'" + (letter === "b" ? " selected" : "") + ">Ý b</option>" +
            "<option value='c'" + (letter === "c" ? " selected" : "") + ">Ý c</option>" +
            "<option value='d'" + (letter === "d" ? " selected" : "") + ">Ý d</option>" +
            "</select>" +
            "</div>" +
            "<div class='flex items-center gap-2'>" +
            "<span class='text-[11px] font-bold text-gray-600'>Mức độ:</span>" +
            "<select class='spec-p2-level text-xs border border-gray-300 rounded px-2 py-0.5 font-bold " + lvlClass + "' onchange='onSpecP2LevelChange(this, " + sIdx + ")'>" +
            "<option value='Biết'" + (st.level === "Biết" ? " selected" : "") + ">Biết (Nhận biết)</option>" +
            "<option value='Hiểu'" + (st.level === "Hiểu" ? " selected" : "") + ">Hiểu (Thông hiểu)</option>" +
            "<option value='Vận dụng'" + (st.level === "Vận dụng" ? " selected" : "") + ">Vận dụng</option>" +
            "</select>" +
            "</div>" +
            "</div>" +
            "<div class='flex flex-col sm:flex-row items-start sm:items-center gap-2'>" +
            "<div class='flex items-center gap-1 shrink-0'>" +
            "<span class='text-[11px] font-bold text-brand-900'>Mã chỉ báo:</span>" +
            "<input type='text' class='spec-p2-indicator w-20 text-xs px-2 py-1 border border-brand-300 rounded font-mono font-bold text-brand-700 uppercase bg-amber-50/50' value='" + (st.indicator || "HH1.1") + "' placeholder='HH1.1'>" +
            "</div>" +
            "<div class='flex-grow w-full'>" +
            "<input type='text' class='spec-p2-desc w-full text-xs px-2 py-1 border border-gray-300 rounded text-gray-700' value='" + (st.desc || "").replace(/'/g, "&#39;") + "' placeholder='YCCĐ tóm tắt đánh giá cho ý này...'>" +
            "</div>" +
            "</div>" +
            "</div>";
    }).join("");

    container.innerHTML = html;
}

function onSpecP2CountsInputChange() {
    if (currentEditingSpecRowIndex < 0 || currentEditingSpecRowIndex >= matrixRows.length) return;
    var row = matrixRows[currentEditingSpecRowIndex];
    var spec = resolveRowSpecification(row);
    delete row.p2_statements;
    renderSpecP2DynamicSection(row, spec);
}

function onSpecP2LevelChange(selectEl, sIdx) {
    var val = selectEl.value;
    if (val === "Biết") selectEl.className = "spec-p2-level text-xs border border-gray-300 rounded px-2 py-0.5 font-bold bg-emerald-50 text-emerald-900 border-emerald-300";
    else if (val === "Hiểu") selectEl.className = "spec-p2-level text-xs border border-gray-300 rounded px-2 py-0.5 font-bold bg-blue-50 text-blue-900 border-blue-300";
    else selectEl.className = "spec-p2-level text-xs border border-gray-300 rounded px-2 py-0.5 font-bold bg-purple-50 text-purple-900 border-purple-300";

    var cards = document.querySelectorAll(".spec-p2-item-card");
    var k = 0, h = 0, v = 0;
    cards.forEach(function(c) {
        var lvl = c.querySelector(".spec-p2-level").value;
        if (lvl === "Biết") k++;
        else if (lvl === "Hiểu") h++;
        else if (lvl === "Vận dụng") v++;
    });
    if (document.getElementById("spec-p2-k")) document.getElementById("spec-p2-k").value = k;
    if (document.getElementById("spec-p2-h")) document.getElementById("spec-p2-h").value = h;
    if (document.getElementById("spec-p2-v")) document.getElementById("spec-p2-v").value = v;
}

function setRowP2Mode(mode) {
    if (currentEditingSpecRowIndex < 0 || currentEditingSpecRowIndex >= matrixRows.length) return;
    var row = matrixRows[currentEditingSpecRowIndex];
    var spec = resolveRowSpecification(row);

    var kCode = (spec.kCodes && spec.kCodes[0]) || "HH1.1";
    var uCode1 = (spec.uCodes && spec.uCodes[0]) || "HH1.2";
    var uCode2 = (spec.uCodes && spec.uCodes[1]) || (spec.uCodes && spec.uCodes[0]) || "HH1.4";
    var aCode = (spec.aCodes && spec.aCodes[0]) || "HH1.6";

    var globalCursor = 0;
    for (var r = 0; r < matrixRows.length; r++) {
        if (matrixRows[r] === row) break;
        var rCnt = (matrixRows[r].p2_k || 0) + (matrixRows[r].p2_h || 0) + (matrixRows[r].p2_v || 0);
        globalCursor += rCnt;
    }

    var setInputs = function(k, h, v) {
        if (document.getElementById("spec-p2-k")) document.getElementById("spec-p2-k").value = k;
        if (document.getElementById("spec-p2-h")) document.getElementById("spec-p2-h").value = h;
        if (document.getElementById("spec-p2-v")) document.getElementById("spec-p2-v").value = v;
    };

    var letters = ["a", "b", "c", "d"];

    if (mode === "full4") {
        setInputs(1, 2, 1);
        var qNum = Math.floor(globalCursor / 4) + 1;
        row.p2_statements = [
            { questionIndex: qNum, letter: "a", label: "Ý a", shortLabel: "C" + qNum + ".a", fullLabel: "Câu " + qNum + ".a", subLabel: "Ý a", level: "Biết", indicator: kCode, desc: spec.knowDesc && spec.knowDesc !== "—" ? spec.knowDesc.slice(0, 80) : "Nhận biết khái niệm, quy tắc hóa học." },
            { questionIndex: qNum, letter: "b", label: "Ý b", shortLabel: "C" + qNum + ".b", fullLabel: "Câu " + qNum + ".b", subLabel: "Ý b", level: "Hiểu", indicator: uCode1, desc: spec.underDesc && spec.underDesc !== "—" ? spec.underDesc.slice(0, 80) : "Giải thích hiện tượng, quy luật biến đổi." },
            { questionIndex: qNum, letter: "c", label: "Ý c", shortLabel: "C" + qNum + ".c", fullLabel: "Câu " + qNum + ".c", subLabel: "Ý c", level: "Hiểu", indicator: uCode2, desc: spec.underDesc && spec.underDesc !== "—" ? spec.underDesc.slice(0, 80) : "Viết PTHH, xác định tính chất." },
            { questionIndex: qNum, letter: "d", label: "Ý d", shortLabel: "C" + qNum + ".d", fullLabel: "Câu " + qNum + ".d", subLabel: "Ý d", level: "Vận dụng", indicator: aCode, desc: spec.applyDesc && spec.applyDesc !== "—" ? spec.applyDesc.slice(0, 80) : "Vận dụng giải bài toán thực tiễn hoặc tính toán." }
        ];
    } else if (mode === "single1") {
        var globalIdx = globalCursor;
        var qNum = Math.floor(globalIdx / 4) + 1;
        var letter = letters[globalIdx % 4];
        var defaultLvl = (globalIdx % 4 === 0) ? "Biết" : ((globalIdx % 4 === 3) ? "Vận dụng" : "Hiểu");
        var defaultInd = (defaultLvl === "Biết") ? kCode : ((defaultLvl === "Hiểu") ? uCode1 : aCode);
        var defaultDesc = (defaultLvl === "Biết") ? spec.knowDesc : ((defaultLvl === "Hiểu") ? spec.underDesc : spec.applyDesc);

        setInputs(defaultLvl === "Biết" ? 1 : 0, defaultLvl === "Hiểu" ? 1 : 0, defaultLvl === "Vận dụng" ? 1 : 0);
        row.p2_statements = [
            {
                questionIndex: qNum,
                letter: letter,
                label: "C" + qNum + "." + letter,
                shortLabel: "C" + qNum + "." + letter,
                fullLabel: "Câu " + qNum + "." + letter,
                subLabel: "Ý " + letter,
                level: defaultLvl,
                indicator: defaultInd,
                desc: (defaultDesc && defaultDesc !== "—") ? defaultDesc.slice(0, 80) : "Đánh giá nhận định theo YCCĐ bài này."
            }
        ];
    } else if (mode === "split2") {
        var globalIdx1 = globalCursor;
        var globalIdx2 = globalCursor + 1;
        var qNum1 = Math.floor(globalIdx1 / 4) + 1;
        var qNum2 = Math.floor(globalIdx2 / 4) + 1;
        var letter1 = letters[globalIdx1 % 4];
        var letter2 = letters[globalIdx2 % 4];

        setInputs(1, 1, 0);
        row.p2_statements = [
            {
                questionIndex: qNum1,
                letter: letter1,
                label: "C" + qNum1 + "." + letter1,
                shortLabel: "C" + qNum1 + "." + letter1,
                fullLabel: "Câu " + qNum1 + "." + letter1,
                subLabel: "Ý " + letter1,
                level: "Biết",
                indicator: kCode,
                desc: (spec.knowDesc && spec.knowDesc !== "—") ? spec.knowDesc.slice(0, 80) : "Nhận biết khái niệm, tính chất bài này."
            },
            {
                questionIndex: qNum2,
                letter: letter2,
                label: "C" + qNum2 + "." + letter2,
                shortLabel: "C" + qNum2 + "." + letter2,
                fullLabel: "Câu " + qNum2 + "." + letter2,
                subLabel: "Ý " + letter2,
                level: "Hiểu",
                indicator: uCode1,
                desc: (spec.underDesc && spec.underDesc !== "—") ? spec.underDesc.slice(0, 80) : "Thông hiểu bản chất hiện tượng bài này."
            }
        ];
    } else if (mode === "clear") {
        setInputs(0, 0, 0);
        delete row.p2_statements;
    }

    renderSpecP2DynamicSection(row, spec);
}

function autoFillStandardP2Statements() {
    setRowP2Mode("full4");
}

function closeSpecPickerModal() {
    var modal = document.getElementById("modal-spec-picker");
    if (modal) modal.classList.add("hidden");
    currentEditingSpecRowIndex = -1;
}

function renderSpecChecklists(lessonId, knowText, underText, applyText) {
    var lessons = (typeof KNTT_LESSONS_STANDARDS !== "undefined" && KNTT_LESSONS_STANDARDS[currentGrade]) ? KNTT_LESSONS_STANDARDS[currentGrade] : [];
    var lesson = lessons.find(function(l) { return l.id === lessonId; });

    var knowContainer = document.getElementById("spec-picker-know-checklist");
    var underContainer = document.getElementById("spec-picker-under-checklist");
    var applyContainer = document.getElementById("spec-picker-apply-checklist");

    if (!lesson) {
        if (knowContainer) knowContainer.innerHTML = "<p class='text-gray-400 italic text-[11px]'>Chọn một bài học ở mục 1 để xem danh sách gợi ý YCCĐ chuẩn SGK KNTT.</p>";
        if (underContainer) underContainer.innerHTML = "<p class='text-gray-400 italic text-[11px]'>Chọn một bài học ở mục 1 để xem danh sách gợi ý YCCĐ chuẩn SGK KNTT.</p>";
        if (applyContainer) applyContainer.innerHTML = "<p class='text-gray-400 italic text-[11px]'>Chọn một bài học ở mục 1 để xem danh sách gợi ý YCCĐ chuẩn SGK KNTT.</p>";
        return;
    }

    if (knowContainer) {
        knowContainer.innerHTML = lesson.knowItems.map(function(item, idx) {
            var isChecked = knowText && (knowText.indexOf(item.text.slice(0, 20)) !== -1 || knowText.indexOf(item.code) !== -1);
            return "<label class='flex items-start gap-2 p-1 hover:bg-emerald-50 rounded cursor-pointer'>" +
                "<input type='checkbox' data-level='know' data-idx='" + idx + "' " + (isChecked ? "checked" : "") + " onchange='onSpecCheckboxToggle(\"know\")' class='mt-0.5 rounded text-emerald-600 focus:ring-emerald-500'>" +
                "<span class='text-gray-800 leading-tight'>" + item.text + " <span class='font-bold text-emerald-800 text-[10px] bg-emerald-100 px-1 rounded border border-emerald-300'>[" + item.code + "]</span></span>" +
                "</label>";
        }).join("");
    }

    if (underContainer) {
        underContainer.innerHTML = lesson.underItems.map(function(item, idx) {
            var isChecked = underText && (underText.indexOf(item.text.slice(0, 20)) !== -1 || underText.indexOf(item.code) !== -1);
            return "<label class='flex items-start gap-2 p-1 hover:bg-blue-50 rounded cursor-pointer'>" +
                "<input type='checkbox' data-level='under' data-idx='" + idx + "' " + (isChecked ? "checked" : "") + " onchange='onSpecCheckboxToggle(\"under\")' class='mt-0.5 rounded text-blue-600 focus:ring-blue-500'>" +
                "<span class='text-gray-800 leading-tight'>" + item.text + " <span class='font-bold text-blue-800 text-[10px] bg-blue-100 px-1 rounded border border-blue-300'>[" + item.code + "]</span></span>" +
                "</label>";
        }).join("");
    }

    if (applyContainer) {
        applyContainer.innerHTML = lesson.applyItems.map(function(item, idx) {
            var isChecked = applyText && (applyText.indexOf(item.text.slice(0, 20)) !== -1 || applyText.indexOf(item.code) !== -1);
            return "<label class='flex items-start gap-2 p-1 hover:bg-purple-50 rounded cursor-pointer'>" +
                "<input type='checkbox' data-level='apply' data-idx='" + idx + "' " + (isChecked ? "checked" : "") + " onchange='onSpecCheckboxToggle(\"apply\")' class='mt-0.5 rounded text-purple-600 focus:ring-purple-500'>" +
                "<span class='text-gray-800 leading-tight'>" + item.text + " <span class='font-bold text-purple-800 text-[10px] bg-purple-100 px-1 rounded border border-purple-300'>[" + item.code + "]</span></span>" +
                "</label>";
        }).join("");
    }
}

function onSpecPickerLessonSelectChange() {
    var select = document.getElementById("spec-picker-lesson-select");
    var lessonId = select ? select.value : "";
    if (!lessonId || lessonId === "custom") return;

    var lessons = (typeof KNTT_LESSONS_STANDARDS !== "undefined" && KNTT_LESSONS_STANDARDS[currentGrade]) ? KNTT_LESSONS_STANDARDS[currentGrade] : [];
    var lesson = lessons.find(function(l) { return l.id === lessonId; });
    if (!lesson) return;

    var topicInput = document.getElementById("spec-picker-topic-input");
    if (topicInput) topicInput.value = lesson.title;

    var kTexts = lesson.knowItems.map(function(i) { return i.text + " (" + i.code + ")"; }).join(" ");
    var uTexts = lesson.underItems.map(function(i) { return i.text + " (" + i.code + ")"; }).join(" ");
    var aTexts = lesson.applyItems.map(function(i) { return i.text + " (" + i.code + ")"; }).join(" ");

    if (document.getElementById("spec-picker-know-textarea")) document.getElementById("spec-picker-know-textarea").value = kTexts;
    if (document.getElementById("spec-picker-under-textarea")) document.getElementById("spec-picker-under-textarea").value = uTexts;
    if (document.getElementById("spec-picker-apply-textarea")) document.getElementById("spec-picker-apply-textarea").value = aTexts;

    renderSpecChecklists(lessonId, kTexts, uTexts, aTexts);
}

function onSpecCheckboxToggle(level) {
    var select = document.getElementById("spec-picker-lesson-select");
    var lessonId = select ? select.value : "";
    var lessons = (typeof KNTT_LESSONS_STANDARDS !== "undefined" && KNTT_LESSONS_STANDARDS[currentGrade]) ? KNTT_LESSONS_STANDARDS[currentGrade] : [];
    var lesson = lessons.find(function(l) { return l.id === lessonId; });
    if (!lesson) return;

    var checkboxes = document.querySelectorAll("#spec-picker-" + level + "-checklist input[type='checkbox']");
    var items = (level === "know") ? lesson.knowItems : ((level === "under") ? lesson.underItems : lesson.applyItems);
    var selectedSentences = [];

    checkboxes.forEach(function(cb, idx) {
        if (cb.checked && items[idx]) {
            selectedSentences.push(items[idx].text + " (" + items[idx].code + ")");
        }
    });

    var textarea = document.getElementById("spec-picker-" + level + "-textarea");
    if (textarea) {
        textarea.value = selectedSentences.join(" ");
    }
}

function insertBadgeToSpec(level, badge) {
    var textarea = document.getElementById("spec-picker-" + level + "-textarea");
    if (!textarea) return;
    var cur = textarea.value.trim();
    if (cur.length > 0 && cur[cur.length - 1] !== ".") {
        textarea.value = cur + " (" + badge + ").";
    } else {
        textarea.value = cur + " (" + badge + ")";
    }
    textarea.focus();
}

function saveSpecPickerChanges() {
    if (currentEditingSpecRowIndex < 0 || currentEditingSpecRowIndex >= matrixRows.length) return;
    var row = matrixRows[currentEditingSpecRowIndex];

    var topicVal = document.getElementById("spec-picker-topic-input").value.trim();
    if (!topicVal) {
        alert("Vui lòng nhập tên chủ đề / bài học!");
        return;
    }
    row.topic = topicVal;
    row.knowDesc = document.getElementById("spec-picker-know-textarea").value.trim();
    row.underDesc = document.getElementById("spec-picker-under-textarea").value.trim();
    row.applyDesc = document.getElementById("spec-picker-apply-textarea").value.trim();

    row.p1_k = parseInt(document.getElementById("spec-p1-k").value) || 0;
    row.p1_h = parseInt(document.getElementById("spec-p1-h").value) || 0;
    row.p1_v = parseInt(document.getElementById("spec-p1-v").value) || 0;

    row.p2_k = parseInt(document.getElementById("spec-p2-k").value) || 0;
    row.p2_h = parseInt(document.getElementById("spec-p2-h").value) || 0;
    row.p2_v = parseInt(document.getElementById("spec-p2-v").value) || 0;

    row.p3_k = parseInt(document.getElementById("spec-p3-k").value) || 0;
    row.p3_h = parseInt(document.getElementById("spec-p3-h").value) || 0;
    row.p3_v = parseInt(document.getElementById("spec-p3-v").value) || 0;

    row.p4_k = parseInt(document.getElementById("spec-p4-k").value) || 0;
    row.p4_h = parseInt(document.getElementById("spec-p4-h").value) || 0;
    row.p4_v = parseInt(document.getElementById("spec-p4-v").value) || 0;

    // Lưu chi tiết các ý của Phần II (Đúng/Sai) theo CV 7991 (Hỗ trợ Cùng bài & Liên bài)
    var savedP2 = [];
    var p2Cards = document.querySelectorAll(".spec-p2-item-card");
    p2Cards.forEach(function(card, idx) {
        var qNumEl = card.querySelector(".spec-p2-qnum");
        var ltrEl = card.querySelector(".spec-p2-letter");
        var lvlEl = card.querySelector(".spec-p2-level");
        var indEl = card.querySelector(".spec-p2-indicator");
        var descEl = card.querySelector(".spec-p2-desc");

        var qNum = qNumEl ? parseInt(qNumEl.value) || 1 : (Math.floor(idx / 4) + 1);
        var ltr = ltrEl ? ltrEl.value : ["a", "b", "c", "d"][idx % 4];
        var lvl = lvlEl ? lvlEl.value : "Hiểu";
        var ind = (indEl && indEl.value.trim()) ? indEl.value.trim().toUpperCase() : "HH1.1";
        var desc = descEl ? descEl.value.trim() : "";

        var isSingle = (p2Cards.length === 4 && qNumEl && qNumEl.value === p2Cards[0].querySelector(".spec-p2-qnum").value);

        savedP2.push({
            questionIndex: qNum,
            letter: ltr,
            label: isSingle ? ("Ý " + ltr) : ("C" + qNum + "." + ltr),
            shortLabel: "C" + qNum + "." + ltr,
            fullLabel: "Câu " + qNum + "." + ltr,
            subLabel: "Ý " + ltr,
            level: lvl,
            indicator: ind,
            desc: desc,
            isIntegrated: !isSingle
        });
    });

    if (savedP2.length > 0) {
        row.p2_statements = savedP2;
        row.p2_k = savedP2.filter(function(s) { return s.level === "Biết"; }).length;
        row.p2_h = savedP2.filter(function(s) { return s.level === "Hiểu"; }).length;
        row.p2_v = savedP2.filter(function(s) { return s.level === "Vận dụng"; }).length;
    } else {
        delete row.p2_statements;
        row.p2_k = parseInt(document.getElementById("spec-p2-k").value) || 0;
        row.p2_h = parseInt(document.getElementById("spec-p2-h").value) || 0;
        row.p2_v = parseInt(document.getElementById("spec-p2-v").value) || 0;
    }

    var allText = row.knowDesc + " " + row.underDesc + " " + row.applyDesc;
    var matches = allText.match(/HH\d\.\d/g);
    row.codes = matches ? Array.from(new Set(matches)) : ["HH1.1"];

    renderMatrixEditor();
    renderSpecificationTable();
    recalculateMatrixTotals();
    closeSpecPickerModal();

    alert("Đã lưu và cập nhật mức độ đánh giá, YCCĐ, chỉ báo và 4 ý Đ/S cho bài: " + row.topic);
}

function openAddNewTopicWithPicker() {
    var newRow = {
        topic: "Chủ đề / Bài học mới",
        p1_k: 1, p1_h: 1, p1_v: 0,
        p2_k: 0, p2_h: 0, p2_v: 0,
        p3_k: 0, p3_h: 0, p3_v: 0,
        p4_k: 0, p4_h: 0, p4_v: 0,
        codes: ["HH1.1", "HH1.2"],
        knowDesc: "",
        underDesc: "",
        applyDesc: ""
    };
    matrixRows.push(newRow);
    renderMatrixEditor();
    renderSpecificationTable();
    recalculateMatrixTotals();
    openSpecPicker(matrixRows.length - 1);
}

/* Ghi đè hàm hiển thị Ngân hàng Chỉ báo YCCĐ (Tab 2) để hiển thị đầy đủ mức độ nhận thức & chỉ báo KNTT */
function getIndicatorLevel(item) {
    if (item.level) return item.level;
    var code = item.code || "";
    var outcome = (item.outcome || "").toLowerCase();
    if (code.startsWith("HH3") || /vận dụng|tính được|thiết kế|thực hiện được|áp dụng/i.test(outcome)) {
        return "Vận dụng";
    }
    if (code === "HH1.1" || code === "HH1.2") {
        return "Biết";
    }
    return "Hiểu";
}

function renderIndicators() {
    var gradeFilterEl = document.getElementById("filter-grade");
    var compFilterEl = document.getElementById("filter-comp");
    var levelFilterEl = document.getElementById("filter-level");
    var keyFilterEl = document.getElementById("filter-keyword");
    var container = document.getElementById("indicator-cards-container");
    if (!container) return;

    var gradeFilter = gradeFilterEl ? gradeFilterEl.value : "ALL";
    var compFilter = compFilterEl ? compFilterEl.value : "ALL";
    var levelFilter = levelFilterEl ? levelFilterEl.value : "ALL";
    var keyFilter = keyFilterEl ? keyFilterEl.value.toLowerCase() : "";

    var dataSource = (typeof INDICATORS_DATABASE !== "undefined") ? [...INDICATORS_DATABASE] : [];
    if (typeof OFFICIAL_KNTT_INDICATORS !== "undefined" && Array.isArray(OFFICIAL_KNTT_INDICATORS)) {
        var existingKeys = new Set(dataSource.map(function(i) { return i.grade + "_" + i.code + "_" + (i.outcome || "").slice(0, 30); }));
        OFFICIAL_KNTT_INDICATORS.forEach(function(item) {
            var k = item.grade + "_" + item.code + "_" + (item.outcome || "").slice(0, 30);
            if (!existingKeys.has(k)) {
                dataSource.push(item);
                existingKeys.add(k);
            }
        });
    }

    var filtered = dataSource.filter(function(item) {
        if (gradeFilter !== "ALL" && item.grade !== gradeFilter) return false;
        if (compFilter !== "ALL" && item.comp !== compFilter) return false;
        var lvl = item.level || getIndicatorLevel(item);
        if (levelFilter !== "ALL" && lvl !== levelFilter) return false;
        if (keyFilter) {
            var matchTopic = item.topic && item.topic.toLowerCase().indexOf(keyFilter) !== -1;
            var matchOutcome = item.outcome && item.outcome.toLowerCase().indexOf(keyFilter) !== -1;
            var matchCode = item.code && item.code.toLowerCase().indexOf(keyFilter) !== -1;
            var matchLevel = lvl && lvl.toLowerCase().indexOf(keyFilter) !== -1;
            if (!matchTopic && !matchOutcome && !matchCode && !matchLevel) return false;
        }
        return true;
    });

    var countText = document.getElementById("indicator-count-text");
    if (countText) {
        countText.innerText = "Hiển thị: " + filtered.length + " Yêu cầu cần đạt (Chuẩn SGK KNTT & CV 7991)";
    }

    if (filtered.length === 0) {
        container.innerHTML = "<div class='text-center text-xs text-gray-500 py-8'>Không tìm thấy yêu cầu cần đạt khớp với bộ lọc.</div>";
        return;
    }

    container.innerHTML = filtered.map(function(item) {
        var lvl = item.level || getIndicatorLevel(item);
        var levelBadge = "";
        if (lvl === "Biết" || lvl === "Nhận biết") {
            levelBadge = "<span class='bg-emerald-100 text-emerald-800 border border-emerald-300 text-xs px-2.5 py-0.5 rounded-full font-bold flex items-center gap-1 shadow-sm'><i class='fa-solid fa-check text-[10px]'></i> Biết</span>";
        } else if (lvl === "Hiểu" || lvl === "Thông hiểu") {
            levelBadge = "<span class='bg-blue-100 text-blue-800 border border-blue-300 text-xs px-2.5 py-0.5 rounded-full font-bold flex items-center gap-1 shadow-sm'><i class='fa-solid fa-brain text-[10px]'></i> Hiểu</span>";
        } else if (lvl === "Vận dụng" || lvl === "VD") {
            levelBadge = "<span class='bg-purple-100 text-purple-800 border border-purple-300 text-xs px-2.5 py-0.5 rounded-full font-bold flex items-center gap-1 shadow-sm'><i class='fa-solid fa-bolt text-[10px]'></i> Vận dụng</span>";
        }

        var lessonHtml = item.lesson ? ("<span class='text-xs text-gray-500 font-medium italic'>• " + item.lesson + "</span>") : "";

        return "<div class='bg-white p-4 rounded-xl border border-gray-200 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-3 hover:border-brand-300 hover:shadow-md transition'>" +
            "<div class='space-y-1.5'>" +
            "<div class='flex flex-wrap items-center gap-2'>" +
            "<span class='bg-brand-50 text-brand-700 text-xs px-2 py-0.5 rounded font-bold border border-brand-200'>Lớp " + item.grade + "</span>" +
            levelBadge +
            "<span class='text-xs font-bold text-gray-800'>" + item.topic + "</span>" +
            lessonHtml +
            "</div>" +
            "<p class='text-xs text-gray-700 leading-relaxed'>" + item.outcome + "</p>" +
            "</div>" +
            "<div class='flex items-center gap-2 shrink-0'>" +
            "<span class='bg-amber-100 text-amber-900 text-xs px-2.5 py-1 rounded-full font-bold border border-amber-300 shadow-sm'>Mã: " + item.code + "</span>" +
            "</div></div>";
    }).join("");

    if (typeof triggerKaTeX === "function") {
        triggerKaTeX();
    }
}

/* 2. BỘ CHUYỂN ĐỔI CÔNG THỨC HÓA HỌC & TOÁN HỌC SANG WORD (.DOC) KHÔNG LỖI FONT
   Xử lý triệt để \ce{...}, đồng vị ^A_Z X, phân số \frac{a}{b}, ion, mũi tên phản ứng, chỉ số dưới/trên
*/
function formatChemSegmentForWord(chemStr) {
    var c = String(chemStr);

    // Đồng vị nguyên tử: ^39_19K hoặc ^{39}_{19}K
    c = c.replace(/\^\{?(\d+)\}?_\{?(\d+)\}?([A-Z][a-z]?)/g, "<sup>$1</sup><sub>$2</sub>$3");
    c = c.replace(/_\{?(\d+)\}?\^\{?(\d+)\}?([A-Z][a-z]?)/g, "<sub>$1</sub><sup>$2</sup>$3");

    // Mũi tên phản ứng kèm điều kiện trong \ce{...}
    c = c.replace(/<=>\[([^\]]+)\]/g, " ⇌ ($1) ");
    c = c.replace(/<=>/g, " ⇌ ");
    c = c.replace(/->\[t\^\\circ\]/g, " ──(t°)──> ");
    c = c.replace(/->\[([^\]]+)\]/g, function(_, cond) {
        return " ──(" + cond.replace(/\\circ/g, "°").replace(/\^°/g, "°") + ")──> ";
    });
    c = c.replace(/->/g, " ──> ");

    // Ký hiệu kết tủa / khí trong mhchem: (v) hoặc v ở cuối, (^) hoặc ^ ở cuối
    c = c.replace(/\s+\^(\s|$)/g, " ↑$1");
    c = c.replace(/\s+v(\s|$)/g, " ↓$1");

    // Điện tích ion: Fe^3+, SO4^2-, Cu^2+, Ag^+, Cl^-, ^{2+}, ^{2-}
    c = c.replace(/\^\{(\d*[+-])\}/g, "<sup>$1</sup>");
    c = c.replace(/\^(\d+[+-])/g, "<sup>$1</sup>");
    c = c.replace(/\^([+-])(?![0-9])/g, "<sup>$1</sup>");

    // Chỉ số dưới đã có dấu gạch dưới: _2, _{2}, _n
    c = c.replace(/_\{([^}]+)\}/g, "<sub>$1</sub>");
    c = c.replace(/_([0-9nmxa-z]+)/g, "<sub>$1</sub>");

    // Số mũ còn lại: ^2, ^{2}
    c = c.replace(/\^\{([^}]+)\}/g, "<sup>$1</sup>");
    c = c.replace(/\^(\d+)/g, "<sup>$1</sup>");

    // Chuẩn hóa điện tích ion trước khi xử lý chỉ số dưới
    c = formatChemicalIons(c);

    // Tự động nhận diện chỉ số dưới trong công thức hóa học: H2SO4, C12H22O11, (NH4)2SO4, Fe3O4...
    // Chỉ biến số thành <sub> khi đứng ngay sau chữ cái nguyên tố hoặc dấu đóng ngoặc ), ]
    c = c.replace(/([A-Za-z\)\]])(\d+)/g, function(match, prevChar, digits) {
        // Tránh biến đổi thẻ HTML như <sup>2</sup> hoặc <sub>2</sub>
        return prevChar + "<sub>" + digits + "</sub>";
    });

    return c;
}

function cleanLatexForWord(text) {
    if (!text) return "";
    var s = String(text);

    function stripOuterParens(str) {
        var t = str.trim();
        while (t.startsWith("(") && t.endsWith(")")) {
            var depth = 0;
            var wrapsAll = true;
            for (var i = 0; i < t.length - 1; i++) {
                if (t[i] === "(") depth++;
                else if (t[i] === ")") depth--;
                if (depth === 0) {
                    wrapsAll = false;
                    break;
                }
            }
            if (wrapsAll) {
                t = t.substring(1, t.length - 1).trim();
            } else {
                break;
            }
        }
        return t;
    }

    function cleanCond(cond) {
        if (!cond) return "";
        var c = cond.replace(/\\circ|\^circ|\^o|o$/gi, "°").trim();
        c = stripOuterParens(c);
        c = c.replace(/\\ge/g, "≥").replace(/\\le/g, "≤").replace(/\\approx/g, "≈");
        c = c.replace(/\\ce\{((?:[^{}]|\{[^{}]*\})*)\}/g, "$1");
        return c;
    }

    function makeChemArrow(cond) {
        if (!cond) return " → ";
        var c = cond.trim();
        if (/^\d+$/.test(c)) c = "(" + c + ")";
        return " <span class='chem-arrow' style='display:inline-block; text-align:center; vertical-align:middle; margin:0 4px; line-height:1;'>" +
               "<span class='arrow-cond' style='font-size:8.5pt; display:block; line-height:1; margin-bottom:2px; font-family:\"Times New Roman\",serif; font-weight:normal;'>" + c + "</span>" +
               "<span class='arrow-sym' style='font-size:12pt; display:block; line-height:1; font-family:\"Times New Roman\",serif;'>⎯⎯⎯→</span>" +
               "</span> ";
    }

    function makeReversibleArrow(cond) {
        if (!cond) return " ⇌ ";
        var c = cond.trim();
        if (/^\d+$/.test(c)) c = "(" + c + ")";
        return " <span class='chem-arrow' style='display:inline-block; text-align:center; vertical-align:middle; margin:0 4px; line-height:1;'>" +
               "<span class='arrow-cond' style='font-size:8.5pt; display:block; line-height:1; margin-bottom:2px; font-family:\"Times New Roman\",serif; font-weight:normal;'>" + c + "</span>" +
               "<span class='arrow-sym' style='font-size:12pt; display:block; line-height:1; font-family:\"Times New Roman\",serif;'>⇌</span>" +
               "</span> ";
    }

    // 0. Xử lý mũi tên phản ứng trước khi bảo vệ thẻ HTML và ngoặc vuông
    // (Đặc biệt: `<=>` có chứa `<` và `>` nên phải chuyển đổi trước regex bảo vệ HTML)
    s = s.replace(/\\xrightarrow\[([^\]]*)\]\{([^}]*)\}/g, function(_, below, above) {
        var a = cleanCond(above);
        var b = cleanCond(below);
        var comb = a + (a && b ? ", " : "") + b;
        return makeChemArrow(comb);
    });
    s = s.replace(/\\xrightarrow\{([^}]*)\}/g, function(_, above) {
        var a = cleanCond(above);
        return makeChemArrow(a);
    });
    s = s.replace(/->\[([^\]]*)\]/g, function(_, cond) {
        var c = cleanCond(cond);
        return makeChemArrow(c);
    });
    s = s.replace(/<=>\[([^\]]*)\]/g, function(_, cond) {
        var c = cleanCond(cond);
        return makeReversibleArrow(c);
    });
    s = s.replace(/<=>|\\rightleftharpoons/g, " ⇌ ");
    s = s.replace(/->|\\rightarrow/g, " → ");

    // Khí bay hơi và kết tủa
    s = s.replace(/\\uparrow|(?<=\b[A-Za-z0-9\)]+)\s*\^(?=\s*[\+\.\,\;\)\}\]]|\s*$)/g, "↑");
    s = s.replace(/\\downarrow|(?<=\b[A-Za-z0-9\)]+)\s+v(?=[\s,.:;)\]\}]|$)/g, "↓");

    // 0.04 Tự động chuyển đổi bảng bị phẳng hóa thành bảng HTML chuẩn
    s = autoFormatFlattenedTables(s);

    // 0.05 Tự động sửa ký hiệu đồng vị / nguyên tử bị dính liền (3919K, 6329Cu...)
    s = autoFixFlattenedIsotopes(s);

    // 0.1 Chuyển đổi ô orbital dạng ngoặc vuông thành bảng ô vuông chuẩn trước khi lưu tags
    s = formatOrbitalBoxes(s);

    // 0.06 Chuyển đổi thẻ HTML isotope <sup>A</sup><sub>Z</sub>Elem sang STACK để Word EQ xếp chồng hoàn hảo
    s = s.replace(/<sup\b[^>]*>(\d+)<\/sup>\s*<sub\b[^>]*>(\d+)<\/sub>\s*([A-Za-z]+)/gi, "###STACK-$1-$2###$3");
    s = s.replace(/<sub\b[^>]*>(\d+)<\/sub>\s*<sup\b[^>]*>(\d+)<\/sup>\s*([A-Za-z]+)/gi, "###STACK-$2-$1###$3");

    var savedHtml = [];
    s = s.replace(/<[^>]+>/g, function(m) {
        savedHtml.push(m);
        return "###HTML-" + (savedHtml.length - 1) + "###";
    });

    // 1. Mở gói các lệnh LaTeX
    s = s.replace(/\\ce\{((?:[^{}]|\{[^{}]*\})*)\}/g, "$1");
    s = s.replace(/\\(?:mathrm|text)\{([^}]*)\}/g, "$1");
    s = s.replace(/\\overline\{([^}]*)\}/g, "<span style='text-decoration: overline;'>$1</span>");
    s = s.replace(/\\frac\{((?:[^{}]|\{[^{}]*\})*)\}\{((?:[^{}]|\{[^{}]*\})*)\}/g, "($1)/($2)");

    // 2. Kí hiệu nhiệt độ và nhiệt động học trước khi xử lý mũ dấu ^
    s = s.replace(/t\^\\circ\s*C|t\^oC|t°C/gi, "t°C");
    s = s.replace(/t\^\\circ|t\^o|t°/gi, "t°");
    s = s.replace(/\^\\circ|\^circ/g, "°");

    s = s.replace(/\\Delta_r\s*H\^\\circ_\{298\}|\\Delta_r\s*H\^0_\{298\}|Δ_r\s*H°_\{298\}|ΔrH°298|ΔrH0298|Δ_rH°_298/gi, "Δ<sub>r</sub>H°<sub>298</sub>");
    s = s.replace(/\\Delta_f\s*H\^\\circ_\{298\}|\\Delta_f\s*H\^0_\{298\}|Δ_f\s*H°_\{298\}|ΔfH°298|ΔfH0298|Δ_fH°_298/gi, "Δ<sub>f</sub>H°<sub>298</sub>");
    s = s.replace(/\\Delta_r\s*H\^\\circ|ΔrH°/gi, "Δ<sub>r</sub>H°");
    s = s.replace(/\\Delta_f\s*H\^\\circ|ΔfH°/gi, "Δ<sub>f</sub>H°");
    s = s.replace(/\\Delta\s*H|ΔH/g, "ΔH");
    s = s.replace(/E\^\\circ_\{pin\}|E°_\{pin\}|E°pin/gi, "E°<sub>pin</sub>");
    s = s.replace(/E\^\\circ_\{([^}]+)\}|E°_\{([^}]+)\}/g, "E°<sub>$1</sub>");
    s = s.replace(/E\^\\circ/g, "E°");
    s = s.replace(/\bK_C\b|\bK_c\b/g, "K<sub>C</sub>");
    s = s.replace(/\bE_b\b/g, "E<sub>b</sub>");

    // 3. Đồng vị hạt nhân (xếp chồng số khối A ở trên, số hiệu nguyên tử Z ở dưới chuẩn IUPAC)
    s = s.replace(/\^\{?(\d+)\}?_\{?(\d+)\}?([A-Z][a-z]?)/g, "###STACK-$1-$2###$3");
    s = s.replace(/_\{?(\d+)\}?\^\{?(\d+)\}?([A-Z][a-z]?)/g, "###STACK-$2-$1###$3");
    s = s.replace(/\^\{?(\d+)\}?([A-Z][a-z]?)/g, "###SUP-$1###END###$2");

    // 4. Cấu hình electron nguyên tử (hỗ trợ cả có dấu cách lẫn viết liền như 1s22s22p63s2...)
    s = s.replace(/([1-7][spdf])\^\{?(\d{1,2})\}?/g, "$1###SUP-$2###END###");
    s = s.replace(/([1-7][spdf])(\d{1,2})(?=[1-7][spdf]|[^0-9a-zA-Z]|$)/g, "$1###SUP-$2###END###");

    // 5. Điện tích ion trên mũ
    s = s.replace(/\^\{(\d*[+-])\}/g, "###SUP-$1###END###");
    s = s.replace(/\^(\d+[+-])/g, "###SUP-$1###END###");
    s = s.replace(/\^([+-])(?![0-9])/g, "###SUP-$1###END###");
    s = s.replace(/\^\{([^}]+)\}/g, "###SUP-$1###END###");
    s = s.replace(/\^([0-9+-]+)/g, "###SUP-$1###END###");

    // 5. Chuẩn hóa điện tích ion toàn diện (đa nguyên tử & đơn nguyên tử: Li+, Na+, N3-, O2-, Mg2+, Al3+, Fe2+, Cu2+, SO4^2-, PO4^3-...)
    s = s.replace(/([0-9A-Za-z\)\]\>])[\u2013\u2014]/g, function(_, g) { return g + "-"; });

    // Ion đa nguyên tử:
    var polyIonConfigs = [
        { re: /\bSO4\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "SO###SUB-4###END###" },
        { re: /\bSO3\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "SO###SUB-3###END###" },
        { re: /\bCO3\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "CO###SUB-3###END###" },
        { re: /\bNO3\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "NO###SUB-3###END###" },
        { re: /\bNO2\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "NO###SUB-2###END###" },
        { re: /\bPO4\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "PO###SUB-4###END###" },
        { re: /\bHPO4\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "HPO###SUB-4###END###" },
        { re: /\bH2PO4\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "H###SUB-2###END###PO###SUB-4###END###" },
        { re: /\bHSO4\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "HSO###SUB-4###END###" },
        { re: /\bHSO3\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "HSO###SUB-3###END###" },
        { re: /\bHCO3\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "HCO###SUB-3###END###" },
        { re: /\bOH\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "OH" },
        { re: /\bClO4\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "ClO###SUB-4###END###" },
        { re: /\bClO3\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "ClO###SUB-3###END###" },
        { re: /\bClO2\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "ClO###SUB-2###END###" },
        { re: /\bClO\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "ClO" },
        { re: /\bMnO4\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "MnO###SUB-4###END###" },
        { re: /\bCrO4\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "CrO###SUB-4###END###" },
        { re: /\bCr2O7\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "Cr###SUB-2###END###O###SUB-7###END###" },
        { re: /\bAlO2\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "AlO###SUB-2###END###" },
        { re: /\bZnO2\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "ZnO###SUB-2###END###" },
        { re: /\bCH3COO\s*(?:\^\{?|\^)?([1-4]?[-–]|[-–][1-4]?)\}?/g, b: "CH###SUB-3###END###COO" },
        { re: /\bNH4\s*(?:\^\{?|\^)?([1-4]?[+]|[+][1-4]?)\}?/g, b: "NH###SUB-4###END###" },
        { re: /\bH3O\s*(?:\^\{?|\^)?([1-4]?[+]|[+][1-4]?)\}?/g, b: "H###SUB-3###END###O" }
    ];

    polyIonConfigs.forEach(function(item) {
        s = s.replace(item.re, function(match, rawCharge) {
            var ch = (rawCharge || "-").replace(/[\u2013\u2014]/g, "-").trim();
            if (/^[+\-]/.test(ch)) ch = ch.slice(1) + ch[0];
            if (ch === "1+" || ch === "+1") ch = "+";
            if (ch === "1-" || ch === "-1") ch = "-";
            return item.b + "###SUP-" + ch + "###END###";
        });
    });

    // Ion đơn nguyên tử (Monoatomic ions) & ion kim loại / phi kim tổng quát (M, X, R):
    var allIonElemSyms = "He|Li|Be|Ne|Na|Mg|Al|Si|Cl|Ar|Ca|Sc|Ti|Cr|Mn|Fe|Co|Ni|Cu|Zn|Ga|Ge|As|Se|Br|Kr|Rb|Sr|Zr|Nb|Mo|Tc|Ru|Rh|Pd|Ag|Cd|In|Sn|Sb|Te|Xe|Cs|Ba|La|Ce|Pr|Nd|Sm|Eu|Gd|Tb|Dy|Ho|Er|Tm|Yb|Lu|Hf|Ta|Re|Os|Ir|Pt|Au|Hg|Tl|Pb|Bi|Po|At|Rn|Fr|Ra|Ac|Th|Pa|H|B|C|N|O|F|P|S|K|V|Y|W|U|I|M|X|R";
    var monoIonRegex = new RegExp("\\b(" + allIonElemSyms + ")(?:\\^\\{?([0-9]*[+\\-\\u2013][0-9]*)\\}?|\\^?([2-4]?[+\\-\\u2013]|[+\\-\\u2013][2-4]?))(?=[\\s,.:;)\\]\\/\\?\\<]|$)", "g");

    s = s.replace(monoIonRegex, function(match, elem, caretCharge, directCharge) {
        var ch = (caretCharge || directCharge || "").replace(/[\u2013\u2014]/g, "-").trim();
        if (!ch) return match;
        if (/^[+\-]/.test(ch)) {
            var sign = ch[0];
            var num = ch.slice(1);
            ch = num + sign;
        }
        if (ch === "1+" || ch === "+1") ch = "+";
        if (ch === "1-" || ch === "-1") ch = "-";
        return elem + "###SUP-" + ch + "###END###";
    });

    // 6. Ký hiệu nguyên tử khối, phân tử khối, khối lượng mol, nồng độ, chỉ số dưới:
    // M_{Cl}, M_Cl, A_{r}, A_r, M_{r}, M_r, m_{dd}, m_{ct}, V_{dd}, C_M, %m_{...}
    s = s.replace(/\b([MAVmnC])_\{([^}]+)\}/g, "$1###SUB-$2###END###");
    s = s.replace(/\b([MAVmnC])_([A-Za-z0-9]+)\b/g, "$1###SUB-$2###END###");
    // Trường hợp sau span: <span ...>M</span>_{Cl} hoặc <span ...>A</span>_{Cu}
    s = s.replace(/_\{([^}]+)\}/g, "###SUB-$1###END###");
    s = s.replace(/_([A-Za-z0-9]+)\b/g, "###SUB-$1###END###");

    // 7. Tự động nhận diện và chuyển đổi chỉ số dưới trong công thức hóa học:
    // Fe2O3, H2SO4, CaCO3, KMnO4, Ba(OH)2, H2O, CO2...
    s = s.replace(/([A-Z][a-z]?|\))(\d+)(?=[A-Za-z0-9\(\)]*)/g, function(match, elem, num) {
        if (!elem) return match;
        return elem + "<sub>" + num + "</sub>";
    });

    // Khôi phục các thẻ SUP và SUB sang HTML (composite trước, đơn sau)
    s = s.replace(/###SUP-([^#]+)###SUB-([^#]+)###END###/g, "<sup>$1</sup><sub>$2</sub>");
    s = s.replace(/###SUP-([^#]+)###END###/g, "<sup>$1</sup>");
    s = s.replace(/###SUB-([^#]+)###END###/g, "<sub>$1</sub>");

    // 8. Kí hiệu toán học & làm sạch
    s = s.replace(/\\times/g, " × ");
    s = s.replace(/\\cdot/g, "·");
    s = s.replace(/\\circ/g, "°");
    s = s.replace(/\^°/g, "°");
    s = s.replace(/\\le/g, "≤");
    s = s.replace(/\\ge/g, "≥");
    s = s.replace(/\\ne/g, "≠");
    s = s.replace(/\\approx/g, "≈");
    s = s.replace(/\\implies/g, " ⇒ ");
    s = s.replace(/\\Rightarrow/g, " ⇒ ");
    s = s.replace(/\\Delta/g, "Δ");
    s = s.replace(/\\gamma/g, "γ");
    s = s.replace(/\\alpha/g, "α");
    s = s.replace(/\\beta/g, "β");
    s = s.replace(/\\pi/g, "π");
    s = s.replace(/\\sigma/g, "σ");

    // Dấu bao LaTeX
    s = s.replace(/\\\\\(/g, "").replace(/\\\\\)/g, "");
    s = s.replace(/\\\(/g, "").replace(/\\\)/g, "");
    s = s.replace(/\\\[/g, "").replace(/\\\]/g, "");
    s = s.replace(/\$\$/g, "").replace(/\$/g, "");
    s = s.replace(/\\/g, "");

    // Khôi phục các thẻ STACK sang HTML (sau khi đã dọn sạch các dấu \ của LaTeX để giữ nguyên cú pháp Word EQ)
    s = s.replace(/###STACK-([^#]+)-([^#]+)###([A-Z][a-z]?)/g, function(_, a, z, elem) {
        return "<!--[if mso]><span style='mso-element:field-begin'></span> EQ \\s\\up 8(" + a + ")\\s\\do 3(" + z + ") <span style='mso-element:field-separator'></span><sup>" + a + "</sup><sub>" + z + "</sub><span style='mso-element:field-end'></span><![endif]-->" +
               "<![if !mso]><span class='isotope-sym' style='display:inline-block; vertical-align:middle; text-align:right; line-height:1; margin-right:1px;'><sup class='iso-a' style='font-size:8.5pt; display:block; line-height:1; margin-bottom:1px; vertical-align:baseline; font-family:\"Times New Roman\",serif; font-weight:bold;'>" + a + "</sup><sub class='iso-z' style='font-size:8.5pt; display:block; line-height:1; vertical-align:baseline; font-family:\"Times New Roman\",serif; font-weight:bold;'>" + z + "</sub></span><![endif]>" + elem;
    });

    // Khôi phục thẻ HTML ban đầu
    s = s.replace(/###HTML-(\d+)###/g, function(_, idx) {
        return savedHtml[parseInt(idx, 10)] || "";
    });

    // Khôi phục thẻ ngoặc vuông metadata
    s = s.replace(/###TAG-(\d+)###/g, function(_, idx) {
        return savedTags[parseInt(idx, 10)] || "";
    });

    // Chuẩn hóa khoảng trắng quanh mũi tên phản ứng
    s = s.replace(/\s*(<span class='chem-arrow'[\s\S]*?<\/span>)\s*/g, " $1 ");
    s = s.replace(/\s+(→|⇌)\s+/g, " $1 ");

    return s;
}

var formatChemistryText = cleanLatexForWord;

function formatSolutionHtml(text) {
    if (!text) return "";
    var s = (typeof stripAuthorAndPromotion === "function") ? stripAuthorAndPromotion(text) : text;
    s = autoFormatFlattenedTables(s);
    s = cleanLatexForWord(s);
    s = s.replace(/\r\n/g, "\n");
    s = s.replace(/\n+/g, function(match, offset, fullStr) {
        var before = fullStr.slice(Math.max(0, offset - 30), offset).trim();
        var after = fullStr.slice(offset + match.length, offset + match.length + 30).trim();
        if (/>$/.test(before) && /^</.test(after)) {
            return "";
        }
        return "<br/>";
    });
    s = s.replace(/(<\/?(?:table|thead|tbody|tr|th|td|ul|ol|li)\b[^>]*>)\s*<br\s*\/?>/gi, "$1");
    s = s.replace(/<br\s*\/?>\s*(<\/?(?:table|thead|tbody|tr|th|td|ul|ol|li)\b)/gi, "$1");

    // Khử max-width/width cố định trên các bảng lồng nhau để Word không kéo giãn cột bảng đáp án
    s = s.replace(/<table\b([^>]*)>/gi, function(match, attrs) {
        if (/ao-boxes|ao-opt-row/i.test(attrs)) return match;
        var cleanAttrs = attrs
            .replace(/width\s*:\s*[^;"]+;?/gi, "")
            .replace(/max-width\s*:\s*[^;"]+;?/gi, "")
            .replace(/width=['"][^'"]*['"]/gi, "");
        return "<table " + cleanAttrs + " style='width: 100% !important; border-collapse: collapse; margin: 4px 0;' border='1'>";
    });

    return s;
}

/* Định dạng phương án trắc nghiệm khi xuất file Word, đảm bảo các ô orbital và nhãn A, B, C, D luôn nằm trên cùng 1 hàng duy nhất chuẩn đẹp */
function formatWordOption(label, optionText) {
    if (!optionText) return "<p style='margin: 2px 0;'><strong>" + label + ".</strong></p>";
    var cleanOpt = String(optionText).replace(/<[^>]+>/g, "").trim();
    if (!cleanOpt || cleanOpt === "." || cleanOpt === "và." || cleanOpt === "và" || cleanOpt === ".." || cleanOpt === "...") {
        return "<p style='margin: 2px 0;'><strong>" + label + ".</strong> [Phương án " + label + "]</p>";
    }
    var rawText = String(optionText);

    // 1. Chuyển đổi toàn bộ HTML entities mũi tên và khoảng trắng về dạng chuẩn
    var s = rawText
        .replace(/&#8593;|&uarr;/gi, "↑")
        .replace(/&#8595;|&darr;/gi, "↓")
        .replace(/&nbsp;/gi, " ");

    // 2. Chuyển đổi các thẻ HTML ao-box hoặc ao-boxes (kể cả đã qua autoSanitize) về dạng ngoặc thô [↑↓], [↑ ], [  ]
    var normalized = s
        .replace(/<table\b[^>]*class=['"]?ao-boxes['"]?[^>]*>[\s\S]*?<\/table>/gi, function(tbl) {
            var m, cells = [];
            var tdRegex = /<td\b[^>]*>([\s\S]*?)<\/td>/gi;
            while ((m = tdRegex.exec(tbl)) !== null) {
                var inner = m[1].replace(/<[^>]+>/g, "").replace(/&#8593;|&uarr;/gi, "↑").replace(/&#8595;|&darr;/gi, "↓").replace(/&nbsp;/g, " ").trim();
                cells.push("[" + (inner || "  ") + "]");
            }
            return cells.join("");
        })
        .replace(/<span\b[^>]*class=['"]?ao-box['"]?[^>]*>([\s\S]*?)<\/span>/gi, function(_, inner) {
            var innerClean = inner.replace(/<[^>]+>/g, "").replace(/&#8593;|&uarr;/gi, "↑").replace(/&#8595;|&darr;/gi, "↓").replace(/&nbsp;/g, " ").trim();
            return "[" + (innerClean || "  ") + "]";
        })
        .replace(/<td\b[^>]*class=['"]?ao-cell['"]?[^>]*>([\s\S]*?)<\/td>/gi, function(_, inner) {
            var innerClean = inner.replace(/<[^>]+>/g, "").replace(/&#8593;|&uarr;/gi, "↑").replace(/&#8595;|&darr;/gi, "↓").replace(/&nbsp;/g, " ").trim();
            return "[" + (innerClean || "  ") + "]";
        });

    var aoRegex = /(?:\[[↑↓⥮\s]{0,4}\]\s*)+/g;

    if (aoRegex.test(normalized)) {
        aoRegex.lastIndex = 0;
        var cellsHtml = "<td class='ao-lbl' style='border: none; font-weight: bold; padding-right: 6px; vertical-align: middle; white-space: nowrap; font-size: 11pt; font-family: \"Times New Roman\", serif; width: 22pt;'>" + label + ".</td>";
        var lastIndex = 0;
        var match;
        var totalWidthPt = 24;

        while ((match = aoRegex.exec(normalized)) !== null) {
            var textBefore = normalized.substring(lastIndex, match.index).trim();
            textBefore = textBefore.replace(/^(?:&nbsp;|\s)+|(?:&nbsp;|\s)+$/g, "");

            if (textBefore) {
                if (lastIndex > 0) {
                    cellsHtml += "<td class='ao-gap' style='border: none; width: 6pt; vertical-align: middle;'>&nbsp;</td>";
                    totalWidthPt += 6;
                }
                var txtWidth = Math.max(20, textBefore.length * 8);
                cellsHtml += "<td class='ao-txt' style='border: none; padding: 0 4px; vertical-align: middle; white-space: nowrap; font-size: 11pt; font-family: \"Times New Roman\", serif; width: " + txtWidth + "pt;'>" + cleanLatexForWord(textBefore) + "</td>";
                totalWidthPt += txtWidth;
            } else if (lastIndex > 0) {
                cellsHtml += "<td class='ao-gap' style='border: none; width: 6pt; vertical-align: middle;'>&nbsp;</td>";
                totalWidthPt += 6;
            }

            var singleBoxRegex = /\[([↑↓⥮\s]{0,4})\]/g;
            var bMatch;
            while ((bMatch = singleBoxRegex.exec(match[0])) !== null) {
                var raw = bMatch[1].replace(/\s+/g, "");
                var content = "&nbsp;&nbsp;";
                if (raw === "↑" || raw === "↑ ") content = "&#8593;&nbsp;";
                else if (raw === "↓" || raw === "↓ ") content = "&#8595;&nbsp;";
                else if (raw === "↑↓" || raw === "⥮") content = "&#8593;&#8595;";
                else if (raw === "↑↑") content = "&#8593;&#8593;";
                else if (raw === "↓↓") content = "&#8595;&#8595;";
                else if (raw) content = raw;

                cellsHtml += "<td class='ao-cell' width='28' style='border: 1.5pt solid black; width: 20pt; height: 20pt; min-width: 20pt; max-width: 20pt; text-align: center; vertical-align: middle; font-weight: bold; font-size: 11pt; font-family: \"Times New Roman\", serif; line-height: 18pt; padding: 0; background-color: #ffffff; color: #000000;'>" + content + "</td>";
                totalWidthPt += 22;
            }

            lastIndex = aoRegex.lastIndex;
        }

        var textAfter = normalized.substring(lastIndex).trim().replace(/^(?:&nbsp;|\s)+|(?:&nbsp;|\s)+$/g, "");
        if (textAfter) {
            var afterWidth = Math.max(20, textAfter.length * 8);
            cellsHtml += "<td class='ao-txt' style='border: none; padding: 0 4px; vertical-align: middle; white-space: nowrap; font-size: 11pt; font-family: \"Times New Roman\", serif; width: " + afterWidth + "pt;'>" + cleanLatexForWord(textAfter) + "</td>";
            totalWidthPt += afterWidth;
        }

        // Đặt thuộc tính width rõ ràng trên thẻ table để Word KHÔNG BAO GIỜ tự động kéo giãn full trang
        return "<table class='ao-opt-row' width='" + totalWidthPt + "' style='border: none; border-collapse: collapse; margin: 3px 0; width: " + totalWidthPt + "pt;'><tr>" + cellsHtml + "</tr></table>";
    }

    return "<p style='margin: 2px 0;'><strong>" + label + ".</strong> " + cleanLatexForWord(optionText) + "</p>";
}

/* 3. Tải tệp Word (.doc) với định dạng bảng biểu chuẩn Bộ GD&ĐT */
function downloadWordDocument(htmlContent, filename) {
    var header = "<html xmlns:o='urn:schemas-microsoft-com:office:office' " +
        "xmlns:w='urn:schemas-microsoft-com:office:word' " +
        "xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math' " +
        "xmlns='http://www.w3.org/TR/REC-html40'>" +
        "<head><meta charset='utf-8'>" +
        "<title>" + filename + "</title>" +
        "<style>" +
        "@page { size: 21cm 29.7cm; margin: 1.8cm 1.8cm 1.8cm 2.0cm; }" +
        "body { font-family: 'Times New Roman', Times, serif; font-size: 12pt; line-height: 1.35; color: #000; }" +
        "h1, h2, h3, h4 { font-family: 'Times New Roman', Times, serif; margin: 6px 0; }" +
        "table { border-collapse: collapse; margin-top: 6px; margin-bottom: 10px; }" +
        "table.data-table, table.header-table, table.matrix-table, table.spec-table, table.answer-table, table.chem-table { width: 100%; border-collapse: collapse; }" +
        "table.data-table { width: 100% !important; border-collapse: collapse !important; margin-top: 8px !important; margin-bottom: 12px !important; }" +
        "table.data-table th, table.data-table td { border: 1px solid #000; padding: 5px 6px; font-size: 10.5pt; text-align: left; vertical-align: middle; word-wrap: break-word !important; overflow-wrap: break-word !important; }" +
        "table.data-table th { background-color: #f2f2f2; font-weight: bold; text-align: center; }" +
        "table.answer-table { width: 100% !important; border-collapse: collapse !important; table-layout: fixed !important; mso-table-layout-alt: fixed !important; margin-top: 8px !important; margin-bottom: 12px !important; }" +
        "table.answer-table th, table.answer-table td { border: 1px solid #000; padding: 5px 6px; font-size: 10.5pt; text-align: left; vertical-align: middle; word-wrap: break-word !important; overflow-wrap: break-word !important; }" +
        "table.answer-table th { background-color: #f2f2f2; font-weight: bold; text-align: center; }" +
        "table.answer-table td table, table.data-table td table { width: 100% !important; max-width: 100% !important; table-layout: fixed !important; mso-table-layout-alt: fixed !important; border-collapse: collapse !important; margin: 4px 0 !important; }" +
        "table.answer-table td table th, table.answer-table td table td, table.data-table td table th, table.data-table td table td { padding: 3px 4px !important; font-size: 9pt !important; border: 1px solid #666 !important; word-wrap: break-word !important; }" +
        "th, td { border: 1px solid #000; padding: 5px 7px; font-size: 11pt; text-align: left; vertical-align: middle; }" +
        "th { background-color: #f2f2f2; font-weight: bold; text-align: center; }" +
        ".text-center { text-align: center; }" +
        ".text-right { text-align: right; }" +
        ".bold { font-weight: bold; }" +
        ".italic { font-style: italic; }" +
        ".header-table { width: 100%; border: none; margin-bottom: 12px; }" +
        ".header-table td { border: none; padding: 2px; vertical-align: top; }" +
        ".page-break { page-break-before: always; }" +
        "sub { font-size: 0.78em; vertical-align: sub; }" +
        "sup { font-size: 0.78em; vertical-align: super; }" +
        ".chem-arrow { display: inline-block; text-align: center; vertical-align: middle; margin: 0 4px; line-height: 1; }" +
        ".arrow-cond { font-size: 8.5pt; display: block; line-height: 1; margin-bottom: 2px; font-family: 'Times New Roman', serif; font-weight: normal; }" +
        ".arrow-sym { font-size: 12pt; display: block; line-height: 1; font-family: 'Times New Roman', serif; }" +
        ".isotope-sym { display: inline-block; vertical-align: middle; text-align: right; line-height: 1; margin-right: 1px; }" +
        ".iso-a { font-size: 8.5pt; display: block; line-height: 1; margin-bottom: 1px; font-family: 'Times New Roman', serif; }" +
        ".iso-z { font-size: 8.5pt; display: block; line-height: 1; font-family: 'Times New Roman', serif; }" +
        ".ao-group { display: inline; white-space: nowrap; margin: 0 4px; }" +
        ".ao-box { mso-border-alt: solid windowtext 1.5pt; border: 1.5pt solid #000; padding: 1pt 3pt; font-family: 'Times New Roman', serif; font-size: 11pt; font-weight: bold; display: inline; line-height: 1; }" +
        "table.ao-boxes { width: auto !important; display: inline !important; margin: 2px 4px !important; border-collapse: collapse !important; }" +
        "table.ao-boxes td { width: 18pt !important; height: 18pt !important; min-width: 18pt !important; max-width: 18pt !important; padding: 1px 2px !important; text-align: center !important; vertical-align: middle !important; font-size: 11pt !important; font-weight: bold !important; border: 1.5pt solid #000 !important; }" +
        "table.ao-opt-row { border: none !important; border-collapse: collapse !important; margin: 3px 0 !important; width: auto !important; }" +
        "table.ao-opt-row td { padding: 0 !important; vertical-align: middle !important; font-size: 11pt !important; font-family: 'Times New Roman', serif !important; }" +
        "td.ao-lbl { border: none !important; font-weight: bold !important; padding-right: 6px !important; white-space: nowrap !important; }" +
        "td.ao-txt { border: none !important; padding: 0 4px !important; white-space: nowrap !important; }" +
        "td.ao-cell { border: 1.5pt solid black !important; width: 20pt !important; height: 20pt !important; min-width: 20pt !important; max-width: 20pt !important; text-align: center !important; vertical-align: middle !important; font-weight: bold !important; font-size: 11pt !important; font-family: 'Times New Roman', serif !important; line-height: 18pt !important; padding: 0 !important; background-color: #ffffff !important; color: #000000 !important; }" +
        "td.ao-gap { border: none !important; width: 6pt !important; }" +
        "</style></head><body>";
    var footer = "</body></html>";
    var sourceHTML = header + htmlContent + footer;

    var blob = new Blob(['\ufeff' + sourceHTML], {
        type: 'application/msword;charset=utf-8'
    });

    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = filename.endsWith('.doc') ? filename : (filename + '.doc');
    document.body.appendChild(a);
    a.click();
    setTimeout(function() {
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }, 200);
}

/* 4. Helper phân tách Chương và Bài học từ tiêu đề chủ đề */
function parseTopicChapterAndLesson(topicStr, defaultIdx) {
    var raw = (topicStr || '').trim();
    var chNum = defaultIdx || 1;
    var chTitle = '';
    var lessonTitle = '';

    var baiIdx = raw.search(/\s*-\s*Bài\s+/i);
    if (baiIdx !== -1) {
        var chPart = raw.substring(0, baiIdx).trim();
        lessonTitle = raw.substring(baiIdx).replace(/^\s*-\s*/, '').trim();
        var chMatch = chPart.match(/^(?:Chương|Chủ đề)\s*([0-9IVXLCDM]+)[\s:.\-]+(.*)$/i);
        if (chMatch) {
            chNum = chMatch[1];
            chTitle = chMatch[2].trim();
        } else {
            chTitle = chPart;
        }
        return { chapterNum: chNum, chapterTitle: chTitle, lessonTitle: lessonTitle };
    }

    var parenBaiIdx = raw.search(/\s*\(\s*Bài\s+/i);
    if (parenBaiIdx !== -1) {
        var chPart = raw.substring(0, parenBaiIdx).trim();
        lessonTitle = raw.substring(parenBaiIdx).replace(/^\s*\(\s*/, '').replace(/\)\s*$/, '').trim();
        var chMatch = chPart.match(/^(?:Chương|Chủ đề)\s*([0-9IVXLCDM]+)[\s:.\-]+(.*)$/i);
        if (chMatch) {
            chNum = chMatch[1];
            chTitle = chMatch[2].trim();
        } else {
            chTitle = chPart;
        }
        return { chapterNum: chNum, chapterTitle: chTitle, lessonTitle: lessonTitle };
    }

    var chMatch2 = raw.match(/^(?:Chương|Chủ đề)\s*([0-9IVXLCDM]+)[\s:.\-]+(.*)$/i);
    if (chMatch2) {
        chNum = chMatch2[1];
        chTitle = chMatch2[2].trim();
        lessonTitle = chTitle;
        return { chapterNum: chNum, chapterTitle: chTitle, lessonTitle: lessonTitle };
    }

    return { chapterNum: chNum, chapterTitle: raw, lessonTitle: raw };
}

/* Helper ánh xạ thứ tự câu hỏi và mã lệnh hỏi cho Ma trận & Bản đặc tả */
function buildStandardMatrixAndSpecMapping(rows) {
    var curP1K = 1;
    var p1KMap = rows.map(function(r) {
        var list = [];
        for (var i = 0; i < (r.p1_k || 0); i++) list.push("C" + (curP1K++));
        return list;
    });
    var curP1H = curP1K;
    var p1HMap = rows.map(function(r) {
        var list = [];
        for (var i = 0; i < (r.p1_h || 0); i++) list.push("C" + (curP1H++));
        return list;
    });
    var curP1V = curP1H;
    var p1VMap = rows.map(function(r) {
        var list = [];
        for (var i = 0; i < (r.p1_v || 0); i++) list.push("C" + (curP1V++));
        return list;
    });

    var p2ItemIdx = 0;
    var p2Map = rows.map(function(r) {
        var kList = [], hList = [], vList = [];
        var getCode = function(idx) {
            var qNum = Math.floor(idx / 4) + 1;
            var letter = String.fromCharCode(97 + (idx % 4));
            return "C" + qNum + letter;
        };
        for (var i = 0; i < (r.p2_k || 0); i++) kList.push(getCode(p2ItemIdx++));
        for (var i = 0; i < (r.p2_h || 0); i++) hList.push(getCode(p2ItemIdx++));
        for (var i = 0; i < (r.p2_v || 0); i++) vList.push(getCode(p2ItemIdx++));
        return { k: kList, h: hList, v: vList };
    });

    var curP3K = 1;
    var p3KMap = rows.map(function(r) {
        var list = [];
        for (var i = 0; i < (r.p3_k || 0); i++) list.push("C" + (curP3K++));
        return list;
    });
    var curP3H = curP3K;
    var p3HMap = rows.map(function(r) {
        var list = [];
        for (var i = 0; i < (r.p3_h || 0); i++) list.push("C" + (curP3H++));
        return list;
    });
    var curP3V = curP3H;
    var p3VMap = rows.map(function(r) {
        var list = [];
        for (var i = 0; i < (r.p3_v || 0); i++) list.push("C" + (curP3V++));
        return list;
    });

    var curP4K = 1;
    var p4KMap = rows.map(function(r) {
        var list = [];
        for (var i = 0; i < (r.p4_k || 0); i++) list.push("C" + (curP4K++));
        return list;
    });
    var curP4H = curP4K;
    var p4HMap = rows.map(function(r) {
        var list = [];
        for (var i = 0; i < (r.p4_h || 0); i++) list.push("C" + (curP4H++));
        return list;
    });
    var curP4V = curP4H;
    var p4VMap = rows.map(function(r) {
        var list = [];
        for (var i = 0; i < (r.p4_v || 0); i++) list.push("C" + (curP4V++));
        return list;
    });

    return rows.map(function(r, idx) {
        var fmt = function(list) { return list && list.length > 0 ? list.join(", ") : "-"; };
        var p1Score = ((r.p1_k||0) + (r.p1_h||0) + (r.p1_v||0)) * 0.25;
        var p2Score = ((r.p2_k||0) + (r.p2_h||0) + (r.p2_v||0)) * 0.25;
        var p3Score = ((r.p3_k||0) + (r.p3_h||0) + (r.p3_v||0)) * 0.50;
        var p4Score = ((r.p4_k||0) + (r.p4_h||0) + (r.p4_v||0)) * 1.00;
        var totalScore = p1Score + p2Score + p3Score + p4Score;

        var parsedTopic = parseTopicChapterAndLesson(r.topic, idx + 1);

        return {
            chapterNum: parsedTopic.chapterNum,
            chapterTitle: parsedTopic.chapterTitle,
            lessonTitle: parsedTopic.lessonTitle,
            p1: { k: fmt(p1KMap[idx]), h: fmt(p1HMap[idx]), v: fmt(p1VMap[idx]) },
            p2: { k: fmt(p2Map[idx].k), h: fmt(p2Map[idx].h), v: fmt(p2Map[idx].v) },
            p3: { k: fmt(p3KMap[idx]), h: fmt(p3HMap[idx]), v: fmt(p3VMap[idx]) },
            p4: { k: fmt(p4KMap[idx]), h: fmt(p4HMap[idx]), v: fmt(p4VMap[idx]) },
            totalScore: totalScore.toFixed(2).replace(".", ",") + " đ"
        };
    });
}

function getTermFullName(term) {
    if (term === "GK1") return "GIỮA HỌC KÌ I";
    if (term === "CK1") return "CUỐI HỌC KÌ I";
    if (term === "GK2") return "GIỮA HỌC KÌ II";
    if (term === "CK2") return "CUỐI HỌC KÌ II";
    return term || "KIỂM TRA ĐỊNH KỲ";
}

/* Helper sinh HTML Khung Ma Trận chuẩn 16 cột + 3 dòng tổng kết */
function generateStandardMatrixHtml(grade, term, variantCode, timeVal, matrixRows, mapping) {
    var termName = getTermFullName(term);
    var chapterSpans = [];
    var currentCh = null;
    var currentCount = 0;
    var startIndex = 0;

    mapping.forEach(function(m, idx) {
        if (m.chapterTitle !== currentCh) {
            if (currentCh !== null) {
                chapterSpans.push({ startIndex: startIndex, count: currentCount, chapterNum: mapping[startIndex].chapterNum, chapterTitle: currentCh });
            }
            currentCh = m.chapterTitle;
            currentCount = 1;
            startIndex = idx;
        } else {
            currentCount++;
        }
    });
    if (currentCh !== null) {
        chapterSpans.push({ startIndex: startIndex, count: currentCount, chapterNum: mapping[startIndex].chapterNum, chapterTitle: currentCh });
    }

    var html = "<table class='header-table' style='width: 100%; border: none; margin-bottom: 8px;'>" +
        "<tr>" +
        "<td style='width: 45%; text-align: center; border: none; vertical-align: top;'>" +
        "<p class='bold' style='margin: 2px 0; font-size: 11pt;'>SỞ GIÁO DỤC VÀ ĐÀO TẠO LÂM ĐỒNG</p>" +
        "<p class='bold' style='margin: 2px 0; font-size: 11pt;'>TRƯỜNG THPT BÙI THỊ XUÂN – MŨI NÉ</p>" +
        "<p class='italic' style='margin: 2px 0; font-size: 10pt;'>Tổ: Hóa - Sinh - KTNN</p>" +
        "</td>" +
        "<td style='width: 55%; text-align: center; border: none; vertical-align: top;'>" +
        "<p class='bold' style='margin: 2px 0; font-size: 11pt;'>KHUNG MA TRẬN ĐỀ KIỂM TRA " + termName.toUpperCase() + "</p>" +
        "<p class='bold' style='margin: 2px 0; font-size: 11pt;'>MÔN: HÓA HỌC " + grade + " – MÃ ĐỀ: " + variantCode + "</p>" +
        "<p class='italic' style='margin: 2px 0; font-size: 10.5pt;'>Thời gian: " + timeVal + " phút – Tổng điểm: 10,0 điểm</p>" +
        "</td>" +
        "</tr>" +
        "</table>";

    html += "<table class='data-table matrix-table' style='width: 100%; border-collapse: collapse; margin-top: 6px; margin-bottom: 12px;'>" +
        "<thead>" +
        "<tr style='background-color: #f1f5f9;'>" +
        "<th rowspan='3' style='width: 3%; text-align: center; border: 1px solid #000; font-size: 9.5pt; font-weight: bold;'>TT</th>" +
        "<th rowspan='3' style='width: 13%; text-align: center; border: 1px solid #000; font-size: 9.5pt; font-weight: bold;'>Chương</th>" +
        "<th rowspan='3' style='width: 18%; text-align: center; border: 1px solid #000; font-size: 9.5pt; font-weight: bold;'>Nội dung / Đơn vị kiến thức</th>" +
        "<th colspan='12' style='width: 58%; text-align: center; border: 1px solid #000; font-size: 9.5pt; font-weight: bold;'>Mức độ đánh giá</th>" +
        "<th rowspan='3' style='width: 8%; text-align: center; border: 1px solid #000; font-size: 9.5pt; font-weight: bold;'>Tổng điểm</th>" +
        "</tr>" +
        "<tr style='background-color: #f1f5f9;'>" +
        "<th colspan='3' style='text-align: center; border: 1px solid #000; font-size: 9pt; font-weight: bold;'>TNKQ nhiều lựa chọn (I)</th>" +
        "<th colspan='3' style='text-align: center; border: 1px solid #000; font-size: 9pt; font-weight: bold;'>TNKQ đúng – sai (II)</th>" +
        "<th colspan='3' style='text-align: center; border: 1px solid #000; font-size: 9pt; font-weight: bold;'>TNKQ trả lời ngắn (III)</th>" +
        "<th colspan='3' style='text-align: center; border: 1px solid #000; font-size: 9pt; font-weight: bold;'>Tự luận (IV)</th>" +
        "</tr>" +
        "<tr style='background-color: #f1f5f9;'>" +
        "<th style='text-align: center; border: 1px solid #000; font-size: 8.5pt; font-weight: bold;'>Biết</th>" +
        "<th style='text-align: center; border: 1px solid #000; font-size: 8.5pt; font-weight: bold;'>Hiểu</th>" +
        "<th style='text-align: center; border: 1px solid #000; font-size: 8.5pt; font-weight: bold;'>VD</th>" +
        "<th style='text-align: center; border: 1px solid #000; font-size: 8.5pt; font-weight: bold;'>Biết</th>" +
        "<th style='text-align: center; border: 1px solid #000; font-size: 8.5pt; font-weight: bold;'>Hiểu</th>" +
        "<th style='text-align: center; border: 1px solid #000; font-size: 8.5pt; font-weight: bold;'>VD</th>" +
        "<th style='text-align: center; border: 1px solid #000; font-size: 8.5pt; font-weight: bold;'>Biết</th>" +
        "<th style='text-align: center; border: 1px solid #000; font-size: 8.5pt; font-weight: bold;'>Hiểu</th>" +
        "<th style='text-align: center; border: 1px solid #000; font-size: 8.5pt; font-weight: bold;'>VD</th>" +
        "<th style='text-align: center; border: 1px solid #000; font-size: 8.5pt; font-weight: bold;'>Biết</th>" +
        "<th style='text-align: center; border: 1px solid #000; font-size: 8.5pt; font-weight: bold;'>Hiểu</th>" +
        "<th style='text-align: center; border: 1px solid #000; font-size: 8.5pt; font-weight: bold;'>VD</th>" +
        "</tr>" +
        "</thead>" +
        "<tbody>";

    var totP1K = 0, totP1H = 0, totP1V = 0;
    var totP2K = 0, totP2H = 0, totP2V = 0;
    var totP3K = 0, totP3H = 0, totP3V = 0;
    var totP4K = 0, totP4H = 0, totP4V = 0;

    matrixRows.forEach(function(r, idx) {
        totP1K += (r.p1_k || 0); totP1H += (r.p1_h || 0); totP1V += (r.p1_v || 0);
        totP2K += (r.p2_k || 0); totP2H += (r.p2_h || 0); totP2V += (r.p2_v || 0);
        totP3K += (r.p3_k || 0); totP3H += (r.p3_h || 0); totP3V += (r.p3_v || 0);
        totP4K += (r.p4_k || 0); totP4H += (r.p4_h || 0); totP4V += (r.p4_v || 0);

        var m = mapping[idx];
        var chSpan = chapterSpans.find(function(cs) { return cs.startIndex === idx; });

        html += "<tr>";
        if (chSpan) {
            html += "<td rowspan='" + chSpan.count + "' style='border: 1px solid #000; text-align: center; font-weight: bold; font-size: 9.5pt;'>" + chSpan.chapterNum + "</td>";
            html += "<td rowspan='" + chSpan.count + "' style='border: 1px solid #000; font-weight: bold; font-size: 9.5pt; padding: 4px;'>" + cleanLatexForWord(chSpan.chapterTitle) + "</td>";
        }
        html += "<td style='border: 1px solid #000; font-size: 9.5pt; padding: 4px;'>" + cleanLatexForWord(m.lessonTitle) + "</td>";
        html += "<td style='border: 1px solid #000; text-align: center; font-size: 9pt;'>" + m.p1.k + "</td>";
        html += "<td style='border: 1px solid #000; text-align: center; font-size: 9pt;'>" + m.p1.h + "</td>";
        html += "<td style='border: 1px solid #000; text-align: center; font-size: 9pt;'>" + m.p1.v + "</td>";
        html += "<td style='border: 1px solid #000; text-align: center; font-size: 9pt;'>" + m.p2.k + "</td>";
        html += "<td style='border: 1px solid #000; text-align: center; font-size: 9pt;'>" + m.p2.h + "</td>";
        html += "<td style='border: 1px solid #000; text-align: center; font-size: 9pt;'>" + m.p2.v + "</td>";
        html += "<td style='border: 1px solid #000; text-align: center; font-size: 9pt;'>" + m.p3.k + "</td>";
        html += "<td style='border: 1px solid #000; text-align: center; font-size: 9pt;'>" + m.p3.h + "</td>";
        html += "<td style='border: 1px solid #000; text-align: center; font-size: 9pt;'>" + m.p3.v + "</td>";
        html += "<td style='border: 1px solid #000; text-align: center; font-size: 9pt;'>" + m.p4.k + "</td>";
        html += "<td style='border: 1px solid #000; text-align: center; font-size: 9pt;'>" + m.p4.h + "</td>";
        html += "<td style='border: 1px solid #000; text-align: center; font-size: 9pt;'>" + m.p4.v + "</td>";
        html += "<td style='border: 1px solid #000; text-align: center; font-weight: bold; font-size: 9.5pt;'>" + m.totalScore + "</td>";
        html += "</tr>";
    });

    html += "<tr style='background-color: #f8fafc; font-weight: bold;'>" +
        "<td colspan='3' style='border: 1px solid #000; text-align: center; font-size: 9.5pt;'>Tổng số câu (Lệnh hỏi) (30 lệnh hỏi / 21 câu)</td>" +
        "<td style='border: 1px solid #000; text-align: center; font-size: 9.5pt;'>" + totP1K + "</td>" +
        "<td style='border: 1px solid #000; text-align: center; font-size: 9.5pt;'>" + totP1H + "</td>" +
        "<td style='border: 1px solid #000; text-align: center; font-size: 9.5pt;'>" + totP1V + "</td>" +
        "<td style='border: 1px solid #000; text-align: center; font-size: 9.5pt;'>" + totP2K + "</td>" +
        "<td style='border: 1px solid #000; text-align: center; font-size: 9.5pt;'>" + totP2H + "</td>" +
        "<td style='border: 1px solid #000; text-align: center; font-size: 9.5pt;'>" + totP2V + "</td>" +
        "<td style='border: 1px solid #000; text-align: center; font-size: 9.5pt;'>" + totP3K + "</td>" +
        "<td style='border: 1px solid #000; text-align: center; font-size: 9.5pt;'>" + totP3H + "</td>" +
        "<td style='border: 1px solid #000; text-align: center; font-size: 9.5pt;'>" + totP3V + "</td>" +
        "<td style='border: 1px solid #000; text-align: center; font-size: 9.5pt;'>" + totP4K + "</td>" +
        "<td style='border: 1px solid #000; text-align: center; font-size: 9.5pt;'>" + totP4H + "</td>" +
        "<td style='border: 1px solid #000; text-align: center; font-size: 9.5pt;'>" + totP4V + "</td>" +
        "<td style='border: 1px solid #000; text-align: center; font-size: 9.5pt;'>30</td>" +
        "</tr>";

    var p1Tot = totP1K + totP1H + totP1V;
    var p2Tot = totP2K + totP2H + totP2V;
    var p3Tot = totP3K + totP3H + totP3V;
    var p4Tot = totP4K + totP4H + totP4V;
    var p1Score = (p1Tot * 0.25).toFixed(2).replace(".", ",");
    var p2Score = (p2Tot * 0.25).toFixed(2).replace(".", ",");
    var p3Score = (p3Tot * 0.50).toFixed(2).replace(".", ",");
    var p4Score = (p4Tot * 1.00).toFixed(2).replace(".", ",");

    html += "<tr style='background-color: #f8fafc; font-weight: bold;'>" +
        "<td colspan='3' style='border: 1px solid #000; text-align: center; font-size: 9.5pt;'>Tổng điểm từng phần</td>" +
        "<td colspan='3' style='border: 1px solid #000; text-align: center; font-size: 9.5pt;'>" + p1Score + " đ (" + p1Tot + " câu × 0,25đ)</td>" +
        "<td colspan='3' style='border: 1px solid #000; text-align: center; font-size: 9.5pt;'>" + p2Score + " đ (2 câu = " + p2Tot + " ý × 0,25đ)</td>" +
        "<td colspan='3' style='border: 1px solid #000; text-align: center; font-size: 9.5pt;'>" + p3Score + " đ (" + p3Tot + " câu × 0,50đ)</td>" +
        "<td colspan='3' style='border: 1px solid #000; text-align: center; font-size: 9.5pt;'>" + p4Score + " đ (" + p4Tot + " câu × 1,00đ)</td>" +
        "<td style='border: 1px solid #000; text-align: center; font-size: 9.5pt;'>10,0 đ</td>" +
        "</tr>";

    html += "<tr style='background-color: #f8fafc; font-weight: bold;'>" +
        "<td colspan='3' style='border: 1px solid #000; text-align: center; font-size: 9.5pt;'>Tổng tỉ lệ điểm theo mức độ</td>" +
        "<td colspan='4' style='border: 1px solid #000; text-align: center; font-size: 9.5pt;'>Biết: 4,00 đ (40%)</td>" +
        "<td colspan='4' style='border: 1px solid #000; text-align: center; font-size: 9.5pt;'>Hiểu: 3,00 đ (30%)</td>" +
        "<td colspan='4' style='border: 1px solid #000; text-align: center; font-size: 9.5pt;'>Vận dụng: 3,00 đ (30%)</td>" +
        "<td style='border: 1px solid #000; text-align: center; font-size: 9.5pt;'>100%</td>" +
        "</tr>";

    html += "</tbody></table>";
    return html;
}

/* Helper sinh HTML Bảng Đặc Tả chuẩn 9 cột */
function generateStandardSpecHtml(grade, term, variantCode, timeVal, matrixRows, mapping) {
    var termName = getTermFullName(term);

    var specRows = [];
    matrixRows.forEach(function(r, rowIdx) {
        var m = mapping[rowIdx];
        var spec = (typeof resolveRowSpecification === 'function') ? resolveRowSpecification(r) : {};
        var kDesc = r.knowDesc || spec.knowDesc || "";
        var uDesc = r.underDesc || spec.underDesc || "";
        var aDesc = r.applyDesc || spec.applyDesc || "";

        var knowBullets = splitSpecBullets(kDesc);
        var underBullets = splitSpecBullets(uDesc);
        var applyBullets = splitSpecBullets(aDesc);

        var hasK = (r.p1_k || 0) + (r.p2_k || 0) + (r.p3_k || 0) + (r.p4_k || 0) > 0 || knowBullets.length > 0;
        var hasH = (r.p1_h || 0) + (r.p2_h || 0) + (r.p3_h || 0) + (r.p4_h || 0) > 0 || underBullets.length > 0;
        var hasVQuestions = (r.p1_v || 0) + (r.p2_v || 0) + (r.p3_v || 0) + (r.p4_v || 0) > 0;
        var hasV = hasVQuestions || (rowIdx === 1 && applyBullets.length > 0);

        var levels = [];
        if (hasK) {
            levels.push({ level: "Biết", bullets: knowBullets, p1: m.p1.k, p2: m.p2.k, p3: m.p3.k, p4: m.p4.k });
        }
        if (hasH) {
            levels.push({ level: "Hiểu", bullets: underBullets, p1: m.p1.h, p2: m.p2.h, p3: m.p3.h, p4: m.p4.h });
        }
        if (hasV) {
            levels.push({ level: "Vận dụng", bullets: applyBullets, p1: m.p1.v, p2: m.p2.v, p3: m.p3.v, p4: m.p4.v });
        }

        specRows.push({
            chapterNum: m.chapterNum,
            chapterTitle: m.chapterTitle,
            lessonTitle: m.lessonTitle,
            levels: levels
        });
    });

    var chapterSpans = [];
    var curCh = null;
    var curChTotalLevels = 0;
    var curChStartLevelIndex = 0;
    var curChChapterNum = null;

    var totalLevelCount = 0;
    specRows.forEach(function(sr) {
        if (sr.chapterTitle !== curCh) {
            if (curCh !== null) {
                chapterSpans.push({ startIndex: curChStartLevelIndex, count: curChTotalLevels, chapterNum: curChChapterNum, chapterTitle: curCh });
            }
            curCh = sr.chapterTitle;
            curChChapterNum = sr.chapterNum;
            curChTotalLevels = sr.levels.length;
            curChStartLevelIndex = totalLevelCount;
        } else {
            curChTotalLevels += sr.levels.length;
        }
        totalLevelCount += sr.levels.length;
    });
    if (curCh !== null) {
        chapterSpans.push({ startIndex: curChStartLevelIndex, count: curChTotalLevels, chapterNum: curChChapterNum, chapterTitle: curCh });
    }

    var html = "<table class='header-table' style='width: 100%; border: none; margin-bottom: 8px;'>" +
        "<tr>" +
        "<td style='width: 45%; text-align: center; border: none; vertical-align: top;'>" +
        "<p class='bold' style='margin: 2px 0; font-size: 11pt;'>SỞ GIÁO DỤC VÀ ĐÀO TẠO LÂM ĐỒNG</p>" +
        "<p class='bold' style='margin: 2px 0; font-size: 11pt;'>TRƯỜNG THPT BÙI THỊ XUÂN – MŨI NÉ</p>" +
        "<p class='italic' style='margin: 2px 0; font-size: 10pt;'>Tổ: Hóa - Sinh - KTNN</p>" +
        "</td>" +
        "<td style='width: 55%; text-align: center; border: none; vertical-align: top;'>" +
        "<p class='bold' style='margin: 2px 0; font-size: 11pt;'>BẢNG ĐẶC TẢ ĐỀ KIỂM TRA " + termName.toUpperCase() + "</p>" +
        "<p class='bold' style='margin: 2px 0; font-size: 11pt;'>MÔN: HÓA HỌC " + grade + " – MÃ ĐỀ: " + variantCode + "</p>" +
        "<p class='italic' style='margin: 2px 0; font-size: 10.5pt;'>Thời gian làm bài: " + timeVal + " phút</p>" +
        "</td>" +
        "</tr>" +
        "</table>";

    html += "<table class='data-table spec-table' style='width: 100%; border-collapse: collapse; margin-top: 6px; margin-bottom: 12px;'>" +
        "<thead>" +
        "<tr style='background-color: #f1f5f9;'>" +
        "<th rowspan='3' style='width: 4%; text-align: center; border: 1px solid #000; font-size: 9.5pt; font-weight: bold;'>TT<br/>(1)</th>" +
        "<th rowspan='3' style='width: 13%; text-align: center; border: 1px solid #000; font-size: 9.5pt; font-weight: bold;'>Chủ đề<br/>(2)</th>" +
        "<th rowspan='3' style='width: 17%; text-align: center; border: 1px solid #000; font-size: 9.5pt; font-weight: bold;'>Nội dung / đơn vị kiến thức<br/>(3)</th>" +
        "<th rowspan='3' style='width: 8%; text-align: center; border: 1px solid #000; font-size: 9.5pt; font-weight: bold;'>Cấp độ tư duy<br/>(4)</th>" +
        "<th rowspan='3' style='width: 34%; text-align: center; border: 1px solid #000; font-size: 9.5pt; font-weight: bold;'>Yêu cầu cần đạt<br/>(5)</th>" +
        "<th colspan='4' style='width: 24%; text-align: center; border: 1px solid #000; font-size: 9.5pt; font-weight: bold;'>Số lượng câu hỏi ở các mức độ</th>" +
        "</tr>" +
        "<tr style='background-color: #f1f5f9;'>" +
        "<th colspan='3' style='text-align: center; border: 1px solid #000; font-size: 9pt; font-weight: bold;'>Trắc nghiệm</th>" +
        "<th rowspan='2' style='width: 6%; text-align: center; border: 1px solid #000; font-size: 9pt; font-weight: bold;'>Tự luận<br/>(Phần IV)<br/>(9)</th>" +
        "</tr>" +
        "<tr style='background-color: #f1f5f9;'>" +
        "<th style='width: 6%; text-align: center; border: 1px solid #000; font-size: 8.5pt; font-weight: bold;'>Nhiều lựa chọn<br/>(Phần I)<br/>(6)</th>" +
        "<th style='width: 6%; text-align: center; border: 1px solid #000; font-size: 8.5pt; font-weight: bold;'>Đúng – Sai<br/>(Phần II)<br/>(7)</th>" +
        "<th style='width: 6%; text-align: center; border: 1px solid #000; font-size: 8.5pt; font-weight: bold;'>Trả lời ngắn<br/>(Phần III)<br/>(8)</th>" +
        "</tr>" +
        "</thead>" +
        "<tbody>";

    var currentLevelGlobalIdx = 0;
    specRows.forEach(function(sr) {
        sr.levels.forEach(function(lvl, lvlIdx) {
            var isFirstOfLesson = (lvlIdx === 0);
            var chSpan = chapterSpans.find(function(cs) { return cs.startIndex === currentLevelGlobalIdx; });

            html += "<tr>";
            if (chSpan) {
                html += "<td rowspan='" + chSpan.count + "' style='border: 1px solid #000; text-align: center; font-weight: bold; font-size: 9.5pt;'>" + chSpan.chapterNum + "</td>";
                html += "<td rowspan='" + chSpan.count + "' style='border: 1px solid #000; font-weight: bold; font-size: 9.5pt; padding: 4px;'>" + cleanLatexForWord(chSpan.chapterTitle) + "</td>";
            }
            if (isFirstOfLesson) {
                html += "<td rowspan='" + sr.levels.length + "' style='border: 1px solid #000; font-weight: bold; font-size: 9.5pt; padding: 4px;'>" + cleanLatexForWord(sr.lessonTitle) + "</td>";
            }

            html += "<td style='border: 1px solid #000; text-align: center; font-size: 9.5pt;'>" + lvl.level + "</td>";

            var reqHtml = lvl.bullets.map(function(b) {
                return "<p style='margin: 2px 0; font-size: 9pt; line-height: 1.3;'>- " + cleanLatexForWord(b.replace(/^[-*•\s]+/, '')) + "</p>";
            }).join("");
            html += "<td style='border: 1px solid #000; padding: 4px; font-size: 9pt;'>" + (reqHtml || "-") + "</td>";

            html += "<td style='border: 1px solid #000; text-align: center; font-size: 9pt;'>" + lvl.p1 + "</td>";
            html += "<td style='border: 1px solid #000; text-align: center; font-size: 9pt;'>" + lvl.p2 + "</td>";
            html += "<td style='border: 1px solid #000; text-align: center; font-size: 9pt;'>" + lvl.p3 + "</td>";
            html += "<td style='border: 1px solid #000; text-align: center; font-size: 9pt;'>" + lvl.p4 + "</td>";
            html += "</tr>";

            currentLevelGlobalIdx++;
        });
    });

    html += "</tbody></table>";
    return html;
}

/* Helper sinh HTML Đề kiểm tra chuẩn CV 7991 */
function generateStandardExamHtml(grade, term, variantCode, timeVal, selected) {
    var termName = getTermFullName(term);
    var part1 = selected.part1 || [];
    var part2 = selected.part2 || [];
    var part3 = selected.part3 || [];
    var part4 = selected.part4 || [];

    var html = "<table class='header-table' style='width: 100%; border: none; margin-bottom: 6px;'>" +
        "<tr>" +
        "<td style='width: 45%; text-align: center; border: none; vertical-align: top;'>" +
        "<p class='bold' style='margin: 2px 0; font-size: 11pt;'>SỞ GIÁO DỤC VÀ ĐÀO TẠO LÂM ĐỒNG</p>" +
        "<p class='bold' style='margin: 2px 0; font-size: 11pt;'>TRƯỜNG THPT BÙI THỊ XUÂN – MŨI NÉ</p>" +
        "<p class='italic' style='margin: 2px 0; font-size: 10pt;'>Tổ: Hóa - Sinh - KTNN</p>" +
        "<p class='italic' style='margin: 2px 0; font-size: 9.5pt;'>(Đề kiểm tra có 03 trang)</p>" +
        "</td>" +
        "<td style='width: 55%; text-align: center; border: none; vertical-align: top;'>" +
        "<p class='bold' style='margin: 2px 0; font-size: 11pt;'>ĐỀ KIỂM TRA " + termName.toUpperCase() + "</p>" +
        "<p class='bold' style='margin: 2px 0; font-size: 11pt;'>MÔN: HÓA HỌC " + grade + " – NĂM HỌC 2025 - 2026</p>" +
        "<p class='italic' style='margin: 2px 0; font-size: 10pt;'>Thời gian làm bài: " + timeVal + " phút (không kể thời gian phát đề)</p>" +
        "</td>" +
        "</tr>" +
        "</table>";

    html += "<table style='width: 100%; border-collapse: collapse; margin: 4px 0 10px 0;'>" +
        "<tr>" +
        "<td style='border: none; width: 75%; font-size: 11pt; padding: 2px;'><strong>Họ và tên thí sinh:</strong> ........................................................................ <strong>Lớp:</strong> .............</td>" +
        "<td style='border: 1px solid #000; width: 25%; text-align: center; font-size: 11pt; padding: 4px; font-weight: bold; background-color: #f8fafc;'>MÃ ĐỀ THI: " + variantCode + "</td>" +
        "</tr>" +
        "</table>" +
        "<hr style='border: 1px solid #000; margin-bottom: 12px;'/>";

    // Part I
    html += "<p class='bold' style='font-size: 11pt; margin-top: 10px;'>PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn (3,0 điểm).</p>" +
        "<p class='italic' style='margin: 2px 0 8px 0; font-size: 10pt;'>Thí sinh trả lời từ câu 1 đến câu " + part1.length + ". Mỗi câu hỏi thí sinh chỉ chọn một phương án.</p>";

    part1.forEach(function(q, idx) {
        html += "<p style='margin: 5px 0 3px 0;'><strong>Câu " + (idx + 1) + ".</strong> " + cleanLatexForWord(q.content) + "</p>" +
            "<div style='margin-left: 18px; margin-bottom: 6px;'>" +
            formatWordOption("A", q.options[0]) +
            formatWordOption("B", q.options[1]) +
            formatWordOption("C", q.options[2]) +
            formatWordOption("D", q.options[3]) +
            "</div>";
    });

    // Part II
    html += "<p class='bold' style='font-size: 11pt; margin-top: 14px;'>PHẦN II. Câu trắc nghiệm đúng sai (2,0 điểm).</p>" +
        "<p class='italic' style='margin: 2px 0 8px 0; font-size: 10pt;'>Thí sinh trả lời từ câu 1 đến câu " + part2.length + ". Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.</p>";

    part2.forEach(function(q, idx) {
        html += "<p style='margin: 5px 0 3px 0;'><strong>Câu " + (idx + 1) + ".</strong> " + cleanLatexForWord(q.content) + "</p>" +
            "<div style='margin-left: 16px; margin-bottom: 8px;'>";
        q.statements.forEach(function(s, sIdx) {
            var letter = String.fromCharCode(97 + sIdx);
            html += "<p style='margin: 2px 0;'><strong>" + letter + ")</strong> " + cleanLatexForWord(s.text) + "</p>";
        });
        html += "</div>";
    });

    // Part III
    html += "<p class='bold' style='font-size: 11pt; margin-top: 14px;'>PHẦN III. Câu trắc nghiệm trả lời ngắn (2,0 điểm).</p>" +
        "<p class='italic' style='margin: 2px 0 8px 0; font-size: 10pt;'>Thí sinh trả lời từ câu 1 đến câu " + part3.length + ".</p>";

    part3.forEach(function(q, idx) {
        html += "<p style='margin: 5px 0 3px 0;'><strong>Câu " + (idx + 1) + ".</strong> " + cleanLatexForWord(q.content) + "</p>" +
            "<p style='margin-left: 16px; color: #333; margin-bottom: 6px;'><em>Đáp án: ........................</em></p>";
    });

    // Part IV
    if (part4.length > 0) {
        html += "<p class='bold' style='font-size: 11pt; margin-top: 14px;'>PHẦN IV. Tự luận (3,0 điểm).</p>" +
            "<p class='italic' style='margin: 2px 0 8px 0; font-size: 10pt;'>Thí sinh trình bày lời giải chi tiết từ câu 1 đến câu " + part4.length + ".</p>";

        part4.forEach(function(q, idx) {
            html += "<p style='margin: 6px 0 4px 0;'><strong>Câu " + (idx + 1) + ".</strong> " + cleanLatexForWord(q.content).replace(/\n/g, "<br/>") + "</p>";
        });
    }

    html += "<p class='bold text-center' style='margin-top: 20px; font-size: 11pt;'>---------------- HẾT ----------------</p>";
    return html;
}

/* Helper sinh HTML Đáp án & Hướng dẫn chấm chi tiết */
function generateStandardAnswerHtml(grade, term, variantCode, timeVal, selected) {
    var termName = getTermFullName(term);
    var part1 = selected.part1 || [];
    var part2 = selected.part2 || [];
    var part3 = selected.part3 || [];
    var part4 = selected.part4 || [];

    var html = "<table class='header-table' style='width: 100%; border: none; margin-bottom: 6px;'>" +
        "<tr>" +
        "<td style='width: 45%; text-align: center; border: none; vertical-align: top;'>" +
        "<p class='bold' style='margin: 2px 0; font-size: 11pt;'>SỞ GIÁO DỤC VÀ ĐÀO TẠO LÂM ĐỒNG</p>" +
        "<p class='bold' style='margin: 2px 0; font-size: 11pt;'>TRƯỜNG THPT BÙI THỊ XUÂN – MŨI NÉ</p>" +
        "<p class='italic' style='margin: 2px 0; font-size: 10pt;'>Tổ: Hóa - Sinh - KTNN</p>" +
        "</td>" +
        "<td style='width: 55%; text-align: center; border: none; vertical-align: top;'>" +
        "<p class='bold' style='margin: 2px 0; font-size: 11pt;'>ĐÁP ÁN VÀ HƯỚNG DẪN CHẤM ĐỀ KIỂM TRA " + termName.toUpperCase() + "</p>" +
        "<p class='bold' style='margin: 2px 0; font-size: 11pt;'>MÔN: HÓA HỌC " + grade + " – MÃ ĐỀ: " + variantCode + "</p>" +
        "<p class='italic' style='margin: 2px 0; font-size: 10pt;'>Giáo viên biên soạn: GV. NGUYỄN VỸ</p>" +
        "</td>" +
        "</tr>" +
        "</table>";

    // Section 1 Answer
    html += "<h4 style='margin-top: 10px; margin-bottom: 4px; font-size: 11pt;'>PHẦN I. CÂU TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (3,0 điểm)</h4>" +
        "<p class='italic' style='margin: 2px 0 6px 0; font-size: 9.5pt;'>(Mỗi câu trả lời đúng được 0,25 điểm)</p>" +
        "<table class='data-table' style='width: 100%; border-collapse: collapse; text-align: center; margin-bottom: 12px;'>" +
        "<tr style='background-color: #f1f5f9; font-weight: bold;'>" +
        "<td style='border: 1px solid #000; width: 8%;'>Câu</td>";
    part1.forEach(function(_, i) {
        html += "<td style='border: 1px solid #000; width: 7.5%; font-weight: bold;'>" + (i + 1) + "</td>";
    });
    html += "</tr><tr>" +
        "<td style='border: 1px solid #000; font-weight: bold; background-color: #f1f5f9;'>Đ/A</td>";
    part1.forEach(function(q) {
        html += "<td style='border: 1px solid #000; font-weight: bold; color: #047857; font-size: 11pt;'>" + String.fromCharCode(65 + q.correct) + "</td>";
    });
    html += "</tr></table>";

    // Section 2 Answer
    html += "<h4 style='margin-top: 12px; margin-bottom: 4px; font-size: 11pt;'>PHẦN II. CÂU TRẮC NGHIỆM ĐÚNG SAI (2,0 điểm)</h4>" +
        "<p class='italic' style='margin: 2px 0 6px 0; font-size: 9.5pt;'>(Điểm tối đa mỗi câu là 1,0 điểm. Đúng 1 ý: 0,10đ; đúng 2 ý: 0,25đ; đúng 3 ý: 0,50đ; đúng 4 ý: 1,00đ)</p>" +
        "<table class='data-table' style='width: 100%; border-collapse: collapse; margin-bottom: 12px;'>" +
        "<thead>" +
        "<tr style='background-color: #f1f5f9; font-weight: bold; text-align: center;'>" +
        "<th style='border: 1px solid #000; width: 10%;'>Câu hỏi</th>" +
        "<th style='border: 1px solid #000; width: 15%;'>Lệnh hỏi (Ý)</th>" +
        "<th style='border: 1px solid #000; width: 25%;'>Đáp án (Đúng / Sai)</th>" +
        "<th style='border: 1px solid #000; width: 50%;'>Nội dung tóm tắt & Yêu cầu cần đạt</th>" +
        "</tr>" +
        "</thead>" +
        "<tbody>";

    part2.forEach(function(q, qIdx) {
        q.statements.forEach(function(s, sIdx) {
            var isTrue = (s.correct === true || s.correct === "Đúng" || s.correct === "Đ");
            var letter = String.fromCharCode(97 + sIdx);
            html += "<tr>";
            if (sIdx === 0) {
                html += "<td rowspan='4' style='border: 1px solid #000; text-align: center; font-weight: bold; vertical-align: middle;'>Câu " + (qIdx + 1) + "</td>";
            }
            html += "<td style='border: 1px solid #000; text-align: center; font-weight: bold;'>Ý " + letter + ")</td>" +
                "<td style='border: 1px solid #000; text-align: center; font-weight: bold; color: " + (isTrue ? "#047857" : "#dc2626") + ";'>" + (isTrue ? "ĐÚNG" : "SAI") + "</td>" +
                "<td style='border: 1px solid #000; font-size: 9.5pt; padding: 4px;'>" + cleanLatexForWord(s.text) + "</td>" +
                "</tr>";
        });
    });
    html += "</tbody></table>";

    // Section 3 Answer
    html += "<h4 style='margin-top: 12px; margin-bottom: 4px; font-size: 11pt;'>PHẦN III. CÂU TRẮC NGHIỆM TRẢ LỜI NGẮN (2,0 điểm)</h4>" +
        "<p class='italic' style='margin: 2px 0 6px 0; font-size: 9.5pt;'>(Mỗi câu trả lời đúng được 0,50 điểm)</p>" +
        "<table class='data-table' style='width: 100%; border-collapse: collapse; text-align: center; margin-bottom: 8px;'>" +
        "<tr style='background-color: #f1f5f9; font-weight: bold;'>" +
        "<td style='border: 1px solid #000; width: 12%;'>Câu</td>";
    part3.forEach(function(_, i) {
        html += "<td style='border: 1px solid #000; width: 22%; font-weight: bold;'>Câu " + (i + 1) + "</td>";
    });
    html += "</tr><tr>" +
        "<td style='border: 1px solid #000; font-weight: bold; background-color: #f1f5f9;'>Đáp án</td>";
    part3.forEach(function(q) {
        html += "<td style='border: 1px solid #000; font-weight: bold; color: #7c3aed; font-size: 11pt;'>" + cleanLatexForWord(q.answer) + "</td>";
    });
    html += "</tr></table>";

    html += "<div style='margin-bottom: 12px; font-size: 9.5pt;'>";
    part3.forEach(function(q, i) {
        if (q.explain) {
            html += "<p style='margin: 2px 0;'><strong>• Hướng dẫn giải Câu " + (i + 1) + ":</strong> " + cleanLatexForWord(q.explain) + "</p>";
        }
    });
    html += "</div>";

    // Section 4 Answer
    if (part4.length > 0) {
        html += "<h4 style='margin-top: 14px; margin-bottom: 4px; font-size: 11pt;'>PHẦN IV. TỰ LUẬN (3,0 điểm)</h4>" +
            "<p class='italic' style='margin: 2px 0 6px 0; font-size: 9.5pt;'>(Hướng dẫn chấm chi tiết và biểu điểm từng bước)</p>" +
            "<table class='data-table answer-table' style='width: 100%; border-collapse: collapse; margin-bottom: 12px;'>" +
            "<thead>" +
            "<tr style='background-color: #f1f5f9; font-weight: bold; text-align: center;'>" +
            "<th style='border: 1px solid #000; width: 10%;'>Câu</th>" +
            "<th style='border: 1px solid #000; width: 75%;'>Nội dung đáp án & Các bước giải</th>" +
            "<th style='border: 1px solid #000; width: 15%;'>Điểm</th>" +
            "</tr>" +
            "</thead>" +
            "<tbody>";

        part4.forEach(function(q, i) {
            var solLines = (q.solution || '').split('\n').filter(Boolean);
            html += "<tr>" +
                "<td style='border: 1px solid #000; text-align: center; font-weight: bold; vertical-align: top; padding: 6px;'>Câu " + (i + 1) + "</td>" +
                "<td style='border: 1px solid #000; padding: 6px; font-size: 9.5pt;'>" +
                "<p style='margin: 0 0 4px 0; font-weight: bold; color: #1e3a8a;'>" + cleanLatexForWord(q.content).replace(/\n/g, '<br/>') + "</p>";
            solLines.forEach(function(line) {
                html += "<p style='margin: 3px 0; line-height: 1.35;'>" + cleanLatexForWord(line) + "</p>";
            });
            html += "</td>" +
                "<td style='border: 1px solid #000; text-align: center; font-weight: bold; vertical-align: top; padding: 6px;'>" + (i === 0 ? "0,50 đ" : (i === 1 ? "1,50 đ" : "1,00 đ")) + "</td>" +
                "</tr>";
        });
        html += "</tbody></table>";
    }

    return html;
}

/* 5. Xuất Ma Trận & Bản Đặc Tả ra file Word (.doc) */
function exportMatrixToWord() {
    var variantCode = (document.getElementById("exam-variant-select") ? document.getElementById("exam-variant-select").value : "101");
    var timeVal = (document.getElementById("exam-time-select") ? document.getElementById("exam-time-select").value : "45");

    var mapping = buildStandardMatrixAndSpecMapping(matrixRows);
    var p1Html = generateStandardMatrixHtml(currentGrade, currentTerm, variantCode, timeVal, matrixRows, mapping);
    var p2Html = generateStandardSpecHtml(currentGrade, currentTerm, variantCode, timeVal, matrixRows, mapping);

    var hardBreak = "<br clear='all' style='page-break-before: always; mso-break-type: section-break;' /><p style='page-break-before: always; mso-break-type: section-break; margin: 0; padding: 0; font-size: 1pt; line-height: 1pt;'>&nbsp;</p>";
    var body = "<h3 class='bold text-center' style='margin-top: 5px; margin-bottom: 8px;'>PHẦN 1: KHUNG MA TRẬN ĐỀ KIỂM TRA ĐỊNH KỲ (CV 7991)</h3>" +
        p1Html +
        hardBreak +
        "<h3 class='bold text-center' style='margin-top: 5px; margin-bottom: 8px;'>PHẦN 2: BẢNG ĐẶC TẢ ĐỀ KIỂM TRA ĐỊNH KỲ (CV 7991)</h3>" +
        p2Html;

    var fileName = "Ma_Tran_Va_Ban_Dac_Ta_Hoa_" + currentGrade + "_" + currentTerm + ".doc";
    downloadWordDocument(body, fileName);
}

/* 6. Xuất Đề Thi & Đáp Án ra file Word (.doc) */
function exportExamToWord() {
    var variantCode = (document.getElementById("exam-variant-select") ? document.getElementById("exam-variant-select").value : "101");
    var timeVal = (document.getElementById("exam-time-select") ? document.getElementById("exam-time-select").value : "45");

    var pool = (typeof QUESTION_BANK !== "undefined" && QUESTION_BANK[currentGrade]) ? QUESTION_BANK[currentGrade] : [];
    var selected = selectQuestionsByMatrix(pool);

    var p3Html = generateStandardExamHtml(currentGrade, currentTerm, variantCode, timeVal, selected);
    var p4Html = generateStandardAnswerHtml(currentGrade, currentTerm, variantCode, timeVal, selected);

    var hardBreak = "<br clear='all' style='page-break-before: always; mso-break-type: section-break;' /><p style='page-break-before: always; mso-break-type: section-break; margin: 0; padding: 0; font-size: 1pt; line-height: 1pt;'>&nbsp;</p>";
    var body = "<h3 class='bold text-center' style='margin-top: 5px; margin-bottom: 8px;'>PHẦN 1: ĐỀ KIỂM TRA CHÍNH THỨC (MÃ ĐỀ " + variantCode + ")</h3>" +
        p3Html +
        hardBreak +
        "<h3 class='bold text-center' style='margin-top: 5px; margin-bottom: 8px;'>PHẦN 2: ĐÁP ÁN VÀ HƯỚNG DẪN CHẤM CHI TIẾT (MÃ ĐỀ " + variantCode + ")</h3>" +
        p4Html;

    var fileName = "De_Thi_Va_Dap_An_Hoa_" + currentGrade + "_" + currentTerm + "_MaDe_" + variantCode + ".doc";
    downloadWordDocument(body, fileName);
}

/* 7. Xuất Trọn Bộ Hồ Sơ 4 Phần (Ma Trận + Đặc Tả + Đề Thi + Đáp Án) vào 1 file Word duy nhất */
function exportAllInOneWord() {
    var variantCode = (document.getElementById("exam-variant-select") ? document.getElementById("exam-variant-select").value : "101");
    var timeVal = (document.getElementById("exam-time-select") ? document.getElementById("exam-time-select").value : "45");

    var mapping = buildStandardMatrixAndSpecMapping(matrixRows);
    var p1Html = generateStandardMatrixHtml(currentGrade, currentTerm, variantCode, timeVal, matrixRows, mapping);
    var p2Html = generateStandardSpecHtml(currentGrade, currentTerm, variantCode, timeVal, matrixRows, mapping);

    var pool = (typeof QUESTION_BANK !== "undefined" && QUESTION_BANK[currentGrade]) ? QUESTION_BANK[currentGrade] : [];
    var selected = selectQuestionsByMatrix(pool);

    var p3Html = generateStandardExamHtml(currentGrade, currentTerm, variantCode, timeVal, selected);
    var p4Html = generateStandardAnswerHtml(currentGrade, currentTerm, variantCode, timeVal, selected);

    var hardBreak = "<br clear='all' style='page-break-before: always; mso-break-type: section-break;' /><p style='page-break-before: always; mso-break-type: section-break; margin: 0; padding: 0; font-size: 1pt; line-height: 1pt;'>&nbsp;</p>";

    var body = "<h3 class='bold text-center' style='margin-top: 5px; margin-bottom: 8px;'>PHẦN 1: KHUNG MA TRẬN ĐỀ KIỂM TRA ĐỊNH KỲ (CV 7991)</h3>" +
        p1Html +
        hardBreak +
        "<h3 class='bold text-center' style='margin-top: 5px; margin-bottom: 8px;'>PHẦN 2: BẢNG ĐẶC TẢ ĐỀ KIỂM TRA ĐỊNH KỲ (CV 7991)</h3>" +
        p2Html +
        hardBreak +
        "<h3 class='bold text-center' style='margin-top: 5px; margin-bottom: 8px;'>PHẦN 3: ĐỀ KIỂM TRA CHÍNH THỨC (MÃ ĐỀ " + variantCode + ")</h3>" +
        p3Html +
        hardBreak +
        "<h3 class='bold text-center' style='margin-top: 5px; margin-bottom: 8px;'>PHẦN 4: ĐÁP ÁN VÀ HƯỚNG DẪN CHẤM CHI TIẾT (MÃ ĐỀ " + variantCode + ")</h3>" +
        p4Html;

    var fileName = "Ho_So_Kiem_Tra_Tron_Bo_Hoa_" + currentGrade + "_" + currentTerm + "_MaDe_" + variantCode + ".doc";
    downloadWordDocument(body, fileName);
}

/* 7. Thuật toán chọn câu hỏi bám sát Ma trận & Đề chính thức của Tổ */
function selectQuestionsByMatrix(pool) {
    // Tính tổng số câu mục tiêu cho từng phần từ Ma trận hiện tại
    var targetP1 = 0, rawP2 = 0, targetP3 = 0, targetP4 = 0;
    matrixRows.forEach(function(r) {
        targetP1 += (r.p1_k + r.p1_h + r.p1_v);
        rawP2 += (r.p2_k + r.p2_h + r.p2_v);
        targetP3 += (r.p3_k + r.p3_h + r.p3_v);
        targetP4 += ((r.p4_k || 0) + (r.p4_h || 0) + (r.p4_v || 0));
    });

    // Quy đổi số lệnh hỏi Phần II (4 ý = 1 câu lớn Đúng/Sai theo chuẩn CV 7991)
    var targetP2 = (rawP2 >= 4) ? Math.round(rawP2 / 4) : (rawP2 > 0 ? 1 : 2);

    if (targetP1 === 0) targetP1 = 12;
    if (targetP2 === 0) targetP2 = 2;
    if (targetP3 === 0) targetP3 = 4;

    // 0. ƯU TIÊN SỐ 1: Nếu người dùng đã gán câu hỏi (tự động hoặc tự chọn) cho các hàng ma trận
    var anyAssigned = matrixRows.some(function(r) { return r.assignedQuestions && r.assignedQuestions.length > 0; });
    if (anyAssigned) {
        var assignedP1 = [];
        var assignedP2 = [];
        var assignedP3 = [];
        var assignedP4 = [];
        var assignedIds = new Set();

        matrixRows.forEach(function(r) {
            if (r.assignedQuestions) {
                r.assignedQuestions.forEach(function(q) {
                    if (!assignedIds.has(q.id)) {
                        assignedIds.add(q.id);
                        if (q.part === 1) assignedP1.push(q);
                        else if (q.part === 2) assignedP2.push(q);
                        else if (q.part === 3) assignedP3.push(q);
                        else if (q.part === 4) assignedP4.push(q);
                    }
                });
            }
        });

        // Bổ sung các câu còn thiếu nếu chưa đủ số lượng mục tiêu
        var remainingPool = pool.filter(function(q) { return !assignedIds.has(q.id); });
        while (assignedP1.length < targetP1) {
            var ex1 = remainingPool.find(function(q) { return q.part === 1 && !assignedIds.has(q.id); });
            if (!ex1) break;
            assignedIds.add(ex1.id);
            assignedP1.push(ex1);
        }
        while (assignedP2.length < targetP2) {
            var ex2 = remainingPool.find(function(q) { return q.part === 2 && !assignedIds.has(q.id); });
            if (!ex2) break;
            assignedIds.add(ex2.id);
            assignedP2.push(ex2);
        }
        while (assignedP3.length < targetP3) {
            var ex3 = remainingPool.find(function(q) { return q.part === 3 && !assignedIds.has(q.id); });
            if (!ex3) break;
            assignedIds.add(ex3.id);
            assignedP3.push(ex3);
        }
        while (assignedP4.length < targetP4) {
            var ex4 = remainingPool.find(function(q) { return q.part === 4 && !assignedIds.has(q.id); });
            if (!ex4) break;
            assignedIds.add(ex4.id);
            assignedP4.push(ex4);
        }

        return {
            part1: assignedP1.slice(0, targetP1),
            part2: assignedP2.slice(0, targetP2),
            part3: assignedP3.slice(0, targetP3),
            part4: assignedP4.slice(0, targetP4)
        };
    }

    // Ưu tiên lấy bộ câu hỏi thuộc đúng Kỳ kiểm tra chính thức của Tổ (GK1, CK1, GK2, CK2) nếu có
    var termPool = [];
    if (typeof OFFICIAL_BTX_EXAMS !== "undefined" && OFFICIAL_BTX_EXAMS[currentGrade] && OFFICIAL_BTX_EXAMS[currentGrade][currentTerm]) {
        // Kết hợp các câu hỏi người dùng vừa tải lên (user_*) ở đầu nếu có, sau đó là bộ đề chính thức của kỳ hiện tại
        var userUploaded = pool.filter(function(q) { return String(q.id).indexOf("user_") === 0; });
        termPool = userUploaded.concat(OFFICIAL_BTX_EXAMS[currentGrade][currentTerm]);
    } else {
        termPool = pool;
    }

    // Nếu đang sử dụng đúng cấu trúc đề chuẩn của Tổ và chưa nạp câu hỏi ngoài, trả về chuẩn xác bộ câu hỏi của Đề chính thức kỳ đó
    var hasUserQuestions = termPool.some(function(q) { return String(q.id).indexOf("user_") === 0; });
    if (!hasUserQuestions && typeof OFFICIAL_BTX_EXAMS !== "undefined" && OFFICIAL_BTX_EXAMS[currentGrade] && OFFICIAL_BTX_EXAMS[currentGrade][currentTerm]) {
        var officialExam = OFFICIAL_BTX_EXAMS[currentGrade][currentTerm];
        var offP1 = officialExam.filter(function(q) { return q.part === 1; });
        var offP2 = officialExam.filter(function(q) { return q.part === 2; });
        var offP3 = officialExam.filter(function(q) { return q.part === 3; });
        var offP4 = officialExam.filter(function(q) { return q.part === 4; });

        if (offP1.length >= targetP1 && offP2.length >= targetP2 && offP3.length >= targetP3 && offP4.length >= targetP4) {
            return {
                part1: offP1.slice(0, targetP1),
                part2: offP2.slice(0, targetP2),
                part3: offP3.slice(0, targetP3),
                part4: offP4.slice(0, targetP4)
            };
        }
    }

    var selectedP1 = [];
    var selectedP2 = [];
    var selectedP3 = [];
    var selectedP4 = [];
    var usedIds = new Set();

    var combinedPool = termPool.concat(pool);

    function pickFromPool(part, level, topic, count) {
        var candidates = combinedPool.filter(function(q) {
            if (q.part !== part || usedIds.has(q.id)) return false;
            if (q.term && q.term !== currentTerm) return false;
            return isQuestionInTopic(q, topic, currentGrade);
        });
        var matched = candidates.filter(function(q) {
            return level ? (q.level.toLowerCase() === level.toLowerCase()) : true;
        });

        if (matched.length < count) {
            var need = count - matched.length;
            var othersInTopic = candidates.filter(function(q) { return matched.indexOf(q) === -1; });
            matched = matched.concat(othersInTopic.slice(0, need));
        }

        var picked = matched.slice(0, count);
        picked.forEach(function(q) { usedIds.add(q.id); });
        return picked;
    }

    matrixRows.forEach(function(row) {
        if (row.p1_k > 0) selectedP1 = selectedP1.concat(pickFromPool(1, 'Biết', row.topic, row.p1_k));
        if (row.p1_h > 0) selectedP1 = selectedP1.concat(pickFromPool(1, 'Hiểu', row.topic, row.p1_h));
        if (row.p1_v > 0) selectedP1 = selectedP1.concat(pickFromPool(1, 'Vận dụng', row.topic, row.p1_v));

        var rowReq = getRequiredCountsForRow(row, (typeof currentExamFormat !== "undefined" ? currentExamFormat : "MIXED"));
        if (rowReq.p2 > 0) {
            selectedP2 = selectedP2.concat(pickFromPool(2, '', row.topic, rowReq.p2));
        }

        if (row.p3_k > 0) selectedP3 = selectedP3.concat(pickFromPool(3, 'Biết', row.topic, row.p3_k));
        if (row.p3_h > 0) selectedP3 = selectedP3.concat(pickFromPool(3, 'Hiểu', row.topic, row.p3_h));
        if (row.p3_v > 0) selectedP3 = selectedP3.concat(pickFromPool(3, 'Vận dụng', row.topic, row.p3_v));

        if (row.p4_k > 0) selectedP4 = selectedP4.concat(pickFromPool(4, 'Biết', row.topic, row.p4_k));
        if (row.p4_h > 0) selectedP4 = selectedP4.concat(pickFromPool(4, 'Hiểu', row.topic, row.p4_h));
        if (row.p4_v > 0) selectedP4 = selectedP4.concat(pickFromPool(4, 'Vận dụng', row.topic, row.p4_v));
    });

    while (selectedP1.length < targetP1) {
        var extra1 = combinedPool.filter(function(q) { return q.part === 1 && !usedIds.has(q.id); })[0];
        if (!extra1) break;
        usedIds.add(extra1.id);
        selectedP1.push(extra1);
    }
    while (selectedP2.length < targetP2) {
        var extra2 = combinedPool.filter(function(q) { return q.part === 2 && !usedIds.has(q.id); })[0];
        if (!extra2) break;
        usedIds.add(extra2.id);
        selectedP2.push(extra2);
    }
    while (selectedP3.length < targetP3) {
        var extra3 = combinedPool.filter(function(q) { return q.part === 3 && !usedIds.has(q.id); })[0];
        if (!extra3) break;
        usedIds.add(extra3.id);
        selectedP3.push(extra3);
    }
    while (selectedP4.length < targetP4) {
        var extra4 = combinedPool.filter(function(q) { return q.part === 4 && !usedIds.has(q.id); })[0];
        if (!extra4) break;
        usedIds.add(extra4.id);
        selectedP4.push(extra4);
    }

    return {
        part1: selectedP1.slice(0, targetP1),
        part2: selectedP2.slice(0, targetP2),
        part3: selectedP3.slice(0, targetP3),
        part4: selectedP4.slice(0, targetP4)
    };
}

/* 8. Nhập câu hỏi từ tệp Word (.docx) hoặc dán văn bản của giáo viên */
function openPasteModal() {
    var modal = document.getElementById("paste-modal");
    if (modal) modal.classList.remove("hidden");
}
function closePasteModal() {
    var modal = document.getElementById("paste-modal");
    if (modal) modal.classList.add("hidden");
}

function handleUserDocxUpload(e) {
    var file = e.target.files[0];
    if (!file) return;

    var badge = document.getElementById("upload-status-badge");
    if (badge) {
        badge.classList.remove("hidden");
        badge.innerText = "Đang xử lý tệp " + file.name + "...";
    }

    if (file.name.endsWith('.docx')) {
        var reader = new FileReader();
        reader.onload = function(loadEvent) {
            var arrayBuffer = loadEvent.target.result;
            if (window.mammoth) {
                mammoth.extractRawText({ arrayBuffer: arrayBuffer })
                    .then(function(result) {
                        var parsed = parseRawExamText(result.value, "File: " + file.name);
                        addParsedQuestionsToBank(parsed);
                        if (badge) badge.innerHTML = "<i class='fa-solid fa-check-circle mr-1'></i> Đã nhập thành công <strong>" + parsed.length + " câu hỏi</strong> từ tệp <em>" + file.name + "</em> vào Ngân hàng!";
                    })
                    .catch(function(err) {
                        if (badge) badge.innerText = "Lỗi đọc file: " + err.message;
                    });
            } else {
                if (badge) badge.innerText = "Thư viện đọc docx chưa sẵn sàng, vui lòng dùng nút 'Dán Văn Bản Đề Thi'.";
            }
        };
        reader.readAsArrayBuffer(file);
    } else {
        var reader2 = new FileReader();
        reader2.onload = function(loadEvent) {
            var parsed = parseRawExamText(loadEvent.target.result, "File: " + file.name);
            addParsedQuestionsToBank(parsed);
            if (badge) badge.innerHTML = "<i class='fa-solid fa-check-circle mr-1'></i> Đã nhập thành công <strong>" + parsed.length + " câu hỏi</strong> vào Ngân hàng!";
        };
        reader2.readAsText(file);
    }
}

function processPastedText() {
    var rawEl = document.getElementById("paste-textarea");
    var targetTopicEl = document.getElementById("paste-target-topic");
    var raw = rawEl ? rawEl.value : "";
    var targetTopic = targetTopicEl ? targetTopicEl.value : "Chủ đề tự chọn";

    if (!raw.trim()) {
        alert("Vui lòng dán nội dung câu hỏi!");
        return;
    }

    var parsed = parseRawExamText(raw, targetTopic);
    addParsedQuestionsToBank(parsed);
    closePasteModal();

    var badge = document.getElementById("upload-status-badge");
    if (badge) {
        badge.classList.remove("hidden");
        badge.innerHTML = "<i class='fa-solid fa-check-circle mr-1'></i> Đã nạp thành công <strong>" + parsed.length + " câu hỏi</strong> vào ngân hàng của khối Lớp " + currentGrade + "!";
    }
}

function parseRawExamText(fullText, defaultTopic) {
    var lines = fullText.split(/\r?\n/).map(function(l) { return l.trim(); }).filter(function(l) { return l.length > 0; });
    var questions = [];
    var currentQ = null;

    for (var i = 0; i < lines.length; i++) {
        var line = lines[i];
        var isQStart = /^câu\s+\d+[:.]/i.test(line) || /^bài\s+\d+[:.]/i.test(line);

        if (isQStart) {
            if (currentQ) finalizeParsedQuestion(currentQ, questions);
            currentQ = {
                id: "user_" + Date.now() + "_" + Math.floor(Math.random() * 10000),
                grade: currentGrade,
                topic: defaultTopic || "Chủ đề tự chọn",
                level: "Hiểu",
                indicator: "HH1.2",
                header: line,
                bodyLines: []
            };
        } else if (currentQ) {
            currentQ.bodyLines.push(line);
        }
    }
    if (currentQ) finalizeParsedQuestion(currentQ, questions);
    return questions;
}

function finalizeParsedQuestion(q, list) {
    var allText = [q.header].concat(q.bodyLines).join('\n');
    var hasStatements = /[a-d]\)\s+/i.test(allText);
    var hasMCQ = /[A-D][.:]\s+/.test(allText);

    if (hasStatements && (allText.toLowerCase().indexOf('đúng') !== -1 || allText.toLowerCase().indexOf('sai') !== -1 || q.bodyLines.some(function(l) { return /^[a-d]\)/i.test(l); }))) {
        q.content = q.header.replace(/^câu\s+\d+[:.]/i, '').trim();
        q.statements = [];
        q.bodyLines.forEach(function(l) {
            if (/^[a-d]\)/i.test(l) || /^\([1-4a-d]\)/i.test(l)) {
                q.statements.push({ text: l.trim(), correct: true });
            }
        });
        if (q.statements.length >= 2) {
            q.part = 2;
            list.push(q);
        } else {
            q.part = 4;
            q.content = allText.replace(/^câu\s+\d+[:.]/i, '').trim();
            q.solution = "Hướng dẫn giải chi tiết.";
            list.push(q);
        }
    } else if (hasMCQ) {
        q.content = q.header.replace(/^câu\s+\d+[:.]/i, '').trim();
        var opts = [];
        q.bodyLines.forEach(function(l) {
            if (/^[A-D][.:]/i.test(l)) {
                opts.push(l.replace(/^[A-D][.:]\s*/i, '').trim());
            }
        });
        // Kiểm tra nếu 4 phương án nằm trên 1 dòng
        if (opts.length < 4) {
            var mSingle = allText.match(/(?:^|[\s\n])A[.:\s]+(.*?)(?:[\s\n]+)B[.:\s]+(.*?)(?:[\s\n]+)C[.:\s]+(.*?)(?:[\s\n]+)D[.:\s]+(.*?)$/is);
            if (mSingle && mSingle[1] && mSingle[2] && mSingle[3] && mSingle[4]) {
                opts = [mSingle[1].trim(), mSingle[2].trim(), mSingle[3].trim(), mSingle[4].trim()];
            }
        }
        if (opts.length === 4) {
            q.part = 1;
            q.options = opts;
            q.correct = 0;
            list.push(q);
        } else {
            q.part = 4;
            q.content = allText.replace(/^câu\s+\d+[:.]/i, '').trim();
            q.solution = "Hướng dẫn chấm chi tiết.";
            list.push(q);
        }
    } else {
        if (allText.toLowerCase().indexOf('tính') !== -1 || allText.toLowerCase().indexOf('bao nhiêu') !== -1) {
            q.part = 3;
            q.content = allText.replace(/^câu\s+\d+[:.]/i, '').trim();
            q.answer = "10";
            q.explain = "Kết quả phân tích từ đề.";
        } else {
            q.part = 4;
            q.content = allText.replace(/^câu\s+\d+[:.]/i, '').trim();
            q.solution = "Hướng dẫn chấm chi tiết.";
        }
        list.push(q);
    }
}

function addParsedQuestionsToBank(questions) {
    if (!Array.isArray(questions)) return;
    var sanitizedList = questions.map(autoSanitizeQuestion);
    if (!QUESTION_BANK[currentGrade]) QUESTION_BANK[currentGrade] = [];
    QUESTION_BANK[currentGrade] = sanitizedList.concat(QUESTION_BANK[currentGrade]);
    if (typeof MASTER_QUESTION_BANK !== "undefined" && MASTER_QUESTION_BANK[currentGrade]) {
        MASTER_QUESTION_BANK[currentGrade] = sanitizedList.concat(MASTER_QUESTION_BANK[currentGrade]);
    }
    renderGeneratedExam();
    if (typeof filterBankQuestions === "function") {
        filterBankQuestions(true);
    }
}

/* =========================================================================================
   PHÂN HỆ QUẢN LÝ, TRA CỨU & KHAI THÁC NGÂN HÀNG CÂU HỎI (1.562 CÂU LỚP 10, 11, 12)
   ========================================================================================= */

var bankCurrentPage = 1;
var bankPageSize = 25;
var bankFilteredList = [];

/* Chuyển đổi giữa Chế độ xem Đề thi và Chế độ Tra cứu Ngân hàng câu hỏi */
function switchTab4View(view) {
    var examView = document.getElementById("tab4-exam-view");
    var bankView = document.getElementById("tab4-bank-view");
    var btnExam = document.getElementById("tab4-btn-exam");
    var btnBank = document.getElementById("tab4-btn-bank");

    if (view === "bank") {
        if (examView) examView.classList.add("hidden");
        if (bankView) bankView.classList.remove("hidden");
        if (btnExam) {
            btnExam.className = "px-4 py-2 text-xs font-bold rounded-lg text-gray-700 hover:bg-gray-100 transition flex items-center gap-2 border border-transparent";
        }
        if (btnBank) {
            btnBank.className = "px-4 py-2 text-xs font-bold rounded-lg bg-brand-500 text-white shadow-sm transition flex items-center gap-2";
        }
        filterBankQuestions(true);
    } else {
        if (bankView) bankView.classList.add("hidden");
        if (examView) examView.classList.remove("hidden");
        if (btnBank) {
            btnBank.className = "px-4 py-2 text-xs font-bold rounded-lg text-gray-700 hover:bg-gray-100 transition flex items-center gap-2 border border-transparent";
        }
        if (btnExam) {
            btnExam.className = "px-4 py-2 text-xs font-bold rounded-lg bg-brand-500 text-white shadow-sm transition flex items-center gap-2";
        }
    }
}

/* Khởi tạo Ngân hàng câu hỏi */
function initBankExplorer() {
    var totalQ = 0;
    ["10", "11", "12"].forEach(function(g) {
        if (QUESTION_BANK[g]) totalQ += QUESTION_BANK[g].length;
    });

    var badgeCount = document.getElementById("bank-badge-count");
    if (badgeCount) badgeCount.innerText = totalQ + "+";

    onBankGradeChange();
}

/* Khi người dùng thay đổi lọc Khối lớp trong Ngân hàng */
function onBankGradeChange() {
    var gradeEl = document.getElementById("bank-filter-grade");
    var chapterEl = document.getElementById("bank-filter-chapter");
    if (!gradeEl || !chapterEl) return;

    var selectedGrade = gradeEl.value;
    var chaptersSet = new Set();

    var pools = (selectedGrade === "ALL") ? [].concat(QUESTION_BANK["10"] || [], QUESTION_BANK["11"] || [], QUESTION_BANK["12"] || []) : (QUESTION_BANK[selectedGrade] || []);
    pools.forEach(function(q) {
        var ch = getCanonicalChapter(q.chapter || q.topic, q.grade || selectedGrade);
        if (ch) chaptersSet.add(ch);
    });

    var optHtml = "<option value='ALL'>Tất cả các chương / chuyên đề</option>";
    Array.from(chaptersSet).sort(function(a, b) {
        return a.localeCompare(b, "vi", { numeric: true });
    }).forEach(function(ch) {
        optHtml += "<option value='" + ch + "'>" + ch + "</option>";
    });
    chapterEl.innerHTML = optHtml;

    filterBankQuestions(true);
}

/* Lọc câu hỏi trong Ngân hàng theo tất cả tiêu chí */
function filterBankQuestions(resetPage) {
    if (resetPage) bankCurrentPage = 1;

    var gradeEl = document.getElementById("bank-filter-grade");
    var chapterEl = document.getElementById("bank-filter-chapter");
    var partEl = document.getElementById("bank-filter-part");
    var levelEl = document.getElementById("bank-filter-level");
    var indicatorEl = document.getElementById("bank-filter-indicator");
    var keywordEl = document.getElementById("bank-filter-keyword");

    var grade = gradeEl ? gradeEl.value : "ALL";
    var chapter = chapterEl ? chapterEl.value : "ALL";
    var part = partEl ? partEl.value : "ALL";
    var level = levelEl ? levelEl.value : "ALL";
    var indicator = indicatorEl ? indicatorEl.value : "ALL";
    var keyword = keywordEl ? keywordEl.value.trim().toLowerCase() : "";

    var allList = [];
    if (grade === "ALL") {
        allList = [].concat(QUESTION_BANK["10"] || [], QUESTION_BANK["11"] || [], QUESTION_BANK["12"] || []);
    } else {
        allList = [].concat(QUESTION_BANK[grade] || []);
    }

    bankFilteredList = allList.filter(function(q) {
        if (chapter !== "ALL") {
            var qChapter = getCanonicalChapter(q.chapter || q.topic, q.grade);
            if (qChapter !== chapter) return false;
        }
        if (part !== "ALL" && String(q.part) !== String(part)) return false;
        if (level !== "ALL" && (q.level || "").toLowerCase() !== level.toLowerCase()) return false;
        if (indicator !== "ALL") {
            var indText = (q.indicator || "").toUpperCase();
            if (indText.indexOf(indicator.toUpperCase()) === -1) return false;
        }
        if (keyword) {
            var fullSearch = (q.content + " " + (q.topic || "") + " " + (q.chapter || "") + " " + (q.explain || "") + " " + (q.solution || "")).toLowerCase();
            if (q.options) fullSearch += " " + q.options.join(" ").toLowerCase();
            if (q.statements) fullSearch += " " + q.statements.map(function(s) { return s.text; }).join(" ").toLowerCase();
            if (fullSearch.indexOf(keyword) === -1) return false;
        }
        return true;
    });

    // Thống kê phân bố
    var cntKnow = bankFilteredList.filter(function(q) { return (q.level || "").toLowerCase() === "biết"; }).length;
    var cntUnder = bankFilteredList.filter(function(q) { return (q.level || "").toLowerCase() === "hiểu"; }).length;
    var cntApply = bankFilteredList.filter(function(q) { return (q.level || "").toLowerCase().indexOf("vận dụng") !== -1; }).length;

    var statsSummary = document.getElementById("bank-stats-summary");
    if (statsSummary) {
        statsSummary.innerHTML = "Hiển thị: <strong>" + bankFilteredList.length + "</strong> câu hỏi phù hợp trong ngân hàng";
    }
    var elK = document.getElementById("bank-cnt-know");
    var elU = document.getElementById("bank-cnt-under");
    var elA = document.getElementById("bank-cnt-apply");
    if (elK) elK.innerText = cntKnow;
    if (elU) elU.innerText = cntUnder;
    if (elA) elA.innerText = cntApply;

    renderBankCards();
}

/* Hiển thị danh sách câu hỏi dạng Card */
function renderBankCards() {
    var container = document.getElementById("bank-questions-container");
    var loadMoreBtn = document.getElementById("bank-load-more-btn");
    if (!container) return;

    if (bankFilteredList.length === 0) {
        container.innerHTML = "<div class='bg-white p-8 rounded-xl border border-gray-200 text-center text-xs text-gray-500'>" +
            "<i class='fa-solid fa-folder-open text-gray-300 text-3xl mb-2 block'></i>" +
            "Không tìm thấy câu hỏi nào phù hợp với bộ lọc hiện tại. Vui lòng thử từ khóa hoặc chọn 'Tất cả'.</div>";
        if (loadMoreBtn) loadMoreBtn.classList.add("hidden");
        return;
    }

    var itemsToShow = bankFilteredList.slice(0, bankCurrentPage * bankPageSize);
    if (loadMoreBtn) {
        if (itemsToShow.length < bankFilteredList.length) loadMoreBtn.classList.remove("hidden");
        else loadMoreBtn.classList.add("hidden");
    }

    var partNames = { 1: "Phần I: Nhiều phương án", 2: "Phần II: Đúng/Sai", 3: "Phần III: TL Ngắn", 4: "Phần IV: Tự Luận" };
    var levelBadges = {
        "Biết": "bg-emerald-100 text-emerald-900 border-emerald-300",
        "Hiểu": "bg-blue-100 text-blue-900 border-blue-300",
        "Vận dụng": "bg-purple-100 text-purple-900 border-purple-300"
    };

    container.innerHTML = itemsToShow.map(function(q, idx) {
        var lvlClass = levelBadges[q.level] || "bg-gray-100 text-gray-800 border-gray-300";
        var partName = partNames[q.part] || ("Phần " + q.part);
        var indicatorStr = q.indicator || "HH1.1";

        var bodyHtml = "";
        if (q.part === 1 && q.options) {
            bodyHtml = "<div class='space-y-1.5 pt-2 pl-2'>" +
                q.options.map(function(opt, optIdx) {
                    var isCorrect = (q.correct === optIdx);
                    var optStyle = isCorrect ? "bg-emerald-50 border-emerald-300 text-emerald-900 font-semibold" : "bg-gray-50/70 border-gray-200 text-gray-800";
                    return "<div class='p-2 rounded-lg border text-xs " + optStyle + "'>" +
                        "<strong>" + String.fromCharCode(65 + optIdx) + ".</strong> " + cleanLatexForWord(opt) +
                        (isCorrect ? " <i class='fa-solid fa-check text-emerald-600 ml-1'></i>" : "") +
                        "</div>";
                }).join("") + "</div>";
        } else if (q.part === 2 && q.statements) {
            bodyHtml = "<div class='space-y-1.5 pt-2 pl-2'>" +
                q.statements.map(function(s) {
                    var statusBadge = s.correct ? "<span class='text-emerald-700 font-bold'>[Đúng]</span>" : "<span class='text-rose-700 font-bold'>[Sai]</span>";
                    return "<div class='text-xs flex items-start gap-2 text-gray-800'>" +
                        "<span>" + cleanLatexForWord(s.text) + "</span> " + statusBadge +
                        "</div>";
                }).join("") + "</div>";
        } else if (q.part === 3) {
            bodyHtml = "<div class='pt-2 text-xs font-semibold text-purple-800 flex items-center gap-2'>" +
                "<span>Đáp số: <strong>" + cleanLatexForWord(q.answer || "10") + "</strong></span>" +
                "</div>";
        } else if (q.part === 4 && (q.solution || q.explain)) {
            var sol = formatDetailedExplanation(q.solution || q.explain, q);
            bodyHtml = "<div class='pt-2 text-xs bg-rose-50/50 p-2.5 rounded-lg border border-rose-100 text-gray-800'>" +
                "<p class='font-bold text-rose-900 mb-1'>Hướng dẫn giải chi tiết:</p>" +
                "<div class='leading-relaxed'>" + formatSolutionHtml(sol) + "</div>" +
                "</div>";
        }

        var detailedExp = formatDetailedExplanation(q.explain, q);
        var explainHtml = "";
        if (detailedExp && q.part !== 4) {
            explainHtml = "<div class='mt-2 pt-2 border-t border-gray-100 text-[11px] text-gray-600 italic'>" +
                "<strong>Giải thích:</strong> " + cleanLatexForWord(detailedExp) +
                "</div>";
        }

        return "<div class='bg-white p-4 rounded-xl border border-gray-200 shadow-sm space-y-2 hover:border-brand-300 transition'>" +
            "<div class='flex flex-wrap items-center justify-between gap-2 border-b border-gray-100 pb-2'>" +
            "<div class='flex flex-wrap items-center gap-1.5'>" +
            "<span class='bg-brand-50 text-brand-700 text-xs px-2 py-0.5 rounded font-bold'>Lớp " + q.grade + "</span>" +
            "<span class='bg-gray-100 text-gray-700 text-xs px-2 py-0.5 rounded font-semibold'>" + partName + "</span>" +
            "<span class='text-xs px-2 py-0.5 rounded font-bold border " + lvlClass + "'>" + q.level + "</span>" +
            "<span class='bg-amber-100 text-amber-900 border border-amber-300 text-xs px-2 py-0.5 rounded-full font-bold'>Mã: " + indicatorStr + "</span>" +
            "<span class='text-xs text-gray-600 font-medium truncate max-w-xs' title='" + (q.chapter || q.topic) + "'>• " + (q.chapter || q.topic) + "</span>" +
            "</div>" +
            "<span class='text-[10px] text-gray-400'>#" + (idx + 1) + "</span>" +
            "</div>" +
            "<div class='text-xs font-semibold text-gray-800 leading-relaxed'>" + cleanLatexForWord(q.content) + "</div>" +
            bodyHtml +
            explainHtml +
            "</div>";
    }).join("");

    if (typeof triggerKaTeX === "function") {
        triggerKaTeX();
    }
}

/* Tải thêm câu hỏi phân trang */
function loadMoreBankQuestions() {
    bankCurrentPage++;
    renderBankCards();
}

/* Đặt lại bộ lọc ngân hàng */
function resetBankFilters() {
    var gradeEl = document.getElementById("bank-filter-grade");
    var partEl = document.getElementById("bank-filter-part");
    var levelEl = document.getElementById("bank-filter-level");
    var indicatorEl = document.getElementById("bank-filter-indicator");
    var keywordEl = document.getElementById("bank-filter-keyword");

    if (gradeEl) gradeEl.value = "ALL";
    if (partEl) partEl.value = "ALL";
    if (levelEl) levelEl.value = "ALL";
    if (indicatorEl) indicatorEl.value = "ALL";
    if (keywordEl) keywordEl.value = "";

    onBankGradeChange();
}

/* Xuất danh sách câu hỏi đang lọc ra file Word (.doc) */
function exportFilteredBankToWord() {
    if (bankFilteredList.length === 0) {
        alert("Không có câu hỏi nào để xuất file!");
        return;
    }

    var gradeEl = document.getElementById("bank-filter-grade");
    var chapterEl = document.getElementById("bank-filter-chapter");
    var gradeVal = gradeEl ? gradeEl.value : "ALL";
    var chapVal = chapterEl ? chapterEl.value : "ALL";

    var title = "BỘ CÂU HỎI HÓA HỌC PHÂN LOẠI (KNTT - GDPT 2018)";
    if (gradeVal !== "ALL") title += " - KHỐI " + gradeVal;
    if (chapVal !== "ALL") title += " - " + chapVal.toUpperCase();

    var html = "<table class='header-table'>" +
        "<tr><td class='text-center' style='width: 45%;'>" +
        "<p class='bold'>BỘ CÂU HỎI THAM KHẢO & GIẢNG DẠY</p>" +
        "<p class='bold'>HÓA HỌC GDPT 2018 (CV 7991)</p></td>" +
        "<td class='text-center' style='width: 55%;'>" +
        "<p class='bold'>" + title + "</p>" +
        "<p class='italic'>Số lượng câu hỏi: " + bankFilteredList.length + " câu (Kèm đáp án & chỉ báo)</p></td></tr></table>" +
        "<hr style='border: 1px solid #000; margin-bottom: 15px;'/>";

    bankFilteredList.forEach(function(q, i) {
        var partTag = q.part === 1 ? "Trắc nghiệm" : (q.part === 2 ? "Đúng/Sai" : (q.part === 3 ? "Trả lời ngắn" : "Tự luận"));
        html += "<div style='margin-bottom: 12px;'>" +
            "<p><strong>Câu " + (i + 1) + ".</strong> [" + q.level + " - " + q.indicator + " - " + cleanLatexForWord(q.chapter || q.topic) + "]: " + cleanLatexForWord(q.content) + "</p>";

        if (q.part === 1 && q.options) {
            html += "<div style='margin-left: 18px; margin-bottom: 6px;'>" +
                formatWordOption("A", q.options[0]) +
                formatWordOption("B", q.options[1]) +
                formatWordOption("C", q.options[2]) +
                formatWordOption("D", q.options[3]) +
                "</div>";
        } else if (q.part === 2 && q.statements) {
            html += "<div style='margin-left: 20px;'>" +
                q.statements.map(function(s) { return "<p style='margin: 2px 0;'>" + cleanLatexForWord(s.text) + "</p>"; }).join("") +
                "</div>";
        } else if (q.part === 3) {
            html += "<p style='margin-left: 20px; color: #444;'><em>Đáp án: .......................................</em></p>";
        }
        html += "</div>";
    });

    // Bảng đáp án
    html += "<div class='page-break'></div><h3 class='text-center bold' style='margin-top: 15px;'>BẢNG ĐÁP ÁN VÀ HƯỚNG DẪN GIẢI CHI TIẾT</h3>" +
        "<p class='text-center bold' style='color: #1e3a8a; margin: 2px 0 10px 0; font-size: 10.5pt;'>Giáo viên biên soạn: GV. NGUYỄN VỸ — Trường THPT Bùi Thị Xuân (Mũi Né)</p>" +
        "<table class='data-table answer-table' style='width: 100%; table-layout: fixed; mso-table-layout-alt: fixed; border-collapse: collapse;'>" +
        "<colgroup>" +
        "<col style='width: 8%;' width='8%'/>" +
        "<col style='width: 44%;' width='44%'/>" +
        "<col style='width: 24%;' width='24%'/>" +
        "<col style='width: 12%;' width='12%'/>" +
        "<col style='width: 12%;' width='12%'/>" +
        "</colgroup>" +
        "<thead><tr>" +
        "<th style='width: 8%; text-align: center;'>Câu</th>" +
        "<th style='width: 44%; text-align: center;'>Đáp Án / Hướng Dẫn Giải</th>" +
        "<th style='width: 24%; text-align: center;'>Chuyên Đề</th>" +
        "<th style='width: 12%; text-align: center;'>Mức Độ</th>" +
        "<th style='width: 12%; text-align: center;'>Mã Chỉ Báo</th>" +
        "</tr></thead><tbody>";

    bankFilteredList.forEach(function(q, i) {
        var ansStr = "";
        if (q.part === 1) ansStr = String.fromCharCode(65 + q.correct);
        else if (q.part === 2 && q.statements) ansStr = q.statements.map(function(s, idx) { return String.fromCharCode(97 + idx) + ") " + (s.correct ? "Đ" : "S"); }).join(" - ");
        else if (q.part === 3) ansStr = q.answer || "10";
        else ansStr = q.solution || "Xem lời giải chi tiết";

        html += "<tr><td class='bold text-center' style='width: 8%;'>" + (i + 1) + "</td>" +
            "<td style='width: 44%; word-wrap: break-word;'>" + formatSolutionHtml(ansStr) + "</td>" +
            "<td style='width: 24%; word-wrap: break-word;'>" + cleanLatexForWord(q.chapter || q.topic) + "</td>" +
            "<td class='text-center' style='width: 12%;'>" + q.level + "</td>" +
            "<td class='text-center bold' style='width: 12%;'>" + q.indicator + "</td></tr>";
    });

    html += "</tbody></table>";
    var fileName = "Ngan_Hang_Cau_Hoi_Hoa_" + gradeVal + "_" + Date.now() + ".doc";
    downloadWordDocument(html, fileName);
}

/* Modal Thêm câu hỏi thủ công */
function openAddQuestionModal() {
    var modal = document.getElementById("manual-question-modal");
    if (modal) modal.classList.remove("hidden");
    onManualQPartChange();
}

function closeAddQuestionModal() {
    var modal = document.getElementById("manual-question-modal");
    if (modal) modal.classList.add("hidden");
}

function onManualQPartChange() {
    var partEl = document.getElementById("manual-q-part");
    var part1Box = document.getElementById("manual-q-part1-box");
    var part2Box = document.getElementById("manual-q-part2-box");
    var part3Box = document.getElementById("manual-q-part3-box");

    var part = partEl ? parseInt(partEl.value) : 1;
    if (part1Box) part1Box.className = (part === 1) ? "space-y-2" : "space-y-2 hidden";
    if (part2Box) part2Box.className = (part === 2) ? "space-y-2" : "space-y-2 hidden";
    if (part3Box) part3Box.className = (part === 3) ? "space-y-2" : "space-y-2 hidden";
}

function saveManualQuestion() {
    var grade = document.getElementById("manual-q-grade").value;
    var part = parseInt(document.getElementById("manual-q-part").value);
    var level = document.getElementById("manual-q-level").value;
    var chapter = document.getElementById("manual-q-chapter").value.trim() || "Chủ đề tự chọn";
    var indicator = document.getElementById("manual-q-indicator").value;
    var content = document.getElementById("manual-q-content").value.trim();
    var solution = document.getElementById("manual-q-solution").value.trim();

    if (!content) {
        alert("Vui lòng nhập nội dung câu hỏi!");
        return;
    }

    var qObj = {
        id: "custom_" + grade + "_" + Date.now(),
        grade: grade,
        chapter: chapter,
        topic: chapter,
        level: level,
        part: part,
        indicator: indicator,
        content: content,
        solution: solution || "Hướng dẫn chấm chi tiết.",
        source: "Giáo viên tự nhập"
    };

    if (part === 1) {
        var optA = document.getElementById("manual-opt-a").value.trim();
        var optB = document.getElementById("manual-opt-b").value.trim();
        var optC = document.getElementById("manual-opt-c").value.trim();
        var optD = document.getElementById("manual-opt-d").value.trim();
        if (!optA || !optB || !optC || !optD) {
            alert("Vui lòng nhập đầy đủ cả 4 phương án A, B, C, D!");
            return;
        }
        var correct = parseInt(document.getElementById("manual-opt-correct").value) || 0;
        qObj.options = [optA, optB, optC, optD];
        qObj.correct = correct;
        qObj.explain = solution || "Đáp án đúng theo phương án đã chọn.";
    } else if (part === 2) {
        var sa = document.getElementById("manual-stmt-a").value.trim();
        var sb = document.getElementById("manual-stmt-b").value.trim();
        var sc = document.getElementById("manual-stmt-c").value.trim();
        var sd = document.getElementById("manual-stmt-d").value.trim();
        if (!sa || !sb || !sc || !sd) {
            alert("Vui lòng nhập đầy đủ nội dung cho cả 4 ý a, b, c, d!");
            return;
        }
        var va = document.getElementById("manual-stmt-a-val").value === "true";
        var vb = document.getElementById("manual-stmt-b-val").value === "true";
        var vc = document.getElementById("manual-stmt-c-val").value === "true";
        var vd = document.getElementById("manual-stmt-d-val").value === "true";
        qObj.statements = [
            { text: "a) " + sa, correct: va },
            { text: "b) " + sb, correct: vb },
            { text: "c) " + sc, correct: vc },
            { text: "d) " + sd, correct: vd }
        ];
        qObj.explain = solution || "Hướng dẫn chấm 4 ý a, b, c, d.";
    } else if (part === 3) {
        var shortAns = document.getElementById("manual-short-ans").value.trim();
        if (!shortAns) {
            alert("Vui lòng nhập đáp số cho câu hỏi trả lời ngắn!");
            return;
        }
        qObj.answer = shortAns;
        qObj.explain = solution || ("Đáp số: " + shortAns);
    }

    qObj = autoSanitizeQuestion(qObj);
    if (!QUESTION_BANK[grade]) QUESTION_BANK[grade] = [];
    QUESTION_BANK[grade].unshift(qObj);

    if (typeof MASTER_QUESTION_BANK !== "undefined" && MASTER_QUESTION_BANK[grade]) {
        MASTER_QUESTION_BANK[grade].unshift(qObj);
    }

    closeAddQuestionModal();
    alert("Đã thêm thành công câu hỏi vào Ngân hàng của Khối Lớp " + grade + "!");

    initBankExplorer();
    renderGeneratedExam();
}

/* =========================================================================================
   HỆ THỐNG BIÊN SOẠN & XUẤT ĐỀ CƯƠNG ÔN TẬP CHO HỌC SINH (KNTT & CV 7991)
   ========================================================================================= */

var KNTT_TERM_LESSONS_MAP = {
    "10": {
        "GK1": ["g10_c1_b1", "g10_c1_b2", "g10_c1_b3", "g10_c2_b5", "g10_c2_b6", "g10_c2_b7"],
        "CK1": ["g10_c1_b1", "g10_c1_b2", "g10_c1_b3", "g10_c2_b5", "g10_c2_b6", "g10_c2_b7", "g10_c3_b9", "g10_c3_b10", "g10_c3_b11", "g10_c3_b12", "g10_c4_b13"],
        "GK2": ["g10_c4_b13", "g10_c5_b15", "g10_c5_b16", "g10_c6_b17", "g10_c6_b18"],
        "CK2": ["g10_c4_b13", "g10_c5_b15", "g10_c5_b16", "g10_c6_b17", "g10_c6_b18", "g10_c7_b19", "g10_c7_b20"]
    },
    "11": {
        "GK1": ["g11_c1_b1", "g11_c1_b2", "g11_c2_b3", "g11_c2_b4", "g11_c2_b5"],
        "CK1": ["g11_c1_b1", "g11_c1_b2", "g11_c2_b3", "g11_c2_b4", "g11_c2_b5", "g11_c2_b6", "g11_c2_b7", "g11_c3_b8", "g11_c3_b9", "g11_c3_b10", "g11_c3_b11", "g11_c4_b12"],
        "GK2": ["g11_c4_b13", "g11_c4_b14", "g11_c5_b15", "g11_c5_b16", "g11_c5_b17"],
        "CK2": ["g11_c4_b13", "g11_c4_b14", "g11_c5_b15", "g11_c5_b16", "g11_c5_b17", "g11_c6_b18", "g11_c6_b19"]
    },
    "12": {
        "GK1": ["g12_c1_b1", "g12_c1_b2", "g12_c2_b3", "g12_c2_b4", "g12_c2_b5", "g12_c3_b6", "g12_c3_b7"],
        "CK1": ["g12_c1_b1", "g12_c1_b2", "g12_c2_b3", "g12_c2_b4", "g12_c2_b5", "g12_c3_b6", "g12_c3_b7", "g12_c3_b8", "g12_c4_b9", "g12_c4_b10", "g12_c5_b12", "g12_c5_b13"],
        "GK2": ["g12_c5_b12", "g12_c5_b13", "g12_c6_b14", "g12_c6_b15", "g12_c6_b16", "g12_c7_b18", "g12_c7_b19", "g12_c7_b20"],
        "CK2": ["g12_c5_b12", "g12_c5_b13", "g12_c6_b14", "g12_c6_b15", "g12_c6_b16", "g12_c7_b18", "g12_c7_b19", "g12_c7_b20", "g12_c8_b21", "g12_c8_b22"]
    }
};

function onSyllabusGradeChange() {
    var sel = document.getElementById("syllabus-grade-select");
    if (!sel) return;
    currentGrade = sel.value;
    var examGradeSel = document.getElementById("exam-grade-select");
    if (examGradeSel) examGradeSel.value = currentGrade;
    if (typeof onGradeSelectChange === "function") {
        onGradeSelectChange();
    }
    initSyllabusLessonRangePickers(true);
    renderSyllabus();
}

function onSyllabusTermChange() {
    var sel = document.getElementById("syllabus-term-select");
    if (!sel) return;
    currentTerm = sel.value;
    var knttTermSel = document.getElementById("kntt-term-select");
    if (knttTermSel) knttTermSel.value = currentTerm;
    if (typeof applyKNTTTermPreset === "function") {
        applyKNTTTermPreset();
    }
    initSyllabusLessonRangePickers(true);
    renderSyllabus();
}

// Khởi tạo các lựa chọn khoảng bài học (Từ bài ... Đến bài ...) cho đề cương
function initSyllabusLessonRangePickers(forcePreset) {
    var gradeSel = document.getElementById("syllabus-grade-select");
    var termSel = document.getElementById("syllabus-term-select");
    var startSel = document.getElementById("syllabus-start-lesson-select");
    var endSel = document.getElementById("syllabus-end-lesson-select");
    if (!startSel || !endSel) return;

    var grade = gradeSel ? gradeSel.value : (currentGrade || "12");
    var term = termSel ? termSel.value : (currentTerm || "GK1");
    var allLessons = (typeof KNTT_LESSONS_STANDARDS !== "undefined" && KNTT_LESSONS_STANDARDS[grade]) ? KNTT_LESSONS_STANDARDS[grade] : [];
    if (!allLessons || allLessons.length === 0) return;

    var prevStartVal = startSel.value;
    var prevEndVal = endSel.value;

    startSel.innerHTML = "";
    endSel.innerHTML = "";

    allLessons.forEach(function(l, idx) {
        var chapPrefix = (l.chapter || "").split(":")[0] || ("C" + (idx + 1));
        var optText = l.lesson + " (" + chapPrefix + ")";

        var opt1 = document.createElement("option");
        opt1.value = l.id;
        opt1.textContent = optText;
        startSel.appendChild(opt1);

        var opt2 = document.createElement("option");
        opt2.value = l.id;
        opt2.textContent = optText;
        endSel.appendChild(opt2);
    });

    // Mặc định cho Giữa kỳ 1 (GK1) theo từng khối:
    var defaultStartId = allLessons[0].id;
    var defaultEndId = allLessons[Math.min(allLessons.length - 1, 6)].id;

    if (term === "GK1") {
        if (grade === "10") {
            // Khối 10 GK1: Bài 1 (Thành phần nguyên tử) -> Bài 7 (Định luật tuần hoàn & Ý nghĩa BTH)
            defaultStartId = "g10_c1_b1";
            defaultEndId = "g10_c2_b7";
        } else if (grade === "11") {
            // Khối 11 GK1: Bài 1 (Cân bằng hóa học) -> Bài 5 (Hợp chất của nitrogen với oxygen)
            defaultStartId = "g11_c1_b1";
            defaultEndId = "g11_c2_b5";
        } else if (grade === "12") {
            // Khối 12 GK1: Bài 1 (Ester) -> Bài 7 (Amino acid & Peptide)
            defaultStartId = "g12_c1_b1";
            defaultEndId = "g12_c3_b7";
        }
    } else {
        var termLessons = (KNTT_TERM_LESSONS_MAP[grade] && KNTT_TERM_LESSONS_MAP[grade][term]) ? KNTT_TERM_LESSONS_MAP[grade][term] : [];
        if (termLessons.length > 0) {
            defaultStartId = termLessons[0];
            defaultEndId = termLessons[termLessons.length - 1];
        }
    }

    if (forcePreset || !prevStartVal || !allLessons.some(function(l){ return l.id === prevStartVal; })) {
        startSel.value = defaultStartId;
    } else {
        startSel.value = prevStartVal;
    }

    if (forcePreset || !prevEndVal || !allLessons.some(function(l){ return l.id === prevEndVal; })) {
        endSel.value = defaultEndId;
    } else {
        endSel.value = prevEndVal;
    }

    updateSyllabusRangeBadge();
}

// Cập nhật huy hiệu hiển thị khoảng bài học đã chọn
function updateSyllabusRangeBadge() {
    var gradeSel = document.getElementById("syllabus-grade-select");
    var startSel = document.getElementById("syllabus-start-lesson-select");
    var endSel = document.getElementById("syllabus-end-lesson-select");
    var badge = document.getElementById("syllabus-range-badge");
    if (!startSel || !endSel || !badge) return;

    var grade = gradeSel ? gradeSel.value : currentGrade;
    var allLessons = (typeof KNTT_LESSONS_STANDARDS !== "undefined" && KNTT_LESSONS_STANDARDS[grade]) ? KNTT_LESSONS_STANDARDS[grade] : [];
    if (!allLessons || allLessons.length === 0) return;

    var startIdx = allLessons.findIndex(function(l) { return l.id === startSel.value; });
    var endIdx = allLessons.findIndex(function(l) { return l.id === endSel.value; });

    if (startIdx === -1) startIdx = 0;
    if (endIdx === -1) endIdx = allLessons.length - 1;

    if (startIdx > endIdx) {
        endIdx = startIdx;
        endSel.value = allLessons[endIdx].id;
    }

    var count = endIdx - startIdx + 1;
    var startName = allLessons[startIdx] ? allLessons[startIdx].lesson.split(":")[0] : "Bài bắt đầu";
    var endName = allLessons[endIdx] ? allLessons[endIdx].lesson.split(":")[0] : "Bài kết thúc";

    if (startIdx === endIdx) {
        badge.textContent = "Đang chọn: " + startName + " (1 bài học)";
    } else {
        badge.textContent = "Đang chọn: Từ " + startName + " đến " + endName + " (" + count + " bài học)";
    }
}

function onSyllabusLessonRangeChange() {
    updateSyllabusRangeBadge();
    renderSyllabus();
}

function applySyllabusGK1DefaultPreset() {
    initSyllabusLessonRangePickers(true);
    renderSyllabus();
}

function applySyllabusSelectAllLessons() {
    var gradeSel = document.getElementById("syllabus-grade-select");
    var startSel = document.getElementById("syllabus-start-lesson-select");
    var endSel = document.getElementById("syllabus-end-lesson-select");
    if (!startSel || !endSel) return;
    var grade = gradeSel ? gradeSel.value : currentGrade;
    var allLessons = (typeof KNTT_LESSONS_STANDARDS !== "undefined" && KNTT_LESSONS_STANDARDS[grade]) ? KNTT_LESSONS_STANDARDS[grade] : [];
    if (allLessons.length > 0) {
        startSel.value = allLessons[0].id;
        endSel.value = allLessons[allLessons.length - 1].id;
        updateSyllabusRangeBadge();
        renderSyllabus();
    }
}

function onSyllabusScopeChange() {
    var scopeSel = document.getElementById("syllabus-scope-select");
    var rangeBox = document.getElementById("syllabus-lesson-range-box");
    if (rangeBox && scopeSel) {
        if (scopeSel.value === "CUSTOM_RANGE") {
            rangeBox.classList.remove("opacity-50");
        } else {
            rangeBox.classList.add("opacity-50");
        }
    }
    renderSyllabus();
}

// Lấy danh sách các bài học thuộc phạm vi đề cương
function getSyllabusLessons() {
    var grade = document.getElementById("syllabus-grade-select") ? document.getElementById("syllabus-grade-select").value : currentGrade;
    var term = document.getElementById("syllabus-term-select") ? document.getElementById("syllabus-term-select").value : currentTerm;
    var scope = document.getElementById("syllabus-scope-select") ? document.getElementById("syllabus-scope-select").value : "CUSTOM_RANGE";

    var allLessons = (typeof KNTT_LESSONS_STANDARDS !== "undefined" && KNTT_LESSONS_STANDARDS[grade]) ? KNTT_LESSONS_STANDARDS[grade] : [];
    if (!allLessons || allLessons.length === 0) return [];

    var startSel = document.getElementById("syllabus-start-lesson-select");
    var endSel = document.getElementById("syllabus-end-lesson-select");

    // Nếu chọn CUSTOM_RANGE hoặc mặc định
    if ((scope === "CUSTOM_RANGE" || !scope) && startSel && endSel && startSel.value && endSel.value) {
        var startIdx = allLessons.findIndex(function(l) { return l.id === startSel.value; });
        var endIdx = allLessons.findIndex(function(l) { return l.id === endSel.value; });

        if (startIdx === -1) startIdx = 0;
        if (endIdx === -1) endIdx = allLessons.length - 1;
        if (startIdx > endIdx) {
            var tmp = startIdx;
            startIdx = endIdx;
            endIdx = tmp;
        }
        return allLessons.slice(startIdx, endIdx + 1);
    }

    if (scope === "MATRIX_TOPICS" && matrixRows && matrixRows.length > 0) {
        var matchedLessons = [];
        var activeChapterKeys = new Set();
        matrixRows.forEach(function(r) {
            var k = getCanonicalChapterKey(grade, r.topic);
            if (k) activeChapterKeys.add(k);
        });

        allLessons.forEach(function(l) {
            var lKey = getCanonicalChapterKey(grade, (l.chapter || "") + " " + (l.lesson || ""));
            if (lKey && activeChapterKeys.has(lKey)) {
                if (matchedLessons.indexOf(l) === -1) {
                    matchedLessons.push(l);
                }
            }
        });
        if (matchedLessons.length > 0) return matchedLessons;
    }

    var termIds = (KNTT_TERM_LESSONS_MAP[grade] && KNTT_TERM_LESSONS_MAP[grade][term]) ? KNTT_TERM_LESSONS_MAP[grade][term] : [];
    if (termIds.length > 0) {
        return allLessons.filter(function(l) { return termIds.indexOf(l.id) !== -1; });
    }
    return allLessons.slice(0, 6);
}

// Thu thập danh sách các đề ôn tập (khoảng 4 đến 5 đề) hoặc bài tập theo chuyên đề cho đề cương
function getSyllabusExamData(lessons) {
    var grade = document.getElementById("syllabus-grade-select") ? document.getElementById("syllabus-grade-select").value : currentGrade;
    var term = document.getElementById("syllabus-term-select") ? document.getElementById("syllabus-term-select").value : currentTerm;
    var scale = document.getElementById("syllabus-scale-select") ? document.getElementById("syllabus-scale-select").value : "EXAM_SET_5";
    var pool = QUESTION_BANK[grade] || [];

    // Bộ lọc nghiêm ngặt đảm bảo 100% câu hỏi chuẩn: không câu hỏi rỗng, không nhận định a-d rỗng, không phương án giả, không cụm từ đứt quãng
    function isValidExamQuestion(q) {
        if (!q || !q.content || q.content.trim().length < 15) return false;
        var c = q.content.trim();
        // Kiểm tra câu hỏi bị mất công thức hóa học: ", , ," hoặc "Tia ."
        if (/,\s*,\s*,/.test(c) || /\bTia\s*\./.test(c)) return false;
        // Kiểm tra câu hỏi chỉ có mở đầu mà không có nội dung phát biểu
        if (/^(?:\[[^\]]+\]\s*)?cho các phát biểu(?: sau)?\s*[:.]?$/i.test(c)) return false;

        if (q.part === 1) {
            if (!Array.isArray(q.options) || q.options.length !== 4) return false;
            return !q.options.some(function(o) {
                if (typeof o !== "string" || o.trim().length < 1) return true;
                if (/^(?:[A-D]\.\s*)?cả [a-d] và\s*$/i.test(o.trim())) return true;
                if (o.indexOf("Phương án") !== -1 || o.indexOf("Đáp án") !== -1) return true;
                return false;
            });
        }
        if (q.part === 2) {
            if (!Array.isArray(q.statements) || q.statements.length < 4) return false;
            return !q.statements.some(function(s) {
                if (typeof s.text !== "string" || s.text.trim().length <= 5) return true;
                if (/^nhận định [a-d]/i.test(s.text.trim()) || /^[a-d][.)]\s*$/i.test(s.text.trim())) return true;
                return false;
            });
        }
        if (q.part === 3) {
            return !!(q.answer || q.explain);
        }
        if (q.part === 4) {
            var sol = (q.solution || q.explain || "").trim();
            return sol.length >= 25 && c.length >= 25;
        }
        return true;
    }

    // Xác định tập các chương được phép theo phạm vi đề cương
    var allowedChapterKeys = new Set();
    lessons.forEach(function(l) {
        var k = getCanonicalChapterKey(grade, (l.chapter || "") + " " + (l.lesson || ""));
        if (k) allowedChapterKeys.add(k);
    });

    if (allowedChapterKeys.size === 0) {
        if (grade === "10") {
            if (term === "GK1") { allowedChapterKeys.add("CH_10_1"); allowedChapterKeys.add("CH_10_2"); }
            else if (term === "CK1") { allowedChapterKeys.add("CH_10_1"); allowedChapterKeys.add("CH_10_2"); allowedChapterKeys.add("CH_10_3"); allowedChapterKeys.add("CH_10_4"); }
            else if (term === "GK2") { allowedChapterKeys.add("CH_10_4"); allowedChapterKeys.add("CH_10_5"); allowedChapterKeys.add("CH_10_6"); }
            else if (term === "CK2") { allowedChapterKeys.add("CH_10_4"); allowedChapterKeys.add("CH_10_5"); allowedChapterKeys.add("CH_10_6"); allowedChapterKeys.add("CH_10_7"); }
        } else if (grade === "11") {
            if (term === "GK1") { allowedChapterKeys.add("CH_11_1"); allowedChapterKeys.add("CH_11_2"); }
            else if (term === "CK1") { allowedChapterKeys.add("CH_11_1"); allowedChapterKeys.add("CH_11_2"); allowedChapterKeys.add("CH_11_3"); allowedChapterKeys.add("CH_11_4"); }
            else if (term === "GK2") { allowedChapterKeys.add("CH_11_4"); allowedChapterKeys.add("CH_11_5"); }
            else if (term === "CK2") { allowedChapterKeys.add("CH_11_4"); allowedChapterKeys.add("CH_11_5"); allowedChapterKeys.add("CH_11_6"); }
        } else if (grade === "12") {
            if (term === "GK1") { allowedChapterKeys.add("CH_12_1"); allowedChapterKeys.add("CH_12_2"); allowedChapterKeys.add("CH_12_3"); }
            else if (term === "CK1") { allowedChapterKeys.add("CH_12_1"); allowedChapterKeys.add("CH_12_2"); allowedChapterKeys.add("CH_12_3"); allowedChapterKeys.add("CH_12_4"); allowedChapterKeys.add("CH_12_5"); }
            else if (term === "GK2") { allowedChapterKeys.add("CH_12_5"); allowedChapterKeys.add("CH_12_6"); allowedChapterKeys.add("CH_12_7"); }
            else if (term === "CK2") { allowedChapterKeys.add("CH_12_5"); allowedChapterKeys.add("CH_12_6"); allowedChapterKeys.add("CH_12_7"); allowedChapterKeys.add("CH_12_8"); }
        }
    }

    var allGradeLessons = (typeof KNTT_LESSONS_STANDARDS !== "undefined" && KNTT_LESSONS_STANDARDS[grade]) ? KNTT_LESSONS_STANDARDS[grade] : [];
    var selectedLessonIds = new Set(lessons.map(function(l) { return l.id; }));
    var excludedLessons = allGradeLessons.filter(function(l) { return !selectedLessonIds.has(l.id); });

    function isQuestionInExcludedLessons(q) {
        if (!excludedLessons || excludedLessons.length === 0) return false;
        var fullText = ((q.content || "") + " " + (q.chapter || "") + " " + (q.topic || "") + " " + (q.lesson || "")).toLowerCase();
        for (var i = 0; i < excludedLessons.length; i++) {
            var ex = excludedLessons[i];
            var exChapKey = getCanonicalChapterKey(grade, (ex.chapter || "") + " " + (ex.lesson || ""));
            if (!allowedChapterKeys.has(exChapKey)) continue;

            if (ex.id === "g12_c3_b8") {
                if (/\b(?:protein|enzym|enzim|biuret|lòng trắng trứng|đông tụ protein)\b/i.test(fullText)) return true;
            } else if (ex.id === "g12_c2_b5") {
                if (/\b(?:tinh bột|cellulose|xenlulozơ|xenlulozo|amilozơ|amilopectin)\b/i.test(fullText)) return true;
            } else if (ex.id === "g12_c2_b4") {
                if (/\b(?:saccharose|saccarozơ|saccarozo|maltose|mantozơ)\b/i.test(fullText)) return true;
            } else if (ex.id === "g12_c1_b2") {
                if (/\b(?:chất béo|lipid|lipit|triglyceride|triglixerit|triolein|tripalmitin|tristearin)\b/i.test(fullText) && !/\beste\b/i.test(fullText)) return true;
            } else if (ex.id === "g11_c2_b6" || ex.id === "g11_c2_b7") {
                if (/\b(?:sulfur|lưu huỳnh|so2|sulfuric acid|h2so4|muối sulfate)\b/i.test(fullText)) return true;
            } else if (ex.id === "g10_c2_b7") {
                if (/\b(?:định luật tuần hoàn|ý nghĩa của bảng tuần hoàn)\b/i.test(fullText)) return true;
            }
        }
        return false;
    }

    var scopePool = pool.filter(function(q) {
        if (!isValidExamQuestion(q)) return false;
        // Loại bỏ câu thuộc kỳ thi khác nếu có gắn tag term
        if (q.term && q.term !== term) return false;
        // Kiểm tra đúng chương nằm trong danh mục cho phép
        var stmtsText = (q.statements && Array.isArray(q.statements)) ? q.statements.map(function(s){ return s.text || ""; }).join(" ") : "";
        var contentKey = getCanonicalChapterKey(grade, (q.content || "") + " " + stmtsText);
        var chapKey = getCanonicalChapterKey(grade, q.chapter || q.topic || "");
        var qKey = chapKey || contentKey;

        if (chapKey && contentKey && chapKey !== contentKey) {
            if (!allowedChapterKeys.has(chapKey) && allowedChapterKeys.has(contentKey)) {
                if (contentKey === "CH_12_1" && /\b(?:ester|este|chất béo|lipid|lipit|triglyceride)\b/i.test(q.content || "")) {
                    qKey = contentKey;
                    q.chapter = "Chương 1: Ester - Lipid";
                    q.topic = "Chương 1: Ester - Lipid";
                }
            } else if (allowedChapterKeys.has(chapKey)) {
                qKey = chapKey;
            }
        }

        if (!qKey || !allowedChapterKeys.has(qKey)) return false;
        if (isQuestionInExcludedLessons(q)) return false;
        return true;
    });

    // Nếu số câu trong phạm vi vẫn chưa phân loại hết, lấy thêm các câu có nội dung khớp với allowedChapterKeys
    if (scopePool.length < 50) {
        var extra = pool.filter(function(q) {
            if (!isValidExamQuestion(q) || scopePool.indexOf(q) !== -1) return false;
            if (q.term && q.term !== term) return false;
            var qKey = getCanonicalChapterKey(grade, q.content || "");
            if (!allowedChapterKeys.has(qKey)) return false;
            if (isQuestionInExcludedLessons(q)) return false;
            return true;
        });
        scopePool = scopePool.concat(extra);
    }

    var numExams = 5;
    var singleExamIndex = 1;
    var isTopical = false;

    if (scale === "EXAM_SET_5") {
        numExams = 5;
    } else if (scale === "EXAM_SET_4") {
        numExams = 4;
    } else if (scale.indexOf("EXAM_") === 0) {
        numExams = 1;
        singleExamIndex = parseInt(scale.replace("EXAM_", "")) || 1;
    } else if (scale === "TOPICAL_BANK" || scale === "COMPACT" || scale === "STANDARD" || scale === "EXTENSIVE") {
        isTopical = true;
    }

    if (isTopical) {
        return {
            mode: "TOPICAL",
            exams: [{
                examNumber: 1,
                name: "HỆ THỐNG BÀI TẬP ÔN TẬP THEO CHUYÊN ĐỀ",
                part1: scopePool.filter(function(q) { return q.part === 1; }).slice(0, 32),
                part2: scopePool.filter(function(q) { return q.part === 2; }).slice(0, 6),
                part3: scopePool.filter(function(q) { return q.part === 3; }).slice(0, 8),
                part4: scopePool.filter(function(q) { return q.part === 4; }).slice(0, 6)
            }]
        };
    }

    // Sinh các đề ôn tập độc lập KHÔNG TRÙNG LẶP NỘI DUNG GIỮA CÁC ĐỀ
    var usedIds = new Set();
    var usedContentKeys = new Set();

    function getNormalizedKey(q) {
        if (!q || !q.content) return "";
        return q.content.toLowerCase().replace(/[^a-z0-9àáảãạăằắẳẵặâầấẩẫậèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđ]/g, "").slice(0, 60);
    }

    function isAvailable(q) {
        if (usedIds.has(q.id)) return false;
        var nk = getNormalizedKey(q);
        if (nk && usedContentKeys.has(nk)) return false;
        return true;
    }

    function markUsed(q) {
        usedIds.add(q.id);
        var nk = getNormalizedKey(q);
        if (nk) usedContentKeys.add(nk);
    }

    var exams = [];
    var count = numExams;

    function pickBatch(part, reqCount) {
        var picked = [];
        for (var i = 0; i < scopePool.length; i++) {
            var q = scopePool[i];
            if (q.part === part && isAvailable(q)) {
                markUsed(q);
                picked.push(q);
                if (picked.length >= reqCount) break;
            }
        }
        return picked;
    }

    if (numExams === 1 && singleExamIndex > 1) {
        for (var s = 1; s < singleExamIndex; s++) {
            pickBatch(1, 12);
            pickBatch(2, 2);
            pickBatch(3, 4);
            pickBatch(4, 3);
        }
    }

    for (var e = 1; e <= count; e++) {
        var examNum = (numExams === 1) ? singleExamIndex : e;
        var p1 = pickBatch(1, 12);
        var p2 = pickBatch(2, 2);
        var p3 = pickBatch(3, 4);
        var p4 = pickBatch(4, 3);

        exams.push({
            examNumber: examNum,
            name: "ĐỀ ÔN TẬP SỐ " + examNum + " (CHUẨN CV 7991 - 10,0 ĐIỂM)",
            part1: p1,
            part2: p2,
            part3: p3,
            part4: p4
        });
    }

    return {
        mode: "EXAMS",
        exams: exams
    };
}

// Hiển thị giao diện Đề Cương Ôn Tập trên Web (Tab 6)
function renderSyllabus() {
    var container = document.getElementById("syllabus-preview-container");
    if (!container) return;

    var startSel = document.getElementById("syllabus-start-lesson-select");
    if (startSel && (!startSel.options || startSel.options.length === 0)) {
        if (typeof initSyllabusLessonRangePickers === "function") {
            initSyllabusLessonRangePickers();
        }
    }

    var grade = document.getElementById("syllabus-grade-select") ? document.getElementById("syllabus-grade-select").value : currentGrade;
    var term = document.getElementById("syllabus-term-select") ? document.getElementById("syllabus-term-select").value : currentTerm;
    var audience = document.getElementById("syllabus-audience-select") ? document.getElementById("syllabus-audience-select").value : "STUDENT_ANS_END";
    var title = document.getElementById("syllabus-title-input") ? document.getElementById("syllabus-title-input").value : "ĐỀ CƯƠNG ÔN TẬP KIỂM TRA HÓA HỌC";

    var termName = (term === "GK1") ? "GIỮA HỌC KỲ 1" : (term === "CK1") ? "CUỐI HỌC KỲ 1" : (term === "GK2") ? "GIỮA HỌC KỲ 2" : "CUỐI HỌC KỲ 2";
    var lessons = getSyllabusLessons();
    var examData = getSyllabusExamData(lessons);

    var rangeDesc = "";
    if (lessons && lessons.length > 0) {
        var startName = lessons[0].lesson ? lessons[0].lesson.split(":")[0] : "Bài bắt đầu";
        var endName = lessons[lessons.length - 1].lesson ? lessons[lessons.length - 1].lesson.split(":")[0] : "Bài kết thúc";
        rangeDesc = (startName === endName) ? startName : ("Từ " + startName + " đến " + endName);
    }

    var html = "<div class='border-b-2 border-brand-600 pb-4 mb-6'>" +
        "<div class='flex flex-col sm:flex-row items-center justify-between text-xs text-gray-700'>" +
        "<div class='text-center sm:text-left font-bold'>" +
        "<p>TRƯỜNG THPT BÙI THỊ XUÂN - MŨI NÉ</p>" +
        "<p class='text-brand-700 uppercase'>TỔ HÓA HỌC - BỘ SÁCH KẾT NỐI TRI THỨC</p>" +
        "<p class='text-amber-700 font-extrabold text-[11px] mt-0.5 flex items-center justify-center sm:justify-start gap-1'><i class='fa-solid fa-chalkboard-user'></i> GV: NGUYỄN VỸ</p>" +
        "</div>" +
        "<div class='text-center sm:text-right mt-2 sm:mt-0'>" +
        "<p class='font-bold'>TÀI LIỆU ÔN TẬP TOÀN DIỆN CHO HỌC SINH</p>" +
        "<p class='italic text-gray-500'>Năm học: 2025 - 2026 • Chuẩn CV 7991</p>" +
        "<p class='text-xs font-semibold text-brand-700 mt-0.5'>Biên soạn: GV. Nguyễn Vỹ</p>" +
        "</div>" +
        "</div>" +
        "<div class='text-center mt-4'>" +
        "<h1 class='text-lg sm:text-xl font-extrabold text-brand-800 uppercase tracking-wide'>" + title + "</h1>" +
        "<p class='text-xs font-semibold text-gray-600 mt-1 uppercase'>MÔN: HÓA HỌC " + grade + " • KỲ: " + termName + (rangeDesc ? " • PHẠM VI: " + rangeDesc + " (" + lessons.length + " BÀI HỌC)" : "") + " • BỘ " + examData.exams.length + " ĐỀ ÔN TẬP RÈN LUYỆN</p>" +
        "</div>" +
        "</div>";

    // PHẦN A: LÝ THUYẾT TRỌNG TÂM & BẢNG YCCĐ
    html += "<div class='space-y-4'>" +
        "<div class='bg-brand-50 p-3 rounded-xl border border-brand-200 flex items-center justify-between'>" +
        "<h2 class='font-bold text-sm sm:text-base text-brand-900 flex items-center gap-2'>" +
        "<i class='fa-solid fa-book-bookmark text-brand-600'></i>" +
        "<span>PHẦN A: HỆ THỐNG LÝ THUYẾT TRỌNG TÂM & YÊU CẦU CẦN ĐẠT (YCCĐ)</span>" +
        "</h2>" +
        "<span class='text-xs font-semibold text-brand-700 bg-white px-2.5 py-1 rounded-lg border border-brand-200'>" + lessons.length + " Bài học trọng tâm</span>" +
        "</div>";

    lessons.forEach(function(lesson, lIdx) {
        var knowList = (lesson.knowItems || []).map(function(item) {
            return "<li><span class='font-bold text-emerald-800'>[" + item.code + "]</span> " + item.text + "</li>";
        }).join("");
        var underList = (lesson.underItems || []).map(function(item) {
            return "<li><span class='font-bold text-blue-800'>[" + item.code + "]</span> " + item.text + "</li>";
        }).join("");
        var applyList = (lesson.applyItems || []).map(function(item) {
            return "<li><span class='font-bold text-purple-800'>[" + item.code + "]</span> " + item.text + "</li>";
        }).join("");

        html += "<div class='p-4 bg-white rounded-xl border border-gray-200 shadow-xs space-y-3'>" +
            "<h3 class='font-bold text-xs sm:text-sm text-gray-900 flex items-center gap-2 border-b border-gray-100 pb-2'>" +
            "<span class='w-6 h-6 rounded-lg bg-brand-100 text-brand-800 flex items-center justify-center text-xs font-bold'>" + (lIdx + 1) + "</span>" +
            "<span>" + lesson.chapter + " — " + lesson.lesson + "</span>" +
            "</h3>" +
            "<div class='grid grid-cols-1 md:grid-cols-3 gap-3 text-xs leading-relaxed'>" +
            "<div class='p-3 bg-emerald-50/40 rounded-lg border border-emerald-100'>" +
            "<h4 class='font-bold text-emerald-900 mb-1 flex items-center gap-1'><i class='fa-solid fa-check text-emerald-600'></i> Nhận Biết:</h4>" +
            "<ul class='space-y-1 list-disc pl-4 text-gray-700'>" + (knowList || "<li>Nắm vững định nghĩa, danh pháp và tính chất đặc trưng.</li>") + "</ul>" +
            "</div>" +
            "<div class='p-3 bg-blue-50/40 rounded-lg border border-blue-100'>" +
            "<h4 class='font-bold text-blue-900 mb-1 flex items-center gap-1'><i class='fa-solid fa-lightbulb text-blue-600'></i> Thông Hiểu:</h4>" +
            "<ul class='space-y-1 list-disc pl-4 text-gray-700'>" + (underList || "<li>Giải thích hiện tượng, quy luật biến đổi và phương trình hóa học.</li>") + "</ul>" +
            "</div>" +
            "<div class='p-3 bg-purple-50/40 rounded-lg border border-purple-100'>" +
            "<h4 class='font-bold text-purple-900 mb-1 flex items-center gap-1'><i class='fa-solid fa-bullseye text-purple-600'></i> Vận Dụng:</h4>" +
            "<ul class='space-y-1 list-disc pl-4 text-gray-700'>" + (applyList || "<li>Tính toán hóa học thực nghiệm, liên hệ thực tiễn sản xuất.</li>") + "</ul>" +
            "</div>" +
            "</div>" +
            "</div>";
    });
    html += "</div>";

    // PHẦN B: HỆ THỐNG CÁC ĐỀ ÔN TẬP
    html += "<div class='space-y-6 mt-8'>" +
        "<div class='bg-amber-50 p-3.5 rounded-xl border border-amber-200 flex items-center justify-between'>" +
        "<h2 class='font-bold text-sm sm:text-base text-amber-900 flex items-center gap-2'>" +
        "<i class='fa-solid fa-file-signature text-amber-600'></i>" +
        "<span>PHẦN B: HỆ THỐNG CÁC ĐỀ ÔN TẬP (GỒM " + examData.exams.length + " ĐỀ ĐỘC LẬP - CHUẨN CV 7991)</span>" +
        "</h2>" +
        "<span class='text-xs font-semibold text-amber-800 bg-white px-2.5 py-1 rounded-lg border border-amber-200'>" +
        "Mỗi đề: 12 câu Nhiều phương án (3đ) + 2 Đ/S 8 ý (2đ) + 4 TLN (2đ) + 3 TL (3đ) = 10,0đ" +
        "</span>" +
        "</div>";

    examData.exams.forEach(function(exam, eIdx) {
        html += "<div class='bg-white p-5 rounded-2xl border-2 border-brand-100 shadow-sm space-y-5'>" +
            "<div class='bg-gradient-to-r from-teal-50 to-emerald-50 p-3 rounded-xl border border-teal-200 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-2'>" +
            "<div>" +
            "<h3 class='font-extrabold text-sm sm:text-base text-teal-900 uppercase flex items-center gap-2'>" +
            "<i class='fa-solid fa-pen-to-square text-teal-600'></i>" +
            "<span>" + exam.name + "</span>" +
            "</h3>" +
            "<p class='text-xs text-teal-700'>Cấu trúc chuẩn: 21 câu/lệnh hỏi • Thời gian làm bài: 45 phút • Thang điểm: 10,0</p>" +
            "</div>" +
            "<span class='bg-white text-teal-800 font-bold px-3 py-1 rounded-lg border border-teal-200 text-xs shadow-xs'>ĐỀ SỐ " + (eIdx + 1) + "</span>" +
            "</div>";

        // 1. Phần I
        if (exam.part1.length > 0) {
            html += "<div class='space-y-3'>" +
                "<h4 class='font-bold text-xs sm:text-sm text-green-900 border-b border-green-100 pb-1.5 flex items-center gap-1.5'>" +
                "<span class='bg-green-600 text-white w-5 h-5 rounded-full flex items-center justify-center text-[10px] font-bold'>I</span>" +
                "<span>PHẦN I: CÂU TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (12 câu - 3,00 điểm)</span>" +
                "</h4>" +
                "<div class='space-y-2.5'>";
            exam.part1.forEach(function(q, qIdx) {
                html += "<div class='text-xs leading-relaxed p-2.5 rounded-lg border border-gray-100 bg-gray-50/30'>" +
                    "<p class='font-semibold text-gray-900'><span class='text-green-800 font-bold'>Câu " + (qIdx + 1) + ".</span> [" + q.level + (q.indicator ? " - " + q.indicator : "") + "] " + cleanLatexForWord(q.content) + "</p>" +
                    "<div class='space-y-1 mt-2 text-[11px] pl-3'>";
                (q.options || []).forEach(function(opt, oIdx) {
                    var isKey = (audience === "TEACHER_KEY" && oIdx === q.correct);
                    html += "<div class='p-1.5 rounded " + (isKey ? "bg-green-100 font-bold text-green-900 border border-green-300" : "bg-white border border-gray-200 text-gray-700") + "'>" +
                        "<strong>" + String.fromCharCode(65 + oIdx) + ".</strong> " + cleanLatexForWord(opt) +
                        "</div>";
                });
                html += "</div>";
                if (audience === "TEACHER_KEY") {
                    var expP1 = formatDetailedExplanation(q.explain, q);
                    html += "<p class='mt-1.5 text-[11px] text-green-800 bg-green-50/60 p-1.5 rounded italic'><strong>Đáp án:</strong> " + String.fromCharCode(65 + q.correct) + " — " + cleanLatexForWord(expP1) + "</p>";
                }
                html += "</div>";
            });
            html += "</div></div>";
        }

        // 2. Phần II
        if (exam.part2.length > 0) {
            html += "<div class='space-y-3 pt-2'>" +
                "<h4 class='font-bold text-xs sm:text-sm text-blue-900 border-b border-blue-100 pb-1.5 flex items-center gap-1.5'>" +
                "<span class='bg-blue-600 text-white w-5 h-5 rounded-full flex items-center justify-center text-[10px] font-bold'>II</span>" +
                "<span>PHẦN II: CÂU TRẮC NGHIỆM ĐÚNG / SAI (2 câu lớn - 8 ý lệnh hỏi - 2,00 điểm)</span>" +
                "</h4>" +
                "<div class='space-y-3'>";
            exam.part2.forEach(function(q, qIdx) {
                html += "<div class='text-xs leading-relaxed p-3 rounded-lg border border-blue-100 bg-blue-50/20'>" +
                    "<p class='font-semibold text-gray-900'><span class='text-blue-800 font-bold'>Câu " + (qIdx + 1) + ".</span> [" + q.level + (q.indicator ? " - " + q.indicator : "") + "] " + cleanLatexForWord(q.content) + "</p>" +
                    "<div class='space-y-1.5 mt-2.5'>";
                (q.statements || []).forEach(function(st, sIdx) {
                    var meta = getStatementMeta(q, sIdx);
                    var metaBadge = "<span class='inline-block bg-blue-100 text-blue-800 font-bold text-[10px] px-1.5 py-0.2 rounded border border-blue-200 mr-1.5'>[" + meta.level + " - " + meta.indicator + "]</span>";
                    var keyBadge = (audience === "TEACHER_KEY") ? ("<span class='font-bold text-[10px] px-2 py-0.5 rounded " + (st.correct ? "bg-green-100 text-green-800" : "bg-rose-100 text-rose-800") + "'>" + (st.correct ? "ĐÚNG" : "SAI") + "</span>") : "<span class='text-gray-400 font-mono text-[11px]'>[Đ / S]</span>";
                    html += "<div class='flex items-center justify-between p-1.5 bg-white rounded border border-gray-200 text-[11px]'>" +
                        "<span>" + metaBadge + cleanLatexForWord(st.text) + "</span>" +
                        keyBadge +
                        "</div>";
                });
                html += "</div>";
                if (audience === "TEACHER_KEY") {
                    var expP2 = formatDetailedExplanation(q.explain, q);
                    html += "<p class='mt-2 text-[11px] text-blue-800 bg-blue-50/60 p-1.5 rounded italic'><strong>Hướng dẫn giải:</strong> " + cleanLatexForWord(expP2) + "</p>";
                }
                html += "</div>";
            });
            html += "</div></div>";
        }

        // 3. Phần III
        if (exam.part3.length > 0) {
            html += "<div class='space-y-3 pt-2'>" +
                "<h4 class='font-bold text-xs sm:text-sm text-purple-900 border-b border-purple-100 pb-1.5 flex items-center gap-1.5'>" +
                "<span class='bg-purple-600 text-white w-5 h-5 rounded-full flex items-center justify-center text-[10px] font-bold'>III</span>" +
                "<span>PHẦN III: CÂU TRẮC NGHIỆM TRẢ LỜI NGẮN (4 câu - 2,00 điểm)</span>" +
                "</h4>" +
                "<div class='space-y-2.5'>";
            exam.part3.forEach(function(q, qIdx) {
                var expP3 = formatDetailedExplanation(q.explain, q);
                html += "<div class='text-xs leading-relaxed p-2.5 rounded-lg border border-purple-100 bg-purple-50/20'>" +
                    "<p class='font-semibold text-gray-900'><span class='text-purple-800 font-bold'>Câu " + (qIdx + 1) + ".</span> [" + q.level + (q.indicator ? " - " + q.indicator : "") + "] " + cleanLatexForWord(q.content) + "</p>";
                if (audience === "TEACHER_KEY") {
                    html += "<div class='mt-1.5 p-1.5 bg-white rounded border border-purple-200 text-[11px] text-purple-900 font-bold'>" +
                        "Đáp số: " + cleanLatexForWord(q.answer || "N/A") + (expP3 ? " — <em>" + cleanLatexForWord(expP3) + "</em>" : "") +
                        "</div>";
                } else {
                    html += "<div class='mt-1.5 flex items-center gap-2 text-[11px] text-gray-500'>" +
                        "<span>Đáp số: </span><span class='inline-block w-28 border-b border-dotted border-gray-400'></span>" +
                        "</div>";
                }
                html += "</div>";
            });
            html += "</div></div>";
        }

        // 4. Phần IV
        if (exam.part4.length > 0) {
            html += "<div class='space-y-3 pt-2'>" +
                "<h4 class='font-bold text-xs sm:text-sm text-rose-900 border-b border-rose-100 pb-1.5 flex items-center gap-1.5'>" +
                "<span class='bg-rose-600 text-white w-5 h-5 rounded-full flex items-center justify-center text-[10px] font-bold'>IV</span>" +
                "<span>PHẦN IV: CÂU HỎI & BÀI TẬP TỰ LUẬN (3 câu - 3,00 điểm)</span>" +
                "</h4>" +
                "<div class='space-y-3'>";
            exam.part4.forEach(function(q, qIdx) {
                html += "<div class='text-xs leading-relaxed p-3 rounded-lg border border-rose-100 bg-rose-50/20'>" +
                    "<p class='font-semibold text-gray-900'><span class='text-rose-800 font-bold'>Câu " + (qIdx + 1) + ".</span> [" + q.level + (q.indicator ? " - " + q.indicator : "") + "] " + cleanLatexForWord(q.content).replace(/\n/g, "<br/>") + "</p>";
                if (audience === "TEACHER_KEY") {
                    html += "<div class='mt-2.5 p-2.5 bg-white rounded-lg border border-rose-200 text-[11px] text-gray-800 leading-relaxed'>" +
                        "<strong class='text-rose-800 block mb-1'>Hướng dẫn giải chi tiết & Biểu điểm:</strong>" +
                        formatSolutionHtml(q.solution || q.explain || "") +
                        "</div>";
                }
                html += "</div>";
            });
            html += "</div></div>";
        }

        html += "</div>";
    });
    html += "</div>";

    // PHẦN C: ĐÁP ÁN Ở CUỐI (Nếu chọn chế độ STUDENT_ANS_END)
    if (audience === "STUDENT_ANS_END") {
        html += "<div class='space-y-6 mt-8 pt-6 border-t-2 border-dashed border-gray-300 page-break'>" +
            "<div class='bg-gray-100 p-3.5 rounded-xl border border-gray-300 text-center'>" +
            "<h2 class='font-bold text-sm sm:text-base text-gray-800 uppercase tracking-wide'>" +
            "PHẦN C: BẢNG ĐÁP ÁN & HƯỚNG DẪN GIẢI CHI TIẾT TỪNG ĐỀ ÔN TẬP" +
            "</h2>" +
            "<p class='text-xs text-gray-500 mt-0.5'>Đáp án ngắn gọn, đầy đủ phương trình hóa học và phép tính chuẩn xác không lỗi.</p>" +
            "</div>";

        examData.exams.forEach(function(exam, eIdx) {
            html += "<div class='bg-white p-4 rounded-xl border border-gray-200 shadow-xs space-y-4'>" +
                "<h3 class='font-bold text-xs sm:text-sm text-brand-900 border-b pb-1.5 flex items-center justify-between'>" +
                "<span>BẢNG ĐÁP ÁN VÀ LỜI GIẢI — " + exam.name + "</span>" +
                "<span class='text-[11px] text-gray-500 font-normal'>Đề số " + (eIdx + 1) + "</span>" +
                "</h3>";

            // Bảng P.I & P.III
            html += "<div class='grid grid-cols-1 md:grid-cols-2 gap-3 text-xs'>" +
                "<div>" +
                "<h4 class='font-bold text-green-900 mb-1.5'>1. Đáp án Phần I (Nhiều phương án 12 câu):</h4>" +
                "<div class='grid grid-cols-4 sm:grid-cols-6 gap-1.5 text-center font-bold'>";
            exam.part1.forEach(function(q, qIdx) {
                html += "<div class='p-1 bg-green-50 rounded border border-green-200 text-green-900'>" +
                    (qIdx + 1) + ": " + String.fromCharCode(65 + q.correct) +
                    "</div>";
            });
            html += "</div></div>";

            html += "<div>" +
                "<h4 class='font-bold text-purple-900 mb-1.5'>2. Đáp án Phần III (TL Ngắn 4 câu):</h4>" +
                "<div class='grid grid-cols-2 gap-1.5 text-center font-bold'>";
            exam.part3.forEach(function(q, qIdx) {
                html += "<div class='p-1 bg-purple-50 rounded border border-purple-200 text-purple-900 truncate' title='" + q.answer + "'>" +
                    (qIdx + 1) + ": " + (q.answer || "N/A") +
                    "</div>";
            });
            html += "</div></div></div>";

            // Bảng P.II
            html += "<div class='text-xs'>" +
                "<h4 class='font-bold text-blue-900 mb-1.5'>3. Đáp án Phần II (Đúng / Sai 2 câu lớn):</h4>" +
                "<div class='grid grid-cols-1 sm:grid-cols-2 gap-2 font-medium'>";
            exam.part2.forEach(function(q, qIdx) {
                var stStr = (q.statements || []).map(function(s, sIdx) {
                    return String.fromCharCode(97 + sIdx) + ": " + (s.correct ? "Đ" : "S");
                }).join("  |  ");
                html += "<div class='p-2 bg-blue-50/60 rounded border border-blue-200'>" +
                    "<strong class='text-blue-900'>Câu " + (qIdx + 1) + ":</strong> " + stStr +
                    "</div>";
            });
            html += "</div></div>";

            // Lời giải vắn tắt P.IV
            if (exam.part4.length > 0) {
                html += "<div class='text-xs space-y-2'>" +
                    "<h4 class='font-bold text-rose-900'>4. Lời giải vắn tắt Phần IV (Tự luận 3 câu):</h4>";
                exam.part4.forEach(function(q, qIdx) {
                    html += "<div class='p-2.5 bg-rose-50/30 rounded border border-rose-100 text-[11px] leading-relaxed'>" +
                        "<strong class='text-rose-900 block mb-0.5'>Câu " + (qIdx + 1) + " (1,0đ):</strong>" +
                        formatSolutionHtml(q.solution || q.explain || "Xem phương trình hóa học và công thức tính toán chuẩn SGK.") +
                        "</div>";
                });
                html += "</div>";
            }

            html += "</div>";
        });
        html += "</div>";
    }

    container.innerHTML = html;
    if (typeof triggerKaTeX === "function") triggerKaTeX();
}

// Xuất file Word (.doc) cho Đề Cương Ôn Tập (Học Sinh) hỗ trợ nhiều đề thi độc lập
function exportSyllabusToWord() {
    var grade = document.getElementById("syllabus-grade-select") ? document.getElementById("syllabus-grade-select").value : currentGrade;
    var term = document.getElementById("syllabus-term-select") ? document.getElementById("syllabus-term-select").value : currentTerm;
    var audience = document.getElementById("syllabus-audience-select") ? document.getElementById("syllabus-audience-select").value : "STUDENT_ANS_END";
    var title = document.getElementById("syllabus-title-input") ? document.getElementById("syllabus-title-input").value : "ĐỀ CƯƠNG ÔN TẬP KIỂM TRA HÓA HỌC";

    var termName = (term === "GK1") ? "GIỮA HỌC KỲ 1" : (term === "CK1") ? "CUỐI HỌC KỲ 1" : (term === "GK2") ? "GIỮA HỌC KỲ 2" : "CUỐI HỌC KỲ 2";
    var lessons = getSyllabusLessons();
    var examData = getSyllabusExamData(lessons);

    var rangeDesc = "";
    if (lessons && lessons.length > 0) {
        var startName = lessons[0].lesson ? lessons[0].lesson.split(":")[0] : "Bài bắt đầu";
        var endName = lessons[lessons.length - 1].lesson ? lessons[lessons.length - 1].lesson.split(":")[0] : "Bài kết thúc";
        rangeDesc = (startName === endName) ? startName : ("Từ " + startName + " đến " + endName);
    }

    var html = "<table class='header-table'>" +
        "<tr><td class='text-center' style='width: 45%;'>" +
        "<p class='bold'>TRƯỜNG THPT BÙI THỊ XUÂN - MŨI NÉ</p>" +
        "<p class='bold'>TỔ HÓA HỌC</p>" +
        "<p style='margin: 2px 0; font-weight: bold; color: #1e3a8a; font-size: 11pt;'>GV: NGUYỄN VỸ</p></td>" +
        "<td class='text-center' style='width: 55%;'>" +
        "<p class='bold'>TÀI LIỆU ÔN TẬP TOÀN DIỆN CHO HỌC SINH</p>" +
        "<p class='italic'>Bộ sách: Kết Nối Tri Thức Với Cuộc Sống - Lớp " + grade + "</p>" +
        "<p style='margin: 2px 0; font-size: 10pt; color: #1e3a8a;'><strong>Biên soạn & Giảng dạy: GV. Nguyễn Vỹ</strong></p></td></tr></table>" +
        "<div class='text-center' style='margin-top: 15px; margin-bottom: 15px;'>" +
        "<h2 class='bold' style='font-size: 14pt; margin: 0;'>" + cleanLatexForWord(title) + "</h2>" +
        "<p class='italic' style='margin: 3px 0;'>Môn: Hóa Học " + grade + " • Kỳ: " + termName + (rangeDesc ? " • Phạm vi ôn tập: " + cleanLatexForWord(rangeDesc) + " (" + lessons.length + " bài học)" : "") + " • Hệ Thống " + examData.exams.length + " Đề Ôn Tập Chuẩn CV 7991</p>" +
        "</div>";

    // PHẦN A: LÝ THUYẾT TRỌNG TÂM
    html += "<h3 style='background-color: #f1f5f9; padding: 4px 8px; border-left: 4px solid #1e5f54; font-size: 12pt;' class='bold'>PHẦN A: HỆ THỐNG LÝ THUYẾT TRỌNG TÂM & YÊU CẦU CẦN ĐẠT (YCCĐ)</h3>";
    lessons.forEach(function(lesson, lIdx) {
        html += "<p style='margin-top: 8px; margin-bottom: 4px;' class='bold'>" + (lIdx + 1) + ". " + cleanLatexForWord(lesson.chapter) + " — " + cleanLatexForWord(lesson.lesson) + "</p>" +
            "<table class='data-table'><thead><tr><th style='width: 33%; background-color: #f8fafc;'>Nhận Biết</th><th style='width: 33%; background-color: #f8fafc;'>Thông Hiểu</th><th style='width: 34%; background-color: #f8fafc;'>Vận Dụng</th></tr></thead><tbody><tr>" +
            "<td style='vertical-align: top;'><ul style='margin: 0; padding-left: 15px;'>" +
            (lesson.knowItems || []).map(function(it) { return "<li><strong>[" + it.code + "]</strong> " + cleanLatexForWord(it.text) + "</li>"; }).join("") +
            "</ul></td>" +
            "<td style='vertical-align: top;'><ul style='margin: 0; padding-left: 15px;'>" +
            (lesson.underItems || []).map(function(it) { return "<li><strong>[" + it.code + "]</strong> " + cleanLatexForWord(it.text) + "</li>"; }).join("") +
            "</ul></td>" +
            "<td style='vertical-align: top;'><ul style='margin: 0; padding-left: 15px;'>" +
            (lesson.applyItems || []).map(function(it) { return "<li><strong>[" + it.code + "]</strong> " + cleanLatexForWord(it.text) + "</li>"; }).join("") +
            "</ul></td>" +
            "</tr></tbody></table>";
    });

    // PHẦN B: CÁC ĐỀ ÔN TẬP
    examData.exams.forEach(function(exam, eIdx) {
        html += "<div class='page-break'></div>";
        html += "<div style='text-align: center; margin-bottom: 12px; border-bottom: 2px solid #0f766e; padding-bottom: 6px;'>" +
            "<h3 class='bold' style='font-size: 13pt; margin: 0; color: #0f766e;'>" + cleanLatexForWord(exam.name) + "</h3>" +
            "<p class='italic' style='margin: 2px 0;'>Thời gian làm bài: 45 phút • Điểm tối đa: 10,0 điểm • Không kể thời gian phát đề</p>" +
            "</div>";

        // P1
        if (exam.part1.length > 0) {
            html += "<p class='bold' style='margin-top: 8px; color: #047857;'>PHẦN I: CÂU TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (12 câu - 3,00 điểm)</p>";
            exam.part1.forEach(function(q, qIdx) {
                html += "<p style='margin-top: 5px; margin-bottom: 3px;'><strong>Câu " + (qIdx + 1) + ".</strong> [" + q.level + (q.indicator ? " - " + q.indicator : "") + "] " + cleanLatexForWord(q.content) + "</p>" +
                    "<div style='margin-left: 18px; margin-bottom: 6px;'>" +
                    formatWordOption("A", q.options[0]) +
                    formatWordOption("B", q.options[1]) +
                    formatWordOption("C", q.options[2]) +
                    formatWordOption("D", q.options[3]) +
                    "</div>";
                if (audience === "TEACHER_KEY") {
                    var expP1w = formatDetailedExplanation(q.explain, q);
                    html += "<p style='margin: 2px 0 5px 10px; font-size: 10pt; color: #047857; font-style: italic;'><strong>Đáp án:</strong> " + String.fromCharCode(65 + q.correct) + " — " + cleanLatexForWord(expP1w) + "</p>";
                }
            });
        }

        // P2
        if (exam.part2.length > 0) {
            html += "<p class='bold' style='margin-top: 10px; color: #1d4ed8;'>PHẦN II: CÂU TRẮC NGHIỆM ĐÚNG / SAI (2 câu lớn - 8 ý lệnh hỏi - 2,00 điểm)</p>";
            exam.part2.forEach(function(q, qIdx) {
                html += "<p style='margin-top: 6px; margin-bottom: 3px;'><strong>Câu " + (qIdx + 1) + ".</strong> [" + q.level + (q.indicator ? " - " + q.indicator : "") + "] " + cleanLatexForWord(q.content) + "</p>" +
                    "<div style='margin-left: 10px;'>" +
                    (q.statements || []).map(function(s, sIdx) {
                        var meta = getStatementMeta(q, sIdx);
                        var metaTag = " <span style='font-size: 9pt; color: #1d4ed8; font-weight: bold;'>[" + meta.level + " - " + meta.indicator + "]</span>";
                        var ansBadge = (audience === "TEACHER_KEY") ? (" <strong>[" + (s.correct ? "ĐÚNG" : "SAI") + "]</strong>") : "";
                        return "<p style='margin: 2px 0;'>" + cleanLatexForWord(s.text) + metaTag + ansBadge + "</p>";
                    }).join("") +
                    "</div>";
                if (audience === "TEACHER_KEY") {
                    var expP2w = formatDetailedExplanation(q.explain, q);
                    html += "<p style='margin: 2px 0 5px 10px; font-size: 10pt; color: #1d4ed8; font-style: italic;'><strong>Hướng dẫn giải:</strong> " + cleanLatexForWord(expP2w) + "</p>";
                }
            });
        }

        // P3
        if (exam.part3.length > 0) {
            html += "<p class='bold' style='margin-top: 10px; color: #7e22ce;'>PHẦN III: CÂU TRẮC NGHIỆM TRẢ LỜI NGẮN (4 câu - 2,00 điểm)</p>";
            exam.part3.forEach(function(q, qIdx) {
                var expP3w = formatDetailedExplanation(q.explain, q);
                html += "<p style='margin-top: 6px; margin-bottom: 3px;'><strong>Câu " + (qIdx + 1) + ".</strong> [" + q.level + (q.indicator ? " - " + q.indicator : "") + "] " + cleanLatexForWord(q.content) + "</p>";
                if (audience === "TEACHER_KEY") {
                    html += "<p style='margin: 2px 0 5px 10px; font-size: 10pt; color: #7e22ce; font-style: italic;'><strong>Đáp án:</strong> " + cleanLatexForWord(q.answer) + (expP3w ? (" — " + cleanLatexForWord(expP3w)) : "") + "</p>";
                } else {
                    html += "<p style='margin: 2px 0 5px 10px; font-size: 10pt;'>Đáp số: .....................................................................</p>";
                }
            });
        }

        // P4
        if (exam.part4.length > 0) {
            html += "<p class='bold' style='margin-top: 10px; color: #be123c;'>PHẦN IV: CÂU HỎI & BÀI TẬP TỰ LUẬN (3 câu - 3,00 điểm)</p>";
            exam.part4.forEach(function(q, qIdx) {
                html += "<div style='margin-top: 6px; margin-bottom: 6px; padding-left: 5px; border-left: 2px solid #be123c;'>" +
                    "<p><strong>Câu " + (qIdx + 1) + ".</strong> [" + q.level + (q.indicator ? " - " + q.indicator : "") + "]: " + cleanLatexForWord(q.content).replace(/\n/g, "<br/>") + "</p></div>";
                if (audience === "TEACHER_KEY") {
                    html += "<div style='margin: 3px 0 8px 10px; background-color: #fff1f2; padding: 4px 8px; border: 1px solid #fecdd3; font-size: 10pt;'>" +
                        "<strong>Hướng dẫn chấm & Lời giải:</strong><br/>" + formatSolutionHtml(q.solution || q.explain || "") +
                        "</div>";
                }
            });
        }
    });

    // PHẦN C (Nếu STUDENT_ANS_END)
    if (audience === "STUDENT_ANS_END") {
        html += "<div class='page-break'></div>";
        html += "<h3 style='background-color: #f1f5f9; padding: 4px 8px; border-left: 4px solid #475569; font-size: 12pt;' class='bold'>PHẦN C: BẢNG ĐÁP ÁN & HƯỚNG DẪN GIẢI CHI TIẾT TỪNG ĐỀ ÔN TẬP</h3>";

        examData.exams.forEach(function(exam, eIdx) {
            html += "<h4 class='bold' style='margin-top: 10px; color: #0f766e;'>BẢNG ĐÁP ÁN — " + cleanLatexForWord(exam.name) + "</h4>";

            // Bảng P1
            html += "<p class='bold' style='margin: 4px 0 2px 0;'>1. Đáp án Phần I (Trắc nghiệm nhiều lựa chọn):</p>" +
                "<table class='data-table answer-table' style='width: 100%; table-layout: fixed; mso-table-layout-alt: fixed; border-collapse: collapse;'>" +
                "<colgroup>" +
                "<col style='width: 12.5%;' width='12.5%'/>" +
                "<col style='width: 12.5%;' width='12.5%'/>" +
                "<col style='width: 12.5%;' width='12.5%'/>" +
                "<col style='width: 12.5%;' width='12.5%'/>" +
                "<col style='width: 12.5%;' width='12.5%'/>" +
                "<col style='width: 12.5%;' width='12.5%'/>" +
                "<col style='width: 12.5%;' width='12.5%'/>" +
                "<col style='width: 12.5%;' width='12.5%'/>" +
                "</colgroup>" +
                "<thead><tr><th>Câu</th><th>Đáp Án</th><th>Câu</th><th>Đáp Án</th><th>Câu</th><th>Đáp Án</th><th>Câu</th><th>Đáp Án</th></tr></thead><tbody><tr>";
            for (var i = 0; i < exam.part1.length; i++) {
                if (i > 0 && i % 4 === 0) html += "</tr><tr>";
                html += "<td class='bold text-center'>" + (i + 1) + "</td><td class='bold text-center' style='color: #047857; font-weight: bold;'>" + String.fromCharCode(65 + exam.part1[i].correct) + "</td>";
            }
            html += "</tr></tbody></table>";

            // Bảng P2
            html += "<p class='bold' style='margin: 6px 0 2px 0;'>2. Đáp án Phần II (Đúng / Sai 4 ý):</p>" +
                "<table class='data-table answer-table' style='width: 100%; table-layout: fixed; mso-table-layout-alt: fixed; border-collapse: collapse;'>" +
                "<colgroup><col style='width: 15%;' width='15%'/><col style='width: 85%;' width='85%'/></colgroup>" +
                "<thead><tr><th style='width: 15%; text-align: center;'>Câu</th><th style='width: 85%; text-align: center;'>Đáp án 4 ý (a - b - c - d)</th></tr></thead><tbody>";
            exam.part2.forEach(function(q, qIdx) {
                var stStr = (q.statements || []).map(function(s, sIdx) {
                    var meta = getStatementMeta(q, sIdx);
                    return String.fromCharCode(97 + sIdx) + ") " + (s.correct ? "Đ" : "S") + " [" + meta.level + " - " + meta.indicator + "]";
                }).join("  •  ");
                html += "<tr><td class='bold text-center' style='width: 15%;'>Câu " + (qIdx + 1) + "</td><td style='width: 85%; word-wrap: break-word;'>" + stStr + "</td></tr>";
            });
            html += "</tbody></table>";

            // Bảng P3
            html += "<p class='bold' style='margin: 6px 0 2px 0;'>3. Đáp án Phần III (Trả lời ngắn):</p>" +
                "<table class='data-table answer-table' style='width: 100%; table-layout: fixed; mso-table-layout-alt: fixed; border-collapse: collapse;'>" +
                "<colgroup><col style='width: 15%;' width='15%'/><col style='width: 85%;' width='85%'/></colgroup>" +
                "<thead><tr><th style='width: 15%; text-align: center;'>Câu</th><th style='width: 85%; text-align: center;'>Đáp số chuẩn</th></tr></thead><tbody>";
            exam.part3.forEach(function(q, qIdx) {
                html += "<tr><td class='bold text-center' style='width: 15%;'>Câu " + (qIdx + 1) + "</td><td class='bold' style='width: 85%; color: #7e22ce; word-wrap: break-word;'>" + cleanLatexForWord(q.answer) + "</td></tr>";
            });
            html += "</tbody></table>";

            // P4
            if (exam.part4.length > 0) {
                html += "<p class='bold' style='margin: 6px 0 2px 0;'>4. Lời giải vắn tắt Phần IV (Tự luận):</p>";
                exam.part4.forEach(function(q, qIdx) {
                    html += "<div style='margin: 4px 0 8px 5px;'><strong>Câu " + (qIdx + 1) + " (1,0đ):</strong> " + formatSolutionHtml(q.solution || q.explain || "") + "</div>";
                });
            }

            html += "<hr style='border: none; border-top: 1px dashed #cbd5e1; margin: 12px 0;'/>";
        });
    }

    var fileName = "De_Cuong_On_Tap_Hoa_" + grade + "_" + term + "_KNTT.doc";
    downloadWordDocument(html, fileName);
}

// Tự động khởi chạy sau khi DOM và các script đã sẵn sàng
window.addEventListener("DOMContentLoaded", function() {
    setTimeout(function() {
        initOfficialBTXSystem();
        if (typeof initSyllabusLessonRangePickers === "function") {
            initSyllabusLessonRangePickers(true);
        }
        if (typeof renderSyllabus === "function") {
            renderSyllabus();
        }
    }, 50);
});

/* Helper toàn cục: Sinh đề và chuyển xem ngay */
function generateExamFromCurrentMatrix() {
    if (typeof generateExamFromMatrix === "function") {
        generateExamFromMatrix();
    }
    if (typeof switchTab === "function") {
        switchTab("generator");
    }
    if (typeof switchTab4View === "function") {
        switchTab4View("exam");
    }
}

/* Helper toàn cục: Thu gọn / Mở rộng Bản Đặc Tả */
function toggleSpecSheetView() {
    var container = document.getElementById("specification-table-container");
    var chevron = document.getElementById("spec-chevron");
    var text = document.getElementById("spec-toggle-text");
    if (!container) return;
    if (container.classList.contains("hidden")) {
        container.classList.remove("hidden");
        if (chevron) chevron.className = "fa-solid fa-chevron-up";
        if (text) text.innerText = "Thu Gọn Bản Đặc Tả";
    } else {
        container.classList.add("hidden");
        if (chevron) chevron.className = "fa-solid fa-chevron-down";
        if (text) text.innerText = "Mở Rộng Bản Đặc Tả";
    }
}


