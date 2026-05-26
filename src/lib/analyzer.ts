import * as XLSX from "xlsx";
import { AnalysisResult, Conflict, ProcessedProcedure } from "../types";
import { isBefore, isEqual, isValid, differenceInMinutes, addMinutes } from "date-fns";
import { formatDateTime } from "./utils";

const standardizeString = (str: string) => {
  if (!str) return "";
  return String(str)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim();
};

const isDienCham = (name: string) => name.trim().toLowerCase() === "điện châm [kim ngắn]";
const isHongNgoai = (name: string) => name.trim().toLowerCase() === "điều trị bằng tia hồng ngoại";
const isXoaBop = (name: string) => {
  return name.toLowerCase().includes("xoa bóp") || standardizeString(name).includes("xoa bop");
};
const isSacThuoc = (name: string) => name.trim().toLowerCase() === "sắc thuốc thang và đóng gói thuốc bằng máy";
const isParafin = (name: string) => name.trim().toLowerCase() === "điều trị bằng parafin";
const isThuyCham = (name: string) => name.toLowerCase().includes("thủy châm") || name.toLowerCase().includes("thuỷ châm");
const isKeoNan = (name: string) => name.toLowerCase().includes("kéo nắn") || standardizeString(name).includes("keo nan");
const isCayChi = (name: string) => name.toLowerCase().includes("cấy chỉ") || standardizeString(name).includes("cay chi");

const checkSpecialOverlapRule = (pA: ProcessedProcedure, pB: ProcessedProcedure) => {
  const pDienCham = isDienCham(pA.procedureName) ? pA : (isDienCham(pB.procedureName) ? pB : null);
  const pHongNgoai = isHongNgoai(pA.procedureName) ? pA : (isHongNgoai(pB.procedureName) ? pB : null);

  if (pDienCham && pHongNgoai) {
    if (isBefore(pHongNgoai.startTime, pDienCham.startTime)) {
      return differenceInMinutes(pDienCham.startTime, pHongNgoai.startTime) >= 1;
    } else {
      return differenceInMinutes(pHongNgoai.startTime, pDienCham.startTime) >= 5;
    }
  }
  return false;
};

const getDoctorBusyEndTime = (p: ProcessedProcedure) => {
  if (isDienCham(p.procedureName)) {
    return addMinutes(p.startTime, 5);
  }
  if (isHongNgoai(p.procedureName)) {
    return addMinutes(p.startTime, 1);
  }
  if (isThuyCham(p.procedureName)) {
    return addMinutes(p.startTime, 10);
  }
  // isXoaBop logic does not have specific busy endTime in the specs explicitly beyond its duration
  return p.endTime;
};

const readExcelFile = (file: File): Promise<any[][]> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: "array" });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        resolve(XLSX.utils.sheet_to_json(worksheet, { header: 1 }));
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(new Error("Lỗi hệ thống khi đọc file."));
    reader.readAsArrayBuffer(file);
  });
};

// Convert excel date to valid JS Date if needed
function parseExcelDate(excelDate: any): Date | null {
  if (excelDate instanceof Date) {
    if (!isValid(excelDate)) return null;
    return new Date(Math.round(excelDate.getTime() / 60000) * 60000);
  }
  if (typeof excelDate === 'number') {
    const parsed = XLSX.SSF.parse_date_code(excelDate);
    if (parsed) {
      return new Date(parsed.y, parsed.m - 1, parsed.d, parsed.H, parsed.M, 0);
    }
  }
  if (typeof excelDate === 'string') {
    // Try to parse typical manual patterns if necessary
    const timeMatch = excelDate.match(/(\d{1,2}):(\d{1,2})/);
    if (timeMatch && excelDate.includes('/')) {
        const dateMatch = excelDate.match(/(\d{1,2})[\/\-\:](\d{1,2})[\/\-\:](\d{4})/);
        if (dateMatch) {
            return new Date(Number(dateMatch[3]), Number(dateMatch[2]) - 1, Number(dateMatch[1]), Number(timeMatch[1]), Number(timeMatch[2]), 0);
        }
    }
    const d = new Date(excelDate);
    if (isValid(d)) return d;
  }
  return null;
}

export const analyzeExcelData = async (mainFile: File, clsFile?: File): Promise<AnalysisResult> => {
  try {
    const mainMatrix = await readExcelFile(mainFile);
    let clsMatrix: any[][] | null = null;
    
    if (clsFile) {
      try {
        clsMatrix = await readExcelFile(clsFile);
      } catch (e) {
        throw new Error("Lỗi khi đọc file Dịch vụ CLS.");
      }
    }

    let colMapping = {
      patientId: -1,
      patientName: -1,
      procedureName: -1,
      startTime: -1,
      endTime: -1,
      time: -1,
      machineId: -1,
      doctorName: -1
    };

    let headerRowIndex = -1;

    // First pass: try to find the full header row
    for (let r = 0; r < Math.min(mainMatrix.length, 100); r++) {
      const row = mainMatrix[r] || [];
      const rowStr = row.map(v => standardizeString(String(v || ""))).join('|');
      
      const hasPatient = rowStr.includes("ma ba") || rowStr.includes("ma bn") || rowStr.includes("ma benh an") || rowStr.includes("ten benh") || rowStr.includes("ho ten");
      const hasProcedure = rowStr.includes("thu thuat") || rowStr.includes("dich vu");
      
      if (hasPatient && hasProcedure) {
        headerRowIndex = r;
        
        // Scan this row and the next two rows to catch all column headers (including merged cells)
        for (let i = 0; i <= 2; i++) {
          const scanRow = mainMatrix[r + i] || [];
          for (let c = 0; c < scanRow.length; c++) {
            const val = standardizeString(String(scanRow[c] || ""));
            if (!val) continue;

            if ((val === "ma ba" || val === "ma benh an" || val === "ma bn" || val.includes("ma benh an")) && colMapping.patientId === -1) colMapping.patientId = c;
            
            if ((val === "ho ten nguoi benh" || val === "ten benh nhan" || val === "ho ten") && colMapping.patientName === -1) colMapping.patientName = c;
            else if (val.includes("ten benh nhan") && colMapping.patientName === -1) colMapping.patientName = c;
            
            if ((val === "phuong phap thu thuat" || val === "thu thuat" || val === "ten dich vu bhyt" || val === "ten dich vu") && colMapping.procedureName === -1) colMapping.procedureName = c;
            else if (val.includes("thu thuat") && colMapping.procedureName === -1 && !val.includes("loai")) colMapping.procedureName = c;
            
            if ((val === "ngay gio thuc hien" || val === "thoi gian bat dau" || val === "ngay gio bat dau" || val.includes("ngay gio thuc hien")) && colMapping.startTime === -1) colMapping.startTime = c;
            
            if ((val === "ngay gio ket thuc" || val === "thoi gian ket thuc" || val.includes("ngay gio ket thuc")) && colMapping.endTime === -1) colMapping.endTime = c;
            
            if ((val === "thoi gian thu thuat" || val === "thoi gian phau thuat" || val === "thoi gian thuc hien") && colMapping.time === -1) colMapping.time = c;
            else if (val.includes("thoi gian phau thuat") && colMapping.time === -1) colMapping.time = c;
            
            if ((val === "ma may" || val.includes("ma may")) && colMapping.machineId === -1) colMapping.machineId = c;
            
            if ((val === "nguoi thuc hien" || val === "ptv tt v" || val === "ptv tt" || val === "ptv" || val === "thuc hien chinh" || val === "phau thuat vien 1") && colMapping.doctorName === -1) colMapping.doctorName = c;
            else if (val.includes("phau thuat vien") && colMapping.doctorName === -1) colMapping.doctorName = c;
          }
        }
        break;
      }
    }
    
    // Fallback defaults if they couldn't be detected (based on legacy file format norms)
    if (colMapping.patientId === -1) colMapping.patientId = 2;
    if (colMapping.patientName === -1) colMapping.patientName = 4;
    if (colMapping.procedureName === -1) colMapping.procedureName = 10;
    if (colMapping.time === -1 && colMapping.startTime === -1) colMapping.time = 12;
    if (colMapping.machineId === -1) colMapping.machineId = 13;
    if (colMapping.doctorName === -1) colMapping.doctorName = 14;
    
    const procedures: ProcessedProcedure[] = [];
    const errors: string[] = [];
    let currentDateStr = "";

    // Data usually starts immediately after the header block
    for (let r = (headerRowIndex !== -1 ? headerRowIndex + 1 : 0); r < mainMatrix.length; r++) {
      const row = mainMatrix[r];
      if (!row || !row.length) continue;
      
      const nonEmptiesRaw = row.filter(v => v !== undefined && v !== null && String(v).trim() !== '');
      if (nonEmptiesRaw.length === 0) continue;

      const patientId = String(row[colMapping.patientId] || '').trim();
      const patientIdLower = standardizeString(patientId);
      
      // Skip if it's a repeated header row or mostly empty
      const isHeaderRow = nonEmptiesRaw.some(c => {
         const std = standardizeString(String(c));
         return std === "stt" || std === "ma ba" || std.includes("ma benh an") || std === "ho ten nguoi benh" || std.includes("ten benh nhan") || std.includes("thoi gian thu thuat");
      });
      if (isHeaderRow) continue;

      // Skip numeric row (e.g. 1, 2, 3 below headers)
      const isNumericRow = nonEmptiesRaw.length >= 8 && nonEmptiesRaw.every(c => typeof c === 'number' || !isNaN(Number(c)) || String(c).trim() === '');
      if (isNumericRow) continue;

      // Handle lone date rows just in case (legacy format)
      if (!patientId || nonEmptiesRaw.length <= 5) {
          const dateObj = nonEmptiesRaw.find(v => v instanceof Date && isValid(v));
          const dateNum = nonEmptiesRaw.find(v => typeof v === 'number' && v > 40000 && v < 50000);
          const dateMatch = nonEmptiesRaw.map(String).find(v => /\d{1,2}[\/\-\:]\d{1,2}[\/\-\:]\d{4}/.test(v));
          
          if (dateObj instanceof Date) {
            const roundedDateObj = new Date(Math.round(dateObj.getTime() / 60000) * 60000);
            currentDateStr = `${roundedDateObj.getDate().toString().padStart(2, '0')}/${(roundedDateObj.getMonth() + 1).toString().padStart(2, '0')}/${roundedDateObj.getFullYear()}`;
          } else if (typeof dateNum === 'number') {
            const parsed = XLSX.SSF.parse_date_code(dateNum);
            if (parsed) {
               currentDateStr = `${parsed.d.toString().padStart(2, '0')}/${parsed.m.toString().padStart(2, '0')}/${parsed.y}`;
            }
          } else if (dateMatch) {
            const m = dateMatch.match(/\d{1,2}[\/\-\:]\d{1,2}[\/\-\:]\d{4}/);
            if (m) {
               currentDateStr = m[0];
            }
          }
          if (!patientId) continue;
      }
      
      const procedureName = String(row[colMapping.procedureName] || '').trim();
      let startTimeValue: Date | null = null;
      let endTimeValue: Date | null = null;

      // Try new format first (separate startTime and endTime columns)
      if (colMapping.startTime !== -1 && colMapping.endTime !== -1) {
         startTimeValue = parseExcelDate(row[colMapping.startTime]);
         endTimeValue = parseExcelDate(row[colMapping.endTime]);
      } 
      
      // Fallback to legacy method (time range in one column)
      if ((!startTimeValue || !endTimeValue) && colMapping.time !== -1) {
         const timeStr = String(row[colMapping.time] || '').trim();
         const timeRegex = /\d{1,2}:\d{1,2}(?::\d{1,2})?/g;
         const timeMatches = timeStr.match(timeRegex);
         
         if (timeMatches && timeMatches.length >= 2) {
             let day = 1, month = 0, year = 1970;
             if (currentDateStr) {
                const dateParts = currentDateStr.split(/[\/\-:]/);
                if (dateParts.length >= 3) {
                   day = parseInt(dateParts[0], 10);
                   month = parseInt(dateParts[1], 10) - 1;
                   year = parseInt(dateParts[2], 10);
                }
             } else {
                const today = new Date();
                day = today.getDate();
                month = today.getMonth();
                year = today.getFullYear();
             }
             
             const [sH, sM, sS] = timeMatches[0].split(':').map(Number);
             startTimeValue = new Date(year, month, day, sH, sM, sS || 0);
             
             const [eH, eM, eS] = timeMatches[1].split(':').map(Number);
             endTimeValue = new Date(year, month, day, eH, eM, eS || 0);
         }
      }

      if (patientId && procedureName && startTimeValue && endTimeValue && isValid(startTimeValue) && isValid(endTimeValue)) {
         const patientName = String(row[colMapping.patientName] || '').trim();
         const doctorName = String(row[colMapping.doctorName] || '').trim();
         let machineId = String(row[colMapping.machineId] || '').trim();
         
         if (machineId === "-1" || machineId.toLowerCase() === "không") {
           machineId = "";
         }

         procedures.push({
            id: `proc_${r}`,
            patientId,
            patientName,
            procedureName,
            doctorName: doctorName || "Không rõ",
            machineId,
            startTime: startTimeValue,
            endTime: endTimeValue,
            rawRow: row,
            rowNumber: r + 1
         });
      }
    }

    if (procedures.length === 0) {
      throw new Error("Không tìm thấy dữ liệu thủ thuật nào trong file chính. Vui lòng kiểm tra lại.");
    }

    // --- Process CLS ---
    if (clsMatrix) {
      let clsColMap = {
        patientId: 3, // Mã BN hoặc Mã BA
        patientName: 5,
        procedureName: 9,
        startTime: 13, // TG tiếp nhận
        endTime: 15 // TG trả KQ
      };

      for (let r = 0; r < Math.min(clsMatrix.length, 10); r++) {
         const row = clsMatrix[r] || [];
         const rowStr = row.map(v => standardizeString(String(v || ""))).join('|');
         if (rowStr.includes("ma bn") && rowStr.includes("tg tiep nhan") && rowStr.includes("tg tra kq")) {
            for (let c = 0; c < row.length; c++) {
              const val = standardizeString(String(row[c] || ""));
              if (!val) continue;
              if (val === "ma ba" || val === "ma bn") clsColMap.patientId = c;
              if (val === "ten bn" || val === "ho ten") clsColMap.patientName = c;
              if (val === "ten dich vu") clsColMap.procedureName = c;
              if (val === "tg tiep nhan") clsColMap.startTime = c;
              if (val === "tg tra kq") clsColMap.endTime = c;
            }
            break;
         }
      }

      for (let r = 1; r < clsMatrix.length; r++) {
         const row = clsMatrix[r];
         if (!row || !row.length) continue;
         const patientId = String(row[clsColMap.patientId] || "").trim();
         if (!patientId || patientId.toLowerCase() === "mã ba" || patientId.toLowerCase() === "mã bn") continue;
         
         let startT = parseExcelDate(row[clsColMap.startTime]);
         let endT = parseExcelDate(row[clsColMap.endTime]);

         if (!startT || !endT) continue;
         
         procedures.push({
            id: `cls_${r}`,
            patientId,
            patientName: String(row[clsColMap.patientName] || "").trim(),
            procedureName: "CLS: " + String(row[clsColMap.procedureName] || "").trim(),
            doctorName: "KTV CLS",
            machineId: "",
            startTime: startT,
            endTime: endT,
            rawRow: row,
            rowNumber: r + 1
         });
      }
    }

    const conflicts: Conflict[] = [];
    
    for (let i = 0; i < procedures.length; i++) {
      const p1 = procedures[i];
      const isCls1 = p1.id.startsWith("cls_");

      if (!isValid(p1.startTime) || !isValid(p1.endTime)) {
        if (!isCls1) {
          conflicts.push({
            id: `err_time_${p1.id}`,
            type: 'INVALID_TIME',
            message: `Thủ thuật "${p1.procedureName}" của bệnh nhân ${p1.patientName} có thời gian không hợp lệ.`,
            procedures: [p1]
          });
        }
        continue;
      }

      if (!isBefore(p1.startTime, p1.endTime) && !isEqual(p1.startTime, p1.endTime)) {
        if (!isCls1) {
          conflicts.push({
            id: `err_time_order_${p1.id}`,
            type: 'INVALID_TIME',
            message: `Thủ thuật "${p1.procedureName}" của bệnh nhân ${p1.patientName} có thời gian kết thúc trước thời gian bắt đầu.`,
            procedures: [p1]
          });
        }
      }

      if (!isCls1) {
        const duration = differenceInMinutes(p1.endTime, p1.startTime);
        if (isDienCham(p1.procedureName) && duration < 25) {
          conflicts.push({
            id: `err_duration_${p1.id}`,
            type: 'INVALID_TIME',
            message: `Thủ thuật "${p1.procedureName}" của bệnh nhân **${p1.patientName}** có thời gian thực hiện (${duration} phút) thấp hơn mức tối thiểu 25 phút.`,
            procedures: [p1]
          });
        } else if (isHongNgoai(p1.procedureName) && duration < 15) {
          conflicts.push({
            id: `err_duration_${p1.id}`,
            type: 'INVALID_TIME',
            message: `Thủ thuật "${p1.procedureName}" của bệnh nhân **${p1.patientName}** có thời gian thực hiện (${duration} phút) thấp hơn mức tối thiểu 15 phút.`,
            procedures: [p1]
          });
        } else if (isXoaBop(p1.procedureName) && duration < 20) {
          conflicts.push({
            id: `err_duration_${p1.id}`,
            type: 'INVALID_TIME',
            message: `Thủ thuật "${p1.procedureName}" của bệnh nhân **${p1.patientName}** có thời gian thực hiện (${duration} phút) thấp hơn mức tối thiểu 20 phút.`,
            procedures: [p1]
          });
        } else if (isKeoNan(p1.procedureName) && (duration < 10 || duration > 20)) {
          conflicts.push({
            id: `err_duration_${p1.id}`,
            type: 'INVALID_TIME',
            message: `Thủ thuật "${p1.procedureName}" của bệnh nhân **${p1.patientName}** có thời gian thực hiện (${duration} phút) nằm ngoài khoảng cho phép 10-20 phút.`,
            procedures: [p1]
          });
        } else if (isCayChi(p1.procedureName) && duration < 15) {
          conflicts.push({
            id: `err_duration_${p1.id}`,
            type: 'INVALID_TIME',
            message: `Thủ thuật "${p1.procedureName}" của bệnh nhân **${p1.patientName}** có thời gian thực hiện (${duration} phút) thấp hơn mức tối thiểu 15 phút.`,
            procedures: [p1]
          });
        }

        const p1BusyEnd = getDoctorBusyEndTime(p1);

        const isMachineRequired = isDienCham(p1.procedureName) || isHongNgoai(p1.procedureName) || isSacThuoc(p1.procedureName);
        if (isMachineRequired && !p1.machineId) {
          conflicts.push({
            id: `err_missing_machine_${p1.id}`,
            type: 'MISSING_MACHINE_ID',
            message: `Thủ thuật "${p1.procedureName}" của bệnh nhân **${p1.patientName}** bắt buộc phải có mã máy nhưng đang để trống.`,
            procedures: [p1]
          });
        }
      }

      for (let j = i + 1; j < procedures.length; j++) {
        const p2 = procedures[j];
        if (!isValid(p2.startTime) || !isValid(p2.endTime)) continue;
        
        const isCls2 = p2.id.startsWith("cls_");

        if (p1.machineId && p2.machineId && p1.machineId === p2.machineId && !isCls1 && !isCls2) {
          const checkMachineOverlap = isBefore(p1.startTime, p2.endTime) && isBefore(p2.startTime, p1.endTime);
          if (checkMachineOverlap) {
            conflicts.push({
              id: `conflict_machine_${p1.id}_${p2.id}`,
              type: 'MACHINE_TIME_OVERLAP',
              message: `Mã máy **${p1.machineId}** đang được sử dụng trùng lịch giữa thủ thuật "${p1.procedureName}" (BN: ${p1.patientName}) và "${p2.procedureName}" (BN: ${p2.patientName}).`,
              procedures: [p1, p2]
            });
          }
        }

        const diffMs = Math.abs(p1.startTime.getTime() - p2.startTime.getTime());
        const isP1SacThuoc = !isCls1 && isSacThuoc(p1.procedureName);
        const isP2SacThuoc = !isCls2 && isSacThuoc(p2.procedureName);

        if (isP1SacThuoc || isP2SacThuoc) {
          if (isP1SacThuoc !== isP2SacThuoc) {
            continue;
          }
          if (diffMs >= 59000) {
            continue;
          }
        }

        const isPatientOverlap = isBefore(p1.startTime, p2.endTime) && isBefore(p2.startTime, p1.endTime);
        
        // For CLS, we only care about Patient overlaps with other CLS or other main procedures
        if (isPatientOverlap && p1.patientId === p2.patientId) {
          // Both are CLS or mixed CLS + Proc
          if (isCls1 || isCls2) {
             conflicts.push({
               id: `conflict_cls_${p1.id}_${p2.id}`,
               type: 'CLS_TIME_OVERLAP',
               message: `Bệnh nhân **${p1.patientName}** (Mã: ${p1.patientId}) đang có dịch vụ ${isCls1 ? "CLS" : "thủ thuật"} "${p1.procedureName}" (từ ${formatDateTime(p1.startTime)} đến ${formatDateTime(p1.endTime)}) trùng thời gian với ${isCls2 ? "CLS" : "thủ thuật"} "${p2.procedureName}" (từ ${formatDateTime(p2.startTime)} đến ${formatDateTime(p2.endTime)}).`,
               procedures: [p1, p2]
             });
          } else {
            // Both are normal procedures
            const isSpecialRuleOk = checkSpecialOverlapRule(p1, p2);

            if (!isSpecialRuleOk) {
              conflicts.push({
                id: `conflict_pt_${p1.id}_${p2.id}`,
                type: 'PATIENT_TIME_OVERLAP',
                message: `Bệnh nhân **${p1.patientName}** (Mã: ${p1.patientId}) đang thực hiện thủ thuật "${p1.procedureName}" (từ ${formatDateTime(p1.startTime)} đến ${formatDateTime(p1.endTime)}) trùng thời gian với thủ thuật "${p2.procedureName}" (từ ${formatDateTime(p2.startTime)} đến ${formatDateTime(p2.endTime)}).`,
                procedures: [p1, p2]
              });
            }
          }
        }

        if (!isCls1 && !isCls2) {
          const p1BusyEnd = getDoctorBusyEndTime(p1);
          const p2BusyEnd = getDoctorBusyEndTime(p2);
          let isDoctorOverlap = isBefore(p1.startTime, p2BusyEnd) && isBefore(p2.startTime, p1BusyEnd);

          const isP1Parafin = isParafin(p1.procedureName);
          const isP2Parafin = isParafin(p2.procedureName);
          if (isP1Parafin && isP2Parafin && p1.patientId !== p2.patientId) {
            if (diffMs >= 59000) {
              isDoctorOverlap = false;
            } else {
              isDoctorOverlap = true;
            }
          }

          if (isDoctorOverlap && p1.doctorName === p2.doctorName && p1.doctorName && p1.doctorName !== "Không rõ" && p1.id !== p2.id) {
             conflicts.push({
              id: `conflict_doc_${p1.id}_${p2.id}`,
              type: 'DOCTOR_TIME_OVERLAP',
              message: `Nhân viên **${p1.doctorName}** đang thực hiện thủ thuật cho ${p1.patientName} (từ ${formatDateTime(p1.startTime)} đến ${formatDateTime(p1.endTime)}) trùng thời gian với thủ thuật của bệnh nhân ${p2.patientName} (từ ${formatDateTime(p2.startTime)} đến ${formatDateTime(p2.endTime)}).`,
              procedures: [p1, p2]
            });
          }
        }
      }
    }

    // Include CLS patients? Only if they have actual conflicts or procedures?
    // Let's keep uniquePatients and uniqueDoctors based on all procedures
    const uniquePatients = new Set(procedures.map(p => p.patientId));
    const uniqueDoctors = new Set(procedures.map(p => p.doctorName).filter(d => d !== "Không rõ" && !d.startsWith("KTV")));

    return {
      procedures,
      conflicts,
      totalPatients: uniquePatients.size,
      totalDoctors: uniqueDoctors.size,
      errors
    };

  } catch (err) {
    if (err instanceof Error) {
      throw err;
    }
    throw new Error("Lỗi khi đọc file Excel. Định dạng file có thể không đúng.");
  }
};
