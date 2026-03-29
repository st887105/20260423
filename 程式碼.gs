/**
 * ============================================================================
 * 車城國小學力檢測考古題輔助系統 - 後端邏輯
 * v5.0 - 多科目 + PDF 教材出題版
 * ============================================================================
 * 新增功能：
 *   1. 支援數學科 / 國語科（科目參數化）
 *   2. 支援從 Google Drive PDF 讀取內容並讓 Gemini 依內容出題
 *   3. Bank 試算表新增「科目」欄位（第 9 欄）
 * ============================================================================
 */

// ==========================================
// API 入口 (供 GitHub Pages 前端呼叫)
// ==========================================
function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const action = params.action;
    const data   = params.data || {};
    let result   = null;

    switch (action) {
      case 'getTasks':                  result = getTasks(); break;
      case 'getStudents':               result = getStudents(data.taskName); break;
      case 'verifyAdmin':               result = verifyAdmin(data.pwd); break;
      case 'getQuizSettings':           result = getQuizSettings(); break;
      case 'updateQuizSettings':        result = updateQuizSettings(data.count); break;
      case 'updateFolderId':            result = updateFolderId(data.folderId); break;
      case 'getFolderId':               result = getFolderId(); break;
      case 'setupDatabase':             result = setupDatabase(); break;
      case 'clearBankCache':            result = clearBankCache(); break;
      case 'uploadTaskData':            result = uploadTaskData(data.taskName, data.grade, data.subject, data.studentData, data.uniqueNodes); break;
      case 'generateBatch':             result = generateBatch(data.nodesArray, data.grade, data.subject, data.batchIndex); break;
      case 'generateBatchFromFolder':   result = generateBatchFromFolder(data.folderId, data.nodesArray, data.grade, data.subject, data.batchIndex); break;
      case 'scanFolder':                result = scanFolder(data.folderId); break;
      case 'generateQuiz':              result = generateQuiz(data.weakNode, data.taskName); break;
      case 'submitQuizResult':          result = submitQuizResult(data.payload); break;
      case 'getTaskResults':            result = getTaskResults(data.taskName); break;
      case 'getQuestionErrorRates':     result = getQuestionErrorRates(data.taskName); break;
      case 'analyzeClassWeakNodes':     result = analyzeClassWeakNodes(data.taskName); break;
      case 'generateTeachingWorksheet': result = generateTeachingWorksheet(data.topNodes, data.grade, data.subject); break;
      default: throw new Error('未知的 API 請求：' + action);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success', data: result }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: error.message || String(error) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput(
    "✅ 車城國小 AI 補救系統 API 伺服器運作中（v5.0 多科目+PDF版）\n請從 GitHub Pages 前端網頁連線。"
  );
}


// ==========================================
// 1. 資料庫初始化（Bank 新增科目欄）
// ==========================================
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = [];

  // ── Config ──
  let configSheet = ss.getSheetByName('Config');
  if (!configSheet) {
    configSheet = ss.insertSheet('Config');
    configSheet.appendRow(['Key', 'Value']);
    configSheet.appendRow(['AdminPassword', '1234']);
    configSheet.appendRow(['GeminiAPIKey', '請在此貼上您的API金鑰']);
    configSheet.appendRow(['QuizCount', '10']);
    configSheet.appendRow(['PdfFolderId', '（選填）貼上 Google Drive 資料夾 ID']);
    log.push('Config 分頁已建立');
  } else {
    // 補齊缺少的 Key
    const configData = configSheet.getDataRange().getValues();
    const existKeys = configData.map(r => String(r[0]));
    if (!existKeys.includes('PdfFolderId')) {
      configSheet.appendRow(['PdfFolderId', '（選填）貼上 Google Drive 資料夾 ID']);
      log.push('Config 已補充 PdfFolderId');
    }
  }

  // ── Bank ──
  let bankSheet = ss.getSheetByName('Bank');
  if (!bankSheet) {
    bankSheet = ss.insertSheet('Bank');
    bankSheet.appendRow(['ID', '知識節點', '題目', '類型(single/fill)', '選項(JSON陣列)', '正解', '難度', '適用年級', '科目']);
    log.push('Bank 分頁已建立（含科目欄）');
  } else {
    // 補齊第 9 欄「科目」
    const header = bankSheet.getRange(1, 1, 1, bankSheet.getLastColumn()).getValues()[0];
    if (!header.includes('科目')) {
      bankSheet.getRange(1, header.length + 1).setValue('科目');
      log.push('Bank 已補充「科目」欄位');
    } else {
      log.push('Bank 科目欄已存在 ✅');
    }
  }

  // ── History ──
  let historySheet = ss.getSheetByName('History');
  if (!historySheet) {
    historySheet = ss.insertSheet('History');
    historySheet.appendRow(['上傳時間', '任務名稱', '適用年級', '科目', '學生人數', '班級弱點節點']);
    log.push('History 分頁已建立（含科目欄）');
  } else {
    const header = historySheet.getRange(1, 1, 1, historySheet.getLastColumn()).getValues()[0];
    if (!header.includes('科目')) {
      // 在適用年級後插入科目欄（第4欄）
      historySheet.insertColumnAfter(3);
      historySheet.getRange(1, 4).setValue('科目');
      log.push('History 已在第4欄插入「科目」欄位');
    } else {
      log.push('History 科目欄已存在 ✅');
    }
  }

  // ── Results ──
  let resultsSheet = ss.getSheetByName('Results');
  if (!resultsSheet) {
    resultsSheet = ss.insertSheet('Results');
    resultsSheet.appendRow(['測驗時間', '任務名稱', '座號', '姓名', '分數', '作答歷時(秒)', '作答明細']);
    log.push('Results 分頁已建立');
  } else {
    log.push('Results 分頁已存在 ✅');
  }

  try { CacheService.getScriptCache().remove('BankData_V3'); } catch (e) {}

  return '✅ 資料庫檢查完成！\n' + log.join('\n');
}


function verifyAdmin(pwd) {
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!configSheet) return false;
  const data = configSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'AdminPassword' && data[i][1].toString() === pwd.toString()) return true;
  }
  return false;
}

function getConfig() {
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!configSheet) return {};
  const result = {};
  configSheet.getDataRange().getValues().forEach(row => { if (row[0]) result[row[0]] = row[1]; });
  return result;
}

function getQuizSettings() {
  const cfg = getConfig();
  return { quizCount: parseInt(cfg.QuizCount, 10) || 10 };
}

function updateQuizSettings(newCount) {
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!configSheet) throw new Error("找不到 Config 設定頁");
  const count = parseInt(newCount, 10);
  if (isNaN(count) || count < 1) throw new Error("請輸入有效的數字");
  const data = configSheet.getDataRange().getValues();
  let found = false;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'QuizCount') { configSheet.getRange(i + 1, 2).setValue(count); found = true; break; }
  }
  if (!found) configSheet.appendRow(['QuizCount', count]);
  return "✅ 題數已更新為 " + count + " 題！";
}


// ==========================================
// 1b. 資料夾 ID 管理
// ==========================================
function updateFolderId(folderId) {
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!configSheet) throw new Error('找不到 Config 設定頁');
  const cleanId = String(folderId || '').split('?')[0].trim();
  const data = configSheet.getDataRange().getValues();
  let found = false;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'PdfFolderId') {
      configSheet.getRange(i + 1, 2).setValue(cleanId);
      found = true;
      break;
    }
  }
  if (!found) configSheet.appendRow(['PdfFolderId', cleanId]);
  // 清除 PDF 快取，下次出題重新讀
  try { CacheService.getScriptCache().remove('PdfText_' + cleanId); } catch(e) {}
  return '✅ 資料夾 ID 已儲存：' + cleanId;
}

function getFolderId() {
  const cfg = getConfig();
  const id = String(cfg.PdfFolderId || '').split('?')[0].trim();
  return (id && !id.includes('選填')) ? id : '';
}

// ==========================================
// 2. 任務與學生管理（新增科目欄）
// ==========================================
function uploadTaskData(taskName, grade, subject, studentData, uniqueNodes) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const subjectLabel = subject || '數學';
  const historySheet = ss.getSheetByName('History');
  if (historySheet) {
    historySheet.appendRow([
      new Date(), "'" + String(taskName), grade, subjectLabel,
      studentData.length, uniqueNodes.join(', ')
    ]);
  }
  let taskSheet = ss.getSheetByName(taskName);
  if (taskSheet) taskSheet.clear(); else taskSheet = ss.insertSheet(taskName);
  taskSheet.appendRow(['座號', '姓名', '答對率', '知識節點(弱項)', '科目']);
  const rows = studentData.map(s => {
    let seat = String(s.seatNo || '').trim(), name = String(s.name || '').trim();
    const match = seat.match(/(\d+)\s*[號]?\s*([A-Za-z\u4e00-\u9fa5]+)$/);
    if (match) { seat = match[1]; name = match[2]; }
    return [seat, name, s.accuracy, s.weakNodes, subjectLabel];
  });
  taskSheet.getRange(2, 1, rows.length, 5).setValues(rows);
  return true;
}

function getTasks() {
  const historySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('History');
  if (!historySheet) return [];
  const data = historySheet.getDataRange().getValues();
  // 回傳包含科目資訊的物件
  return data.slice(1).filter(r => r[1]).map(r => ({
    name:    String(r[1]).replace(/'/g, ''),
    grade:   String(r[2] || ''),
    subject: String(r[3] || '數學'),
  })).reverse();
}

function getStudents(taskName) {
  const taskSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(taskName);
  if (!taskSheet) return [];
  return taskSheet.getDataRange().getValues().slice(1).map(r => ({
    seatNo:  String(r[0] || ''),
    name:    String(r[1] || ''),
    weakNode: String(r[3] || ''),
    subject: String(r[4] || '數學'),
  }));
}


// ==========================================
// ==========================================
// 3. PDF 工具函數
// 【v5.1】完全不需要 Drive API 進階服務
// scanFolder：只用內建 DriveApp
// extractPdfText：用 ScriptApp.getOAuthToken() + UrlFetchApp 做 OCR
// ==========================================

/**
 * 掃描資料夾：列出所有 PDF（不需任何進階服務）
 */
function scanFolder(folderId) {
  if (!folderId || folderId.trim() === '') throw new Error('請先填入有效的資料夾 ID');
  const cleanId = folderId.split('?')[0].trim(); // 去掉 ?usp=... 等多餘參數
  try {
    const folder = DriveApp.getFolderById(cleanId);
    const iter   = folder.getFilesByMimeType('application/pdf');
    const files  = [];
    while (iter.hasNext()) {
      const f = iter.next();
      files.push({ id: f.getId(), name: f.getName() });
    }
    return { folderName: folder.getName(), count: files.length, files: files };
  } catch (err) {
    throw new Error('無法讀取資料夾：' + err.message +
      '\n請重新部署 GAS 並在授權視窗點「允許」。');
  }
}

/**
 * 從單一 PDF 讀取文字
 * 【不需要 Drive API 進階服務】
 * 改用 ScriptApp.getOAuthToken() + UrlFetchApp 呼叫 Drive REST API
 */
function extractPdfText(fileId) {
  const token = ScriptApp.getOAuthToken();

  try {
    // 方法一：先嘗試直接 export 為純文字（適用有文字層的 PDF）
    const exportResp = UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files/' + fileId +
      '/export?mimeType=text%2Fplain',
      {
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true
      }
    );
    if (exportResp.getResponseCode() === 200) {
      const txt = exportResp.getContentText();
      if (txt && txt.trim().length > 30) {
        return txt.length > 6000 ? txt.substring(0, 6000) + '\n...(略)' : txt;
      }
    }
  } catch(e) {}

  // 方法二：OCR（適用掃描版 PDF）
  // 用 multipart upload 把 PDF 上傳並讓 Drive 自動 OCR 轉成 Google Doc
  try {
    const file     = DriveApp.getFileById(fileId);
    const blob     = file.getBlob();
    const boundary = 'ocr_' + Date.now();
    const metadata = JSON.stringify({
      name    : '_ocr_' + Date.now(),
      mimeType: 'application/vnd.google-apps.document'
    });
    const b64 = Utilities.base64Encode(blob.getBytes());
    const body =
      '--' + boundary + '\r\n' +
      'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
      metadata + '\r\n' +
      '--' + boundary + '\r\n' +
      'Content-Type: application/pdf\r\n' +
      'Content-Transfer-Encoding: base64\r\n\r\n' +
      b64 + '\r\n--' + boundary + '--';

    const uploadResp = UrlFetchApp.fetch(
      'https://www.googleapis.com/upload/drive/v3/files' +
      '?uploadType=multipart',
      {
        method     : 'POST',
        contentType: 'multipart/related; boundary=' + boundary,
        payload    : body,
        headers    : { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true
      }
    );

    if (uploadResp.getResponseCode() !== 200) {
      throw new Error('OCR 上傳失敗（HTTP ' + uploadResp.getResponseCode() + '）');
    }

    const newFileId = JSON.parse(uploadResp.getContentText()).id;
    Utilities.sleep(2000); // 等 Google 完成 OCR

    const docText = DocumentApp.openById(newFileId).getBody().getText();

    // 刪除暫存 Doc
    try {
      UrlFetchApp.fetch(
        'https://www.googleapis.com/drive/v3/files/' + newFileId,
        { method: 'DELETE', headers: { Authorization: 'Bearer ' + token },
          muteHttpExceptions: true }
      );
    } catch(e) {}

    return docText.length > 6000 ? docText.substring(0, 6000) + '\n...(略)' : docText;

  } catch (err) {
    throw new Error('PDF 讀取失敗：' + err.message);
  }
}

/**
 * 讀取資料夾內所有 PDF，合併文字後出題
 */
function generateBatchFromFolder(folderId, nodesArray, grade, subject, batchIndex) {
  if (!folderId || folderId.trim() === '') throw new Error('請先設定 Google Drive 資料夾 ID');
  const cleanId = folderId.split('?')[0].trim();

  const cache   = CacheService.getScriptCache();
  const PDF_KEY = 'PdfText_' + cleanId;
  let   pdfText = cache.get(PDF_KEY);

  if (!pdfText) {
    let folder;
    try {
      folder = DriveApp.getFolderById(cleanId);
    } catch (err) {
      throw new Error('無法開啟資料夾（' + err.message + '）');
    }

    const iter  = folder.getFilesByMimeType('application/pdf');
    const texts = [];

    while (iter.hasNext()) {
      const f = iter.next();
      try {
        const txt = extractPdfText(f.getId());
        if (txt && txt.trim().length > 0) {
          texts.push('【' + f.getName() + '】\n' + txt);
          Logger.log('已讀取：' + f.getName() + '（' + txt.length + ' 字）');
        }
      } catch (readErr) {
        Logger.log('跳過：' + f.getName() + ' → ' + readErr.message);
        texts.push('【' + f.getName() + '】（讀取失敗，略過）');
      }
    }

    if (texts.length === 0) throw new Error('資料夾內沒有可讀取的 PDF，請確認檔案格式正確。');

    pdfText = texts.join('\n\n');
    if (pdfText.length > 10000) {
      pdfText = pdfText.substring(0, 10000) + '\n\n...(內容過長，已截斷)';
    }
    try { cache.put(PDF_KEY, pdfText, 1800); } catch (e) {}
  }

  return _doGenerateBatch(nodesArray, grade, subject || '數學', batchIndex, pdfText);
}

// 4. 出題 Prompt 工廠
// 依科目回傳對應的命題指令
// ==========================================
function buildPrompt(subject, grade, batchNodes, pdfText) {
  const nodeList = batchNodes.map((n, i) => `${i + 1}. 「${n}」`).join('\n');
  const pdfSection = pdfText
    ? `\n\n【參考教材內容（請依據此內容出題，不可超出範圍）】\n---\n${pdfText}\n---\n`
    : '';

  if (subject === '國語') {
    return `你是台灣資深國小國語科命題專家。
請為「${grade}」學生，針對以下 ${batchNodes.length} 個知識節點，各設計 6 道題：${pdfSection}
${nodeList}

【國語科命題規範】
題型分配（每個節點）：
- 第1-2題：字詞辨識單選題（4選1），考字形、字音、字義辨別
- 第3-4題：填充題，填入正確詞語、標點或語詞搭配
- 第5題：句子改寫填充題（給範例句型，學生仿造填空）
- 第6題：閱讀理解單選題（4選1）

難度分配：第1-2題 easy，第3-5題 medium，第6題 hard
注意事項：
- 選擇題的干擾選項要合理，不能明顯錯誤
- 填充題答案必須是明確的詞語，避免開放性答案
- 若有提供教材內容，題目必須有教材依據
- answer 欄位：選擇題填完整選項文字，填充題填正確詞語

只回傳 JSON 陣列，不含任何說明：
[{"node":"節點","text":"題目文字","type":"single","options":["選項A","選項B","選項C","選項D"],"answer":"選項A","difficulty":"easy"}]`;
  }

  // 預設：數學科
  return `你是台灣資深國小數學命題專家。
請為「${grade}」學生，針對以下 ${batchNodes.length} 個知識節點，各設計 6 道題：${pdfSection}
${nodeList}

【數學科命題規範】
- 第1-3題：單選題（type:"single"），4個選項，選項不含雙引號
- 第4-6題：填充題（type:"fill"），options為[]
- 難度：第1題 easy，第2-4題 medium，第5-6題 hard
- 情境多樣：純計算、生活情境、錯誤辨析各至少出現一次
- answer 必須與 options 某一項完全相同（選擇題），或為純數字/分數（填充題）
${pdfText ? '- 題目必須有教材依據，數字與情境取自教材' : ''}

只回傳 JSON 陣列，不含任何說明：
[{"node":"節點","text":"題目","type":"single","options":["A","B","C","D"],"answer":"A","difficulty":"easy"}]`;
}


// ==========================================
// 5. 批次出題引擎（AI 自由發揮版）
// ==========================================
function generateBatch(nodesArray, grade, subject, batchIndex) {
  return _doGenerateBatch(nodesArray, grade, subject || '數學', batchIndex, null);
}

/**
 * 從 PDF 教材內容出題
 */
function generateBatchFromPdf(fileId, nodesArray, grade, subject, batchIndex) {
  if (!fileId || fileId.trim() === '') throw new Error('請提供 PDF 檔案 ID');
  const pdfText = extractPdfText(fileId);
  if (!pdfText || pdfText.trim().length < 50) throw new Error('PDF 內容讀取失敗或內容過少，請確認檔案可讀取。');
  return _doGenerateBatch(nodesArray, grade, subject || '數學', batchIndex, pdfText);
}

/**
 * 核心出題邏輯（AI自由 / PDF 共用）
 */
function _doGenerateBatch(nodesArray, grade, subject, batchIndex, pdfText) {
  const cfg = getConfig();
  const apiKey = String(cfg.GeminiAPIKey || '');
  if (!apiKey || apiKey.includes('請在此')) throw new Error("請先在 Config 試算表設定 Gemini API Key！");

  const validNodes = nodesArray.filter(n => n && n.trim() !== '');
  const BATCH_SIZE = 5;
  const batches = [];
  for (let i = 0; i < validNodes.length; i += BATCH_SIZE) batches.push(validNodes.slice(i, i + BATCH_SIZE));
  const totalBatches = batches.length;

  if (batchIndex >= totalBatches) {
    _clearAllCaches();
    return { done: true, current: totalBatches, total: totalBatches, message: `✅ 全部完成！共 ${totalBatches} 批次。` };
  }

  const batchNodes = batches[batchIndex];
  const prompt = buildPrompt(subject, grade, batchNodes, pdfText);

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
  const fetchOpts = {
    method: "post", contentType: "application/json",
    payload: JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { responseMimeType: "application/json", temperature: 0.85 }
    }),
    muteHttpExceptions: true
  };

  let result = null;
  for (let retry = 0; retry < 3; retry++) {
    const resp = UrlFetchApp.fetch(url, fetchOpts);
    result = JSON.parse(resp.getContentText());
    if (result.error) {
      if (result.error.message.toLowerCase().includes("quota") || result.error.message.includes("429")) {
        Utilities.sleep(20000);
      } else {
        throw new Error("Gemini API 錯誤: " + result.error.message);
      }
    } else break;
  }
  if (result && result.error) throw new Error("API 持續失敗，請等待 1 分鐘後重試。");

  const rawText = result.candidates[0].content.parts[0].text.replace(/```json/g, '').replace(/```/g, '').trim();
  const questions = JSON.parse(rawText);

  const bankSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Bank');
  const lastRow = bankSheet.getLastRow() || 1;
  const newRows = questions.map((q, idx) => [
    `AI-${subject}-${Date.now().toString().slice(-6)}-${batchIndex}-${idx}`,
    q.node, q.text, q.type,
    JSON.stringify(q.options || []),
    normalizeAnswer(q.answer),
    q.difficulty, grade,
    subject   // ← 第 9 欄：科目
  ]);

  if (newRows.length > 0) {
    bankSheet.getRange(lastRow + 1, 1, newRows.length, 9).setValues(newRows);
    bankSheet.getRange(lastRow + 1, 6, newRows.length, 1).setNumberFormat('@STRING@');
  }

  const isLast = (batchIndex + 1 >= totalBatches);
  if (isLast) _clearAllCaches();

  return {
    done: isLast,
    current: batchIndex + 1,
    total: totalBatches,
    addedThisBatch: newRows.length,
    nodes: batchNodes,
    message: `[${subject}] 第 ${batchIndex + 1}/${totalBatches} 批完成（${batchNodes.join('、')}）— 新增 ${newRows.length} 題${pdfText ? '（PDF出題）' : ''}`
  };
}


// ==========================================
// 6. 派題引擎（新增科目篩選）
// ==========================================
function normalizeAnswer(str) {
  if (str === null || str === undefined) return '';
  if (str instanceof Date) return '';
  if (typeof str === 'number') return String(Number.isInteger(str) ? str : str).trim().toLowerCase();
  return String(str).trim()
    .replace(/\s+/g, '')
    .replace(/[\uff10-\uff19]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
    .replace(/，/g, ',')
    .toLowerCase();
}

function generateQuiz(weakNode, taskName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getConfig();
  const quizCount = parseInt(cfg.QuizCount, 10) || 10;

  // 取得任務的年級與科目
  // 【防呆】先從任務分頁第5欄讀科目，再從 History 讀年級
  let targetGrade = '', targetSubject = '數學';

  // 方法一：從任務分頁第 5 欄直接讀科目（最可靠，不受試算表版本影響）
  const taskSheetForSubject = ss.getSheetByName(taskName);
  if (taskSheetForSubject && taskSheetForSubject.getLastRow() > 1) {
    try {
      const subjectVal = String(taskSheetForSubject.getRange(2, 5, 1, 1).getValue() || '').trim();
      if (subjectVal === '國語' || subjectVal === '數學') targetSubject = subjectVal;
    } catch(e) {}
  }

  // 方法二：從 History 讀年級，科目只有是明確文字才採用（防止讀到舊格式的數字）
  const historySheet = ss.getSheetByName('History');
  if (historySheet && taskName) {
    historySheet.getDataRange().getValues().slice(1).forEach(r => {
      if (String(r[1]).replace(/'/g, '').trim() === String(taskName).trim()) {
        targetGrade = String(r[2] || '').trim();
        const r3 = String(r[3] || '').trim();
        if (r3 === '國語' || r3 === '數學') targetSubject = r3;
      }
    });
  }

  const cache = CacheService.getScriptCache();
  const CACHE_KEY = 'BankData_V3';
  let allQuestions = [];
  const cachedData = cache.get(CACHE_KEY);

  if (cachedData) {
    allQuestions = JSON.parse(cachedData);
  } else {
    const bankSheet = ss.getSheetByName('Bank');
    if (!bankSheet || bankSheet.getLastRow() <= 1) return [];
    const lastRow = bankSheet.getLastRow();
    bankSheet.getRange(2, 6, lastRow - 1, 1).setNumberFormat('@STRING@');
    const data        = bankSheet.getRange(2, 1, lastRow - 1, 9).getValues();
    const displayData = bankSheet.getRange(2, 1, lastRow - 1, 9).getDisplayValues();

    data.forEach((row, i) => {
      if (!row[0] && !row[2]) return;
      let options = [];
      const rawOpts = row[4] ? String(row[4]).trim() : '';
      if (rawOpts) {
        if (rawOpts.startsWith('[')) { try { options = JSON.parse(rawOpts); } catch (e) {} }
        else { options = rawOpts.split(',').map(o => o.trim()).filter(o => o); }
      }
      const rawAnswer = (displayData[i] && displayData[i][5]) ? displayData[i][5] : String(row[5] || '');
      allQuestions.push({
        id:            row[0] || `Q${i + 2}`,
        node:          row[1] ? String(row[1]).trim() : '',
        text:          row[2],
        type:          row[3],
        options,
        answer:        normalizeAnswer(rawAnswer),
        displayAnswer: String(rawAnswer).trim(),
        difficulty:    String(row[6] || 'medium').trim(),
        grade:         String(row[7] || '').trim(),
        subject:       (row[8] && (String(row[8]).trim() === '國語' || String(row[8]).trim() === '數學')) ? String(row[8]).trim() : '',  // 第 9 欄，空白表示不限科目
      });
    });
    try { cache.put(CACHE_KEY, JSON.stringify(allQuestions), 3600); } catch (e) {}
  }

  const shuffle = arr => {
    for (let i = arr.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [arr[i], arr[j]] = [arr[j], arr[i]];
    }
    return arr;
  };

  // 篩選：科目 + 年級 + 弱點節點
  let targets = [], fallbacks = [];
  const safeNode = String(weakNode || '').trim();
  allQuestions.forEach(q => {
    // 科目不符合就跳過
    if (q.subject && q.subject !== targetSubject) return;
    // 年級不符合就跳過
    if (targetGrade && q.grade && q.grade !== targetGrade) return;

    if (q.node && safeNode && (q.node.includes(safeNode) || safeNode.includes(q.node))) {
      targets.push(q);
    } else if (!targetGrade || !q.grade || q.grade === targetGrade) {
      fallbacks.push(q);
    }
  });

  const pickByDiff = (arr, n) => {
    const e = shuffle(arr.filter(q => q.difficulty === 'easy'));
    const m = shuffle(arr.filter(q => q.difficulty === 'medium'));
    const h = shuffle(arr.filter(q => q.difficulty === 'hard'));
    const ec = Math.round(n * 0.3), hc = Math.round(n * 0.2), mc = n - ec - hc;
    let res = [...e.slice(0, ec), ...m.slice(0, mc), ...h.slice(0, hc)];
    if (res.length < n) {
      const rest = shuffle([...e.slice(ec), ...m.slice(mc), ...h.slice(hc)]);
      res = res.concat(rest.slice(0, n - res.length));
    }
    return res;
  };

  shuffle(targets);
  let final = pickByDiff(targets, Math.min(quizCount, targets.length));
  if (final.length < quizCount) {
    shuffle(fallbacks);
    final = final.concat(pickByDiff(fallbacks, Math.min(quizCount - final.length, fallbacks.length)));
  }
  final = final.map(q => q.type === 'single' && q.options.length > 1 ? { ...q, options: shuffle([...q.options]) } : q);
  return shuffle(final);
}

function submitQuizResult(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let resultSheet = ss.getSheetByName('Results');

  // 若 Results 分頁不存在，自動建立
  if (!resultSheet) {
    resultSheet = ss.insertSheet('Results');
    resultSheet.appendRow(['測驗時間', '任務名稱', '座號', '姓名', '分數', '作答歷時(秒)', '作答明細']);
  }

  if (!data || !data.taskName) throw new Error('作答資料格式錯誤，缺少 taskName');

  resultSheet.appendRow([
    new Date(),
    data.taskName,
    data.seatNo  || '',
    data.name    || '',
    data.score   || 0,
    data.timeSpent || 0,
    JSON.stringify(data.details || [])
  ]);
  return true;
}


// ==========================================
// 7. 教學講義生成（新增科目支援）
// ==========================================
function generateTeachingWorksheet(topNodes, grade, subject) {
  const cfg = getConfig();
  const apiKey = String(cfg.GeminiAPIKey || '');
  if (!apiKey || apiKey.includes('請在此')) throw new Error("請先設定 Gemini API Key！");

  const subj = subject || '數學';
  const nodeList = topNodes.map((n, i) => `${i + 1}. 「${n}」`).join('\n');

  let prompt = '';
  if (subj === '國語') {
    prompt = `你是台灣國小國語科教師，擅長設計補救教學練習。
請針對「${grade}」學生，為以下知識節點設計「教師教學用講義」：

${nodeList}

規範：
- 每個節點設計 5 題，題型為填充題（type: "fill"）或簡單選擇題（type: "single"）
- 難度全部 easy（最基礎，只考核心概念）
- step：寫出學習步驟（如「先找出生字的部首」）
- hint：一句教學提示（如「注意左右結構的寫法！」）
- answer：選擇題填完整選項，填充題填正確詞語

只回傳 JSON 陣列：
[{"node":"節點","step":"步驟","text":"題目","type":"fill","answer":"答案","hint":"提示"}]`;
  } else {
    prompt = `你是台灣國小數學科教師，擅長設計補救教學練習。
請針對「${grade}」學生，為以下知識節點設計「教師教學用講義」：

${nodeList}

規範：
- 每個節點設計 5 題，全部為填充計算題（type: "fill"）
- 難度全部 easy（最基礎的直接計算）
- step：引導步驟（如「先通分 → 再相加」）
- hint：一句教學提示
- answer：最終數字或分數答案

只回傳 JSON 陣列：
[{"node":"節點","step":"步驟","text":"題目","type":"fill","answer":"答案","hint":"提示"}]`;
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
  const opts = {
    method: "post", contentType: "application/json",
    payload: JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { responseMimeType: "application/json", temperature: 0.5 }
    }),
    muteHttpExceptions: true
  };

  let result = null;
  for (let retry = 0; retry < 3; retry++) {
    const resp = UrlFetchApp.fetch(url, opts);
    result = JSON.parse(resp.getContentText());
    if (result.error) {
      if (result.error.message.toLowerCase().includes("quota")) { Utilities.sleep(15000); }
      else throw new Error("API 錯誤: " + result.error.message);
    } else break;
  }
  if (result && result.error) throw new Error("API 持續失敗，請稍後再試。");

  const rawText = result.candidates[0].content.parts[0].text.replace(/```json/g, '').replace(/```/g, '').trim();
  return JSON.parse(rawText);
}


// ==========================================
// 8. 快取管理
// ==========================================
function _clearAllCaches() {
  ['BankData_V3', 'BankData_V2', 'BankData_V1'].forEach(key => {
    try { CacheService.getScriptCache().remove(key); } catch (e) {}
  });
}

function clearBankCache() {
  _clearAllCaches();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bankSheet = ss.getSheetByName('Bank');
  if (bankSheet && bankSheet.getLastRow() > 1) {
    bankSheet.getRange(2, 6, bankSheet.getLastRow() - 1, 1).setNumberFormat('@STRING@');
  }
  return "✅ 快取已清除，正解欄格式已修正！";
}


// ==========================================
// 9. 教師儀表板：成績分析
// ==========================================
function getTaskResults(taskName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultSheet = ss.getSheetByName('Results');
  if (!resultSheet || resultSheet.getLastRow() <= 1) return [];

  const data = resultSheet.getDataRange().getValues();
  const results = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowTask   = String(row[1] || '').replace(/'/g, '').trim();
    const cleanTask = String(taskName || '').replace(/'/g, '').trim();
    if (rowTask !== cleanTask) continue;
    let details = [];
    try { details = JSON.parse(row[6] || '[]'); } catch (e) {}
    results.push({
      submittedAt: row[0] ? new Date(row[0]).toLocaleString('zh-TW') : '',
      taskName:    rowTask,
      seatNo:      String(row[2] || ''),
      name:        String(row[3] || ''),
      score:       Number(row[4] || 0),
      timeSpent:   Number(row[5] || 0),
      details:     details
    });
  }

  const seen = {}, deduped = [];
  for (let i = results.length - 1; i >= 0; i--) {
    const key = results[i].seatNo + '_' + results[i].name;
    if (!seen[key]) { seen[key] = true; deduped.unshift(results[i]); }
  }
  return deduped.sort((a, b) => Number(a.seatNo) - Number(b.seatNo));
}

function analyzeClassWeakNodes(taskName) {
  const results = getTaskResults(taskName);
  if (!results.length) return [];
  const nodeMap = {};
  results.forEach(r => {
    (r.details || []).forEach(d => {
      if (!d.isCorrect && d.node) {
        const node = String(d.node).trim();
        if (!nodeMap[node]) nodeMap[node] = new Set();
        nodeMap[node].add(r.name || r.seatNo);
      }
    });
  });
  const totalStudents = results.length;
  return Object.entries(nodeMap).map(([node, students]) => ({
    node,
    wrongCount:   students.size,
    studentCount: totalStudents,
    wrongRate:    Math.round((students.size / totalStudents) * 100),
    students:     Array.from(students)
  })).sort((a, b) => b.wrongCount - a.wrongCount);
}

function getQuestionErrorRates(taskName) {
  const results = getTaskResults(taskName);
  if (!results.length) return [];
  const qMap = {};
  results.forEach(r => {
    (r.details || []).forEach(d => {
      const qid = d.id || d.questionText || 'unknown';
      if (!qMap[qid]) {
        qMap[qid] = {
          questionText: d.questionText || qid,
          node:         d.node || '',
          correctAns:   d.displayCorrectAns || d.correctAns || '',
          total: 0, wrong: 0, wrongAnswers: []
        };
      }
      qMap[qid].total++;
      if (!d.isCorrect) {
        qMap[qid].wrong++;
        qMap[qid].wrongAnswers.push({ name: r.name, seatNo: r.seatNo, ans: d.userAns || '未作答' });
      }
    });
  });
  return Object.values(qMap)
    .map(q => ({ ...q, wrongRate: Math.round((q.wrong / q.total) * 100) }))
    .sort((a, b) => b.wrongRate - a.wrongRate);
}

// ==========================================
// 10. 測試用：確認 UrlFetchApp 權限
// ==========================================
function testAIGeneration() {
  UrlFetchApp.fetch("https://www.google.com");
  Logger.log("外部連線權限正常！");
}
