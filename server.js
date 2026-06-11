/**
 * asahipack-plate-assistant - Standalone Desktop Server
 * This script starts a local HTTP server to host the static Next.js export
 * and implements the file-processing API. It is packaged into a standalone EXE using pkg.
 */

const http = require('http');
const fs = require('fs').promises;
const path = require('path');
const os = require('os');
const { exec } = require('child_process');
const xlsx = require('xlsx');
const ExcelJS = require('exceljs');
const JapaneseHolidays = require('japanese-holidays');

const PORT = 3000;

// MIMEタイプの定義
const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.gif': 'image/gif',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',
  '.txt': 'text/plain; charset=utf-8'
};

/**
 * 日本の祝日を考慮して、指定された年月の最終営業日を YYYY/MM/DD 形式で取得する
 * @param {number} year 年
 * @param {number} month 月 (1-12)
 * @returns {string} YYYY/MM/DD 形式の最終営業日文字列
 */
function getLastBusinessDay(year, month) {
  const date = new Date(year, month, 0); // 月末日
  while (true) {
    const dayOfWeek = date.getDay();
    const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
    const isHoliday = JapaneseHolidays.isHoliday(date) !== undefined;
    
    if (!isWeekend && !isHoliday) {
      break;
    }
    date.setDate(date.getDate() - 1);
  }
  
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}/${m}/${d}`;
}

/**
 * 仕入先名から種別を特定する
 * @param {string} supplierName 仕入先名
 * @returns {string} 種別名
 */
function getProductType(supplierName) {
  if (!supplierName) return '別注';
  const name = supplierName.trim();
  if (name.includes('シルク印刷')) return 'シルク印刷';
  if (name.includes('３Ｆロール印刷')) return 'SP';
  if (name.includes('オクダ') || name.includes('エイト')) return 'オフセット';
  if (name.includes('３Ｆロールカット')) return 'カット';
  if (name.includes('アサヒパック')) return 'アサヒパック';
  return '別注';
}

/**
 * ExcelのB2セルから担当者名を取得する
 * @param {Buffer} fileBuffer 
 * @returns {string} 担当者名
 */
function parseStaffName(fileBuffer) {
  try {
    const workbook = xlsx.read(fileBuffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const staffCell = worksheet['B2'];
    if (staffCell && staffCell.v) {
      const val = String(staffCell.v).trim();
      return val.replace(/^(担当者|担当)[：:]\s*/, '');
    }
  } catch (e) {
    console.error('Staff name parsing error:', e);
  }
  return '担当者';
}

/**
 * 得意先名を正規化し、(株)や（株）などの略称を「株式会社」に展開する。有限会社・合同会社も同様。
 * @param {string} name 元の得意先名
 * @returns {string} 正規化された得意先名
 */
function normalizeClientName(name) {
  if (!name) return '';
  let normalized = name
    .replace(/\(株\)/g, '株式会社')
    .replace(/（株）/g, '株式会社')
    .replace(/㈱/g, '株式会社')
    .replace(/\(有\)/g, '有限会社')
    .replace(/（有）/g, '有限会社')
    .replace(/㈲/g, '有限会社')
    .replace(/\(合\)/g, '合同会社')
    .replace(/（合）/g, '合同会社')
    .replace(/\(同\)/g, '合同会社')
    .replace(/（同）/g, '合同会社')
    .replace(/㈏/g, '合同会社');
  
  normalized = normalized.replace(/\s*株式会社\s*/g, '株式会社');
  normalized = normalized.replace(/\s*有限会社\s*/g, '有限会社');
  normalized = normalized.replace(/\s*合同会社\s*/g, '合同会社');
  return normalized.trim();
}

/**
 * アップロードされたExcelを解析しレコード配列にする
 * @param {Buffer} fileBuffer 
 * @returns {Array} レコードの配列
 */
function parseExcel(fileBuffer) {
  const workbook = xlsx.read(fileBuffer, { type: 'buffer' });
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];

  const range = xlsx.utils.decode_range(worksheet['!ref'] || 'A1:R100');
  const records = [];

  for (let r = 4; r <= range.e.r; r++) {
    const clientCodeCell = worksheet[xlsx.utils.encode_cell({ r, c: 0 })];
    if (!clientCodeCell || !clientCodeCell.v) {
      continue;
    }

    const clientCode = String(clientCodeCell.v).trim();
    const clientNameRaw = worksheet[xlsx.utils.encode_cell({ r, c: 1 })]?.v ? String(worksheet[xlsx.utils.encode_cell({ r, c: 1 })].v).trim() : '';
    const clientName = normalizeClientName(clientNameRaw);
    const supplierCode = worksheet[xlsx.utils.encode_cell({ r, c: 2 })]?.v ? String(worksheet[xlsx.utils.encode_cell({ r, c: 2 })].v).trim() : '';
    const supplierName = worksheet[xlsx.utils.encode_cell({ r, c: 3 })]?.v ? String(worksheet[xlsx.utils.encode_cell({ r, c: 3 })].v).trim() : '';
    const plateNo = worksheet[xlsx.utils.encode_cell({ r, c: 4 })]?.v ? String(worksheet[xlsx.utils.encode_cell({ r, c: 4 })].v).trim() : '';
    const orderNo = worksheet[xlsx.utils.encode_cell({ r, c: 5 })]?.v !== undefined ? String(worksheet[xlsx.utils.encode_cell({ r, c: 5 })].v).trim() : '';
    const orderSuffix = worksheet[xlsx.utils.encode_cell({ r, c: 6 })]?.v !== undefined ? String(worksheet[xlsx.utils.encode_cell({ r, c: 6 })].v).trim() : '';
    const orderSub = worksheet[xlsx.utils.encode_cell({ r, c: 7 })]?.v !== undefined ? String(worksheet[xlsx.utils.encode_cell({ r, c: 7 })].v).trim() : '';
    const weight = worksheet[xlsx.utils.encode_cell({ r, c: 8 })]?.v !== undefined ? worksheet[xlsx.utils.encode_cell({ r, c: 8 })].v : '';
    const finish = worksheet[xlsx.utils.encode_cell({ r, c: 9 })]?.v ? String(worksheet[xlsx.utils.encode_cell({ r, c: 9 })].v).trim() : '';
    const storeName = worksheet[xlsx.utils.encode_cell({ r, c: 10 })]?.v ? String(worksheet[xlsx.utils.encode_cell({ r, c: 10 })].v).trim() : '';
    const brandName = worksheet[xlsx.utils.encode_cell({ r, c: 11 })]?.v ? String(worksheet[xlsx.utils.encode_cell({ r, c: 11 })].v).trim() : '';
    const colorCount = typeof worksheet[xlsx.utils.encode_cell({ r, c: 12 })]?.v === 'number' ? Number(worksheet[xlsx.utils.encode_cell({ r, c: 12 })].v) : 0;
    const colors = worksheet[xlsx.utils.encode_cell({ r, c: 13 })]?.v ? String(worksheet[xlsx.utils.encode_cell({ r, c: 13 })].v).trim() : '';
    const lastUsedDate = worksheet[xlsx.utils.encode_cell({ r, c: 14 })]?.v ? String(worksheet[xlsx.utils.encode_cell({ r, c: 14 })].v).trim() : '';
    const expiryDate = worksheet[xlsx.utils.encode_cell({ r, c: 15 })]?.v ? String(worksheet[xlsx.utils.encode_cell({ r, c: 15 })].v).trim() : '';
    const profit = typeof worksheet[xlsx.utils.encode_cell({ r, c: 16 })]?.v === 'number' ? Number(worksheet[xlsx.utils.encode_cell({ r, c: 16 })].v) : 0;

    records.push({
      clientCode, clientName, supplierCode, supplierName, plateNo, orderNo, orderSuffix, orderSub,
      weight, finish, storeName, brandName, colorCount, colors, lastUsedDate, expiryDate, profit
    });
  }
  return records;
}

/**
 * 得意先別の確認Excelを作成する
 * @param {object} group 
 * @param {string} staffName 
 * @returns {Promise<Buffer>}
 */
async function generateClientExcel(group, staffName) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('落版確認書', {
    pageSetup: {
      paperSize: 9,              // A4
      orientation: 'landscape',  // 横向き
      fitToPage: true,           // ページに合わせる
      fitToWidth: 1,             // 横幅を1ページに収める
      fitToHeight: 0             // 縦は自動
    }
  });

  worksheet.views = [{ showGridLines: true }];

  worksheet.columns = [
    { key: 'orderNo', width: 12 }, { key: 'orderSuffix', width: 4 }, { key: 'orderSub', width: 6 },
    { key: 'weight', width: 8 }, { key: 'finish', width: 8 }, { key: 'productType', width: 12 },
    { key: 'storeName', width: 15 }, { key: 'brandName', width: 25 }, { key: 'colorCount', width: 8 },
    { key: 'colors', width: 25 }, { key: 'lastUsedDate', width: 14 }, { key: 'expiryDate', width: 14 },
    { key: 'answer', width: 12 }
  ];

  const today = new Date();
  const dateString = `${today.getFullYear()}/${String(today.getMonth() + 1).padStart(2, '0')}/${String(today.getDate()).padStart(2, '0')}`;

  worksheet.getCell('A1').value = `${group.clientName} 御中`;
  worksheet.getCell('A1').font = { name: 'MS Gothic', size: 14, bold: true };

  worksheet.getCell('K1').value = '株式会社アサヒパック';
  worksheet.getCell('K1').font = { name: 'MS Gothic', size: 10, bold: true };
  worksheet.getCell('K2').value = `作成日: ${dateString}`;
  worksheet.getCell('K2').font = { name: 'MS Gothic', size: 10 };
  worksheet.getCell('K3').value = `担当者: ${staffName}`;
  worksheet.getCell('K3').font = { name: 'MS Gothic', size: 10 };

  worksheet.mergeCells('A4:M4');
  const titleCell = worksheet.getCell('A4');
  titleCell.value = '落版候補リストのご確認について';
  titleCell.font = { name: 'MS Gothic', size: 16, bold: true, underline: true };
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
  worksheet.getRow(4).height = 30;

  worksheet.mergeCells('A5:M5');
  const descCell = worksheet.getCell('A5');
  descCell.value = 'いつも大変お世話になっております。弊社で保管しております御社の版につきまして、最終使用日より期間が経過したものを落版（廃棄）候補としてリストアップいたしました。お手数ですが、落版の可否（廃棄/継続）をご記入の上、ご返送いただけますようお願い申し上げます。';
  descCell.font = { name: 'MS Gothic', size: 10 };
  descCell.alignment = { wrapText: true, vertical: 'top' };
  worksheet.getRow(5).height = 45;

  const headerRowNumber = 7;
  const headerRow = worksheet.getRow(headerRowNumber);
  headerRow.height = 25;

  worksheet.mergeCells(`A${headerRowNumber}:C${headerRowNumber}`);
  worksheet.getCell(`A${headerRowNumber}`).value = '受注No.';
  worksheet.getCell(`D${headerRowNumber}`).value = 'Ｋg';
  worksheet.getCell(`E${headerRowNumber}`).value = '仕上';
  worksheet.getCell(`F${headerRowNumber}`).value = '種別';
  worksheet.getCell(`G${headerRowNumber}`).value = '店名';
  worksheet.getCell(`H${headerRowNumber}`).value = '銘柄';
  worksheet.getCell(`I${headerRowNumber}`).value = '色数';
  worksheet.getCell(`J${headerRowNumber}`).value = '色';
  worksheet.getCell(`K${headerRowNumber}`).value = '最終使用日';
  worksheet.getCell(`L${headerRowNumber}`).value = '落版予定日';
  
  const answerHeaderCell = worksheet.getCell(`M${headerRowNumber}`);
  answerHeaderCell.value = '回答（必須）';

  const headerCols = ['A', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'];
  headerCols.forEach((col) => {
    const cell = worksheet.getCell(`${col}${headerRowNumber}`);
    cell.font = { name: 'MS Gothic', size: 10, bold: true };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = {
      top: { style: 'medium' }, left: { style: 'thin' }, bottom: { style: 'medium' }, right: { style: 'thin' }
    };
    
    if (col === 'M') {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF0D0' } };
    } else {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE6F2FF' } };
    }
  });

  ['B', 'C'].forEach((col) => {
    const cell = worksheet.getCell(`${col}${headerRowNumber}`);
    cell.border = {
      top: { style: 'medium' }, bottom: { style: 'medium' },
      left: col === 'B' ? undefined : { style: 'thin' },
      right: col === 'C' ? { style: 'thin' } : undefined
    };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE6F2FF' } };
  });

  let currentRow = 8;
  group.records.forEach((record) => {
    const row = worksheet.getRow(currentRow);
    row.height = 20;

    const productType = getProductType(record.supplierName);
    let displayExpiryDate = record.expiryDate;
    
    if (['SP', 'シルク印刷', 'オフセット'].includes(productType)) {
      const match = record.expiryDate.match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})/);
      if (match) {
        const year = parseInt(match[1], 10);
        const month = parseInt(match[2], 10);
        displayExpiryDate = getLastBusinessDay(year, month);
      } else {
        const now = new Date();
        displayExpiryDate = getLastBusinessDay(now.getFullYear(), now.getMonth() + 1);
      }
    }

    worksheet.getCell(`A${currentRow}`).value = record.orderNo;
    worksheet.getCell(`B${currentRow}`).value = record.orderSuffix;
    worksheet.getCell(`C${currentRow}`).value = record.orderSub ? Number(record.orderSub) : '';
    worksheet.getCell(`D${currentRow}`).value = record.weight ? Number(record.weight) : '';
    worksheet.getCell(`E${currentRow}`).value = record.finish;
    worksheet.getCell(`F${currentRow}`).value = productType;
    worksheet.getCell(`G${currentRow}`).value = record.storeName;
    worksheet.getCell(`H${currentRow}`).value = record.brandName;
    worksheet.getCell(`I${currentRow}`).value = record.colorCount ? Number(record.colorCount) : '';
    worksheet.getCell(`J${currentRow}`).value = record.colors;
    worksheet.getCell(`K${currentRow}`).value = record.lastUsedDate;
    worksheet.getCell(`L${currentRow}`).value = displayExpiryDate;
    worksheet.getCell(`M${currentRow}`).value = '';

    const cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'];
    cols.forEach((col) => {
      const cell = worksheet.getCell(`${col}${currentRow}`);
      cell.font = { name: 'MS Gothic', size: 10 };
      cell.border = {
        top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }
      };

      if (['A', 'B', 'C', 'D', 'F', 'I', 'K', 'L', 'M'].includes(col)) {
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      } else {
        cell.alignment = { horizontal: 'left', vertical: 'middle' };
      }

      if (col === 'M') {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFDF0' } };
        cell.dataValidation = {
          type: 'list',
          allowBlank: true,
          formulae: ['"廃棄,継続"'],
          showErrorMessage: true,
          errorTitle: '入力エラー',
          error: 'リストから「廃棄」または「継続」を選択してください。'
        };
      }
    });

    currentRow++;
  });

  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer);
}

/**
 * 共有フォルダから受注Noの画像を探す
 * @param {string} srcDir 
 * @param {string} orderNo 
 * @returns {Promise<string[]>}
 */
async function scanImages(srcDir, orderNo) {
  if (!orderNo) return [];
  try {
    const files = await fs.readdir(srcDir);
    const lowerOrderNo = orderNo.toLowerCase();
    return files.filter(f => f.toLowerCase().startsWith(lowerOrderNo));
  } catch (e) {
    return [];
  }
}

/**
 * 日付スタンプ文字列
 * @returns {string} YYYYMMDD_HHMMSS
 */
function getTimestampString() {
  const now = new Date();
  return `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}_${String(now.getHours()).padStart(2, '0')}${String(now.getMinutes()).padStart(2, '0')}${String(now.getSeconds()).padStart(2, '0')}`;
}

// HTTPサーバーの構築
const server = http.createServer(async (req, res) => {
  // CORSヘッダーの設定 (ローカル検証用)
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.writeHead(200);
    res.end();
    return;
  }

  // フォルダ選択APIルート
  if (req.url === '/api/select-directory' && req.method === 'POST') {
    const psCommand = `powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -sta -Command "Add-Type -AssemblyName System.Windows.Forms; $dialog = New-Object System.Windows.Forms.FolderBrowserDialog; $dialog.Description = 'フォルダを選択してください'; $dialog.ShowNewFolderButton = $true; if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { Write-Output $dialog.SelectedPath }"`;

    exec(psCommand, (err, stdout, stderr) => {
      if (err) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ success: false, error: err.message }));
        return;
      }
      const selectedPath = stdout.trim();
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ success: true, selectedPath }));
    });
    return;
  }

  // APIルート
  if (req.url === '/api/process' && req.method === 'POST') {
    let bodyData = '';
    req.on('data', chunk => {
      bodyData += chunk;
    });

    req.on('end', async () => {
      const logs = [];
      const addLog = (type, message) => {
        logs.push({
          type,
          message,
          timestamp: new Date().toLocaleTimeString('ja-JP')
        });
      };

      try {
        const params = JSON.parse(bodyData);
        const { imageSrcDir, outputDestDir, excelData } = params;

        if (!excelData) {
          addLog('error', 'Excelデータがありません。');
          res.writeHead(400, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ success: false, message: 'Excel data is required', logs }));
          return;
        }

        addLog('info', 'Excelデータの解析を開始します...');
        const buffer = Buffer.from(excelData.split(',')[1] || excelData, 'base64');
        
        const records = parseExcel(buffer);
        const staffName = parseStaffName(buffer);
        addLog('success', `Excelの解析が完了しました。担当者: ${staffName}様, レコード数: ${records.length}件`);

        if (records.length === 0) {
          addLog('warning', '有効なレコードが見つかりませんでした。');
          res.writeHead(200, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ success: true, clientGroups: [], logs }));
          return;
        }

        // 得意先ごとにグループ化
        const groupsMap = new Map();
        records.forEach((rec) => {
          const key = `${rec.clientCode}___${rec.clientName}`;
          if (!groupsMap.has(key)) {
            groupsMap.set(key, []);
          }
          groupsMap.get(key).push(rec);
        });

        const clientGroups = Array.from(groupsMap.entries()).map(([key, recs]) => {
          const [clientCode, clientName] = key.split('___');
          return { clientCode, clientName, records: recs };
        });

        const defaultDesktop = path.join(os.homedir(), 'Desktop');
        const rootOutputDir = outputDestDir ? outputDestDir : path.join(defaultDesktop, `落版連絡_${getTimestampString()}`);

        await fs.mkdir(rootOutputDir, { recursive: true });
        addLog('info', `出力先フォルダを作成しました: ${rootOutputDir}`);

        let isSrcDirAccessible = true;
        try {
          await fs.access(imageSrcDir);
          addLog('info', `画像フォルダに接続しました: ${imageSrcDir}`);
        } catch (e) {
          isSrcDirAccessible = false;
          addLog('warning', `画像フォルダにアクセスできません。画像コピーをスキップします。`);
        }

        let totalImagesCopied = 0;
        let processedClientsCount = 0;

        for (const group of clientGroups) {
          const safeClientName = group.clientName.replace(/[\\/:*?"<>|]/g, '_').trim();
          const clientFolder = path.join(rootOutputDir, `${group.clientCode}_${safeClientName}様`);
          
          await fs.mkdir(clientFolder, { recursive: true });

          const clientExcelBuffer = await generateClientExcel(group, staffName);
          const excelPath = path.join(clientFolder, `落版確認書_${safeClientName}.xlsx`);
          await fs.writeFile(excelPath, clientExcelBuffer);
          addLog('info', `[${safeClientName}] 得意先確認Excelを作成しました。`);

          let clientImagesCount = 0;
          if (isSrcDirAccessible) {
            const processedOrderNos = new Set();
            for (const record of group.records) {
              const orderNo = record.orderNo;
              if (!orderNo || processedOrderNos.has(orderNo)) continue;
              processedOrderNos.add(orderNo);

              const matchedFiles = await scanImages(imageSrcDir, orderNo);
              for (const file of matchedFiles) {
                const srcPath = path.join(imageSrcDir, file);
                const destPath = path.join(clientFolder, file);
                try {
                  await fs.copyFile(srcPath, destPath);
                  clientImagesCount++;
                  totalImagesCopied++;
                } catch (e) {
                  addLog('warning', `画像コピー失敗: ${file}`);
                }
              }
            }
            if (clientImagesCount > 0) {
              addLog('success', `[${safeClientName}] 画像を ${clientImagesCount} 件コピーしました。`);
            }
          }
          processedClientsCount++;
        }

        addLog('success', 'すべての処理が正常に完了しました！');
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({
          success: true,
          processedClientsCount,
          totalImagesCopied,
          outputDir: rootOutputDir,
          logs,
          clientGroups,
          staffName
        }));

      } catch (e) {
        addLog('error', `システム例外: ${e.message}`);
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ success: false, message: e.message, logs }));
      }
    });
    return;
  }

  // 静的アセットの配信
  // pkgでアセットを埋め込んだ場合、__dirnameからの相対パスでアクセスできます
  let reqPath = req.url === '/' ? '/index.html' : req.url;
  // URLパラメータ（?filePath=...等）をトリミング
  reqPath = reqPath.split('?')[0];

  const filePath = path.join(__dirname, 'out', reqPath);
  const ext = path.extname(filePath);
  const contentType = MIME_TYPES[ext] || 'application/octet-stream';

  try {
    const data = await fs.readFile(filePath);
    res.writeHead(200, { 'Content-Type': contentType });
    res.end(data);
  } catch (e) {
    // 404エラーの場合は、Next.jsのHTMLにルーティングする (SPA対応)
    try {
      const fallbackData = await fs.readFile(path.join(__dirname, 'out', 'index.html'));
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      res.end(fallbackData);
    } catch (err) {
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('Not Found');
    }
  }
});

// サーバーエラーハンドリング（ポート衝突など）
server.on('error', (err) => {
  if (err.code === 'EADDRINUSE') {
    console.error(`\n【エラー】ポート ${PORT} がすでに使用されています。`);
    console.error(`・すでにこのツールが起動している可能性があります。別のウィンドウで起動中ではないかご確認ください。`);
    console.error(`・または、他のアプリケーションがポート ${PORT} を使用しています。\n`);
  } else {
    console.error(`\n【サーバーエラー】起動中にエラーが発生しました: ${err.message}\n`);
  }
  console.log('エンターキーを押すと終了します...');
  process.stdin.resume();
  process.stdin.on('data', () => {
    process.exit(1);
  });
});

// サーバー起動
server.listen(PORT, () => {
  console.log(`====================================================`);
  console.log(`  落版連絡ツール (asahipack-plate-assistant)`);
  console.log(`  ローカルサーバー起動完了: http://localhost:${PORT}`);
  console.log(`====================================================`);
  
  // 自動的にブラウザでページを開く
  const url = `http://localhost:${PORT}`;
  const startCmd = process.platform === 'win32' ? `start ${url}` : `open ${url}`;
  exec(startCmd, (err) => {
    if (err) {
      console.log(`ブラウザの自動起動に失敗しました。お手数ですがブラウザで ${url} に直接アクセスしてください。`);
    }
  });
});
