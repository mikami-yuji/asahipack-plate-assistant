import * as xlsx from 'xlsx';
import ExcelJS from 'exceljs';
import * as JapaneseHolidays from 'japanese-holidays';
import { PlateRecord, ClientGroup } from '../types';

/**
 * Excelファイルから担当者名 (B2セル) を取得する
 * @param fileBuffer Excelファイルのバイナリデータ
 * @returns 担当者名 (見つからない場合は '担当者')
 */
export function parseStaffName(fileBuffer: Buffer): string {
  const workbook = xlsx.read(fileBuffer, { type: 'buffer' });
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const staffCell = worksheet['B2'];
  if (staffCell && staffCell.v) {
    const val = String(staffCell.v).trim();
    return val.replace(/^(担当者|担当)[：:]\s*/, '');
  }
  return '担当者';
}

/**
 * 指定された年月の最終営業日（土日祝除く）を YYYY/MM/DD 形式で取得する
 * @param year 年
 * @param month 月 (1-12)
 * @returns YYYY/MM/DD 形式の最終営業日
 */
export function getLastBusinessDay(year: number, month: number): string {
  const date = new Date(year, month, 0); // その月の最終日
  
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
 * アップロードされたExcelファイルのバイナリデータを解析し、レコードの配列に変換する
 * @param fileBuffer Excelファイルのバイナリデータ
 * @returns 解析されたレコードの配列
 */
export async function parseExcel(fileBuffer: Buffer): Promise<PlateRecord[]> {
  const workbook = xlsx.read(fileBuffer, { type: 'buffer' });
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];

  const range = xlsx.utils.decode_range(worksheet['!ref'] || 'A1:R100');
  const records: PlateRecord[] = [];

  // 5行目 (インデックス 4) からデータが始まる
  for (let r = 4; r <= range.e.r; r++) {
    const clientCodeCell = worksheet[xlsx.utils.encode_cell({ r, c: 0 })];
    if (!clientCodeCell || !clientCodeCell.v) {
      continue; // 得意先コードが空の行はスキップ
    }

    const clientCode = String(clientCodeCell.v).trim();
    const clientName = worksheet[xlsx.utils.encode_cell({ r, c: 1 })]?.v ? String(worksheet[xlsx.utils.encode_cell({ r, c: 1 })].v).trim() : '';
    const supplierCode = worksheet[xlsx.utils.encode_cell({ r, c: 2 })]?.v ? String(worksheet[xlsx.utils.encode_cell({ r, c: 2 })].v).trim() : '';
    const supplierName = worksheet[xlsx.utils.encode_cell({ r, c: 3 })]?.v ? String(worksheet[xlsx.utils.encode_cell({ r, c: 3 })].v).trim() : '';
    const plateNo = worksheet[xlsx.utils.encode_cell({ r, c: 4 })]?.v ? String(worksheet[xlsx.utils.encode_cell({ r, c: 4 })].v).trim() : '';
    
    // 受注No.と枝番は数値の場合があるため文字列に変換
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
      clientCode,
      clientName,
      supplierCode,
      supplierName,
      plateNo,
      orderNo,
      orderSuffix,
      orderSub,
      weight,
      finish,
      storeName,
      brandName,
      colorCount,
      colors,
      lastUsedDate,
      expiryDate,
      profit
    });
  }

  return records;
}

/**
 * 仕入先名から製品の「種別」を判定する
 * @param supplierName 元の仕入先名
 * @returns 判定された種別文字列
 */
export function getProductType(supplierName: string): string {
  if (!supplierName) {
    return '別注';
  }
  const name = supplierName.trim();
  if (name.includes('シルク印刷')) {
    return 'シルク印刷';
  }
  if (name.includes('３Ｆロール印刷')) {
    return 'SP';
  }
  if (name.includes('オクダ') || name.includes('エイト')) {
    return 'オフセット';
  }
  if (name.includes('３Ｆロールカット')) {
    return 'カット';
  }
  if (name.includes('アサヒパック')) {
    return 'アサヒパック';
  }
  return '別注';
}

/**
 * 得意先ごとの情報から、落版確認用Excelファイルを生成する
 * @param group 得意先のレコードグループ
 * @param staffName 担当者名
 * @returns 生成されたExcelファイルのBuffer
 */
export async function generateClientExcel(group: ClientGroup, staffName: string): Promise<Buffer> {
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

  // グリッド線を表示する
  worksheet.views = [{ showGridLines: true }];

  // 列幅の設定 (A〜N列)
  worksheet.columns = [
    { key: 'orderNo', width: 12 },      // A: 受注No
    { key: 'orderSuffix', width: 4 },   // B: 枝番1 (-)
    { key: 'orderSub', width: 6 },      // C: 枝番2 (数値)
    { key: 'weight', width: 8 },        // D: Kg
    { key: 'finish', width: 8 },        // E: 仕上
    { key: 'productType', width: 12 },   // F: 種別
    { key: 'storeName', width: 15 },     // G: 店名
    { key: 'brandName', width: 25 },     // H: 銘柄
    { key: 'colorCount', width: 8 },    // I: 色数
    { key: 'colors', width: 25 },       // J: 色
    { key: 'lastUsedDate', width: 14 }, // K: 最終使用日
    { key: 'expiryDate', width: 14 },   // L: 落版予定日
    { key: 'answer', width: 12 },       // M: 回答 (廃棄/継続)
    { key: 'reason', width: 30 }        // N: 継続理由
  ];

  // 宛名と作成情報の記述
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

  // タイトル
  worksheet.mergeCells('A4:N4');
  const titleCell = worksheet.getCell('A4');
  titleCell.value = '落版候補リストのご確認について';
  titleCell.font = { name: 'MS Gothic', size: 16, bold: true, underline: true };
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
  worksheet.getRow(4).height = 30;

  // 説明文
  worksheet.mergeCells('A5:N5');
  const descCell = worksheet.getCell('A5');
  descCell.value = 'いつも大変お世話になっております。弊社で保管しております御社の版につきまして、最終使用日より期間が経過したものを落版（廃棄）候補としてリストアップいたしました。お手数ですが、落版の可否（廃棄/継続）をご記入の上、ご返送いただけますようお願い申し上げます。';
  descCell.font = { name: 'MS Gothic', size: 10 };
  descCell.alignment = { wrapText: true, vertical: 'top' };
  worksheet.getRow(5).height = 45;

  // 空行を挟んで表の開始

  // ヘッダー行 (7行目)
  const headerRowNumber = 7;
  const headerRow = worksheet.getRow(headerRowNumber);
  headerRow.height = 25;

  // 受注No.の列を結合
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
  
  // 回答・理由欄 (背景色を薄い黄色にして目立たせる)
  const answerHeaderCell = worksheet.getCell(`M${headerRowNumber}`);
  answerHeaderCell.value = '回答（必須）';
  
  const reasonHeaderCell = worksheet.getCell(`N${headerRowNumber}`);
  reasonHeaderCell.value = '継続の場合の理由';

  // ヘッダーのスタイル適用
  const headerCols = ['A', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N'];
  headerCols.forEach((col) => {
    const cell = worksheet.getCell(`${col}${headerRowNumber}`);
    cell.font = { name: 'MS Gothic', size: 10, bold: true };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = {
      top: { style: 'medium' },
      left: { style: 'thin' },
      bottom: { style: 'medium' },
      right: { style: 'thin' }
    };
    
    // M, N列は回答エリアとして薄い黄色にする
    if (col === 'M' || col === 'N') {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFF0D0' } // 薄いオレンジ/黄色
      };
    } else {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE6F2FF' } // 薄いブルー
      };
    }
  });



  // 残りのB, Cのヘッダー境界も設定
  ['B', 'C'].forEach((col) => {
    const cell = worksheet.getCell(`${col}${headerRowNumber}`);
    cell.border = {
      top: { style: 'medium' },
      bottom: { style: 'medium' },
      left: col === 'B' ? undefined : { style: 'thin' },
      right: col === 'C' ? { style: 'thin' } : undefined
    };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFE6F2FF' }
    };
  });

  // データ行の書き込み
  let currentRow = 8;
  group.records.forEach((record) => {
    const row = worksheet.getRow(currentRow);
    row.height = 20;

    const productType = getProductType(record.supplierName);
    let displayExpiryDate = record.expiryDate;
    
    // SP、シルク印刷、オフセットのみ落版予定日を当月（expiryDateの月）の最終営業日に変更 (土日祝除く)
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
    worksheet.getCell(`M${currentRow}`).value = ''; // 初期値は空
    worksheet.getCell(`N${currentRow}`).value = ''; // 初期値は空

    // データセルのアライメントと罫線
    const cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N'];
    cols.forEach((col) => {
      const cell = worksheet.getCell(`${col}${currentRow}`);
      cell.font = { name: 'MS Gothic', size: 10 };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };

      // 配置
      if (['A', 'B', 'C', 'D', 'F', 'I', 'K', 'L', 'M'].includes(col)) {
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      } else {
        cell.alignment = { horizontal: 'left', vertical: 'middle' };
      }

      // 回答セル (M列) にデータ入力規則 (廃棄 / 継続) を設定
      if (col === 'M') {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFFFDF0' } // 入力エリアとしての超薄黄色
        };
        // データ入力規則 (プルダウン)
        cell.dataValidation = {
          type: 'list',
          allowBlank: true,
          formulae: ['"廃棄,継続"'],
          showErrorMessage: true,
          errorTitle: '入力エラー',
          error: 'リストから「廃棄」または「継続」を選択してください。'
        };
      }
      
      // 理由セル (N列) も薄い黄色にする
      if (col === 'N') {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFFFDF0' }
        };
      }
    });

    currentRow++;
  });

  // バッファとしてエクスポート
  const excelBuffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(excelBuffer);
}
