import * as fs from 'fs/promises';
import * as path from 'path';
import { parseExcel, generateClientExcel, getProductType } from './src/utils/excelUtils';
import { PlateRecord, ClientGroup } from './src/types';

// 入力データおよび設定
const EXCEL_FILE_PATH = 'C:\\Users\\asahi\\Desktop\\フォーマット_JI020L_落版候補リスト_054_見上_20260531232354069134.xlsx';
const IMAGE_SRC_DIR = '\\\\asahipack01\\画像';
const OUTPUT_DEST_DIR = 'C:\\Users\\asahi\\Desktop\\テスト_種別追加';
const STAFF_NAME = '見上';

/**
 * 画像フォルダから受注Noにマッチする画像をスキャンする
 * @param srcDir 画像取得元
 * @param orderNo 受注No
 * @returns 見つかったファイル名の配列
 */
async function scanImages(srcDir: string, orderNo: string): Promise<string[]> {
  if (!orderNo) return [];
  try {
    const files = await fs.readdir(srcDir);
    const lowerOrderNo = orderNo.toLowerCase();
    return files.filter((file) => {
      const lowerFile = file.toLowerCase();
      return lowerFile.startsWith(lowerOrderNo);
    });
  } catch {
    return [];
  }
}

/**
 * テスト実行用メイン関数
 */
async function runTest(): Promise<void> {
  console.log('--- 落版仕分け処理テスト開始 ---');
  console.log(`入力Excel: ${EXCEL_FILE_PATH}`);
  console.log(`出力先: ${OUTPUT_DEST_DIR}`);
  console.log(`画像元: ${IMAGE_SRC_DIR}`);

  try {
    // 1. 入力Excelの読み込み
    const excelBuffer = await fs.readFile(EXCEL_FILE_PATH);
    const records = await parseExcel(excelBuffer);
    console.log(`Excelパース完了。総レコード数: ${records.length}件`);

    if (records.length === 0) {
      console.log('レコードが見つかりませんでした。');
      return;
    }

    // 2. 得意先ごとにグループ化
    const groupsMap = new Map<string, PlateRecord[]>();
    records.forEach((rec) => {
      const key = `${rec.clientCode}___${rec.clientName}`;
      if (!groupsMap.has(key)) {
        groupsMap.set(key, []);
      }
      groupsMap.get(key)!.push(rec);
    });

    const clientGroups: ClientGroup[] = Array.from(groupsMap.entries()).map(([key, recs]) => {
      const [clientCode, clientName] = key.split('___');
      return { clientCode, clientName, records: recs };
    });

    console.log(`得意先数: ${clientGroups.length}社`);

    // 3. 出力先フォルダの作成
    await fs.mkdir(OUTPUT_DEST_DIR, { recursive: true });

    // 画像フォルダへのアクセス確認
    let isSrcDirAccessible = true;
    try {
      await fs.access(IMAGE_SRC_DIR);
      console.log(`共有画像フォルダに接続成功: ${IMAGE_SRC_DIR}`);
    } catch {
      isSrcDirAccessible = false;
      console.warn(`[警告] 共有画像フォルダにアクセスできません: ${IMAGE_SRC_DIR} (画像コピーはスキップします)`);
    }

    // 各グループの処理
    let totalImagesCopied = 0;
    for (const group of clientGroups) {
      const safeClientName = group.clientName.replace(/[\\/:*?"<>|]/g, '_').trim();
      const clientFolder = path.join(OUTPUT_DEST_DIR, `${group.clientCode}_${safeClientName}様`);
      
      // 得意先フォルダ作成
      await fs.mkdir(clientFolder, { recursive: true });

      // 得意先用Excelの生成・保存
      const clientExcelBuffer = await generateClientExcel(group, STAFF_NAME);
      const excelPath = path.join(clientFolder, `落版確認書_${safeClientName}.xlsx`);
      await fs.writeFile(excelPath, clientExcelBuffer);
      
      const firstRec = group.records[0];
      const testType = getProductType(firstRec.supplierName);
      console.log(`[${safeClientName}] -> Excel作成完了: ${path.basename(excelPath)} (最初の仕入先: "${firstRec.supplierName}" -> 種別: "${testType}")`);

      // 画像コピー
      if (isSrcDirAccessible) {
        const processedOrderNos = new Set<string>();
        let clientImagesCount = 0;
        for (const record of group.records) {
          const orderNo = record.orderNo;
          if (!orderNo || processedOrderNos.has(orderNo)) continue;
          processedOrderNos.add(orderNo);

          const matchedFiles = await scanImages(IMAGE_SRC_DIR, orderNo);
          for (const file of matchedFiles) {
            const srcFilePath = path.join(IMAGE_SRC_DIR, file);
            const destFilePath = path.join(clientFolder, file);
            try {
              await fs.copyFile(srcFilePath, destFilePath);
              clientImagesCount++;
              totalImagesCopied++;
            } catch (err) {
              console.error(`[エラー] 画像コピー失敗 (${file}):`, err);
            }
          }
        }
        if (clientImagesCount > 0) {
          console.log(`[${safeClientName}] -> 画像コピー完了: ${clientImagesCount}件`);
        }
      }
    }

    console.log('\n--- 処理完了 ---');
    console.log(`出力先を確認してください: ${OUTPUT_DEST_DIR}`);
    console.log(`総コピー画像数: ${totalImagesCopied}件`);

  } catch (error) {
    console.error('テスト実行中に致命的なエラーが発生しました:', error);
  }
}

runTest();
