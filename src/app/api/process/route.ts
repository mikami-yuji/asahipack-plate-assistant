import { NextRequest, NextResponse } from 'next/server';
import { promises as fs } from 'fs';
import * as path from 'path';
import * as os from 'os';
import { parseExcel, generateClientExcel } from '../../../utils/excelUtils';
import { ProcessRequestData, ProcessResponseData, ClientGroup, PlateRecord, ProcessingLogItem } from '../../../types';

/**
 * 日時文字列から安全なディレクトリ名用の文字列を生成する (YYYYMMDD_HHMMSS)
 * @returns フォーマットされた日時文字列
 */
function getTimestampString(): string {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const date = String(now.getDate()).padStart(2, '0');
  const hours = String(now.getHours()).padStart(2, '0');
  const minutes = String(now.getMinutes()).padStart(2, '0');
  const seconds = String(now.getSeconds()).padStart(2, '0');
  return `${year}${month}${date}_${hours}${minutes}${seconds}`;
}

/**
 * 共有フォルダやローカルディレクトリから、指定された受注No.に対応する画像ファイルをスキャンする
 * @param srcDir 画像の取得元ディレクトリパス
 * @param orderNo 検索対象の受注No
 * @returns マッチした画像ファイル名の配列
 */
async function scanImages(srcDir: string, orderNo: string): Promise<string[]> {
  if (!orderNo) {
    return [];
  }
  try {
    const files = await fs.readdir(srcDir);
    // 受注No.から始まる、または受注No.を含む画像ファイルを探す (大文字小文字を区別しない)
    const lowerOrderNo = orderNo.toLowerCase();
    return files.filter((file) => {
      const lowerFile = file.toLowerCase();
      // orderNoで始まる、または orderNo_ , orderNo- で始まるものをマッチさせる
      return lowerFile.startsWith(lowerOrderNo);
    });
  } catch {
    // フォルダが見つからない等の場合は空配列を返す
    return [];
  }
}

/**
 * 落版リストを処理するメインAPIハンドラー
 * @param request HTTPリクエスト
 * @returns 処理結果を含むレスポンス
 */
export async function POST(request: NextRequest): Promise<NextResponse<ProcessResponseData>> {
  const logs: ProcessingLogItem[] = [];
  const addLog = (type: 'info' | 'success' | 'warning' | 'error', message: string): void => {
    logs.push({
      type,
      message,
      timestamp: new Date().toLocaleTimeString('ja-JP')
    });
  };

  try {
    const body: ProcessRequestData = await request.json();
    const { imageSrcDir, outputDestDir, excelData } = body;

    if (!excelData) {
      addLog('error', 'Excelデータが提供されていません。');
      return NextResponse.json(
        { success: false, message: 'Excel data is required', processedClientsCount: 0, totalImagesCopied: 0, outputDir: '', logs },
        { status: 400 }
      );
    }

    addLog('info', 'Excelデータの解析を開始します...');
    
    // Base64からバイナリバッファへ変換
    const buffer = Buffer.from(excelData.split(',')[1] || excelData, 'base64');
    let records: PlateRecord[];
    try {
      records = await parseExcel(buffer);
      addLog('success', `Excelの解析が完了しました。総レコード数: ${records.length}件`);
    } catch (err) {
      const errMsg = err instanceof Error ? err.message : String(err);
      addLog('error', `Excelのパースに失敗しました: ${errMsg}`);
      throw new Error(`Excel parse failed: ${errMsg}`);
    }

    if (records.length === 0) {
      addLog('warning', '有効な落版データが見つかりませんでした。得意先コードが入力されているか確認してください。');
      return NextResponse.json({
        success: true,
        message: 'No records found',
        processedClientsCount: 0,
        totalImagesCopied: 0,
        outputDir: '',
        logs
      });
    }

    // 担当者の抽出 (元Excelから担当者名を特定するか、デフォルト名)
    // 今回はリストの元データが「見上」さんなので、デフォルトは「見上」とする
    const staffName = '見上'; 

    // 得意先ごとにグループ化
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

    addLog('info', `得意先の仕分けが完了しました。得意先数: ${clientGroups.length}社`);

    // 出力先フォルダの決定
    const defaultDesktop = path.join(os.homedir(), 'Desktop');
    const rootOutputDir = outputDestDir ? outputDestDir : path.join(defaultDesktop, `落版連絡_${getTimestampString()}`);

    try {
      await fs.mkdir(rootOutputDir, { recursive: true });
      addLog('info', `出力先フォルダを作成しました: ${rootOutputDir}`);
    } catch (err) {
      const errMsg = err instanceof Error ? err.message : String(err);
      addLog('error', `出力先フォルダの作成に失敗しました: ${errMsg}`);
      throw new Error(`Failed to create output directory: ${errMsg}`);
    }

    // 画像フォルダの存在確認
    let isSrcDirAccessible = true;
    try {
      await fs.access(imageSrcDir);
      addLog('info', `共有画像フォルダに接続しました: ${imageSrcDir}`);
    } catch {
      isSrcDirAccessible = false;
      addLog('warning', `共有画像フォルダ ${imageSrcDir} にアクセスできません。画像コピーはスキップし、Excelの生成のみ実行します。`);
    }

    let totalImagesCopied = 0;
    let processedClientsCount = 0;

    // 各得意先ごとに処理を実行
    for (const group of clientGroups) {
      // フォルダ名から特殊文字を除外 (Windowsのフォルダ名に使えない文字対策)
      const safeClientName = group.clientName.replace(/[\\/:*?"<>|]/g, '_').trim();
      const clientDirName = `${group.clientCode}_${safeClientName}様`;
      const clientFolder = path.join(rootOutputDir, clientDirName);

      try {
        // 得意先用フォルダの作成
        await fs.mkdir(clientFolder, { recursive: true });
        
        // 1. 得意先用Excelの生成と保存
        const excelBuffer = await generateClientExcel(group, staffName);
        const excelPath = path.join(clientFolder, `落版確認書_${safeClientName}.xlsx`);
        await fs.writeFile(excelPath, excelBuffer);
        
        addLog('info', `[${safeClientName}] 得意先フォルダと確認書Excelを作成しました。`);
        
        // 2. 画像の検索とコピー
        let clientImagesCount = 0;
        if (isSrcDirAccessible) {
          // 重複した受注No.でのスキャンを防ぐための一時セット
          const processedOrderNos = new Set<string>();
          
          for (const record of group.records) {
            const orderNo = record.orderNo;
            if (!orderNo || processedOrderNos.has(orderNo)) {
              continue;
            }
            processedOrderNos.add(orderNo);

            const matchedFiles = await scanImages(imageSrcDir, orderNo);
            
            for (const file of matchedFiles) {
              const srcFilePath = path.join(imageSrcDir, file);
              const destFilePath = path.join(clientFolder, file);
              try {
                await fs.copyFile(srcFilePath, destFilePath);
                clientImagesCount++;
                totalImagesCopied++;
              } catch (copyErr) {
                const errMsg = copyErr instanceof Error ? copyErr.message : String(copyErr);
                addLog('warning', `[${safeClientName}] 画像のコピーに失敗しました (${file}): ${errMsg}`);
              }
            }
          }
          
          if (clientImagesCount > 0) {
            addLog('success', `[${safeClientName}] 画像を ${clientImagesCount} 件コピーしました。`);
          } else {
            addLog('info', `[${safeClientName}] 対応する画像は見つかりませんでした。`);
          }
        }

        processedClientsCount++;
      } catch (err) {
        const errMsg = err instanceof Error ? err.message : String(err);
        addLog('error', `[${group.clientName}] の処理中にエラーが発生しました: ${errMsg}`);
      }
    }

    addLog('success', `すべての処理が完了しました！ 得意先: ${processedClientsCount}社、コピーした画像: ${totalImagesCopied}件`);

    return NextResponse.json({
      success: true,
      message: 'Processing completed successfully',
      processedClientsCount,
      totalImagesCopied,
      outputDir: rootOutputDir,
      logs,
      clientGroups // クライアント一覧とレコード情報を返す
    });

  } catch (globalErr) {
    const errMsg = globalErr instanceof Error ? globalErr.message : String(globalErr);
    addLog('error', `システム例外が発生しました: ${errMsg}`);
    return NextResponse.json(
      {
        success: false,
        message: errMsg,
        processedClientsCount: 0,
        totalImagesCopied: 0,
        outputDir: '',
        logs
      },
      { status: 500 }
    );
  }
}
