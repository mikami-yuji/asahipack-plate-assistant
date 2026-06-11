// 共通で使用する型定義

// Excelからパースされた版データの1レコードを表す型
export type PlateRecord = {
  clientCode: string;
  clientName: string;
  supplierCode: string;
  supplierName: string;
  plateNo: string;
  orderNo: string;
  orderSuffix: string;
  orderSub: string;
  weight: string | number;
  finish: string;
  storeName: string;
  brandName: string;
  colorCount: number;
  colors: string;
  lastUsedDate: string;
  expiryDate: string;
  profit: number;
};

// 得意先ごとのグループを表す型
export type ClientGroup = {
  clientCode: string;
  clientName: string;
  records: PlateRecord[];
};

// API処理リクエストのパラメータ型
export type ProcessRequestData = {
  imageSrcDir: string;  // 入力画像フォルダ (e.g. \\asahipack01\画像)
  outputDestDir: string; // 出力先フォルダ (e.g. デスクトップ下の特定フォルダ)
  excelData: string;    // Base64エンコードされたExcelファイルデータ
};

// API処理の個別の進捗や結果を表す型
export type ProcessingLogItem = {
  type: 'info' | 'success' | 'warning' | 'error';
  message: string;
  timestamp: string;
};

// API処理レスポンスの型
export type ProcessResponseData = {
  success: boolean;
  message: string;
  processedClientsCount: number;
  totalImagesCopied: number;
  outputDir: string;
  logs: ProcessingLogItem[];
  clientGroups?: ClientGroup[];
};
