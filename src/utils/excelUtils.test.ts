import ExcelJS from 'exceljs';
import { generateClientExcel, getProductType, parseStaffName, getLastBusinessDay } from './excelUtils';
import { ClientGroup } from '../types';

describe('excelUtils', (): void => {
  describe('getLastBusinessDay', (): void => {
    test('should return correct last business day of the month', (): void => {
      // 2026年6月30日(火) -> 平日なので 2026/06/30
      expect(getLastBusinessDay(2026, 6)).toBe('2026/06/30');

      // 2026年5月31日(日), 30日(土) -> 金曜日の 2026/05/29
      expect(getLastBusinessDay(2026, 5)).toBe('2026/05/29');
    });
  });
  describe('parseStaffName', (): void => {
    test('should correctly extract staff name from cell B2', async (): Promise<void> => {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Sheet1');
      worksheet.getCell('B2').value = '担当者：佐藤';
      
      const buffer = await workbook.xlsx.writeBuffer();
      const staffName = parseStaffName(Buffer.from(buffer));
      
      expect(staffName).toBe('佐藤');
    });

    test('should return default value if B2 is empty', async (): Promise<void> => {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Sheet1');
      
      const buffer = await workbook.xlsx.writeBuffer();
      const staffName = parseStaffName(Buffer.from(buffer));
      
      expect(staffName).toBe('担当者');
    });
  });
  describe('getProductType', (): void => {
    test('should correctly classify supplier names into product types', (): void => {
      expect(getProductType('シルク印刷')).toBe('シルク印刷');
      expect(getProductType('３Ｆロール印刷')).toBe('SP');
      expect(getProductType('オクダ印刷')).toBe('オフセット');
      expect(getProductType('（株）エイト')).toBe('オフセット');
      expect(getProductType('３Ｆロールカット')).toBe('カット');
      expect(getProductType('（株）アサヒパック')).toBe('アサヒパック');
      expect(getProductType('（株）コバヤシ')).toBe('別注');
      expect(getProductType('')).toBe('別注');
    });
  });

  describe('generateClientExcel', (): void => {
    test('should generate a beautiful Excel buffer with validations and data', async (): Promise<void> => {
      // テストデータの準備
      const mockGroup: ClientGroup = {
        clientCode: '0027099',
        clientName: '（株）みどりフーズ',
        records: [
          {
            clientCode: '0027099',
            clientName: '（株）みどりフーズ',
            supplierCode: '0000281',
            supplierName: '（株）コバヤシ',
            plateNo: '01243939',
            orderNo: '1243939',
            orderSuffix: '-',
            orderSub: '1',
            weight: 5,
            finish: 'RA',
            storeName: 'みどり',
            brandName: 'お米',
            colorCount: 5,
            colors: '墨 マゼンダ イエロー シアン 白',
            lastUsedDate: '2025/02/14',
            expiryDate: '2026/08/14',
            profit: -34036
          },
          {
            clientCode: '0027099',
            clientName: '（株）みどりフーズ',
            supplierCode: '0000705',
            supplierName: '３Ｆロール印刷',
            plateNo: '01106750',
            orderNo: '1106750',
            orderSuffix: '-',
            orderSub: '25',
            weight: 5,
            finish: 'RA',
            storeName: 'みどり',
            brandName: 'お米',
            colorCount: 1,
            colors: '墨',
            lastUsedDate: '2024/06/21',
            expiryDate: '2026/06/21',
            profit: 527770
          }
        ]
      };

      const staffName = '見上';

      // エクセル生成の実行
      const excelBuffer = await generateClientExcel(mockGroup, staffName);
      
      // 生成されたエクセルバッファをロードして内容を検証
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(excelBuffer);
      const worksheet = workbook.getWorksheet('落版確認書');

      expect(worksheet).toBeDefined();

      // 1. 宛先の検証
      const clientCell = worksheet!.getCell('A1');
      expect(clientCell.value).toBe('（株）みどりフーズ 御中');

      // 2. 担当者の検証
      const companyCell = worksheet!.getCell('K1');
      expect(companyCell.value).toBe('株式会社アサヒパック');

      const staffCell = worksheet!.getCell('K3');
      expect(staffCell.value).toBe('担当者: 見上');

      // 3. データ行の検証
      // ヘッダーが7行目、最初のデータ行は8行目
      const orderNoCell = worksheet!.getCell('A8');
      expect(orderNoCell.value).toBe('1243939');

      // 新設された種別列の検証
      const typeCell = worksheet!.getCell('F8');
      expect(typeCell.value).toBe('別注'); // 「（株）コバヤシ」はその他に該当するため「別注」

      const brandCell = worksheet!.getCell('H8');
      expect(brandCell.value).toBe('お米');

      // 1行目は「別注」なので落版予定日は元のまま
      const expiryCell8 = worksheet!.getCell('L8');
      expect(expiryCell8.value).toBe('2026/08/14');

      // 2行目の検証 (SPタイプなので落版予定日がシフトされている)
      const orderNoCell9 = worksheet!.getCell('A9');
      expect(orderNoCell9.value).toBe('1106750');

      const typeCell9 = worksheet!.getCell('F9');
      expect(typeCell9.value).toBe('SP');

      const expiryCell9 = worksheet!.getCell('L9');
      expect(expiryCell9.value).toBe('2026/06/30'); // 2026年6月の最終営業日

      // 4. データバリデーションの検証 (M8セル)
      const answerCell = worksheet!.getCell('M8');
      expect(answerCell.dataValidation).toBeDefined();
      expect(answerCell.dataValidation!.type).toBe('list');
      expect(answerCell.dataValidation!.formulae).toEqual(['"廃棄,継続"']);
    });
  });
});
