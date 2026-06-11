import ExcelJS from 'exceljs';
import { generateClientExcel, getProductType } from './excelUtils';
import { ClientGroup } from '../types';

describe('excelUtils', (): void => {
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

      // 4. データバリデーションの検証 (M8セル)
      const answerCell = worksheet!.getCell('M8');
      expect(answerCell.dataValidation).toBeDefined();
      expect(answerCell.dataValidation!.type).toBe('list');
      expect(answerCell.dataValidation!.formulae).toEqual(['"廃棄,継続"']);
    });
  });
});
