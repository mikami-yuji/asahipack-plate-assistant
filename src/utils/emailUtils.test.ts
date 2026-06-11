import { generateEmailDraft } from './emailUtils';
import { ClientGroup } from '../types';

describe('emailUtils', (): void => {
  describe('generateEmailDraft', (): void => {
    test('should correctly generate email draft subject and body with client details', (): void => {
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
            storeName: '',
            brandName: '',
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
            storeName: 'みどりフーズ',
            brandName: '京丹波のお米',
            colorCount: 1,
            colors: '墨',
            lastUsedDate: '2024/06/21',
            expiryDate: '2026/06/21',
            profit: 527770
          }
        ]
      };

      const staffName = '見上';
      
      // テスト対象関数の実行
      const draft = generateEmailDraft(mockGroup, staffName);

      // 検証
      // 1. 件名に顧客名が含まれること
      expect(draft.subject).toContain('（株）みどりフーズ');
      expect(draft.subject).toContain('【落版（版廃棄）のご確認】');

      // 2. 本文に顧客名、担当者名、版の合計件数が含まれること
      expect(draft.body).toContain('（株）みどりフーズ様');
      expect(draft.body).toContain('担当：見上');
      expect(draft.body).toContain('合計 2 件');
      
      // 3. 期日の年が含まれること (今年は2026年だが、現在日付から14日後なので年度は入るはず)
      const expectedYear = new Date().getFullYear();
      expect(draft.body).toContain(`${expectedYear}年`);
    });
  });
});
