import { ClientGroup } from '../types';

/**
 * 得意先ごとの落版連絡メールの下書き文面を生成する
 * @param group 得意先のレコードグループ
 * @param staffName 弊社担当者名
 * @returns 生成されたメールの件名と本文オブジェクト
 */
export function generateEmailDraft(group: ClientGroup, staffName: string): { subject: string; body: string } {
  const clientName = group.clientName;
  const plateCount = group.records.length;
  
  // 返答期日を現在から2週間後に設定
  const deadlineDate = new Date();
  deadlineDate.setDate(deadlineDate.getDate() + 14);
  const deadlineString = `${deadlineDate.getFullYear()}年${deadlineDate.getMonth() + 1}月${deadlineDate.getDate()}日`;

  const subject = `【落版（版廃棄）のご確認】御社製品用版の保管につきまして（${clientName}様）`;
  
  const body = `${clientName}様
いつもお世話になっております。
アサヒパック株式会社の${staffName}です。

平素は格別のご愛顧を賜り、心より御礼申し上げます。

さて、弊社にて大切に保管させていただいております御社のパッケージ用「版」につきまして、長期間（最終ご使用日より1年以上）経過したものを整理させていただきたく存じます。

つきましては、落版（廃棄）の候補となる版（合計 ${plateCount} 件）のリスト「落版確認書.xlsx」および、製品のデザイン画像をお送りいたします。

大変お手数ではございますが、添付のExcelファイルを開いていただき、各項目について「廃棄」または「継続」のご意向を回答欄にてご選択いただけますでしょうか。

勝手ながら、弊社での保管スペースの都合もあり、ご回答は 【 ${deadlineString} 】 までにご返送いただけますよう、ご協力をお願い申し上げます。

お忙しいところ誠に恐縮ではございますが、ご確認とご回答のほど、何卒よろしくお願い申し上げます。

--------------------------------------------------
アサヒパック株式会社
担当：${staffName}
--------------------------------------------------`;

  return { subject, body };
}
export type EmailDraft = { subject: string; body: string };
