import React from 'react';
import './globals.css';

export const metadata = {
  title: '落版連絡ツール | アサヒパック株式会社',
  description: 'システムから出力された落版候補リストを各顧客用に自動仕分けするツールです。',
};

type RootLayoutProps = {
  children: React.ReactNode;
};

/**
 * アプリケーションのルートレイアウト
 * @param props children を含むプロップス
 * @returns ルートレイアウトのJSX
 */
export default function RootLayout({ children }: RootLayoutProps): React.JSX.Element {
  return (
    <html lang="ja">
      <body>{children}</body>
    </html>
  );
}
