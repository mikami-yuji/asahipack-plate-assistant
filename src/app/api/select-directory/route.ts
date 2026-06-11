import { NextRequest, NextResponse } from 'next/server';
import { exec } from 'child_process';

/**
 * Windowsのフォルダ選択ダイアログを開くためのPOSTハンドラー
 * @param request HTTPリクエスト
 * @returns 選択されたフォルダパスを含むレスポンス
 */
export async function POST(request: NextRequest): Promise<NextResponse> {
  return new Promise<NextResponse>((resolve) => {
    const psCommand = `powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -sta -Command "Add-Type -AssemblyName System.Windows.Forms; $dialog = New-Object System.Windows.Forms.FolderBrowserDialog; $dialog.Description = 'フォルダを選択してください'; $dialog.ShowNewFolderButton = $true; if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { Write-Output $dialog.SelectedPath }"`;

    exec(psCommand, (err, stdout, stderr) => {
      if (err) {
        resolve(NextResponse.json({ success: false, error: err.message }, { status: 500 }));
        return;
      }
      const selectedPath = stdout.trim();
      resolve(NextResponse.json({ success: true, selectedPath }));
    });
  });
}
