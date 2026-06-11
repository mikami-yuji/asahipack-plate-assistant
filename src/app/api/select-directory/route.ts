import { NextRequest, NextResponse } from 'next/server';
import { exec } from 'child_process';

/**
 * Windowsのフォルダ選択ダイアログを開くためのPOSTハンドラー
 * @param request HTTPリクエスト
 * @returns 選択されたフォルダパスを含むレスポンス
 */
export async function POST(request: NextRequest): Promise<NextResponse> {
  return new Promise<NextResponse>((resolve) => {
    const psCommand = `powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -sta -Command "Add-Type -AssemblyName System.Windows.Forms; $form = New-Object System.Windows.Forms.Form; $form.TopMost = $true; $form.Width = 0; $form.Height = 0; $form.WindowState = 'Minimized'; $dialog = New-Object System.Windows.Forms.FolderBrowserDialog; $dialog.Description = 'フォルダを選択してください'; $dialog.ShowNewFolderButton = $true; if ($dialog.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) { [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($dialog.SelectedPath)) }"`;

    exec(psCommand, (err, stdout, stderr) => {
      if (err) {
        resolve(NextResponse.json({ success: false, error: err.message }, { status: 500 }));
        return;
      }
      const base64Path = stdout.trim();
      let selectedPath = '';
      if (base64Path) {
        selectedPath = Buffer.from(base64Path, 'base64').toString('utf8');
      }
      resolve(NextResponse.json({ success: true, selectedPath }));
    });
  });
}
