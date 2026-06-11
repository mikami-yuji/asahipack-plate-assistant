'use client';

import React, { useState, useRef, DragEvent, ChangeEvent } from 'react';
import { 
  Upload, 
  Folder, 
  Copy, 
  Check, 
  FileSpreadsheet, 
  AlertTriangle, 
  Info, 
  ArrowRight,
  FolderOpen
} from 'lucide-react';
import { ProcessingLogItem, ClientGroup, PlateRecord } from '../types';
import { generateEmailDraft } from '../utils/emailUtils';

export default function Home(): React.JSX.Element {
  // 状態定義
  const [imageSrcDir, setImageSrcDir] = useState<string>('\\\\asahipack01\\画像');
  const [outputDestDir, setOutputDestDir] = useState<string>('');
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [isDragging, setIsDragging] = useState<boolean>(false);
  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [progress, setProgress] = useState<number>(0);
  const [statusMessage, setStatusMessage] = useState<string>('');
  
  // APIレスポンスデータ
  const [logs, setLogs] = useState<ProcessingLogItem[]>([]);
  const [clientGroups, setClientGroups] = useState<ClientGroup[]>([]);
  const [selectedGroup, setSelectedGroup] = useState<ClientGroup | null>(null);
  const [outputDir, setOutputDir] = useState<string>('');
  
  // コピー状態管理
  const [subjectCopied, setSubjectCopied] = useState<boolean>(false);
  const [bodyCopied, setBodyCopied] = useState<boolean>(false);

  const fileInputRef = useRef<HTMLInputElement>(null);
  const logEndRef = useRef<HTMLDivElement>(null);

  // ファイルサイズ変換
  const formatFileSize = (bytes: number): string => {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  };

  // ドラッグ＆ドロップハンドラー
  const handleDragOver = (e: DragEvent<HTMLDivElement>): void => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = (e: DragEvent<HTMLDivElement>): void => {
    e.preventDefault();
    setIsDragging(false);
  };

  const handleDrop = (e: DragEvent<HTMLDivElement>): void => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const file = e.dataTransfer.files[0];
      if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
        setSelectedFile(file);
      } else {
        alert('Excelファイル (.xlsx, .xls) のみアップロード可能です。');
      }
    }
  };

  const handleFileChange = (e: ChangeEvent<HTMLInputElement>): void => {
    if (e.target.files && e.target.files.length > 0) {
      setSelectedFile(e.target.files[0]);
    }
  };

  const triggerFileSelect = (): void => {
    fileInputRef.current?.click();
  };

  // ファイルのBase64エンコード
  const fileToBase64 = (file: File): Promise<string> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.readAsDataURL(file);
      reader.onload = (): void => resolve(reader.result as string);
      reader.onerror = (error): void => reject(error);
    });
  };

  // 処理の実行
  const handleProcess = async (): Promise<void> => {
    if (!selectedFile) return;

    setIsProcessing(true);
    setProgress(10);
    setStatusMessage('Excelファイルを読み込んでいます...');
    setLogs([]);
    setClientGroups([]);
    setSelectedGroup(null);
    setOutputDir('');

    try {
      const base64Data = await fileToBase64(selectedFile);
      setProgress(30);
      setStatusMessage('データをサーバーへ送信し、仕分け処理を開始します...');

      const response = await fetch('/api/process', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          imageSrcDir,
          outputDestDir,
          excelData: base64Data
        }),
      });

      const result = await response.json();
      setProgress(90);
      setStatusMessage('処理結果を受け取っています...');

      if (response.ok && result.success) {
        setLogs(result.logs);
        setOutputDir(result.outputDir);
        
        // パースしたレコードから、もう一度クライアント側でグループ化を再現してUI表示用にする
        // APIから返されたログからパース情報を推測するか、
        // あるいはパース結果そのものをクライアント側でシミュレートして構築する。
        // ここではパース済みのExcelデータをBase64からクライアント側でも簡易パースして得意先一覧を再構築する
        // または、API側から得意先リストとレコード構造をそのまま送り返すように設計するのが最も綺麗。
        // route.ts で PlateRecord[] または ClientGroup[] もレスポンスに含めるように修正しよう。
        // 一旦、APIが成功したため、モック的にAPI結果から得意先を再抽出できるように、
        // API (route.ts) のレスポンススキーマを少し更新して、グループ情報も返せるようにしよう。
        // (一旦この処理が終わった後に、route.ts の戻り値に clientGroups も追加する)
      } else {
        setLogs(result.logs || [{ type: 'error', message: result.message || '不明なエラーが発生しました。', timestamp: new Date().toLocaleTimeString() }]);
      }
    } catch (err) {
      const errMsg = err instanceof Error ? err.message : String(err);
      setLogs((prev) => [
        ...prev,
        { type: 'error', message: `通信エラーが発生しました: ${errMsg}`, timestamp: new Date().toLocaleTimeString() }
      ]);
    } finally {
      setIsProcessing(false);
      setProgress(100);
      setStatusMessage('すべての処理が完了しました。');
    }
  };

  // メールテキストのコピー
  const copyToClipboard = async (text: string, type: 'subject' | 'body'): Promise<void> => {
    try {
      await navigator.clipboard.writeText(text);
      if (type === 'subject') {
        setSubjectCopied(true);
        setTimeout(() => setSubjectCopied(false), 2000);
      } else {
        setBodyCopied(true);
        setTimeout(() => setBodyCopied(false), 2000);
      }
    } catch {
      alert('クリップボードへのコピーに失敗しました。');
    }
  };

  // ログ自動スクロール
  React.useEffect(() => {
    logEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [logs]);

  // route.ts からのレスポンスに clientGroups も追加されたと仮定して、
  // 解析処理で取得できた得意先一覧を状態として管理する。
  // route.ts 側のレスポンスに `clientGroups` フィールドを含めるよう後で修正する。
  const handleProcessWrapper = async (): Promise<void> => {
    if (!selectedFile) return;

    setIsProcessing(true);
    setProgress(10);
    setStatusMessage('Excelファイルを読み込んでいます...');
    setLogs([]);
    setClientGroups([]);
    setSelectedGroup(null);
    setOutputDir('');

    try {
      const base64Data = await fileToBase64(selectedFile);
      setProgress(30);
      setStatusMessage('データをサーバーへ送信し、仕分け処理を開始します...');

      const response = await fetch('/api/process', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          imageSrcDir,
          outputDestDir,
          excelData: base64Data
        }),
      });

      const result = await response.json();
      setProgress(90);
      setStatusMessage('処理結果を受信しました。');

      setLogs(result.logs || []);
      
      if (response.ok && result.success) {
        setOutputDir(result.outputDir);
        // API側で仕分けされたグループデータをroute.tsから受け取る
        // (API route.ts にて groups も返却するように更新する)
        if (result.clientGroups) {
          setClientGroups(result.clientGroups);
          if (result.clientGroups.length > 0) {
            setSelectedGroup(result.clientGroups[0]);
          }
        }
      }
    } catch (err) {
      const errMsg = err instanceof Error ? err.message : String(err);
      setLogs((prev) => [
        ...prev,
        { type: 'error', message: `通信エラーが発生しました: ${errMsg}`, timestamp: new Date().toLocaleTimeString() }
      ]);
    } finally {
      setIsProcessing(false);
      setProgress(100);
      setStatusMessage('すべての処理が完了しました。');
    }
  };

  return (
    <div className="container">
      {/* ヘッダー */}
      <header className="header">
        <div className="header-title-container">
          <span className="logo-badge">AP-TOOL</span>
          <div>
            <h1>落版連絡フォルダ自動生成ツール</h1>
            <p className="header-subtitle">得意先ごとの確認書作成および画像仕分け業務を効率化します</p>
          </div>
        </div>
        <div className="output-info-content" style={{ textAlign: 'right' }}>
          <span style={{ fontSize: '0.8rem', color: 'var(--text-muted)' }}>アサヒパック株式会社</span>
        </div>
      </header>

      {/* メインレイアウト */}
      <div className="layout-grid">
        {/* 左カラム: 設定 & アップロード */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: '2rem' }}>
          {/* パス設定カード */}
          <section className="card">
            <h2 className="card-title">
              <FolderOpen size={18} /> 処理設定
            </h2>
            
            <div className="form-group">
              <label className="form-label">共有画像フォルダの場所 (画像取得元)</label>
              <input 
                type="text" 
                className="form-input" 
                value={imageSrcDir}
                onChange={(e) => setImageSrcDir(e.target.value)}
                placeholder="\\\\asahipack01\\画像"
                disabled={isProcessing}
              />
              <p className="form-input-help">※受注No.と一致する画像ファイルをスキャンするディレクトリです</p>
            </div>

            <div className="form-group">
              <label className="form-label">出力先フォルダの場所</label>
              <input 
                type="text" 
                className="form-input" 
                value={outputDestDir}
                onChange={(e) => setOutputDestDir(e.target.value)}
                placeholder="空欄の場合はデスクトップに新規作成します"
                disabled={isProcessing}
              />
              <p className="form-input-help">※得意先ごとの確認用Excelおよび仕分け画像を保存する先です</p>
            </div>
          </section>

          {/* アップロードカード */}
          <section className="card">
            <h2 className="card-title">
              <FileSpreadsheet size={18} /> 落版候補リストのアップロード
            </h2>

            <div 
              className={`upload-area ${isDragging ? 'dragging' : ''}`}
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
              onDrop={handleDrop}
              onClick={triggerFileSelect}
            >
              <input 
                type="file" 
                ref={fileInputRef} 
                onChange={handleFileChange}
                accept=".xlsx, .xls"
                style={{ display: 'none' }}
              />
              <div className="upload-icon">
                <Upload size={24} />
              </div>
              <p className="upload-text">ファイルをドラッグ＆ドロップするか、クリックして選択</p>
              <p className="upload-subtext">Excelファイル (.xlsx, .xls) のみ対応</p>
            </div>

            {selectedFile && (
              <div className="selected-file-info">
                <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', overflow: 'hidden' }}>
                  <FileSpreadsheet size={16} style={{ color: 'var(--color-success)', flexShrink: 0 }} />
                  <span className="file-name">{selectedFile.name}</span>
                  <span className="file-size">({formatFileSize(selectedFile.size)})</span>
                </div>
                <button 
                  className="btn-copy" 
                  style={{ border: 'none', background: 'transparent', padding: '0.25rem', color: 'var(--color-error)' }}
                  onClick={(e) => { e.stopPropagation(); setSelectedFile(null); }}
                  disabled={isProcessing}
                >
                  削除
                </button>
              </div>
            )}

            <div style={{ marginTop: '1.5rem' }}>
              <button 
                className="btn btn-primary"
                disabled={!selectedFile || isProcessing}
                onClick={handleProcessWrapper}
              >
                {isProcessing ? '処理を実行中...' : '落版仕分け・生成を開始'}
                {!isProcessing && <ArrowRight size={16} />}
              </button>
            </div>

            {/* 進捗プログレス */}
            {isProcessing && (
              <div className="progress-container">
                <div className="progress-bar-bg">
                  <div className="progress-bar-fill" style={{ width: `${progress}%` }}></div>
                </div>
                <div className="progress-text">
                  <span>{statusMessage}</span>
                  <span>{progress}%</span>
                </div>
              </div>
            )}
          </section>

          {/* ログ出力 */}
          {logs.length > 0 && (
            <section className="card">
              <h2 className="card-title">
                <Info size={18} /> 処理ログ
              </h2>
              <div className="log-viewer">
                {logs.map((log, index) => (
                  <div key={index} className="log-item">
                    <span className="log-time">[{log.timestamp}]</span>
                    <span className={`log-msg ${log.type}`}>{log.message}</span>
                  </div>
                ))}
                <div ref={logEndRef}></div>
              </div>
            </section>
          )}
        </div>

        {/* 右カラム: 得意先リスト & メール下書き */}
        <div>
          {clientGroups.length > 0 ? (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '2rem' }}>
              {/* 出力完了お知らせ */}
              {outputDir && (
                <div className="output-info-box">
                  <Folder className="output-info-icon" size={20} />
                  <div className="output-info-content">
                    <h4>顧客別フォルダを生成しました</h4>
                    <p>保存場所:</p>
                    <p className="output-path">{outputDir}</p>
                  </div>
                </div>
              )}

              {/* メール下書きカード */}
              <div className="card" style={{ paddingBottom: '1.5rem' }}>
                <h2 className="card-title">
                  <Folder size={18} /> 仕分け結果と連絡先一覧
                </h2>
                <p style={{ fontSize: '0.85rem', color: 'var(--text-secondary)', marginBottom: '1rem' }}>
                  得意先を選択すると、連絡用メールの下書きが表示されます。
                </p>

                <div className="result-section" style={{ border: 'none', paddingTop: 0, marginTop: 0 }}>
                  {/* 得意先リスト */}
                  <div className="client-list">
                    {clientGroups.map((group) => (
                      <div 
                        key={group.clientCode} 
                        className={`client-item ${selectedGroup?.clientCode === group.clientCode ? 'active' : ''}`}
                        onClick={() => setSelectedGroup(group)}
                      >
                        <div className="client-info">
                          <span className="client-name-text">{group.clientName}</span>
                          <span className="client-code-text">コード: {group.clientCode}</span>
                        </div>
                        <span className="client-badge success">{group.records.length} 件</span>
                      </div>
                    ))}
                  </div>

                  {/* メールドラフト表示 */}
                  {selectedGroup && (() => {
                    const draft = generateEmailDraft(selectedGroup, '見上');
                    return (
                      <div className="draft-container">
                        <div>
                          <div className="draft-header">
                            <span className="draft-title">件名</span>
                            <button 
                              className={`btn-copy ${subjectCopied ? 'success' : ''}`}
                              onClick={() => copyToClipboard(draft.subject, 'subject')}
                            >
                              {subjectCopied ? <Check size={12} /> : <Copy size={12} />}
                              {subjectCopied ? 'コピー完了' : '件名をコピー'}
                            </button>
                          </div>
                          <div className="draft-subject-box">
                            <span>{draft.subject}</span>
                          </div>
                        </div>

                        <div>
                          <div className="draft-header">
                            <span className="draft-title">本文の下書き</span>
                            <button 
                              className={`btn-copy ${bodyCopied ? 'success' : ''}`}
                              onClick={() => copyToClipboard(draft.body, 'body')}
                            >
                              {bodyCopied ? <Check size={12} /> : <Copy size={12} />}
                              {bodyCopied ? 'コピー完了' : '本文をコピー'}
                            </button>
                          </div>
                          <div className="draft-body-box">
                            {draft.body}
                          </div>
                        </div>
                      </div>
                    );
                  })()}
                </div>
              </div>
            </div>
          ) : (
            <div className="card" style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', height: '100%', minHeight: '350px', color: 'var(--text-muted)', textAlign: 'center' }}>
              <Info size={48} style={{ marginBottom: '1rem', strokeWidth: 1.5 }} />
              <h3>仕分け結果は未生成です</h3>
              <p style={{ fontSize: '0.85rem', maxWidth: '300px', marginTop: '0.5rem' }}>
                左側のパネルから「落版候補リスト」をアップロードして処理を実行してください。
              </p>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
