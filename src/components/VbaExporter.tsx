import React, { useState } from 'react';
import { 
  Download, 
  Copy, 
  Check, 
  FileSpreadsheet, 
  FileText, 
  Layers, 
  CheckSquare, 
  Square 
} from 'lucide-react';
import { VbaSample } from '../types';

interface VbaExporterProps {
  filteredSamples: VbaSample[];
  allSamples: VbaSample[];
}

export const VbaExporter: React.FC<VbaExporterProps> = ({ filteredSamples, allSamples }) => {
  const [exportScope, setExportScope] = useState<'filtered' | 'all'>('filtered');
  const [copyingText, setCopyingText] = useState(false);
  const [downloadingCSV, setDownloadingCSV] = useState(false);
  const [downloadingTXT, setDownloadingTXT] = useState(false);

  const targetSamples = exportScope === 'filtered' ? filteredSamples : allSamples;

  // CSV変換処理（Excel用BOM付きUTF-8）
  const convertToCSV = (samples: VbaSample[]): string => {
    const headers = ['ID', 'タイトル', 'カテゴリ', '難易度', '概要解説', 'タグ', 'VBAソースコード'];
    
    const escapeField = (val: string): string => {
      if (!val) return '';
      const str = String(val).replace(/"/g, '""');
      if (str.includes(',') || str.includes('\n') || str.includes('\r') || str.includes('"')) {
        return `"${str}"`;
      }
      return str;
    };

    const rows = samples.map(s => [
      s.id,
      s.title,
      s.category,
      s.difficulty,
      s.description + '\n\n【詳細解説】\n' + s.explanation.join('\n'),
      s.tags.join(','),
      s.code
    ]);

    const csvContent = [
      headers.map(escapeField).join(','),
      ...rows.map(row => row.map(escapeField).join(','))
    ].join('\r\n');

    return '\ufeff' + csvContent; // Excel用のBOM
  };

  // テキスト連結処理
  const convertToTXT = (samples: VbaSample[]): string => {
    let output = `========================================================================\n`;
    output += `■ Curated VBA Library Auto-Export Data\n`;
    output += `■ 生成日時: ${new Date().toLocaleString()}\n`;
    output += `■ 対象件数: ${samples.length} 件 (${exportScope === 'filtered' ? '検索・フィルタ適用中' : '全件一覧'})\n`;
    output += `========================================================================\n\n`;

    samples.forEach((s, idx) => {
      output += `------------------------------------------------------------------------\n`;
      output += `[マクロ No. ${idx + 1}] ${s.title}\n`;
      output += `------------------------------------------------------------------------\n`;
      output += `【ID】: ${s.id}\n`;
      output += `【カテゴリ】: ${s.category}  |  【難易度】: ${s.difficulty}\n`;
      output += `【タグ】: ${s.tags.join(', ')}\n`;
      output += `【機能説明】: ${s.description}\n\n`;
      
      output += `【解説・最適化方針】:\n`;
      s.explanation.forEach((exp, i) => {
        output += `  ${i + 1}. ${exp}\n`;
      });
      output += `\n`;

      output += `【VBAソースコード】:\n`;
      output += `${s.code}\n\n\n`;
    });

    output += `\n========================================================================\n`;
    output += `[EOF] End of VBA Code Transfer Document\n`;
    output += `========================================================================\n`;

    return output;
  };

  // ダウンロード実行ヘルパー
  const triggerDownload = (content: string, filename: string, mimeType: string) => {
    const blob = new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  // CSVダウンロード
  const handleDownloadCSV = () => {
    if (targetSamples.length === 0) return;
    setDownloadingCSV(true);
    const csvStr = convertToCSV(targetSamples);
    const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    const filename = `vba_library_${exportScope}_${dateStr}.csv`;
    triggerDownload(csvStr, filename, 'text/csv;charset=utf-8;');
    setTimeout(() => setDownloadingCSV(false), 1200);
  };

  // テキストダウンロード
  const handleDownloadTXT = () => {
    if (targetSamples.length === 0) return;
    setDownloadingTXT(true);
    const txtStr = convertToTXT(targetSamples);
    const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    const filename = `vba_code_consolidated_${exportScope}_${dateStr}.txt`;
    triggerDownload(txtStr, filename, 'text/plain;charset=utf-8;');
    setTimeout(() => setDownloadingTXT(false), 1200);
  };

  // クリップボード一括コピー
  const handleCopyConsolidated = () => {
    if (targetSamples.length === 0) return;
    setCopyingText(true);
    const txtStr = convertToTXT(targetSamples);
    navigator.clipboard.writeText(txtStr).then(() => {
      setTimeout(() => setCopyingText(false), 1500);
    }).catch(err => {
      console.error('一括コピーに失敗しました:', err);
      setCopyingText(false);
    });
  };

  return (
    <div className="p-5 rounded-2xl border border-stone-250/60 dark:border-stone-900 bg-[#FAF7F2]/90 dark:bg-[#121212]/80 backdrop-blur-md shadow-sm space-y-4 interact-glow font-sans select-none">
      <div className="flex items-center justify-between pb-1 border-b border-stone-200/40 dark:border-stone-900/40">
        <h3 className="text-[10px] font-bold uppercase tracking-widest text-[#8C8A85] dark:text-stone-500 flex items-center gap-2 font-mono">
          <Layers className="w-3.5 h-3.5 text-indigo-500" />
          EXPORT_DATA_CENTER
        </h3>
        <span className="text-[9px] font-mono font-bold bg-[#1A1A1A] text-[#FBF9F6] dark:bg-[#FAF7F2] dark:text-[#101010] px-1.5 py-0.5 rounded">
          CSV_TXT
        </span>
      </div>

      <div className="space-y-1.5">
        <div className="text-xs font-semibold text-stone-700 dark:text-stone-300">
          ライブラリの一括エクスポート
        </div>
        <p className="text-[10px] text-stone-500 dark:text-stone-400 leading-relaxed">
          複数のVBAマクロを連結し、ファイル保存やクリップボード蓄積を一瞬で行う高効率システム。
        </p>
      </div>

      {/* 対象切り替えトグル */}
      <div className="grid grid-cols-2 gap-1.5 bg-stone-100/70 dark:bg-stone-950/60 p-1.0 rounded-xl border border-stone-200 dark:border-stone-900/60">
        <button
          onClick={() => setExportScope('filtered')}
          className={`flex items-center justify-center gap-1.5 py-1.5 px-1 rounded-lg text-[10px] font-bold transition-all cursor-pointer ${
            exportScope === 'filtered'
              ? 'bg-white dark:bg-stone-850 text-stone-900 dark:text-white shadow-xs'
              : 'text-stone-500 dark:text-stone-400 hover:text-stone-800'
          }`}
        >
          <span>絞り込み中 ({filteredSamples.length}件)</span>
        </button>
        <button
          onClick={() => setExportScope('all')}
          className={`flex items-center justify-center gap-1.5 py-1.5 px-1 rounded-lg text-[10px] font-bold transition-all cursor-pointer ${
            exportScope === 'all'
              ? 'bg-white dark:bg-stone-850 text-stone-900 dark:text-white shadow-xs'
              : 'text-stone-500 dark:text-stone-400 hover:text-stone-800'
          }`}
        >
          <span>全件対象 ({allSamples.length}件)</span>
        </button>
      </div>

      {/* 操作ボタン軍 */}
      <div className="space-y-2 pt-1 font-sans">
        <button
          onClick={handleDownloadCSV}
          disabled={targetSamples.length === 0}
          className="w-full flex items-center justify-between px-3 py-2 text-xs font-semibold rounded-xl border border-stone-200 dark:border-stone-850/80 bg-white hover:bg-stone-50 dark:bg-[#181818] dark:hover:bg-stone-900 text-stone-850 dark:text-stone-200 transition-all cursor-pointer shadow-2xs hover:shadow-xs disabled:opacity-50 disabled:cursor-not-allowed group/btn"
        >
          <div className="flex items-center gap-2">
            <FileSpreadsheet className="w-4 h-4 text-emerald-600 dark:text-emerald-500 group-hover/btn:scale-110 transition-transform" />
            <span>Excel CSV形式でダウンロード</span>
          </div>
          {downloadingCSV ? (
            <Check className="w-3.5 h-3.5 text-emerald-500" />
          ) : (
            <Download className="w-3.5 h-3.5 text-stone-400 group-hover/btn:translate-y-0.5 transition-transform" />
          )}
        </button>

        <button
          onClick={handleDownloadTXT}
          disabled={targetSamples.length === 0}
          className="w-full flex items-center justify-between px-3 py-2 text-xs font-semibold rounded-xl border border-stone-200 dark:border-stone-850/80 bg-white hover:bg-stone-50 dark:bg-[#181818] dark:hover:bg-stone-900 text-stone-850 dark:text-stone-200 transition-all cursor-pointer shadow-2xs hover:shadow-xs disabled:opacity-50 disabled:cursor-not-allowed group/btn"
        >
          <div className="flex items-center gap-2">
            <FileText className="w-4 h-4 text-blue-650 dark:text-blue-400 group-hover/btn:scale-110 transition-transform" />
            <span>テキスト連結形式 (.txt)</span>
          </div>
          {downloadingTXT ? (
            <Check className="w-3.5 h-3.5 text-emerald-500" />
          ) : (
            <Download className="w-3.5 h-3.5 text-stone-400 group-hover/btn:translate-y-0.5 transition-transform" />
          )}
        </button>

        <button
          onClick={handleCopyConsolidated}
          disabled={targetSamples.length === 0}
          className="w-full flex items-center justify-between px-3 py-2.5 text-xs font-bold rounded-xl border border-stone-950/20 dark:border-[#FAF7F2]/10 bg-stone-950 hover:bg-stone-900 dark:bg-[#FAF7F2] dark:hover:bg-[#FAF7F2]/90 text-white dark:text-stone-950 transition-all cursor-pointer shadow-sm disabled:opacity-50 disabled:cursor-not-allowed group/btn"
        >
          <div className="flex items-center gap-2">
            <Copy className="w-4 h-4 group-hover/btn:scale-110 transition-transform" />
            <span>表示コードを一括コピー</span>
          </div>
          {copyingText ? (
            <span className="text-[10px] font-mono text-emerald-500 dark:text-emerald-600 font-bold">COPIED!</span>
          ) : (
            <span className="text-[10px] font-mono opacity-55">Ready</span>
          )}
        </button>
      </div>
    </div>
  );
};
