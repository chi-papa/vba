import React, { useState } from 'react';
import { Copy, Check } from 'lucide-react';

interface SyntaxHighlighterProps {
  code: string;
  language: 'vba' | 'html';
  id?: string;
}

export const SyntaxHighlighter: React.FC<SyntaxHighlighterProps> = ({ code, language, id }) => {
  const [copied, setCopied] = useState(false);

  const handleCopy = async () => {
    try {
      await navigator.clipboard.writeText(code);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch (err) {
      console.error('コピー失敗:', err);
    }
  };

  const highlightVBA = (rawCode: string) => {
    // 1. HTMLエスケープして直接記述
    let escaped = rawCode
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');

    // 一時退避用のTokenマップ。文字列やコメントが他の正規表現で破壊されるのを防ぐ。
    const tokens: { [key: string]: string } = {};
    let tokenIndex = 0;

    const addToken = (html: string) => {
      const tokenId = `___TOKEN_HG_${tokenIndex}___`;
      tokens[tokenId] = html;
      tokenIndex++;
      return tokenId;
    };

    // 1. VBA コメント (シングルクォートから行末) の退避
    escaped = escaped.replace(/('[^\r\n]*)/g, (match) => {
      return addToken(`<span style="text-shadow: 0 1px 2px rgba(0,0,0,0.9);" class="text-emerald-400 font-normal italic">${match}</span>`);
    });

    // 2. ダブルクォーテーションで囲まれた文字列 ("...") の退避
    escaped = escaped.replace(/("[^"\r\n]*")/g, (match) => {
      return addToken(`<span style="text-shadow: 0 1px 2px rgba(0,0,0,0.9);" class="text-amber-300 font-medium">${match}</span>`);
    });

    // 3. キーワードハイライト (単語境界 \\b を使用)
    const keywords = [
      'Sub', 'End', 'Function', 'Dim', 'As', 'If', 'Then', 'Next', 'For', 'To',
      'ByVal', 'With', 'Exit', 'On', 'Error', 'GoTo', 'Set', 'Nothing', 'True',
      'False', 'Else', 'Each', 'In', 'Select', 'Case', 'Const', 'Private', 'Public',
      'Call', 'Do', 'Loop', 'While', 'Until', 'Resume'
    ];
    
    const excelObjects = [
      'Application', 'ScreenUpdating', 'Calculation', 'xlCalculationManual', 'xlCalculationAutomatic',
      'EnableEvents', 'ActiveWorkbook', 'ActiveSheet', 'Cells', 'Rows', 'Columns', 'Range',
      'Workbook', 'Worksheet', 'UsedRange', 'SpecialCells', 'AutoFilter', 'AutoFilterMode', 'FileDialog',
      'SaveAs', 'Close', 'Rows', 'Count', 'Round', 'MsgBox', 'Err', 'Description', 'FSO', 'FileSystemObject',
      'Folder', 'File', 'GetFolder', 'CreateObject', 'Dir'
    ];

    const types = [
      'Long', 'String', 'Integer', 'Boolean', 'Variant', 'Double', 'Object', 'Single'
    ];

    // キーワード、オブジェクト、型
    const keywordRegex = new RegExp(`\\b(${keywords.join('|')})\\b`, 'gi');
    escaped = escaped.replace(keywordRegex, (match) => {
      return `<span style="text-shadow: 0 1px 2px rgba(0,0,0,0.9);" class="text-sky-300 font-bold">${match}</span>`;
    });

    const objRegex = new RegExp(`\\b(${excelObjects.join('|')})\\b`, 'gi');
    escaped = escaped.replace(objRegex, (match) => {
      return `<span style="text-shadow: 0 1px 2px rgba(0,0,0,0.9);" class="text-fuchsia-300 font-medium">${match}</span>`;
    });

    const typeRegex = new RegExp(`\\b(${types.join('|')})\\b`, 'g');
    escaped = escaped.replace(typeRegex, (match) => {
      return `<span style="text-shadow: 0 1px 2px rgba(0,0,0,0.9);" class="text-teal-300 font-medium">${match}</span>`;
    });

    // 4. 数値リテラルのハイライト
    escaped = escaped.replace(/\b(\d+)\b/g, (match) => {
      return `<span style="text-shadow: 0 1px 2px rgba(0,0,0,0.9);" class="text-indigo-300 font-mono">${match}</span>`;
    });

    // 退避したコメント・文字列を復元
    Object.keys(tokens).reverse().forEach((tokenId) => {
      escaped = escaped.replace(tokenId, tokens[tokenId]);
    });

    return escaped;
  };

  const highlightHTML = (rawCode: string) => {
    let escaped = rawCode
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');

    const tokens: { [key: string]: string } = {};
    let tokenIndex = 0;

    const addToken = (html: string) => {
      const tokenId = `___TOKEN_HTML_${tokenIndex}___`;
      tokens[tokenId] = html;
      tokenIndex++;
      return tokenId;
    };

    // 1. HTMLコメント ( &lt;!-- [...] --&gt; ) の退避
    escaped = escaped.replace(/(&lt;!--[\s\S]*?--&gt;)/g, (match) => {
      return addToken(`<span style="text-shadow: 0 1px 2px rgba(0,0,0,0.9);" class="text-emerald-400 italic">${match}</span>`);
    });

    // 2. ダブルクォーテーションで囲まれた文字列 ("...") の退避
    escaped = escaped.replace(/("[^"\r\n]*")/g, (match) => {
      return addToken(`<span style="text-shadow: 0 1px 2px rgba(0,0,0,0.9);" class="text-amber-300">${match}</span>`);
    });

    // 3. シングルクォーテーションで囲まれた文字列 ('...') の退避
    escaped = escaped.replace(/('[^'\r\n]*')/g, (match) => {
      return addToken(`<span style="text-shadow: 0 1px 2px rgba(0,0,0,0.9);" class="text-amber-300">${match}</span>`);
    });

    // 4. HTMLのタグ (&lt; / &gt;で囲まれている部分) の基本カラー、および属性
    // まず開始タグと終了タグのキーワード部分を調整
    escaped = escaped.replace(/(&lt;\/?[a-zA-Z0-9:-]+)/g, (match) => {
      return `<span style="text-shadow: 0 1px 2px rgba(0,0,0,0.9);" class="text-pink-400 font-medium">${match}</span>`;
    });

    escaped = escaped.replace(/(&gt;)/g, (match) => {
      return `<span style="text-shadow: 0 1px 2px rgba(0,0,0,0.9);" class="text-pink-400 font-semibold">${match}</span>`;
    });

    // 5. styleやonclickなどの属性名、変数の色
    const attributes = [
      'class', 'id', 'onclick', 'onchange', 'style', 'src', 'href', 'type', 'accept', 'lang', 'charset', 'name', 'content'
    ];
    attributes.forEach((attr) => {
      const attrRegex = new RegExp(`\\b(${attr})=`, 'gi');
      escaped = escaped.replace(attrRegex, (match, p1) => {
        return `<span style="text-shadow: 0 1px 2px rgba(0,0,0,0.9);" class="text-violet-300">${p1}</span>=`;
      });
    });

    // 各種JavaScriptキーワード、変数
    const jsKeywords = [
      'const', 'let', 'var', 'function', 'return', 'if', 'else', 'new', 'Set', 'Array', 'JSON', 'document', 'window', 'localStorage', 'sessionStorage', 'console'
    ];
    jsKeywords.forEach((kw) => {
      const kwRegex = new RegExp(`\\b(${kw})\\b`, 'g');
      escaped = escaped.replace(kwRegex, (match) => {
        return `<span style="text-shadow: 0 1px 2px rgba(0,0,0,0.9);" class="text-sky-300 font-bold">${match}</span>`;
      });
    });

    // 退避したコメント・文字列を復元
    Object.keys(tokens).reverse().forEach((tokenId) => {
      escaped = escaped.replace(tokenId, tokens[tokenId]);
    });

    return escaped;
  };

  const getHighlightedHtml = () => {
    if (language === 'vba') {
      return highlightVBA(code);
    }
    return highlightHTML(code);
  };

  // 行番号をカウントして配列化
  const lines = code.split('\n');

  return (
    <div className="relative group rounded-xl border border-stone-250/60 dark:border-stone-850 bg-[#0A0A0A] overflow-hidden font-mono text-sm shadow-md leading-relaxed select-none" id={id}>
      
      {/* ターミナル・Macウインドウ風ヘッダー */}
      <div className="flex items-center justify-between px-4 py-3 border-b border-stone-850 bg-[#121212] backdrop-blur-md">
        
        {/* 左側: ミニマルなデジタルドットインジゲーター + 言語ラベル */}
        <div className="flex items-center gap-3">
          <div className="flex gap-1.5 shrink-0 select-none">
            <span className="w-2 h-2 rounded-full bg-stone-700 dark:bg-stone-500 animate-pulse block" />
            <span className="w-2 h-2 rounded-full bg-stone-800 dark:bg-stone-700 block" />
          </div>
          <span className="text-[10px] font-bold uppercase tracking-widest text-stone-400 dark:text-stone-400 font-mono">
            {language === 'vba' ? 'VBA CORE MODULE' : 'HTML EMULATOR CODE'}
          </span>
        </div>

        {/* 右側: コピーボタン */}
        <button
          onClick={handleCopy}
          className="flex items-center gap-1.5 px-3 py-1.5 text-[10px] font-bold font-mono rounded-lg transition-all text-stone-300 hover:text-white bg-white/[0.04] hover:bg-white/[0.08] border border-white/[0.03] hover:border-white/[0.1] shadow-sm cursor-pointer"
          title="コードをクリップボードにコピー"
        >
          {copied ? (
            <>
              <Check className="w-3.5 h-3.5 text-emerald-400 animate-none" />
              <span className="text-emerald-400 font-bold tracking-wider">COPIED_OK</span>
            </>
          ) : (
            <>
              <Copy className="w-3.5 h-3.5" />
              <span className="tracking-wider">COPY_CODE</span>
            </>
          )}
        </button>
      </div>

      {/* コード表示セクション (行番号＋コード本体) */}
      <div className="flex overflow-x-auto max-h-[500px] xl:max-h-[600px] divide-x divide-white/[0.04] leading-relaxed text-stone-200 bg-[#080808] select-text">
        
        {/* 行番号表示 (選択不可) */}
        <div className="py-4 px-3 bg-[#0c0c0c] text-right select-none text-stone-600 dark:text-stone-600 min-w-[3.5rem] text-[11px] font-bold font-mono">
          {lines.map((_, idx) => (
            <div key={idx} className="h-5">
              {idx + 1}
            </div>
          ))}
        </div>

        {/* コード本体 */}
        <pre className="py-4 px-5 flex-1 overflow-x-auto select-text font-mono text-xs md:text-sm">
          <code
            dangerouslySetInnerHTML={{ __html: getHighlightedHtml() }}
            className="block whitespace-pre select-text"
          />
        </pre>
      </div>
    </div>
  );
};
