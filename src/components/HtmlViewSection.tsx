import React, { useState } from 'react';
import { htmlNotebookCode } from '../data/htmlSample';
import { SyntaxHighlighter } from './SyntaxHighlighter';
import { Play, Code, Star, CheckCircle, Info, Sparkles } from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';

export const HtmlViewSection: React.FC = () => {
  const [activeTab, setActiveTab] = useState<'preview' | 'code'>('preview');

  // 特徴と解説項目
  const highlights = [
    { title: 'バグ対応済み v13設計', desc: '循環参照による無限ループ・スタックオーバーフロー回避、オートセーブ機能やJSON破損時の初期化フローを完全カバー。' },
    { title: '4px グリッド & CSS変数で調和', desc: '細部のマージン、パディングに至るまで統一されたコンシスエントなUI設計。ダークモードの切り替えが完璧。' },
    { title: 'ツリー階層構造 ＆ フォルダ開閉', desc: '子ノートを作成し、ツリーの中に格納。サイドバーでのフォルダ折りたたみによるスマートなノート整理。' },
    { title: 'ローカル強固な仕組み', desc: 'localStorageを用いたオフラインファースト設計。他形式に依存せずJSON形式での安全な一括エクスポートと追加インポートに対応。' }
  ];

  return (
    <div className="space-y-8">
      {/* 導入バナー - 未来的なポータルガード */}
      <div className="p-6 rounded-2xl border border-stone-250/50 dark:border-stone-900 bg-[#FAF8F5]/70 dark:bg-[#121212]/85 backdrop-blur-md shadow-sm relative overflow-hidden group interact-glow">
        
        {/* 精密なコプロセッサー角の未来風ライン */}
        <div className="absolute top-0 left-0 w-3 h-[1px] bg-stone-300 dark:bg-stone-750 pointer-events-none" />
        <div className="absolute top-0 left-0 h-3 w-[1px] bg-stone-300 dark:bg-stone-750 pointer-events-none" />
        <div className="absolute top-0 right-0 w-3 h-[1px] bg-stone-300 dark:bg-stone-750 pointer-events-none" />
        <div className="absolute top-0 right-0 h-3 w-[1px] bg-stone-300 dark:bg-stone-750 pointer-events-none" />

        <div className="absolute right-0 top-0 w-48 h-48 bg-stone-900/[0.02] dark:bg-white/[0.01] rounded-full blur-3xl pointer-events-none" />
        
        <div className="flex gap-4 items-start relative z-10">
          <div className="p-3 bg-stone-950 dark:bg-stone-200 text-[#FBF9F6] dark:text-[#101010] rounded-xl shadow-lg shrink-0">
            <Info className="w-4.5 h-4.5" />
          </div>
          <div>
            <h3 className="text-lg font-bold text-stone-950 dark:text-white mb-2 font-display tracking-wider flex items-center gap-2">
              <span>Secure Notes</span>
              <span className="text-[9px] font-mono font-bold bg-[#1A1A1A] text-white dark:bg-stone-800 dark:text-stone-300 px-1.5 py-0.5 rounded border border-transparent dark:border-stone-700 tracking-widest uppercase">
                SANDBOXED_APP
              </span>
            </h3>
            <p className="text-stone-600 dark:text-stone-400 text-xs md:text-xs leading-relaxed mb-4 font-sans font-medium">
              こちらに統合されたHTMLコードは、一切の不透明な通信を排除し、ブラウザの完全隔離されたローカル環境のみで駆動するハイエンドな私的ノートブックです。極限設計された親子構造型ツリー、リアルタイムでのオートセーブ、スマートフォルダの折り畳み整理、堅牢なJSON形式でのバックアップおよびインポート。オフィス等で一切のインストールが不可能なセキュアな状況にあっても、コピペ１つで極上の管理環境を用意できます。
            </p>
            <div className="flex flex-wrap gap-2 select-none font-mono">
              <span className="text-[9px] font-bold bg-stone-100 text-stone-600 dark:bg-stone-900 dark:text-stone-400 px-3 py-1 rounded-lg border border-stone-200/50 dark:border-stone-800/80 uppercase tracking-widest">
                HTML5 + Pure CSS3 + JS (Vanilla)
              </span>
              <span className="text-[9px] font-bold bg-stone-100 text-stone-600 dark:bg-stone-900 dark:text-stone-400 px-3 py-1 rounded-lg border border-stone-200/50 dark:border-stone-800/80 uppercase tracking-widest">
                STANDALONE LOCAL INSTINTCT
              </span>
            </div>
          </div>
        </div>
      </div>

      {/* プレビュー・コード切り替えタブ - 未来型ラボインターフェース */}
      <div className="border border-stone-250/50 dark:border-stone-900 rounded-2xl bg-[#FAF8F5]/70 dark:bg-stone-950/80 backdrop-blur-md shadow-sm overflow-hidden interact-glow relative">
        
        {/* 装飾 */}
        <div className="absolute top-0 right-0 w-2.5 h-[1px] bg-stone-300 dark:bg-stone-800 pointer-events-none" />
        <div className="absolute top-0 right-0 h-2.5 w-[1px] bg-stone-300 dark:bg-stone-800 pointer-events-none" />

        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 px-5 py-4 border-b border-stone-200 dark:border-stone-900 bg-stone-100/[0.2] dark:bg-stone-900/[0.1] select-none">
          <div className="flex items-center gap-1 bg-stone-200/55 dark:bg-stone-900/60 p-1 rounded-xl max-w-max font-sans">
            <button
              onClick={() => setActiveTab('preview')}
              className={`flex items-center gap-1.5 px-4 py-1.5 text-xs font-bold rounded-lg transition-all cursor-pointer ${
                activeTab === 'preview'
                  ? 'bg-stone-950 text-white dark:bg-stone-200 dark:text-stone-950 shadow-sm font-bold'
                  : 'text-stone-500 hover:text-stone-950 dark:text-stone-400 dark:hover:text-white'
              }`}
            >
              <Play className="w-3 h-3 fill-current shrink-0" />
              LIVE EMULATION
            </button>
            <button
              onClick={() => setActiveTab('code')}
              className={`flex items-center gap-1.5 px-4 py-1.5 text-xs font-bold rounded-lg transition-all cursor-pointer ${
                activeTab === 'code'
                  ? 'bg-stone-950 text-white dark:bg-stone-200 dark:text-stone-950 shadow-sm font-bold'
                  : 'text-stone-500 hover:text-stone-950 dark:text-stone-400 dark:hover:text-white'
              }`}
            >
              <Code className="w-3.5 h-3.5 shrink-0" />
              SOURCE_VIEW
            </button>
          </div>

          <div className="text-[10px] text-stone-500 dark:text-stone-500 font-bold font-mono tracking-widest">
            {activeTab === 'preview' ? 'STATE: ACTIVE_SANDBOX_STABLE' : 'STATE: COMPILATION_READY_V13'}
          </div>
        </div>

        {/* タブの中身 */}
        <div className="p-5 bg-stone-100/[0.05] dark:bg-stone-950/[0.05]">
          <AnimatePresence mode="wait">
            {activeTab === 'preview' ? (
              <motion.div
                key="preview-tab"
                initial={{ opacity: 0, y: 5 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -5 }}
                transition={{ duration: 0.15 }}
                className="space-y-4"
              >
                <div className="rounded-xl border border-stone-200 dark:border-stone-900 overflow-hidden bg-[#FAF8F5] dark:bg-[#080808] shadow-inner relative">
                  {/* エミュレータのトップラウンダー */}
                  <div className="absolute top-0 left-0 right-0 h-1 bg-gradient-to-r from-stone-300 dark:from-stone-800 to-transparent pointer-events-none" />
                  <iframe
                    title="Secure Notes Demo"
                    srcDoc={htmlNotebookCode}
                    className="w-full h-[580px] border-0 bg-[#FAF8F5]"
                    sandbox="allow-scripts allow-modals allow-downloads"
                    referrerPolicy="no-referrer"
                  />
                </div>
                <div className="text-[10px] text-stone-500 dark:text-stone-400 leading-relaxed px-1 flex gap-2.5 items-start font-mono font-medium tracking-wider">
                  <Sparkles className="w-4 h-4 text-stone-400 dark:text-stone-600 shrink-0 mt-0.5 animate-none" />
                  <span>
                    // NOTICE_INFO: この実証モデルのエミュレーターデータは、ブラウザ内の独立保護Storage領域に隔離管理されます。PCのCドライブやデスクトップへの実用保存は、上の「SOURCE_VIEW」タブからワンタッチで全コードを保存（.html 形式など）し、ただちにオフラインでご利用いただけます。
                  </span>
                </div>
              </motion.div>
            ) : (
              <motion.div
                key="code-tab"
                initial={{ opacity: 0, y: 5 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -5 }}
                transition={{ duration: 0.15 }}
                className="space-y-3"
              >
                <div className="text-xs font-bold text-stone-800 dark:text-stone-300 flex items-center gap-2 font-mono uppercase tracking-widest pl-1">
                  <Code className="w-3.5 h-3.5 text-stone-400" />
                  <span>APP_EMBED_HTML_SOURCE (100% VANILLA)</span>
                </div>
                <div className="rounded-xl border border-stone-250/50 dark:border-stone-900 overflow-hidden shadow-inner bg-[#1A1A1A]">
                  <SyntaxHighlighter code={htmlNotebookCode} language="html" id="html-notebook-code" />
                </div>
              </motion.div>
            )}
          </AnimatePresence>
        </div>
      </div>

      {/* HTMLメモ帳の設計アーキテクチャ特徴解説 - ハイテク bento 格子 */}
      <div className="space-y-5">
        <h3 className="text-sm font-bold text-[#8C8A85] dark:text-stone-400 flex items-center gap-2 select-none font-mono tracking-widest uppercase">
          <Star className="w-4 h-4 text-stone-500 fill-current" />
          SYSTEM_ARCH_ADVANTAGES
        </h3>
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          {highlights.map((item, idx) => (
            <div key={idx} className="p-5 rounded-2xl border border-stone-250/50 dark:border-stone-900 bg-[#FAF8F5]/70 dark:bg-[#121212]/80 backdrop-blur-md shadow-sm hover:border-stone-400 dark:hover:border-stone-700 transition-all duration-300 relative group overflow-hidden interact-glow">
              
              {/* 装飾 */}
              <div className="absolute top-0 left-0 w-2 h-[1px] bg-stone-300 dark:bg-stone-800 pointer-events-none" />
              <div className="absolute top-0 left-0 h-2 w-[1px] bg-stone-300 dark:bg-stone-800 pointer-events-none" />

              <div className="flex gap-4.5 items-start">
                <div className="p-2 bg-stone-950 dark:bg-stone-200 text-white dark:text-stone-950 text-[10px] font-bold font-mono rounded-lg shrink-0 select-none">
                  {idx + 1 < 10 ? `0${idx + 1}` : idx + 1}
                </div>
                <div>
                  <h4 className="text-sm font-bold text-stone-950 dark:text-white mb-1.5 font-display tracking-tight">
                    {item.title}
                  </h4>
                  <p className="text-xs text-stone-500 dark:text-stone-400 leading-relaxed font-normal font-sans font-medium">
                    {item.desc}
                  </p>
                </div>
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
};
