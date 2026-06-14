import React, { useState } from 'react';
import { VbaSample } from '../types';
import { SyntaxHighlighter } from './SyntaxHighlighter';
import { Sparkles, Calendar, BookOpen, Star, ChevronDown, ChevronUp, CheckCircle } from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';

interface VbaCodeCardProps {
  sample: VbaSample;
  isFavorite: boolean;
  onToggleFavorite: (id: string) => void;
}

export const VbaCodeCard: React.FC<VbaCodeCardProps> = ({ sample, isFavorite, onToggleFavorite }) => {
  const [isOpenExplanation, setIsOpenExplanation] = useState(false);

  const getCategoryTheme = (category: string) => {
    switch (category) {
      case 'excel':
        return { name: 'Excel 操作', color: 'bg-stone-100 text-stone-700 border-stone-200 dark:bg-stone-900/60 dark:text-stone-300 dark:border-stone-800' };
      case 'outlook':
        return { name: 'Outlook 連携', color: 'bg-stone-100 text-stone-700 border-stone-200 dark:bg-stone-900/60 dark:text-stone-300 dark:border-stone-800' };
      case 'file':
        return { name: 'ファイル操作', color: 'bg-stone-100 text-stone-700 border-stone-200 dark:bg-stone-900/60 dark:text-stone-300 dark:border-stone-800' };
      case 'utility':
        return { name: '共通設計・汎用', color: 'bg-stone-100 text-stone-700 border-stone-200 dark:bg-stone-900/60 dark:text-stone-300 dark:border-stone-800' };
      case 'speed':
        return { name: '高速化 & 最適化', color: 'bg-stone-100 text-stone-700 border-stone-200 dark:bg-stone-900/60 dark:text-stone-300 dark:border-stone-800' };
      default:
        return { name: 'その他', color: 'bg-stone-100 text-stone-700 border-stone-200 dark:bg-stone-900/20 dark:text-stone-300 dark:border-stone-800' };
    }
  };

  const getDifficultyBadge = (difficulty: string) => {
    switch (difficulty) {
      case 'easy':
        return { name: '初級', css: 'text-stone-500 bg-stone-50 border-stone-200/60 dark:bg-stone-900/30 dark:text-stone-400 dark:border-stone-850' };
      case 'medium':
        return { name: '中級', css: 'text-stone-800 bg-stone-100 border-stone-300 dark:bg-stone-800 dark:text-stone-250 dark:border-stone-700 font-medium' };
      case 'hard':
        return { name: '上級', css: 'text-white bg-stone-900 border-stone-950 dark:bg-stone-100 dark:text-stone-950 dark:border-white font-bold' };
      default:
        return { name: '共通', css: 'text-stone-700 bg-stone-50 border-stone-200' };
    }
  };

  const cat = getCategoryTheme(sample.category);
  const diff = getDifficultyBadge(sample.difficulty);

  return (
    <motion.div
      layout
      initial={{ opacity: 0, y: 15 }}
      animate={{ opacity: 1, y: 0 }}
      exit={{ opacity: 0, scale: 0.95 }}
      transition={{ duration: 0.25 }}
      className="p-6 rounded-2xl border border-stone-250/50 dark:border-stone-900 bg-[#FAF8F5]/80 dark:bg-[#121212]/90 hover:bg-[#FAF8F5] dark:hover:bg-[#121212] hover:border-stone-400 dark:hover:border-stone-700 shadow-sm transition-all duration-300 relative group interact-glow backdrop-blur-md overflow-hidden"
    >
      {/* 未来的なコンポーネント角を飾る極小のテック装飾線 */}
      <div className="absolute top-0 left-0 w-2.5 h-[1px] bg-stone-300 dark:bg-stone-750 pointer-events-none" />
      <div className="absolute top-0 left-0 h-2.5 w-[1px] bg-stone-300 dark:bg-stone-750 pointer-events-none" />
      <div className="absolute top-0 right-0 w-2.5 h-[1px] bg-stone-300 dark:bg-stone-750 pointer-events-none" />
      <div className="absolute top-0 right-0 h-2.5 w-[1px] bg-stone-300 dark:bg-stone-750 pointer-events-none" />

      <div className="flex flex-col md:flex-row md:items-start justify-between gap-4 mb-4">
        {/* タイトルとタグ */}
        <div className="space-y-2.5 max-w-full md:max-w-[75%]">
          <div className="flex flex-wrap items-center gap-2">
            <span className={`px-2.5 py-0.5 text-[9px] font-bold rounded-lg border uppercase tracking-widest font-mono ${cat.color}`}>
              {cat.name}
            </span>
            <span className={`px-2.5 py-0.5 text-[9px] font-bold rounded-lg border uppercase tracking-widest font-mono ${diff.css}`}>
              難易度: {diff.name}
            </span>
          </div>
          <h3 className="text-xl font-bold tracking-tight text-stone-950 dark:text-white transition-colors group-hover:text-stone-900 dark:group-hover:text-[#FAFAFA]">
            {sample.title}
          </h3>
        </div>

        {/* アクションボタン（お気に入り、コピペ、など） */}
        <div className="flex items-center gap-2 self-end md:self-auto select-none shrink-0">
          <button
            onClick={() => onToggleFavorite(sample.id)}
            className={`flex items-center justify-center p-2.5 rounded-xl border transition-all cursor-pointer ${
              isFavorite
                ? 'bg-[#1A1A1A] border-[#1A1A1A] text-white dark:bg-white dark:border-white dark:text-[#101010] shadow-md'
                : 'bg-stone-50 hover:bg-stone-100 border-stone-200 text-stone-400 hover:text-stone-700 dark:bg-stone-900/40 dark:hover:bg-stone-850 dark:border-stone-850 dark:text-stone-500 dark:hover:text-stone-300'
            }`}
            title={isFavorite ? 'お気に入りから削除' : 'お気に入りに追加'}
          >
            <Star className={`w-3.5 h-3.5 ${isFavorite ? 'fill-current' : ''}`} />
          </button>
        </div>
      </div>

      <p className="text-slate-600 dark:text-slate-350 text-sm mb-5 leading-relaxed">
        {sample.description}
      </p>

      {/* タグ表示 */}
      <div className="flex flex-wrap items-center gap-1.5 mb-5 select-none text-[11px] font-medium font-mono">
        {sample.tags.map((tag) => (
          <span key={tag} className="text-stone-500 hover:text-stone-900 dark:text-stone-400 dark:hover:text-stone-100 bg-stone-100 dark:bg-stone-900 px-2.5 py-0.5 rounded transition-colors">
            #{tag}
          </span>
        ))}
      </div>

      {/* シンタックスハイライターによるコード描画 */}
      <div className="mb-5 overflow-hidden rounded-xl border border-stone-200/65 dark:border-stone-900 shadow-inner">
        <SyntaxHighlighter code={sample.code} language="vba" id={`vba-${sample.id}`} />
      </div>

      {/* 詳細解説アコーディオン */}
      <div className="border border-stone-200 dark:border-stone-900 rounded-xl overflow-hidden bg-stone-50/50 dark:bg-stone-900/10">
        <button
          onClick={() => setIsOpenExplanation(!isOpenExplanation)}
          className="w-full flex items-center justify-between px-4 py-3 select-none text-left font-medium text-xs md:text-sm text-stone-700 hover:text-stone-950 dark:text-stone-300 dark:hover:text-stone-100 hover:bg-stone-100/50 dark:hover:bg-stone-900/55 transition-colors cursor-pointer"
        >
          <span className="flex items-center gap-2 font-bold uppercase tracking-wider text-stone-600 dark:text-stone-400 font-display">
            <BookOpen className="w-4 h-4 text-stone-400" />
            VBAコードの解説と安全・最適化のポイント
          </span>
          {isOpenExplanation ? <ChevronUp className="w-4.5 h-4.5 text-stone-600 dark:text-stone-400" /> : <ChevronDown className="w-4.5 h-4.5 text-stone-400" />}
        </button>

        <AnimatePresence initial={false}>
          {isOpenExplanation && (
            <motion.div
              initial={{ height: 0, opacity: 0 }}
              animate={{ height: 'auto', opacity: 1 }}
              exit={{ height: 0, opacity: 0 }}
              transition={{ duration: 0.2 }}
              className="overflow-hidden border-t border-stone-200 dark:border-stone-900"
            >
              <div className="p-4 space-y-3 bg-[#FAF8F5]/50 dark:bg-stone-950/20 text-xs md:text-sm text-stone-600 dark:text-stone-400 leading-relaxed font-sans">
                {sample.explanation.map((item, idx) => {
                  const [title, desc] = item.split('：');
                  return (
                    <div key={idx} className="flex gap-2.5 items-start">
                      <CheckCircle className="w-4 h-4 text-stone-400 dark:text-stone-500 shrink-0 mt-0.5" />
                      <div>
                        {desc ? (
                          <>
                            <strong className="text-stone-900 dark:text-stone-200 block md:inline mr-1">{title}：</strong>
                            <span>{desc}</span>
                          </>
                        ) : (
                          <span>{item}</span>
                        )}
                      </div>
                    </div>
                  );
                })}
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </div>
    </motion.div>
  );
};
