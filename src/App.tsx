import { useState, useEffect } from 'react';
import { vbaSamples } from './data/vbaSamples';
import { VbaCodeCard } from './components/VbaCodeCard';
import { HtmlViewSection } from './components/HtmlViewSection';
import { VbaLabSection } from './components/VbaLabSection';
import { VbaPaletteSection } from './components/VbaPaletteSection';
import { 
  Sparkles, 
  Search, 
  Layers, 
  Star, 
  Sun, 
  Moon, 
  Code2, 
  BookMarked, 
  Grid, 
  Filter, 
  Check, 
  Award, 
  ChevronRight, 
  Database,
  Cpu,
  Layout
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';

export default function App() {
  const [activeSection, setActiveSection] = useState<'vba' | 'vba-lab' | 'html' | 'palette'>('vba');
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedCategory, setSelectedCategory] = useState<string>('all');
  const [selectedDifficulty, setSelectedDifficulty] = useState<string>('all');
  const [onlyShowFavorites, setOnlyShowFavorites] = useState(false);
  const [favorites, setFavorites] = useState<string[]>([]);
  const [darkMode, setDarkMode] = useState(true);

  // お気に入り & ダークモードの初期化
  useEffect(() => {
    const savedFavorites = localStorage.getItem('vba_favorites');
    if (savedFavorites) {
      try {
        setFavorites(JSON.parse(savedFavorites));
      } catch (e) {
        console.error(e);
      }
    }

    const savedTheme = localStorage.getItem('vba_theme');
    const systemPrefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
    const isDark = savedTheme ? savedTheme === 'dark' : systemPrefersDark;
    setDarkMode(isDark);
    applyTheme(isDark);
  }, []);

  const applyTheme = (isDark: boolean) => {
    const root = document.documentElement;
    if (isDark) {
      root.classList.add('dark');
      root.setAttribute('data-theme', 'dark');
    } else {
      root.classList.remove('dark');
      root.setAttribute('data-theme', 'light');
    }
  };

  const toggleDarkMode = () => {
    const newMode = !darkMode;
    setDarkMode(newMode);
    localStorage.setItem('vba_theme', newMode ? 'dark' : 'light');
    applyTheme(newMode);
  };

  const handleToggleFavorite = (id: string) => {
    const updated = favorites.includes(id)
      ? favorites.filter((favId) => favId !== id)
      : [...favorites, id];
    setFavorites(updated);
    localStorage.setItem('vba_favorites', JSON.stringify(updated));
  };

  // 検索・フィルタリング処理
  const filteredVbaSamples = vbaSamples.filter((sample) => {
    // 1. お気に入り絞り込み
    if (onlyShowFavorites && !favorites.includes(sample.id)) {
      return false;
    }

    // 2. カテゴリ絞り込み
    if (selectedCategory !== 'all' && sample.category !== selectedCategory) {
      return false;
    }

    // 3. 難易度絞り込み
    if (selectedDifficulty !== 'all' && sample.difficulty !== selectedDifficulty) {
      return false;
    }

    // 4. フリーワード検索 (タイトル、説明、タグ、ソースコードの中身)
    if (searchQuery.trim() !== '') {
      const query = searchQuery.toLowerCase();
      const inTitle = sample.title.toLowerCase().includes(query);
      const inDesc = sample.description.toLowerCase().includes(query);
      const inCode = sample.code.toLowerCase().includes(query);
      const inTags = sample.tags.some((tag) => tag.toLowerCase().includes(query));
      
      return inTitle || inDesc || inCode || inTags;
    }

    return true;
  });

  return (
    <div className="min-h-screen bg-tech-grid-light dark:bg-tech-grid-dark text-[#1A1A1A] dark:text-[#E2E2E2] transition-colors duration-300 font-sans">
      
      {/* ナビゲーションバー - ガラス質感と極細ライン */}
      <header className="sticky top-0 z-50 border-b border-stone-250/40 dark:border-stone-900/80 bg-[#FBF9F6]/85 dark:bg-[#0d0d0d]/85 backdrop-blur-lg">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between gap-4">
          
          {/* ロゴ - 研ぎ澄まされた近未来的モジュールスタイル */}
          <div className="flex items-center gap-3 shrink-0 select-none">
            <div className="relative w-9 h-9 bg-[#1A1A1A] dark:bg-[#FAF8F5] rounded-lg flex items-center justify-center shadow-lg transition-transform duration-300 group-hover/logo:scale-105">
                <span className="text-[#FBF9F6] dark:text-[#101010] font-black text-xs font-mono tracking-widest pl-[1px]">VBA</span>
              </div>
            <div>
              <h1 className="text-sm sm:text-base font-bold font-display text-stone-950 dark:text-[#FAF7F2] leading-none flex items-center gap-2 tracking-[0.05em]">
                <span>CodeHub Studio</span>
                <span className="text-[9px] font-mono font-bold bg-[#1A1A1A] text-[#FBF9F6] dark:bg-[#FAF7F2] dark:text-[#101010] px-1.5 py-0.5 rounded uppercase tracking-widest">
                  v3.4
                </span>
              </h1>
              <p className="text-[9px] text-stone-400 dark:text-stone-500 font-mono font-medium mt-1 pr-2 hidden sm:block uppercase tracking-wider">
                Excel Engine ✕ Sandboxed Local Storage
              </p>
            </div>
          </div>

          {/* セクション切り替えタブ - 未来的なスイッチデザイン */}
          <div className="flex items-center gap-1 bg-[#F5F2EC]/80 dark:bg-[#151515] p-1 rounded-xl select-none border border-stone-250/40 dark:border-stone-900/60 shadow-sm overflow-x-auto max-w-full">
            <button
              onClick={() => {
                setActiveSection('vba');
                setSearchQuery('');
              }}
              className={`flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer shrink-0 ${
                activeSection === 'vba'
                  ? 'bg-[#FAF7F2] dark:bg-stone-850 text-stone-950 dark:text-white shadow-sm'
                  : 'text-stone-500 dark:text-stone-400 hover:text-stone-800 dark:hover:text-stone-200'
              }`}
            >
              <Database className="w-3.5 h-3.5 animate-none" />
              <span>VBAライブラリ</span>
            </button>
            <button
              onClick={() => {
                setActiveSection('vba-lab');
                setSearchQuery('');
              }}
              className={`flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer shrink-0 ${
                activeSection === 'vba-lab'
                  ? 'bg-stone-950 text-white dark:bg-[#FAF7F2] dark:text-stone-950 shadow-sm font-black'
                  : 'text-stone-500 dark:text-stone-400 hover:text-stone-800 dark:hover:text-stone-200'
              }`}
            >
              <Cpu className="w-3.5 h-3.5 animate-pulse" />
              <span>VBA最適化ラボ</span>
            </button>
            <button
              onClick={() => {
                setActiveSection('palette');
                setSearchQuery('');
              }}
              className={`flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer shrink-0 ${
                activeSection === 'palette'
                  ? 'bg-stone-950 text-white dark:bg-[#FAF7F2] dark:text-stone-950 shadow-sm font-black'
                  : 'text-stone-500 dark:text-stone-400 hover:text-stone-800 dark:hover:text-stone-200'
              }`}
            >
              <Layout className="w-3.5 h-3.5" />
              <span>VBAパレット</span>
            </button>
            <button
              onClick={() => {
                setActiveSection('html');
                setSearchQuery('');
              }}
              className={`flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer shrink-0 ${
                activeSection === 'html'
                  ? 'bg-stone-950 text-white dark:bg-[#FAF7F2] dark:text-stone-950 shadow-sm font-black'
                  : 'text-stone-500 dark:text-stone-400 hover:text-stone-800 dark:hover:text-stone-200'
              }`}
            >
              <BookMarked className="w-3.5 h-3.5" />
              <span>HTMLセキュアメモ</span>
            </button>
          </div>

          {/* 右部分: ステータスインジゲーターとダークモードスイッチ */}
          <div className="flex items-center gap-4">
            {/* 本質的な安全ステータス（ローカルストレージ駆動である実態を示すインジケーター） */}
            <div className="hidden lg:flex items-center gap-2 px-3 py-1 bg-stone-200/40 dark:bg-stone-900/50 border border-stone-250/50 dark:border-stone-850/80 rounded-full font-mono text-[9px] text-stone-500 dark:text-stone-400">
              <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-tech-pulse shrink-0"></span>
              <span className="tracking-wider">SECURE SHIELD: LOCALONLY</span>
            </div>

            {/* ダークモード切り替え */}
            <div className="flex items-center gap-2 border-l pl-4 border-stone-200/60 dark:border-stone-900/60">
              <span className="text-[9px] font-bold font-mono text-stone-400 dark:text-stone-550 uppercase tracking-widest hidden md:block">MODE_SYS</span>
              <button
                onClick={toggleDarkMode}
                className={`w-11 h-6 rounded-full relative flex items-center transition-colors shadow-inner select-none cursor-pointer ${
                  darkMode ? 'bg-stone-800' : 'bg-stone-250'
                }`}
                title={darkMode ? 'ライトモードに切り替え' : 'ダークモードに切り替え'}
              >
                <motion.div
                  layout
                  className="w-4 h-4 bg-[#FAF7F2] dark:bg-stone-950 rounded-full shadow absolute left-1 flex items-center justify-center animate-none"
                  animate={{ x: darkMode ? 20 : 0 }}
                  transition={{ type: 'spring', stiffness: 500, damping: 30 }}
                >
                  {darkMode ? (
                    <Moon className="w-2.5 h-2.5 text-stone-200 animate-none" />
                  ) : (
                    <Sun className="w-2.5 h-2.5 text-stone-900 animate-none" />
                  )}
                </motion.div>
              </button>
            </div>
          </div>

        </div>
      </header>

      {/* メインヒーローエリア - 未来派グリッド＆極細ボーダーアクセント */}
      <div className="relative border-b border-stone-250/30 dark:border-stone-900 py-14 md:py-20 overflow-hidden select-none">
        
        {/* 背景に近未来的な光の筋やボケを、うるさくないモノトーンで演出 */}
        <div className="absolute inset-0 bg-gradient-to-b from-stone-500/[0.02] to-transparent pointer-events-none" />
        <div className="absolute right-[10%] top-[20%] w-[500px] h-[500px] rounded-full bg-stone-900/[0.01] dark:bg-white/[0.01] blur-3xl pointer-events-none" />
        
        {/* レーザーのように繊細な1pxの未来的飾り罫線 */}
        <div className="absolute left-0 right-0 top-0 h-[10px] flex items-center justify-between px-8 text-stone-300 dark:text-stone-800 font-mono text-[8px] pointer-events-none select-none opacity-40">
          <span>// CORE_ENGINE_INITIALIZED</span>
          <span>SYSTEM_STABLE: 100%</span>
        </div>

        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 text-center relative z-10 space-y-6">
          <motion.div
            initial={{ opacity: 0, y: 8 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.4 }}
            className="flex justify-center"
          >
            <span className="inline-flex items-center gap-2 px-3 py-1 rounded-full text-[9px] md:text-[10px] font-bold font-mono tracking-widest bg-stone-100 dark:bg-stone-900/60 text-stone-600 dark:text-stone-400 border border-stone-250/60 dark:border-stone-800/80 uppercase">
              <Sparkles className="w-3 h-3 animate-none text-stone-500 dark:text-stone-400" />
              CURATED DEVELOPMENT HARBOUR
            </span>
          </motion.div>
          
          <div className="space-y-4">
            {activeSection === 'vba' ? (
              <h2 className="leading-tight max-w-4xl mx-auto select-none">
                <span className="text-xs font-mono font-bold uppercase tracking-[0.25em] text-stone-500 dark:text-stone-400 block mb-3">
                  [ SYSTEM_MANEUVERING_PORTAL ]
                </span>
                <span className="text-3xl sm:text-4xl md:text-5.5xl font-black tracking-tighter text-stone-950 dark:text-white font-display block leading-none">
                  Excel を自動操縦する
                </span>
                <span className="text-2xl sm:text-3xl md:text-4.5xl font-light tracking-tight text-stone-600 dark:text-stone-300 font-display block mt-1.5 font-sans">
                  極限まで最適化された実務 VBA ライブラリ
                </span>
              </h2>
            ) : activeSection === 'vba-lab' ? (
              <h2 className="leading-tight max-w-4xl mx-auto select-none">
                <span className="text-xs font-mono font-bold uppercase tracking-[0.25em] text-stone-500 dark:text-stone-400 block mb-3">
                  [ INTERACTIVE_BENCHMARK_LAB ]
                </span>
                <span className="text-3xl sm:text-4xl md:text-5.5xl font-black tracking-tighter text-stone-950 dark:text-white font-display block leading-none">
                  プロセスの極限高速化
                </span>
                <span className="text-2xl sm:text-3xl md:text-4.5xl font-light tracking-tight text-stone-600 dark:text-stone-300 font-display block mt-1.5 font-sans">
                  インタラクティブ自動診断 & 効果シミュレーター
                </span>
              </h2>
            ) : activeSection === 'palette' ? (
              <h2 className="leading-tight max-w-4xl mx-auto select-none">
                <span className="text-xs font-mono font-bold uppercase tracking-[0.25em] text-stone-500 dark:text-stone-400 block mb-3">
                  [ MULTI_FUNCTION_FLOATING_PALETTE ]
                </span>
                <span className="text-3xl sm:text-4xl md:text-5.5xl font-black tracking-tighter text-stone-950 dark:text-white font-display block leading-none">
                  手のひらにカスタムツール
                </span>
                <span className="text-2xl sm:text-3xl md:text-4.2xl font-light tracking-tight text-stone-600 dark:text-stone-300 font-display block mt-1.5 font-sans">
                  最前面で浮遊する、超高機能フローティングパレット VBA
                </span>
              </h2>
            ) : (
              <h2 className="leading-tight max-w-4xl mx-auto select-none">
                <span className="text-xs font-mono font-bold uppercase tracking-[0.25em] text-stone-500 dark:text-stone-400 block mb-3">
                  [ STANDALONE_SANDBOX_SECURE ]
                </span>
                <span className="text-3xl sm:text-4xl md:text-5.5xl font-black tracking-tighter text-stone-950 dark:text-white font-display block leading-none">
                  安心安全を形にした
                </span>
                <span className="text-2xl sm:text-3xl md:text-4.2xl font-light tracking-tight text-stone-600 dark:text-stone-300 font-display block mt-1.5 font-sans">
                  隔離動作設計・高機能ローカル HTML メモ帳
                </span>
              </h2>
            )}
            
            <p className="text-[11px] md:text-xs text-stone-500 dark:text-stone-400 max-w-2xl mx-auto leading-relaxed font-sans max-w-[65ch]">
              {activeSection === 'vba' && '非同期的な描画カット、厳密な型定義、安全なオブジェクト解放。企業の制約された環境でも、そのままコピペして最高のスピードを叩き出せるライブラリ。'}
              {activeSection === 'vba-lab' && '実際に自分のコードを評価し、ボトルネックを特定。さらにボタン一つでエラーに強いテンプレート自動生成、業務の劇的な時間・費用削減シミュレーションを行なえます。'}
              {activeSection === 'palette' && 'モーダレスフォーム機能(UserForm)をフル活用。ドラッグ操作、動的な設定シートからのボタン展開、複数バグを完璧にケアした安心設計の一大モジュール一式。'}
              {activeSection === 'html' && 'サーバー連携、外部APIを一切必要としない究極のローカル駆動。完璧な親子ツリー、スマート開閉フォルダー、堅牢なJSONエクスポート機能をご体験ください。'}
            </p>
          </div>

          {/* 装飾用の未来的ミニマルドット、クロス */}
          <div className="flex items-center justify-center gap-12 text-stone-300 dark:text-stone-800 font-mono text-[9px] pt-4 pointer-events-none">
            <div className="flex items-center gap-1.5">
              <span className="inline-block w-1.5 h-1.5 bg-stone-400 dark:bg-stone-600 rounded-full animate-pulse" />
              <span>OFFLINE FIRST</span>
            </div>
            <span>+</span>
            <div className="flex items-center gap-1.5">
              <span className="inline-block w-1.5 h-1.5 bg-stone-400 dark:bg-stone-600 rounded-full animate-pulse" />
              <span>COPY READY</span>
            </div>
            <span>+</span>
            <div className="flex items-center gap-1.5">
              <span className="inline-block w-1.5 h-1.5 bg-stone-400 dark:bg-stone-600 rounded-full animate-pulse" />
              <span>ZERO TRACKING</span>
            </div>
          </div>

        </div>
      </div>

      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 md:py-12">
        
        {/* === SECTION 1: VBA サンプル展示集 === */}
        {activeSection === 'vba' && (
          <div className="grid grid-cols-1 lg:grid-cols-4 gap-8">
            
            {/* 左側フィルターメニュー */}
            <div className="lg:col-span-1 space-y-6">
              
              {/* 重宝する検索ボックス */}
              <div className="p-5 rounded-2xl border border-stone-250/60 dark:border-stone-900 bg-[#FAF7F2]/90 dark:bg-[#121212]/80 backdrop-blur-md shadow-sm space-y-3 interact-glow">
                <h3 className="text-[10px] font-bold uppercase tracking-widest text-stone-500 dark:text-stone-500 flex items-center gap-2 font-mono">
                  <Search className="w-3.5 h-3.5" />
                  QUERY_ENGINE_SEARCH
                </h3>
                <div className="relative">
                  <input
                    type="text"
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    placeholder="キーワード、API、関数名..."
                    className="w-full pl-9 pr-4 py-2 text-xs rounded-xl border border-stone-200 dark:border-stone-850 bg-stone-50/50 dark:bg-stone-950 text-stone-900 dark:text-stone-100 placeholder-stone-400 dark:placeholder-stone-600 focus:outline-none focus:ring-1 focus:ring-stone-400 dark:focus:ring-[#FAF7F2] transition-all duration-300"
                  />
                  <Search className="absolute left-3 top-2.5 w-3.5 h-3.5 text-stone-500 dark:text-stone-600" />
                  {searchQuery && (
                    <button 
                      onClick={() => setSearchQuery('')}
                      className="absolute right-3 top-2.5 text-[10px] text-stone-500 hover:text-stone-900 dark:text-stone-500 dark:hover:text-[#FAF7F2] font-bold cursor-pointer font-mono"
                    >
                      RESET
                    </button>
                  )}
                </div>
              </div>

              {/* フィルター条件 */}
              <div className="p-5 rounded-2xl border border-stone-250/60 dark:border-stone-900 bg-[#FAF7F2]/90 dark:bg-[#121212]/80 backdrop-blur-md shadow-sm space-y-5 interact-glow">
                
                {/* お気に入りトグル */}
                <div className="flex items-center justify-between pb-4 border-b border-stone-200/50 dark:border-stone-900/40 font-sans">
                  <span className="text-xs font-semibold text-stone-700 dark:text-stone-300 tracking-wide font-display">
                    お気に入りのみ瞬時表示
                  </span>
                  <button
                    onClick={() => setOnlyShowFavorites(!onlyShowFavorites)}
                    className={`relative inline-flex h-5 w-9 shrink-0 items-center rounded-full transition-colors cursor-pointer ${
                      onlyShowFavorites ? 'bg-stone-950 dark:bg-[#FAF7F2]' : 'bg-stone-200 dark:bg-stone-800'
                    }`}
                  >
                    <span
                      className={`inline-block h-3 w-3 transform rounded-full bg-[#FAF7F2] dark:bg-stone-950 transition-transform ${
                        onlyShowFavorites ? 'translate-x-5' : 'translate-x-1'
                      }`}
                    />
                  </button>
                </div>

                {/* カテゴリ */}
                <div className="space-y-3">
                  <h3 className="text-[10px] font-bold uppercase tracking-widest text-stone-500 dark:text-stone-500 flex items-center gap-2 select-none font-mono">
                    <Filter className="w-3.5 h-3.5" />
                    FILTER_CATEGORY
                  </h3>
                  <div className="flex flex-col gap-1 select-none font-sans">
                    {[
                      { key: 'all', label: 'すべてのコード' },
                      { key: 'speed', label: '高速化 & 最適化' },
                      { key: 'excel', label: 'Excel 操作 & 構造化転記' },
                      { key: 'outlook', label: 'Outlook 連携自動作成' },
                      { key: 'file', label: 'ファイル操作汎用系' },
                      { key: 'utility', label: '共通設計 ＆ 汎用ロジック' }
                    ].map((cat) => {
                      const isActive = selectedCategory === cat.key;
                      return (
                        <button
                          key={cat.key}
                          onClick={() => setSelectedCategory(cat.key)}
                          className={`w-full text-left px-3 py-2 rounded-xl text-xs font-semibold transition-all flex items-center justify-between cursor-pointer border ${
                            isActive
                              ? 'bg-stone-950 text-[#FAF7F2] dark:bg-[#FAF7F2] dark:text-stone-950 font-bold border-stone-950 dark:border-[#FAF7F2] shadow-sm'
                              : 'text-stone-600 hover:text-stone-950 border-transparent hover:bg-stone-100/60 dark:text-stone-400 dark:hover:text-white dark:hover:bg-stone-900/50'
                          }`}
                        >
                          <span>{cat.label}</span>
                          {isActive && <Check className="w-3.5 h-3.5 border-none" />}
                        </button>
                      );
                    })}
                  </div>
                </div>

                {/* 難易度 */}
                <div className="space-y-3 pt-4 border-t border-stone-200/50 dark:border-stone-900/60">
                  <h3 className="text-[10px] font-bold uppercase tracking-widest text-[#8C8A85] dark:text-stone-500 flex items-center gap-2 select-none font-mono">
                    <Award className="w-3.5 h-3.5" />
                    FILTER_DIFFICULTY
                  </h3>
                  <div className="flex flex-col gap-1 select-none font-sans">
                    {[
                      { key: 'all', label: 'すべての難易度' },
                      { key: 'easy', label: '初級 (基本構文、高速化宣言など)' },
                      { key: 'medium', label: '中級 (ファイル操作、Outlook連携)' },
                      { key: 'hard', label: '上級 (オートフィルタ連動、重複排除)' }
                    ].map((diff) => {
                      const isActive = selectedDifficulty === diff.key;
                      return (
                        <button
                          key={diff.key}
                          onClick={() => setSelectedDifficulty(diff.key)}
                          className={`w-full text-left px-3 py-2 rounded-xl text-xs font-semibold transition-all flex items-center justify-between cursor-pointer border ${
                            isActive
                              ? 'bg-stone-950 text-[#FAF7F2] dark:bg-[#FAF7F2] dark:text-stone-950 font-bold border-stone-950 dark:border-[#FAF7F2] shadow-sm'
                              : 'text-stone-600 hover:text-stone-950 border-transparent hover:bg-stone-100/60 dark:text-stone-400 dark:hover:text-white dark:hover:bg-stone-900/50'
                          }`}
                        >
                          <span>{diff.label}</span>
                          {isActive && <Check className="w-3.5 h-3.5 border-none animate-none" />}
                        </button>
                      );
                    })}
                  </div>
                </div>

              </div>

              {/* 便利なVBA豆知識 - 究極のミニマル・インフォメーション */}
              <div className="p-5 rounded-2xl border border-stone-250/60 dark:border-stone-900/80 bg-white/50 dark:bg-[#121212]/85 text-stone-850 dark:text-stone-200 shadow-sm space-y-3 hidden lg:block backdrop-blur-md interact-glow">
                <div className="flex items-center gap-1 text-stone-900 dark:text-white font-mono text-[10px]">
                  <Sparkles className="w-4 h-4 text-stone-400 dark:text-stone-500 animate-none" />
                  <span className="font-bold tracking-widest pl-[1px]">ENGINE_GUIDE_SYS</span>
                </div>
                <h4 className="text-xs font-bold font-display tracking-tight text-stone-950 dark:text-white pt-1">高速化とエラーハンドリングの極意</h4>
                <p className="text-[10px] text-stone-500 dark:text-stone-400 leading-relaxed font-sans font-medium">
                  画面の更新を止め（ScreenUpdating = False）、On Errorで確実にエラーを捕捉し、Set Object = Nothingによるメモリ解放フローを加える。この数秒の配慮が、100倍の処理効率と壊れない堅牢性を実証します。
                </p>
              </div>

            </div>

            {/* 右側：VBAコードカード一覧 */}
            <div className="lg:col-span-3 space-y-6">
              
              <div className="flex items-center justify-between border-b border-stone-250/30 dark:border-stone-900/60 pb-3">
                <div className="text-[11px] font-mono text-stone-500 dark:text-stone-500 select-none uppercase tracking-wider">
                  MATCHING_RECORDS_FOUND: <span className="font-bold text-stone-900 dark:text-white font-sans">{filteredVbaSamples.length}</span> UNITS
                </div>
                
                {/* 簡易タグ */}
                <div className="flex gap-2 text-[9px] font-mono">
                  {onlyShowFavorites && (
                    <span className="bg-stone-950 text-[#FBF9F6] dark:bg-[#FAF7F2] dark:text-stone-950 px-2.5 py-0.5 rounded-lg font-bold uppercase tracking-wider shadow-sm">
                      ★ FAVORITES_ONLY
                    </span>
                  )}
                  {selectedCategory !== 'all' && (
                    <span className="bg-stone-200/60 dark:bg-stone-900 text-stone-800 dark:text-stone-300 px-2.5 py-0.5 rounded-lg font-bold uppercase tracking-wider border border-transparent dark:border-stone-800">
                      CAT: {selectedCategory}
                    </span>
                  )}
                </div>
              </div>

              {/* カード本体のグリッド/リスト */}
              <div className="space-y-6">
                <AnimatePresence mode="popLayout">
                  {filteredVbaSamples.length > 0 ? (
                    filteredVbaSamples.map((sample) => (
                      <VbaCodeCard
                        key={sample.id}
                        sample={sample}
                        isFavorite={favorites.includes(sample.id)}
                        onToggleFavorite={handleToggleFavorite}
                      />
                    ))
                  ) : (
                    <motion.div
                      layout
                      initial={{ opacity: 0 }}
                      animate={{ opacity: 1 }}
                      exit={{ opacity: 0 }}
                      className="text-center py-20 p-8 rounded-2xl border border-dashed border-stone-250/50 dark:border-stone-850 bg-[#FAF7F2]/80 dark:bg-[#121212]/40"
                    >
                      <Search className="w-12 h-12 text-slate-300 dark:text-slate-700 mx-auto mb-4" />
                      <h4 className="text-sm font-bold text-slate-700 dark:text-slate-300 mb-1">
                        一致するVBAマクロが見つかりません
                      </h4>
                      <p className="text-xs text-slate-500 dark:text-slate-500 max-w-sm mx-auto">
                        検索キーワードやカテゴリ、難易度フィルターの設定を変更してもう一度お試しください。
                      </p>
                    </motion.div>
                  )}
                </AnimatePresence>
              </div>

            </div>

          </div>
        )}

        {/* === SECTION 2: HTML セキュアメモ展示・プレビュー === */}
        {activeSection === 'html' && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            transition={{ duration: 0.3 }}
          >
            <HtmlViewSection />
          </motion.div>
        )}

        {/* === SECTION 3: VBAラボ（自動最適化 & 診断ツール） === */}
        {activeSection === 'vba-lab' && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            transition={{ duration: 0.3 }}
          >
            <VbaLabSection />
          </motion.div>
        )}

        {/* === SECTION 4: VBAパレット === */}
        {activeSection === 'palette' && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            transition={{ duration: 0.3 }}
          >
            <VbaPaletteSection />
          </motion.div>
        )}

      </main>

      {/* 美しいおしゃーフッター */}
      <footer className="border-t border-stone-200/80 dark:border-stone-900 bg-[#FBF9F6] dark:bg-[#101010]/95 py-12 select-none">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 text-center space-y-4">
          <div className="text-stone-400 dark:text-stone-500 text-xs sm:text-sm font-medium tracking-wider font-display uppercase">
            VBA & HTML CodeHub Portal — 実践的・セキュア・効率化
          </div>
          <p className="text-[11px] text-stone-500 dark:text-stone-600 max-w-md mx-auto leading-relaxed">
            すべてのコードはブラウザのみ、またはPC内のExcel環境のみで動作し、データや個人特定情報を外部のサーバーへ送信することはありません。ローカルファースト、オフライン重視で安心安全。
          </p>
          <div className="text-[10px] text-stone-350 dark:text-stone-700 font-mono">
            &copy; 2026 CodeHub Studio. All rights reserved.
          </div>
        </div>
      </footer>
    </div>
  );
}
