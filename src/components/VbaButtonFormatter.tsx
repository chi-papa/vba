import React, { useState, useEffect } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  Download, 
  Copy, 
  Check, 
  Palette, 
  Type, 
  Maximize2, 
  FileCode, 
  Layers, 
  Sliders, 
  RefreshCw, 
  LayoutGrid,
  Info
} from 'lucide-react';
import { SyntaxHighlighter } from './SyntaxHighlighter';

// Excel/VBA カラー構造体を想定した変換用のヘルパー
interface ColorPreset {
  name: string;
  hex: string;
  vbaColor: string; // RGB(R, G, B) または Long値
}

const COMMON_FONTS = [
  { value: 'Meiryo UI', label: 'メイリオ (Meiryo UI)' },
  { value: 'Yu Gothic UI', label: '遊ゴシック (Yu Gothic UI)' },
  { value: 'MS PGothic', label: 'ＭＳ Ｐゴシック' },
  { value: 'BIZ UDPGothic', label: 'BIZ UDPゴシック' },
  { value: 'Segoe UI', label: 'Segoe UI (Global Standard)' },
  { value: 'Calibri', label: 'Calibri (Clean English)' },
];

const BACK_COLOR_PRESETS: ColorPreset[] = [
  { name: 'モダンブルー', hex: '#2B6CB0', vbaColor: 'RGB(43, 108, 176)' },
  { name: 'ミッドナイト', hex: '#1A202C', vbaColor: 'RGB(26, 32, 44)' },
  { name: 'ソフトエメラルド', hex: '#319795', vbaColor: 'RGB(49, 151, 149)' },
  { name: 'クレイオレンジ', hex: '#DD6B20', vbaColor: 'RGB(221, 107, 32)' },
  { name: 'クリムゾンレッド', hex: '#E53E3E', vbaColor: 'RGB(229, 62, 62)' },
  { name: 'ロイヤルパープル', hex: '#805AD5', vbaColor: 'RGB(128, 90, 213)' },
  { name: 'シックグレー', hex: '#4A5568', vbaColor: 'RGB(74, 85, 104)' },
  { name: 'クラシックExcelブルー', hex: '#2980B9', vbaColor: 'RGB(41, 128, 185)' },
  { name: 'Excelグリーン', hex: '#107C41', vbaColor: 'RGB(16, 124, 65)' },
  { name: 'ピュアホワイト', hex: '#FFFFFF', vbaColor: 'RGB(255, 255, 255)' },
];

const FORE_COLOR_PRESETS: ColorPreset[] = [
  { name: 'プレーンホワイト', hex: '#FFFFFF', vbaColor: 'RGB(255, 255, 255)' },
  { name: 'ダークチャコール', hex: '#1A202C', vbaColor: 'RGB(26, 32, 44)' },
  { name: 'ソフトゴールド', hex: '#FEFCBF', vbaColor: 'RGB(254, 252, 191)' },
  { name: 'ミントシアン', hex: '#E6FFFA', vbaColor: 'RGB(230, 255, 250)' },
  { name: 'ライトグレー', hex: '#EDF2F7', vbaColor: 'RGB(237, 242, 247)' },
];

export const VbaButtonFormatter: React.FC = () => {
  // 基本状態
  const [targetType, setTargetType] = useState<'userform' | 'worksheet'>('userform');
  const [btnName, setBtnName] = useState('btnAction');
  const [btnCaption, setBtnCaption] = useState('🚀 処理を実行');
  const [width, setWidth] = useState<number>(120); // Point (VBA標準ユニット)
  const [height, setHeight] = useState<number>(40); // Point
  
  // スタイル・配色状態
  const [backColor, setBackColor] = useState('#2B6CB0');
  const [foreColor, setForeColor] = useState('#FFFFFF');
  const [fontSize, setFontSize] = useState<number>(11);
  const [fontName, setFontName] = useState('Meiryo UI');
  const [fontBold, setFontBold] = useState(true);
  const [fontItalic, setFontItalic] = useState(false);

  // シートボタン形状専用の追加オプション
  const [roundedAmount, setRoundedAmount] = useState<number>(0.15); // Adjustments(1) 0.0〜0.5
  const [hasShadow, setHasShadow] = useState(true);
  const [borderWidth, setBorderWidth] = useState<number>(1.5);
  const [borderColor, setBorderColor] = useState('#1A202C');

  // UserForm専用の追加オプション
  const [specialEffect, setSpecialEffect] = useState<number>(0); // 0=fmSpecialEffectFlat, 1=fmSpecialEffectRaised, etc.
  const [wordWrap, setWordWrap] = useState(true);

  // コピー状態、リセット完了状態など
  const [copied, setCopied] = useState(false);
  const [showGrid, setShowGrid] = useState(true);

  // HEXからVBAで使えるRGB(R, G, B)形式への変換
  const hexToVbaRgb = (hexStr: string): string => {
    let hex = hexStr.replace('#', '');
    if (hex.length === 3) {
      hex = hex.split('').map(char => char + char).join('');
    }
    const r = parseInt(hex.substring(0, 2), 16) || 0;
    const g = parseInt(hex.substring(2, 4), 16) || 0;
    const b = parseInt(hex.substring(4, 6), 16) || 0;
    return `RGB(${r}, ${g}, ${b})`;
  };

  const getVbaRgbStr = (hex: string) => {
    const preset = [...BACK_COLOR_PRESETS, ...FORE_COLOR_PRESETS].find(p => p.hex.toLowerCase() === hex.toLowerCase());
    return preset ? preset.vbaColor : hexToVbaRgb(hex);
  };

  // リアルタイムに生成されるVBAソースコード
  const [vbaCode, setVbaCode] = useState('');

  useEffect(() => {
    if (targetType === 'userform') {
      // ユーザーフォーム用コントロールコード生成
      const code = `' ============================================================
' ■ UserForm：「${btnName}」の動的プロパティスタイリング
'   設置方法：UserFormの「Initialize」イベント、または追加・起動時に
'   以下のプロパティ設定コードを貼り付けることで、洗練された外観を確実に強制します。
' ============================================================
Private Sub UserForm_Initialize()
    With Me.${btnName}
        .Caption = "${btnCaption}"
        .Width = ${width}
        .Height = ${height}
        
        ' スタイリング配色
        .BackColor = ${getVbaRgbStr(backColor)}
        .ForeColor = ${getVbaRgbStr(foreColor)}
        
        ' フォント設定の最適化
        With .Font
            .Name = "${fontName}"
            .Size = ${fontSize}
            .Bold = ${fontBold ? 'True' : 'False'}
            .Italic = ${fontItalic ? 'True' : 'False'}
        End With
        
        ' その他特殊効果
        .SpecialEffect = ${specialEffect} ' ${
          specialEffect === 0 ? '0=fmSpecialEffectFlat' : 
          specialEffect === 1 ? '1=fmSpecialEffectRaised' : 
          specialEffect === 2 ? '2=fmSpecialEffectSunken' : '3=fmSpecialEffectEtched'
        }
        .WordWrap = ${wordWrap ? 'True' : 'False'}
        
        ' フォーカス線やテキストアライメントの中央寄せ
        .TextAlign = 2 ' 2=fmTextAlignCenter (真ん中揃え)
    End With
End Sub`;
      setVbaCode(code);
    } else {
      // ワークシート上に形状(Shape)マクロボタンを動的に作成・整形するコード生成
      const code = `' ============================================================
' ■ Worksheet：美しいモダン Shape ボタンの自動配置＆スタイルマクロ
'   機能：Excelシート上の指定したセルの位置に合わせて、角丸
'   高機能ボタンスタイルを完全自動生成し、指定マクロを登録します。
'   貼り付け先：標準モジュール
' ============================================================
Sub CreateModernShapeButton()
    Dim ws As Worksheet
    Set ws = ActiveSheet ' もしくは ThisWorkbook.Sheets("コントロールシート")
    
    Dim btnShape As Shape
    Dim targetRange As Range
    
    ' 配置位置となる基準セル（例: D2セル、あるいはアクティブセルの右など）
    On Error Resume Next
    Set targetRange = ws.Range("D2")
    If targetRange Is Nothing Then Set targetRange = ActiveCell
    On Error GoTo 0
    
    ' 既存の同名ボタンがあれば重複防止のため先行削除
    On Error Resume Next
    ws.Shapes("${btnName}").Delete
    On Error GoTo 0
    
    ' 1. 指定の位置・サイズに角丸四角形（msoShapeRoundedRectangle）を配置
    '    ※ targetRange.Left, targetRange.Top を使用して美しく吸着させられます。
    Set btnShape = ws.Shapes.AddShape( _
        Type:=5, _ ' 5=msoShapeRoundedRectangle（角丸長方形）
        Left:=targetRange.Left + 5, _
        Top:=targetRange.Top + 5, _
        Width:=${width}, _
        Height:=${height} _
    )
    
    ' 2. ボタンの各種属性・形状プロパティの最適化
    With btnShape
        .Name = "${btnName}"
        
        ' 角丸の鋭さコントロール (0.0=直角 〜 0.5=完全半円形)
        .Adjustments(1) = ${roundedAmount}
        
        ' 塗りつぶし設定 (単色)
        With .Fill
            .Solid
            .ForeColor.RGB = ${getVbaRgbStr(backColor)}
            .Transparency = 0#
        End With
        
        ' 枠線（ボーダー）
        With .Line
            .Visible = msoTrue
            .ForeColor.RGB = ${getVbaRgbStr(borderColor)}
            .Weight = ${borderWidth}
        End With
        
        ' 3. テキストの表示・アライメント・フォント設定
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle ' 縦方向の中央寄せ
            With .TextRange
                .Text = "${btnCaption}"
                
                ' 段落プロパティで横方向の中央寄せ
                .ParagraphFormat.Alignment = msoCenter
                
                ' フォントの適用
                With .Font
                    .NameComplexScript = "${fontName}"
                    .NameFarEast = "${fontName}"
                    .Name = "${fontName}"
                    .Size = ${fontSize}
                    .Bold = ${fontBold ? 'msoTrue' : 'msoFalse'}
                    .Italic = ${fontItalic ? 'msoTrue' : 'msoFalse'}
                    .Fill.ForeColor.RGB = ${getVbaRgbStr(foreColor)}
                End With
            End With
            
            ' マージンを詰めてはみ出しを防ぐ
            .MarginLeft = 2
            .MarginRight = 2
            .MarginTop = 1
            .MarginBottom = 1
        End With
        
        ' 影(シャドウ)効果のコントロール
        With .Shadow
            .Visible = ${hasShadow ? 'msoTrue' : 'msoFalse'}
            If .Visible Then
                .Type = msoShadow25 ' 下に沈み込むソフトな落ち影
                .Blur = 4
                .OffsetX = 1.5
                .OffsetY = 1.5
                .Transparency = 0.5
            End If
        End With
        
        ' 4. ボタンをクリックしたときに実行するマクロを紐づける
        '    ※実行したい標準モジュールのマクロ名を指定
        .OnAction = "CheckFileUpdates" ' ← 例：更新チェックマクロを設定
        
        ' マウスホバー時にカーソルを指マーク(手型)にする設定
        .AlternativeText = "ボタンをクリックして処理を実行"
    End With
    
    MsgBox "セル " & targetRange.Address(False, False) & " 近傍にモダン形状ボタンを生成しました！", vbInformation, "生成完了"
End Sub`;
      setVbaCode(code);
    }
  }, [
    targetType, btnName, btnCaption, width, height, 
    backColor, foreColor, fontSize, fontName, fontBold, fontItalic,
    roundedAmount, hasShadow, borderWidth, borderColor,
    specialEffect, wordWrap
  ]);

  const handleCopyCode = () => {
    navigator.clipboard.writeText(vbaCode).then(() => {
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    }).catch(err => console.error('コピーに失敗しました:', err));
  };

  const syncPresetColors = (bgHex: string, forceFgHex?: string) => {
    setBackColor(bgHex);
    if (forceFgHex) {
      setForeColor(forceFgHex);
    } else {
      // 背景が明るい場合はフォアを暗く、暗い場合は明るく自動誘導
      const lightThresholdHex = bgHex.replace('#', '');
      const r = parseInt(lightThresholdHex.substring(0, 2), 16) || 0;
      const g = parseInt(lightThresholdHex.substring(2, 4), 16) || 0;
      const b = parseInt(lightThresholdHex.substring(4, 6), 16) || 0;
      const brightness = (r * 299 + g * 587 + b * 114) / 1000;
      setForeColor(brightness > 140 ? '#1A202C' : '#FFFFFF');
    }
  };

  // ピクセルからPoint換算値の概算表示用 (VBA 1pt = 約1.33px)
  const pxToPt = (px: number) => Math.round(px / 1.33);
  const ptToPx = (pt: number) => Math.round(pt * 1.33);

  return (
    <div className="space-y-12 select-none">
      
      {/* 情報案内カード */}
      <div className="p-8 md:p-10 rounded-3xl border border-stone-250/60 dark:border-stone-900 bg-[#FAF7F2]/90 dark:bg-[#121212]/80 backdrop-blur-md shadow-sm relative overflow-hidden group interact-glow">
        <div className="absolute right-0 top-0 w-80 h-80 rounded-full bg-amber-500/[0.015] dark:bg-amber-400/[0.015] blur-3xl pointer-events-none" />
        <div className="absolute left-0 top-0 w-2 h-full bg-indigo-600 dark:bg-indigo-400 rounded-l-3xl" />
        
        <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-6">
          <div className="space-y-3">
            <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-[9px] font-bold font-mono tracking-widest bg-amber-500/10 text-amber-600 dark:text-amber-400 border border-amber-500/20 uppercase">
              <Palette className="w-3.5 h-3.5" />
              BUTTON_SHAPE_SIMULATOR_PRO
            </span>
            <h2 className="text-2xl md:text-3.5xl font-black tracking-tight text-stone-950 dark:text-white font-display">
              🎨 VBAボタン形状フォーマッター
            </h2>
            <p className="text-xs md:text-sm text-stone-500 dark:text-stone-400 max-w-2xl leading-relaxed">
              VBAやExcelシートで利用するボタンの「外観（形状、サイズ、色、フォント、影）」を100%リアルタイムに直感的シミュレーション。設定に沿った、ExcelとUserForm用の美しく強固なプロパティ初期化・動的生成VBAマクロを自動で吐き出します。
            </p>
          </div>
          
          <div className="flex items-center gap-1 bg-[#F5F2EC]/80 dark:bg-[#151515] p-1 rounded-xl border border-stone-250/40 dark:border-stone-900/60 shadow-sm shrink-0">
            <button
              onClick={() => setTargetType('userform')}
              className={`flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg text-[11px] font-bold transition-all cursor-pointer ${
                targetType === 'userform'
                  ? 'bg-amber-500 text-white shadow-sm'
                  : 'text-stone-500 dark:text-stone-400 hover:text-stone-800 dark:hover:text-stone-200'
              }`}
            >
              <span>UserFormボタン</span>
            </button>
            <button
              onClick={() => setTargetType('worksheet')}
              className={`flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg text-[11px] font-bold transition-all cursor-pointer ${
                targetType === 'worksheet'
                  ? 'bg-amber-500 text-white shadow-sm'
                  : 'text-stone-500 dark:text-stone-400 hover:text-stone-800 dark:hover:text-stone-200'
              }`}
            >
              <span>シート通常形状ボタン</span>
            </button>
          </div>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-12 gap-8 items-start">
        
        {/* 左側：シミュレータープレビュー & コントロール */}
        <div className="lg:col-span-7 space-y-8">
          
          {/* 1. ライブグラフィックスプレビューステージ */}
          <div className="border border-stone-250/55 dark:border-stone-900 rounded-3xl bg-white dark:bg-[#0c0c0c] overflow-hidden shadow-sm relative">
            <div className="flex items-center justify-between px-5 py-4 border-b border-stone-200/50 dark:border-stone-900/45 bg-[#FBF9F6] dark:bg-[#121212]">
              <div className="flex items-center gap-2">
                <span className="w-2.5 h-2.5 rounded-full bg-emerald-500 animate-pulse" />
                <span className="text-[10px] font-bold uppercase tracking-wider font-mono text-stone-550">
                  LIVE_SHAPE_STAGE ({targetType === 'userform' ? 'UserForm Control' : 'Sheet Shape Object'})
                </span>
              </div>
              
              <div className="flex items-center gap-2.5 text-[10px] font-mono text-stone-400">
                <button 
                  onClick={() => setShowGrid(!showGrid)} 
                  className={`px-2 py-1 rounded border transition-colors ${showGrid ? 'bg-indigo-500/10 text-indigo-500 border-indigo-500/20' : 'border-stone-250'}`}
                >
                  GRID_NET:{showGrid ? 'ON' : 'OFF'}
                </button>
                <span>WIDTH: {width}pt ({ptToPx(width)}px)</span>
                <span>HEIGHT: {height}pt ({ptToPx(height)}px)</span>
              </div>
            </div>

            {/* プレビュー本体。VBA環境や擬似Excelグリッドを再現 */}
            <div 
              className={`h-72 flex items-center justify-center relative p-8 transition-colors ${
                showGrid 
                  ? 'bg-[radial-gradient(#e5e5e5_1.2px,transparent_1.2px)] dark:bg-[radial-gradient(#262626_1.2px,transparent_1.2px)] [background-size:16px_16px] bg-[#FAF9F6] dark:bg-[#0d0d0d]' 
                  : 'bg-stone-50 dark:bg-stone-950'
              }`}
            >
              <div className="absolute top-4 left-6 pointer-events-none select-none flex items-center gap-1">
                <LayoutGrid className="w-4 h-4 text-stone-300 dark:text-stone-850" />
                <span className="text-[9px] font-bold font-mono text-stone-300 dark:text-stone-800">
                  {targetType === 'userform' ? 'Me.UserForm_Design' : 'ActiveSheet.GridAlign'}
                </span>
              </div>

              {/* 動的ボタンスタイリングの実写シミュレーション */}
              <motion.button
                layout
                style={{
                  width: `${ptToPx(width)}px`,
                  height: `${ptToPx(height)}px`,
                  backgroundColor: backColor,
                  color: foreColor,
                  fontFamily: fontName,
                  fontSize: `${fontSize * 1.25}px`, // プレビュー用に適宜拡大スケール
                  fontWeight: fontBold ? 'bold' : 'normal',
                  fontStyle: fontItalic ? 'italic' : 'normal',
                  whiteSpace: wordWrap ? 'normal' : 'nowrap',
                  
                  // シート用シェイプの場合のみ、角丸プロパティや国境枠を有効化
                  borderRadius: targetType === 'worksheet' 
                    ? `${(height * roundedAmount * 1.33)}px` 
                    : specialEffect === 0 ? '0px' : '2px', // UserFormのボタンは基本フラットか微角
                  
                  borderWidth: targetType === 'worksheet' ? `${borderWidth}px` : undefined,
                  borderColor: targetType === 'worksheet' ? borderColor : undefined,
                  borderStyle: targetType === 'worksheet' ? 'solid' : undefined,

                  // 影
                  boxShadow: (targetType === 'worksheet' && hasShadow)
                    ? '0 4px 10px rgba(0,0,0,0.22), 0 1px 3px rgba(0,0,0,0.12)'
                    : specialEffect === 1 ? '1px 1.5px 0px #ffffff inset, -1px -1.5px 0px rgba(0,0,0,0.4) inset' // 擬似立体
                    : specialEffect === 2 ? '1px 1.5px 0px rgba(0,0,0,0.5) inset, -1.2px -1.2px 0px #ffffff inset' // 擬似凹み
                    : 'none',
                }}
                className={`transition-all duration-150 flex items-center justify-center text-center px-4 overflow-hidden outline-none ${
                  specialEffect === 3 ? 'border border-[#777] border-b-[#f6f6f6] border-r-[#f6f6f6]' : ''
                }`}
              >
                {btnCaption || 'ボタンテキスト'}
              </motion.button>
            </div>
          </div>

          {/* 2. 形状カスタマイズコントローラー */}
          <div className="p-6 md:p-8 rounded-3xl border border-stone-250/55 dark:border-stone-900 bg-[#FAF7F2]/45 dark:bg-[#121212]/80 backdrop-blur-md shadow-xs grid grid-cols-1 md:grid-cols-2 gap-6 md:gap-8">
            
            {/* 左コラム：テキストと基本寸法、フォント */}
            <div className="space-y-5">
              <h3 className="text-xs font-bold uppercase tracking-widest text-[#8C8A85] flex items-center gap-2 font-mono pb-2 border-b border-stone-250/20">
                <Sliders className="w-3.5 h-3.5" />
                BASIC_PROPERTIES
              </h3>
              
              {/* マクロ名 */}
              <div className="space-y-1.5">
                <label className="text-[10px] font-bold text-stone-550 dark:text-stone-400 font-mono">
                  OBJECT_NAME (.Name)
                </label>
                <input 
                  type="text" 
                  value={btnName} 
                  onChange={(e) => setBtnName(e.target.value.replace(/[^a-zA-Z0-9_]/g, ''))}
                  placeholder="btnAction"
                  className="w-full px-3 py-2 text-xs rounded-xl border border-stone-250 dark:border-stone-850 bg-white dark:bg-[#181818] text-stone-900 dark:text-white font-mono"
                />
              </div>

              {/* キャプション */}
              <div className="space-y-1.5">
                <label className="text-[10px] font-bold text-stone-550 dark:text-stone-400 font-mono">
                  DISPLAY_TEXT (.Caption / .Text)
                </label>
                <input 
                  type="text" 
                  value={btnCaption} 
                  onChange={(e) => setBtnCaption(e.target.value)}
                  placeholder="処理を実行する"
                  className="w-full px-3 py-2 text-xs rounded-xl border border-stone-250 dark:border-stone-850 bg-white dark:bg-[#181818] text-stone-900 dark:text-white"
                />
              </div>

              {/* 寸法 */}
              <div className="grid grid-cols-2 gap-4">
                <div className="space-y-1.5">
                  <label className="text-[10px] font-bold text-stone-550 dark:text-stone-400 font-mono flex justify-between">
                    <span>WIDTH (.Width)</span>
                    <span className="text-stone-400">{width}pt</span>
                  </label>
                  <input 
                    type="range"
                    min="40"
                    max="280"
                    step="5"
                    value={width}
                    onChange={(e) => setWidth(Number(e.target.value))}
                    className="w-full accent-amber-500 bg-stone-200 dark:bg-stone-800 h-1.5 rounded-lg appearance-none"
                  />
                </div>
                <div className="space-y-1.5">
                  <label className="text-[10px] font-bold text-stone-550 dark:text-stone-400 font-mono flex justify-between">
                    <span>HEIGHT (.Height)</span>
                    <span className="text-stone-400">{height}pt</span>
                  </label>
                  <input 
                    type="range"
                    min="20"
                    max="100"
                    step="2"
                    value={height}
                    onChange={(e) => setHeight(Number(e.target.value))}
                    className="w-full accent-amber-500 bg-stone-200 dark:bg-stone-800 h-1.5 rounded-lg appearance-none"
                  />
                </div>
              </div>

              {/* フォント選択ファミリーとサイズ */}
              <div className="space-y-3.5 pt-2">
                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-1.5">
                    <label className="text-[10px] font-bold text-stone-550 dark:text-stone-400 font-mono">
                      FONT_NAME (.Font.Name)
                    </label>
                    <select
                      value={fontName}
                      onChange={(e) => setFontName(e.target.value)}
                      className="w-full px-2 py-1.5 text-xs rounded-xl border border-stone-250 dark:border-stone-850 bg-white dark:bg-[#181818] text-stone-900 dark:text-white"
                    >
                      {COMMON_FONTS.map(f => (
                        <option key={f.value} value={f.value}>{f.label}</option>
                      ))}
                    </select>
                  </div>

                  <div className="space-y-1.5">
                    <label className="text-[10px] font-bold text-stone-550 dark:text-stone-400 font-mono flex justify-between">
                      <span>SIZE (.Font.Size)</span>
                      <span className="text-stone-400">{fontSize}pt</span>
                    </label>
                    <input 
                      type="range"
                      min="8"
                      max="24"
                      step="0.5"
                      value={fontSize}
                      onChange={(e) => setFontSize(Number(e.target.value))}
                      className="w-full accent-amber-500 bg-stone-200 dark:bg-stone-800 h-1.5 rounded-lg appearance-none"
                    />
                  </div>
                </div>

                <div className="flex gap-4">
                  <label className="flex items-center gap-2 cursor-pointer text-xs font-semibold text-stone-700 dark:text-stone-300">
                    <input 
                      type="checkbox" 
                      checked={fontBold}
                      onChange={(e) => setFontBold(e.target.checked)}
                      className="rounded border-stone-300 text-amber-500 focus:ring-amber-500 w-3.5 h-3.5"
                    />
                    <span>太字 (Bold)</span>
                  </label>
                  <label className="flex items-center gap-2 cursor-pointer text-xs font-semibold text-stone-700 dark:text-stone-300">
                    <input 
                      type="checkbox" 
                      checked={fontItalic}
                      onChange={(e) => setFontItalic(e.target.checked)}
                      className="rounded border-stone-300 text-amber-500 focus:ring-amber-500 w-3.5 h-3.5"
                    />
                    <span>斜体 (Italic)</span>
                  </label>
                </div>
              </div>

            </div>

            {/* 右コラム：配色コントロールとターゲット固有設定 */}
            <div className="space-y-5">
              <h3 className="text-xs font-bold uppercase tracking-widest text-[#8C8A85] flex items-center gap-2 font-mono pb-2 border-b border-stone-250/20">
                <Palette className="w-3.5 h-3.5" />
                STYLING_DETAILED
              </h3>

              {/* 背景色プレセット & カスタム */}
              <div className="space-y-2">
                <label className="text-[10px] font-bold text-stone-550 dark:text-stone-400 font-mono">
                  BACK_COLOR (.BackColor / .Fill)
                </label>
                <div className="flex gap-2">
                  <input 
                    type="color" 
                    value={backColor}
                    onChange={(e) => syncPresetColors(e.target.value)}
                    className="w-10 h-8 rounded-lg border border-stone-250 cursor-pointer bg-transparent"
                  />
                  <input 
                    type="text" 
                    value={backColor}
                    onChange={(e) => syncPresetColors(e.target.value)}
                    className="px-2 py-1.5 text-xs rounded-xl border border-stone-250 dark:border-stone-850 bg-white dark:bg-[#181818] w-24 text-stone-900 dark:text-white font-mono"
                  />
                  <span className="text-[10px] text-stone-400 font-mono self-center">
                     {getVbaRgbStr(backColor)}
                  </span>
                </div>
                
                {/* プリセット丸並び */}
                <div className="flex flex-wrap gap-1.5 pt-1.5">
                  {BACK_COLOR_PRESETS.map((p) => (
                    <button
                      key={p.name}
                      onClick={() => syncPresetColors(p.hex)}
                      title={p.name}
                      style={{ backgroundColor: p.hex }}
                      className={`w-5.5 h-5.5 rounded-full border border-stone-350 dark:border-stone-800 relative transition-transform hover:scale-110 cursor-pointer ${
                        backColor.toLowerCase() === p.hex.toLowerCase() ? 'ring-2 ring-amber-500 ring-offset-2 dark:ring-offset-stone-950' : ''
                      }`}
                    >
                      {backColor.toLowerCase() === p.hex.toLowerCase() && (
                        <Check className={`w-3 h-3 absolute inset-0 m-auto ${p.hex === '#FFFFFF' ? 'text-black' : 'text-white'}`} />
                      )}
                    </button>
                  ))}
                </div>
              </div>

              {/* 文字色カラー */}
              <div className="space-y-2">
                <label className="text-[10px] font-bold text-stone-550 dark:text-stone-400 font-mono">
                  FORE_COLOR (.ForeColor)
                </label>
                <div className="flex gap-2">
                  <input 
                    type="color" 
                    value={foreColor}
                    onChange={(e) => setForeColor(e.target.value)}
                    className="w-10 h-8 rounded-lg border border-stone-250 cursor-pointer bg-transparent"
                  />
                  <input 
                    type="text" 
                    value={foreColor}
                    onChange={(e) => setForeColor(e.target.value)}
                    className="px-2 py-1.5 text-xs rounded-xl border border-stone-250 dark:border-stone-850 bg-white dark:bg-[#181818] w-24 text-stone-900 dark:text-white font-mono"
                  />
                </div>
              </div>

              {/* ターゲット固有パラメータ */}
              {targetType === 'worksheet' ? (
                <div className="space-y-4 pt-2 border-t border-stone-250/20">
                  {/* 角丸の強さ */}
                  <div className="space-y-1">
                    <div className="flex justify-between text-[10px] font-bold text-stone-550 dark:text-stone-400 font-mono">
                      <span>ROUNDED_CORNER (.Adjustments(1))</span>
                      <span>{Math.round(roundedAmount * 100)}%</span>
                    </div>
                    <input 
                      type="range"
                      min="0.0"
                      max="0.5"
                      step="0.03"
                      value={roundedAmount}
                      onChange={(e) => setRoundedAmount(Number(e.target.value))}
                      className="w-full accent-indigo-500 bg-stone-200 dark:bg-stone-800 h-1 rounded-lg appearance-none"
                    />
                  </div>

                  {/* ボーダーサイズ・色 */}
                  <div className="grid grid-cols-2 gap-4">
                    <div className="space-y-1">
                      <label className="text-[10px] font-bold text-stone-550 font-mono">
                        BORDER_WIDTH
                      </label>
                      <select
                        value={borderWidth}
                        onChange={(e) => setBorderWidth(Number(e.target.value))}
                        className="w-full px-2 py-1.5 text-xs rounded-xl border border-stone-250 dark:border-stone-850 bg-white dark:bg-[#181818] text-stone-900 dark:text-white font-mono"
                      >
                        <option value="0">0 (なし)</option>
                        <option value="0.75">0.75pt</option>
                        <option value="1.5">1.5pt</option>
                        <option value="2.25">2.25pt</option>
                        <option value="3.5">3.5pt</option>
                      </select>
                    </div>

                    <div className="space-y-1">
                      <label className="text-[10px] font-bold text-stone-550 font-mono">
                        BORDER_COLOR
                      </label>
                      <input 
                        type="color" 
                        value={borderColor}
                        onChange={(e) => setBorderColor(e.target.value)}
                        className="w-full h-8 rounded-xl border border-stone-250 cursor-pointer bg-transparent"
                      />
                    </div>
                  </div>

                  {/* 影チェックボックス */}
                  <label className="flex items-center gap-2 cursor-pointer text-xs font-semibold text-stone-700 dark:text-stone-300">
                    <input 
                      type="checkbox" 
                      checked={hasShadow}
                      onChange={(e) => setHasShadow(e.target.checked)}
                      className="rounded border-stone-300 text-amber-500 focus:ring-amber-500 w-3.5 h-3.5"
                    />
                    <span>ソフトな落ち影 (Shadow) 効果を適用</span>
                  </label>
                </div>
              ) : (
                <div className="space-y-4 pt-2 border-t border-stone-250/20">
                  {/* 立体特殊効果 */}
                  <div className="space-y-1.5">
                    <label className="text-[10px] font-bold text-stone-550 dark:text-stone-400 font-mono">
                      SPECIAL_EFFECT (.SpecialEffect)
                    </label>
                    <select
                      value={specialEffect}
                      onChange={(e) => setSpecialEffect(Number(e.target.value))}
                      className="w-full px-2 py-1.5 text-xs rounded-xl border border-stone-250 dark:border-stone-850 bg-white dark:bg-[#181818] text-stone-900 dark:text-white"
                    >
                      <option value="0">0 - fmSpecialEffectFlat (平坦)</option>
                      <option value="1">1 - fmSpecialEffectRaised (隆起立体)</option>
                      <option value="2">2 - fmSpecialEffectSunken (沈み込み)</option>
                      <option value="3">3 - fmSpecialEffectEtched (彫刻風枠)</option>
                    </select>
                  </div>

                  {/* 字詰めWordWrap */}
                  <label className="flex items-center gap-2 cursor-pointer text-xs font-semibold text-stone-700 dark:text-stone-300">
                    <input 
                      type="checkbox" 
                      checked={wordWrap}
                      onChange={(e) => setWordWrap(e.target.checked)}
                      className="rounded border-stone-300 text-amber-500 focus:ring-amber-500 w-3.5 h-3.5"
                    />
                    <span>自動折り返しを行う (WordWrap)</span>
                  </label>
                </div>
              )}

            </div>

          </div>

        </div>

        {/* 右側：出力された精密VBAソースコード */}
        <div className="lg:col-span-5 space-y-4">
          <div className="flex items-center justify-between border-b border-stone-250/35 dark:border-stone-900/60 pb-3">
            <div className="flex items-center gap-2.5">
              <span className="text-[10px] font-bold px-2.5 py-0.5 rounded-md border font-mono tracking-wider bg-indigo-500/10 text-indigo-500 dark:bg-indigo-400/10 dark:text-indigo-400 border-indigo-500/20">
                AUTO_GENERATED_VBA
              </span>
              <span className="text-xs font-bold text-stone-900 dark:text-white font-mono">
                {targetType === 'userform' ? 'For UserForm' : 'For Worksheet'}
              </span>
            </div>
            
            <button
              onClick={handleCopyCode}
              className="flex items-center gap-1.5 px-3 py-1.5 text-[10px] font-mono font-bold rounded-lg transition-all text-stone-700 hover:text-stone-950 dark:text-stone-300 dark:hover:text-white bg-stone-100 dark:bg-white/[0.04] border border-stone-250/50 dark:border-white/[0.03] cursor-pointer"
            >
              {copied ? <Check className="w-3.5 h-3.5 text-emerald-500" /> : <Copy className="w-3.5 h-3.5" />}
              <span>{copied ? 'VBA_COPIED_OK' : 'COPY_CLIPBOARD'}</span>
            </button>
          </div>

          <div className="p-4 rounded-2xl border border-stone-250/50 dark:border-stone-900 bg-amber-500/5 text-amber-900 dark:text-amber-300 text-[11px] leading-relaxed flex items-start gap-2 select-none">
            <Info className="w-4 h-4 shrink-0 mt-0.5" />
            <p>
              生成されたコードをコピーし、VBE（VBAエディタ）に直接コピー＆ペーストするだけで、Excelの画面解像度やテーマ設定に左右されない完璧な配色ボタン形状を再現可能です。
            </p>
          </div>

          <SyntaxHighlighter code={vbaCode} language="vba" />
        </div>

      </div>

    </div>
  );
};
