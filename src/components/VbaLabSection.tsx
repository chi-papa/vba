import React, { useState, useEffect } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  Cpu, 
  Zap, 
  CheckCircle, 
  AlertTriangle, 
  FileText, 
  Play, 
  Copy, 
  Check, 
  RefreshCw,
  TrendingUp,
  Award,
  BookOpen,
  Trash2,
  History,
  Plus,
  Terminal,
  FileSpreadsheet
} from 'lucide-react';
import { SyntaxHighlighter } from './SyntaxHighlighter';

// Detailed catalog of common VBA errors
interface VbaErrorDefinition {
  code: string;
  name: string;
  english: string;
  reason: string;
  solution: string;
  sampleCode: string;
}

const COMMON_VBA_ERRORS: VbaErrorDefinition[] = [
  {
    code: "9",
    name: "インデックスが有効範囲にありません",
    english: "Subscript out of range",
    reason: "Excelオブジェクトの不一致による最大の原因ナンバーワンです。「Worksheets(\"売上データ\")」などの指定で、エクセル内にその名称のシートが実在しなかったり、文字の前後に余剰な半角・全角スペースが含まれている際によく発生します。また、配列で宣言した次元以上の要素（例: Dim arr(5) に対する arr(6)）を参照した場合にも出ます。",
    solution: "1. 参照先のシート名が完全に一致しているか（半角/全角、大文字/小文字、不要な内包スペースなど）を確認して修正します。\n2. 確実に存在する「インデックス番号（例: Worksheets(1)」を使用、もしくはシートの保護状態を確認します。\n3. 動的配列の境界確認に UBound や LBound を使用してください。",
    sampleCode: `' ❌ NG: 存在しない名称を指定（末尾の無駄なスペース等）
Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets("売上データ ") ' スペース違いでエラー9発生

'  OK: 正確にシート名を合わせる
Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets("売上データ")

' もしくは最初のシートをインデックス指定で安全に取得
Set ws = ThisWorkbook.Worksheets(1)`
  },
  {
    code: "91",
    name: "オブジェクト変数または With ブロック変数が設定されていません",
    english: "Object variable or With block variable not set",
    reason: "WorksheetやRange、ユーザーフォームなどの『オブジェクトオブジェクト（参照型）』を変数にセットする際、VBAの必須予約語である『Set』を忘れて代入を試みたか、Findメソッド等で探し出したセル対象が「Nothing（見つからなかった）」であるのに対象セルのメソッドを処理しようとした時に頻発します。",
    solution: "1. 変数への代入行で「Set 変数名 = ...」と正しく Set 記述が記述されていますか？\n2. Findメソッド等の検索時は、事前に「If Not 変数 Is Nothing Then」の条件式を挟んで Nothing 回避設計にしてください。",
    sampleCode: `' ❌ NG: オブジェクトアサイン時の Set 抜け
Dim ws As Worksheet
ws = ThisWorkbook.Sheets(1) ' Set が抜けているためエラー91

Dim target As Range
Set target = Cells.Find("売上総額")
target.Value = 10000 ' 検索に落ちた（Nothingの）場合にエラー91

'  OK: Setを堅牢に付加し、Nothing防御を行う
Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets(1)

Dim target As Range
Set target = Cells.Find("売上総額")
If Not target Is Nothing Then
    target.Value = 10000
Else
    MsgBox "「売上総額」セルが見つかりませんでした。", vbExclamation
End If`
  },
  {
    code: "13",
    name: "型が一致しません",
    english: "Type mismatch",
    reason: "宣言された変数の型と、そこに実際に代入しようとしたデータの型が相違している場合に発生します。特に、Integer型やDouble型の数値型変数に対し、文字列セルを入力したり、セルの中身がエラー値（#N/A, #VALUE!, #DIV/0! など）を含む状態で直接変数に格納するとこのクラッシュが発生します。",
    solution: "1. 入力値が数値かどうか「IsNumeric(対象)」で事前にバリデーション検算してください。\n2. 動的に変化するセル値は一旦「Variant（バリアント型）」で受け、エラー状態ではないか「IsError(対象)」等で回避分岐します。",
    sampleCode: `' ❌ NG: 文字やエラーセルの数値を直接型決定した変数へ格納
Dim payment As Long
payment = Range("B12").Value ' B12が「請求書未回収」や「#N/A」だとエラー13

'  OK: 安全なVariantでバッファ受けし、検算
Dim rawVal As Variant
rawVal = Range("B12").Value

If IsError(rawVal) Then
    MsgBox "セルの値がエラー表記されています。", vbCritical
ElseIf IsNumeric(rawVal) Then
    Dim payment As Long
    payment = CLng(rawVal)
Else
    MsgBox "セルに数値以外の文字列が入力されています。", vbInformation
End If`
  },
  {
    code: "1004",
    name: "アプリケーション定義またはオブジェクト定義のエラー",
    english: "Application-defined or object-defined error",
    reason: "VBAの不調で最もカバー範囲が広く原因追求が難解なエラーです。最も代表的なトリガーは、「アクティブになっていない別シートのセル範囲を Range.Select でつかもうとした」「行や列に「0」や負数などの限界不正値を渡した（Cells(0, 1) など）」「セルの指定文字列（Range(\"A0\")など）が物理的に破綻している」のいずれかです。",
    solution: "1. セル指定の行条件が 0 以下の数値に陥っていないかループカウンタをデバッグします。\n2. 他シートのセル操作をする際「Select / Activate」は行わず、「Worksheets(\"Sheet1\").Range(\"A1\").Value = ... 」のように完全に『親オブジェクト修飾付き』で記述してください。",
    sampleCode: `' ❌ NG: アクティブではないシートの要素を無理やり Select しようとする
Sheets("集計表").Activate
Sheets("台帳マスタ").Range("C5").Select ' アクティブでない別シートをSelectするとエラー1004

'  OK: 操作を修飾型で物理的にダイレクト制御する (Select不要論)
Sheets("台帳マスタ").Range("C5").Value2 = "転記マスタ"`
  },
  {
    code: "424",
    name: "オブジェクトが必要です",
    english: "Object required",
    reason: "オブジェクト型ではない通常プリミティブ（StringやDouble）に対して、オブジェクトのプロパティやメソッド名を実行しようとした際、あるいはワークシート内のコントロール（TextBoxやComboBoxなど）の名前を打ち間違えて存在しない対象に対してオブジェクトプロトコルを請求した際に発生します。",
    solution: "1. 参照オブジェクト名のアルファベット綴り（つづり）ミスが無いかエディタで確認します。\n2. 変数のDim宣言タイプが正しいか検証します。",
    sampleCode: `' ❌ NG: 通常データ型変数に対して Range や Sheet のような操作を実行
Dim userName As String
userName = "鈴木"
userName.Select ' userNameは単なる文字列のためオブジェクトが必要です(エラー424)

'  OK: 正確に Range オブジェクトに対して実行
Dim userRange As Range
Set userRange = Range("A2")
userRange.Value = "鈴木"`
  },
  {
    code: "438",
    name: "オブジェクトは、このプロパティまたはメソッドをサポートしていません",
    english: "Object doesn't support this property or method",
    reason: "ほぼすべての要因は、プロパティやメソッドの「英単語のスペルタイポ」です。（例: Range(\"A1\").Value と書くべき場所を `.Valuu` や `.Valuee` とする）。もしくは、Rangeクラスには存在する命令をWorksheetオブジェクトに対して無理やり実行させた場合なども該当します。",
    solution: "1. VBAの予約語が自動で大文字・小文字に置換されるか確認し、スペルミスを見つけます。\n2. 操作しようとしているオブジェクトが、そのメソッドを公式にサポートしているかOfficeドキュメントを確認します。",
    sampleCode: `' ❌ NG: 存在定義のない間違ったスペルのプロパティ呼び出し
Dim cell As Range
Set cell = Range("C3")
cell.TextLength = 20 ' Range には TextLength というプロパティはないためエラー438

'  OK: 正しいプロパティ名（.Value / .Text など）を使用
cell.Value = "20文字以上のテキスト"`
  },
  {
    code: "COMPILE_VARIABLE",
    name: "コンパイル エラー: 変数が定義されていません",
    english: "Variable not defined",
    reason: "モジュールの先頭に変数定義の強制宣言『Option Explicit』が記述されている健全なプログラムにおいて、宣言（Dimステートメント）を行なっていない全く新しい変数をスペルミス等でいきなり処理の中に導入したため、コンパイラから差し止められた状態です。",
    solution: "1. 変数名に打ち間違い（例： targetData と tragetData）がないか注意深く確認します。\n2. 変数を使用する前に、必ず「Dim 変数名 As 型」でスコープ宣言を記述してください。",
    sampleCode: `' Option Explicit 指定状態
Dim currentTotal As Long
curntTotal = 50000 ' ❌ NG: curntTotal は変数定義されていないためコンパイルエラー

'  OK: スペルを正しい変数名へ修正する
currentTotal = 50000`
  }
];

export const VbaLabSection: React.FC = () => {
  // Navigation Tabs for dynamic lab
  const [activeLabTab, setActiveLabTab] = useState<'boiler' | 'audit' | 'solver' | 'roi'>('boiler');

  // Tool 1: Preamble Generator State
  const [subName, setSubName] = useState('My_Automated_Process');
  const [useScreenUpdating, setUseScreenUpdating] = useState(true);
  const [useCalculation, setUseCalculation] = useState(true);
  const [useEnableEvents, setUseEnableEvents] = useState(true);
  const [useErrorHandler, setUseErrorHandler] = useState(true);
  const [useStatusMsg, setUseStatusMsg] = useState(true);
  const [copiedBoilerplate, setCopiedBoilerplate] = useState(false);

  // Tool 2: Code Analyzer State
  const [auditCode, setAuditCode] = useState(`Sub TransferData()
    ' Excelシート間のデータ転記マクロ
    Sheets("Sheet1").Select
    Range("A1:D10").Select
    Selection.Copy
    Sheets("Sheet2").Select
    Range("A1").Select
    ActiveSheet.Paste
End Sub`);
  const [auditResult, setAuditResult] = useState<{
    score: string;
    scoreColor: string;
    scoreBg: string;
    findings: { status: 'pass' | 'fail' | 'warn'; title: string; desc: string }[];
    analyzed: boolean;
  } | null>(null);

  // Tool 3: Error Solver & History State
  const [errorInputText, setErrorInputText] = useState('実行時エラー \'9\': インデックスが有効範囲にありません');
  const [solvedAnalysis, setSolvedAnalysis] = useState<VbaErrorDefinition | null>(null);
  const [savedErrorLogs, setSavedErrorLogs] = useState<{
    id: string;
    errorCode: string;
    errorTitle: string;
    timestamp: string;
    notes: string;
    isFixed: boolean;
  }[]>([]);

  // Tool 4: Savings Simulator State
  const [minutesPerTask, setMinutesPerTask] = useState(45);
  const [frequencyPerWeek, setFrequencyPerWeek] = useState(5);
  const [hourlyWage, setHourlyWage] = useState(2500);

  // Load Saved Error Logbook from localStorage
  useEffect(() => {
    try {
      const stored = localStorage.getItem('vba_sandbox_error_logbook');
      if (stored) {
        setSavedErrorLogs(JSON.parse(stored));
      } else {
        // Initial mock logs to show usability
        const initialMock = [
          {
            id: 'mock-1',
            errorCode: '9',
            errorTitle: 'インデックスが有効範囲にありません',
            timestamp: '06/14 11:30',
            notes: '台帳転記のシート名の全角スペース違いを修正し解決。',
            isFixed: true,
          },
          {
            id: 'mock-2',
            errorCode: '91',
            errorTitle: 'オブジェクト変数または With ブロック変数が設定されていません',
            timestamp: '06/14 14:15',
            notes: 'FindセルのNothingチェックを忘れていたため、安全判定を追加した。',
            isFixed: true,
          }
        ];
        setSavedErrorLogs(initialMock);
        localStorage.setItem('vba_sandbox_error_logbook', JSON.stringify(initialMock));
      }
    } catch (e) {
      console.warn('LocalStorage is offline or restricted.', e);
    }
  }, []);

  // Sync Saved Error Logbook
  const saveLogsToStorage = (updatedLogs: typeof savedErrorLogs) => {
    setSavedErrorLogs(updatedLogs);
    try {
      localStorage.setItem('vba_sandbox_error_logbook', JSON.stringify(updatedLogs));
    } catch (e) {
      console.warn(e);
    }
  };

  // Generate Boilerplate Code
  const getBoilerplateCode = () => {
    let code = `Option Explicit\n\n`;
    code += `Sub ${subName || 'Optimize_Process'}()\n`;
    code += `    ' =========================================================\n`;
    code += `    ' CURATED HIGH-SPEED ENGINE WRAPPER\n`;
    code += `    ' =========================================================\n`;
    
    if (useScreenUpdating || useCalculation || useEnableEvents) {
      code += `    ' 1. 高速化プロトコル：バックグラウンド処理の最適化\n`;
    }
    if (useScreenUpdating) code += `    Application.ScreenUpdating = False\n`;
    if (useCalculation) code += `    Application.Calculation = xlCalculationManual\n`;
    if (useEnableEvents) code += `    Application.EnableEvents = False\n`;
    code += `\n`;

    if (useErrorHandler) {
      code += `    On Error GoTo ErrorHandler\n\n`;
    }

    code += `    ' ---------------------------------------------------------\n`;
    code += `    ' [メイン処理] ここへ自動化ロジックを記述します\n`;
    code += `    ' ---------------------------------------------------------\n`;
    code += `    Dim wsSource As Worksheet, wsTarget As Worksheet\n`;
    code += `    Set wsSource = ThisWorkbook.Sheets(1)\n`;
    code += `    ' (例: wsSource.Range("A1").Value = "Processing Ready")\n\n`;
    code += `    ' ---------------------------------------------------------\n`;

    // Normal Exit
    if (useErrorHandler) {
      code += `\nCleanExit:\n`;
    }
    if (useScreenUpdating || useCalculation || useEnableEvents) {
      code += `    ' 2. 環境パラメータの完全復元フロー\n`;
    }
    if (useScreenUpdating) code += `    Application.ScreenUpdating = True\n`;
    if (useCalculation) code += `    Application.Calculation = xlCalculationAutomatic\n`;
    if (useEnableEvents) code += `    Application.EnableEvents = True\n`;

    if (useStatusMsg) {
      code += `\n    MsgBox "マクラ処理が正常に完了しました！", vbInformation + vbOKOnly, "ENGINE STATUS"\n`;
    }
    code += `    Exit Sub\n`;

    if (useErrorHandler) {
      code += `\nErrorHandler:\n`;
      code += `    ' 3. ロギング ＆ 復元フェイルセーフ\n`;
      code += `    MsgBox "実行中に致命的なエラーが発生しました。" & vbCrLf & _\n`;
      code += `           "エラー内容: " & Err.Description, vbCritical, "SYS_FATAL_ERROR"\n`;
      code += `    Resume CleanExit\n`;
    }

    code += `End Sub`;
    return code;
  };

  const currentGeneratedCode = getBoilerplateCode();

  const handleCopyBoilerplate = async () => {
    try {
      await navigator.clipboard.writeText(currentGeneratedCode);
      setCopiedBoilerplate(true);
      setTimeout(() => setCopiedBoilerplate(false), 2000);
    } catch (err) {
      console.error(err);
    }
  };

  // Run Code Analysis
  const performAnalysis = () => {
    const findings: { status: 'pass' | 'fail' | 'warn'; title: string; desc: string }[] = [];
    const lower = auditCode.toLowerCase();

    // 1. Option Explicit
    if (lower.indexOf('option explicit') !== -1) {
      findings.push({
        status: 'pass',
        title: 'Option Explicit 宣言検出',
        desc: '型宣言が強制されており、変数名のうち間違いによるバグを未然に防止します。素晴らしい設計です。'
      });
    } else {
      findings.push({
        status: 'fail',
        title: 'Option Explicit 未検出',
        desc: 'モジュール先頭への記述を推奨。変数の誤字脱字によるサイレントバグ（値が代入されないなど）の原因になります。'
      });
    }

    // 2. Select, Activate
    const selectCount = (lower.match(/\.select|\.activate|selection\./g) || []).length;
    if (selectCount === 0) {
      findings.push({
        status: 'pass',
        title: 'Select非依存・直接参照設計',
        desc: '画面の切り替えラグ（.Select）を伴わないダイレクトなデータアクセスが行われています。最速動作の基本です。'
      });
    } else {
      findings.push({
        status: 'warn',
        title: `警告: Select/Activate の多用 (${selectCount}件)`,
        desc: 'セルやシートを物理選択する処理は極めて低速です。Range("A1").Value = ... のように直接代入する設計に切り替えましょう。'
      });
    }

    // 3. ScreenUpdating
    if (lower.indexOf('screenupdating') !== -1) {
      findings.push({
        status: 'pass',
        title: 'ScreenUpdating 制御プロトコル',
        desc: '画面のちらつき・再描画を制御して実行パフォーマンスを最大限引き出せています。'
      });
    } else {
      findings.push({
        status: 'warn',
        title: 'ScreenUpdating 未制御',
        desc: '処理中にエクセル画面がチカチカ再描画され、大きなロスが生じます。Application.ScreenUpdating = Falseを導入してください。'
      });
    }

    // 4. Object Release
    const setDeclarations = (lower.match(/set\s+\w+\s*=/g) || []).length;
    const releaseDeclarations = (lower.match(/set\s+\w+\s*=\s*nothing/g) || []).length;
    if (setDeclarations > 0) {
      if (releaseDeclarations >= setDeclarations) {
        findings.push({
          status: 'pass',
          title: 'メモリリーク対策 (Set = Nothing) 記述済み',
          desc: '生成した外部オブジェクトが適切にメモリから解放されています。オフィスシステムに優しい設計です。'
        });
      } else {
        findings.push({
          status: 'warn',
          title: 'オブジェクト開放の漏れ懸念',
          desc: 'オブジェクト数（Set）に対して Nothing による解放が不足している可能性があります。処理末尾で Nothing にセットして解放しましょう。'
        });
      }
    } else {
      findings.push({
        status: 'pass',
        title: 'メモリ解放考慮なし（単純処理）',
        desc: '追加オブジェクト変数が使用されていないため、メモリ蓄積の恐れはありません。'
      });
    }

    // 5. Error GoTo
    if (lower.indexOf('on error') !== -1) {
      findings.push({
        status: 'pass',
        title: '例外処理 (Error Handling) 実装済み',
        desc: '想定外のエラーによりエクセルが異常フリーズして強制終了するのを防ぐ監視網が整備されています。'
      });
    } else {
      findings.push({
        status: 'warn',
        title: 'On Error 例外監視の不備',
        desc: '実行エラー時にVBA警告ダイアログが直接表示されマクロが中断します。On Error GoTo でのエラー回避が推奨されます。'
      });
    }

    // Determine Score
    const fails = findings.filter(f => f.status === 'fail').length;
    const warns = findings.filter(f => f.status === 'warn').length;

    let score = 'A+';
    let scoreBg = 'bg-emerald-100/60 dark:bg-emerald-950/20';
    let scoreColor = 'text-emerald-600 dark:text-emerald-400';

    if (fails > 0 || warns > 2) {
      score = 'B';
      scoreBg = 'bg-amber-100/60 dark:bg-amber-950/20';
      scoreColor = 'text-amber-600 dark:text-amber-400';
    }
    if (fails > 1 || warns > 3) {
      score = 'C';
      scoreBg = 'bg-orange-100/60 dark:bg-orange-950/20';
      scoreColor = 'text-orange-600';
    }
    if (fails > 2) {
      score = 'D-';
      scoreBg = 'bg-rose-100/60 dark:bg-rose-950/20';
      scoreColor = 'text-rose-600';
    }

    setAuditResult({
      score,
      scoreBg,
      scoreColor,
      findings,
      analyzed: true
    });
  };

  // Perform Error Analysis (Solver)
  const performErrorAnalysis = () => {
    if (!errorInputText.trim()) {
      setSolvedAnalysis(null);
      return;
    }

    const norm = errorInputText.toLowerCase();
    
    // Fuzzy matching against error definitions
    let match = COMMON_VBA_ERRORS.find(err => {
      // Direct numeric check (e.g. "9", "91")
      if (err.code !== "COMPILE_VARIABLE") {
        const regex = new RegExp(`\\b${err.code}\\b`);
        if (regex.test(norm)) return true;
      }
      // Name keyword check
      if (norm.includes(err.name.toLowerCase()) || norm.includes(err.english.toLowerCase())) {
        return true;
      }
      return false;
    });

    // Special checks for verbal descriptions
    if (!match) {
      if (norm.includes('有効範囲') || norm.includes('シート名') || norm.includes('インデックス')) {
        match = COMMON_VBA_ERRORS[0]; // Error 9
      } else if (norm.includes('オブジェクト変数') || norm.includes('set') || norm.includes('nothing') || norm.includes('with block')) {
        match = COMMON_VBA_ERRORS[1]; // Error 91
      } else if (norm.includes('型') || norm.includes('まみれ') || norm.includes('mismatch') || norm.includes('一致しません')) {
        match = COMMON_VBA_ERRORS[2]; // Error 13
      } else if (norm.includes('定義のエラー') || norm.includes('1004')) {
        match = COMMON_VBA_ERRORS[3]; // Error 1004
      } else if (norm.includes('オブジェクトが必要') || norm.includes('424')) {
        match = COMMON_VBA_ERRORS[4]; // Error 424
      } else if (norm.includes('サポート') || norm.includes('438') || norm.includes('メソッド')) {
        match = COMMON_VBA_ERRORS[5]; // Error 438
      } else if (norm.includes('未定義') || norm.includes('定義されていません') || norm.includes('explicit')) {
        match = COMMON_VBA_ERRORS[6]; // Compile 
      }
    }

    if (match) {
      setSolvedAnalysis(match);
    } else {
      // Default fuzzy analysis for unknown errors
      setSolvedAnalysis({
        code: "GEN_SYS",
        name: "汎用エラー・その他のバグ可能性",
        english: "Unclassified System Exception",
        reason: "入力されたログから直接のVBAコードナンバーを断定できませんでした。しかし、VBA開発の初期段階で生じるフリーズ・クラッシュの多くは『セルの誤指定』『シート保護の競合』、または『オブジェクト型変数宣言時の「Set」宣言漏れ』から発生します。",
        solution: "1. エラー行に黄色ハイライトが入っている時に「値を格納するはずの変数」にカーソルを合わせ、値が Nothing 又は Empty になっていないか確認します。\n2. Excelシートの保護（閲覧制限）がかかっていると更新しようとした瞬間エラーになりますのでロックを事前に解除します。\n3. 「Option Explicit」を先頭行に入れて「コンパイル」を事前に呼び出してください。",
        sampleCode: `' 汎用デバッグ対策：
' 各関数の最上位に、現在対象としているシートの情報を出力して現在地を正確にトラッキング
Debug.Print "Current active sheet is: " & ActiveSheet.Name
Debug.Print "Is target cell content error? " & IsError(ActiveCell.Value)`
      });
    }
  };

  // Add parsed error solver output directly to persistent history logbook
  const handleSaveToLogbook = () => {
    if (!solvedAnalysis) return;
    const newLog = {
      id: 'err-' + Date.now(),
      errorCode: solvedAnalysis.code,
      errorTitle: solvedAnalysis.name,
      timestamp: new Date().toLocaleString('ja-JP', { month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' }),
      notes: 'ログ解析から原因を解明「' + solvedAnalysis.english + '」。',
      isFixed: false,
    };
    const updated = [newLog, ...savedErrorLogs];
    saveLogsToStorage(updated);
  };

  // Toggle solved state of error history logs
  const handleToggleLogFixed = (id: string) => {
    const updated = savedErrorLogs.map(log => 
      log.id === id ? { ...log, isFixed: !log.isFixed } : log
    );
    saveLogsToStorage(updated);
  };

  // Delete log entry
  const handleDeleteLog = (id: string) => {
    const updated = savedErrorLogs.filter(log => log.id !== id);
    saveLogsToStorage(updated);
  };

  useEffect(() => {
    // Perform standard initial states
    performAnalysis();
    performErrorAnalysis();
  }, []);

  // Simulator value calculators
  const yearlyTasks = frequencyPerWeek * 52;
  const timeSavedPerYearHours = parseFloat(((minutesPerTask * yearlyTasks) / 60).toFixed(1));
  const moneySavedPerYear = Math.round(timeSavedPerYearHours * hourlyWage);

  return (
    <div className="space-y-12">
      
      {/* イントロバナー - 洗練されたコントラストで描く未来調ヘッダー */}
      <div className="p-6 md:p-8 rounded-2xl border border-stone-250/50 dark:border-stone-900 bg-[#FAF7F2]/90 dark:bg-[#121212]/85 backdrop-blur-md relative overflow-hidden group interact-glow select-none">
        <div className="absolute top-0 left-0 w-3.5 h-[1px] bg-stone-300 dark:bg-stone-750 pointer-events-none" />
        <div className="absolute top-0 left-0 h-3.5 w-[1px] bg-stone-300 dark:bg-stone-750 pointer-events-none" />
        <div className="absolute top-0 right-0 w-3.5 h-[1px] bg-stone-300 dark:bg-stone-750 pointer-events-none" />
        <div className="absolute top-0 right-0 h-3.5 w-[1px] bg-stone-300 dark:bg-stone-750 pointer-events-none" />
        
        <div className="flex gap-4 md:gap-5 items-start relative z-10">
          <div className="p-3.5 bg-stone-950 dark:bg-stone-200 text-[#FAF9F6] dark:text-[#101010] rounded-xl shadow-lg shrink-0">
            <Cpu className="w-5.5 h-5.5 animate-tech-pulse text-stone-100 dark:text-stone-900" />
          </div>
          <div className="space-y-1.5">
            <h3 className="text-xl md:text-2xl font-bold text-stone-950 dark:text-stone-100 font-display tracking-wide flex items-center gap-2">
              <span>VBA Optimization Interactive Lab</span>
              <span className="text-[9px] font-mono font-bold bg-[#1A1A1A] text-white dark:bg-stone-850 dark:text-stone-300 px-1.5 py-0.5 rounded border border-transparent dark:border-stone-700 tracking-widest uppercase">
                LAB_MODULE_V1.1
              </span>
            </h3>
            <p className="text-[#595551] dark:text-stone-400 text-xs md:text-sm leading-relaxed font-sans font-medium">
              マクロ開発における「書き方による処理速度の差」や「予測不可能なエラーを回避する防御ロジック」を、直感的かつ高い視覚コントラストで体験できるインタラクティブ機能群です。
              最適化、品質監査、そしてリアルタイムエラー解読ログにより、あなたの開発をハイエンドに引き上げます。
            </p>
          </div>
        </div>
      </div>

      {/* 4つのツールを切り替えるセグメントコントロール（洗練されたマテリアルタブ） */}
      <div className="border border-stone-250/50 dark:border-stone-900 rounded-2xl bg-[#EFECE4]/50 dark:bg-[#121212]/40 p-1.5 flex flex-wrap gap-1.5 select-none font-sans shadow-sm">
        <button
          onClick={() => setActiveLabTab('boiler')}
          className={`flex-1 min-w-[125px] flex items-center justify-center gap-2 py-2.5 rounded-xl text-xs font-bold transition-all cursor-pointer ${
            activeLabTab === 'boiler'
              ? 'bg-[#FAF7F2] dark:bg-stone-850 text-stone-950 dark:text-white shadow-sm'
              : 'text-stone-600 hover:text-stone-950 dark:text-stone-400 dark:hover:text-stone-200'
          }`}
        >
          <Zap className="w-3.5 h-3.5 text-amber-500" />
          <span>🚀 高速化構築</span>
        </button>
        <button
          onClick={() => {
            setActiveLabTab('audit');
            performAnalysis();
          }}
          className={`flex-1 min-w-[125px] flex items-center justify-center gap-2 py-2.5 rounded-xl text-xs font-bold transition-all cursor-pointer ${
            activeLabTab === 'audit'
              ? 'bg-[#FAF7F2] dark:bg-stone-850 text-stone-950 dark:text-white shadow-sm'
              : 'text-stone-600 hover:text-stone-950 dark:text-stone-400 dark:hover:text-stone-200'
          }`}
        >
          <FileText className="w-3.5 h-3.5 text-stone-500" />
          <span>🔍 コード品質監査</span>
        </button>
        <button
          onClick={() => {
            setActiveLabTab('solver');
            performErrorAnalysis();
          }}
          className={`flex-1 min-w-[125px] flex items-center justify-center gap-2 py-2.5 rounded-xl text-xs font-bold transition-all cursor-pointer ${
            activeLabTab === 'solver'
              ? 'bg-[#FAF7F2] dark:bg-stone-850 text-stone-950 dark:text-white shadow-sm'
              : 'text-stone-600 hover:text-stone-950 dark:text-stone-400 dark:hover:text-stone-200'
          }`}
        >
          <Terminal className="w-3.5 h-3.5 text-indigo-500" />
          <span>🛠️ エラー解決＆ログ</span>
        </button>
        <button
          onClick={() => setActiveLabTab('roi')}
          className={`flex-1 min-w-[125px] flex items-center justify-center gap-2 py-2.5 rounded-xl text-xs font-bold transition-all cursor-pointer ${
            activeLabTab === 'roi'
              ? 'bg-[#FAF7F2] dark:bg-stone-850 text-stone-950 dark:text-white shadow-sm'
              : 'text-stone-600 hover:text-stone-950 dark:text-stone-400 dark:hover:text-stone-200'
          }`}
        >
          <TrendingUp className="w-3.5 h-3.5 text-emerald-500" />
          <span>📈 投資対効果 (ROI)</span>
        </button>
      </div>

      {/* タブ切り替え表示エリア */}
      <AnimatePresence mode="wait">
        
        {/* ================= TAB 1: BOILERPLATE OPTIMIZER ================= */}
        {activeLabTab === 'boiler' && (
          <motion.div
            key="boiler-tab"
            initial={{ opacity: 0, y: 15 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: -15 }}
            transition={{ duration: 0.2 }}
            className="grid grid-cols-1 lg:grid-cols-12 gap-8"
          >
            {/* Setting controller */}
            <div className="lg:col-span-5 p-6 rounded-2xl border border-stone-250/50 dark:border-stone-900 bg-[#FAF7F2]/90 dark:bg-[#121212]/90 backdrop-blur-sm space-y-6 relative flex flex-col justify-between shadow-sm">
              <div className="absolute top-0 left-0 w-2.5 h-[1px] bg-stone-300 dark:bg-stone-700 pointer-events-none" />
              <div className="absolute top-0 left-0 h-2.5 w-[1px] bg-stone-300 dark:bg-stone-700 pointer-events-none" />
              
              <div className="space-y-4">
                <div className="flex items-center gap-2">
                  <span className="p-1.5 bg-[#FAF7F2]/40 rounded-lg text-amber-500">
                    <Zap className="w-4 h-4 animate-pulse" />
                  </span>
                  <h3 className="text-sm md:text-base font-bold font-display tracking-tight text-stone-950 dark:text-white uppercase select-none">
                    プロトコル自動追加型テンプレート構成
                  </h3>
                </div>
                <p className="text-stone-600 dark:text-stone-400 text-xs font-sans font-medium leading-relaxed">
                  トグルスイッチを操作することで、お好みの「高速化プロトコル」と「防御網」をマクロへ組み込んだ極めて頑丈なVBAコード雛形を構築します。
                </p>

                <div className="space-y-4 pt-2">
                  {/* Process Name Input */}
                  <div className="space-y-1.5">
                    <label className="text-[10px] uppercase font-mono tracking-widest text-stone-500 dark:text-stone-500 font-bold block select-none">
                      SUB_PROCEDURE_NAME (プロシージャ名)
                    </label>
                    <input
                      type="text"
                      value={subName}
                      onChange={(e) => setSubName(e.target.value.replace(/[^a-zA-Z0-9_]/g, '_'))}
                      placeholder="My_Automated_Process"
                      className="w-full px-3 py-2 text-xs rounded-xl border border-stone-200 dark:border-stone-850 bg-stone-50/50 dark:bg-stone-950 text-stone-950 dark:text-stone-100 placeholder-stone-400 focus:outline-none focus:ring-1 focus:ring-stone-400 font-mono transition-colors"
                    />
                  </div>

                  {/* Toggles */}
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-1 select-none">
                    
                    <label className="flex items-center justify-between p-3 rounded-xl border border-stone-200/50 dark:border-stone-850/50 bg-[#EFECE4]/30 dark:bg-stone-950/20 cursor-pointer hover:border-stone-300 dark:hover:border-stone-800 transition-all">
                      <div className="space-y-0.5">
                        <span className="text-xs font-bold text-stone-900 dark:text-stone-200 block font-display">画面更新の停止</span>
                        <span className="text-[10px] text-stone-500 dark:text-stone-400 block font-mono">ScreenUpdating</span>
                      </div>
                      <input
                        type="checkbox"
                        checked={useScreenUpdating}
                        onChange={(e) => setUseScreenUpdating(e.target.checked)}
                        className="w-4 h-4 accent-stone-900 dark:accent-stone-300 cursor-pointer rounded"
                      />
                    </label>

                    <label className="flex items-center justify-between p-3 rounded-xl border border-stone-200/50 dark:border-stone-850/50 bg-[#EFECE4]/30 dark:bg-stone-950/20 cursor-pointer hover:border-stone-300 dark:hover:border-stone-800 transition-all">
                      <div className="space-y-0.5">
                        <span className="text-xs font-bold text-stone-900 dark:text-stone-200 block font-display">再計算の手動化</span>
                        <span className="text-[10px] text-stone-500 dark:text-stone-400 block font-mono">Calculation=Manual</span>
                      </div>
                      <input
                        type="checkbox"
                        checked={useCalculation}
                        onChange={(e) => setUseCalculation(e.target.checked)}
                        className="w-4 h-4 accent-stone-900 dark:accent-stone-300 cursor-pointer rounded"
                      />
                    </label>

                    <label className="flex items-center justify-between p-3 rounded-xl border border-stone-200/50 dark:border-stone-850/50 bg-[#EFECE4]/30 dark:bg-stone-950/20 cursor-pointer hover:border-stone-300 dark:hover:border-stone-800 transition-all">
                      <div className="space-y-0.5">
                        <span className="text-xs font-bold text-stone-900 dark:text-stone-200 block font-display">イベントの抑制</span>
                        <span className="text-[10px] text-stone-500 dark:text-stone-400 block font-mono">EnableEvents</span>
                      </div>
                      <input
                        type="checkbox"
                        checked={useEnableEvents}
                        onChange={(e) => setUseEnableEvents(e.target.checked)}
                        className="w-4 h-4 accent-stone-900 dark:accent-stone-300 cursor-pointer rounded"
                      />
                    </label>

                    <label className="flex items-center justify-between p-3 rounded-xl border border-stone-200/50 dark:border-stone-850/50 bg-[#EFECE4]/30 dark:bg-stone-950/20 cursor-pointer hover:border-stone-300 dark:hover:border-stone-800 transition-all">
                      <div className="space-y-0.5">
                        <span className="text-xs font-bold text-stone-900 dark:text-stone-200 block font-display">例外検知網の整備</span>
                        <span className="text-[10px] text-stone-500 dark:text-stone-400 block font-mono">On Error GoTo</span>
                      </div>
                      <input
                        type="checkbox"
                        checked={useErrorHandler}
                        onChange={(e) => setUseErrorHandler(e.target.checked)}
                        className="w-4 h-4 accent-stone-900 dark:accent-stone-300 cursor-pointer rounded"
                      />
                    </label>

                    <label className="flex items-center justify-between p-3 rounded-xl border border-stone-200/50 dark:border-stone-850/50 bg-[#EFECE4]/30 dark:bg-stone-950/20 cursor-pointer hover:border-stone-300 dark:hover:border-stone-800 transition-all sm:col-span-2">
                      <div className="space-y-0.5">
                        <span className="text-xs font-bold text-stone-900 dark:text-stone-200 block font-display">終了ステータスダイアログ</span>
                        <span className="text-[10px] text-stone-500 dark:text-stone-400 block font-mono">MsgBox Exit Info</span>
                      </div>
                      <input
                        type="checkbox"
                        checked={useStatusMsg}
                        onChange={(e) => setUseStatusMsg(e.target.checked)}
                        className="w-4 h-4 accent-stone-900 dark:accent-stone-300 cursor-pointer rounded"
                      />
                    </label>

                  </div>
                </div>
              </div>

              <div className="pt-4 border-t border-stone-200/50 dark:border-stone-900 font-sans text-[11px] text-stone-500 leading-relaxed space-y-2">
                <span className="font-bold flex items-center gap-1.5 text-stone-750 dark:text-stone-300 select-none">// SYSTEM PROTOCOL TIPS:</span>
                <p>画面更新の停止と手動計算化を同時に行なうだけで、数万件セルの処理ラグが最大で元の10%以下の時間に短縮されます。</p>
              </div>
            </div>

            {/* Generated Code display */}
            <div className="lg:col-span-7 flex flex-col justify-between p-6 rounded-2xl border border-stone-250/50 dark:border-stone-900 bg-[#FAF7F2]/90 dark:bg-[#121212]/90 backdrop-blur-sm relative h-full">
              <div className="absolute top-0 right-0 w-2.5 h-[1px] bg-stone-300 dark:bg-stone-700 pointer-events-none" />
              <div className="absolute top-0 right-0 h-2.5 w-[1px] bg-stone-300 dark:bg-stone-700 pointer-events-none" />

              <div className="space-y-3 font-mono h-full flex flex-col justify-between">
                <div className="flex items-center justify-between select-none">
                  <span className="text-[10px] font-bold text-stone-500 uppercase tracking-widest font-mono">
                    BOILERPLATE_PRECISE_OUTPUT
                  </span>
                  <button
                    onClick={handleCopyBoilerplate}
                    className="flex items-center gap-1.5 text-[10px] font-bold text-stone-700 dark:text-stone-300 hover:text-stone-950 dark:hover:text-[#FAF7F2] transition-colors cursor-pointer"
                  >
                    {copiedBoilerplate ? (
                      <>
                        <Check className="w-3.5 h-3.5 text-emerald-500" />
                        <span className="text-emerald-500 tracking-wider">COPIED_OK</span>
                      </>
                    ) : (
                      <>
                        <Copy className="w-3.5 h-3.5" />
                        <span>COPY_BOILERPLATE</span>
                      </>
                    )}
                  </button>
                </div>
                
                <div className="rounded-xl border border-stone-200/70 dark:border-stone-900 overflow-hidden font-mono text-xs flex-1 max-h-[480px] overflow-y-auto mt-2">
                  <SyntaxHighlighter code={currentGeneratedCode} language="vba" id="vba-optimised-boilerplate-tab" />
                </div>
              </div>
            </div>
          </motion.div>
        )}

        {/* ================= TAB 2: CODE QUALITY STATIC AUDIT ================= */}
        {activeLabTab === 'audit' && (
          <motion.div
            key="audit-tab"
            initial={{ opacity: 0, y: 15 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: -15 }}
            transition={{ duration: 0.2 }}
            className="grid grid-cols-1 lg:grid-cols-12 gap-8"
          >
            {/* Input target panel */}
            <div className="lg:col-span-6 p-6 rounded-2xl border border-stone-250/50 dark:border-stone-900 bg-[#FAF7F2]/90 dark:bg-[#121212]/90 backdrop-blur-sm space-y-4 relative flex flex-col justify-between shadow-sm">
              <div className="absolute top-0 left-0 w-2.5 h-[1px] bg-stone-300 dark:bg-stone-700 pointer-events-none" />
              <div className="absolute top-0 left-0 h-2.5 w-[1px] bg-stone-300 dark:bg-stone-700 pointer-events-none" />

              <div className="space-y-4">
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-2">
                    <span className="p-1.5 bg-[#FAF7F2]/40 rounded-lg text-stone-500">
                      <FileText className="w-4 h-4" />
                    </span>
                    <h3 className="text-sm md:text-base font-bold font-display tracking-tight text-stone-950 dark:text-white uppercase select-none">
                      VBA設計品質評価システム
                    </h3>
                  </div>
                  <span className="text-[9px] font-mono font-bold tracking-widest text-stone-400 select-none">STATIC_SCANNER</span>
                </div>
                <p className="text-[#595551] dark:text-stone-400 text-xs font-sans font-medium leading-relaxed">
                  作成中または社内の共有マクロのソースコードを貼り付け、ボタンを押します。処理遅延を招く「.Select」や変数不正などを瞬間的に全スキャンします。
                </p>

                <div className="space-y-2">
                  <textarea
                    value={auditCode}
                    onChange={(e) => setAuditCode(e.target.value)}
                    placeholder="Sub ...  End Sub マクロコードを貼り付けてください"
                    className="w-full h-[260px] px-3 py-2.5 text-xs font-mono rounded-xl border border-stone-200 dark:border-stone-850 bg-stone-50/50 dark:bg-stone-950 text-stone-950 dark:text-stone-100 placeholder-stone-500 focus:outline-none focus:ring-1 focus:ring-stone-400 leading-relaxed"
                  />
                  <button
                    onClick={performAnalysis}
                    className="w-full py-3 bg-stone-950 hover:bg-stone-900 dark:bg-stone-100 dark:hover:bg-stone-200 text-[#FAF9F6] dark:text-[#101010] font-bold text-xs sm:text-sm font-display rounded-xl tracking-widest flex items-center justify-center gap-2 cursor-pointer transition-all active:scale-[0.99] shadow-md select-none mt-2"
                  >
                    <RefreshCw className="w-4 h-4 animate-spin-slow" />
                    マクロを高速診断・最適化点をチェック
                  </button>
                </div>
              </div>
            </div>

            {/* Results pane */}
            <div className="lg:col-span-6 p-6 rounded-2xl border border-stone-250/50 dark:border-stone-900 bg-[#FAF7F2]/90 dark:bg-[#121212]/90 backdrop-blur-sm relative flex flex-col justify-between shadow-sm">
              <div className="absolute top-0 right-0 w-2.5 h-[1px] bg-stone-300 dark:bg-stone-700 pointer-events-none" />
              <div className="absolute top-0 right-0 h-2.5 w-[1px] bg-stone-300 dark:bg-stone-700 pointer-events-none" />

              <AnimatePresence mode="wait">
                {auditResult && (
                  <motion.div
                    initial={{ opacity: 0, y: 10 }}
                    animate={{ opacity: 1, y: 0 }}
                    exit={{ opacity: 0, y: -10 }}
                    transition={{ duration: 0.15 }}
                    className="space-y-5 h-full flex flex-col justify-between"
                  >
                    <div className="space-y-4">
                      {/* Grade Banner */}
                      <div className="flex items-center justify-between border-b border-stone-200/30 pb-3">
                        <span className="text-[10px] font-bold text-stone-500 font-mono tracking-widest uppercase select-none">
                          ANALYSIS_AUDIT_REPORT
                        </span>
                        <div className="flex items-center gap-2">
                          <span className="text-xs font-bold text-stone-700">品質推定スコア:</span>
                          <span className={`px-3 py-1 rounded-lg text-sm font-black font-mono tracking-wide border border-stone-250/50 dark:border-stone-850/60 ${auditResult.scoreBg} ${auditResult.scoreColor}`}>
                            {auditResult.score}
                          </span>
                        </div>
                      </div>

                      {/* Items scroll */}
                      <div className="space-y-3 max-h-[310px] overflow-y-auto pr-1">
                        {auditResult.findings.map((item, idx) => {
                          const isPass = item.status === 'pass';
                          const isWarn = item.status === 'warn';
                          return (
                            <div 
                              key={idx} 
                              className={`p-3 rounded-xl border text-xs leading-relaxed flex items-start gap-3 transition-colors ${
                                isPass 
                                  ? 'bg-[#EFECE4]/20 border-stone-200/40 dark:bg-stone-950/20 dark:border-stone-900'
                                  : isWarn
                                  ? 'bg-amber-50/10 dark:bg-amber-950/5 border-amber-200/20 dark:border-amber-950/20'
                                  : 'bg-rose-50/10 dark:bg-rose-950/5 border-rose-200/20 dark:border-rose-950/20'
                              }`}
                            >
                              {isPass ? (
                                <CheckCircle className="w-4.5 h-4.5 text-emerald-500 shrink-0 mt-0.5" />
                              ) : isWarn ? (
                                <AlertTriangle className="w-4.5 h-4.5 text-amber-500 shrink-0 mt-0.5" />
                              ) : (
                                <AlertTriangle className="w-4.5 h-4.5 text-rose-500 shrink-0 mt-0.5" />
                              )}
                              <div className="space-y-1">
                                <h4 className="font-bold text-stone-950 dark:text-stone-300 leading-tight">
                                  {item.title}
                                </h4>
                                <p className="text-[11px] text-stone-500 dark:text-stone-400">
                                  {item.desc}
                                </p>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </div>

                    <div className="pt-4 border-t border-stone-200/50 dark:border-stone-900 flex items-center gap-2 select-none text-[10px] md:text-xs text-stone-500 font-medium">
                      <Award className="w-4 h-4 text-stone-400" />
                      <span>診断により可読性・フリーズ安全性・実行速度の指標を総合スキャンできます。</span>
                    </div>
                  </motion.div>
                )}
              </AnimatePresence>
            </div>
          </motion.div>
        )}

        {/* ================= TAB 3: VBA ERROR RESOLVER & LOGBOOK ================= */}
        {activeLabTab === 'solver' && (
          <motion.div
            key="solver-tab"
            initial={{ opacity: 0, y: 15 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: -15 }}
            transition={{ duration: 0.2 }}
            className="grid grid-cols-1 lg:grid-cols-12 gap-8"
          >
            {/* Input and Quick select panel */}
            <div className="lg:col-span-5 p-6 rounded-2xl border border-stone-250/50 dark:border-stone-900 bg-[#FAF7F2]/90 dark:bg-[#121212]/90 backdrop-blur-sm space-y-5 relative flex flex-col justify-between shadow-sm">
              <div className="absolute top-0 left-0 w-2.5 h-[1px] bg-stone-300 dark:bg-stone-700 pointer-events-none" />
              <div className="absolute top-0 left-0 h-2.5 w-[1px] bg-stone-300 dark:bg-stone-700 pointer-events-none" />

              <div className="space-y-4">
                <div className="flex items-center gap-2 select-none">
                  <span className="p-1.5 bg-[#FAF7F2]/40 rounded-lg text-indigo-500">
                    <Terminal className="w-4.5 h-4.5" />
                  </span>
                  <h3 className="text-sm md:text-base font-bold font-display tracking-tight text-stone-950 dark:text-white uppercase">
                    エラー自動解読＆対策ロガー
                  </h3>
                </div>
                <p className="text-stone-600 dark:text-stone-400 text-xs font-sans font-medium leading-relaxed">
                  VBAで遭遇したクラッシュメッセージやデバッグウインドウに浮上した「実行時エラー XX」を貼り付けると、頻発する原因（シート不一致など）と解決策を瞬時にリストアップします。
                </p>

                {/* Error textarea */}
                <div className="space-y-2">
                  <label className="text-[10px] uppercase font-mono tracking-widest text-stone-500 font-bold block select-none">
                    PASTE_VBA_ERROR_LOG (エラー文貼り付け)
                  </label>
                  <textarea
                    value={errorInputText}
                    onChange={(e) => {
                      setErrorInputText(e.target.value);
                    }}
                    placeholder="例: 実行時エラー '9': インデックスが有効範囲にありません"
                    className="w-full h-[90px] px-3 py-2 text-xs font-mono rounded-xl border border-stone-200 dark:border-stone-850 bg-stone-50/50 dark:bg-stone-950 text-stone-950 dark:text-stone-100 placeholder-stone-550 focus:outline-none focus:ring-1 focus:ring-stone-400 leading-normal"
                  />
                  <button
                    onClick={performErrorAnalysis}
                    className="w-full py-2.5 bg-stone-950 hover:bg-stone-900 dark:bg-stone-100 dark:hover:bg-stone-200 text-[#FAF9F6] dark:text-[#101010] font-bold text-xs font-display rounded-xl tracking-widest flex items-center justify-center gap-1.5 cursor-pointer transition-all active:scale-[0.99]"
                  >
                    <span>🔍 エラーを解読＆特定</span>
                  </button>
                </div>

                {/* Sample error triggers */}
                <div className="space-y-1.5 pt-1.5">
                  <span className="text-[10px] font-bold font-mono text-stone-500 block uppercase tracking-wider select-none">デモ用エラーログ瞬間挿入</span>
                  <div className="flex flex-wrap gap-1.5">
                    <button
                      onClick={() => {
                        setErrorInputText("実行時エラー '9': インデックスが有効範囲にありません");
                        setTimeout(() => performErrorAnalysis(), 50);
                      }}
                      className="px-2 py-1 text-[10px] font-bold text-stone-700 bg-stone-150/50 hover:bg-stone-200/50 dark:text-stone-300 dark:bg-stone-900 dark:hover:bg-stone-850 rounded-lg transition-colors cursor-pointer border border-transparent dark:border-stone-800"
                    >
                      エラー9 (シート不一致)
                    </button>
                    <button
                      onClick={() => {
                        setErrorInputText("実行時エラー '91': オブジェクト変数または With ブロック変数が設定されていません");
                        setTimeout(() => performErrorAnalysis(), 50);
                      }}
                      className="px-2 py-1 text-[10px] font-bold text-stone-700 bg-stone-150/50 hover:bg-stone-200/50 dark:text-stone-300 dark:bg-stone-900 dark:hover:bg-stone-850 rounded-lg transition-colors cursor-pointer border border-transparent dark:border-stone-800"
                    >
                      エラー91 (Set忘れ)
                    </button>
                    <button
                      onClick={() => {
                        setErrorInputText("実行時エラー '13': 型が一致しません");
                        setTimeout(() => performErrorAnalysis(), 50);
                      }}
                      className="px-2 py-1 text-[10px] font-bold text-stone-700 bg-stone-150/50 hover:bg-stone-200/50 dark:text-stone-300 dark:bg-stone-900 dark:hover:bg-stone-850 rounded-lg transition-colors cursor-pointer border border-transparent dark:border-stone-800"
                    >
                      エラー13 (型不一致)
                    </button>
                    <button
                      onClick={() => {
                        setErrorInputText("コンパイル エラー: 変数が定義されていません");
                        setTimeout(() => performErrorAnalysis(), 50);
                      }}
                      className="px-2 py-1 text-[10px] font-bold text-stone-700 bg-stone-150/50 hover:bg-stone-200/50 dark:text-stone-300 dark:bg-stone-900 dark:hover:bg-stone-850 rounded-lg transition-colors cursor-pointer border border-transparent dark:border-stone-800"
                    >
                      コンパイルエラー
                    </button>
                  </div>
                </div>
              </div>

              {/* Action output check */}
              {solvedAnalysis && (
                <div className="pt-4 border-t border-stone-200/50 dark:border-stone-900">
                  <button
                    onClick={handleSaveToLogbook}
                    className="w-full flex items-center justify-center gap-1.5 py-2 border border-dashed border-stone-300 dark:border-stone-800 text-[11px] font-bold text-indigo-600 dark:text-indigo-400 hover:text-indigo-700 hover:border-indigo-400 dark:hover:text-indigo-300 rounded-xl transition-all cursor-pointer bg-indigo-50/5 dark:bg-indigo-950/5"
                  >
                    <Plus className="w-3.5 h-3.5" />
                    <span>このデバッグ情報を記録ログへ保存</span>
                  </button>
                </div>
              )}
            </div>

            {/* Diagnostics details & Sample Fix */}
            <div className="lg:col-span-7 space-y-6 flex flex-col justify-between">
              
              <AnimatePresence mode="wait">
                {solvedAnalysis ? (
                  <motion.div
                    key={solvedAnalysis.code}
                    initial={{ opacity: 0, y: 10 }}
                    animate={{ opacity: 1, y: 0 }}
                    exit={{ opacity: 0, scale: 0.98 }}
                    transition={{ duration: 0.15 }}
                    className="p-6 rounded-2xl border border-stone-250/50 dark:border-stone-900 bg-[#FAF7F2]/95 dark:bg-[#121212]/90 backdrop-blur-sm space-y-4 shadow-sm"
                  >
                    {/* Header info */}
                    <div className="flex flex-col md:flex-row md:items-center justify-between gap-3 border-b border-stone-200/40 dark:border-stone-900 pb-3">
                      <div>
                        <span className="text-[9px] font-bold font-mono tracking-widest text-[#8C8A85] block select-none">
                          ERROR_CODE_DETECTED_OK
                        </span>
                        <h4 className="text-base font-bold text-stone-950 dark:text-stone-100 font-display mt-0.5">
                          {solvedAnalysis.code !== "GEN_SYS" && `実行時エラー '${solvedAnalysis.code}': `}
                          {solvedAnalysis.name}
                        </h4>
                      </div>
                      <span className="px-2.5 py-0.5 text-[9px] font-bold font-mono bg-stone-100 text-stone-700 dark:bg-stone-900 dark:text-stone-400 border border-stone-200 dark:border-stone-850 rounded uppercase tracking-widest">
                        {solvedAnalysis.english}
                      </span>
                    </div>

                    {/* Description Paragraph */}
                    <div className="space-y-4">
                      <div className="space-y-1">
                        <span className="text-[10px] font-bold text-stone-500 uppercase tracking-wide flex items-center gap-1 font-mono">// コア引き金・原因：</span>
                        <p className="text-xs text-stone-600 dark:text-stone-400 leading-relaxed font-sans font-medium">
                          {solvedAnalysis.reason}
                        </p>
                      </div>

                      <div className="space-y-1">
                        <span className="text-[10px] font-bold text-stone-500 uppercase tracking-wide flex items-center gap-1 font-mono">// 解決・対抗策：</span>
                        <div className="text-xs text-stone-600 dark:text-stone-400 font-medium font-sans leading-relaxed whitespace-pre-line">
                          {solvedAnalysis.solution}
                        </div>
                      </div>

                      <div className="space-y-1.5 pt-2">
                        <span className="text-[10px] font-bold text-stone-500 uppercase tracking-wide block font-mono">// NG / OK コード対比例：</span>
                        <div className="rounded-xl border border-stone-200/60 dark:border-stone-900 overflow-hidden text-xs">
                          <SyntaxHighlighter code={solvedAnalysis.sampleCode} language="vba" id={`error-fix-${solvedAnalysis.code}`} />
                        </div>
                      </div>
                    </div>
                  </motion.div>
                ) : (
                  <div className="p-8 rounded-2xl border border-stone-250/20 dark:border-stone-900 bg-stone-50/10 dark:bg-stone-950/20 text-center text-stone-500 text-xs md:text-sm font-sans select-none flex flex-col items-center justify-center h-[260px] space-y-2">
                    <Terminal className="w-8 h-8 text-stone-400 animate-tech-pulse" />
                    <span>エラーログを入力、または左側のデモボタンをクリックして解読を走らせてください。</span>
                  </div>
                )}
              </AnimatePresence>

              {/* Debug history logbook list - VERY premium and functional */}
              <div className="p-6 rounded-2xl border border-stone-250/50 dark:border-stone-900 bg-[#FAF7F2]/90 dark:bg-[#121212]/90 backdrop-blur-sm space-y-4">
                <div className="flex items-center justify-between select-none">
                  <div className="flex items-center gap-2">
                    <History className="w-4 h-4 text-[#8C8A85]" />
                    <h4 className="text-xs font-black font-mono tracking-widest text-[#8C8A85] uppercase">
                      お仕事デバッグ歴史記録帳 (Logbook)
                    </h4>
                  </div>
                  <span className="text-[9px] font-mono text-stone-400 tracking-wide font-black">
                    SAVED: {savedErrorLogs.length} CASES
                  </span>
                </div>

                <div className="space-y-2 max-h-[160px] overflow-y-auto pr-1">
                  {savedErrorLogs.length > 0 ? (
                    savedErrorLogs.map((log) => (
                      <div 
                        key={log.id} 
                        className={`p-3 rounded-xl border flex items-center justify-between gap-4 text-xs transition-colors ${
                          log.isFixed 
                            ? 'bg-stone-50/40 border-stone-200/30 text-stone-400 dark:bg-stone-950/10 dark:border-stone-900' 
                            : 'bg-[#EFECE4]/30 border-stone-300 dark:bg-stone-900/30 dark:border-stone-800'
                        }`}
                      >
                        <div className="flex items-start gap-2.5 min-w-0">
                          {/* Checked box */}
                          <input
                            type="checkbox"
                            checked={log.isFixed}
                            onChange={() => handleToggleLogFixed(log.id)}
                            className="w-4 h-4 shrink-0 mt-0.5 accent-stone-900 dark:accent-white cursor-pointer"
                            title="解決・完了マークトグ"
                          />
                          <div className="min-w-0 font-sans">
                            <div className="flex items-center gap-2 flex-wrap text-[11px] leading-tight font-medium">
                              <span className="font-bold text-stone-800 dark:text-stone-300 font-mono">
                                [Err {log.errorCode}]
                              </span>
                              <span className={`font-semibold overflow-hidden text-ellipsis whitespace-nowrap ${log.isFixed ? 'line-through text-stone-400' : 'text-stone-900 dark:text-stone-200'}`}>
                                {log.errorTitle}
                              </span>
                            </div>
                            <p className="text-[10px] text-stone-500 mt-0.5 whitespace-nowrap overflow-hidden text-ellipsis">
                              {log.notes} <span className="text-stone-400 text-[9px] font-mono">({log.timestamp})</span>
                            </p>
                          </div>
                        </div>

                        <button 
                          onClick={() => handleDeleteLog(log.id)}
                          className="p-1.5 hover:bg-stone-200/50 dark:hover:bg-stone-850 rounded-lg text-stone-400 hover:text-stone-700 dark:hover:text-stone-200 transition-colors cursor-pointer"
                          title="履歴を削除"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    ))
                  ) : (
                    <div className="text-center py-4 text-[11px] text-stone-400 font-medium select-none">
                      履歴はありません。上のボタンから現在のデバッグを記録できます。
                    </div>
                  )}
                </div>
              </div>

            </div>
          </motion.div>
        )}

        {/* ================= TAB 4: EFFICIENCY SAVINGS SIMULATOR ================= */}
        {activeLabTab === 'roi' && (
          <motion.div
            key="roi-tab"
            initial={{ opacity: 0, y: 15 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: -15 }}
            transition={{ duration: 0.2 }}
            className="p-6 md:p-8 rounded-2xl border border-stone-250/50 dark:border-stone-900 bg-[#FAF7F2]/90 dark:bg-[#121212]/90 backdrop-blur-md relative overflow-hidden group shadow-sm"
          >
            <div className="absolute top-0 left-0 w-3.5 h-[1px] bg-stone-300 dark:bg-stone-750 pointer-events-none" />
            <div className="absolute top-0 left-0 h-3.5 w-[1px] bg-stone-300 dark:bg-stone-750 pointer-events-none" />

            <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-8">
              
              <div className="space-y-4 max-w-xl flex-1">
                <div className="flex items-center gap-2 text-stone-950 dark:text-white select-none">
                  <TrendingUp className="w-4.5 h-4.5 text-stone-500" />
                  <h3 className="text-sm md:text-base font-bold font-display uppercase tracking-wider">
                    自動化投資対効果（ROI）計算機
                  </h3>
                </div>
                <p className="text-stone-600 dark:text-stone-400 text-xs md:text-sm leading-relaxed font-sans font-medium">
                  VBAやRPAツールへの自社移行が、年間ベースでどれほどの人的コストと不要な残業時間を削減するのかをシミュレーションします。
                </p>

                <div className="space-y-5 pt-2 font-sans select-none">
                  {/* Slider 1: Minutes per task */}
                  <div className="space-y-1.5">
                    <div className="flex justify-between items-center text-xs font-semibold">
                      <span className="text-stone-600 dark:text-stone-300">手作業時の1プロセス所要時間:</span>
                      <span className="font-bold font-mono text-stone-900 dark:text-[#FAF7F2] bg-stone-200/60 dark:bg-stone-850 px-2 py-0.5 rounded-lg">{minutesPerTask} 分</span>
                    </div>
                    <input
                      type="range"
                      min="5"
                      max="180"
                      step="5"
                      value={minutesPerTask}
                      onChange={(e) => setMinutesPerTask(Number(e.target.value))}
                      className="w-full accent-stone-900 dark:accent-stone-300 cursor-pointer"
                    />
                  </div>

                  {/* Slider 2: Times executed per week */}
                  <div className="space-y-1.5">
                    <div className="flex justify-between items-center text-xs font-semibold">
                      <span className="text-stone-600 dark:text-stone-300">マクロの想定週次実行頻度:</span>
                      <span className="font-bold font-mono text-stone-900 dark:text-[#FAF7F2] bg-stone-200/60 dark:bg-stone-850 px-2 py-0.5 rounded-lg">{frequencyPerWeek} 回 / 週</span>
                    </div>
                    <input
                      type="range"
                      min="1"
                      max="40"
                      step="1"
                      value={frequencyPerWeek}
                      onChange={(e) => setFrequencyPerWeek(Number(e.target.value))}
                      className="w-full accent-stone-900 dark:accent-stone-300 cursor-pointer"
                    />
                  </div>

                  {/* Slider 3: Target hourly wage */}
                  <div className="space-y-1.5">
                    <div className="flex justify-between items-center text-xs font-semibold">
                      <span className="text-stone-600 dark:text-stone-300">設計エンジニアや作業員の想定時給換算:</span>
                      <span className="font-bold font-mono text-stone-900 dark:text-[#FAF7F2] bg-stone-200/60 dark:bg-stone-850 px-2 py-0.5 rounded-lg">¥{hourlyWage.toLocaleString()} / 時間</span>
                    </div>
                    <input
                      type="range"
                      min="1000"
                      max="8000"
                      step="250"
                      value={hourlyWage}
                      onChange={(e) => setHourlyWage(Number(e.target.value))}
                      className="w-full accent-stone-900 dark:accent-stone-300 cursor-pointer"
                    />
                  </div>
                </div>
              </div>

              {/* Display Output Cards */}
              <div className="lg:w-[400px] bg-[#EFECE4]/30 dark:bg-stone-950/40 border border-stone-200 dark:border-stone-900 p-6 rounded-2xl flex flex-col justify-between space-y-6 relative shrink-0">
                <div className="absolute top-0 right-0 w-2.5 h-[1px] bg-stone-300 dark:bg-stone-750 pointer-events-none" />
                <div className="absolute top-0 right-0 h-2.5 w-[1px] bg-stone-300 dark:bg-stone-750 pointer-events-none" />

                <span className="text-[9px] font-mono tracking-widest text-stone-500 block font-bold leading-none uppercase select-none">
                  CALCULATED_INVESTMENT_RETURN
                </span>

                <div className="grid grid-cols-2 gap-4 select-none">
                  <div className="space-y-1">
                    <span className="text-[10px] md:text-xs text-stone-500 block font-bold leading-none">年間削減労働時間</span>
                    <span className="text-2xl md:text-3xl font-black font-display tracking-tight text-emerald-600 dark:text-emerald-400 block whitespace-nowrap">
                      {timeSavedPerYearHours.toLocaleString()} <span className="text-xs font-bold text-stone-500 font-sans">h</span>
                    </span>
                  </div>
                  <div className="space-y-1">
                    <span className="text-[10px] md:text-xs text-stone-500 block font-bold leading-none">年間コスト解放推測</span>
                    <span className="text-2xl md:text-3xl font-black font-display tracking-tight text-stone-950 dark:text-stone-100 block whitespace-nowrap">
                      ¥{moneySavedPerYear.toLocaleString()}
                    </span>
                  </div>
                </div>

                {/* Simulated comparison chart bars */}
                <div className="space-y-3.5 pt-4 border-t border-stone-200 dark:border-stone-900 select-none">
                  <div className="space-y-1 text-xs">
                    <div className="flex justify-between items-center text-[10px] font-mono tracking-wider">
                      <span className="text-stone-550 dark:text-stone-400">手作業に必要な負担時間 (BEFORE)</span>
                      <span className="text-stone-550 dark:text-stone-400 font-bold">100% (基準)</span>
                    </div>
                    <div className="w-full h-2 rounded bg-stone-250 dark:bg-stone-850 overflow-hidden">
                      <div className="h-full bg-stone-400 dark:bg-stone-600 rounded" style={{ width: '100%' }} />
                    </div>
                  </div>

                  <div className="space-y-1 text-xs">
                    <div className="flex justify-between items-center text-[10px] font-mono tracking-wider text-emerald-600 dark:text-emerald-400 font-semibold">
                      <span>最適化マクロ自動化後 (AFTER)</span>
                      <span className="font-bold">約 3.2% 解放</span>
                    </div>
                    <div className="w-full h-2 rounded bg-stone-250 dark:bg-stone-850 overflow-hidden">
                      <div className="h-full bg-emerald-500 rounded animate-pulse" style={{ width: '3.2%' }} />
                    </div>
                  </div>
                </div>

                <p className="text-[10px] text-stone-500 dark:text-stone-600 leading-relaxed font-sans font-medium text-center">
                  ※ 週計算 {yearlyTasks} 回の動作、Excel最適化制御（画面再描画停止・バックグラウンド手動計算化）を内包した数値をベースにした算出です。
                </p>
              </div>

            </div>
          </motion.div>
        )}

      </AnimatePresence>

    </div>
  );
};
