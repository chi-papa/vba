import { VbaSample } from '../types';

export const vbaSamples: VbaSample[] = [
  {
    id: 'shared-data-sync',
    title: '【共有同期＆履歴】入荷状況・動的色判定＆差分更新履歴抽出システム',
    description: '「ControlBox」シートの設定情報をトリガーに、共有（ReadOnly）ブックから商品データとそのセル背景色（入荷状況）を取得。前回取得データとの自動差分比較から「新規追加」履歴をログシートにタイムスタンプ形式で追記し、「本日追加」されたものだけを別シートに瞬時に二次抽出する実務特化型システムです。',
    category: 'excel',
    difficulty: 'hard',
    tags: ['共有ファイル連携', '状態色判定', '履歴自動生成', '本日抽出'],
    code: `'===========================================================
' ■ ControlBoxシート（マスターコントローラー）構成
'   B2：共有ファイルのパス (例: C:\\shared\\items.xlsx)
'   B3：読み取り対象シート名 (例: Sheet1)
'   B4：商品リストの開始行 (例: 2)
'   B5：商品名の列 (例: 1 = A列)
'   B6：背景色が付いている＝入荷済 の色コード (自動取得)
'   B7：更新履歴シート名
'   B8：本日追加シート名
'===========================================================

'===========================================================
' ① 共有ファイルを読み取り、商品名と状態（色）を取得する
'===========================================================
Sub ReadSharedFile()

    Dim sharedPath As String
    Dim targetSheet As String
    Dim startRow As Long
    Dim colName As Long
    Dim inStockColor As Long

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim r As Long
    Dim lastRow As Long
    Dim outRow As Long

    ' ControlBox から設定を取得
    With Sheets("ControlBox")
        sharedPath = .Range("B2").Value
        targetSheet = .Range("B3").Value
        startRow = .Range("B4").Value
        colName = .Range("B5").Value
        inStockColor = .Range("B6").Interior.Color
    End With

    ' 読み取り専用で開く
    Set wb = Workbooks.Open(sharedPath, ReadOnly:=True)
    Set ws = wb.Sheets(targetSheet)

    ' 最終行取得
    lastRow = ws.Cells(ws.Rows.Count, colName).End(xlUp).Row

    ' 一旦、前回の読み取り結果をクリア
    Sheets("ReadData").Cells.ClearContents

    outRow = 2

    ' 商品を読み取り
    For r = startRow To lastRow

        Dim name As String
        Dim color As Long
        Dim status As String

        name = ws.Cells(r, colName).Value
        color = ws.Cells(r, colName).Interior.Color

        If name <> "" Then

            ' 色で状態判定
            If color = inStockColor Then
                status = "入荷済"
            Else
                status = "発注中"
            End If

            ' ReadData に書き込み
            With Sheets("ReadData")
                .Cells(outRow, 1).Value = name
                .Cells(outRow, 2).Value = status
                .Cells(outRow, 3).Value = Now
            End With

            outRow = outRow + 1
        End If

    Next r

    wb.Close False

End Sub


'===========================================================
' ② 前回データと今回データを比較し、更新履歴を記録
'===========================================================
Sub SaveUpdateHistory()

    Dim wsNew As Worksheet
    Dim wsOld As Worksheet
    Dim wsLog As Worksheet

    Set wsNew = Sheets("ReadData")
    Set wsOld = Sheets("OldData")
    Set wsLog = Sheets(Sheets("ControlBox").Range("B7").Value)

    Dim r As Long
    Dim lastNew As Long
    Dim lastOld As Long
    Dim logRow As Long

    lastNew = wsNew.Cells(wsNew.Rows.Count, 1).End(xlUp).Row
    lastOld = wsOld.Cells(wsOld.Rows.Count, 1).End(xlUp).Row
    logRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1

    ' 新しいデータをループ
    For r = 2 To lastNew

        Dim name As String
        Dim status As String

        name = wsNew.Cells(r, 1).Value
        status = wsNew.Cells(r, 2).Value

        ' OldData に存在しない → 新規追加
        If WorksheetFunction.CountIf(wsOld.Range("A:A"), name) = 0 Then

            wsLog.Cells(logRow, 1).Value = Now
            wsLog.Cells(logRow, 2).Value = name
            wsLog.Cells(logRow, 3).Value = status
            wsLog.Cells(logRow, 4).Value = "新規追加"

            logRow = logRow + 1
        End If

    Next r

    ' 今回データを OldData にコピー
    wsOld.Cells.ClearContents
    wsNew.Range("A:C").Copy wsOld.Range("A1")

End Sub

'===========================================================
' ③ 更新履歴から「今日追加された商品」だけ抽出
'===========================================================
Sub ExtractTodayAdded()

    Dim wsLog As Worksheet
    Dim wsToday As Worksheet
    Dim logRow As Long
    Dim lastLog As Long
    Dim outRow As Long

    Set wsLog = Sheets(Sheets("ControlBox").Range("B7").Value)
    Set wsToday = Sheets(Sheets("ControlBox").Range("B8").Value)

    wsToday.Cells.ClearContents
    wsToday.Range("A1").Value = "本日追加された商品"

    lastLog = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row
    outRow = 2

    For logRow = 2 To lastLog

        If DateValue(wsLog.Cells(logRow, 1).Value) = Date Then

            wsToday.Cells(outRow, 1).Value = wsLog.Cells(logRow, 2).Value
            wsToday.Cells(outRow, 2).Value = wsLog.Cells(logRow, 3).Value
            wsToday.Cells(outRow, 3).Value = wsLog.Cells(logRow, 4).Value

            outRow = outRow + 1
        End If

    Next logRow

End Sub`,
    explanation: [
      '【セキュアなReadOnlyファイル同期】: 他部署や共有ディレクトリにあるブックを「ReadOnly:=True」で開くことで、元の集計データを不用意に変更あるいはロックしてしまう二次障害を100%防止する実務的テクニックです。',
      '【動的な入荷済色（RGB）判定】: 固定的なカラー値ではなく、ControlBoxのB6セルの「Interior.Color」をそのまま判定基準値として代入（`.Interior.Color`）しているため、Excelシステム管理者側でB6の塗りつぶし色を変えるだけで、自動判定基準色が瞬時に最適化されます。',
      '【CountIfを活用した超低遅延差分比較】: リストループの中で「WorksheetFunction.CountIf」を用いて前回退避データ（`OldData`）の全量中にその商品があるかないかをシンプルさで高速スキャンし、「新規追加」された商品を劇的に高速検出します。',
      '【DateValue安全フィルタリング】: ログシートの「日時（Now）」情報から、時間帯を削ぎ落として日付だけを取り出す「DateValue(cell) = Date」の厳密一致判定を利用し、本日追加された情報だけを極めて正確に別シートへ高速転記完了させます。'
    ]
  },
  {
    id: 'order-transfer-v3',
    title: '【受注自動転記】受注転記マクロ v3（再バグ修正・最適化版）',
    description: '入力シートにCtrl+Vで貼り付けたテキストを自動で項目分解し、注文書シートの任意セルへ転記を完結。バグを引き起こしやすいInteger型やオブジェクト初期化不備、シート名エスケープ問題を完全に解消したプロ仕様の実務向け完全版コードです。',
    category: 'excel',
    difficulty: 'hard',
    tags: ['業務自動化', '文字列分解', 'イベント制御', 'バグ対応版'],
    code: `' ================================================
' ■ 受注転記マクロ  VBAコード v3（再バグ修正版）
' ================================================

Option Explicit

' シート名定数（変更時はここだけ直す）
Private Const SH_INPUT  As String = "入力シート"
Private Const SH_ORDER  As String = "注文書"
Private Const SH_CONFIG As String = "設定"

' -----------------------------------------------
' リセット処理 ※「リセット」ボタンに登録
' ① =参照を注文書にセット → ② A1・B列クリア
' -----------------------------------------------
Sub ResetOrder()
    Dim wsIn As Worksheet
    Set wsIn = ThisWorkbook.Sheets(SH_INPUT)

    ' [修正4] SetFormulas内エラーでEnableEventsがFalseのまま
    '         止まらないようOn Errorで保護
    On Error Resume Next
    Call SetFormulas
    On Error GoTo 0

    ' A1・B列クリア（Worksheet_Change発火を抑止）
    Application.EnableEvents = False
    wsIn.Range("A1").ClearContents
    wsIn.Range("B3:B7").ClearContents
    Application.EnableEvents = True

    wsIn.Activate
    wsIn.Range("A1").Select

    MsgBox "リセット完了。A1に受注情報を貼り付けてください。", _
           vbInformation, "リセット完了"
End Sub

' -----------------------------------------------
' 確定処理 ※「確定」ボタンに登録
' 注文書の =参照 を値に変換
' -----------------------------------------------
Sub ConfirmOrder()
    Dim wsOrd As Worksheet, wsCfg As Worksheet
    Dim i As Long, lastRow As Long          ' [修正1] Long型
    Dim dstCell As String
    Dim r As Range

    Set wsOrd = ThisWorkbook.Sheets(SH_ORDER)
    Set wsCfg = ThisWorkbook.Sheets(SH_CONFIG)

    If MsgBox("注文書を確定します。" & Chr(10) & _
              "=参照を値に変換してよいですか？", _
              vbYesNo + vbQuestion, "確定確認") <> vbYes Then Exit Sub

    lastRow = GetConfigLastRow(wsCfg)
    If lastRow < 2 Then                     ' [修正6] 設定が空のとき
        MsgBox "設定シートに項目が登録されていません。", vbExclamation, "設定エラー"
        Exit Sub
    End If

    Dim hasFormula As Boolean
    hasFormula = False

    For i = 2 To lastRow
        dstCell = Trim(wsCfg.Cells(i, 3).Value)
        If dstCell = "" Then GoTo NextRow

        Set r = Nothing                     ' [修正2] 必ずNothingで初期化
        On Error Resume Next
        Set r = wsOrd.Range(dstCell)
        On Error GoTo 0

        If Not r Is Nothing Then
            If r.HasFormula Then
                r.Value = r.Value
                hasFormula = True
            End If
        End If
NextRow:
    Next i

    If Not hasFormula Then
        MsgBox "転記先に =参照 が見つかりませんでした。" & Chr(10) & _
               "先にリセットを実行してください。", vbExclamation, "確定エラー"
        Exit Sub
    End If

    MsgBox "確定しました！印刷後、リセットボタンで次の受注へ。", _
           vbInformation, "完了"
End Sub

' -----------------------------------------------
' =参照を注文書の転記先セルにセット
' -----------------------------------------------
Sub SetFormulas()
    Dim wsOrd As Worksheet, wsCfg As Worksheet
    Dim i As Long, lastRow As Long          ' [修正1] Long型
    Dim srcRow As Long, dstCell As String   ' [修正1] Long型
    Dim r As Range

    Set wsOrd = ThisWorkbook.Sheets(SH_ORDER)
    Set wsCfg = ThisWorkbook.Sheets(SH_CONFIG)

    lastRow = GetConfigLastRow(wsCfg)
    If lastRow < 2 Then Exit Sub            ' [修正6] 設定が空のとき

    For i = 2 To lastRow
        srcRow  = 0
        dstCell = ""
        On Error Resume Next
        srcRow  = CLng(wsCfg.Cells(i, 2).Value)   ' [修正1] CLng
        dstCell = Trim(wsCfg.Cells(i, 3).Value)
        On Error GoTo 0

        If dstCell = "" Or srcRow <= 0 Then GoTo NextRow2

        Set r = Nothing                     ' [修正2] 必ずNothingで初期化
        On Error Resume Next
        Set r = wsOrd.Range(dstCell)
        On Error GoTo 0

        If Not r Is Nothing Then
            ' [修正3] シート名をシングルクォートで囲む
            r.Formula = "='" & SH_INPUT & "'!B" & srcRow
        End If
NextRow2:
    Next i
End Sub

' -----------------------------------------------
' 貼り付けテキストを分解 → 入力シートB列に書き込み
' Worksheet_Changeからのみ呼ばれる
' -----------------------------------------------
Sub ParsePastedText(ByVal pastedText As String)
    Dim wsIn  As Worksheet, wsCfg As Worksheet
    Dim i As Long, j As Long, lastRow As Long  ' [修正1] Long型
    Dim lines() As String, oneLine As String
    Dim colonPos As Long                        ' [修正1] Long型
    Dim lbl As String, val As String
    Dim cfgLabel As String, cfgSrcRow As Long   ' [修正1] Long型

    Set wsIn  = ThisWorkbook.Sheets(SH_INPUT)
    Set wsCfg = ThisWorkbook.Sheets(SH_CONFIG)

    wsIn.Range("B3:B7").ClearContents

    If Trim(pastedText) = "" Then Exit Sub

    lastRow = GetConfigLastRow(wsCfg)
    If lastRow < 2 Then Exit Sub            ' [修正6] 設定が空のとき

    ' CR+LF / LF / CR すべて対応
    pastedText = Replace(pastedText, Chr(13) & Chr(10), Chr(10))
    pastedText = Replace(pastedText, Chr(13), Chr(10))
    lines = Split(pastedText, Chr(10))

    For i = 0 To UBound(lines)
        oneLine = Trim(lines(i))
        If oneLine = "" Then GoTo NextLine

        ' コロン位置（半角・全角両対応）
        colonPos = InStr(oneLine, ":")
        If colonPos = 0 Then colonPos = InStr(oneLine, Chr(65306))
        If colonPos = 0 Then GoTo NextLine

        lbl = Trim(Left(oneLine, colonPos - 1))
        ' [ポイント] 値はcolonPos以降すべて取得（値中のコロン対応）
        val = Trim(Mid(oneLine, colonPos + 1))

        For j = 2 To lastRow
            cfgLabel  = Trim(wsCfg.Cells(j, 1).Value)
            cfgSrcRow = 0
            On Error Resume Next
            cfgSrcRow = CLng(wsCfg.Cells(j, 2).Value)  ' [修正1] CLng
            On Error GoTo 0
            If cfgLabel <> "" And cfgSrcRow > 0 Then
                If InStr(lbl, cfgLabel) > 0 Or InStr(cfgLabel, lbl) > 0 Then
                    wsIn.Cells(cfgSrcRow, 2).Value = val
                    Exit For
                End If
            End If
        Next j
NextLine:
    Next i

    Call SetFormulas
End Sub

' -----------------------------------------------
' 設定シートの最終行を取得
' -----------------------------------------------
Private Function GetConfigLastRow(wsCfg As Worksheet) As Long  ' [修正1] Long型
    GetConfigLastRow = wsCfg.Cells(wsCfg.Rows.Count, 1).End(xlUp).Row
End Function



' ================================================
' 【 Sheet1「入力シート」のコードウィンドウ 】
' ================================================
' (※ シート1に記述するワークシートChangeイベント)
' Private Sub Worksheet_Change(ByVal Target As Range)
'     If Target.Address <> "$A$1" Then Exit Sub
'     Application.EnableEvents = False
'     On Error GoTo Cleanup
'     Call Module1.ParsePastedText(Target.Value)
'     GoTo Done
' Cleanup:
'     ' エラー防止処理
' Done:
'     Application.EnableEvents = True
' End Sub`,
    explanation: [
      '【Integer → Long型への一貫統一】: Excelの最大行（104万余り）制限や、メモリ演算でより安全な32/64ビット符号付き変数の最適化を追求するため、古いInteger(上限32,767)を排除し、すべて「Long型」に一律修正しています。',
      '【Nothingでの安全な初期化】: VBAで「On Error Resume Next」を併用した範囲取得ループを行う際、取得失敗時に直前のセルrをキャリーオーバーさせないための「Set r = Nothing」処理を、ループ直前に必ず注入します。',
      '【シート名スペースへの完全防護】: シート名定数「SH_INPUT」がスペースや特殊記号を将来含むようになってもExcel数式構造が狂わないよう、文字列エスケープ記号としてシングルクォーテーション「`\'`」を追加加工しています。',
      '【エラー時発動トラップによるイベント崩壊回避】: Worksheet_Change監視内でWorksheetをいじる際、「EnableEvents = False」のままマクロが落丁することで一切のExcelセル編集監視がストップする惨事を、On Error GoTo構造で安全に保護します。'
    ]
  },
  {
    id: 'clear-cells-controlbox',
    title: '【セル一括クリア】ControlBoxシートのリスト連携・安全クリアマクロ',
    description: '「ControlBox」シートのセル番地リスト（A5:B14など）に記入された10個のセル候補を動的にロードし、安全に内容を一括クリア（ClearContents）します。未記入セルの存在や、無効な番地が記述されていてもエラーを出さずに進む堅牢な設計です。',
    category: 'utility',
    difficulty: 'easy',
    tags: ['セルクリア', 'シート制御', '動的管理', '安全設計'],
    code: `'===========================================
' 10個のセル候補を ControlBox から読み取り
' 未記入セルはスキップしてエラーを回避
'===========================================
Sub ClearCellsFromControlBox()

    Dim i As Long
    Dim addr As String
    Dim wsCtrl As Worksheet
    Dim wsTarget As Worksheet

    ' ControlBox シート
    Set wsCtrl = Sheets("ControlBox")

    ' クリア対象のシート（必要に応じて変更）
    Set wsTarget = Sheets("Sheet1")

    ' A5:B14 に 1〜10 のセル番地が入っている想定
    For i = 5 To 14

        addr = wsCtrl.Range("B" & i).Value

        ' 空欄はスキップ
        If addr <> "" Then

            ' セル番地が不正でもエラーを出さない
            On Error Resume Next
            wsTarget.Range(addr).ClearContents
            On Error GoTo 0

        End If

    Next i

End Sub`,
    explanation: [
      '【ControlBoxによる外部動的管理】: セルの消去範囲やアドレス情報をマクロプログラム中に直接埋め込まないため、セルのレイアウト変更があった際でもユーザー自身が「ControlBox」シートを書き換えるだけで動作範囲を変更可能にします。',
      '【未記入セルの安全スルー】: 「B5:B14」の全セルに入力値が入っている必要はなく、「If addr <> "" Then」により空欄行は綺麗にスキップされ、エラーやデバッグ停止を発生させません。',
      '【万が一の無効入力対策】: ユーザーが「ControlBox」シートに誤って存在しない範囲（例：「B9999999」や無効な英字など）を入力した場合でも、VBAが中断してしまわない「On Error Resume Next」＆「On Error GoTo 0」の局所エラー保護を記述して保護しています。'
    ]
  },
  {
    id: 'clear-cells-multisheet',
    title: '【複数シート横断クリア】ControlBox指定・複数シート横断安全クリアマクロ',
    description: '「ControlBox」シートのA5:C14で設定された「シート名（B列）」と「セル番地（C列）」のペアを1つずつ読み取り、複数シートに跨るセル内容を一括クリア（ClearContents）します。シートが存在しない場合や無効なセル番地が指定されても、スキップして処理を継続する堅牢設計です。',
    category: 'utility',
    difficulty: 'medium',
    tags: ['複数シート', '一括クリア', '動的シート指定', 'エラー回避'],
    code: `'===========================================
' 複数シートに跨る10個のセル候補を
' ControlBox の「シート名」「セル番地」から読み取ってクリア
' 未記入分、または存在しないシートは安全にスキップ
'===========================================
Sub ClearMultiSheetsFromControlBox()

    Dim i As Long
    Dim shName As String
    Dim addr As String
    Dim wsCtrl As Worksheet
    Dim wsTarget As Worksheet

    ' ControlBox シートの設定
    On Error Resume Next
    Set wsCtrl = ThisWorkbook.Sheets("ControlBox")
    On Error GoTo 0
    
    If wsCtrl Is Nothing Then
        MsgBox "「ControlBox」シートが見つかりません。", vbCritical, "エラー"
        Exit Sub
    End If

    ' A5:C14 に 1〜10 の設定を記入している想定
    ' B列：シート名、C列：セル番地
    For i = 5 To 14

        shName = Trim(wsCtrl.Range("B" & i).Value)
        addr = Trim(wsCtrl.Range("C" & i).Value)

        ' シート名とセル番地の両方が空欄でない場合のみ処理
        If shName <> "" And addr <> "" Then

            ' 指定された名前のシートが存在するか安全に確認
            Set wsTarget = Nothing
            On Error Resume Next
            Set wsTarget = ThisWorkbook.Sheets(shName)
            On Error GoTo 0

            If Not wsTarget Is Nothing Then
                ' セル番地が不正（例：無効な定義や空文字）でもエラーを出さない
                On Error Resume Next
                wsTarget.Range(addr).ClearContents
                On Error GoTo 0
            Else
                ' シートが存在しない場合はイミディエイトログに出力（デバッグ確認用）
                Debug.Print "警告: シート 「" & shName & "」 が無いためスキップされました。"
            End If

        End If

    Next i
    
    MsgBox "複数シートにまたがる対象セルのクリアを完了しました！", vbInformation, "処理完了"

End Sub`,
    explanation: [
      '【複数シート間を動的横断】: 単一シートに縛られず、B列に入力されたシート情報を動的に検索し、指定シートの指定セルを一括クリアしていく柔軟な実務用設計です。',
      '【存在しないシートのスキップ処理】: 「Set wsTarget = ThisWorkbook.Sheets(shName)」をエラー無視ブロックで挟み、「wsTarget Is Nothing」の検知ロジックによって、消滅したシートへのアクセスエラーによる強制終了（デバッグ停止）を完全に防護しています。',
      '【イミディエイトログログ出力】: エラーをサイレントで潰すだけでなく、「Debug.Print」に警告内容を出力しておくことで、マクロを実行した後にVBE画面で「どれがスキップされたか」を確認できます。'
    ]
  },
  {
    id: 'clear-backup-initialize',
    title: '【高機能アレンジ版】履歴バックアップ＆数式保護付き・セル一括初期化',
    description: 'ControlBoxシートのリスト連携をさらにパワーアップ！クリア前に消去データを「履歴」シートへ日付付きで自動バックアップ。更に、初期設定値（D列）を指定可能にし、数式（=）の入った集計用セルの誤消去を防ぐ「数式保護（E列）」機能まで備えた実務向け最高水準マクロです。',
    category: 'utility',
    difficulty: 'hard',
    tags: ['バックアップ', '初期化', '安全設計', '数式保護'],
    code: `'===========================================================
' 【アレンジ版】履歴自動バックアップ ＆ 数式保護付き
'  シート横断セル一括初期化・クリアマクロ
'===========================================================
Sub InitializeAndBackupWithControlBox()

    Dim i As Long
    Dim shName As String
    Dim addr As String
    Dim defVal As Variant
    Dim isProtectFormula As Boolean
    
    Dim wsCtrl As Worksheet
    Dim wsTarget As Worksheet
    Dim wsLog As Worksheet
    Dim rTarget As Range
    Dim logRow As Long

    ' 1. 各シートの関連付け
    On Error Resume Next
    Set wsCtrl = ThisWorkbook.Sheets("ControlBox")
    Set wsLog = ThisWorkbook.Sheets("履歴") ' 消去前のデータを退避するシート
    On Error GoTo 0
    
    If wsCtrl Is Nothing Then
        MsgBox "「ControlBox」シートが見つかりません。", vbCritical, "エラー"
        Exit Sub
    End If
    
    ' バックアップ用「履歴」シートがない場合に自動作成する親切機能
    If wsLog Is Nothing Then
        If MsgBox("履歴退避用の「履歴」シートが作成されていません。" & Chr(10) & _
                  "自動作成してもよろしいですか？", vbYesNo + vbQuestion, "シート自動作成") = vbYes Then
            Set wsLog = ThisWorkbook.Sheets.Add(After:=wsCtrl)
            wsLog.Name = "履歴"
            wsLog.Range("A1:E1").Value = Array("実行日時", "対象シート", "セル番地", "消去前のデータ", "新しく入れた値")
            wsLog.Columns("A:E").AutoFit
        Else
            MsgBox "履歴シートが必要な設計のため、処理を中断します。", vbExclamation, "中断"
            Exit Sub
        End If
    End If

    ' 高速化＆警告表示OFF（大量クリア時のチラつきや、上書き確認の警告をカット）
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' バックアップ用の書き込み先最終行の検出
    logRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1

    ' A5:E14 に 1〜10 の設定があると想定
    ' B列:シート名, C列:セル番地, D列:デフォルト初期値(未設定ならクリア), E列:数式保護(TRUE=数式セルを保護)
    For i = 5 To 14

        shName = Trim(wsCtrl.Range("B" & i).Value)
        addr = Trim(wsCtrl.Range("C" & i).Value)
        defVal = wsCtrl.Range("D" & i).Value
        
        ' E列：数式保護の設定。デフォルト（空白時、またはTRUE指定）では数式は保護されてクリア対象外
        If UCase(Trim(wsCtrl.Range("E" & i).Value)) = "FALSE" Then
            isProtectFormula = False
        Else
            isProtectFormula = True
        End If

        ' 必要キーのチェック
        If shName <> "" And addr <> "" Then

            Set wsTarget = Nothing
            On Error Resume Next
            Set wsTarget = ThisWorkbook.Sheets(shName)
            On Error GoTo 0

            If Not wsTarget Is Nothing Then
                
                Set rTarget = Nothing
                On Error Resume Next
                Set rTarget = wsTarget.Range(addr)
                On Error GoTo 0

                If Not rTarget Is Nothing Then
                    
                    ' 【数式保護】数式セル（=SUM()など）を誤ってクリアしない保護設定
                    If isProtectFormula And rTarget.HasFormula Then
                        Debug.Print "【数式保護により通過】 " & shName & "!" & addr
                        GoTo NextItem
                    End If
                    
                    ' --- ① データを「履歴」シートにバックアップ退避 ---
                    wsLog.Cells(logRow, 1).Value = Now                         ' 実行日付
                    wsLog.Cells(logRow, 2).Value = shName                      ' シート名
                    wsLog.Cells(logRow, 3).Value = addr                        ' セル番地
                    wsLog.Cells(logRow, 4).Value = rTarget.Value               ' 古いデータ
                    wsLog.Cells(logRow, 5).Value = IIf(IsNull(defVal) Or defVal = "", "【クリア】", defVal) ' 設定した値
                    logRow = logRow + 1
                    
                    ' --- ② 初期化実行 (D列に値があれば代入、なければ完全クリア) ---
                    If IsNull(defVal) Or defVal = "" Then
                        rTarget.ClearContents
                    Else
                        rTarget.Value = defVal
                    End If
                    
                End If
            End If
        End If
NextItem:
    Next i

    ' 画面描画とアラート設定を元に戻す
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "データのクリア初期化と、「履歴」シートへのバックアップ書き込みが完了しました！", vbInformation, "正常終了"

End Sub`,
    explanation: [
      '【タイムスタンプ付き履歴自動保存】: 誤って実行してしまった場合や、過去に消した値を再確認したいときのために、消去前のデータを別の「履歴」シートに日時・対象場所と一緒に自動退避させる究極の安心・安全機能です。',
      '【シートの自動構築コンストラクタ】: 「履歴」シートが無いとき、ユーザーがわざわざ手作業で作らなくても、VBAが「自動でシートを追加して見出し行を追加する」自動修復設計になっています。',
      '【値の消去 vs 初期値セットの二刀流】: D列（デフォルト値）を空白にすると単純に「ClearContents」され、文字や数値を書いておくと「その値でセルをリセット」する多目的イニシャライズ機能を搭載しています。',
      '【数式保護（HasFormula）ガード】: 集計用の `SUM()` や `VLOOKUP()` などの重要数式がセルに入っていた場合、数式を壊さずに実データだけを初期化できるようにする、実務構築で極めて高評価なロジックをE列の保護フラグ連動で解決しています。'
    ]
  },
  {
    id: 'speed-process',
    title: '【高速化】画面更新と自動計算の停止・再開',
    description: 'VBAの処理速度を数倍〜数十倍に高速化するための必須テンプレートコードです。大量のセル操作を行う前に必ず呼び出します。',
    category: 'speed',
    difficulty: 'easy',
    tags: ['高速化', 'お決まり', 'パフォーマンス'],
    code: `Sub FastProcessing()
    ' 画面更新と自動計算を停止（高速化）
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    ' === ここにメインの処理を記述します ===
    Dim i As Long
    For i = 1 To 10000
        Cells(i, 1).Value = "処理中: " & i
    Next i
    ' =====================================
    
    ' 設定を元に戻す
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    MsgBox "処理が完了しました！", vbInformation
    Exit Sub

ErrorHandler:
    ' エラーが発生した場合も必ず設定を元に戻す
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
End Sub`,
    explanation: [
      'Application.ScreenUpdating = False：画面の描画更新を止め、セル編集時のチラつきを防ぎます。',
      'Application.Calculation = xlCalculationManual：数式の自動再計算を停止し、処理ごとの計算負荷を省きます。',
      'Application.EnableEvents = False：マクロ起動中シート内のイベントマクロ（Worksheet_Changeなど）の連鎖起動を防ぎます。',
      'On Error GoTo ErrorHandler：予期せぬエラーが起きた際、Excelの画面更新停止状態が維持されたままロックしないよう復旧用ラベルへ移します。'
    ]
  },
  {
    id: 'outlook-mail',
    title: 'Outlook連携による個別メール自動作成＆送信',
    description: 'Excelシートのアドレス帳から、宛先・件名・本文・添付ファイルを読み込み、自動でメールを作成（またはそのまま送信）します。',
    category: 'outlook',
    difficulty: 'medium',
    tags: ['Outlook', '自動送信', 'メール連携'],
    code: `Sub CreateIndividualMails()
    Dim OutlookApp As Object
    Dim MailItem As Object
    Dim r As Long
    Dim lastRow As Long
    
    ' Outlookアプリケーションインスタンスの作成（起動していない場合は自動起動）
    Set OutlookApp = CreateObject("Outlook.Application")
    
    ' A列の最終行を取得（1行目はヘッダー [宛先|件名|本文|添付パス]）
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For r = 2 To lastRow
        ' 新規メールの作成
        Set MailItem = OutlookApp.CreateItem(0) ' 0はolMailItem
        
        With MailItem
            .To = Cells(r, 1).Value       ' 宛先
            .Subject = Cells(r, 2).Value  ' 件名
            .Body = Cells(r, 3).Value     ' 本文
            
            ' 添付ファイルのパスが設定されている場合
            Dim attachPath As String
            attachPath = Cells(r, 4).Value
            If attachPath <> "" And Dir(attachPath) <> "" Then
                .Attachments.Add attachPath
            End If
            
            ' .Send で直接送信、.Display で下書き確認画面を表示
            .Display
        End With
    Next r
    
    ' 各種オブジェクトの解放
    Set MailItem = Nothing
    Set OutlookApp = Nothing
    MsgBox "メールの作成が完了しました！", vbInformation
End Sub`,
    explanation: [
      'CreateObject("Outlook.Application")：参照設定不要（Late Binding法）で他OfficeアプリからOutlookを呼び出します。',
      'CreateItem(0)：引数「0」はOutlookにおけるメールオブジェクト(olMailItem)に相当します。',
      '.Display：メールの送信画面が立ち上がります。すぐに送る場合は .Send に書き換えます。',
      'Dir(attachPath) <> ""：ファイルが存在するかどうかを事前にVBAでチェックし、添付エラーを防ぎます。'
    ]
  },
  {
    id: 'get-file-list',
    title: '特定フォルダ内のファイル一覧をシートへ書き出し',
    description: '指定したフォルダの中にあるすべてのファイル名、ファイルサイズ、最終更新日時を再帰的に、または同階層一覧でシートに展開します。',
    category: 'file',
    difficulty: 'medium',
    tags: ['ファイル操作', 'FSO', '一覧表示'],
    code: `Sub ExportFileList()
    Dim FSO As Object
    Dim targetFolder As Object
    Dim fileItem As Object
    Dim folderPath As String
    Dim r As Long
    
    ' ダイアログでフォルダを選択させる
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then
            folderPath = .SelectedItems(1)
        Else
            Exit Sub ' キャンセルされた場合
        End If
    End With
    
    ' FileSystemObjectを生成
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set targetFolder = FSO.GetFolder(folderPath)
    
    ' 出力シートの準備
    Cells.Clear
    Cells(1, 1).Value = "ファイル名"
    Cells(1, 2).Value = "ファイルサイズ(KB)"
    Cells(1, 3).Value = "最終更新日時"
    Cells(1, 4).Value = "絶対パス"
    Range("A1:D1").Font.Bold = True
    
    r = 2
    For Each fileItem In targetFolder.Files
        Cells(r, 1).Value = fileItem.Name
        Cells(r, 2).Value = Round(fileItem.Size / 1024, 2)
        Cells(r, 3).Value = fileItem.DateLastModified
        Cells(r, 4).Value = fileItem.Path
        r = r + 1
    Next fileItem
    
    ' 列幅を自動調整
    Columns("A:D").AutoFit
    Set FSO = Nothing
    MsgBox "ファイル一覧を取得しました！", vbInformation
End Sub`,
    explanation: [
      'Application.FileDialog：フォルダ選択用のWindowsダイアログを起動。マクロユーザーにパスを手動で選ばせます。',
      'Scripting.FileSystemObject：Windowsのファイルシステム（FSO）に安全にアクセスし、細かな属性情報を高速で取り出せるライブラリです。',
      'fileItem.Size / 1024：バイト単位の値を、見やすいキロバイト（KB）単位に換算しています。',
      'Columns("A:D").AutoFit：一覧構築後に文字列幅に応じてエクセルシートの列幅を自動整形します。'
    ]
  },
  {
    id: 'last-row-col',
    title: '【超多用】正確な最終行と最終列の番号を取得',
    description: 'データが入力されているエクセルの本当の最終行・最終列を取得します。途中に空欄のある表でも問題なく動作するため一番推奨される手法です。',
    category: 'utility',
    difficulty: 'easy',
    tags: ['最終行', '最終列', 'お決まり'],
    code: `Function GetRealLastRow(ByVal ws As Worksheet) As Long
    ' 下から上に検索させて最後の有効なセルの行番号を得る
    GetRealLastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
End Function

Function GetRealLastColumn(ByVal ws As Worksheet) As Long
    ' 右から左に検索させて非空セルの列番号を得る
    GetRealLastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
End Function

' 使い方サンプル
Sub UseSample()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lr As Long, lc As Long
    lr = GetRealLastRow(ws)
    lc = GetRealLastColumn(ws)
    
    MsgBox "最終行: " & lr & " / 最終列: " & lc, vbInformation
End Sub`,
    explanation: [
      'ws.Rows.Count：Excelの上限行（Excel2007以降は1,048,576行）を指定します。',
      'End(xlUp).Row：最大行から上向きにセルを駆け上がり、最初にぶつかったデータのある行を返します。これによってA列中間の空欄を飛び越えます。',
      'End(xlToLeft).Column：同様に最大列（XFD列）から左向きに移動し、最初にある入力列を見つけ出します。'
    ]
  },
  {
    id: 'sheet-split-save',
    title: '特定列の値に基づいてデータを別ブックに分割して保存',
    description: '「担当者名」や「支店名」などの列の値でデータをふるい、合致するデータのみの新しいExcelブックをそれぞれ自動で作って保存します。',
    category: 'excel',
    difficulty: 'hard',
    tags: ['データ分割', 'ブック保存', '自動書き出し'],
    code: `Sub SplitSheetAndSave()
    Dim wsSource As Worksheet
    Dim wsNew As Worksheet
    Dim dict As Object
    Dim r As Long, lastRow As Long
    Dim keyVal As Variant
    Dim filterCol As Integer
    Dim saveDir As String
    
    Set wsSource = ActiveSheet
    filterCol = 1 ' 例：A列（1列目）に入っている分類名（支店名など）で分割
    saveDir = ActiveWorkbook.Path & "\\" ' 元のブックと同階層に保存
    
    ' 重複を排除した分類キーリストを作成
    Set dict = CreateObject("Scripting.Dictionary")
    lastRow = wsSource.Cells(wsSource.Rows.Count, filterCol).End(xlUp).Row
    
    ' A列の2行目から最終行まで分類を取得しDictionaryに格納
    For r = 2 To lastRow
        keyVal = wsSource.Cells(r, filterCol).Value
        If keyVal <> "" Then
            dict(keyVal) = True
        End If
    Next r
    
    ' 高速化設定
    Application.ScreenUpdating = False
    
    ' 分類ごとに新規ブックへ転記して保存
    Dim k As Variant
    For Each k In dict.Keys
        ' 新規ブック作成
        Dim wbNew As Workbook
        Set wbNew = Workbooks.Add
        Set wsNew = wbNew.Sheets(1)
        
        ' ヘッダー(1行目)をコピー
        wsSource.Rows(1).EntireRow.Copy wsNew.Rows(1)
        
        ' オートフィルターで抽出しコピー
        wsSource.UsedRange.AutoFilter Field:=filterCol, Criteria1:=k
        wsSource.UsedRange.SpecialCells(xlCellTypeVisible).Copy wsNew.Range("A1")
        
        ' フィルター解除
        wsSource.AutoFilterMode = False
        
        ' 新しいブックを保存して閉じる
        On Error Resume Next
        wbNew.SaveAs saveDir & k & ".xlsx"
        wbNew.Close SaveChanges:=False
        On Error GoTo 0
    Next k
    
    Set dict = Nothing
    Application.ScreenUpdating = True
    MsgBox "データの分割保存が完了しました！", vbInformation
End Sub`,
    explanation: [
      'Scripting.Dictionary：キーの重複を防ぐハッシュマップ。これを用いて「支店別」や「担当別」などのユニークリストを洗い出します。',
      'wsSource.UsedRange.AutoFilter：エクセルの標準機能である「オートフィルター」をVBAから作動させ、該当担当者のみ絞り込みます。',
      'SpecialCells(xlCellTypeVisible)：目に映る「可視セル」のみを限定対象としてコピー。これによりフィルタリングされた裏データを除外して美しくコピーできます。',
      'wbNew.SaveAs：新規ファイルを元のブックと同じフォルダ（ActiveWorkbook.Path）に値（キー）ベースのファイル名で自動保存します。'
    ]
  },
  {
    id: 'safe-file-open',
    title: 'ファイルの存在確認と安全なオープン処理',
    description: 'ファイルが他のユーザーに現在開かれて（ロックされて）いるか、あるいはパスが間違っていないかを事前にテストして安全にファイルを開く関数です。',
    category: 'file',
    difficulty: 'medium',
    tags: ['ファイル確認', 'エラー回避', '安全設計'],
    code: `Function IsFileLocked(filePath As String) As Boolean
    Dim fileNum As Integer
    Dim errNum As Long
    
    ' パスが存在しない場合はロック以前に不可能として返す
    If Dir(filePath) = "" Then
        IsFileLocked = False
        Exit Function
    End If
    
    fileNum = FreeFile()
    On Error Resume Next
    ' 排他的アクセスでファイルを開いてみる
    Open filePath For Input Lock Read Write As #fileNum
    errNum = Err.Number
    Close #fileNum
    On Error GoTo 0
    
    ' エラーが発生した(記述子競合など)場合は他の誰かが開いている
    If errNum <> 0 Then
        IsFileLocked = True
    Else
        IsFileLocked = False
    End If
End Function

' 使い方サンプル
Sub OpenExcelBookSafely()
    Dim targetPath As String
    targetPath = "C:\\test\\data_sheet.xlsx"
    
    If Dir(targetPath) = "" Then
        MsgBox "指定のファイルが見つかりません。", vbExclamation
        Exit Sub
    End If
    
    If IsFileLocked(targetPath) Then
        MsgBox "対象ファイルは編集中にロックされています。時間をおいてから実行してください。", vbCritical
    Else
        Workbooks.Open targetPath
    End If
End Sub`,
    explanation: [
      'FreeFile()：未割り当ての安全なファイル登録番号を動的に取得します。',
      'Open filePath For Input Lock Read Write：ファイルを読み書き権限込みの排他モードで一時アクセスします。編集中ロックされている場合はエラー70(書き込み権限なし)等が発行されます。',
      'On Error Resume Next & On Error GoTo 0：一時的にエラーをスリープさせ、直後のErr.Number値を検証することでプログラムを強制落とさずにロック可否を調査します。'
    ]
  }
];
