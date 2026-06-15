import React, { useState } from 'react';
import { motion } from 'motion/react';
import { 
  Download, 
  Copy, 
  Check, 
  FileCode, 
  HelpCircle, 
  BookOpen, 
  Layout, 
  Compass, 
  Settings, 
  Zap,
  Info
} from 'lucide-react';
import { SyntaxHighlighter } from './SyntaxHighlighter';

export const VbaPaletteSection: React.FC = () => {
  const [activeTab, setActiveTab] = useState<'bas' | 'form' | 'guide'>('bas');
  const [copiedStates, setCopiedStates] = useState<{ [key: string]: boolean }>({});

  const files = [
    {
      id: 'bas',
      name: 'FloatingPalette.bas',
      type: '標準モジュール',
      badgeClass: 'bg-emerald-500/10 text-emerald-500 border-emerald-500/20 dark:bg-emerald-400/10 dark:text-emerald-400 dark:border-emerald-400/20',
      description: 'パレットの表示/表示状態チェックおよび各ボタンのアクションロジック、Excel設定シートの自動読み込み・作成を司る中核モジュール。',
      content: `Option Explicit

'------ パレット表示 / 非表示 ------
Public Sub ShowPalette()
    If Not IsPaletteOpen() Then
        frmPalette.Show vbModeless
    Else
        frmPalette.SetFocus
    End If
End Sub

Public Sub HidePalette()
    If IsPaletteOpen() Then Unload frmPalette
End Sub

Private Function IsPaletteOpen() As Boolean
    Dim f As Object
    For Each f In VBA.UserForms
        If f.Name = "frmPalette" Then IsPaletteOpen = True: Exit Function
    Next
End Function

'------ ボタンアクション ------

' テキスト挿入
Public Sub Action_InsertText(sText As String)
    If ActiveCell Is Nothing Then Exit Sub
    If sText = "" Then
        Dim result As String
        result = InputBox("挿入するテキストを入力してください", "テキスト挿入")
        If result = "" Then Exit Sub
        ActiveCell.Value = result
    Else
        ActiveCell.Value = sText
    End If
End Sub

' フォントサイズ変更
Public Sub Action_FontSize()
    If Selection Is Nothing Then Exit Sub
    Dim sz As String
    sz = InputBox("フォントサイズを入力してください（例: 12）", "フォントサイズ", _
                  Selection.Font.Size)
    If sz = "" Then Exit Sub
    If Not IsNumeric(sz) Then MsgBox "数値を入力してください": Exit Sub
    Selection.Font.Size = CDbl(sz)
End Sub

' 背景色変更
Public Sub Action_BgColor()
    If Selection Is Nothing Then Exit Sub
    Dim c As Long
    c = PickColor(Selection.Interior.Color)
    If c = -1 Then Exit Sub
    Selection.Interior.Color = c
End Sub

' 文字色変更
Public Sub Action_FontColor()
    If Selection Is Nothing Then Exit Sub
    Dim c As Long
    c = PickColor(Selection.Font.Color)
    If c = -1 Then Exit Sub
    Selection.Font.Color = c
End Sub

' カメラ（選択範囲をリンク図として貼り付け）
Public Sub Action_Camera()
    If TypeName(Selection) <> "Range" Then
        MsgBox "まずセル範囲を選択してください": Exit Sub
    End If
    Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Paste
    MsgBox "選択範囲を図として貼り付けました。" & Chr(10) & _
           "移動・リサイズしてご利用ください。", vbInformation, "カメラ"
End Sub

' 印刷
Public Sub Action_Print()
    ActiveSheet.PrintOut
End Sub

' カスタムアクション（設定シートから）
Public Sub Action_Custom(sCmd As String)
    On Error Resume Next
    Application.CommandBars.ExecuteMso sCmd
    If Err.Number <> 0 Then
        Err.Clear
        Application.CommandBars.FindControl(ID:=CInt(sCmd)).Execute
    End If
    On Error GoTo 0
End Sub

'------ カラーピッカー（ダイアログ代替）------
Private Function PickColor(defaultColor As Long) As Long
    On Error GoTo ErrHandler
    Dim s As String
    Dim r As Long, g As Long, b As Long
    
    r = (defaultColor Mod 256)
    g = ((defaultColor \\ 256) Mod 256)
    b = ((defaultColor \\ 65536) Mod 256)
    
    s = InputBox("RGB値を入力してください（カンマ区切り: R,G,B）" & Chr(10) & _
                 "例: 255,0,0 = 赤 0,0,255 = 青", "色選択", r & "," & g & "," & b)
    If s = "" Then PickColor = -1: Exit Function
    
    Dim parts() As String
    parts = Split(s, ",")
    If UBound(parts) < 2 Then MsgBox "入力形式が正しくありません": PickColor = -1: Exit Function
    
    PickColor = RGB(CLng(Trim(parts(0))), CLng(Trim(parts(1))), CLng(Trim(parts(2))))
    Exit Function
ErrHandler:
    MsgBox "色の指定に問題があります。": PickColor = -1
End Function

'==============================================================
' 設定シートからボタン情報を読み込む
'==============================================================
Public Function LoadButtonConfig() As Collection
    Dim col As New Collection
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("設定")
    On Error GoTo 0
    
    ' 設定シートがなければデフォルト6ボタン
    If ws Is Nothing Then
        col.Add Array("📝 テキスト", "InsertText", "")
        col.Add Array("🔤 フォントサイズ", "FontSize", "")
        col.Add Array("🎨 背景色", "BgColor", "")
        col.Add Array("🖊 文字色", "FontColor", "")
        col.Add Array("📷 カメラ", "Camera", "")
        col.Add Array("🖨 印刷", "Print", "")
        Set LoadButtonConfig = col
        Exit Function
    End If
    
    ' 設定シートから読み込み（2行目以降）
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 2 Then
        col.Add Array("📝 テキスト", "InsertText", "")
        col.Add Array("🔤 フォントサイズ", "FontSize", "")
        col.Add Array("🎨 背景色", "BgColor", "")
        col.Add Array("🖊 文字色", "FontColor", "")
        col.Add Array("📷 カメラ", "Camera", "")
        col.Add Array("🖨 印刷", "Print", "")
        Set LoadButtonConfig = col
        Exit Function
    End If
    
    Dim i As Long
    For i = 2 To lastRow
        Dim lbl As String, act As String, prm As String
        lbl = Trim(ws.Cells(i, 1).Value)
        act = Trim(ws.Cells(i, 2).Value)
        prm = Trim(ws.Cells(i, 3).Value)
        If lbl <> "" And act <> "" Then
            col.Add Array(lbl, act, prm)
        End If
    Next i
    
    Set LoadButtonConfig = col
End Function

'==============================================================
' 設定シートを自動作成するユーティリティ
'==============================================================
Public Sub CreateSettingSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("設定")
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        If MsgBox("「設定」シートは既に存在します。上書きしますか？", vbYesNo) = vbNo Then Exit Sub
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = "設定"
    
    ' ヘッダー
    With ws
        .Cells(1, 1).Value = "ボタンラベル"
        .Cells(1, 2).Value = "アクション"
        .Cells(1, 3).Value = "パラメータ（任意）"
        .Range("A1:C1").Font.Bold = True
        .Range("A1:C1").Interior.Color = RGB(70, 130, 180)
        .Range("A1:C1").Font.Color = RGB(255, 255, 255)
        
        ' サンプルデータ
        .Cells(2, 1).Value = "📝 テキスト挿入"
        .Cells(2, 2).Value = "InsertText"
        .Cells(2, 3).Value = ""
        
        .Cells(3, 1).Value = "🔤 フォントサイズ"
        .Cells(3, 2).Value = "FontSize"
        .Cells(3, 3).Value = ""
        
        .Cells(4, 1).Value = "🎨 背景色"
        .Cells(4, 2).Value = "BgColor"
        .Cells(4, 3).Value = ""
        
        .Cells(5, 1).Value = "🖊 文字色"
        .Cells(5, 2).Value = "FontColor"
        .Cells(5, 3).Value = ""
        
        .Cells(6, 1).Value = "📷 カメラ"
        .Cells(6, 2).Value = "Camera"
        .Cells(6, 3).Value = ""
        
        .Cells(7, 1).Value = "🖨 印刷"
        .Cells(7, 2).Value = "Print"
        .Cells(7, 3).Value = ""
        
        ' コメント欄
        .Cells(9, 1).Value = "【アクション一覧】"
        .Cells(9, 1).Font.Bold = True
        .Cells(10, 1).Value = "InsertText"
        .Cells(10, 2).Value = "テキスト挿入"
        .Cells(11, 1).Value = "FontSize"
        .Cells(11, 2).Value = "フォントサイズ変更"
        .Cells(12, 1).Value = "BgColor"
        .Cells(12, 2).Value = "背景色変更"
        .Cells(13, 1).Value = "FontColor"
        .Cells(13, 2).Value = "文字色変更"
        .Cells(14, 1).Value = "Camera"
        .Cells(14, 2).Value = "選択範囲を図として貼り付け"
        .Cells(15, 1).Value = "Print"
        .Cells(15, 2).Value = "印刷"
        
        .Columns("A:C").AutoFit
    End With
    
    MsgBox "「設定」シートを作成しました。", vbInformation
End Sub`
    },
    {
      id: 'form',
      name: 'frmPalette_Code.bas',
      type: 'UserForm コード',
      badgeClass: 'bg-indigo-500/10 text-indigo-500 border-indigo-500/20 dark:bg-indigo-400/10 dark:text-indigo-400 dark:border-indigo-400/20',
      description: 'VBAのUserFormである `frmPalette` に貼り付ける、UI配置ロジックと、マウスドラッグでのフローティング追従・動的クリックスルー転記を行うコード。',
      content: `Option Explicit

'------ ドラッグ用変数 ------
Private m_dragging As Boolean
Private m_startX   As Single
Private m_startY   As Single

'------ ボタン配列 ------
Private m_buttons(1 To 6) As MSForms.CommandButton
Private m_config  As Collection

'==============================================================
' フォーム初期化
'==============================================================
Private Sub UserForm_Initialize()
    '--- フォーム外観 ---
    With Me
        .Caption       = "🎨 パレット"
        .Width         = 180
        .Height        = 230
        .BackColor     = RGB(40, 40, 50)
        .BorderStyle   = fmBorderStyleSingle
        .StartUpPosition = 0
        .Left          = Application.Left + Application.Width - 210
        .Top           = Application.Top + 60
        .ShowModal     = False
    End With

    '--- 設定読み込み ---
    Set m_config = LoadButtonConfig()

    '--- タイトルバー代わりのラベル ------
    Dim lbl As MSForms.Label
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblTitle")
    With lbl
        .Caption    = "≡  MyPalette"
        .Left       = 0
        .Top        = 0
        .Width      = Me.Width - 10
        .Height     = 22
        .BackColor  = RGB(60, 60, 80)
        .ForeColor  = RGB(220, 220, 255)
        .Font.Bold  = True
        .Font.Size  = 9
        .TextAlign  = fmTextAlignCenter
    End With

    '--- ボタン配置 ---
    BuildButtons

    '--- 閉じるボタン ---
    Dim btnClose As MSForms.CommandButton
    Set btnClose = Me.Controls.Add("Forms.CommandButton.1", "btnClose")
    With btnClose
        .Caption   = "✕"
        .Left      = Me.Width - 30
        .Top       = 2
        .Width     = 20
        .Height    = 18
        .BackColor = RGB(180, 60, 60)
        .ForeColor = RGB(255, 255, 255)
        .Font.Size = 8
    End With
End Sub

'==============================================================
' ボタン生成
'==============================================================
Private Sub BuildButtons()
    Const BTN_W  As Integer = 72
    Const BTN_H  As Integer = 52
    Const PAD    As Integer = 6
    Const TOP_OFF As Integer = 26
    
    Dim total As Integer
    total = m_config.Count
    If total > 6 Then total = 6
    
    For i = 1 To total
        Dim info As Variant
        info = m_config(i)
        
        Dim col As Integer, row As Integer
        col = (i - 1) Mod 2
        row = (i - 1) \\ 2
        
        Dim btn As MSForms.CommandButton
        Set btn = Me.Controls.Add("Forms.CommandButton.1", "btn" & i)
        With btn
            .Caption   = info(0)
            .Left      = PAD + col * (BTN_W + PAD)
            .Top       = TOP_OFF + PAD + row * (BTN_H + PAD)
            .Width     = BTN_W
            .Height    = BTN_H
            .BackColor = RGB(55, 55, 75)
            .ForeColor = RGB(230, 230, 255)
            .Font.Size = 9
            .WordWrap  = True
        End With
        Set m_buttons(i) = btn
    Next i
    
    Dim rows As Integer
    rows = WorksheetFunction.RoundUp(total / 2, 0)
    Me.Height = TOP_OFF + PAD + rows * (BTN_H + PAD) + PAD + 30
End Sub

'==============================================================
' ドラッグ移動
'==============================================================
Private Sub lblTitle_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then
        m_dragging = True
        m_startX = X
        m_startY = Y
    End If
End Sub

Private Sub lblTitle_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    If m_dragging Then
        Me.Left = Me.Left + (X - m_startX)
        Me.Top  = Me.Top  + (Y - m_startY)
    End If
End Sub

Private Sub lblTitle_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, _
                              ByVal X As Single, ByVal Y As Single)
    m_dragging = False
End Sub

'==============================================================
' ボタンクリック
'==============================================================
Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btn1_Click():  DispatchAction 1:  End Sub
Private Sub btn2_Click():  DispatchAction 2:  End Sub
Private Sub btn3_Click():  DispatchAction 3:  End Sub
Private Sub btn4_Click():  DispatchAction 4:  End Sub
Private Sub btn5_Click():  DispatchAction 5:  End Sub
Private Sub btn6_Click():  DispatchAction 6:  End Sub

Private Sub DispatchAction(idx As Integer)
    If idx > m_config.Count Then Exit Sub
    
    Dim info As Variant
    info = m_config(idx)
    
    Dim act As String, prm As String
    act = info(1)
    prm = info(2)
    
    Select Case act
        Case "InsertText":  Action_InsertText prm
        Case "FontSize":    Action_FontSize
        Case "BgColor":     Action_BgColor
        Case "FontColor":   Action_FontColor
        Case "Camera":      Action_Camera
        Case "Print":       Action_Print
        Case Else:          Action_Custom act
    End Select
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Set m_config = Nothing
End Sub`
    }
  ];

  const handleCopyText = (text: string, key: string) => {
    navigator.clipboard.writeText(text).then(() => {
      setCopiedStates(prev => ({ ...prev, [key]: true }));
      setTimeout(() => {
        setCopiedStates(prev => ({ ...prev, [key]: false }));
      }, 2000);
    }).catch(err => console.error('コピー失敗:', err));
  };

  const handleDownloadFile = (fileName: string, content: string) => {
    // VBAのバックスラッシュの二重表現を解決してエクスポートする
    const fixedContent = content.replace(/\\\\/g, '\\');
    const blob = new Blob([fixedContent], { type: 'text/plain;charset=shift_jis' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const downloadAll = () => {
    files.forEach((file, index) => {
      setTimeout(() => {
        handleDownloadFile(file.name, file.content);
      }, index * 400);
    });
  };

  return (
    <div className="space-y-12 select-none">
      
      {/* イントロ・ヘッダーカード */}
      <div className="p-8 md:p-10 rounded-3xl border border-stone-250/60 dark:border-stone-900 bg-[#FAF7F2]/90 dark:bg-[#121212]/80 backdrop-blur-md shadow-sm relative overflow-hidden group interact-glow">
        <div className="absolute right-[5%] top-0 w-80 h-80 rounded-full bg-indigo-500/[0.02] dark:bg-indigo-400/[0.02] blur-3xl pointer-events-none" />
        <div className="absolute left-0 top-0 w-2 h-full bg-stone-900 dark:bg-[#FAF7F2] rounded-l-3xl" />
        
        <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-6">
          <div className="space-y-3">
            <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-[9px] font-bold font-mono tracking-widest bg-indigo-500/10 text-indigo-600 dark:text-indigo-400 border border-indigo-500/20 uppercase">
              <Zap className="w-3 h-3 text-indigo-500 animate-none" />
               ADVANCED_FLOATING_UI
            </span>
            <h2 className="text-2xl md:text-3.5xl font-black tracking-tight text-stone-950 dark:text-white font-display">
              🎨 フローティングパレット VBA
            </h2>
            <p className="text-xs md:text-sm text-stone-500 dark:text-stone-400 max-w-2xl leading-relaxed">
              Excelのシートやセルの上に常に最前面表示（モーダレス）で浮遊させ、よく使う書式変更やテキスト挿入などの便利アクションをワンクリックで実行できる完全ローカルなカスタムツールバーパレットです。
            </p>
          </div>
          
          <button
            onClick={downloadAll}
            className="flex items-center gap-2 px-5 py-3 rounded-xl bg-stone-950 text-[#FAF7F2] dark:bg-[#FAF7F2] dark:text-stone-950 hover:opacity-90 font-bold text-xs tracking-wider transition-all shadow-sm select-none shrink-0 cursor-pointer border border-transparent"
          >
            <Download className="w-4 h-4" />
            <span>全VBAファイルを一括保存</span>
          </button>
        </div>
      </div>

      {/* メインタブコントロール & コードビューアー */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-8 items-start">
        
        {/* 左: タブナビゲーション & モジュール解説 */}
        <div className="lg:col-span-4 space-y-6">
          <div className="p-5 rounded-2xl border border-stone-250/60 dark:border-stone-900 bg-[#FAF7F2]/90 dark:bg-[#121212]/80 backdrop-blur-md shadow-sm space-y-4">
            <h3 className="text-[10px] font-bold uppercase tracking-widest text-[#8C8A85] dark:text-stone-500 flex items-center gap-2 font-mono">
              <Compass className="w-3.5 h-3.5" />
              PALETTE_MODULE_SELECT
            </h3>

            <div className="flex flex-col gap-1.5">
              {[
                { id: 'bas', label: 'FloatingPalette.bas', desc: '標準モジュール', icon: FileCode },
                { id: 'form', label: 'frmPalette_Code.bas', desc: 'UserForm コード', icon: Layout },
                { id: 'guide', label: '導入手順とドキュメント', desc: 'インストール・パラメータ設定', icon: BookOpen }
              ].map((tab) => {
                const Icon = tab.icon;
                const isActive = activeTab === tab.id;
                return (
                  <button
                    key={tab.id}
                    onClick={() => setActiveTab(tab.id as any)}
                    className={`w-full text-left p-3.5 rounded-xl transition-all flex items-start gap-3 cursor-pointer border ${
                      isActive
                        ? 'bg-stone-950 text-[#FAF7F2] dark:bg-[#FAF7F2] dark:text-stone-950 font-bold border-stone-950 dark:border-[#FAF7F2] shadow-sm'
                        : 'text-stone-600 hover:text-stone-950 border-transparent hover:bg-stone-100/60 dark:text-stone-400 dark:hover:text-white dark:hover:bg-stone-900/50'
                    }`}
                  >
                    <Icon className={`w-4 h-4 mt-0.5 shrink-0 ${isActive ? 'text-indigo-400 dark:text-indigo-600' : 'text-stone-400 dark:text-stone-500'}`} />
                    <div className="space-y-0.5">
                      <div className="text-xs font-bold leading-tight flex items-center gap-1.5 font-mono">
                        {tab.label}
                      </div>
                      <div className={`text-[10px] font-medium leading-normal ${isActive ? 'text-stone-300 dark:text-stone-700' : 'text-stone-400 dark:text-stone-500'}`}>
                        {tab.desc}
                      </div>
                    </div>
                  </button>
                );
              })}
            </div>
          </div>

          {/* クイック統計 / 特徴 */}
          <div className="p-5 rounded-2xl border border-stone-250/60 dark:border-stone-900 bg-white/50 dark:bg-[#121212]/85 text-stone-850 dark:text-stone-200 shadow-sm space-y-4 backdrop-blur-md">
            <h4 className="text-[10px] font-bold uppercase tracking-widest text-[#8C8A85] dark:text-stone-500 flex items-center gap-2 font-mono">
              <Info className="w-3.5 h-3.5" />
              SYSTEM_ADVANTAGES
            </h4>
            
            <div className="space-y-3.5 text-xs">
              <div className="flex items-start gap-2.5">
                <span className="w-1.5 h-1.5 bg-emerald-500 rounded-full mt-1.5 block shrink-0" />
                <p className="leading-relaxed text-stone-600 dark:text-stone-400 font-medium">
                  <strong className="text-stone-900 dark:text-white font-bold">完全モーダレス対応 (vbModeless)</strong><br />
                  パレットを立ち上げたままでもセルの編集やスクロールなどのExcel通常操作を100%妨げません。
                </p>
              </div>
              <div className="flex items-start gap-2.5">
                <span className="w-1.5 h-1.5 bg-emerald-500 rounded-full mt-1.5 block shrink-0" />
                <p className="leading-relaxed text-stone-600 dark:text-stone-400 font-medium">
                  <strong className="text-stone-900 dark:text-white font-bold">設定シート自動連携読み込み</strong><br />
                  ボタンラベルやマクロアクションは、「設定」シートを作成すれば、コード側の修正なしで動的に書き換え可能です。
                </p>
              </div>
              <div className="flex items-start gap-2.5">
                <span className="w-1.5 h-1.5 bg-emerald-500 rounded-full mt-1.5 block shrink-0" />
                <p className="leading-relaxed text-stone-600 dark:text-stone-400 font-medium">
                  <strong className="text-stone-900 dark:text-white font-bold">フォーム擬似ドラッグ追従制御</strong><br />
                  標準の野暮ったいWindowsタイトルバーを非表示にし、モダンに設計されたユーザーフォーム内部の見出しラベルから滑らかに掴んで移動可能です。
                </p>
              </div>
            </div>
          </div>
        </div>

        {/* 右: ソースコード / ガイド本体 */}
        <div className="lg:col-span-8 space-y-6">
          {activeTab !== 'guide' ? (
            files.map((file) => {
              if (file.id !== activeTab) return null;
              return (
                <div key={file.id} className="space-y-4">
                  {/* ヘッダー */}
                  <div className="flex items-center justify-between border-b border-stone-250/30 dark:border-stone-900/60 pb-3">
                    <div className="flex items-center gap-3">
                      <span className={`text-[10px] font-bold px-2 py-0.5 rounded-md border font-mono tracking-wider ${file.badgeClass}`}>
                        {file.type}
                      </span>
                      <span className="text-xs font-bold text-stone-900 dark:text-white font-mono">
                        {file.name}
                      </span>
                    </div>
                    
                    <button
                      onClick={() => handleDownloadFile(file.name, file.content)}
                      className="flex items-center gap-1.5 px-3 py-1.5 text-[10px] font-mono font-bold rounded-lg transition-all text-stone-700 hover:text-stone-950 dark:text-stone-300 dark:hover:text-white bg-stone-100 dark:bg-white/[0.04] border border-stone-250/50 dark:border-white/[0.03] cursor-pointer"
                    >
                      <Download className="w-3.5 h-3.5" />
                      <span>DOWNLOAD_FILE</span>
                    </button>
                  </div>

                  <p className="text-[11px] md:text-xs text-stone-500 dark:text-stone-400 leading-relaxed font-sans font-medium px-1 bg-stone-100/50 dark:bg-[#000000]/10 p-2 rounded-xl">
                    {file.description}
                  </p>

                  <SyntaxHighlighter code={file.content} language="vba" />
                </div>
              );
            })
          ) : (
            <div className="space-y-8">
              {/* === 導入手順ドキュメントの内容 === */}
              <div className="flex items-center justify-between border-b border-stone-250/30 dark:border-stone-900/60 pb-3">
                <div className="flex items-center gap-3">
                  <span className="text-[10px] font-bold px-2 py-0.5 rounded-md border font-mono tracking-wider bg-amber-500/10 text-amber-500 dark:bg-amber-400/10 dark:text-amber-400 border-amber-500/20 dark:border-amber-400/20">
                    ドキュメント
                  </span>
                  <span className="text-xs font-bold text-stone-900 dark:text-white font-mono">
                    導入ガイド ＆ 設計書
                  </span>
                </div>
                
                <button
                  onClick={() => handleCopyText(`VBAフローティングパレットガイド\n\n1. 標準モジュールのインポート\n2. UserForm作成（名前：frmPalette）とコードペースト\n3. ショートカットキーの設定\n4. 設定シート作成`, 'guide-text')}
                  className="flex items-center gap-1.5 px-3 py-1.5 text-[10px] font-mono font-bold rounded-lg transition-all text-stone-700 hover:text-stone-950 dark:text-stone-300 dark:hover:text-white bg-stone-100 dark:bg-white/[0.04] border border-stone-250/50 dark:border-white/[0.03] cursor-pointer"
                >
                  {copiedStates['guide-text'] ? <Check className="w-3.5 h-3.5 text-emerald-500" /> : <Copy className="w-3.5 h-3.5" />}
                  <span>{copiedStates['guide-text'] ? 'COPIED_OK' : 'COPY_SUMMARY'}</span>
                </button>
              </div>

              {/* 手順の視覚化ステップ */}
              <div className="space-y-6">
                <h3 className="text-sm font-bold text-stone-900 dark:text-white flex items-center gap-2 select-none font-mono tracking-widest uppercase">
                  <Settings className="w-4 h-4 text-stone-500" />
                  STEPS_TO_INSTALL
                </h3>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  {[
                    {
                      step: 'STEP 1',
                      title: '標準モジュールインポート',
                      desc: 'Excelで対象ブックを開き、「Alt + F11」でVBEを開きます。メニューから「ファイルのインポート」を選択し「FloatingPalette.bas」を取り込みます。'
                    },
                    {
                      step: 'STEP 2',
                      title: 'UserForm の作成 ＆ 設定',
                      desc: 'VBEメニュー「挿入」→「ユーザーフォーム」を選択。追加されたフォームのオブジェクト名を必ず「frmPalette」に変更。ダブルクリックして中身に「frmPalette_Code.bas」をペーストします。'
                    },
                    {
                      step: 'STEP 3',
                      title: 'ショートカットキー割り当て',
                      desc: 'Excel画面に戻り、「開発」タブ → 「マクロ」→ 「ShowPalette」を選択して「オプション」から任意のキー（例: Ctrl + Shift + P）を設定します。'
                    },
                    {
                      step: 'STEP 4',
                      title: '動的な設定シート自動構築',
                      desc: 'マクロから「CreateSettingSheet」を実行すると、Excel上にボタンレイアウトや割り当てアクションをカスタマイズ可能な「設定」シートが自動生成されます。'
                    }
                  ].map((s, idx) => (
                    <div key={idx} className="p-5 rounded-2xl border border-stone-250/50 dark:border-stone-900 bg-[#FAF8F5]/50 dark:bg-[#121212]/80 backdrop-blur-md shadow-sm transition-all duration-300 relative group overflow-hidden interact-glow">
                      <div className="absolute top-0 right-0 p-3 font-mono text-[9px] font-bold text-stone-400 dark:text-stone-600 select-none">
                        {s.step}
                      </div>
                      <h4 className="text-xs font-bold text-stone-900 dark:text-white mb-2 tracking-tight">
                        {s.title}
                      </h4>
                      <p className="text-[11px] text-stone-500 dark:text-stone-400 leading-relaxed font-sans font-medium">
                        {s.desc}
                      </p>
                    </div>
                  ))}
                </div>
              </div>

              {/* 設定シート・アクション仕様 */}
              <div className="p-6 rounded-2xl border border-stone-250/60 dark:border-stone-900/85 bg-[#FAF7F2]/40 dark:bg-[#121212]/30 space-y-4">
                <h4 className="text-xs font-bold text-stone-900 dark:text-white uppercase tracking-wider flex items-center gap-2">
                  <Layout className="w-4 h-4 text-indigo-500" />
                  カスタムアクション ＆ 設定シート構成
                </h4>
                <p className="text-xs text-stone-500 dark:text-stone-400 leading-relaxed">
                  マクロ「CreateSettingSheet」を実行すると作成されるシート。テーブルにそって最大6個（2列×3行）のグリッドボタンが初期化されます。A列が「ラベル名」、B列が「アクション定義名」です。
                </p>

                <div className="overflow-x-auto rounded-xl border border-stone-250 dark:border-stone-850">
                  <table className="w-full text-xs text-left text-stone-600 dark:text-stone-300">
                    <thead className="bg-stone-100 dark:bg-stone-900 text-stone-700 dark:text-stone-300 font-bold uppercase tracking-wider">
                      <tr>
                        <th className="px-4 py-2.5">登録アクション</th>
                        <th className="px-4 py-2.5">実行される内容</th>
                        <th className="px-4 py-2.5">パラメータ（初期値等）</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-stone-200 dark:divide-stone-850">
                      <tr>
                        <td className="px-4 py-2 font-semibold font-mono text-indigo-600 dark:text-indigo-400">InsertText</td>
                        <td className="px-4 py-2 font-medium">セルに任意の文字列・初期テキストを流し込みます。</td>
                        <td className="px-4 py-2 font-mono text-stone-400">省略時はInputBoxが出現</td>
                      </tr>
                      <tr>
                        <td className="px-4 py-2 font-semibold font-mono text-indigo-600 dark:text-indigo-400">FontSize</td>
                        <td className="px-4 py-2 font-medium">選択されたセル領域全体のフォントサイズ（CDbl）を一瞬で調整します。</td>
                        <td className="px-4 py-2 font-mono text-stone-400">不要</td>
                      </tr>
                      <tr>
                        <td className="px-4 py-2 font-semibold font-mono text-indigo-600 dark:text-indigo-400">BgColor</td>
                        <td className="px-4 py-2 font-medium">セル背景色を選択します（擬似RGB変換カラーダイアログ代替）。</td>
                        <td className="px-4 py-2 font-mono text-stone-400">不要</td>
                      </tr>
                      <tr>
                        <td className="px-4 py-2 font-semibold font-mono text-indigo-600 dark:text-indigo-400">FontColor</td>
                        <td className="px-4 py-2 font-medium">セルの文字カラーを更新します。</td>
                        <td className="px-4 py-2 font-mono text-stone-400">不要</td>
                      </tr>
                      <tr>
                        <td className="px-4 py-2 font-semibold font-mono text-indigo-600 dark:text-indigo-400">Camera</td>
                        <td className="px-4 py-2 font-medium">選択範囲を図としてキャプチャし、リンク図（CopyPicture）を同じシートに。</td>
                        <td className="px-4 py-2 font-mono text-stone-400">不要</td>
                      </tr>
                      <tr>
                        <td className="px-4 py-2 font-semibold font-mono text-indigo-600 dark:text-indigo-400">Print</td>
                        <td className="px-4 py-2 font-medium">アクティブな全体シートを安全に用紙出力（PrintOut）します。</td>
                        <td className="px-4 py-2 font-mono text-stone-400">不要</td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              </div>

              {/* F.A.Q.アコーディオン */}
              <div className="space-y-4">
                <h4 className="text-xs font-bold text-stone-900 dark:text-white uppercase tracking-widest flex items-center gap-2 font-mono">
                  <HelpCircle className="w-4 h-4 text-stone-500" />
                  FREQUENTLY_ASKED_QUESTIONS
                </h4>

                <div className="space-y-3">
                  {[
                    {
                      q: 'Q: Excelを開いたと同時にパレットを自動的に起動させたい',
                      a: 'VBE画面で、ファイル共通オブジェクトである「ThisWorkbook」モジュールに以下のようにコードを記述します。\n\nPrivate Sub Workbook_Open()\n    ShowPalette\nEnd Sub\n\nこれにより、読み込み完了と同時に自動起動が実行されます。'
                    },
                    {
                      q: 'Q: ボタンの数を6個以上に増やすことはできませんか？',
                      a: '初期設定では「2列×3行」のレイアウトに最大化制限（6個）を設けていますが、「frmPalette_Code.bas」内の「BuildButtons」にある「If total > 6 Then total = 6」や定数を拡張し、グリッドの割り算（コロン Mod / Row 引数など）を調整することで「3×3=9個」などに広げることができます。'
                    },
                    {
                      q: 'Q: フォームを閉じても時折Excel側に残っている気がします',
                      a: '通常の「✕」ボタンクリックで「Unload Me」が確実に走り、メモリ解放（QueryCloseイベント内にてSet m_config = Nothing）を行っています。また、「HidePalette」マクロをマクロ一覧（F8）から手動実行することで、安全に解放させることができます。'
                    }
                  ].map((faq, idx) => (
                    <div key={idx} className="p-4 rounded-xl border border-stone-200 dark:border-stone-900 bg-white dark:bg-stone-950 shadow-xs space-y-2 select-none">
                      <h5 className="text-xs font-bold text-stone-900 dark:text-white flex items-center gap-2">
                        <span className="text-indigo-500 font-mono">●</span>
                        {faq.q}
                      </h5>
                      <p className="text-[11px] text-stone-500 dark:text-stone-400 leading-relaxed whitespace-pre-line font-medium pl-4.5">
                        {faq.a}
                      </p>
                    </div>
                  ))}
                </div>
              </div>

            </div>
          )}
        </div>

      </div>

    </div>
  );
};
