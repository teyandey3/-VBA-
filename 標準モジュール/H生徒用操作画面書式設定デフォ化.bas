Attribute VB_Name = "h生徒用操作画面書式設定デフォ化"
Option Explicit

Sub ボタン⑧生徒用操作画面_表の設定をすべてデフォルトにする()


    Application.ScreenUpdating = False '画面描画を停止
    Application.Cursor = xlWait 'ウエイトカーソル
    Application.EnableEvents = False 'イベントを抑止
    Application.DisplayAlerts = False '確認メッセージを抑止
    Application.Calculation = xlCalculationManual '計算を手動に

    Dim 生徒用ws As Worksheet
    Set 生徒用ws = Worksheets("生徒用操作画面")

    Dim デフォ化 As h書式設定デフォルト化
    Set デフォ化 = New h書式設定デフォルト化
    Call デフォ化.書式設定デフォルト化(生徒用ws)

    '教科名の初期化
    生徒用ws.Range("Q11:XFD11").ClearContents
    生徒用ws.Range("Q13:XFD13").ClearContents
    
    Dim 通教科名 As Variant: 通教科名 = Array("英語", "数学", "国語", "理科", "社会")
    Dim 季教科名 As Variant: 季教科名 = Array("季　英語", "季　数学", "季　国語", "季　理科", "季　社会")
    生徒用ws.Range("R11:V11").Value = 通教科名
    生徒用ws.Range("R13:V13").Value = 季教科名
    
    '季節科目数の初期化
    生徒用ws.Range("Q11").Value = 5
    
    '通常科目数の初期化
    生徒用ws.Range("Q13").Value = 5
    
    Application.Calculation = xlCalculationAutomatic '計算を自動に
    Application.DisplayAlerts = True '確認メッセージを開始
    Application.EnableEvents = True 'イベントを開始
    Application.Cursor = xlDefault '標準カーソル
    Application.ScreenUpdating = True '画面描画を開始

  
End Sub
