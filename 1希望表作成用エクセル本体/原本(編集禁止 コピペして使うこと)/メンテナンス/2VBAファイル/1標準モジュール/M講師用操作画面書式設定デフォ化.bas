Attribute VB_Name = "M講師用操作画面書式設定デフォ化"
Option Explicit

Sub ボタン⑮講師用操作画面_表の設定をすべてデフォルトにする()


    Application.ScreenUpdating = False '画面描画を停止
    Application.Cursor = xlWait 'ウエイトカーソル
    Application.EnableEvents = False 'イベントを抑止
    Application.DisplayAlerts = False '確認メッセージを抑止
    Application.Calculation = xlCalculationManual '計算を手動に

    Dim 講師用ws As Worksheet
    Set 講師用ws = Worksheets("講師用操作画面")

    Dim デフォ化 As h書式設定デフォルト化
    Set デフォ化 = New h書式設定デフォルト化
    Call デフォ化.書式設定デフォルト化(講師用ws)

    Application.Calculation = xlCalculationAutomatic '計算を自動に
    Application.DisplayAlerts = True '確認メッセージを開始
    Application.EnableEvents = True 'イベントを開始
    Application.Cursor = xlDefault '標準カーソル
    Application.ScreenUpdating = True '画面描画を開始

    
End Sub

