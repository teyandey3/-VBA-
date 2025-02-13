Attribute VB_Name = "L講師用操作画面書式設定"
Option Explicit

Sub ボタン�M講師用操作画面書式設定()


    Application.ScreenUpdating = False '画面描画を停止
    Application.Cursor = xlWait 'ウエイトカーソル
    Application.EnableEvents = False 'イベントを抑止
    Application.DisplayAlerts = False '確認メッセージを抑止
    Application.Calculation = xlCalculationManual '計算を手動に

    Dim 講師用ws As Worksheet
    Set 講師用ws = Worksheets("講師用操作画面")
   
    Dim 書式 As gレイアウトと書式クラス
    Set 書式 = New gレイアウトと書式クラス

    Call 書式.ページレイアウトと文字書式(講師用ws, 講師用ws, "講師用ws")

    'K5セルの文字書式
    講師用ws.Range("K5").HorizontalAlignment = xlLeft '水平方向の文字配置を左揃えにする
    'Q3セルの文字書式
    講師用ws.Range("U4").HorizontalAlignment = xlLeft '水平方向の文字配置を左揃えにする
    
    'AF1セルの文字書式
    講師用ws.Range("AF1").HorizontalAlignment = xlLeft '水平方向の文字配置を左揃えにする
    'AF2セルの文字書式
    講師用ws.Range("AF2").HorizontalAlignment = xlLeft '水平方向の文字配置を左揃えにする

    Application.Calculation = xlCalculationAutomatic '計算を自動に
    Application.DisplayAlerts = True '確認メッセージを開始
    Application.EnableEvents = True 'イベントを開始
    Application.Cursor = xlDefault '標準カーソル
    Application.ScreenUpdating = True '画面描画を開始


End Sub
