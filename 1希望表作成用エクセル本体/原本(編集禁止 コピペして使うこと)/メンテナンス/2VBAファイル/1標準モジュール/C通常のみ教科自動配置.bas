Attribute VB_Name = "C通常のみ教科自動配置"
Option Explicit

Sub ボタン③通常教科のみの表を作成するプログラム()


    Application.ScreenUpdating = False '画面描画を停止
    Application.Cursor = xlWait 'ウエイトカーソル
    Application.EnableEvents = False 'イベントを抑止
    Application.DisplayAlerts = False '確認メッセージを抑止
    Application.Calculation = xlCalculationManual '計算を手動に

    Dim 生徒用ws As Worksheet
    Set 生徒用ws = Worksheets("生徒用操作画面")

    Dim 表変 As a表作成用変数クラス
    Set 表変 = New a表作成用変数クラス
    Call 表変.表作成用変数初期化(生徒用ws, "生徒用ws")

    'シャッフル後の配列"教科"をセルに貼り付け
    Dim 貼付 As z教科自動配置クラス
    Set 貼付 = New z教科自動配置クラス
    
    '通常教科のみ入力するため、コ数合計には0を代入する。
    Call 貼付.教科探索とセル貼付(生徒用ws, 表変.表行始, 表変.表列始, 表変.表行終, 表変.表列終, 0)

    '次の作業を指示するメッセージの表示
    MsgBox "ボタン⑤を押してください。"

    Application.Calculation = xlCalculationAutomatic '計算を自動に
    Application.DisplayAlerts = True '確認メッセージを開始
    Application.EnableEvents = True 'イベントを開始
    Application.Cursor = xlDefault '標準カーソル
    Application.ScreenUpdating = True '画面描画を開始


End Sub
