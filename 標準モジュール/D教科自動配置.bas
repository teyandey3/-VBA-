Attribute VB_Name = "d教科自動配置"
Option Explicit

Sub ボタン④教科を自動で振り分けるプログラム()


    Dim 生徒用ws As Worksheet
    Set 生徒用ws = Worksheets("生徒用操作画面")

    Dim 表変 As a表作成用変数クラス
    Set 表変 = New a表作成用変数クラス
    Call 表変.表作成用変数初期化(生徒用ws, "生徒用ws")
    
    '入力された各教科の合計数を変数に格納
    Dim 季コマ数合計 As Integer: 季コマ数合計 = 生徒用ws.Range("E2").Value - 1 '配列が0番目から始まるため-1する
    
    '各教科のコマ数の合計が希望表にあるコマ数の合計と一致しているか確認
    If Not 生徒用ws.Range("E3").Value - 1 = 季コマ数合計 Then
        MsgBox "季節講習の各教科のコマ数の合計と希望コマ数の合計が一致していません。" 'もし一致していなければ、エラーメッセージを表示
        Exit Sub
    ElseIf 生徒用ws.Range("E3").Value = "" _
    Or 生徒用ws.Range("E3").Value = 0 Then
        MsgBox "上記の季節講習のコマ数表に各教科のコマ数を入力してください。" 'もし合計が0なら、エラーメッセージを表示
        Exit Sub
    End If 'もし一致していたら、以下のプログラムを実行
    
    Application.ScreenUpdating = False '画面描画を停止
    Application.Cursor = xlWait 'ウエイトカーソル
    Application.EnableEvents = False 'イベントを抑止
    Application.DisplayAlerts = False '確認メッセージを抑止
    Application.Calculation = xlCalculationManual '計算を手動に
    
    'シャッフル後の配列"教科"をセルに貼り付け
    Dim 貼付 As z教科自動配置クラス
    Set 貼付 = New z教科自動配置クラス
    
    Call 貼付.教科探索とセル貼付(生徒用ws, 表変.表行始, 表変.表列始, 表変.表行終, 表変.表列終, 季コマ数合計)
    
    '選択範囲を消去
    生徒用ws.Range("E3").ClearContents
    
    '関数を挿入（条件付き書式で色を付けたい教科を増やすときはこのコードを増やす）
    生徒用ws.Range("E3").Formula = "=SUM(F3:J3)"               '（VBA画面でのみ変更可能）
    生徒用ws.Range("F3").Formula = "=COUNTIF(17:1048576,R13)"  '（VBA画面でのみ変更可能）
    生徒用ws.Range("G3").Formula = "=COUNTIF(17:1048576,S13)"  '（VBA画面でのみ変更可能）
    生徒用ws.Range("H3").Formula = "=COUNTIF(17:1048576,T13)"  '（VBA画面でのみ変更可能）
    生徒用ws.Range("I3").Formula = "=COUNTIF(17:1048576,U13)"  '（VBA画面でのみ変更可能）
    生徒用ws.Range("J3").Formula = "=COUNTIF(17:1048576,V13)"  '（VBA画面でのみ変更可能）
    生徒用ws.Range("K3").Formula = "=COUNTIF(17:1048576,W13)"  '（VBA画面でのみ変更可能）
    生徒用ws.Range("L3").Formula = "=COUNTIF(17:1048576,X13)"  '（VBA画面でのみ変更可能）
    生徒用ws.Range("M3").Formula = "=COUNTIF(17:1048576,Y13)"  '（VBA画面でのみ変更可能）
    生徒用ws.Range("N3").Formula = "=COUNTIF(17:1048576,Z13)"  '（VBA画面でのみ変更可能）
    生徒用ws.Range("O3").Formula = "=COUNTIF(17:1048576,AA13)" '（VBA画面でのみ変更可能）

    '次の作業を指示するメッセージの表示
    MsgBox "「確認用」の欄の数字が入力したコマ数と同じかを確認して" + vbCrLf + "ボタン④を押してください。"

    Application.Calculation = xlCalculationAutomatic '計算を自動に
    Application.DisplayAlerts = True '確認メッセージを開始
    Application.EnableEvents = True 'イベントを開始
    Application.Cursor = xlDefault '標準カーソル
    Application.ScreenUpdating = True '画面描画を開始


End Sub
