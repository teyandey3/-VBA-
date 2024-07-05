Attribute VB_Name = "C生徒用操作画面表加工"
Option Explicit

Sub ボタン②生徒用操作画面新規シートを日程表に加工するプログラム()


    Dim 生徒用ws As Worksheet
    Set 生徒用ws = Worksheets("生徒用操作画面")

    Dim 表変 As a表作成用変数クラス
    Set 表変 = New a表作成用変数クラス
    Call 表変.表作成用変数初期化(生徒用ws, "生徒用ws")

    'ボタン②の二度押しを防ぐためのプログラム
    If 生徒用ws.Range("K5").Value = "判定マーカー(消さないで)" Then 'K5セル（VBA画面でのみ変更可能）の文字を取得して、もし間違えてボタン②を連続で二度押したら、エラーコードを表示
        MsgBox "ボタン②を連続で二度押しています。この機能は現在使うことができません。" + vbCrLf + "ボタン③を押すか、最初から作業をやり直してください。"
        Exit Sub
    End If 'エラーがなければ、以下のプログラムを実行
    
    '開始日を入力したか判定するためのプログラム
    If 生徒用ws.Cells(表変.表行始, 表変.開始日列).Value = "開始日" Then 'もし講習開始日をB15のセルに入力していなかったらエラーコードを表示
        MsgBox "講習開始日を入力してください。"
        Exit Sub                                                   'プログラムの終了
    End If
    
    Application.ScreenUpdating = False '画面描画を停止
    Application.Cursor = xlWait 'ウエイトカーソル
    Application.EnableEvents = False 'イベントを抑止
    Application.DisplayAlerts = False '確認メッセージを抑止
    Application.Calculation = xlCalculationManual '計算を手動に
    
    '日付を自動入力するクラスの呼び出し
    Dim 日付 As e日付自動入力クラス
    Set 日付 = New e日付自動入力クラス
    Call 日付.日付自動入力(生徒用ws, 表変.表行始, 表変.表列始, 表変.表行終, 表変.表列終, 表変.コマ数, 表変.見出束数, 表変.開始日列)

    '土日に色を付ける条件付き書式を設定するクラス(プロシージャ)の呼び出し
    Dim 土書式 As f土日条件付書式クラス
    Set 土書式 = New f土日条件付書式クラス
    Call 土書式.土日条件付書式(生徒用ws)
    
    '選択範囲を消去
    Dim 表削 As b削除クラス
    Set 表削 = New b削除クラス
    Call 表削.表周辺削除(生徒用ws, 表変.表行始, 表変.表列始, 表変.表行終, 表変.表列終)
    
    '上記のプログラムで消えてしまった表の下外枠を付けなおす
    生徒用ws.Range(Cells(表変.表行終, 表変.表列始), Cells(表変.表行終, 表変.表列終)).Borders(xlEdgeBottom).Weight = xlThick
    
    '関数を挿入
    生徒用ws.Range("E2").Formula = "=SUM(F2:O2)"
    生徒用ws.Range("E3").Formula = "=COUNTIF(17:1048576,""=0"")" '1048576はエクセルの最終行なので変更不要。（"15"→最初の行はVBA画面でのみ変更可能）
    生徒用ws.Range("K4").Formula = "=IF(E2=E3, """", ""エラー: 各教科のコマ数の合計と希望日程表に入力されたコマ数の合計が一致しません。"")"
        
    'ボタン②を連続で二度押すことを防ぐためのエラー判別マーカーを挿入
    生徒用ws.Range("K5") = "判定マーカー(消さないで)" '（VBA画面でのみ変更可能）
        
    '次の操作を指示するメッセージの表示
    MsgBox "各教科のコマ数を入力してください。"
    
    Application.Calculation = xlCalculationAutomatic '計算を自動に
    Application.DisplayAlerts = True '確認メッセージを開始
    Application.EnableEvents = True 'イベントを開始
    Application.Cursor = xlDefault '標準カーソル
    Application.ScreenUpdating = True '画面描画を開始
    
    
End Sub
