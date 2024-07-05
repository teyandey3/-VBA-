Attribute VB_Name = "E生徒用操作画面複製"
Option Explicit

Sub ボタン⑤生徒用操作画面シートを複製するプログラム()
Attribute ボタン⑤生徒用操作画面シートを複製するプログラム.VB_ProcData.VB_Invoke_Func = " \n14"


    Dim 生徒用ws As Worksheet
    Set 生徒用ws = Worksheets("生徒用操作画面")

    Dim 表変 As a表作成用変数クラス
    Set 表変 = New a表作成用変数クラス
    Call 表変.表作成用変数初期化(生徒用ws, "生徒用ws")
    
    '同じシートを二度作成することを防ぐ
    Dim 名前 As String: 名前 = 生徒用ws.Range("B1") 'Worksheets("2操作画面（編集厳禁）")から生徒名をコピー（VBA画面でのみ変更可能）
    Dim シート検索 As Worksheet
    
    If 名前 = "" Then 'もし生徒名が入力されていない場合、エラーコードを表示
        MsgBox "生徒名を入力してください。"
        Exit Sub
    End If 'エラーがなければ、以下のプログラムを実行
    
    For Each シート検索 In Sheets                               'シートの中から同じ生徒名のシートがないかFor Eachループで探す
        If シート検索.Name = 名前 Then
            MsgBox "同じ生徒の希望日程表を二度作成しています。" 'あればエラーメッセージを表示
            Exit Sub
        End If
    Next シート検索 'エラーがなければ、以下のプログラムを実行
   
    Application.ScreenUpdating = False '画面描画を停止
    Application.Cursor = xlWait 'ウエイトカーソル
    Application.EnableEvents = False 'イベントを抑止
    Application.DisplayAlerts = False '確認メッセージを抑止
    Application.Calculation = xlCalculationManual '計算を手動に
   
    'Worksheets("2操作画面（編集厳禁）")を複製
    生徒用ws.Copy after:=生徒用ws
    ActiveSheet.Name = "仮名"  '複製したシートに仮名をつける
    
    '新しい表の調整
    Dim 新表変 As a表作成用変数クラス
    Set 新表変 = New a表作成用変数クラス
    Call 新表変.表作成用変数初期化(生徒用ws, "仮名")
    
    Dim 表調整 As b削除クラス
    Set 表調整 = New b削除クラス
    Call 表調整.新シート表加工用(Worksheets("仮名"), 新表変.表行始, 新表変.表列始, 表変.表行始, 新表変.表列終) '新シートの表の始点の行から、旧シートの表の始点の一つ上の行まで削除
                                                                                   '(注）この部分は旧シートの表の始点にセットすること
    Dim 土書式 As gレイアウトと書式クラス
    Set 土書式 = New gレイアウトと書式クラス
    Call 土書式.ページレイアウトと文字書式(生徒用ws, Worksheets("仮名"), "仮名")

    Dim 条書式 As y条件付書式設定
    Set 条書式 = New y条件付書式設定
    Call 条書式.条件付書式設定(生徒用ws, Worksheets("仮名"), 新表変.表行始, 新表変.表列始, 新表変.表行終, 新表変.表列終)
    
    '作成日を入力
    Worksheets("仮名").Range("B2").Value = Date
    
    '複製したシートの名前を生徒名に変更
    Worksheets("仮名").Name = 名前
    
    '初期化できるようにK5セルの判定マーカーを削除
    生徒用ws.Range("K5").ClearContents
    
    'ファイルを保存
    ActiveWorkbook.Save
    
    Application.Calculation = xlCalculationAutomatic '計算を自動に
    Application.DisplayAlerts = True '確認メッセージを開始
    Application.EnableEvents = True 'イベントを開始
    Application.Cursor = xlDefault '標準カーソル
    Application.ScreenUpdating = True '画面描画を開始
    
    
End Sub
