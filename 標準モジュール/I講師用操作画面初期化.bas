Attribute VB_Name = "I講師用操作画面初期化"
Option Explicit

Sub ボタン⑪講師用操作画面初期化プログラム()


    Dim 講師用ws As Worksheet
    Set 講師用ws = Worksheets("講師用操作画面")

    Dim 表変 As a表作成用変数クラス
    Set 表変 = New a表作成用変数クラス
    Call 表変.表作成用変数初期化(講師用ws, "講師用ws")

    'ボタン②の二度押しを防ぐためのプログラム
    If 講師用ws.Range("K5").Value = "判定マーカー(消さないで)" Then 'K5セル（VBA画面でのみ変更可能）の文字を取得して、もし間違えてボタン②を連続で二度押したら、エラーコードを表示
        MsgBox "表の編集途中でボタン①を押しています。編集途中で初期化したい場合は" + vbCrLf + "K5セルの「判定マーカー(消さないで)」を消した後にボタン①を押してください。"
        Exit Sub
    End If 'エラーがなければ、以下のプログラムを実行

    Application.ScreenUpdating = False '画面描画を停止
    Application.Cursor = xlWait 'ウエイトカーソル
    Application.EnableEvents = False 'イベントを抑止
    Application.DisplayAlerts = False '確認メッセージを抑止
    Application.Calculation = xlCalculationManual '計算を手動に

    '表の削除
    Dim 表削 As b削除クラス
    Set 表削 = New b削除クラス
    Call 表削.表全削除(講師用ws, 表変.表行始, 表変.表列始)

    講師用ws.Range("B2") = ""             'セルB2の内容を消去（VBA画面でのみ変更可能）
    講師用ws.Range("E2:O2").ClearContents 'セルE2からJ2の内容を消去（VBA画面でのみ変更可能）
    講師用ws.Range("E3:O3").ClearContents 'セルE3からJ3の内容を消去（VBA画面でのみ変更可能）
    
    '時刻見出作成を作成するクラスの呼び出し
    Dim 時見出 As c時刻見出作成クラス
    Set 時見出 = New c時刻見出作成クラス
    Call 時見出.縦時刻見出作成(講師用ws, "講師用ws", 表変.表行始, 表変.表列始, 表変.表行終, 表変.コマ数)

    '罫線を引くクラスの呼び出し
    Dim 罫線 As d罫線引くクラス
    Set 罫線 = New d罫線引くクラス
    Call 罫線.罫線引く(講師用ws, 表変.表行始, 表変.表列始, 表変.表行終, 表変.表列終, 表変.コマ数)
    
    '選択範囲に文字を入力
    講師用ws.Cells(表変.表行始, 表変.開始日列).Value = "開始日"
    
    '次の作業を指示するメッセージの表示
    MsgBox "講習開始日をB17のセルに入力してください。"

    Application.Calculation = xlCalculationAutomatic '計算を自動に
    Application.DisplayAlerts = True '確認メッセージを開始
    Application.EnableEvents = True 'イベントを開始
    Application.Cursor = xlDefault '標準カーソル
    Application.ScreenUpdating = True '画面描画を開始


End Sub
