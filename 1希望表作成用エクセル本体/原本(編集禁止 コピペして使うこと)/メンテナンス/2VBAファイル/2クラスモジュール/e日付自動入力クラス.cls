VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "e日付自動入力クラス"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub 日付自動入力(ByVal 対象ws As Worksheet, ByVal 表行始 As Integer, ByVal 表列始 As Integer, ByVal 表行終 As Integer, ByVal 表列終 As Integer, ByVal コマ数 As Integer, ByVal 見出束数 As Integer, ByVal 開始日列 As Integer)

  
    '日付を保管する配列の作成
    Dim 日付 As Date: 日付 = 対象ws.Cells(表行始, 開始日列).Value '入力したデータを参照して日付を自動入力する
    Dim 日配() As Variant
    Dim 日列数 As Integer: 日列数 = 表列終 - 表列始               '一束に入る日付の個数
    Dim 日数 As Integer: 日数 = 日列数 * 見出束数 + 見出束数 - 1  '配列"日配"の要素数。-1は配列の0番目から格納するため-。+日束数は"日束数"個分の"番兵"を追加するため
    ReDim 日配(日数) As Variant                                   '配列を初期化
    
    '日付を配列に格納
    Dim 日数変 As Integer
    For 日数変 = 0 To 日数
        If 日数変 = 日列数 Then        'もし変数 "日変"が行幅と同じとき
            日配(日数変) = "番兵"      '配列"日付配"に"番兵"というマーカーを追加
            日列数 = 日列数 + 表列終
        Else
            日配(日数変) = 日付
            日付 = 日付 + 1          '日付を一つずらす
            GoTo Continue
        End If
Continue:
    Next 日数変

    '日付をセルに代入
    Dim 日貼配 As Variant                                   'セルに一括貼り付けをするために、配列"日配"から要素を表の横のサイズに合う個数だけ格納する配列
    Dim 日貼列数 As Integer: 日貼列数 = 表列終 - 表列始 - 1 '一束に入る日付の個数を再定義。ただし、配列"日貼配"が0番目から始まるので-1する。
    ReDim 日貼配(日貼列数) As Variant                       '配列を初期化、配列は"日列数"個を格納可能
    Dim 日取Index As Integer: 日取Index = 0                 '配列"日付配"から要素を取り出す変数
    
    For 日数変 = 0 To 日数
        If 日配(日数変) = "番兵" Then                                         'もし配列"日付配"から"番兵"というマーカーを取り出したとき
            With 対象ws.Range(Cells(表行始, 開始日列), Cells(表行始, 表列終))
                .Value = 日貼配                                               '配列"日貼配"をセルに代入する。(注)以下五行の順番を変えないこと
                .NumberFormatLocal = "m/d(aaa)"                               '日付表記をyyyy/mm/ddから"m""月""d""日"" (月)に変更"
            End With
            日取Index = 0                                                     '配列"日貼配"の0番目の要素から配列"日配"の要素を再び格納開始
            表行始 = 表行始 + コマ数
        Else
            日貼配(日取Index) = 日配(日数変)                                  '配列"日付配"の0番目から"日列数"番目の要素を取り出し順に配列"日貼配"の0番目から"日列数"番目まで順に代入していく
            日取Index = 日取Index + 1                                         '変数をひとつずらす
            GoTo 継続
        End If
継続:
  Next 日数変
    
    
End Sub
