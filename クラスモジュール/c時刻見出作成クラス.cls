VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c時刻見出作成クラス"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private 時取行 As Integer
Private 時取列始 As Integer

Private Function 時刻見出配列作成(ByVal 対象ws As Worksheet, ByVal シート名 As String, ByVal コマ数 As Integer)


    If シート名 = "TMS用ws" Then
        時取行 = 対象ws.Range("A9").Row
        時取列始 = 対象ws.Range("A9").Column
    Else
        時取行 = 対象ws.Range("Q9").Row
        時取列始 = 対象ws.Range("Q9").Column
    End If
    
    Dim 時取回数 As Integer: 時取回数 = コマ数 - 1 '授業開始時刻を入力するセル(Q14から始まる)の最終列を測定（配列が0番目からスタートするため-1する）
    ReDim 時刻配(時取回数) As String
    
    Dim 時取回変 As Integer
    For 時取回変 = 0 To 時取回数
        時刻配(時取回変) = Format(対象ws.Cells(時取行, 時取列始).Value, "hh:mm") 'セル内の時刻をStringに直しながら配列"時刻配"に格納
        時取列始 = 時取列始 + 1
    Next 時取回変
    
    時刻見出配列作成 = 時刻配
    
    
End Function

Public Sub 縦時刻見出作成(ByVal 対象ws As Worksheet, ByVal シート名 As String, ByVal 表行始 As Integer, ByVal 表列始 As Integer, ByVal 表行終 As Integer, ByVal コマ数 As Integer)


    Dim 時刻配() As String: 時刻配() = 時刻見出配列作成(対象ws, シート名, コマ数)
        
    Dim 見行始 As Integer
    Dim 一見行終 As Integer: 一見行終 = 表行始 + コマ数 - 1 '一束の時刻見出しの最終行
    For 見行始 = 表行始 To 表行終 Step コマ数                                                                    '入力された見出し数の分だけ繰り返す
        対象ws.Range(Cells(見行始, 表列始), Cells(一見行終, 表列始)).Value = WorksheetFunction.Transpose(時刻配) '配列"時刻配()"をTransposeで横→縦に転置し表の始点から順に配置
        一見行終 = 一見行終 + コマ数
    Next 見行始
 

End Sub
