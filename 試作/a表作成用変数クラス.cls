VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "a表作成用変数クラス"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
'表の形を決定する変数
Private 表行始Pri As Integer
Private 表行終Pri As Integer
Private 表列始Pri As Integer
Private 開始日列Pri As Integer
Private 表列終Pri As Integer
Private コマ数Pri As Integer
Private 見出束数Pri As Integer

Public Sub 表作成用変数初期化(ByVal 変数取得元ws As Worksheet, シート名 As String)


    If シート名 = "生徒用ws" Then                    'もしシート名が"生徒用ws"なら
        表行始Pri = 変数取得元ws.Range("A17").Row    '表の始点をA17セルにセットする（VBA画面でのみ変更可能）
        表列始Pri = 変数取得元ws.Range("A17").Column '（VBA画面でのみ変更可能）
    ElseIf シート名 = "講師用ws" Then                'もしシート名が"講師用ws"なら
        表行始Pri = 変数取得元ws.Range("A13").Row    '表の始点をA13セルにセットする（VBA画面でのみ変更可能）
        表列始Pri = 変数取得元ws.Range("A13").Column '（VBA画面でのみ変更可能）
    ElseIf シート名 = "仮名" Then                    'もしシート名が"仮名"なら
        表行始Pri = 変数取得元ws.Range("A3").Row     '表の始点をA3セルにセットする（VBA画面でのみ変更可能）
        表列始Pri = 変数取得元ws.Range("A3").Column  '（VBA画面でのみ変更可能）
    ElseIf シート名 = "TMS用ws" Then                     'もしシート名が"TMS用ws"なら
        表行始Pri = 変数取得元ws.Range("A45").Row        '表の始点をA37セルにセットする（VBA画面でのみ変更可能）
        表列始Pri = 変数取得元ws.Range("A45").Column     '（VBA画面でのみ変更可能）
        
        表列終Pri = 変数取得元ws.Range("D7").Value       '時刻見出しが横向きの表の最終列はコマ数と同じになる
        コマ数Pri = 変数取得元ws.Range("G7").Value + 1   '講師一人の担当生徒数（生徒数+講師名を記入するセル一つ分）
        見出束数Pri = 変数取得元ws.Range("J7").Value     '一日に出勤する講師数
        表行終Pri = コマ数Pri * 見出束数Pri + 表行始 - 1 '表の最終行をコマ数と見出しの束数から計算して変数に格納
        Exit Sub
    End If
    
    開始日列Pri = 表列始Pri + 1                      '日付の開始場所を特定する（"日/時"見出しが入る一つ分を除く）

    表列終Pri = 変数取得元ws.Range("Q7").Value + 1   '表の最終列を変数に格納（列に入力する日数+"日付/時間"セル一つ分）

    コマ数Pri = 変数取得元ws.Range("U7").Value + 1   'コマ数+"日付/時間"セル一つ分
 
    見出束数Pri = 変数取得元ws.Range("Y7").Value     '日付見出しの個数
    
    表行終Pri = コマ数Pri * 見出束数Pri + 表行始 - 1 '表の最終行をコマ数と見出しの束数から計算して変数に格納
    

End Sub

'---以下の"Property Get"はPrivate指定した変数の取得用アクセサ---
Property Get 表行始() As Integer
  表行始 = 表行始Pri
End Property
Property Get 表列始() As Integer
  表列始 = 表列始Pri
End Property
Property Get 開始日列() As Integer
  開始日列 = 開始日列Pri
End Property
Property Get 表行終() As Integer
  表行終 = 表行終Pri
End Property
Property Get 表列終() As Integer
  表列終 = 表列終Pri
End Property
Property Get コマ数() As Integer
  コマ数 = コマ数Pri
End Property
Property Get 見出束数() As Integer
  見出束数 = 見出束数Pri
End Property
