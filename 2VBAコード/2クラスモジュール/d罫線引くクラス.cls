VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "d罫線引くクラス"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub 罫線引く(ByVal 対象ws As Worksheet, ByVal 表行始 As Integer, ByVal 表列始 As Integer, ByVal 表行終 As Integer, ByVal 表列終 As Integer, ByVal コマ数 As Integer)

    
    '格子状の罫線の設定
    対象ws.Range(Cells(表行始, 表列始), Cells(表行終, 表列終)).Borders.LineStyle = xlContinuous 'A15セルから表の最終セルまで、線の細さ(Weight)はデフォルトの細(xlThin)
    
    '列の罫線
    With 対象ws.Range(Cells(表行始, 表列始), Cells(表行終, 表列始))
        .Borders(xlEdgeLeft).Weight = xlThick                                       '外枠(左)を引く　線の細さ(Weight)は最太(xlThin)
        .Borders(xlEdgeRight).Weight = xlThick                                      'A列(時刻見出し)の右側に太線を引く
    End With
    対象ws.Range(Cells(表行始, 表列終), Cells(表行終, 表列終)).Borders(xlEdgeRight).Weight = xlThick '外枠(右)を引く
    
    '行の罫線
    Dim 罫行始 As Integer
    Dim 罫線終点 As Integer: 罫線終点 = 表行終 + 1                     '(ループの回数分+1)のスペースを確保することで、罫線を最終行{外枠(下)}まで引く
    For 罫行始 = 表行始 To 罫線終点 Step コマ数
        With 対象ws.Range(Cells(罫行始, 表列始), Cells(罫行始, 表列終))
            .Borders(xlEdgeTop).Weight = xlThick                        '時刻見出しの上の行に引く{表の外枠(上)を作る}
            .Borders(xlEdgeBottom).Weight = xlThick                     '時刻見出しの下の行に太線を引く
        End With
    Next 罫行始


End Sub
