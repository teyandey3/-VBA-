VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "b削除クラス"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub 表全削除(ByVal 対象ws As Worksheet, ByVal 表行始 As Integer, ByVal 表列始 As Integer)


    Dim 表上削 As Integer: 表上削 = 表行始 - 1                                  '表の一行上の位置を特定
    対象ws.Range(Cells(表上削, 表列始), Cells(Rows.Count, Columns.Count)).Clear '表の一行上から最終行まで削除


End Sub

Public Sub 表周辺削除(ByVal 対象ws As Worksheet, ByVal 表行始 As Integer, ByVal 表列始 As Integer, ByVal 表行終 As Integer, ByVal 表列終 As Integer)


    Dim 表上削 As Integer: 表上削 = 表行始 - 1                       '表の一行上の位置を特定
    Dim 表下削 As Integer: 表下削 = 表行終 + 1                       '表の一行下の位置を特定
    Dim 表右削 As Integer: 表右削 = 表列終 + 1                       '表の一行右の位置を特定
    対象ws.Range(Cells(表上削, 表列始), Cells(表上削, 表右削)).Clear '表の一行上を削除
    対象ws.Range(Cells(表下削, 表列始), Cells(表下削, 表右削)).Clear '表の一行下を削除
    対象ws.Range(Cells(表上削, 表右削), Cells(表下削, 表右削)).Clear '表の一行右を削除
    
    '余分な部分についた条件付き書式設定を初期
    対象ws.Range(Cells(2, 1), Cells(表上削, Columns.Count)).FormatConditions.Delete '一行目は条件付き書式の基準点となっているため書式を削除しない


End Sub

Public Sub 新シート表加工用(ByVal 対象ws As Worksheet, ByVal 表行始 As Integer, 表列始 As Integer, ByVal 表行終 As Integer, 表列終 As Integer)

    
    Dim 表右削 As Integer: 表右削 = 表列終 + 1 '表の一行右の位置を特定
    Dim 旧表上削 As Integer: 旧表上削 = 表行終 - 1 '表の一行上の位置を特定
    
    With 対象ws                  '余分な部分の削除
        .Range(Cells(表行始, 表列始), Cells(旧表上削, Columns.Count)).Delete 'シート"仮名"の3行目から16行目を削除
        .Range(Cells(1, 表右削), Cells(Rows.Count, Columns.Count)).Delete  'P1セルから右の列を削除
    End With
    
    '余分な部分についた条件付き書式設定を初期
    Dim 表上削 As Integer: 表上削 = 表行始 - 1 '表の一行上の位置を特定
    
    対象ws.Range(Cells(2, 1), Cells(表上削, Columns.Count)).FormatConditions.Delete '一行目は条件付き書式の基準点となっているため書式を削除しない


End Sub
