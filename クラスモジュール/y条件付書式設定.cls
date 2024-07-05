VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "y条件付書式設定"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub 条件付書式設定(ByVal 変数取得元ws As Worksheet, ByVal 編集先ws As Worksheet, 表行始 As Integer, 表列始 As Integer, 表行終 As Integer, 表列終 As Integer)


    Dim 書式設定一  As FormatCondition
    Dim 書式設定二  As FormatCondition
    Dim 書式設定三  As FormatCondition
    Dim 書式設定四  As FormatCondition
    Dim 書式設定五  As FormatCondition
    
    '---（以下のセル表記はすべてVBA画面でのみ変更可能、条件付き書式で色を付けたい教科を増やすときはこのコードを増やす）---
    With 編集先ws.Range(Cells(表行始, 表列始), Cells(表行終, 表列終)) '教科別の書式設定
        Set 書式設定一 = .FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:=変数取得元ws.Range("R13").Value) 'セルの値(Formula1:=Range("Q11"))を参照して(Type:=xlCellValue、例:季　英語)と完全一致(Operator:=xlEqual)しているセルに書式を適用
        With 書式設定一.Interior
            .ThemeColor = 変数取得元ws.Range("AD2")   'ThemeColorは色一覧の列
            .TintAndShade = 変数取得元ws.Range("AE2") 'TintAndShadeは基本色 +QQ%の場所（％を小数表記）
        End With
        Set 書式設定二 = .FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:=変数取得元ws.Range("S13").Value)
        With 書式設定二.Interior
            .ThemeColor = 変数取得元ws.Range("AD3")
            .TintAndShade = 変数取得元ws.Range("AE3")
        End With
        Set 書式設定三 = .FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:=変数取得元ws.Range("T13").Value)
        With 書式設定三.Interior
            .ThemeColor = 変数取得元ws.Range("AD4")
            .TintAndShade = 変数取得元ws.Range("AE4")
        End With
        Set 書式設定四 = .FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:=変数取得元ws.Range("U13").Value)
        With 書式設定四.Interior
            .ThemeColor = 変数取得元ws.Range("AD5")
            .TintAndShade = 変数取得元ws.Range("AE5")
        End With
        Set 書式設定五 = .FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:=変数取得元ws.Range("V13").Value)
        With 書式設定五.Interior
            .ThemeColor = 変数取得元ws.Range("AD6")
            .TintAndShade = 変数取得元ws.Range("AE6")
        End With
    End With
    '---ここまで---


End Sub
