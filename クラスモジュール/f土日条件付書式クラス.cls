VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "f土日条件付書式クラス"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub 土日条件付書式(ByVal 対象ws As Worksheet)


    Dim 書式設定一  As FormatCondition
    
    With 対象ws.Rows("1:1048576")                                                                                     'A1からexcelの最終行である1048576行目を利用しているため修正不要
        .FormatConditions.Delete                                                                                      '条件付き書式設定を初期
        Set 書式設定一 = .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(WEEKDAY(A1,2)>=6,A1<>"""",A1<>1)") '選択範囲に(Formula1:=)と(Type:=xlExpression)を用いて関数(=AND(WEEKDAY(A1,2)>=6,A1<>"",A1<>1))を適用
        With 書式設定一.Interior
            .ThemeColor = 対象ws.Range("AD7").Value                                                                   '色と濃度を"色番号と濃度表"の各セルの値を参照して指定
            .TintAndShade = 対象ws.Range("AE7").Value
        End With
    End With
            
            
End Sub
