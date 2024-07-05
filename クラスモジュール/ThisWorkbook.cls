VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()


    Dim TMS用原本ws As Worksheet
    Set TMS用原本ws = Worksheets("TMS反映用表原本")

    Dim 生徒用原本ws As Worksheet
    Set 生徒用原本ws = Worksheets("生徒用操作画面原本")
    
    Dim 講師用原本ws As Worksheet
    Set 講師用原本ws = Worksheets("講師用操作画面原本")
    
    '非表示シートを一度表示
    TMS用原本ws.Visible = True
    生徒用原本ws.Visible = True
    講師用原本ws.Visible = True
    
    Call 起動時にTMS反映用表を複製(TMS用原本ws)
    Call 起動時に生徒用操作画面を複製(生徒用原本ws)
    Call 起動時に講師用操作画面を複製(講師用原本ws)
    
    '原本を再度非表示
    TMS用原本ws.Visible = False
    生徒用原本ws.Visible = False
    講師用原本ws.Visible = False
    
    
End Sub

Private Sub 起動時にTMS反映用表を複製(ByVal TMS用原本ws As Variant)


    For Each シート検索 In Sheets               'すでに操作画面のシートがないかFor Eachループで探す
    If シート検索.Name = "TMS反映用表" Then '"TMS反映用表"があれば
        Exit Sub
    End If
    Next シート検索
    
    TMS用原本ws.Copy after:=TMS用原本ws 'Worksheets("TMS反映用表原本")を複製
    ActiveSheet.Name = "TMS反映用表"    '複製したシートに名前をつける
        
        
End Sub

Private Sub 起動時に講師用操作画面を複製(ByVal 講師用原本ws As Variant)


    For Each シート検索 In Sheets               'すでに操作画面のシートがないかFor Eachループで探す
    If シート検索.Name = "講師用操作画面" Then '"講師用操作画面"があれば
        Exit Sub
    End If
    Next シート検索
    
    講師用原本ws.Copy after:=講師用原本ws 'Worksheets("講師用操作画面原本")を複製
    ActiveSheet.Name = "講師用操作画面"   '複製したシートに名前をつける


End Sub

Private Sub 起動時に生徒用操作画面を複製(ByVal 生徒用原本ws As Variant)


    For Each シート検索 In Sheets               'すでに操作画面のシートがないかFor Eachループで探す
    If シート検索.Name = "生徒用操作画面" Then '"生徒用操作画面"があれば
        Exit Sub
    End If
    Next シート検索
    
    生徒用原本ws.Copy after:=生徒用原本ws 'Worksheets("生徒用操作画面原本")を複製
    ActiveSheet.Name = "生徒用操作画面"   '複製したシートに名前をつける


End Sub
