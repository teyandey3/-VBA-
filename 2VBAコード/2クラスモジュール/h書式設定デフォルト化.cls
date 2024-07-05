VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "h書式設定デフォルト化"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub 書式設定デフォルト化(ByVal 対象ws As Worksheet)


    '余白の上下左右、行の高さ、列の幅、フォント、字大きさの初期化
    対象ws.Range("Q3").Value = 2.5         '余白の上
    対象ws.Range("R3").Value = 2.5         '余白の下
    対象ws.Range("S3").Value = 2           '余白の右
    対象ws.Range("T3").Value = 2           '余白の左
    対象ws.Range("U3").Value = 0.48        '行の高さ
    対象ws.Range("V3").Value = 7.35        '列の幅
    対象ws.Range("W3").Value = "Yu Gothic" 'フォント
    対象ws.Range("X3").Value = 10          '字大きさ
    
    '列に入力する日数、一日のコマ数、見出しの束数の初期化
    対象ws.Range("Q7").Value = 14 '列に入力する日数
    対象ws.Range("U7").Value = 9  '一日のコマ数
    対象ws.Range("Y7").Value = 3  '見出しの束数の初期化
    
        '各コマの開始時間の初期化
    対象ws.Range("Q9:XFD9").ClearContents
    
    Dim 時刻 As Variant: 時刻 = Array("日/時", "13:00", "14:00", "15:00", "16:00", "17:00", "18:00", "19:00", "20:00", "21:00")
    
    対象ws.Range("Q9:Z9").Value = 時刻
    
    '条件付き書式の色設定
    '色番号
    対象ws.Range("AD2").Value = 9
    対象ws.Range("AD3").Value = 5
    対象ws.Range("AD4").Value = 6
    対象ws.Range("AD5").Value = 10
    対象ws.Range("AD6").Value = 8
    対象ws.Range("AD7").Value = 7
    '色濃度
    対象ws.Range("AE2").Value = 0.6
    対象ws.Range("AE3").Value = 0.8
    対象ws.Range("AE4").Value = 0.8
    対象ws.Range("AE5").Value = 0.6
    対象ws.Range("AE6").Value = 0.6
    対象ws.Range("AE7").Value = 0.6


End Sub
