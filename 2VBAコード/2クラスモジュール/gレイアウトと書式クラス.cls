VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "gレイアウトと書式クラス"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private 余白上 As Double
Private 余白下 As Double
Private 余白右 As Double
Private 余白左 As Double
Private 行高 As Double
Private 列幅 As Double
Private フォント As String
Private 文字大 As Double

Public Sub ページレイアウトと文字書式(ByVal 変数取得元ws As Worksheet, ByVal 編集先ws As Worksheet, シート名 As String)


    If シート名 = "TMS用ws" Then
        余白上 = 変数取得元ws.Range("A3").Value   '上の余白はA3セルを参照
        余白下 = 変数取得元ws.Range("B3").Value   '下の余白はB3セルを参照
        余白右 = 変数取得元ws.Range("C3").Value   '左の余白はC3セルを参照
        余白左 = 変数取得元ws.Range("D3").Value   '右の余白はD3セルを参照
        行高 = 変数取得元ws.Range("E3").Value     '行の高さはE3セルを参照
        列幅 = 変数取得元ws.Range("F3").Value     '列の幅はF3セルを参照
        フォント = 変数取得元ws.Range("G3").Value '文字書式はG3セルを参照
        文字大 = 変数取得元ws.Range("H3").Value   '文字サイズはH3セルを参照
    Else
        余白上 = 変数取得元ws.Range("Q3").Value   '上の余白はQ3セルを参照
        余白下 = 変数取得元ws.Range("R3").Value   '下の余白はR3セルを参照
        余白右 = 変数取得元ws.Range("S3").Value   '左の余白はS3セルを参照
        余白左 = 変数取得元ws.Range("T3").Value   '右の余白はT3セルを参照
        行高 = 変数取得元ws.Range("U3").Value     '行の高さはU3セルを参照
        列幅 = 変数取得元ws.Range("V3").Value     '列の幅はV3セルを参照
        フォント = 変数取得元ws.Range("W3").Value '文字書式はW3セルを参照
        文字大 = 変数取得元ws.Range("X3").Value   '文字サイズはX3セルを参照
    End If

    'ページレイアウト
    With 編集先ws.PageSetup
        .Orientation = xlLandscape                              '印刷向きを横にする
        .PaperSize = xlPaperA4                                  'ページサイズをA4に設定する（VBA画面でのみ変更可能）
        .TopMargin = Application.CentimetersToPoints(余白上)    '上の余白を変更 (Application.CentimetersToPointsはcmを1ポイント=約0.35mmに変換するメソッド)
        .BottomMargin = Application.CentimetersToPoints(余白下) '下の余白を変更
        .LeftMargin = Application.CentimetersToPoints(余白右)   '左の余白を変更
        .RightMargin = Application.CentimetersToPoints(余白左)  '右の余白を変更
        .HeaderMargin = 0                                       'ヘッダーを0にする
        .FooterMargin = 0                                       'フッターを0にする
        .CenterVertically = True                                '印刷位置を垂直方向のページ中央にする
        .CenterHorizontally = True                              '印刷位置を水平方向のページ中央にする
    End With
    
    '文字書式
    With 編集先ws.Cells
        .RowHeight = Application.CentimetersToPoints(行高) '行の高さを変更
        .ColumnWidth = 列幅                                '列の幅を変更
        .Font.Name = フォント                              'フォントを変更
        .Font.Size = 文字大                                '文字サイズを変更
        .VerticalAlignment = xlTop                         '垂直方向の文字配置を上揃えにする
        .HorizontalAlignment = xlCenter                    '水平方向の文字配置を中央揃えにする
    End With


End Sub
