VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "z教科自動配置クラス"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Function 季節教科名格納(ByVal 対象ws As Worksheet) '季節講習の教科名を格納する配列


    Dim 季教名取行 As Integer: 季教名取行 = 対象ws.Range("R13").Row                                 '（VBA画面でのみ変更可能）
    Dim 季名取列始 As Integer: 季名取列始 = 対象ws.Range("R13").Column                              '（VBA画面でのみ変更可能）
    Dim 季名取列終 As Integer: 季名取列終 = 対象ws.Range("R13").Column + 対象ws.Range("Q13").Value - 1 '（VBA画面でのみ変更可能）
    
    Dim 季教名配 As Variant: 季教名配 = 対象ws.Range(Cells(季教名取行, 季名取列始), Cells(季教名取行, 季名取列終))
    
    季節教科名格納 = 季教名配


End Function

Private Function 通常教科名格納(ByVal 対象ws As Worksheet) '通常授業の教科名を格納する配列


    Dim 通教名取行 As Integer: 通教名取行 = 対象ws.Range("R11").Row                                    '（VBA画面でのみ変更可能）
    Dim 通名取列始 As Integer: 通名取列始 = 対象ws.Range("R11").Column                                 '（VBA画面でのみ変更可能
    Dim 通名取列終 As Integer: 通名取列終 = 対象ws.Range("R11").Column + 対象ws.Range("Q11").Value - 1 '（VBA画面でのみ変更可能
    Dim 通教名配 As Variant: 通教名配 = 対象ws.Range(Cells(通教名取行, 通名取列始), Cells(通教名取行, 通名取列終))
    
    通常教科名格納 = 通教名配
    
    
End Function

Private Function 貼付用配列作成(ByVal 対象ws As Worksheet, ByVal コ数合計 As Integer)
    
    
    '季節講習の教科名を格納するための関数呼び出し
    ReDim 季教名配(コ数合計) As Variant
    季教名配() = 季節教科名格納(対象ws)
    
    '各教科のコマ数を格納する配列
    Dim 季教科数 As Variant: 季教科数 = 対象ws.Range("Q13").Value     '表から季節講習で実施する教科数を取得（VBA画面でのみ変更可能）
    Dim コマ数取行 As Integer: コマ数取行 = 対象ws.Range("F2").Row    '（VBA画面でのみ変更可能）
    Dim コ数取列始 As Integer: コ数取列始 = 対象ws.Range("F2").Column '（VBA画面でのみ変更可能）
    Dim コ数取列終 As Integer: コ数取列終 = コ数取列始 + 季教科数 - 1 '（VBA画面でのみ変更可能）
    Dim コマ数配 As Variant: コマ数配 = 対象ws.Range(Cells(コマ数取行, コ数取列始), Cells(コマ数取行, コ数取列終)).Value '表から各教科のコマ数を取得

    '季節講習の各教科をセルに貼り付けるためのシャッフル前(未置換)の配列
    ReDim 教貼未換配(コ数合計) As Variant
    
    
    '日付を配列"教貼未換配()"に格納
    Dim コ数配Index As Variant                       '二次元配列"コマ数配"から要素を取り出すための変数
    Dim 季名配limits As Integer: 季名配limits = -1   '配列"季教名配"から教科を取り出す回数を決める（ループの回数の上限値）配列は0番目からのため-1する。
    Dim 教貼配Index As Integer
    For コ数配Index = 1 To 季教科数
        季名配limits = 季名配limits + コマ数配(1, コ数配Index)
        For 教貼配Index = 教貼配Index To 季名配limits
            教貼未換配(教貼配Index) = 季教名配(1, コ数配Index)
        Next 教貼配Index
    Next
    
    貼付用配列作成 = 教貼未換配()
    
    
End Function
    
Private Function 教科シャッフル(ByVal 対象ws As Worksheet, ByVal コ数合計 As Integer)


    '季節講習の各教科をセルに貼り付けるための配列を格納する関数呼び出し
    ReDim 教貼換済配(コ数合計) As Variant
    教貼換済配() = 貼付用配列作成(対象ws, コ数合計) 'シャッフル後(置換済)の教科を格納するための配列を用意

    Dim Index終 As Integer: Index終 = UBound(教貼換済配) '配列の最後のインデックス番号を格納する変数
    Dim 乱数 As Integer                                  '乱数を格納する変数
    Dim 教科tmp As String                                'ランダムに選ばれた教科を一時的に格納する変数
    Randomize                                            '乱数の初期化
    
    Dim 換Index As Integer
    For 換Index = Index終 To 0 Step -1         '最大のIndex番号から0まで減少させながら、以下を(最大のIndex番号+1=配列の要素数)回分実行する
        乱数 = Int((換Index + 1) * Rnd)        '乱数の生成
        教科tmp = 教貼換済配(乱数)             '一次元配列"教納済()"に0以上要素数以下のランダムな整数を代入する
        教貼換済配(乱数) = 教貼換済配(換Index) '教納済(乱数)を教納済(換変)に交換する
        教貼換済配(換Index) = 教科tmp          '教納済(換変)を"教科tmp"(=教貼換済配(換変)代入前の教貼換済配(乱数))に交換する
    Next 換Index

    教科シャッフル = 教貼換済配()


End Function

Public Sub 教科探索とセル貼付(ByVal 対象ws As Worksheet, ByVal 表行始 As Integer, ByVal 表列始 As Integer, ByVal 表行終 As Integer, ByVal 表列終 As Integer, ByVal コ数合計 As Integer)

    
    '季節講習の各教科をセルに貼り付けるため配列をシャッフルする関数呼び出し
    ReDim 教貼用配(コ数合計) As Variant
    教貼用配() = 教科シャッフル(対象ws, コ数合計)
    
    '季節講習の教科名を格納するための関数呼び出し
    Dim 通教科数 As Integer: 通教科数 = 対象ws.Range("Q11").Value
    ReDim 通教名配(通教科数) As Variant
    通教名配() = 通常教科名格納(対象ws)
    
    Dim 表行数 As Integer: 表行数 = 表行終 - 表行始 + 1                                          '表の行数
    Dim セルtmp As Integer                                                                       'セルの値を一時保存する変数
    Dim セル As Variant: セル = 対象ws.Range(Cells(表行始, 表列始), Cells(表行終, 表列終)).Value 'シート指定された範囲(A15セルから表の最終セルまで)を配列"セル"に格納する
    Dim 教貼用配Index As Integer: 教貼用配Index = 0                                              '配列"教貼用配"から要素を取り出すための変数
    
    Dim 表列Index As Integer
    Dim 表行Index As Integer
    For 表列Index = 1 To 表列終 '←もっと早いアルゴリズムがあるのでは？
        For 表行Index = 1 To 表行数                                  '縦一列目から下方向に検索開始
            If セル(表行Index, 表列Index) = "0" Then                 'もしセルに"0"と表示されていたら
                セル(表行Index, 表列Index) = 教貼用配(教貼用配Index) '一次元配列"探貼教乱済"の0番目から順に教科を選んで"0"を置き換える
                教貼用配Index = 教貼用配Index + 1                    '教貼換済配Indexを一つずらす
            ElseIf セル(表行Index, 表列Index) >= "1" _
            And セル(表行Index, 表列Index) <= 通教科数 Then          'もしセルに通常科目の識別番号が入力されていたら
                セルtmp = セル(表行Index, 表列Index)
                セル(表行Index, 表列Index) = 通教名配(1, セルtmp)    'セルを参照して二次元配列"探通教名配"の中から識別番号の教科を選んで、そのセルに代入
            Else                                                     'もし空欄や"0"と通常科目の識別番号外の数字が入力されていたら
                GoTo Continue                                        '上記の作業をスキップする
            End If
Continue:
        Next 表行Index
    Next 表列Index                                                    '一列目が終わったら二列目の検索を開始
    
    対象ws.Range(Cells(表行始, 表列始), Cells(表行終, 表列終)).Value = セル '編集した配列"セル"をシートに貼り付け
    
    
End Sub
