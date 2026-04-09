Attribute VB_Name = "Module1"
Option Explicit

' ============================================
' 在庫管理チェック処理
' 在庫切れ（0）と在庫が少ない商品（5以下）を判定
' ============================================

Sub RunInventoryChecK()

    Dim wsStock As Worksheet  '在庫データのシート
    Dim lastRow As Long  '最終行
    Dim i As Long  'ループ用
    Dim stockText As String  '在庫（文字として扱う）
    Dim msg As String  '表示のメッセージをまとめる

    'シート設定
    Set wsStock = ThisWorkbook.Worksheets("stock")
    
    '最終行取得（A列基準）
    lastRow = wsStock.Cells(wsStock.Rows.Count, 1).End(xlUp).Row
    
    msg = ""  'ループの前で初期化

    '２行目からループ
    For i = 2 To lastRow

        'C列の在庫を余計な「空白なしの文字」にして取得
        stockText = Trim(CStr(wsStock.Cells(i, 3).Value))  'wsStock.Cells(i, 3).Value...i行目のC列の値（在庫）
                   'Trim(...)前後のスペースを削除/CStr(...)文字に変換
        
        '在庫切れチェック
        If stockText = "0" Then
            msg = msg & wsStock.Cells(i, 1).Value & " は在庫切れです" & vbCrLf
        End If
        
        
        '在庫状況の判断結果をD列に書き込む
        If stockText = 0 Then
            wsStock.Cells(i, 4).Value = "在庫切れ"  'D列に書く
        ElseIf Val(stockText) <= 5 Then
            wsStock.Cells(i, 4).Value = "在庫少"  'D列に在庫少と書く
        Else
            wsStock.Cells(i, 4).Value = "正常"  'D列に正常と書く
        End If
        
        '在庫が少ないチェック（１～５）
        If Val(stockText) <= 5 And stockText <> "0" Then
            msg = msg & wsStock.Cells(i, 1).Value & " は在庫が少ないです" & vbCrLf
        End If

    Next i

'ループの後に１回だけ表示
If msg <> "" Then
    MsgBox msg
End If

End Sub
