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
    Dim outOfStockCount As Long  '在庫切れ件数
    Dim lowStocKCount As Long  '在所少件数
    Dim wsResult As Worksheet  '結果出力用のシート
    Dim resultRow As Long  'resultシートの書き込み行

    'シート設定
    Set wsStock = ThisWorkbook.Worksheets("stock")
    
    '最終行取得（A列基準）
    lastRow = wsStock.Cells(wsStock.Rows.Count, 1).End(xlUp).Row
    
    'resultシートを取得し、なければ新しく作る
    On Error Resume Next
    Set wsResult = ThisWorkbook.Worksheets("result")
    On Error GoTo 0
    
    If wsResult Is Nothing Then
        Set wsResult = ThisWorkbook.Worksheets.Add
        wsResult.Name = "result"
    End If
    
    'resultシートを初期化する
    wsResult.Cells.Clear
    
    '見出しを入れる
    wsResult.Cells(1, 1).Value = "商品名"
    wsResult.Cells(1, 2).Value = "在庫数"
    wsResult.Cells(1, 3).Value = "判定"
    
     
    
    
    msg = ""  'ループの前で初期化

    '２行目からループ
    For i = 2 To lastRow

        'C列の在庫を余計な「空白なしの文字」にして取得
        stockText = Trim(CStr(wsStock.Cells(i, 3).Value))  'wsStock.Cells(i, 3).Value...i行目のC列の値（在庫）
                   'Trim(...)前後のスペースを削除/CStr(...)文字に変換
        
        '在庫切れチェック
        If stockText = "0" Then
            msg = msg & wsStock.Cells(i, 1).Value & " は在庫切れです" & vbCrLf
            outOfStockCount = outOfStockCount + 1
        End If
        
        '在庫状況の判断結果をD列に書き込む
        If stockText = 0 Then
            wsStock.Cells(i, 4).Value = "在庫切れ"  'D列に書く
        ElseIf Val(stockText) <= 5 Then
            wsStock.Cells(i, 4).Value = "在庫少"  'D列に在庫少と書く
        Else
            wsStock.Cells(i, 4).Value = "正常"  'D列に正常と書く
        End If
        
        '在庫切れと在庫少の商品をresultシートに書き出す
        If stockText = "0" Then
            resultRow = wsResult.Cells(wsResult.Rows.Count, 1).End(xlUp).Row + 1
            wsResult.Cells(resultRow, 1).Value = wsStock.Cells(i, 1).Value  '商品名
            wsResult.Cells(resultRow, 2).Value = stockText  '在庫数
            wsResult.Cells(resultRow, 3).Value = "在庫切れ"  '判定
        ElseIf Val(stockText) <= 5 Then
            resultRow = wsResult.Cells(wsResult.Rows.Count, 1).End(xlUp).Row + 1
            wsResult.Cells(resultRow, 1).Value = wsStock.Cells(i, 1).Value '商品名
            wsResult.Cells(resultRow, 2).Value = stockText  '在庫数
            wsResult.Cells(resultRow, 3).Value = "在庫切れ"  '判定
        End If
        
        
        
        '判定結果に応じて行の色を変える
        If stockText = "0" Then
            wsStock.Range(wsStock.Cells(i, 1), wsStock.Cells(i, 4)).Interior.Color = RGB(255, 199, 206)  '薄い赤
        ElseIf Val(stockText) <= 5 Then
            wsStock.Range(wsStock.Cells(i, 1), wsStock.Cells(i, 4)).Interior.Color = RGB(255, 235, 156)  '薄い黄
        Else
            wsStock.Range(wsStock.Cells(i, 1), wsStock.Cells(i, 4)).Interior.Color = xlNone
            
        End If
        
        
        '在庫が少ないチェック（１～５）
        If Val(stockText) <= 5 And stockText <> "0" Then
            msg = msg & wsStock.Cells(i, 1).Value & " は在庫が少ないです" & vbCrLf
            lowStocKCount = lowStocKCount + 1
        End If

    Next i

'ループの後に１回だけ表示
If msg <> "" Then
MsgBox msg & vbCrLf & _
       "在庫切れ：" & outOfStockCount & "件" & vbCrLf & _
       "在庫少：" & lowStocKCount & "件"
End If

'resultシートをcvsとして保存する
Dim csvPath As String

'保存先（このExcelと同じフォルダ）
csvPath = ThisWorkbook.Path & "\result.csv"

'resultシートをコピーしてcsvで保存
wsResult.Copy
ActiveWorkbook.SaveAs Filename:=csvPath, FileFormat:=xlCSV
ActiveWorkbook.Close False
End Sub
