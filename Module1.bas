Attribute VB_Name = "Module1"
Option Explicit

' ============================================
' resultシート上部に件数サマリーを表示する
' 在庫切れ件数・在庫少件数を一覧の前に表示し、
' 実行結果の全体像をすぐ確認できるようにする
'
' 【役割】
' ① 在庫切れ件数を表示する
' ② 在庫少件数を表示する
' ③ 詳細一覧の前に集計結果を見せる
'
' 【実務上の目的】
' ・結果を報告しやすくする
' ・CSV保存前に件数確認をしやすくする
' ・resultシートを簡易レポートとして使えるようにする
' ============================================

Sub RunInventoryChecK()

    Dim wsStock As Worksheet  '在庫データのシート
    Dim lastRow As Long  '最終行
    Dim i As Long  'ループ用
    Dim stockText As String  '在庫（文字として扱う）
    Dim stockValue As Long  '在庫数を数値として扱う
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
        
    'resultシートの一覧データは５行目から書き込む
    resultRow = 5

   'resultシート上部に件数欄を作る/1～2行目は件数表示用、4行目は一覧の見出しにする
    wsResult.Cells(1, 1).Value = "在庫切れ件数"
    wsResult.Cells(2, 1).Value = "在庫少件数"
    '太文字にする
    wsResult.Cells(1, 1).Font.Bold = True
    wsResult.Cells(2, 1).Font.Bold = True
    
    wsResult.Cells(4, 1).Value = "商品名"
    wsResult.Cells(4, 2).Value = "在庫数"
    wsResult.Cells(4, 3).Value = "判定"
    '太文字にする
    wsResult.Range(wsResult.Cells(4, 1), wsResult.Cells(4, 3)).Font.Bold = True

    msg = ""  'ループの前で初期化

    '２行目からループ
    For i = 2 To lastRow
    
        'C列の在庫を余計な「空白なしの文字」にして取得
        stockText = Trim(CStr(wsStock.Cells(i, 3).Value))  'wsStock.Cells(i, 3).Value...i行目のC列の値（在庫）
                   'Trim(...)前後のスペースを削除/CStr(...)文字に変換
        
        '在庫数が文字か確認し、数字でなければどこの行が原因かを表示して処理を止める
        If stockText = "" Or Not IsNumeric(stockText) Then
            MsgBox "C列の在庫が数字ではありません" & vbCrLf & _
                        "行番号：" & i & vbCrLf & _
                        "商品名：" & wsStock.Cells(i, 1).Value & vbCrLf & _
                        "入力値：" & stockText, vbExclamation
            Exit Sub
        End If
        
        stockValue = CLng(stockText)  '数字チェック後に在庫数を数値へ変換する
                        
        
        '在庫切れチェック
        If stockText = "0" Then
            msg = msg & wsStock.Cells(i, 1).Value & " は在庫切れです" & vbCrLf
            outOfStockCount = outOfStockCount + 1
        End If
        
        '在庫状況の判断結果をD列に書き込む
        If stockText = 0 Then
            wsStock.Cells(i, 4).Value = "在庫切れ"  'D列に書く
        ElseIf stockValue <= 5 Then
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
        ElseIf stockValue <= 5 Then
            resultRow = wsResult.Cells(wsResult.Rows.Count, 1).End(xlUp).Row + 1
            wsResult.Cells(resultRow, 1).Value = wsStock.Cells(i, 1).Value '商品名
            wsResult.Cells(resultRow, 2).Value = stockText  '在庫数
            wsResult.Cells(resultRow, 3).Value = "在庫少"  '判定
        End If
        
        '判定結果に応じて行の色を変える
        If stockText = "0" Then
            wsStock.Range(wsStock.Cells(i, 1), wsStock.Cells(i, 4)).Interior.Color = RGB(255, 199, 206)  '薄い赤
        ElseIf stockValue <= 5 Then
            wsStock.Range(wsStock.Cells(i, 1), wsStock.Cells(i, 4)).Interior.Color = RGB(255, 235, 156)  '薄い黄
        Else
            wsStock.Range(wsStock.Cells(i, 1), wsStock.Cells(i, 4)).Interior.Color = xlNone
            
        End If
        
        '在庫が少ないチェック（１～５）
        If stockValue <= 5 And stockText <> "0" Then
            msg = msg & wsStock.Cells(i, 1).Value & " は在庫が少ないです" & vbCrLf
            lowStocKCount = lowStocKCount + 1
        End If

    Next i
    
    'resultシート上部の件数サマリー欄に集計結果を書き込む/ループで数えた在庫切れ件数・在庫少件数を見える形で残す
    wsResult.Cells(1, 2).Value = outOfStockCount
    wsResult.Cells(2, 2).Value = lowStocKCount

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

' resultシートの列幅を自動調整する/すべてのデータが見切れないようにする
wsResult.Columns("A:C").AutoFit

' resultシートの表に罫線をつける/見出し～データまでを1つの表として見やすくする
Dim lastResultRow As Long

'A列の一番下から上に向かって最後にデータがある行を探す※最後の行にするのはデータ変動に対応するため
lastResultRow = wsResult.Cells(wsResult.Rows.Count, 1).End(xlUp).Row

'４行目から最後の行までA～C列に枠を付ける※最後の行にするのはデータ変動に対応するため
wsResult.Range(wsResult.Cells(4, 1), wsResult.Cells(lastResultRow, 3)).Borders.LineStyle = xlContinuous

'見出し行（4行目）に背景色をつけて視認性を上げる/表のタイトル行として分かりやすくする
wsResult.Range(wsResult.Cells(4, 1), wsResult.Cells(4, 3)).Interior.Color = RGB(200, 200, 200)

'見出し行（4行目）を中央揃えにする/列タイトルとして見やすく整える
wsResult.Range(wsResult.Cells(4, 1), wsResult.Cells(4, 3)).HorizontalAlingnment = xlCenter

'resultシートをコピーしてcsvで保存
wsResult.Copy

'保存確認ダイアログを出さない
Application.DisplayAlerts = False

'今開いているブックをCSVとして保存する
ActiveWorkbook.SaveAs Filename:=csvPath, FileFormat:=xlCSV
'今開いているその一時Excelを閉じる
ActiveWorkbook.Close False

'元に戻す
Application.DisplayAlerts = True

End Sub
