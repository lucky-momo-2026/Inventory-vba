Attribute VB_Name = "Module1"
Option Explicit

Sub RunInventoryChecK()

    Dim wsStock As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim stockText As String

    MsgBox "Ç±ÇÃÉ}ÉNÉçÇÕìÆÇ¢ÇƒÇ¢Ç‹Ç∑"

    Set wsStock = ThisWorkbook.Worksheets("stock")
    lastRow = wsStock.Cells(wsStock.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow

        stockText = Trim(CStr(wsStock.Cells(i, 3).Value))

        If stockText = "0" Then
            MsgBox wsStock.Cells(i, 1).Value & " ÇÕç›å…êÿÇÍÇ≈Ç∑"
        End If

    Next i

End Sub
