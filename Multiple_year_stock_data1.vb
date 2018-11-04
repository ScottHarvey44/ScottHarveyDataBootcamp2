Sub Stock_Volume()

Dim Ticker As String
Dim Unique As Integer
Dim LastRow As Long

For Each ws In Worksheets

Unique = 2
LastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
ActiveSheet.Cells(1, 8).Value = "Ticker"
ActiveSheet.Cells(1, 9).Value = "Total Stock Volume"

For i = 2 To LastRow

If ActiveSheet.Cells(i + 1, 1).Value <> ActiveSheet.Cells(i, 1).Value Then

Ticker = ActiveSheet.Cells(i, 1)
ActiveSheet.Cells(Unique, 8).Value = Ticker
ActiveSheet.Cells(Unique, 9).Value = ActiveSheet.Cells(Unique, 9).Value + Cells(i, 7).Value
Unique = Unique + 1

ElseIf ActiveSheet.Cells(i + 1, 1).Value = ActiveSheet.Cells(i, 1).Value Then
ActiveSheet.Cells(Unique, 9).Value = ActiveSheet.Cells(Unique, 9).Value + Cells(i, 7).Value

End If

Next i

If ActiveSheet.Index = Worksheets.Count Then
Worksheets(1).Select
Else
ActiveSheet.Next.Select
End If

Next ws

End Sub


