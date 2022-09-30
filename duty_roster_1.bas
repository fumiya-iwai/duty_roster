Attribute VB_Name = "Module1"
Sub make_table()
  
Dim days As New Collection
Dim days_length As Integer
Dim d As Variant
Dim n As Integer
Dim i As Integer
Dim count As Integer

count = 0
For i = 12 To 18
    If Sheets("Sheet1").Cells(i, 2).Value = "" Then
        count = count + 1
    End If
Next

If count = 7 Then
    Err.Raise Number:=888, Description:="日付を入力してください"
End If

For i = 12 To Sheets("Sheet1").Cells(Rows.count, 2).End(xlUp).Row
    If Sheets("Sheet1").Cells(i, 2).Value = "" Then
        Err.Raise Number:=999, Description:="日付は上に詰めて、入力してください"
    Else
        days.Add Sheets("Sheet1").Cells(i, 2).Value
    End If
Next

days_length = days.count

count = 12
For Each d In days
    If IsDate(d) Then
        If Weekday(d) = 3 Then
            n = 4
        ElseIf Weekday(d) = 4 Then
            n = 6
        ElseIf Weekday(d) = 5 Then
            n = 8
        ElseIf Weekday(d) = 6 Then
            n = 10
        ElseIf Weekday(d) = 7 Then
            n = 12
        ElseIf Weekday(d) = 1 Then
            n = 14
        ElseIf Weekday(d) = 2 Then
            n = 16
        End If
        
        Sheets("Sheet1").Cells(count, 4).Value = d
        
        For i = 6 To 10
            If Sheets("Sheet2").Cells(i, n) = "" Then
                Err.Raise Number:=555, Description:="Sheet2で出力する曜日の当番を埋めてください"
            Else
                Sheets("Sheet1").Cells(count, i - 1).Value = Sheets("Sheet2").Cells(i, n).Value
                If Sheets("Sheet2").Cells(i, n + 1) = "x" Then
                    Sheets("Sheet1").Cells(count, i - 1).Value = Sheets("Sheet1").Cells(count, i - 1).Value + "  (リトライ)"
                    Sheets("Sheet2").Cells(i, n + 1) = ""
                    Sheets("Sheet2").Cells(i, n).Font.Color = RGB(0, 0, 0)
                End If
            End If
        Next
    Else
        Err.Raise Number:=666, Description:="oooo/oo/ooの形で日付を入力してください"
    End If
count = count + 1
Next d

Sheets("Sheet1").Cells(9, 7).Value = days(1)
Sheets("Sheet1").Cells(9, 9).Value = days(days_length)

delete_from = days_length + 12
If delete_from < 19 Then
    Sheets("Sheet1").Range(Cells(delete_from, 4), Cells(18, 9)).ClearContents
End If

End Sub

 
 
Sub print_table()
 
Dim last_row As Integer
 
last_row = Sheets("Sheet1").Cells(Rows.count, 4).End(xlUp).Row
 
With Sheets("Sheet1")
    .PageSetup.PrintArea = "D4:I" & last_row
    .PageSetup.Orientation = xlLandscape
    .PageSetup.PaperSize = xlPaperA4
    .PageSetup.Zoom = False
    .PageSetup.FitToPagesWide = 1
    .PageSetup.FitToPagesTall = 1
    .PrintOut
End With
    
End Sub
 


