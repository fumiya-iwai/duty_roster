Attribute VB_Name = "Module1"
Sub make_table()

' 教室業務当番表を作る

Dim days As New Collection
Dim day As Variant
Dim day_of_week As Integer
Dim i As Integer
Dim n As Integer
Dim count As Integer
Dim delete_from As Integer

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

n = 12
For Each day In days
    If IsDate(day) Then
        Select Case Weekday(day)
            Case 3
                day_of_week = 4
            Case 4
                day_of_week = 6
            Case 5
                day_of_week = 8
            Case 6
                day_of_week = 10
            Case 7
                day_of_week = 12
            Case 1
                day_of_week = 14
            Case 2
                day_of_week = 16
        End Select
        
        Sheets("Sheet1").Cells(n, 4).Value = day
        
        For i = 6 To 10
            If Sheets("Sheet2").Cells(i, day_of_week) = "" Then
                Err.Raise Number:=555, Description:="Sheet2で出力する曜日の当番を埋めてください"
            Else
                Sheets("Sheet1").Cells(n, i - 1).Value = Sheets("Sheet2").Cells(i, day_of_week).Value
                If Sheets("Sheet2").Cells(i, day_of_week + 1) = "x" Then
                    Sheets("Sheet1").Cells(n, i - 1).Value = Sheets("Sheet1").Cells(n, i - 1).Value + "  (リトライ)"
                    Sheets("Sheet2").Cells(i, day_of_week + 1) = ""
                    Sheets("Sheet2").Cells(i, day_of_week).Font.Color = RGB(0, 0, 0)
                End If
            End If
        Next
    Else
        Err.Raise Number:=666, Description:="oooo/oo/ooの形で日付を入力してください"
    End If
n = n + 1
Next day

Sheets("Sheet1").Cells(9, 7).Value = days(1)
Sheets("Sheet1").Cells(9, 9).Value = days(days.count)

delete_from = days.count + 12
If delete_from < 19 Then
    Sheets("Sheet1").Range(Cells(delete_from, 4), Cells(18, 9)).ClearContents
End If

End Sub

 
 
Sub print_table()
 
' 教室業務当番表を印刷する
 
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
 


