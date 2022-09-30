Attribute VB_Name = "Module2"
Function CreateShuffledCollection(c As Collection) As Collection
    Dim cc As Collection: Set cc = New Collection
    Dim i
    For i = 1 To c.count
        cc.Add c.Item(i)
    Next
    
    Dim ccc As Collection: Set ccc = New Collection
    Dim m As Long
    Do While cc.count > 0
        m = Int(cc.count * Rnd + 1)
        ccc.Add cc.Item(m)
        cc.Remove m
    Loop
    Set CreateShuffledCollection = ccc
End Function

Sub assignment()

Dim teachers As New Collection
Dim teacher As Variant
Dim days_of_week As Integer
Dim i As Integer
Dim j As Integer
Dim n As Integer
Dim count As Integer
Dim teachers_const As Collection


For days_of_week = 4 To 16 Step 2
    
    count = 0
    For i = 17 To 40
        If Sheets("Sheet2").Cells(i, days_of_week).Value = "" Then
            count = count + 1
        End If
    Next
    
    If count = 24 Then
        For j = 6 To 10
            Sheets("Sheet2").Cells(j, days_of_week).Value = ""
        Next
    Else
        For i = 17 To Sheets("Sheet2").Cells(Rows.count, days_of_week).End(xlUp).Row
            If Sheets("Sheet2").Cells(i, days_of_week).Value = "" Then
                Err.Raise Number:=777, Description:="Åuè„Ç…ãlÇﬂÇÈÅvÇé¿çsÇµÇƒÇ≠ÇæÇ≥Ç¢"
            Else
                teachers.Add Sheets("Sheet2").Cells(i, days_of_week).Value
            End If
        Next
        
        While teachers.count < 5
            For i = 17 To Sheets("Sheet2").Cells(Rows.count, days_of_week).End(xlUp).Row
                teachers.Add Sheets("Sheet2").Cells(i, days_of_week).Value
            Next
        Wend
        
       Set teachers_const = teachers
        
        For j = 6 To 10
            If Sheets("Sheet2").Cells(j, days_of_week + 1).Value = "x" Then
                For Each teacher In teachers_const
                    If teacher = Sheets("Sheet2").Cells(j, days_of_week) Then
                       Sheets("Sheet2").Cells(j, days_of_week).Font.Color = RGB(255, 0, 0)
                       GoTo Continue
                    End If
                Next teacher
            End If
                Set teachers = CreateShuffledCollection(teachers)
                Sheets("Sheet2").Cells(j, days_of_week).Value = teachers(1)
                Sheets("Sheet2").Cells(j, days_of_week).Font.Color = RGB(0, 0, 0)
                Sheets("Sheet2").Cells(j, days_of_week + 1).Value = ""
                teachers.Remove (1)
Continue:
        Next
        
        For n = teachers.count To 1 Step -1
            teachers.Remove (n)
        Next
    End If
Next

End Sub


Sub packing()

Dim days_of_week As Integer
Dim i As Integer
Dim count As Integer
Dim teachers As New Collection

For days_of_week = 4 To 16 Step 2
     
    count = 0
    For i = 17 To 40
        If Sheets("Sheet2").Cells(i, days_of_week).Value = "" Then
            count = count + 1
        End If
    Next
    
    If count = 24 Then
        
    Else
        For i = 17 To Sheets("Sheet2").Cells(Rows.count, days_of_week).End(xlUp).Row
            If Sheets("Sheet2").Cells(i, days_of_week).Value = "" Then
                
            Else
                teachers.Add Sheets("Sheet2").Cells(i, days_of_week).Value
                Sheets("Sheet2").Cells(i, days_of_week).Value = ""
            End If
        Next
        
        count = 17
        For Each teacher In teachers
            Sheets("Sheet2").Cells(count, days_of_week).Value = teacher
            count = count + 1
        Next teacher
        For n = teachers.count To 1 Step -1
            teachers.Remove (n)
        Next
    End If
Next

End Sub


