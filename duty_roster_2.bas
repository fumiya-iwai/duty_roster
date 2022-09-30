Attribute VB_Name = "Module2"
Function CreateShuffledCollection(c As Collection) As Collection

' コレクションの要素をランダムに並べ替えて返す
' 引数と返り値はコレクション

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

' 各曜日の当番割り当て表にラスコマ講師一覧から講師を割り当てる

Dim teachers As New Collection
Dim teachers_list As Collection
Dim teacher As Variant
Dim count As Integer
Dim day_of_week As Integer
Dim i As Integer



For day_of_week = 4 To 16 Step 2
    
    count = 0
    For i = 17 To 40
        If Sheets("Sheet2").Cells(i, day_of_week).Value = "" Then
            count = count + 1
        End If
    Next
    
    If count = 24 Then
        For i = 6 To 10
            Sheets("Sheet2").Cells(i, day_of_week).Value = ""
        Next
    Else
        For i = 17 To Sheets("Sheet2").Cells(Rows.count, day_of_week).End(xlUp).Row
            If Sheets("Sheet2").Cells(i, day_of_week).Value = "" Then
                Err.Raise Number:=777, Description:="「上に詰める」を実行してください"
            Else
                teachers.Add Sheets("Sheet2").Cells(i, day_of_week).Value
            End If
        Next
        
        While teachers.count < 5
            For i = 17 To Sheets("Sheet2").Cells(Rows.count, day_of_week).End(xlUp).Row
                teachers.Add Sheets("Sheet2").Cells(i, day_of_week).Value
            Next
        Wend
        
       Set teachers_list = teachers
        
        For i = 6 To 10
            If Sheets("Sheet2").Cells(i, day_of_week + 1).Value = "x" Then
                For Each teacher In teachers_list
                    If teacher = Sheets("Sheet2").Cells(i, day_of_week) Then
                       Sheets("Sheet2").Cells(i, day_of_week).Font.Color = RGB(255, 0, 0)
                       GoTo Continue
                    End If
                Next teacher
            End If
                Set teachers = CreateShuffledCollection(teachers)
                With Sheets("Sheet2")
                    .Cells(i, day_of_week).Value = teachers(1)
                    .Cells(i, day_of_week).Font.Color = RGB(0, 0, 0)
                    .Cells(i, day_of_week + 1).Value = ""
                End With
                teachers.Remove (1)
Continue:
        Next
        
        For i = teachers.count To 1 Step -1
            teachers.Remove (i)
        Next
    End If
Next

End Sub


Sub packing()

' ラスコマ講師一覧の講師名を上に詰める

Dim teachers As New Collection
Dim day_of_week As Integer
Dim count As Integer
Dim i As Integer
Dim n As Integer


For day_of_week = 4 To 16 Step 2
     
    count = 0
    For i = 17 To 40
        If Sheets("Sheet2").Cells(i, day_of_week).Value = "" Then
            count = count + 1
        End If
    Next
    
    If count = 24 Then
        
    Else
        For i = 17 To Sheets("Sheet2").Cells(Rows.count, day_of_week).End(xlUp).Row
            If Sheets("Sheet2").Cells(i, day_of_week).Value = "" Then
                
            Else
                teachers.Add Sheets("Sheet2").Cells(i, day_of_week).Value
                Sheets("Sheet2").Cells(i, day_of_week).Value = ""
            End If
        Next
        
        n = 17
        For Each teacher In teachers
            Sheets("Sheet2").Cells(n, day_of_week).Value = teacher
            n = n + 1
        Next teacher
        For i = teachers.count To 1 Step -1
            teachers.Remove (i)
        Next
    End If
Next

End Sub


