Attribute VB_Name = "Module1"
Sub MakeProgRep()
Attribute MakeProgRep.VB_Description = "Make a progress report for the selected student."
Attribute MakeProgRep.VB_ProcData.VB_Invoke_Func = "R\n14"
'
' MakeProgRep Macro
'

    ' Path & worksheet vars, cat, and student vars - don't touch
    Dim aws As Worksheet, wspr As Worksheet
    Dim PathToPR As String
    Dim studentName As String
    Dim prTitle As String
    Dim cat1Col As Integer
    
    ' This is the number of categories for the progress report
    Const numCat = 5
    Dim Cat(1 To numCat) As String
    
    ' This reads the current sheet into a worksheet object - don't touch
    Set aws = ThisWorkbook.ActiveSheet

    ' This is the path to the progress reports folder - set to your desired path
    PathToPR = "C:\Users\Admin\Documents\Work Small\Comm & Forms\Progress Reports\"
    
    ' The number of columns to display for the first (assignment) category,
    ' As mine can get quite large. Currently supports 1 or 2 columns.
    cat1Col = 2
    
    ' These are all of the assignment category titles
    Cat(1) = "Assignments"
    Cat(2) = "Attendance & Participation"
    Cat(3) = "Tests"
    Cat(4) = "Midterm & Final Exam"
    Cat(5) = "Semester Grade"

    ' Get the active cell's row; column should always be 1, as this macro makes
    ' a progress report based on the student chosen
    With ActiveCell
        Dim r As Integer
        r = .Row
        studentName = .Value
    End With
    
    ' Get the active sheet (class)'s name, which will be reset when the new workbook is made
    With ActiveSheet
        Dim className As String
        Dim numStu As Integer, numStuI As Integer
        
        className = .Name
        numStu = aws.Range("A3", aws.Range("A3").End(xlDown)).Rows.Count
        numStuI = numStu + 3 ' Row incrementer used often
    End With
    
    ' Declare array var for the active sheet's (aws) numer of assignments
    ' in each category
    Dim numAsmt(1 To numCat) As Integer
    
    ' Declare array vars for the active sheet's (aws)
    ' assignment category title row and the assignment title row
    Dim aws_actR(1 To numCat) As Integer, aws_atR(1 To numCat) As Integer
    Dim aws_gR(1 To numCat) As Integer

    'Declare adder vars
    Dim aws_actRa As Integer, aws_atRa As Integer, aws_gRa As Integer
    
    ' Assign start values to the adders for the loop which makes the array
    aws_actRa = 1
    aws_atRa = 2
    aws_gRa = r
    
    ' This loop gathers the row numbers from the source sheet (aws)
    Dim i As Integer

    For i = 1 To numCat
        ' This loads the row number for each assignment category title and assignment title start
        aws_actR(i) = aws_actRa
        aws_atR(i) = aws_atRa
        aws_gR(i) = aws_gRa

        ' Debug.Print i & "|" & aws_actR(i) & "|" & aws_atR(i) & "|"; aws_gR(i)

        ' This gets the number of assignments for each category
        numAsmt(i) = ActiveSheet.Cells(aws_atR(i), Columns.Count).End(xlToLeft).Column

        ' Debug.Print i & "|"; numAsmt(i)

        ' Increment each of the row adder vars
        aws_actRa = aws_actRa + numStuI
        aws_atRa = aws_atRa + numStuI
        aws_gRa = aws_gRa + numStuI

    Next i

    ' Declare/set vars relevant to splitting the columns in the first category
    Dim asmtCol1 As Integer, asmtCol2L As Integer

    asmtCol1 = numAsmt(1) / cat1Col
    asmtCol2L = asmtCol1 + 1
    
    ' Declare vars for destination sheet's (wspr) assignment category title row
    ' and assignment title row. Note there's no separate grade row, because
    ' grades go on the same row as the specific assignment in the destination.
    Dim wspr_actR(1 To numCat) As Integer, wspr_atR(1 To numCat) As Integer

    ' Declare adder vars
    Dim wspr_actRa As Integer, wspr_atRa As Integer
    
    ' Assign start values to the adders for the loop which makes the array
    wspr_actRa = 2
    wspr_atRa = 3
    
    ' This loop calculates the required row positions for
    ' the destination sheet (wspr)
    Dim j As Integer

    For j = 1 To numCat
        ' This loads the row number for each assignment category title and assignment title start
        wspr_actR(j) = wspr_actRa
        wspr_atR(j) = wspr_atRa

        Debug.Print j & "|" & wspr_actR(j) & "|" & wspr_atR(j) _
        & "|"; wspr_actRa & "|" & wspr_atRa

        ' Debug.Print j & "|"; numAsmt(j)

        ' Increment each of the row adder vars
        If j = 1 Then
            If cat1Col = 2 Then
                wspr_actRa = wspr_actRa + Round(numAsmt(j) / 2) + 1
                wspr_atRa = wspr_actRa + 1
            Else
                wspr_actRa = wspr_actRa + numAsmt(j) + 1
                wspr_atRa = wspr_actRa + 1
            End If
        Else
            wspr_actRa = wspr_actRa + numAsmt(j) + 1
            wspr_atRa = wspr_actRa + 1
        End If

    Next j

    ' Make & name the new workbook to hold the progress report
    Set wbpr = Workbooks.Add

    With wbpr
        .Title = "Student Progress Reports " & className
        .Subject = "Progress Reports " & className
        .SaveAs Filename:=PathToPR & "Progress Reports " & className & ".xls"
    End With
    
    ' Create a new sheet with the student's name
    prTitle = "Progress Report " & studentName
    ' Sheets.Add(After:=Sheets(Sheets.Count)).Name = prTitle
    
    Sheets("Sheet1").Name = prTitle
    
    Set wspr = Sheets(prTitle)
    
    ' Add the PR title, first assignment category, & all assignments from the first category
    With Range("A1")
        .Value = prTitle
        .Font.Bold = True
    End With
    
    ' Start the loop that copies and pastes each category from the active sheet to the new workbook
    ' Dim i As Integer (already declared)
    
    For i = 1 To numCat
        ' Set the value and bold the cell for the assignment category title
        ' Debug.Print Cat(i)
        ' The assignment category title is being overwritten by the 1st assignment title
        ' Debug.Print wspr_actR(i)
        With Cells(wspr_actR(i), 1)
            .Value = Cat(i)
            .Font.Bold = True
        End With

        If i = 1 Then
            If cat1Col = 2 Then
                ' Add the assignment titles, transposed from columns to rows (column 1)
                aws.Range(aws.Cells(aws_atR(i), 2), aws.Cells(aws_atR(i), asmtCol1)).Copy
                wspr.Cells(wspr_atR(i), 1).PasteSpecial Transpose:=True
                Application.CutCopyMode = False

                ' Add student's assignment grades from the first category (column 1)
                aws.Range(aws.Cells(aws_gR(i), 2), aws.Cells(aws_gR(i), asmtCol1)).Copy
                wspr.Cells(wspr_atR(i), 2).PasteSpecial Transpose:=True

                ' Add the assignment titles, transposed from columns to rows (column 2)
                aws.Range(aws.Cells(aws_atR(i), asmtCol2L), aws.Cells(aws_atR(i), numAsmt(i))).Copy
                wspr.Cells(wspr_atR(i), 3).PasteSpecial Transpose:=True

                ' Add student's assignment grades from the first category (column 2)
                aws.Range(aws.Cells(aws_gR(i), asmtCol2L), aws.Cells(aws_gR(i), numAsmt(i))).Copy
                wspr.Cells(wspr_atR(i), 4).PasteSpecial Transpose:=True

                Application.CutCopyMode = False
                
            Else
                ' Add the assignment titles, transposed from columns to rows
                aws.Range(aws.Cells(aws_atR(i), 2), aws.Cells(aws_atR(i), numAsmt(i))).Copy
                wspr.Cells(wspr_atR(i), 1).PasteSpecial Transpose:=True
                Application.CutCopyMode = False
        
                ' Add student's assignment grades from the category
                aws.Range(aws.Cells(r, 2), aws.Cells(r, numAsmt(i))).Copy
                wspr.Cells(wspr_atR(i), 2).PasteSpecial Transpose:=True

                Application.CutCopyMode = False
            End If
        Else
            ' Add the assignment titles, transposed from columns to rows
            aws.Range(aws.Cells(aws_atR(i), 2), aws.Cells(aws_atR(i), numAsmt(i))).Copy
            wspr.Cells(wspr_atR(i), 1).PasteSpecial Transpose:=True
            ' Application.CutCopyMode = False

            ' Add student's assignment grades from the category
            aws.Range(aws.Cells(aws_gR(i), 2), aws.Cells(aws_gR(i), numAsmt(i))).Copy
            wspr.Cells(wspr_atR(i), 2).PasteSpecial Transpose:=True
        End If
    Next i

End Sub
Sub Play()
Attribute Play.VB_ProcData.VB_Invoke_Func = "P\n14"
'
' Play Macro
'
    ' Category 2 Grades
    
    
    ' A variable to locate the proper row to get student assignment titles from category 2
'    Dim cat2R As Integer
'    cat2R = numStu + 5
'
'    ' The row locator for a student's grades from category 2
'
'    With Range(Cells(1, cat2R))
'        .Value = "Attendance & Participation"
'        .Font.Bold = True
'    End With

'    Dim PathToPR As String
'
'    ' This is the path to the progress reports folder
'    PathToPR = "C:\Users\james\Documents\Work Mini\Comm\Progress Reports\"
'
'    With ActiveSheet
'        Dim actShtName As String
'        actShtName = .Name
'        Debug.Print actShtName
'    End With
'
'    Set NewBook = Workbooks.Add
'
'    With NewBook
'        .Title = "Student Progress Reports " & actShtName
'        .Subject = "Progress Reports " & actShtName
'        .SaveAs Filename:=PathToPR & "Progress Reports " & actShtName & ".xls"
'    End With

'    Dim ws As Worksheet
'
'    Set ws = ThisWorkbook.ActiveSheet
'
'    Debug.Print ws.Name

End Sub
Sub DelPR()
Attribute DelPR.VB_Description = "Deletes a progress report."
Attribute DelPR.VB_ProcData.VB_Invoke_Func = "D\n14"
'
' DelPR Macro
' Deletes a progress report sheet

Sheets("Progress Report Apple").Delete

End Sub
Sub Play2()
'
Const numCat = 5

Dim numStu As Integer, numStuI

numStu = 12
numStuI = numStu + 3

' Declare array vars and the adders for the active sheet's (aws)
' assignment category title row and the assignment title row
Dim aws_atR(1 To numCat) As Integer, aws_actR(1 To numCat) As Integer
Dim aws_atRa As Integer, aws_actRa As Integer

' Assign start values to the adders for the loop which makes the array
aws_actRa = 1
aws_atRa = 2

Dim i As Integer

For i = 1 To numCat
    ' This loads the row number for each assignment category title and assignment title start
    aws_actR(i) = aws_actRa
    aws_atR(i) = aws_atRa
    aws_actRa = aws_actRa + numStuI
    aws_atRa = aws_atRa + numStuI
Next i

End Sub
Sub Play3()

    With ActiveSheet
        Dim numAsmt As Integer

        numAsmt = ActiveSheet.Range("A1").Cells(1, Columns.Count).End(xlToLeft).Column - 1
        Debug.Print numAsmt
    End With

End Sub
