Sub ImportLambdaFunctions()

    Dim fileContent() As String
    Dim i As Long
    Dim commentStr As String, commentStr2 As String, lambdaName As String, lambdaBody As String
    Dim inLambda As Boolean
    Dim MyName As Name
    
    Dim regex As Object
    Dim match As Variant

    ' Create regex object
    Set regex = CreateObject("VBScript.RegExp")
    ' Set the regex pattern
    regex.Pattern = "^[a-zA-Z0-9._]+\s*\=\s*lambda\("
    regex.Global = False
    regex.MultiLine = True
    
    Dim filePath As Variant

    filePath = Excel.Application.GetOpenFilename("Text Files (*.txt), *.txt, All Files (*.*), *.*")
    ' Check if user clicked Cancel
    If filePath = False Then
        Exit Sub
    End If
    
    ' Read the file into an array
    fileContent = Split(CreateObject("Scripting.FileSystemObject").OpenTextFile(filePath, 1).ReadAll, vbCrLf)
    

    Dim wb As Workbook
    Set wb = ActiveWorkbook
    Dim ws As Worksheet
    ' Delete the worksheet if it exists
    For Each ws In wb.Worksheets
        If ws.Name = "Custom Functions" Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
            Exit For
        End If
    Next ws
    ' Create a new worksheet at the end of the workbook
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "Custom Functions"
    Set ws = wb.ActiveSheet
    ' Start from the first cell in the active sheet
    Dim startRow As Long, startCol As Long
    startRow = 1
    startCol = 1
    
    For i = LBound(fileContent) To UBound(fileContent)
        Dim line As String
        line = Trim(fileContent(i))
        
        ' Ignore lines with ##
        If Not Left(line, 2) = "##" Then
            ' If line starts with #, it's a comment
            If Left(line, 1) = "#" And Not inLambda Then
                If commentStr <> "" Then
                    commentStr = commentStr & vbLf
                End If
                commentStr = commentStr & Mid(line, 2)
            ' If the line includes a lambda definition start
            ElseIf regex.Test(line) Then
                ' Write comment to Excel
                ws.Cells(startRow, startCol).Value = commentStr
                ws.Cells(startRow, startCol).Font.Bold = True
                commentStr2 = commentStr
                startRow = startRow + 1
                commentStr = ""  ' Reset commentStr for next iteration
                
                ' Capture the lambda name
                lambdaName = Trim(Split(line, "=")(0))
                
                ' Start capturing the lambda body
                lambdaBody = Trim(Split(line, "=")(1))
                inLambda = True
            ' If we're inside a lambda, continue capturing its body
            ElseIf inLambda Then
                lambdaBody = lambdaBody & vbLf & line
            End If
        End If
        
        Dim isLastLine As Boolean
        Dim isNextLineEmptyOrComment As Boolean
        
        isLastLine = (i = UBound(fileContent))
        If Not isLastLine Then
            isNextLineEmptyOrComment = (fileContent(i + 1) = "" Or Left(fileContent(i + 1), 1) = "#")
        End If
        
        If isLastLine Or isNextLineEmptyOrComment Then
            If inLambda Then
                ' Output the lambda name and body
                ws.Cells(startRow, startCol).Value = lambdaName
                lambdaBody = "= " & lambdaBody
                ws.Cells(startRow, startCol + 1).Value = ("'" & lambdaBody)
                With ActiveWorkbook
                    Set MyName = .Names.Add(Name:=lambdaName, _
                        RefersTo:=Replace(Replace(Replace(Replace(lambdaBody, vbTab, ""), vbCrLf, ""), vbCr, ""), vbLf, "")) ' Remove Tabs And Line Breaks
                    With MyName
                        .Comment = commentStr2
                    End With
                End With
                ' Move the start position
                startRow = startRow + 1
                startCol = 1
                
                ' Reset the lambda variables for next iteration
                lambdaName = ""
                lambdaBody = ""
                inLambda = False
            End If
        End If
    Next i
    
    ActiveSheet.Cells.WrapText = False
    ' Activate Excel
    AppActivate Application.Caption
    ' Use SendKeys to simulate the keyboard shortcut for Name Manager (Ctrl + F3)
    SendKeys "^{F3}", True

End Sub

