Sub ImportLambdasFromTextFile()

    Dim fileContent() As String, parts() As String
    Dim i As Long, j As Long
    Dim commentStr As String, commentStr2 As String, lambdaName As String, lambdaBody As String
    Dim inLambda As Boolean
    Dim MyName As Name
    
    Dim regex As Object, regex2 As Object
    Dim match As Variant

    ' Create regex object
    Set regex = CreateObject("VBScript.RegExp")
    Set regex2 = CreateObject("VBScript.RegExp")
    ' Set the regex pattern
    regex.Pattern = "^[a-zA-Z0-9._]+\s*\=\s*lambda\("
    regex.Global = False
    regex.MultiLine = True
    regex.IgnoreCase = True
    
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
        If UCase(ws.Name) = "CUSTOM FUNCTIONS" Then
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
                ' Split the string by equal sign
                parts = Split(line, "=")
                lambdaBody = ""
                ' Concatenate all chunks from the second till the last one
                For j = 1 To UBound(parts)
                    lambdaBody = lambdaBody & parts(j)
                Next j
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
                ' Set the regex properties
                regex2.Pattern = "\s"  ' Matches any whitespace character
                regex2.Global = True   ' Replace all occurrences
                ' Use regex to replace all whitespace characters with an empty string
                lambdaName = regex2.Replace(lambdaName, "")
                ' Output the lambda name and body
                ws.Cells(startRow, startCol).Value = lambdaName
                lambdaBody = "= " & lambdaBody
                ws.Cells(startRow, startCol + 1).Value = ("'" & lambdaBody)
                With ActiveWorkbook
                    Set MyName = .Names.Add(Name:=lambdaName, _
                        RefersTo:=Replace(Replace(Replace(Replace(lambdaBody, vbTab, ""), vbCrLf, ""), vbCr, ""), vbLf, ""))  ' Remove Tabs And Line Breaks
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
    ws.Columns("A").columnWidth = 25
    ws.Cells.Font.Name = "Consolas"
    ActiveWindow.Zoom = 80
    ' Activate Excel
    AppActivate Application.Caption
    ' Use SendKeys to simulate the keyboard shortcut for Name Manager (Ctrl + F3)
    SendKeys "^{F3}", True

End Sub

Sub ExportLambdasToTextFile()

    Dim nameItem As Name
    Dim filePath As Variant
    Dim fileNum As Integer
    Dim outputStr As String
    Dim lines() As String
    Dim i As Integer

    ' Ask user for the output text file path
    filePath = Application.GetSaveAsFilename(title:="Please choose a file path", _
                                              FileFilter:="Text Files (*.txt), *.txt")

    ' Exit if user cancels the file dialog
    If filePath = False Then Exit Sub

    Dim wb As Workbook
    Set wb = ActiveWorkbook
    ' Loop through each name in the workbook
    For Each nameItem In wb.Names
        ' Split the string into lines
        lines = Split(nameItem.Comment, vbLf)
        ' Add the # prefix to each line
        For i = LBound(lines) To UBound(lines)
            lines(i) = "#" & lines(i)
        Next i
        ' Check if the name refers to a LAMBDA function
        If InStr(1, nameItem.RefersTo, "LAMBDA(", vbTextCompare) > 0 Then
            outputStr = outputStr & _
                Join(lines, vbCrLf) & vbCrLf & _
                nameItem.Name & " " & nameItem.RefersTo & vbCrLf & vbCrLf
        End If
    Next nameItem

    ' Write the results to the file
    If outputStr <> "" Then
        fileNum = FreeFile
        Open filePath For Output As fileNum
        Print #fileNum, outputStr
        Close fileNum
    Else
        MsgBox "No LAMBDA functions found in the Name Manager."
    End If
    
    Shell "notepad.exe " & filePath, vbNormalFocus

End Sub
