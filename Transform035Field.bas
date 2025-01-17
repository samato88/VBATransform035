' written via ChatGPT 4o - Jan 1, 2025
Sub Transform035Field()
    Dim ws As Worksheet
    Dim targetColumn As Range
    Dim newColIndex As Long
    Dim headerCell As Range
    Dim validPrefixes As Variant
    Dim dataRange As Range
    Dim inputArray As Variant
    Dim outputArray() As String
    Dim parts As Variant
    Dim part As Variant
    Dim extractedText As String
    Dim uniqueValues As Collection
    Dim i As Long, rowCount As Long
    Dim maxLength As Integer
    Dim chunkSize As Long
    Dim startRow As Long, endRow As Long
    Dim totalRows As Long

    ' Set the worksheet (assumes the macro runs on the active sheet)
    Set ws = ActiveSheet

    ' Define valid prefixes
    validPrefixes = Array("ocm", "ocn", "on", "(OCoLC)ocm", "(OCoLC)ocn", "(OCoLC)on", "(OCoLC)")

    ' Find the column labeled "035 field"
    On Error Resume Next
    Set targetColumn = ws.Rows(1).Find(What:="035 field", LookIn:=xlValues, LookAt:=xlWhole)
    On Error GoTo 0

    If Not targetColumn Is Nothing Then
        ' Check if the column "Extracted OCLC Number" already exists
        On Error Resume Next
        Set headerCell = ws.Rows(1).Find(What:="Extracted OCLC Number", LookIn:=xlValues, LookAt:=xlWhole)
        On Error GoTo 0

        If Not headerCell Is Nothing Then
            newColIndex = headerCell.Column
        Else
            ' Insert a new column to the right of "035 field"
            newColIndex = targetColumn.Column + 1
            ws.Columns(newColIndex).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        End If

        ' Add or overwrite the header for the new column
        ws.Cells(1, newColIndex).Value = "Extracted OCLC Number"

        ' Set the new column format to Text
        ws.Columns(newColIndex).NumberFormat = "@"

        ' Determine the total number of rows
        totalRows = ws.Cells(ws.Rows.Count, targetColumn.Column).End(xlUp).Row
        chunkSize = 50000 ' Process rows in chunks of 50,000

        ' Process rows in chunks
        For startRow = 2 To totalRows Step chunkSize
            endRow = Application.Min(startRow + chunkSize - 1, totalRows)

            ' Read the data into an array for faster processing
            Set dataRange = ws.Range(ws.Cells(startRow, targetColumn.Column), ws.Cells(endRow, targetColumn.Column))
            inputArray = dataRange.Value
            rowCount = dataRange.Rows.Count
            ReDim outputArray(1 To rowCount, 1 To 1)

            ' Process each row in the chunk
            For i = 1 To rowCount
                Dim results As String
                results = ""
                Dim val As Variant

                ' Initialize unique values collection
                Set uniqueValues = New Collection

                ' Split the input text by "$"
                If Not IsEmpty(inputArray(i, 1)) Then
                    parts = Split(inputArray(i, 1), "$")
                    For Each part In parts
                        If Left(Trim(part), 1) = "a" Then
                            extractedText = Mid(Trim(part), 2)

                            ' Check for valid prefixes
                            Dim prefix As Variant
                            For Each prefix In validPrefixes
                                If Left(Trim(extractedText), Len(prefix)) = prefix Then
                                    ' Remove the prefix
                                    extractedText = Mid(Trim(extractedText), Len(prefix) + 1)

                                    ' Validate that the extracted text is numeric
                                    If IsNumeric(extractedText) Then
                                        On Error Resume Next
                                        ' Remove leading zeros safely
                                        extractedText = CStr(CLng(extractedText))
                                        On Error GoTo 0
                                    Else
                                        ' Skip non-numeric entries
                                        extractedText = ""
                                    End If

                                    ' Add to unique collection if valid
                                    If Len(extractedText) > 0 Then
                                        On Error Resume Next
                                        uniqueValues.Add extractedText, extractedText
                                        On Error GoTo 0
                                    End If

                                    Exit For
                                End If
                            Next prefix
                        End If
                    Next part
                End If

                ' Concatenate unique values with a semicolon delimiter
                If uniqueValues.Count > 0 Then
                    For Each val In uniqueValues
                        If results <> "" Then results = results & "; "
                        results = results & val
                    Next val
                End If

                outputArray(i, 1) = results
            Next i

            ' Write the processed data back to the worksheet
            ws.Range(ws.Cells(startRow, newColIndex), ws.Cells(endRow, newColIndex)).Value = outputArray
        Next startRow

        ' Auto-adjust column width
        ws.Columns(newColIndex).AutoFit
        MsgBox "Transformation complete!", vbInformation
    Else
        MsgBox "Column labeled '035 field' not found.", vbExclamation
    End If
End Sub