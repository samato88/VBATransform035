' written via ChatGPT 4o - Jan 1, 2025
Sub Transform035Field()
    Dim ws As Worksheet
    Dim targetColumn As Range
    Dim newColIndex As Long
    Dim cell As Range
    Dim inputText As String
    Dim extractedText As String
    Dim validPrefixes As Variant
    Dim i As Long
    Dim headerCell As Range
    Dim maxLength As Integer
    Dim parts As Variant
    Dim part As Variant
    Dim results As String
    Dim uniqueValues As Collection

    ' Set the worksheet (assumes the macro runs on the active sheet)
    Set ws = ActiveSheet

    ' Define valid prefixes
    validPrefixes = Array("ocm", "ocn", "on", "(OCoLC)ocm", "(OCoLC)ocn", "(OCoLC)on", "(OCoLC)")

    ' Find the column labeled "035 field"
    On Error Resume Next
    Set targetColumn = ws.Rows(1).Find(What:="035 field", LookIn:=xlValues, LookAt:=xlWhole)
    On Error GoTo 0

    ' If the column is found, proceed
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

        ' Initialize maxLength to track the longest entry
        maxLength = Len(ws.Cells(1, newColIndex).Value)

        ' Loop through the cells in the "035 field" column (assuming data starts in row 2)
        For Each cell In ws.Range(ws.Cells(2, targetColumn.Column), ws.Cells(ws.Rows.Count, targetColumn.Column).End(xlUp))
            inputText = cell.Value
            results = ""

            ' Split the input text by "$"
            parts = Split(inputText, "$")

            ' Initialize collection for unique values
            Set uniqueValues = New Collection

            ' Loop through each part to find valid "$a" entries
            For Each part In parts
                If Left(Trim(part), 1) = "a" Then
                    extractedText = Mid(Trim(part), 2)
                    
                    ' Check if extracted text starts with any valid prefix
                    For i = LBound(validPrefixes) To UBound(validPrefixes)
                        If Left(Trim(extractedText), Len(validPrefixes(i))) = validPrefixes(i) Then
                            ' Remove the full matched prefix before copying
                            extractedText = Mid(Trim(extractedText), Len(validPrefixes(i)) + 1)
                            On Error Resume Next
                            uniqueValues.Add Trim(extractedText), Trim(extractedText)
                            On Error GoTo 0
                            Exit For
                        End If
                    Next i
                End If
            Next part

            ' Concatenate unique values with a semicolon delimiter
            If uniqueValues.Count > 0 Then
                Dim val As Variant
                For Each val In uniqueValues
                    If results <> "" Then
                        results = results & "; "
                    End If
                    results = results & val
                Next val
            End If

            ' Set the results in the new column
            ws.Cells(cell.Row, newColIndex).Value = results

            ' Update maxLength if necessary
            If Len(results) > maxLength Then
                maxLength = Len(results)
            End If
        Next cell

        ' Adjust the column width to fit the longest entry
        ws.Columns(newColIndex).ColumnWidth = maxLength + 2

        MsgBox "Transformation complete!", vbInformation
    Else
        MsgBox "Column labeled '035 field' not found.", vbExclamation
    End If
End Sub
 