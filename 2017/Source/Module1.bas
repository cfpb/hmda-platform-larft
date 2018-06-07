Attribute VB_Name = "Module1"
Private Sub Export()

Dim export_string As String
Dim new_lines As Integer
Dim lar_lines As Integer
Dim path As String


Worksheets("Export").Columns(1).ClearContents

'Generate the transmittal sheet
export_string = "1"
For i = 1 To 20 '1 to 20 are the columns the ts are contained in (ts is 20 fields)
    export_string = export_string + "|" + CStr(Cells(3, i))
Next i

'Write the transmittal sheet into "Export" sheet
new_lines = 1
Worksheets("Export").Cells(new_lines, 1).Value = export_string
new_lines = new_lines + 1

i = 5 '5 is the first row that lars are contained on
'Generate and write lars
Do While Cells(i, 1).Value <> ""
    export_string = "2"
    For j = 1 To 38 '1 to 38 are the columns the lars are contained in (a lar is 38 fields)
        export_string = export_string + "|" + CStr(Cells(i, j))
    Next j
    Worksheets("Export").Cells(new_lines, 1).Value = export_string
    new_lines = new_lines + 1
    i = i + 1
Loop

path = Application.GetSaveAsFilename( _
    FileFilter:="Text Files (*.txt), *.txt")

If path <> "False" Then
    Open path For Output Access Write As #1 'creates and opens the document

    i = 1
    Do While Worksheets("Export").Cells(i, 1).Value <> ""
        Print #1, Sheets("Export").Cells(i, 1).Value 'writes in the data line by line
        i = i + 1
    Loop

    Close #1 'closes the document
End If
    
End Sub
