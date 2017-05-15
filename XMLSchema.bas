'Macro to create XML schema based on current file
'Finds blanks, formats dates to company spec, converts to filetypes

Attribute VB_Name = "XMLSchema"

Sub autoSchema()

Dim x As Integer
Dim origWS As Worksheet

    Set origWS = ActiveSheet

'Delete blank columns in data
Call deleteBlankColumns

'Scans header row for blanks
    Range("A1").Select
    Do Until IsEmpty(ActiveCell) And IsEmpty(ActiveCell.Offset(1, 0))
        If IsEmpty(ActiveCell) Then
            MsgBox "                   A header cell is empty." & vbCr & _
                    "Please fill in all header cells and try again."
                    End
        Else
            ActiveCell.Offset(0, 1).Select
        End If
    Loop
        
'Header formatting - replaces spaces for underscore
Call replaceSpaceTopRow
'Format date ranges - based on "date" in header name
Call findDate


    Range("A1").Select
    ActiveCell.CurrentRegion.Select
    Selection.Copy
    
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, transpose:=False
    ActiveSheet.Name = ActiveSheet.Name & "Schema Build"
      
'Transposes
    Rows("3:" & Rows.Count).ClearContents
    ActiveCell.CurrentRegion.Copy
    Range("A3").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, transpose:=True
    Rows("1:2").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    
'Changes cell text to data type
    Range("B1").Select
    Do Until IsEmpty(ActiveCell.Offset(0, -1))
      If InStr(1, ActiveCell.Offset(0, -1), "date", 1) Then
          Selection.Value = "1999/12/31"
      ElseIf IsEmpty(Selection) Then
          Selection.Value = "ABC"
      ElseIf IsNumeric(Selection) Then
          Selection.Value = "123"
      Else
          Selection.Value = "ABC"
      End If
       ActiveCell.Offset(1, 0).Select
    Loop
      
'Adds relevent text to file
    Range("C1").Select
    Do Until IsEmpty(ActiveCell.Offset(, -2))
        ActiveCell.Value = "<" & ActiveCell.Offset(0, -2) & ">" _
                        & ActiveCell.Offset(0, -1) _
                        & "</" & ActiveCell.Offset(0, -2) & ">"
        ActiveCell.Offset(1, 0).Select
    Loop
    
'Final format for SCHEMA export
    Columns("A:B").EntireColumn.Delete
    Range("A1").Insert Shift:=xlDown
    Range("A1").Insert Shift:=xlDown
    Range("A1").Insert Shift:=xlDown
    Range("A1").Value = "<?xml version='1.0'?>"
    ' Range("A2").Value = "<BookInfo>"
    ' Range("A3").Value = "<Book>"
    Range("A2").Value = "<RecordInfo>"
    Range("A3").Value = "<Record>"
    Range("A1").End(xlDown).Offset(1, 0).Select
    'Selection.Value = "</Book>"
    Selection.Value = "</Record>"
    
    Range("A3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A1").End(xlDown).Offset(1, 0).Select
    ActiveSheet.Paste
    
   Range("A1").End(xlDown).Offset(1, 0).Select
   ' Selection.Value = "</BookInfo>"
   Selection.Value = "</RecordInfo>"

Call exportSchema

    Application.DisplayAlerts = False
    ActiveSheet.Delete
    Application.DisplayAlerts = True
    
    origWS.Activate
    
    FName = Application.GetOpenFilename
    ' ActiveWorkbook.XmlMaps.Add(FName, _
        ' "BookInfo").Name = "BookInfo_Map"
    ActiveWorkbook.XmlMaps.Add(FName, _
        "RecordInfo").Name = "RecordInfo_Map"
        

End Sub

Sub exportSchema()
    Dim FileName As Variant
    Dim Sep As String
    FileName = Application.GetSaveAsFilename(InitialFileName:=vbNullString)
    If FileName = False Then
        ' user cancelled, get out
        Exit Sub
    End If
    Sep = "^l"
    Debug.Print "FileName: " & FileName, "Separator: " & Sep
    exportToTextFile FName:=CStr(FileName), Sep:=CStr(Sep), _
       SelectionOnly:=False, AppendData:=False
End Sub


Sub findDate()

Dim lastRow As Long

lastRow = Cells(Cells.Rows.Count, "A").End(xlUp).Row

'Changes cell text to data type
    Range("A1").Select
    Do Until IsEmpty(ActiveCell) And IsEmpty(ActiveCell.Offset(1, 0))
      If InStr(1, Selection, "date", 1) Then
        Selection.Offset(0, 1).EntireColumn.Insert
        ActiveCell.Offset(0, 1).Select
        ActiveCell.FormulaR1C1 = "=RC[-1]"
        ActiveCell.Offset(1, 0).Select
        Selection.FormulaR1C1 = _
            "=TEXT(RC[-1], ""mm/dd/yyyy"")"
        ActiveCell.Copy ActiveCell.Resize(lastRow)
        ActiveCell.EntireColumn.Copy
        ActiveCell.EntireColumn.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, transpose:=False
        
        Columns(ActiveCell.Column - 1).Delete
        
        Range(ActiveCell.EntireColumn.Address)(1, 0).Select
        
        Do Until IsEmpty(ActiveCell)
            If ActiveCell.Value = "01/00/1900" Then
                ActiveCell.ClearContents
                ActiveCell.Offset(1, 0).Select
            Else
                ActiveCell.Offset(1, 0).Select
            End If
        Loop
        
        Range(ActiveCell.EntireColumn.Address)(1, 2).Select
        
      Else
       ActiveCell.Offset(0, 1).Select
      End If
    Loop

End Sub


Sub replaceSpaceTopRow()

    Dim startCol As String
    Dim startRow As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim myCol As String
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    
    Set ws = ActiveSheet
    startCol = "A"
    startRow = 1
    lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    myCol = GetColumnLetter(lastCol)
    
    Set rng = ws.Range(startCol & startRow & ":" & myCol & startRow)
    
    For Each cell In rng
        cell = Replace(cell, " ", "_")
    Next cell
    
End Sub


Function GetColumnLetter(colNum As Long) As String
    Dim vArr
    vArr = Split(Cells(1, colNum).Address(True, False), "$")
    GetColumnLetter = vArr(0)
End Function


Sub deleteBlankColumns()
'Step1:  Declare your variables.
    Dim MyRange As Range
    Dim iCounter As Long
'Step 2:  Define the target Range.
    Set MyRange = ActiveSheet.UsedRange
    
'Step 3:  Start reverse looping through the range.
    For iCounter = MyRange.Columns.Count To 1 Step -1
    
'Step 4: If entire column is empty then delete it.
       If Application.CountA(Columns(iCounter).EntireColumn) = 0 Then
       Columns(iCounter).Delete
       End If
'Step 5: Increment the counter down
    Next iCounter
End Sub
