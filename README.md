
Sub ProcessAndSaveShipments()

    Dim wsRaw As Worksheet, wsGrouped As Worksheet
    Dim DesktopPath As String, FileName As String
    Dim NewWB As Workbook
    Dim tbl As ListObject
    Dim LastRow As Long
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Set wsRaw = ThisWorkbook.Sheets("Raw Data")
    Set wsGrouped = ThisWorkbook.Sheets("Grouped Gates")
    
    ' === CLEAR OLD RAW DATA (keep headers in row 2) ===
    LastRow = wsRaw.Cells(wsRaw.Rows.Count, "A").End(xlUp).Row
    If LastRow > 2 Then
        wsRaw.Rows("3:" & LastRow).ClearContents
    End If
    
    ' === PASTE THE COPIED DATA STARTING AT A2 (your current layout) ===
    wsRaw.Activate
    wsRaw.Range("A2").Select          ' This is where your headers and data start
    wsRaw.Paste                       ' Pastes whatever is in your clipboard
    
    ' === MAKE SURE IT'S AN EXCEL TABLE (Power Query source) ===
    On Error Resume Next
    Set tbl = wsRaw.ListObjects(1)
    On Error GoTo 0
    
    If tbl Is Nothing Then
        ' Create table if it doesn't exist yet
        LastRow = wsRaw.Cells(wsRaw.Rows.Count, "A").End(xlUp).Row
        If LastRow >= 2 Then
            wsRaw.Range("A2").CurrentRegion.Select
            Set tbl = wsRaw.ListObjects.Add(xlSrcRange, Selection, , xlYes)
            tbl.Name = "RawDataTable"
        End If
    Else
        ' Refresh the table range
        tbl.Resize tbl.Range
    End If
    
    ' === REFRESH THE POWER QUERY ===
    ThisWorkbook.RefreshAll
    
    ' Small delay to let refresh finish
    Application.Wait (Now + TimeValue("0:00:02"))
    
    ' === COPY GROUPED RESULT AS VALUES ONLY TO NEW FILE ===
    wsGrouped.Cells.Copy
    Set NewWB = Workbooks.Add(xlWBATWorksheet)
    
    With NewWB.Sheets(1)
        .Range("A1").PasteSpecial Paste:=xlPasteValues
        .Range("A1").PasteSpecial Paste:=xlPasteFormats
        .Columns.AutoFit
    End With
    
    Application.CutCopyMode = False
    
    ' === SAVE TO DESKTOP WITH DATE + TIME ===
    DesktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
    FileName = "Shipments_" & Format(Now, "yyyymmdd_hhmm") & ".xlsx"
    
    NewWB.SaveAs DesktopPath & FileName, FileFormat:=xlOpenXMLWorkbook
    NewWB.Close False
    
    ' === RESET RAW DATA SHEET (clear below headers) ===
    LastRow = wsRaw.Cells(wsRaw.Rows.Count, "A").End(xlUp).Row
    If LastRow > 2 Then
        wsRaw.Rows("3:" & LastRow).ClearContents
    End If
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "✅ Done!" & vbCrLf & vbCrLf & _
           "Clean file saved to your Desktop as:" & vbCrLf & _
           FileName & vbCrLf & vbCrLf & _
           "Template is ready for the next paste.", vbInformation, "Shipments Processed"

End Sub
