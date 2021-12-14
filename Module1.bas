Attribute VB_Name = "Module1"
Dim ClosingMessageByOtherSub As Boolean

Sub LoadMCCs()

    Dim ExcelFiles As FileDialog
    Dim vrtSelectedFiles As Variant
    Dim ProcessedFiles As Long
        
    ClosingMessageByOtherSub = True
    
    Set ExcelFiles = Application.FileDialog(msoFileDialogFilePicker)
    ExcelFiles.AllowMultiSelect = True
    ExcelFiles.Filters.Clear
    ExcelFiles.Filters.Add "Excel Files", "*.xlsx", 1
    ExcelFiles.Title = "Seleccione el/los ficheros a cargar"
    If (ExcelFiles.Show = -1) Then
        For Each vrtSelectedFiles In ExcelFiles.SelectedItems
            ProcessedFiles = ProcessedFiles + HandleOneExcelFile(vrtSelectedFiles)
        Next vrtSelectedFiles
    End If
   
    If ProcessedFiles > 0 Then
        ShowMessage (ProcessedFiles & " files have been processed")
        Call OrderSheets
        
        Call UpdateSummary
    Else
        ShowMessage ("No files have been processed")
    End If
    
    ShowMessage ("Finished - window can be closed")
    ClosingMessageByOtherSub = False
End Sub

Function HandleOneExcelFile(vrtFileNameWithPath As Variant) As Long
    Dim strFileName As String
    Dim strMCCName As String
    Dim lngPosition As Long
    Dim StartCopyFromLine As Long

    lngPosition = InStrRev(vrtFileNameWithPath, "\")
    
    If lngPosition > 0 Then
            strFileName = Mid(vrtFileNameWithPath, lngPosition + 1)
    Else
        HandleOneExcelFile = 0
        Exit Function
    End If
    
    ShowMessage ("Start processing: " & strFileName)
    lngPosition = InStr(1, strFileName, "MCC_")
    If lngPosition > 0 Then
        strMCCName = Mid(strFileName, lngPosition + 4, 3)
    Else
        ShowMessage ("ERROR: Not a valid filename for an MCC Excel")
        HandleOneExcelFile = 0
        Exit Function
    End If
    
    On Error GoTo ErrorHandler1
    Set MCCSourceWorkbook = Workbooks.Open(vrtFileNameWithPath)
    Set MCCSourceWorkSheet = MCCSourceWorkbook.Sheets("MCC")
    GoTo NoError1

ErrorHandler1:
    ShowMessage ("ERROR: Not a valid MCC Excel file. Must contain a tab MCC")
    MCCSourceWorkbook.Close
    HandleOneExcelFile = 0
    Exit Function

NoError1:
    ' Now locate the MCC sheet and if it not exist, create it
    On Error GoTo ErrorHandler2
    Set MCCTargetWorkSheet = ThisWorkbook.Sheets(strMCCName)
    StartCopyFromLine = 3
    GoTo NoError2
    
ErrorHandler2:
    'MCC tab must be created
    ShowMessage ("Creating tab for MCC: " & strMCCName)
    ThisWorkbook.Sheets.Add.Name = strMCCName
    StartCopyFromLine = 1
    Set MCCTargetWorkSheet = ThisWorkbook.Sheets(strMCCName)

NoError2:
    'Empty mcc worksheet
    '***
    ShowMessage ("Copying data from MCC: " & strMCCName)
    Application.ScreenUpdating = False
    MCCTargetWorkSheet.Activate
    While MCCTargetWorkSheet.Cells(3, 2).Value = strMCCName
        MCCTargetWorkSheet.Rows(3).EntireRow.Delete
    Wend
    Application.ScreenUpdating = True
    
    'Copying the new data
    '***
    Dim MaxLineToCopy As Long
    
    For maxLinesToCopy = 3 To 999
        If (MCCSourceWorkSheet.Cells(maxLinesToCopy, 2).Value <> strMCCName) Then
            Exit For
        End If
    Next maxLinesToCopy
    
    MCCSourceWorkSheet.Activate
    ' copying the title lines?
    If StartCopyFromLine = 1 Then
        MCCSourceWorkSheet.Range("A1:W2").Copy
        
        Application.DisplayAlerts = False
        MCCTargetWorkSheet.Activate
        MCCTargetWorkSheet.Cells(1, 1).Select
    
        Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone _
                , SkipBlanks:=False, Transpose:=False
        
        Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
                , SkipBlanks:=False, Transpose:=False
        Application.DisplayAlerts = True
    End If
    
    MCCSourceWorkSheet.Range("A3:W" & maxLinesToCopy - 1).Copy
    
    MCCTargetWorkSheet.Activate
    MCCTargetWorkSheet.Cells(3, 1).Select
    Application.DisplayAlerts = False
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
            , SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    MCCSourceWorkbook.Close SaveChanges:=False
    Application.DisplayAlerts = True
    ShowMessage ("Finished processing: " & strFileName)
    HandleOneExcelFile = 1
End Function


Sub UpdateSummary()
    ShowMessage ("Start updating Summary")
        
    Set SummaryWorkSheet = ThisWorkbook.Sheets("Summary")

    Application.ScreenUpdating = False
    SummaryWorkSheet.Activate
    
    ShowMessage ("Deleting existing content")
    
    'Deleting
    Dim DeleteCounter As Long
    DeleteCounter = 1
    
    While SummaryWorkSheet.Cells(3, 2).Value <> ""
        SummaryWorkSheet.Rows(3).EntireRow.Delete
        DeleteCounter = DeleteCounter + 1
        If DeleteCounter Mod 200 = 0 Then
            ShowMessage ("Lines deleted: " & DeleteCounter)
        End If
    Wend
    ShowMessage ("Lines deleted: " & DeleteCounter)
    Application.ScreenUpdating = True
    
    Dim vrtAllSheets As Variant
    Dim MaxLineToCopy As Long, CurrentLineToPaste As Long
    
    CurrentLineToPaste = 3
    
    Application.ScreenUpdating = False
    For Each vrtAllSheets In ThisWorkbook.Sheets
        If Len(vrtAllSheets.Name) = 3 And (InStr(1, vrtAllSheets.Name, UCase(vrtAllSheets.Name), vbBinaryCompare) = 1) Then
            
            ShowMessage ("Inserting content from " & vrtAllSheets.Name)
            For maxLinesToCopy = 3 To 999
                If (vrtAllSheets.Cells(maxLinesToCopy, 2).Value <> vrtAllSheets.Name) Then
                    Exit For
                End If
            Next maxLinesToCopy
            
            vrtAllSheets.Range("A3:W" & maxLinesToCopy - 1).Copy
            
            SummaryWorkSheet.Activate
            SummaryWorkSheet.Cells(CurrentLineToPaste, 1).Select
        
            Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
                    , SkipBlanks:=False, Transpose:=False
                    
            CurrentLineToPaste = CurrentLineToPaste + maxLinesToCopy - 3
            
        End If
    Next vrtAllSheets
    
    Call UpdatePivot
    
    ShowMessage ("Finished updating Summary")
    If (ClosingMessageByOtherSub = False) Then ShowMessage ("Finished - This window can be closed")
    Application.ScreenUpdating = True
    
End Sub

Sub testtextmessagebox()

Dim i As Integer

For i = 0 To 20
    ShowMessage ("Test Message")
Next i

End Sub

Sub ShowMessage(TextToAdd As String)

    If (StatusWindow.Visible = False) Then StatusWindow.Show
    
    StatusWindow.MsgTextBox.SetFocus
    StatusWindow.MsgTextBox.Text = StatusWindow.MsgTextBox.Text & TextToAdd & vbCrLf
    StatusWindow.MsgTextBox.SelStart = Len(StatusWindow.MsgTextBox.Text)
    
    StatusWindow.Repaint
    Application.Wait (Now + TimeValue("0:00:01"))
    
End Sub

Sub UpdatePivot()

    ShowMessage ("Start updating Pivot table")
    
    Dim maxLinesToCopy As Long
    Dim rng As Range
    
    Set SummaryWorkSheet = ThisWorkbook.Sheets("Summary")

    For maxLinesToCopy = 3 To 3000
        If (SummaryWorkSheet.Cells(maxLinesToCopy, 2).Value = "") Then
            Exit For
        End If
    Next maxLinesToCopy
    
    Set Pivottbl = ThisWorkbook.Sheets("Pivot").PivotTables(1).PivotCache
    
    Set rng = SummaryWorkSheet.Range("A2:W" & maxLinesToCopy - 1)
    Pivottbl.SourceData = rng.Address(True, True, xlR1C1, True)
    Pivottbl.Refresh
    
    ShowMessage ("Finished updating Pivot table")
    If (ClosingMessageByOtherSub = False) Then ShowMessage ("Finished - This window can be closed")
End Sub

Sub OrderSheets()
    ShowMessage ("Start ordering Sheets")
    'Getting total no. of worsheets in workbook
    SCount = Worksheets.Count
    
    'Checking condition whether count of worksheets is greater than 1, If count is one then exit the procedure
    If SCount = 1 Then Exit Sub
    
    'Using Bubble sort as sorting algorithm
    'Looping through all worksheets
    For i = 1 To SCount - 1
    
        'Making comparison of selected sheet name with other sheets for moving selected sheet to appropriate position
        For j = i + 1 To SCount
            If Worksheets(j).Name < Worksheets(i).Name Then
                Worksheets(j).Move Before:=Worksheets(i)
            End If
        Next j
    Next i

    ' move admin page to front
    For i = 1 To SCount - 1
        If Worksheets(i).Name = "Admin_Sheet" Then
            Worksheets(i).Move Before:=Worksheets(1)
        End If
        If Worksheets(i).Name = "Pivot" Then
            Worksheets(i).Move Before:=Worksheets(2)
        End If
        If Worksheets(i).Name = "Summary" Then
            Worksheets(i).Move Before:=Worksheets(3)
        End If
    Next i
    
    Application.ScreenUpdating = False
    
    ShowMessage ("Finished ordering Sheets")
    If (ClosingMessageByOtherSub = False) Then ShowMessage ("Finished - This window can be closed")
End Sub
