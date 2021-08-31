Attribute VB_Name = "mod6_Format_Reports"
Sub FormatPayroll()
'
' Format worksheet view
'
Application.PrintCommunication = False
    
    With wsPayroll.PageSetup
        'Headers & footers
        .LeftHeader = "Payroll - Invoice Week " & invoiceWeek
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .DifferentFirstPageHeaderFooter = False
        
        'Margins
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        
        'View
        .Zoom = 70
        .PrintErrors = xlPrintErrorsNA
        .ScaleWithDocHeaderFooter = True
    End With

Application.PrintCommunication = True

End Sub

Sub PoolPayroll()
'Add Bad Debt Cancel to Payroll Report
    
    'Get Pool Saves from wsPool and paste to wsPayroll
    Sheets("Pool").Activate
    Sheets("Pool").Range("Q4:W4").Copy
    wsPayroll.Activate
    wsPayroll.Range("B9").PasteSpecial xlPasteValues
    
    'Call UDF getBulkTechID and insert value into payroll template.
    wsPayroll.Range("B2").Value = getBulkTechID
    
    'Call UDF count functions to fill other payroll features
    wsPayroll.Range("B25").Value = countJobs(getBulkTechID)
    wsPayroll.Range("B26").Value = countCancels(getBulkTechID)
    wsPayroll.Range("B27").Value = countSaves(getBulkTechID)
    wsPayroll.Range("B28").Value = countUTCs(getBulkTechID)
    
End Sub

Sub TechPayroll()
'Populate values in payroll template for each techID

Dim techName As String
Dim pastecol As Integer
Dim oldSaves As Range, newSaves As Range
Dim oldDiscos As Range, newDiscos As Range

Sheets("Payroll").Cells.PageBreak = xlPageBreakNone
pastecol = 10

For Each techID In techIDRange
    If techID.Value = "" Then Exit For
    If techID.Value <> "" Then
        
        'Populate Variables
        techNumber = techID.Value
        techName = Sheets(techNumber).Range("F2").Value
        With wsPayroll
            Set oldSaves = .Range(.Cells(7, pastecol + 1), .Cells(7, pastecol + 7))
            Set newSaves = .Range(.Cells(8, pastecol + 1), .Cells(8, pastecol + 7))
            Set oldDiscos = .Range(.Cells(20, pastecol + 1), .Cells(20, pastecol + 7))
            Set newDiscos = .Range(.Cells(21, pastecol + 1), .Cells(21, pastecol + 7))
        End With
        
        'Copy/paste new payroll template on new page in wsPayroll.  Add tech name&id.
        With wsPayroll
            .Activate
            .Range("A1:I34").Copy
            .Cells(1, pastecol).PasteSpecial xlPasteValuesAndNumberFormats
            .Cells(1, pastecol).PasteSpecial xlPasteFormats
            .Cells(1, pastecol).PasteSpecial xlPasteColumnWidths
            .Cells(1, pastecol).PasteSpecial xlPasteFormulas
            .Cells(1, pastecol).PageBreak = xlPageBreakManual
            .Cells(1, pastecol + 1).Value = techName
            .Cells(2, pastecol + 1).Value = techID.Value
        End With
        
        'Monday-Friday Saves/Discos
            oldSaves.Value = Sheets(techNumber).Range("U4:AA4").Value
            newSaves.Value = Sheets(techNumber).Range("AK4:AQ4").Value
            oldDiscos.Value = Sheets(techNumber).Range("AC4:AI4").Value
            newDiscos.Value = Sheets(techNumber).Range("AS4:AY4").Value
            
        'Equipment
        With wsPayroll
            .Cells(11, pastecol + 7).Value = Sheets(techNumber).Range("BA4").Value
            .Cells(12, pastecol + 7).Value = Sheets(techNumber).Range("BB4").Value
            .Cells(13, pastecol + 7).Value = Sheets(techNumber).Range("BC4").Value
            .Cells(14, pastecol + 7).Value = Sheets(techNumber).Range("BD4").Value
        End With
        
       'Summary job counts
        With Sheets("Payroll")
            .Cells(25, pastecol + 1).Value = countJobs(techNumber)
            .Cells(26, pastecol + 1).Value = countCancels(techNumber)
            .Cells(27, pastecol + 1).Value = countSaves(techNumber)
            .Cells(28, pastecol + 1).Value = countUTCs(techNumber)
        End With
        
       'Set pasteCol for next techID
       pastecol = pastecol + 9
    End If
Next techID

End Sub

Sub Production()

Dim tblProduction As ListObject
Set tblProduction = wsReports.ListObjects("Production")

'Issue date and turn-in date
With wsReports
    .Range("C2").Value = dateLastClosed
    .Range("C3").Value = dateLastClosed + 14
End With

'Pool Production
With wsReports
    .Range("A6").Value = getBulkTechID
    .Range("B6").Value = countJobs(getBulkTechID)
    .Range("C6").Value = sumDupCancelRS(getBulkTechID)
    .Range("E6").Value = countSaves(getBulkTechID)
    .Range("G6").Value = sumSaves(getBulkTechID)
End With

'Tech Production
Dim techRow As ListRow

For Each techID In techIDRange
    If techID.Value = "" Then Exit For
    If techID.Value <> "" Then
        techNumber = techID.Value
        Set techRow = tblProduction.ListRows.Add(AlwaysInsert:=True)
        With techRow
            .Range.Cells(1, 1).Value = techNumber
            .Range.Cells(1, 2).Value = countJobs(techNumber)
            .Range.Cells(1, 3).Value = sumDupCancelRS(techNumber)
            .Range.Cells(1, 5).Value = countSaves(techNumber)
            .Range.Cells(1, 7).Value = sumSaves(techNumber)
            .Range.Cells(1, 17).Value = countDiscos(techNumber)
        End With
    End If
Next techID

End Sub


Public Function getBulkTechID() As String
    If Sheets("Pool").Range("H2").Value <> "" Then
        getBulkTechID = Sheets("Pool").Range("H2").Value
        Else: getBulkTechID = "Bulk Tech ID Not Found"
    End If
End Function

Public Function getTechName(techNumber As String) As String
    Sheets(techNumber).Activate
    getTechName = Sheets(techNumber).Range("F2").Value
    Debug.Print getTechName
End Function

Public Function countJobs(techNumber As String) As Integer
    Sheets("LastClosedIssue").Activate
    lastRow = Sheets("LastClosedIssue").Cells(Rows.Count, 1).End(xlUp).row
    countJobs = WorksheetFunction.CountIf(Sheets("LastClosedIssue").Range("H2:H" & lastRow), techNumber)
    Debug.Print countJobs
End Function

Public Function countSaves(techNumber As String) As Integer
    Sheets("LastClosedIssue").Activate
    lastRow = Sheets("LastClosedIssue").Cells(Rows.Count, 1).End(xlUp).row
    countSaves = WorksheetFunction.CountIfs(Sheets("LastClosedIssue").Range("H2:H" & lastRow), techNumber, Sheets("LastClosedIssue").Range("K2:K" & lastRow), "SAVE")
    Debug.Print countSaves
End Function
Public Function sumSaves(techNumber As String) As Integer
    Sheets("LastClosedIssue").Activate
    lastRow = Sheets("LastClosedIssue").Cells(Rows.Count, 1).End(xlUp).row
    sumSaves = WorksheetFunction.SumIf(Sheets("LastClosedIssue").Range("H2:H" & lastRow), _
            techNumber, Sheets("LastClosedIssue").Range("L2:L" & lastRow))
    Debug.Print sumSaves
End Function
Public Function countCancels(techNumber As String) As Integer
    Sheets("LastClosedIssue").Activate
    lastRow = Sheets("LastClosedIssue").Cells(Rows.Count, 1).End(xlUp).row
    countCancels = WorksheetFunction.CountIfs(Sheets("LastClosedIssue").Range("H2:H" & lastRow), techNumber, Sheets("LastClosedIssue").Range("K2:K" & lastRow), "CANCEL")
    Debug.Print countCancels
End Function

Public Function countUTCs(techNumber As String) As Integer
    Sheets("LastClosedIssue").Activate
    lastRow = Sheets("LastClosedIssue").Cells(Rows.Count, 1).End(xlUp).row
    countUTCs = WorksheetFunction.CountIfs(Sheets("LastClosedIssue").Range("H2:H" & lastRow), techNumber, Sheets("LastClosedIssue").Range("K2:K" & lastRow), "UTC")
    Debug.Print countUTCs
End Function
Public Function countDUPS(techNumber As String) As Integer
    Sheets("LastClosedIssue").Activate
    lastRow = Sheets("LastClosedIssue").Cells(Rows.Count, 1).End(xlUp).row
    countDUPS = WorksheetFunction.CountIfs(Sheets("LastClosedIssue").Range("H2:H" & lastRow), techNumber, Sheets("LastClosedIssue").Range("K2:K" & lastRow), "DUP")
    
End Function
Public Function countRSs(techNumber As String) As Integer
    Sheets("LastClosedIssue").Activate
    lastRow = Sheets("LastClosedIssue").Cells(Rows.Count, 1).End(xlUp).row
    countRSs = WorksheetFunction.CountIfs(Sheets("LastClosedIssue").Range("H2:H" & lastRow), techNumber, Sheets("LastClosedIssue").Range("K2:K" & lastRow), "R/S")
    End Function
Public Function countDiscos(techNumber As String) As Integer
    Sheets("LastClosedIssue").Activate
    lastRow = Sheets("LastClosedIssue").Cells(Rows.Count, 1).End(xlUp).row
    countDiscos = WorksheetFunction.CountIfs(Sheets("LastClosedIssue").Range("H2:H" & lastRow), techNumber, Sheets("LastClosedIssue").Range("K2:K" & lastRow), "DISCO")
    
End Function
Public Function sumDupCancelRS(techNumber As String) As Integer
    sumDupCancelRS = countDUPS(techNumber) + countRSs(techNumber) + countCancels(techNumber)
End Function



