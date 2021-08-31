Attribute VB_Name = "mod3_Organize_Data"
Option Explicit

'
' mod3_Organize_Data
' Sort data for use with calculations
'
Sub addWeekNumberWS()
'wsWeekNumber

'Create WeekNumber Worksheet for current invoice week data
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "Week_" & invoiceWeek
    
    Set wsWeekNumber = ActiveWorkbook.Sheets("Week_" & invoiceWeek)

'Filter tblTracking for current invoice week data
    tblTracking.Range.AutoFilter _
        field:=13, _
        Criteria1:=invoiceWeek
                        
'Copy/Paste filtered tblTracking data to WeekNumber worksheet
    tblTracking.Range.SpecialCells(xlCellTypeVisible).Copy
        wsWeekNumber.Activate
        wsWeekNumber.Range("A1").PasteSpecial xlPasteAll

'Format data in wsWeekNumber as Table
    lastRow = wsWeekNumber.Cells(Rows.Count, 1).End(xlUp).row
    
    With wsWeekNumber
        .Range("A1:N" & lastRow).Select
        .ListObjects.Add(xlSrcRange, Range("A1:N" & lastRow), , xlYes).Name = "tblWeekNumber"
    End With

    Set tblWeekNumber = wsWeekNumber.ListObjects("tblWeekNumber")
                        
'Set Column Widths
    
    With tblWeekNumber
        .ListColumns(1).Range.ColumnWidth = 10
        .ListColumns(2).Range.ColumnWidth = 10
        .ListColumns(3).Range.ColumnWidth = 20
        .ListColumns(4).Range.ColumnWidth = 10
        .ListColumns(5).Range.ColumnWidth = 5
        .ListColumns(6).Range.ColumnWidth = 8
        .ListColumns(7).Range.ColumnWidth = 5
        .ListColumns(8).Range.ColumnWidth = 8
        .ListColumns(9).Range.ColumnWidth = 5
        .ListColumns(10).Range.ColumnWidth = 10
        .ListColumns(11).Range.ColumnWidth = 10
        .ListColumns(12).Range.ColumnWidth = 10
        .ListColumns(13).Range.ColumnWidth = 5
        .ListColumns(14).Range.ColumnWidth = 5
    End With
    
'Unfilter Tracking Worksheet
    wsTracking.Activate
    tblTracking.AutoFilter.ShowAllData
                       
End Sub

Sub addLastClosedWS()
'wsLastClosed

'Create LastClosedIssue Worksheet for data from issue most recently closed
    Sheets.Add(After:=Sheets("Tracking")).Name = "LastClosedIssue"
                    
    Set wsLastClosed = ActiveWorkbook.Sheets("LastClosedIssue")
                    
'Filter tblTracking for schedule date=dateLastClosed
    wsTracking.Activate
    tblTracking.Range.AutoFilter _
        field:=4, _
        Criteria1:=dateLastClosed
                        
'Copy/Paste filtered tblTracking data to wsLastClosed
    tblTracking.Range.SpecialCells(xlCellTypeVisible).Copy
    
    wsLastClosed.Activate
    wsLastClosed.Range("A1").PasteSpecial xlPasteAll
                       
'Format data in wsLastClosed as Table
    lastRow = wsLastClosed.Cells(Rows.Count, 1).End(xlUp).row
    
    With wsLastClosed
        .Range("A1:N" & lastRow).Select
        .ListObjects.Add(xlSrcRange, Range("A1:N" & lastRow), , xlYes).Name = "tblLastClosed"
    End With

    Set tblLastClosed = wsLastClosed.ListObjects("tblLastClosed")
                       
'Set Column Widths
    
    With tblLastClosed
        .ListColumns(1).Range.ColumnWidth = 10
        .ListColumns(2).Range.ColumnWidth = 10
        .ListColumns(3).Range.ColumnWidth = 20
        .ListColumns(4).Range.ColumnWidth = 10
        .ListColumns(5).Range.ColumnWidth = 5
        .ListColumns(6).Range.ColumnWidth = 8
        .ListColumns(7).Range.ColumnWidth = 5
        .ListColumns(8).Range.ColumnWidth = 8
        .ListColumns(9).Range.ColumnWidth = 5
        .ListColumns(10).Range.ColumnWidth = 10
        .ListColumns(11).Range.ColumnWidth = 10
        .ListColumns(12).Range.ColumnWidth = 10
        .ListColumns(13).Range.ColumnWidth = 5
        .ListColumns(14).Range.ColumnWidth = 5
    End With
    
'Unfilter Tracking Worksheet
    wsTracking.Activate
    tblTracking.AutoFilter.ShowAllData

End Sub
Sub addPoolWS()
'wsPool

'Create Pool Worksheet for BulkTech data
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "Pool"
    Set wsPool = ActiveWorkbook.Sheets("Pool")

'Filter tblWeekNumber for Name="Pool"
    wsWeekNumber.Activate
    tblWeekNumber.Range.AutoFilter _
        field:=6, _
        Criteria1:="Pool"
                        
'Copy/Paste filtered tblWeekNumber data to wsPool
    tblWeekNumber.Range.SpecialCells(xlCellTypeVisible).Copy
    
    wsPool.Activate
    wsPool.Range("A1").PasteSpecial xlPasteAll
    
'Fix some formatting
    With wsPool
        .Columns(15).Delete
        .Columns(14).ColumnWidth = 10
        .Columns(15).ColumnWidth = 2
        .Columns(16).ColumnWidth = 2
    End With
    
'Unfilter tblWeekNumber
    wsWeekNumber.Activate
    tblWeekNumber.AutoFilter.ShowAllData
    
End Sub

Sub addTechWorksheets()

'Loop:  Create new WS for each TechID and copy/paste Tech's data from tblWeekNumber

    For Each techID In techIDRange
        If techID = "" Then Exit For
        If techID <> "" Then
            'Populate techNumber variable with current techID in techIDRange
            techNumber = techID.Value
                         
            'Create Worksheet for Individual TechID
            Sheets.Add(After:=Sheets(Sheets.Count)).Name = techID.Value
                         
            'Filter WeekNumber data for records corresponding to techID
            wsWeekNumber.Activate
            tblWeekNumber.Range.AutoFilter _
                field:=8, _
                Criteria1:=techID.Value
                tblWeekNumber.Range.SpecialCells(xlCellTypeVisible).Copy
            
            'Paste data to corresponding techID worksheet
            Sheets(techNumber).Activate
            Sheets(techNumber).Range("A1").PasteSpecial xlPasteAll
            
            'Unfilter tblWeekNumber
            wsWeekNumber.Activate
            tblWeekNumber.AutoFilter.ShowAllData
        End If
    Next techID
End Sub
Sub SortTechEQ()
'Sort EQLog data into techID worksheets

Dim sqlEQ As String
Dim rsEQ As ADODB.Recordset

'Establish ADODB connection to workbook
Call ConnectToSelf
    
'Loop through techID worksheets and add equipment from wsEQ
For Each techID In techIDRange
    If techID = "" Then Exit For
    If techID <> "" Then
        'Populate variables for this techID
        techNumber = techID.Value
                    
        'Copy EQLogs Column Headers to tech worksheet
        Sheets(techNumber).Activate
        Sheets(techNumber).Range("P1:S1").Value = wsEQ.Range("E1:H1").Value
                                      
        'Use SQL query of wsEQ to create recordset of tech's equipment.  Copy rs to tech ws.
        sqlEQ = "SELECT [Serial Number], [Item], [Description], [Type] FROM [EQLogs$] WHERE [TechNbr] =" & techNumber
        Set rsEQ = Nothing
        Set rsEQ = New ADODB.Recordset
        rsEQ.Open sqlEQ, selfConn, adOpenStatic, adLockReadOnly
        Sheets(techNumber).Range("P2").CopyFromRecordset rsEQ
        rsEQ.Close
    End If
Next techID

'Close ADODB connection
selfConn.Close

End Sub

Sub CreateTechTables()

Dim tblTechNumber As ListObject

'Loop through each tech ID
    For Each techID In techIDRange
        If techID.Value = "" Then Exit For
        If techID.Value <> "" Then
            techNumber = techID.Value
            With Sheets(techNumber)
                .Activate
                lastRow = .UsedRange.Rows(.UsedRange.Rows.Count).row
                'Create tbl(techNumber) List Object (store data as table)
                .Range("A1:S" & lastRow).Select
                .ListObjects.Add(xlSrcRange, .Range("A1:S" & lastRow), , xlYes).Name = "tbl" & techNumber
            End With

            Set tblTechNumber = Sheets(techNumber).ListObjects("tbl" & techNumber)
            
            'Set Column Widths
            With Sheets(techNumber)
                .Columns(1).ColumnWidth = 10
                .Columns(2).ColumnWidth = 10
                .Columns(3).ColumnWidth = 20
                .Columns(4).ColumnWidth = 10
                .Columns(5).ColumnWidth = 5
                .Columns(6).ColumnWidth = 8
                .Columns(7).ColumnWidth = 5
                .Columns(8).ColumnWidth = 8
                .Columns(9).ColumnWidth = 5
                .Columns(10).ColumnWidth = 10
                .Columns(11).ColumnWidth = 10
                .Columns(12).ColumnWidth = 10
                .Columns(13).ColumnWidth = 5
                .Columns(14).ColumnWidth = 5
                .Columns(15).ColumnWidth = 8
                .Columns(16).ColumnWidth = 15
                .Columns(17).ColumnWidth = 8
                .Columns(18).ColumnWidth = 12
                .Columns(19).ColumnWidth = 10
                .Columns(20).ColumnWidth = 1
            End With
           
            'Group tech tracking to clean-up worksheet view
            With Sheets(techNumber)
                .Columns("A:S").Columns.Group
                .Outline.ShowLevels ColumnLevels:=1
            End With
        End If
    Next techID
End Sub

