Attribute VB_Name = "mod2_Format_Data"
Option Explicit
'
'mod2_Format_Source_Data
'

Sub CombineTracking()
'Take user-provided tracking data and compile on one worksheet
    
    Dim sourceTracking1 As Worksheet
    Dim sourceTracking2 As Worksheet
    Dim sourceTracking3 As Worksheet
                  
    Set sourceTracking1 = Sheets("Tracking1")
    Set sourceTracking2 = Sheets("Tracking2")
    Set sourceTracking3 = Sheets("Tracking3")
    
    lastRow = sourceTracking1.Cells(Rows.Count, 1).End(xlUp).row
               
'Copy contents of "Tracking1" worksheet and paste to "Tracking" worksheet
    'includes headers
    sourceTracking1.Activate
    sourceTracking1.Range("A1:BP" & lastRow).Copy
                      
    wsTracking.Activate
    wsTracking.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
   
'Repeat with "Tracking2" and "Tracking3" worksheets. Paste below UsedRange. Offset to exclude header row.
     With sourceTracking2
        .Activate
        Intersect(.UsedRange, .UsedRange.Offset(1)).Copy
    End With
     
    wsTracking.Activate
    wsTracking.Cells(Rows.Count, 1).End(xlUp).Offset(1).PasteSpecial xlPasteValuesAndNumberFormats
            
    With sourceTracking3
        .Activate
        Intersect(.UsedRange, .UsedRange.Offset(1)).Copy
    End With

    wsTracking.Activate
    wsTracking.Cells(Rows.Count, 1).End(xlUp).Offset(1).PasteSpecial xlPasteValuesAndNumberFormats

'Delete source tracking worksheets
    sourceTracking1.Delete
    sourceTracking2.Delete
    sourceTracking3.Delete

End Sub

Sub FormatTracking()
'Format tracking data

lastRow = wsTracking.Cells(Rows.Count, 1).End(xlUp).row
       
    'Delete unnecessary columns
    With wsTracking
        Columns("A:A").Delete Shift:=xlToLeft
        Columns("E:BC").Delete Shift:=xlToLeft
    End With
       
    'Create tblTracking List Object (store data in table)
    With wsTracking
        .Range("A1:N" & lastRow).Select
        .ListObjects.Add(xlSrcRange, Range("A1:N" & lastRow), , xlYes).Name = "tblTracking"
    End With

    Set tblTracking = wsTracking.ListObjects("tblTracking")

    With tblTracking
        'Set Column Widths
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
    
    'Sort tblTracking
    With tblTracking.Sort
        .SortFields.Clear
        .SortFields.Add _
            Key:=Range("tblTracking[TechNbr]")
        .SortFields.Add _
            Key:=Range("tblTracking[STATUS]"), _
            CustomOrder:="SAVE,DISCO,R/S,CANCEL,DUP,UTC"
        .SortFields.Add _
            Key:=Range("tblTracking[DATE]")
        .SortFields.Add _
            Key:=Range("tblTracking[JobNbr]")
        .Header = xlYes
        .Apply
    End With
    
End Sub

Sub AddGroupColumn()
'Add a column to tblTracking that identifies each job as "old issue" or "new issue"
    
    'Insert Column and Name it
    wsTracking.Activate
    wsTracking.ListObjects("tblTracking").ListColumns.Add _
        Position:=14
    wsTracking.ListObjects("tblTracking").ListColumns(14).Name = "IssueGroup"

    'Loop through cells in schedule_date column (row,4) and compare value to issue dates
        'Set corresponding value in issue_group column (r,14)
    Dim row As Long
    lastRow = wsTracking.Cells(Rows.Count, 1).End(xlUp).row

    For row = 2 To lastRow
        If IsDate(wsTracking.Cells(row, 4).Value) And wsTracking.Cells(row, 4).Value = dateIssue1 Then
            wsTracking.Cells(row, 14).Value = "Old Issue"
        ElseIf IsDate(wsTracking.Cells(row, 4).Value) And wsTracking.Cells(row, 4).Value = dateIssue2 Then
            wsTracking.Cells(row, 14).Value = "Old Issue"
        ElseIf IsDate(wsTracking.Cells(row, 4).Value) And wsTracking.Cells(row, 4).Value = dateIssue3 Then
            wsTracking.Cells(row, 14).Value = "New Issue"
        Else:  wsTracking.Cells(row, 14).Value = "Check Schedule Date"
        End If
    Next row

End Sub

Sub FormatEQ()
'Format Equipment Logs

lastRow = wsEQ.Cells(Rows.Count, 1).End(xlUp).row
 
'Change column header to avoid hash tag problems
    wsEQ.Range("A1").Value = "TechNbr"

'Format data in wsEQ as List Object (table)
    With wsEQ
        .ListObjects.Add(xlSrcRange, wsEQ.Range("A1:H" & lastRow), , xlYes).Name = "tblEQ"
    End With

    Set tblEQ = wsEQ.ListObjects("tblEQ")
          
'Loop techID Numbers.  Delete empty rows and format techID as Long number
Dim cell As Range, col As Range
Dim row As Long

Set col = tblEQ.ListColumns(1).DataBodyRange

    For Each cell In col
        If cell.Value = "" Then
            Rows(cell.row).Delete Shift:=xlUp
        End If
    Next cell
          
End Sub

Sub FormatNumbers()
'Fix data mismatch bug
    
'Set techIDRange Cells to Correct Number Format
    wsUserInput.Range("F5:F14").NumberFormat = "#"

'Set tblTracking Cells to Correct Number Format
    With tblTracking
        .ListColumns(1).DataBodyRange.NumberFormat = "#"
        .ListColumns(2).DataBodyRange.NumberFormat = "#"
        .ListColumns(4).DataBodyRange.NumberFormat = "m/d/yyyy"
        .ListColumns(5).DataBodyRange.NumberFormat = "#"
        .ListColumns(8).DataBodyRange.NumberFormat = "#"
        .ListColumns(10).DataBodyRange.NumberFormat = "m/d/yyyy"
        .ListColumns(12).DataBodyRange.NumberFormat = "$#,##0.00"
    End With
    
'Set tblEQ Cells to Correct Number Format
    With tblEQ
        .ListColumns(1).DataBodyRange.NumberFormat = "#"
        .ListColumns(2).DataBodyRange.NumberFormat = "@"
        .ListColumns(5).DataBodyRange.NumberFormat = "@"
        .ListColumns(8).DataBodyRange.NumberFormat = "@"
    End With

'Convert numeric values to Long data type

Dim cell As Range, col As Range
Dim rCol As Variant, element As Variant
  
    'Convert techID values to Long datatype
    For Each techID In techIDRange
        If techID.Value <> "" And IsNumeric(techID.Value) Then
            techID.Value = CLng(techID.Value)
        End If
    Next techID
    
    'Convert tblTracking WO#, acct#, TechNbr to Long datatype
    rCol = Array(1, 2, 8)
    For Each element In rCol
        Set col = tblTracking.ListColumns(element).DataBodyRange
        For Each cell In col
            If cell.Value <> "" And IsNumeric(cell.Value) Then
                cell.Value = CLng(cell.Value)
            End If
        Next cell
    Next element
    
    'Convert tblEQ TechNbr to Long datatype
    Set col = tblEQ.ListColumns(1).DataBodyRange
        For Each cell In col
            If cell.Value <> "" And IsNumeric(cell.Value) Then
                cell.Value = CLng(cell.Value)
            End If
        Next cell

End Sub
 
