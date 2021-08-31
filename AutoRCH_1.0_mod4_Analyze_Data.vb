Attribute VB_Name = "mod4_Summarize_Data"
Option Explicit

Sub ListPoolSaves()

Dim wsPool As Worksheet
Dim dateRange As Range, dateCell As Range
Dim jobDate As Date
Dim sqlString As String

Dim pastecol As Long

Dim rsSaveAmt As ADODB.Recordset

Set wsPool = ActiveWorkbook.Sheets("Pool")

'Copy/paste skeleton list format from wsReference.
    wsReference.Activate
    wsReference.Range("Z1:AF100").Copy
    wsPool.Activate
    wsPool.Range("Q1").PasteSpecial xlPasteValuesAndNumberFormats
    wsPool.Range("Q1").PasteSpecial xlPasteFormats
    wsPool.Range("Q1").PasteSpecial xlPasteColumnWidths

'Establish ADODB connection to workbook
    Call ConnectToSelf
    
'Query NEW ISSUE SAVES AMOUNT For Each Date
    Set dateRange = wsPool.Range("Q3:W3")
            
    For Each dateCell In dateRange
        jobDate = dateCell.Value
        pastecol = dateCell.Column
              
        'Set query parameters
        sqlString = "SELECT [AMOUNT] FROM [Pool$] WHERE ([DATE]=#" & jobDate & "#)"
                
        'Create RecordSet with SQL query and paste in row 5 of date column
        Set rsSaveAmt = Nothing
        Set rsSaveAmt = New ADODB.Recordset
        rsSaveAmt.Open sqlString, selfConn, adOpenStatic, adLockReadOnly
        wsPool.Cells(5, pastecol).CopyFromRecordset rsSaveAmt
        rsSaveAmt.Close
    Next dateCell

'Paste aggregate formulas from wsReference
    wsReference.Activate
    wsReference.Range("Z4:AF4").Copy
            
    wsPool.Activate
    wsPool.Range("Q4").PasteSpecial xlPasteFormulas
           
'Close connection
    selfConn.Close
                
End Sub

Sub ListJobsByDate()
'Summarize completed jobs for each techID

Dim dateRange As Range, dateCell As Range
Dim jobDate As Date
Dim sqlString As String

Dim pastecol As Long

Dim rsAcctNbr As ADODB.Recordset
Dim rsSaveAmt As ADODB.Recordset

'Establish ADODB connection to workbook
Call ConnectToSelf
    
'Loop through techID worksheets
    For Each techID In techIDRange
        If techID = "" Then Exit For
        If techID <> "" Then
            'Populate variables for this techID
            techNumber = techID.Value
            lastRow = Sheets(techNumber).Cells(Rows.Count, 1).End(xlUp).row
            
            'Copy/paste skeleton list format from wsReference.
            wsReference.Activate
            wsReference.Range("J1:AS100").Copy
            
            Sheets(techNumber).Activate
            Sheets(techNumber).Range("U1").PasteSpecial xlPasteValuesAndNumberFormats
            Sheets(techNumber).Range("U1").PasteSpecial xlPasteFormats
            Sheets(techNumber).Range("U1").PasteSpecial xlPasteColumnWidths
                            
            'Query OLD ISSUE SAVES AMOUNT For Each Date
            Set dateRange = Sheets(techNumber).Range("U3:AA3")
            
            For Each dateCell In dateRange
                jobDate = dateCell.Value
                pastecol = dateCell.Column
                
                'Set query parameters
                sqlString = "SELECT [AMOUNT] FROM [" & techNumber & "$] WHERE ([STATUS] = 'SAVE' AND [IssueGroup]='Old Issue' AND [DATE]=#" & jobDate & "#)"
                
                'Create RecordSet with SQL query and paste in row 5 of date column
                Set rsSaveAmt = Nothing
                Set rsSaveAmt = New ADODB.Recordset
                rsSaveAmt.Open sqlString, selfConn, adOpenStatic, adLockReadOnly
                Sheets(techNumber).Cells(5, pastecol).CopyFromRecordset rsSaveAmt
                rsSaveAmt.Close
            Next dateCell
            
            'Query NEW ISSUE SAVES AMOUNT For Each Date
            Set dateRange = Sheets(techNumber).Range("AK3:AQ3")
            
            For Each dateCell In dateRange
                jobDate = dateCell.Value
                pastecol = dateCell.Column
                
                'Set query parameters
                sqlString = "SELECT [AMOUNT] FROM [" & techNumber & "$] WHERE ([STATUS] = 'SAVE' AND [IssueGroup]='New Issue' AND [DATE]=#" & jobDate & "#)"
                
                'Create RecordSet with SQL query and paste in row 5 of date column
                Set rsSaveAmt = Nothing
                Set rsSaveAmt = New ADODB.Recordset
                rsSaveAmt.Open sqlString, selfConn, adOpenStatic, adLockReadOnly
                Sheets(techNumber).Cells(5, pastecol).CopyFromRecordset rsSaveAmt
                rsSaveAmt.Close
            Next dateCell
            
            'Query OLD ISSUE DISCO ACCT# For Each Date
            Set dateRange = Sheets(techNumber).Range("AC3:AI3")
            
            For Each dateCell In dateRange
                jobDate = dateCell.Value
                pastecol = dateCell.Column
                
                Debug.Print jobDate
                Debug.Print pastecol
                'Set query parameters
                sqlString = "SELECT [account_number] FROM [" & techNumber & "$] WHERE ([STATUS] = 'DISCO' AND [IssueGroup]='Old Issue' AND [DATE]=#" & jobDate & "#)"
            
                'Create RecordSet with SQL query
                Set rsAcctNbr = Nothing
                Set rsAcctNbr = New ADODB.Recordset
                rsAcctNbr.Open sqlString, selfConn, adOpenStatic, adLockReadOnly
                Sheets(techNumber).Cells(5, pastecol).CopyFromRecordset rsAcctNbr
                rsAcctNbr.Close
            Next dateCell
                        
            'Query NEW ISSUE DISCO ACCT# For Each Date
            Set dateRange = Sheets(techNumber).Range("AS3:AY3")
            
            For Each dateCell In dateRange
                jobDate = dateCell.Value
                pastecol = dateCell.Column

                'Set query parameters
                sqlString = "SELECT [account_number] FROM [" & techNumber & "$] WHERE ([STATUS] = 'DISCO' AND [IssueGroup]='New Issue' AND [DATE]=#" & jobDate & "#)"
            
                'Create RecordSet with SQL query
                Set rsAcctNbr = Nothing
                Set rsAcctNbr = New ADODB.Recordset
                rsAcctNbr.Open sqlString, selfConn, adOpenStatic, adLockReadOnly
                Sheets(techNumber).Cells(5, pastecol).CopyFromRecordset rsAcctNbr
                rsAcctNbr.Close
            Next dateCell
            
            'Paste aggregate formulas from wsReference
            wsReference.Activate
            wsReference.Range("J4:AS4").Copy
            
            Sheets(techNumber).Activate
            Sheets(techNumber).Range("U4").PasteSpecial xlPasteFormulas
        End If
    Next techID

'Close ADODB connection
selfConn.Close

End Sub

Sub ListEQByType()
'Sort EQLog data into techID worksheets
Dim sqlEQ As String
Dim rsEQ As ADODB.Recordset

Dim eqRange As Range, eqCell As Range
Dim eqType As String
Dim pastecol As Long

'Establish ADODB connection to workbook
Call ConnectToSelf
    
'Loop through techID worksheets and add equipment from wsEQ
For Each techID In techIDRange
    If techID = "" Then Exit For
    If techID <> "" Then
        'Populate variables for this techID
        techNumber = techID.Value
                                
        'Loop through eq type headers ("eqRange") and list corresponding serial numbers
        Set eqRange = Sheets(techNumber).Range("BA3:BD3")
        For Each eqCell In eqRange
            eqType = eqCell.Value
            pastecol = eqCell.Column
                      
            'Use SQL query wsEQ to create recordsets for tech's eq by type/category.  Copy rs to tech ws.
            sqlEQ = "SELECT [Serial Number] FROM [EQLogs$] WHERE ([TechNbr] = " & techNumber & " AND [Type] = '" & eqType & "')"
            Set rsEQ = Nothing
            Set rsEQ = New ADODB.Recordset
            rsEQ.Open sqlEQ, selfConn, adOpenStatic, adLockReadOnly
            Sheets(techNumber).Cells(5, pastecol).CopyFromRecordset rsEQ
            rsEQ.Close
        Next eqCell
    End If
Next techID

'Close ADODB connection
selfConn.Close

End Sub


