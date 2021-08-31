Attribute VB_Name = "mod5_Analyze_Data"
Option Explicit

Sub ClearPivotTables()
'Deletes existing pivot tables from reportWB

Dim sht As Worksheet
Dim pvt As PivotTable

'Loop Through Each Pivot Table In ActiveWorkbook
  For Each sht In ActiveWorkbook.Worksheets
    For Each pvt In sht.PivotTables
      pvt.TableRange2.Clear
    Next pvt
  Next sht
  
End Sub

Sub PivotLastClosed()
'Summarizes Last Closed Issue Data in Pivot Table

Dim rngLastClosed As Range
Dim pcLastClosed As PivotCache
Dim ptLastClosed As PivotTable
Dim ptDestination As Range
       
    'Define Source Data Range
    wsLastClosed.Activate
    Set rngLastClosed = wsLastClosed.Range("tblLastClosed[#All]")
    
    'Create the cache from the source data
    Set pcLastClosed = ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=rngLastClosed _
        )
    
    'Set table destination range
    Set ptDestination = wsLastClosed.Range("Q2")
    
    'Create the Pivot table
    Set ptLastClosed = wsLastClosed.PivotTables.Add( _
        PivotCache:=pcLastClosed, _
        TableDestination:=ptDestination, _
        TableName:="Pivot2")

    ActiveWorkbook.ShowPivotTableFieldList = True
 
    'Adding fields
    With ptLastClosed
        
        With .PivotFields("STATUS")
             .Orientation = xlColumnField
             .Position = 1
        End With

        With .PivotFields("TechNbr")
             .Orientation = xlRowField
             .Position = 1
        End With
        
        With .PivotFields("STATUS")
             .Orientation = xlDataField
             .Position = 1
             .Caption = "TechTracker"
             .Function = xlCount
        End With

        'Adjusting some settings
        .RowGrand = True
        .DisplayFieldCaptions = True
        .HasAutoFormat = True
        .RowAxisLayout xlOutlineRow
        
    End With

    ActiveWorkbook.ShowPivotTableFieldList = False

End Sub

Sub PivotWeekNumber()
'Summarizes Invoice Week Data in Pivot Table

Dim rngWeekNumber As Range
Dim pcWeekNumber As PivotCache
Dim ptWeekNumber As PivotTable
Dim ptDestination As Range
       
    'Define Source Data Range
    Set wsWeekNumber = ActiveWorkbook.Sheets("Week_" & invoiceWeek)
    wsWeekNumber.Activate
    Set rngWeekNumber = wsWeekNumber.Range("tblWeekNumber[#All]")
    
    'Create the cache from the source data
    Set pcWeekNumber = ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=rngWeekNumber _
        )
    
    'Set table destination range
    Set ptDestination = wsWeekNumber.Range("Q2")
    
    'Create the Pivot table
    Set ptWeekNumber = wsWeekNumber.PivotTables.Add( _
        PivotCache:=pcWeekNumber, _
        TableDestination:=ptDestination, _
        TableName:="Pivot2")

    ActiveWorkbook.ShowPivotTableFieldList = True
 
    'Adding fields
    With ptWeekNumber
        
        With .PivotFields("DATE")
             .Orientation = xlColumnField
             .Position = 1
        End With

        With .PivotFields("TechNbr")
             .Orientation = xlRowField
             .Position = 1
        End With
        
        With .PivotFields("IssueGroup")
             .Orientation = xlRowField
             .Position = 2
        End With
        
        With .PivotFields("STATUS")
             .Orientation = xlRowField
             .Position = 3
        End With
        
        With .PivotFields("STATUS")
             .Orientation = xlDataField
             .Position = 1
             .Caption = "Summary"
             .Function = xlCount
        End With

        'Adjusting some settings
        .RowGrand = True
        .DisplayFieldCaptions = True
        .HasAutoFormat = True
        .RowAxisLayout xlOutlineRow
        
    End With

    ActiveWorkbook.ShowPivotTableFieldList = False

End Sub






