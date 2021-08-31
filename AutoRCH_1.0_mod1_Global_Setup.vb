Attribute VB_Name = "mod1_Global_Setup"
Option Explicit
'
' mod1_Global_Setup
'

' *** PROJECT-LEVEL DECLARATIONS ***

'Workbook Objects
Public wbReport As Workbook

'Worksheet Objects
Public wsTracking As Worksheet
Public wsLastClosed As Worksheet
Public wsWeekNumber As Worksheet
Public wsPool As Worksheet
Public wsPayroll As Worksheet
Public wsReports As Worksheet
Public wsEQ As Worksheet
Public wsReference As Worksheet
Public wsTech As Worksheet

'Range Objects
Public techIDRange As Range
Public techID As Range
Public eqRange As Range
Public eqCell As Range
Public dateRange As Range
Public dateCell As Range

'Tables/List Objects
Public tblTracking As ListObject
Public tblWeekNumber As ListObject
Public tblLastClosed As ListObject
Public tblEQ As ListObject
Public tblProduction As ListObject

'Connection Objects
Public selfConn As ADODB.Connection
Public connString As String
Public selfPath As String

'String Variables
Public invoiceWeek As String
Public techNumber As String
Public eqType As String

'Date Variables
Public dateLastClosed As Date
Public dateIssue1 As Date
Public dateIssue2 As Date
Public dateIssue3 As Date
Public dateMon As Date
Public dateTues As Date
Public dateWeds As Date
Public dateThurs As Date
Public dateFri As Date
Public dateSat As Date
Public dateSun As Date

'Numeric Variables
Public lastRow As Long
Public pasteRow As Long
Public pastecol As Long

Sub setGlobalSettings()

    Application.DisplayAlerts = False
    'Application.ScreenUpdating = False
    'Application.Calculate=Manual (check syntax)
    'Application.Events=Disabled (check syntax)
    
End Sub

Sub setStaticVariables()

'UserInput
    invoiceWeek = wsUserInput.Range("F4").Value
    Set techIDRange = wsUserInput.Range("F5:F24")

'Workbooks
    Set wbReport = ThisWorkbook

'Worksheets
    Set wsTracking = Sheets("Tracking")
    Set wsPayroll = Sheets("Payroll")
    Set wsReports = Sheets("Reports")
    Set wsEQ = Sheets("EQLogs")
    Set wsReference = Sheets("Reference")
    
'Dates
    dateIssue1 = wsReference.Range("E4").Value
    dateIssue2 = wsReference.Range("E5").Value
    dateIssue3 = wsReference.Range("E6").Value
    dateLastClosed = wsReference.Range("E7").Value
    dateMon = wsReference.Range("E8").Value
    dateTues = wsReference.Range("E9").Value
    dateWeds = wsReference.Range("E10").Value
    dateThurs = wsReference.Range("E11").Value
    dateFri = wsReference.Range("E12").Value
    dateSat = wsReference.Range("E13").Value
    dateSun = wsReference.Range("E14").Value
       
End Sub

Sub SaveFiles()
'Saves AutoRCH Macro Workbook before making changes
'Backs-up user-provided data before making changes
'Saves Report workbook to continue macro

'Find path to directory where wbReport is saved
    Dim directoryPath As String
    Dim selfPath As String
    Dim dataPath As String
    Dim reportPath As String
    
    directoryPath = Application.ThisWorkbook.Path
    
    selfPath = ThisWorkbook.FullName
    dataPath = directoryPath & "\Data_Backup\Week_" & invoiceWeek & "_Source_Data_" & _
        Format(Now(), "yyyy-mm-dd_hhmm_AMPM") & ".xlsm"
    reportPath = directoryPath & "\Reports\Week_" & invoiceWeek & "_Reports_" & _
        Format(Now(), "yyyy-mm-dd_hhmm_AMPM") & ".xlsm"

    wbReport.SaveAs Filename:=selfPath
    wbReport.SaveAs Filename:=dataPath
    wbReport.SaveAs Filename:=reportPath

End Sub

Sub DisplayWorksheets()

    wsPayroll.Visible = True
    wsReports.Visible = True
    wsTracking.Visible = True
    
End Sub

Sub ConnectToSelf()
' Connects to this workbook as "external" data source for use with SQL queries
    wbReport.Save
        
    Set selfConn = New ADODB.Connection
    selfPath = ThisWorkbook.FullName
    
    connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & selfPath _
        & ";Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1"";"
    
    If selfConn.State = adStateOpen Then selfConn.Close

    selfConn.Open connString

End Sub

