Attribute VB_Name = "main_MasterMacro"
Option Explicit

Sub MasterMacro()

    'mod1_Global_Settings
        progress (1)
        Call setGlobalSettings
        Call setStaticVariables
        Call SaveFiles
        Call DisplayWorksheets
    'mod2_Format_Data
        progress (2)
        Call CombineTracking
        Call FormatTracking
        Call AddGroupColumn
        Call FormatEQ
        progress (3)
        Call FormatNumbers
    'mod3_Organize_Data
        progress (4)
        Call addWeekNumberWS
        Call addLastClosedWS
        Call addPoolWS
        Call addTechWorksheets
        Call SortTechEQ
        Call CreateTechTables
    'mod4_Summarize_Data
        progress (5)
        Call ListPoolSaves
        Call ListJobsByDate
        Call ListEQByType
    'mod5_Analyze_Data
        progress (6)
        Call ClearPivotTables
        Call PivotWeekNumber
        Call PivotLastClosed
    'mod6_Format_Reports
        progress (7)
        Call FormatPayroll
        Call PoolPayroll
        Call TechPayroll
        Call Production
    'mod7_Clean_Up
        progress (8)
        Call CloseConnections
        'Call DestroyVariables
        Call setDisplay
        Call SaveReport
        Call unsetGlobalSettings
    'Grand Finale
        Unload ProgressBar
        Call msgAllDone
End Sub

Sub progress(step As Single)

Dim step1 As String
Dim step2 As String
Dim step3 As String
Dim step4 As String
Dim step5 As String
Dim step6 As String
Dim step7 As String
Dim step8 As String

step1 = "Initializing Workbook..."
step2 = "Formatting Source Data..."
step3 = "FIXING MISMATCH BUG!!!"
step4 = "Sorting Source Data..."
step5 = "Summarizing Data..."
step6 = "Creating Pivot Tables..."
step7 = "Formatting Reports..."
step8 = "Saving Reports..."

ProgressBar.Bar.Width = step * 24

If step = 1 Then
        ProgressBar.Text.Caption = step1
    ElseIf step = 2 Then
        ProgressBar.Text.Caption = step2
    ElseIf step = 3 Then
        ProgressBar.Text.Caption = step3
    ElseIf step = 4 Then
        ProgressBar.Text.Caption = step4
    ElseIf step = 5 Then
        ProgressBar.Text.Caption = step5
    ElseIf step = 6 Then
        ProgressBar.Text.Caption = step6
    ElseIf step = 7 Then
        ProgressBar.Text.Caption = step7
    ElseIf step = 8 Then
        ProgressBar.Text.Caption = step8
End If

DoEvents

End Sub

