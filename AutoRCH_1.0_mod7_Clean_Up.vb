Attribute VB_Name = "mod7_Clean_Up"
Option Explicit
Sub setDisplay()

'Put worksheets in order
    wsWeekNumber.Move After:=wsReports
    wsLastClosed.Move After:=wsWeekNumber
    wsPool.Move After:=wsReference
    
'Delete extra worksheets
    wsUserInput.Delete
    wsTracking.Delete
    wsEQ.Delete
    wsReference.Visible = xlSheetVisible
    wsReference.Delete
    
'Activate Payroll Sheet
    wsPayroll.Activate
    wsPayroll.Range("A1").Select

End Sub

Sub SaveReport()
'Save report to Reports Folder
  
'Find path to directory where wbReport is saved
    
    wbReport.Save
    
End Sub

Sub CloseConnections()

'Connection to Self

If selfConn.State = adStateOpen Then selfConn.Close
Set selfConn = Nothing

End Sub

Sub DestroyVariables()


End Sub

Sub unsetGlobalSettings()

    Application.DisplayAlerts = True
    'Application.ScreenUpdating = True

End Sub


Sub msgAllDone()

MsgBox ("ALL DONE! I hope this worked!!")

End Sub



