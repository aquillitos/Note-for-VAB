Attribute VB_Name = "01_App_Control"
Option Compare Database:    Option Explicit

Public Function f_RunQuery(stSQL As String) As Boolean
    f_RunQuery = False
    On Error GoTo e
        DoCmd.SetWarnings False
        Debug.Print stSQL
        DoCmd.RunSQL stSQL: stSQL = ""
        DoCmd.SetWarnings True
        f_RunQuery = True
        DoCmd.SetWarnings True
        Exit Function
e:
    f_RunQuery = False
    DoCmd.SetWarnings True
End Function
