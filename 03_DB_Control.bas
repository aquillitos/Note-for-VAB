Attribute VB_Name = "03_DB_Control"
Option Compare Database:    Option Explicit


Public Function s_Connect_Server() As Boolean
    Dim e As Integer:   e = 0
    s_Connect_Server = False
    If s_Local_Server_Connect = False Then e = e + 1
    If s_System_Server_Connect = False Then e = e + 1
    If s_Database_Server_Connect = False Then e = e + 1
    If e = 0 Then
            s_Connect_Server = True
    Else
            s_Connect_Server = False
    End If
    e = 0
End Function

Public Function f_PT_Query(stSQL As String, queryNm As String, ReturnRec As Boolean) As Boolean
        'Pass Through Query - to run query on the server side and return result.
        f_PT_Query = False
        Dim dbc As DAO.Database
        Dim queryDf As DAO.QueryDef
        Dim PT_Con As String
        On Error GoTo Er
        
        If s_System_Server_Connect = True Then
                Set dbc = CurrentDb
                PT_Con = "ODBC;DSN=BM System;UID=KOSTF8Y;Trusted_Connection=Yes;DATABASE=BM_System;"    'Const Value
                
                Call f_Delete_LocalQuery(queryNm)
                Set queryDf = dbc.CreateQueryDef(queryNm)
                    With queryDf
                        .Connect = PT_Con
                        .ReturnsRecords = ReturnRec
                        .SQL = stSQL
                        .Close
                    End With
                Set queryDf = Nothing
                Set dbc = Nothing
                f_PT_Query = True
                Exit Function
        Else
Er:
                Set queryDf = Nothing
                Set dbc = Nothing
                f_PT_Query = False
        End If

End Function

Private Function s_Server_Connect(conADOString As String) As Boolean
        '== General Connection Function ==================
        s_Server_Connect = False
        On Error GoTo ErrSub
                Set C_Con = New ADODB.Connection
                C_Con.ConnectionString = conADOString
                C_Con.Open
        On Error GoTo 0
        s_Server_Connect = True
        Exit Function
ErrSub:
        Call s_Server_Connect_Fail
        s_Server_Connect = False
End Function

Private Function s_Local_Server_Connect() As Boolean   'Access Database
    s_Local_Server_Connect = False
    On Error GoTo ErrSub
        Set C_ConLo = CurrentProject.Connection
    On Error GoTo 0
    s_Local_Server_Connect = True
    Exit Function
ErrSub:
    Call s_Server_Connect_Fail
    s_Local_Server_Connect = False
End Function

    Private Function s_System_Server_Connect() As Boolean  'This System's SQL server
    s_System_Server_Connect = False
    On Error GoTo ErrSub
        Set C_ConSys = New ADODB.Connection
        C_ConSys.ConnectionString = C_ConADOSys
        C_ConSys.Open
    On Error GoTo 0
    s_System_Server_Connect = True
    Exit Function
ErrSub:
    Call s_Server_Connect_Fail
    s_System_Server_Connect = False
End Function

Private Function s_Database_Server_Connect() As Boolean  'This System's SQL server
    s_Database_Server_Connect = False
    On Error GoTo ErrSub
        Set C_ConDat = New ADODB.Connection
        C_ConDat.ConnectionString = C_ConADODat
        C_ConDat.Open
    On Error GoTo 0
    s_Database_Server_Connect = True
    Exit Function
ErrSub:
    Call s_Server_Connect_Fail
    s_Database_Server_Connect = False
End Function

Public Function s_Server_Connect_Fail()
    ANS = MsgBox("Server Connection Failed." & vbCrLf & "Please check the Server Status and the Connection String then try again", vbCritical + vbOKOnly, C_msgTitle)
End Function

Public Function f_Copy_TableLL(FromTable, ToTable) As Boolean 'Local to Local
    On Error Resume Next: C_ConLo.Execute "SELECT * INTO [" & ToTable & "] FROM [" & FromTable & "];":  On Error GoTo 0
End Function

Public Function f_Copy_TableLS(FromTable, ToTable) As Boolean 'Local to System Database
    On Error Resume Next:   C_ConLo.Execute "SELECT * INTO " & C_ConSQLSys & ".[" & ToTable & "] FROM " & FromTable:  On Error GoTo 0
End Function

Public Function f_Copy_TableSL(FromTable, ToTable) As Boolean 'System Database to LocaL
    f_Copy_TableSL = False
    Call f_Delete_LocalTable(ToTable)
    On Error GoTo Er
    C_ConLo.Execute "SELECT * INTO [" & ToTable & "] FROM " & C_ConSQLSys & ".[" & FromTable & "]"
    f_Copy_TableSL = True
    Exit Function
Er:
    f_Copy_TableSL = False
    On Error GoTo 0
End Function

Public Function f_Copy_TableDL(FromTable, ToTable) As Boolean 'Database Database to LocaL
    f_Copy_TableDL = False
    Call f_Delete_LocalTable(ToTable)
    On Error GoTo Er
    C_ConLo.Execute "SELECT * INTO [" & ToTable & "] FROM " & C_ConSQLDat & ".[" & FromTable & "]"
    f_Copy_TableDL = True
    Exit Function
Er:
    f_Copy_TableDL = False
    On Error GoTo 0
End Function

Public Function f_Copy_TableSS(FromTable, ToTable) As Boolean
    f_Copy_TableSS = False
    On Error GoTo Er
    C_ConSys.Execute "SELECT * INTO [" & ToTable & "] FROM [" & FromTable & "]"
    f_Copy_TableSS = True
    Exit Function
Er:
    f_Copy_TableSS = False
    On Error GoTo 0
End Function

Public Function f_Exist_LocalTable(tableName) As Boolean
   On Error Resume Next:    f_Exist_LocalTable = CurrentDb.TableDefs(tableName).Name = tableName:   On Error GoTo 0
End Function

Public Function f_Delete_LocalTable(tableName) As Boolean
    On Error Resume Next:   DoCmd.Close acTable, tableName: DoCmd.DeleteObject acTable, tableName:  On Error GoTo 0
End Function

Public Function f_Exist_LocalQuery(queryName) As Boolean
    On Error Resume Next:   f_Exist_Query = CurrentDb.QueryDefs(queryName).Name = queryName:    On Error GoTo 0
End Function

Public Function f_Delete_LocalQuery(queryName) As Boolean
    On Error Resume Next:   DoCmd.Close acQuery, queryName: DoCmd.DeleteObject acQuery, queryName:  On Error GoTo 0
End Function

Public Function f_Delete_LocalTables_Like(tableName As String) As Boolean
    Dim tTable As DAO.TableDef
    On Error Resume Next
            For Each tTable In db.TableDefs
                If tTable.Name Like tableName Then
                        DoCmd.Close acTable, tTable.Name
                        DoCmd.DeleteObject acTable, tTable.Name
                End If
            Next
    On Error GoTo 0
End Function

Public Function f_Delete_LocalQueries_Like(queryName As String) As Boolean
        Dim queryTable As DAO.QueryDef
        On Error Resume Next
                For Each queryTable In db.QueryDefs
                        If queryTable.Name Like queryName Then
                            DoCmd.Close acQuery, queryTable.Name
                            DoCmd.DeleteObject acQuery, queryTable.Name
                        End If
                Next
        On Error GoTo 0
End Function

Public Function f_Drop_System_Table(tableName) As Boolean
    Dim rsD As ADODB.Recordset
    Set rsD = New ADODB.Recordset
    f_Drop_System_Table = False
    On Error GoTo Er
    rsD.Open "DROP TABLE " & tableName & ";", C_ConSys:
    Set rsD = Nothing
    f_Drop_System_Table = True
    Exit Function
Er:
    f_Drop_System_Table = False
    On Error GoTo 0
End Function

Public Function f_Update_System_Data(tableName) As Boolean
    f_Update_System_Data = False
    On Error GoTo Er
    C_ConLo.Execute "INSERT INTO " & C_ConSQLSys & ".[" & tableName & "] SELECT * FROM [" & tableName & "] WHERE Month = #" & Format(ciMonth, "yyyy/mm/dd") & "#;"
    f_Update_System_Data = True
    On Error GoTo 0
    Exit Function
Er:
    f_Update_System_Data = False
    On Error GoTo 0
End Function

Public Function f_Delete_System_Data(tableName) As Boolean
    f_Delete_System_Data = False
    On Error GoTo Er
    C_ConSys.Execute "DELETE FROM [" & tableName & "] WHERE Month = '" & Format(ciMonth, "yyyy/mm/dd") & "';"
    f_Delete_System_Data = True
    On Error GoTo 0
    Exit Function
Er:
    f_Delete_System_Data = False = False
    On Error GoTo 0
End Function

Public Function f_RunQuery(stSQL As String) As Boolean
    f_RunQuery = False
    On Error GoTo Er
        DoCmd.SetWarnings False
            Debug.Print stSQL
            DoCmd.RunSQL stSQL
            stSQL = ""
            f_RunQuery = True
        DoCmd.SetWarnings True
        Exit Function
Er:
    f_RunQuery = False
    stSQL = ""
    DoCmd.SetWarnings True
End Function

