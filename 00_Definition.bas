Attribute VB_Name = "00_Definition"
Option Compare Database:    Option Explicit
    Public db As DAO.Database
    Public Form01 As String
    
    '=== SQL Server ====================
    Public C_ConSQLSys As String
    Public C_ConADOSys As String
    Public C_ConLo As ADODB.Connection  'Constant: Requires "Microsoft ActiveX Data Object X.X Library"
    Public C_ConSys As ADODB.Connection  'Constant: Requires "Microsoft ActiveX Data Object X.X Library"
    Public C_msgTitle As String
    
    '=== Folders =======================
    Public DeskTopPath As String
    Public tFolder As String
    Public FilePath As String
    Public saveFolder As String

    '=== Error & Message ================
    Public ANS As String

Public Function f_Initial()
    Set db = CurrentDb()
    Form01 = "Form1"
End Function

