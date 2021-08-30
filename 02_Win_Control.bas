Attribute VB_Name = "02_Win_Control"
Option Compare Database:    Option Explicit

Public Function f_DeskTopPath() As Boolean
    DeskTopPath = ""
    Dim WSH As Object
    Set WSH = CreateObject("Wscript.Shell")
    DeskTopPath = WSH.SpecialFolders("Desktop")
    Set WSH = Nothing
End Function

Public Function f_CreateFolder(FolderName As String) As String
    Dim objFileSys As Object
    Dim strCreateFolder As String
    Dim tFolder As String
    
    Call f_DeskTopPath
    tFolder = DeskTopPath & "\" & FolderName
        
    Set objFileSys = CreateObject("Scripting.fileSystemobject")
    If objFileSys.FolderExists(tFolder) = False Then
            strCreateFolder = objFileSys.BuildPath(DeskTopPath & "\", FolderName)
            objFileSys.CreateFolder strCreateFolder
    End If
    f_CreateFolder = tFolder
    Set objFileSys = Nothing:  strCreateFolder = ""
End Function

Public Function f_DeleteFiles(fileName As String)
    If Len(Dir(fileName)) > 0 Then
        Kill fileName
    End If
End Function

Public Function f_File_PickUp(tTitle As String) As Boolean
    f_File_PickUp = False
    
    Dim MFile As MsoFileDialogType ' Requires Reference : Microsoft Office 15.0 Object Library
    Dim FDialog As FileDialog
    
    MFile = msoFileDialogFilePicker
    Set FDialog = FileDialog(MFile)
    With FDialog
        .AllowMultiSelect = False
        .Title = tTitle
        .ButtonName = "Import"
        '.InitialFIleName = DeskTopPath
    End With
        
    FDialog.Show
    If FDialog.SelectedItems.Count = 1 Then
        f_File_PickUp = True
        FilePath = FDialog.SelectedItems(1)
        Debug.Print FilePath
    End If

End Function

Public Function f_Get_FolderPath(tTitle As String) As Boolean
    f_Get_FolderPath = False
    
    Dim MFile As MsoFileDialogType ' Requires Reference : Microsoft Office 15.0 Object Library
    Dim FDialog As FileDialog
    
    MFile = msoFileDialogFolderPicker
    Set FDialog = FileDialog(MFile)
    With FDialog
        .AllowMultiSelect = False
        .Title = tTitle
        .Filters.Clear
        .ButtonName = "Select Folder"
        If .Show = True Then
                folderPath = .SelectedItems(1)
                f_Folder_Path = True
        Else
                folderPath = ""
                f_Get_FolderPath = False
        End If
    End With
End Function

Public Function f_List_Files(folderPath) As Boolean
        Dim objFs As Object, objItem As Object, objFiles As Object
        
        Set objFs = CreateObject("Scripting.FileSystemObject")
        Dim m As Integer, i As Integer
            m = objFs.GetFolder(fPath).subFolders.Count
            m = m + objFs.GetFolder(fPath).Files.Count
            ReDim arrayFiles(m - 1)
            i = 0
        For Each objItem In objFs.GetFolder(fPath).subFolders
                arrayFiles(i) = objItem.Path
                i = i + 1
        Next
        For Each objItem In objFs.GetFolder(fPath).Files
                arrayFiles(i) = objItem.Path
                i = i + 1
        Next
End Function

Public Function f_Kill_File(fileName As String)
    On Error Resume Next
    Kill fileName
End Function

Public Function f_FormtExcel(expFileB As String)
        Dim AppObj As Object    'Excel.Applicationオブジェクトの宣言
        Dim WBObj As Object     'Excel.Workbookオブジェクトの宣言
        Dim WsObj As Object      'Excel.WorkSheetオブジェクトの宣言
        Dim i As Integer
        Set AppObj = CreateObject("Excel.Application")  '実行時バインディング
        Set WBObj = AppObj.Workbooks.Open(expFileB)  'ワークブックを開く
        
        AppObj.Visible = False
        AppObj.DisplayAlerts = False
        
        For i = 1 To WBObj.Worksheets.Count
                Set WsObj = WBObj.Worksheets(i)
                    With WsObj
                            .Cells.Font.Name = "Arial Unicode MS"
                            .Cells.Font.Size = 9
                            .Rows(1).Interior.Color = RGB(183, 222, 232)
                    End With
        Next i
        
        WBObj.sheets(1).Select
        
        WBObj.Save   'ワークブックを保存する
        WBObj.Close 'ワークブックを閉じる
        AppObj.Quit
        i = 0
End Function
