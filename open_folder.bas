Attribute VB_Name = "Module1"
Sub open_folder_test()
Call open_folder("C:\Users\user\Downloads\sample_folder")
End Sub
Sub open_folder(ByVal FilePath As String)
    If Dir(FilePath, vbDirectory) = "" Then
        MkDir FilePath
        MsgBox "指定したフォルダを作成しました"
    End If
    Shell "C:\Windows\explorer.exe " & FilePath, vbNormalFocus
End Sub
