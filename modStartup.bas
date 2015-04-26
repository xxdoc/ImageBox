Attribute VB_Name = "modStartup"

Public Function AppPath() As String
    On Error Resume Next
    AppPath = App.Path
    If Right(AppPath, 1) <> "\" Then AppPath = AppPath & "\"
End Function

Public Sub Main()
    On Error Resume Next
    'Set ImageWindows = New Collection
    
    'AddWindow
    
    frmMDI.Show
End Sub
