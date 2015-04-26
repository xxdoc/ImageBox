Attribute VB_Name = "modImageWindow"
'Global ImageWindows As Collection
'
'Public Function AddWindow(Optional ByVal k As Integer) As frmImageWindow
'    On Error Resume Next
'    Dim i As Integer, f As frmImageWindow
'
'    If IsMissing(k) Then i = ImageWindows.Count Else i = k
'
'    Set f = New frmImageWindow
'
'    f.Caption = "Í¼Ïó´°¿Ú #" & i
'    f.Tag = i
'    f.Show
'
'    Set AddWindow = f
'
'    ImageWindows.Add f, CStr(i)
'
'    Set f = Nothing
'End Function
'
'Public Sub DeleteWindow(ByVal i As Integer)
'    On Error Resume Next
'    ImageWindows.Remove CStr(i)
'End Sub
'
'Public Property Get ImageWindow(ByVal i As Integer) As frmImageWindow
'    On Error Resume Next
'    If ImageWindows(CStr(i)) Is Nothing Then AddWindow i
'
'    Set ImageWindow = ImageWindows(CStr(i))
'
'End Property
