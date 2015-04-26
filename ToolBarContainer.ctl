VERSION 5.00
Begin VB.UserControl ToolBarContainer 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5415
   ControlContainer=   -1  'True
   ScaleHeight     =   63
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   361
End
Attribute VB_Name = "ToolBarContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub Refresh()
    On Error Resume Next
    UserControl.Cls
    UserControl.Line (0, 0)-(UserControl.ScaleWidth, 0), vb3DHighlight
    UserControl.Line (0, 0)-(0, UserControl.ScaleHeight), vb3DHighlight
    UserControl.Line (UserControl.ScaleWidth - 1, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight), vbButtonShadow
    UserControl.Line (0, UserControl.ScaleHeight - 1)-(UserControl.ScaleWidth, UserControl.ScaleHeight - 1), vbButtonShadow
End Sub

Private Sub UserControl_Resize()
    Refresh
End Sub
