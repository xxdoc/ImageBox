VERSION 5.00
Object = "{797ED6C1-1DCC-489A-973F-BA2F31915A6C}#2.0#0"; "ImageControl.ocx"
Begin VB.Form frmImageWindow 
   Caption         =   "Í¼Ïó´°¿Ú"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmImageWindow.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin ImageControl.ImageBox ibImage 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   5530
      ThetaMin        =   -3.142
      ThetaMax        =   3.142
   End
End
Attribute VB_Name = "frmImageWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    On Error Resume Next
    ibImage.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    'DeleteWindow CInt(Me.Tag)
End Sub
