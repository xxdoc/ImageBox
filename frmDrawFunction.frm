VERSION 5.00
Begin VB.Form frmDrawFunction 
   Caption         =   "绘制任意函数"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4980
   Icon            =   "frmDrawFunction.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   207
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   332
   Begin VB.CheckBox chkDraw 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4620
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   120
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   "求值(&V)"
      Height          =   2475
      Left            =   120
      TabIndex        =   3
      Top             =   540
      Width           =   4755
      Begin VB.CommandButton cmdF4 
         Caption         =   "确定"
         Height          =   255
         Left            =   4020
         TabIndex        =   20
         Top             =   2040
         Width           =   615
      End
      Begin VB.CommandButton cmdF3 
         Caption         =   "确定"
         Height          =   255
         Left            =   4020
         TabIndex        =   19
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton cmdF2 
         Caption         =   "确定"
         Height          =   255
         Left            =   4020
         TabIndex        =   18
         Top             =   660
         Width           =   615
      End
      Begin VB.CommandButton cmdF1 
         Caption         =   "确定"
         Height          =   255
         Left            =   4020
         TabIndex        =   17
         Top             =   300
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   420
         TabIndex        =   15
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox txtA 
         Height          =   270
         Left            =   420
         TabIndex        =   12
         Top             =   1680
         Width           =   555
      End
      Begin VB.TextBox txtB 
         Height          =   270
         Left            =   420
         TabIndex        =   11
         Top             =   1020
         Width           =   555
      End
      Begin VB.TextBox txtF2 
         Height          =   270
         Left            =   420
         TabIndex        =   8
         Top             =   660
         Width           =   1455
      End
      Begin VB.TextBox txtF1 
         Height          =   270
         Left            =   420
         TabIndex        =   5
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label lblR4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "时，f(x)→？"
         Height          =   180
         Left            =   1920
         TabIndex        =   16
         Top             =   2100
         Width           =   1080
      End
      Begin VB.Label lblF4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "x→"
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   2100
         Width           =   270
      End
      Begin VB.Label lblR3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "f(x)dx＝？"
         Height          =   180
         Left            =   660
         TabIndex        =   13
         Top             =   1380
         Width           =   900
      End
      Begin VB.Label lblF3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "∫"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   42
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   225
      End
      Begin VB.Label lblR2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ")＝？"
         Height          =   180
         Left            =   1920
         TabIndex        =   9
         Top             =   705
         Width           =   450
      End
      Begin VB.Label lblF2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "f'("
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   705
         Width           =   270
      End
      Begin VB.Label lblR1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ")＝？"
         Height          =   180
         Left            =   1920
         TabIndex        =   6
         Top             =   345
         Width           =   450
      End
      Begin VB.Label lblF1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "f("
         Height          =   180
         Left            =   120
         TabIndex        =   4
         Top             =   345
         Width           =   180
      End
   End
   Begin VB.CommandButton cmdDraw 
      Caption         =   "绘制(&D)"
      Default         =   -1  'True
      Height          =   315
      Left            =   3720
      TabIndex        =   2
      Top             =   120
      Width           =   915
   End
   Begin VB.TextBox txtExpression 
      Height          =   270
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label lblFunction 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "f(x)＝"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   540
   End
End
Attribute VB_Name = "frmDrawFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDraw_Click()
    On Error Resume Next
    'ImageWindow(0).ibImage.Functions.Add "k" & Rnd, txtExpression.Text
    'ImageWindow(0).ibImage.Refresh
    'ImageWindow(0).SetFocus
End Sub

Private Sub cmdF1_Click()
    On Error Resume Next
    Dim b As Boolean, d As Double
    d = CalculateString(txtExpression.Text, b, Val(txtF1.Text), "X")
    If b Then
        lblR1.Caption = ")＝" & d
    Else
        lblR1.Caption = ")＝错误"
        txtF1.SetFocus
        txtF1.SelStart = 0
        txtF1.SelLength = Len(txtF1.Text)
    End If
End Sub
