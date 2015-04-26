VERSION 5.00
Object = "{797ED6C1-1DCC-489A-973F-BA2F31915A6C}#2.0#0"; "ImageControl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmLinearFunction 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "一次函数"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   Icon            =   "frmLinearFunction.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   392
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   299
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   257
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   960
      Width           =   3855
      Begin VB.TextBox txtY 
         Height          =   270
         Index           =   1
         Left            =   0
         TabIndex        =   11
         Top             =   300
         Width           =   1575
      End
      Begin VB.TextBox txtX 
         Height          =   270
         Index           =   1
         Left            =   1860
         TabIndex        =   13
         Top             =   300
         Width           =   1575
      End
      Begin VB.TextBox txtY 
         Height          =   270
         Index           =   0
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   1575
      End
      Begin VB.TextBox txtX 
         Height          =   270
         Index           =   0
         Left            =   1860
         TabIndex        =   9
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label lblE 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "＝"
         Height          =   180
         Index           =   1
         Left            =   1620
         TabIndex        =   12
         Top             =   330
         Width           =   180
      End
      Begin VB.Label lblKB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "k＋b"
         Height          =   180
         Index           =   1
         Left            =   3480
         TabIndex        =   14
         Top             =   330
         Width           =   360
      End
      Begin VB.Label lblE 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "＝"
         Height          =   180
         Index           =   0
         Left            =   1620
         TabIndex        =   8
         Top             =   30
         Width           =   180
      End
      Begin VB.Label lblKB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "k＋b"
         Height          =   180
         Index           =   0
         Left            =   3480
         TabIndex        =   10
         Top             =   30
         Width           =   360
      End
   End
   Begin VB.PictureBox picKB 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   120
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   257
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   3855
      Begin VB.TextBox txtB 
         Height          =   270
         Left            =   2220
         TabIndex        =   5
         Top             =   0
         Width           =   1575
      End
      Begin VB.TextBox txtK 
         Height          =   270
         Left            =   300
         TabIndex        =   3
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label lblX 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "x＋"
         Height          =   180
         Left            =   1920
         TabIndex        =   4
         Top             =   30
         Width           =   270
      End
      Begin VB.Label lblY 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "y＝"
         Height          =   180
         Left            =   0
         TabIndex        =   2
         Top             =   30
         Width           =   270
      End
   End
   Begin FunctionImage.ToolButton tbXY 
      Height          =   360
      Left            =   4020
      ToolTipText     =   "通过已知条件求得解析式"
      Top             =   1080
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   635
      Picture         =   "frmLinearFunction.frx":014A
   End
   Begin FunctionImage.ToolButton tbExpression 
      Height          =   360
      Left            =   4020
      ToolTipText     =   "直接输入解析式"
      Top             =   420
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   635
      Picture         =   "frmLinearFunction.frx":02A4
   End
   Begin TabDlg.SSTab sstFrame 
      Height          =   3795
      Left            =   120
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1980
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   6694
      _Version        =   393216
      Style           =   1
      TabHeight       =   529
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "分析"
      TabPicture(0)   =   "frmLinearFunction.frx":03FE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "求值"
      TabPicture(1)   =   "frmLinearFunction.frx":041A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtY(3)"
      Tab(1).Control(1)=   "txtX(3)"
      Tab(1).Control(2)=   "ToolButton3"
      Tab(1).Control(3)=   "txtD(1)"
      Tab(1).Control(4)=   "txtD(0)"
      Tab(1).Control(5)=   "txtY(2)"
      Tab(1).Control(6)=   "txtX(2)"
      Tab(1).Control(7)=   "Line1(2)"
      Tab(1).Control(8)=   "Label1(7)"
      Tab(1).Control(9)=   "Label12"
      Tab(1).Control(10)=   "Label11"
      Tab(1).Control(11)=   "Label10"
      Tab(1).Control(12)=   "Line1(1)"
      Tab(1).Control(13)=   "Label1(10)"
      Tab(1).Control(14)=   "Label1(9)"
      Tab(1).Control(15)=   "Label1(8)"
      Tab(1).Control(16)=   "Label1(6)"
      Tab(1).Control(17)=   "lblE(2)"
      Tab(1).Control(18)=   "lblKB(2)"
      Tab(1).ControlCount=   19
      TabCaption(2)   =   "图象"
      TabPicture(2)   =   "frmLinearFunction.frx":0436
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ibImage"
      Tab(2).ControlCount=   1
      Begin VB.TextBox txtY 
         Height          =   270
         Index           =   3
         Left            =   -72840
         TabIndex        =   31
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtX 
         Height          =   270
         Index           =   3
         Left            =   -74640
         TabIndex        =   29
         Top             =   1320
         Width           =   1575
      End
      Begin FunctionImage.ToolButton ToolButton3 
         Height          =   360
         Left            =   -71220
         Top             =   2940
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   635
         Picture         =   "frmLinearFunction.frx":0452
      End
      Begin VB.TextBox txtD 
         Height          =   270
         Index           =   1
         Left            =   -74520
         TabIndex        =   38
         Top             =   3360
         Width           =   555
      End
      Begin VB.TextBox txtD 
         Height          =   270
         Index           =   0
         Left            =   -74520
         TabIndex        =   36
         Top             =   2580
         Width           =   555
      End
      Begin VB.TextBox txtY 
         Height          =   270
         Index           =   2
         Left            =   -74820
         TabIndex        =   23
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtX 
         Height          =   270
         Index           =   2
         Left            =   -72960
         TabIndex        =   25
         Top             =   480
         Width           =   1575
      End
      Begin ImageControl.ImageBox ibImage 
         Height          =   3375
         Left            =   -74940
         TabIndex        =   39
         Top             =   360
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   5953
         ThetaMin        =   -3.142
         ThetaMax        =   3.142
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   -74880
         X2              =   -70920
         Y1              =   1980
         Y2              =   1980
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "点到直线的距离是 0"
         Height          =   180
         Index           =   7
         Left            =   -74820
         TabIndex        =   33
         Top             =   1680
         Width           =   3855
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "）"
         Height          =   180
         Left            =   -71220
         TabIndex        =   32
         Top             =   1380
         Width           =   180
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "，"
         Height          =   180
         Left            =   -73020
         TabIndex        =   30
         Top             =   1380
         Width           =   180
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "（"
         Height          =   180
         Left            =   -74820
         TabIndex        =   28
         Top             =   1380
         Width           =   180
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   -74880
         X2              =   -70920
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "（kx＋b）dx＝0"
         Height          =   180
         Index           =   10
         Left            =   -74340
         TabIndex        =   37
         Top             =   3000
         Width           =   1260
      End
      Begin VB.Label Label1 
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
         Index           =   9
         Left            =   -74820
         TabIndex        =   35
         Top             =   2640
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "y'＝k，即该函数代表的直线的斜率是 k，倾斜角是 atan(k)（弧度），（角度）"
         Height          =   360
         Index           =   8
         Left            =   -74820
         TabIndex        =   34
         Top             =   2100
         Width           =   3915
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "输入 y 或 x，即可求得另外一个值。"
         Height          =   180
         Index           =   6
         Left            =   -74820
         TabIndex        =   27
         Top             =   900
         Width           =   2970
      End
      Begin VB.Label lblE 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "＝"
         Height          =   180
         Index           =   2
         Left            =   -73200
         TabIndex        =   24
         Top             =   510
         Width           =   180
      End
      Begin VB.Label lblKB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "k＋b"
         Height          =   180
         Index           =   2
         Left            =   -71340
         TabIndex        =   26
         Top             =   510
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "该函数化为直线方程为 kx﹣y＋b＝0，化为法线式为（kx﹣y＋b）／（±√k^2＋1）"
         Height          =   360
         Index           =   5
         Left            =   180
         TabIndex        =   22
         Top             =   2460
         Width           =   3900
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "该函数的反函数是 x＝y／k﹣b／k"
         Height          =   180
         Index           =   4
         Left            =   180
         TabIndex        =   21
         Top             =   2100
         Width           =   3900
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "该函数的图象在 x 轴上的截距是 0，在 y 轴上的截距是 0"
         Height          =   360
         Index           =   3
         Left            =   180
         TabIndex        =   20
         Top             =   1560
         Width           =   3900
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "该函数是奇函数"
         Height          =   180
         Index           =   2
         Left            =   180
         TabIndex        =   19
         Top             =   1200
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "该函数在（﹣∞，﹢∞）内是单调增加的"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   18
         Top             =   840
         Width           =   3240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "定义域、值域：全体实数  （﹣∞，﹢∞）"
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   17
         Top             =   480
         Width           =   3420
      End
   End
   Begin VB.Label lblResult 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "求得的解析式是：y＝kx＋b"
      Height          =   180
      Left            =   120
      TabIndex        =   15
      Top             =   1680
      Width           =   2160
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   4
      X2              =   292
      Y1              =   56
      Y2              =   56
   End
   Begin VB.Label lblExpression 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "输入解析式(&E)：点击右面的按钮进行分析等操作。"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4050
   End
End
Attribute VB_Name = "frmLinearFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public K As Double, B As Double

Public Function GetExpression(ByVal K As Double, ByVal B As Double) As String
    On Error Resume Next
    If K = 0 Then
        If B < 0 Then
            GetExpression = "y＝﹣" & -B
        ElseIf B > 0 Then
            GetExpression = "y＝" & B
        Else
            GetExpression = "y＝0"
        End If
    ElseIf K < 0 Then
        If B < 0 Then
            GetExpression = "y＝﹣" & -K & "k﹣" & -B
        ElseIf B > 0 Then
            GetExpression = "y＝﹣" & -K & "k＋" & B
        Else
            GetExpression = "y＝﹣" & -K
        End If
    Else
        If B < 0 Then
            GetExpression = "y＝" & K & "k﹣" & -B
        ElseIf B > 0 Then
            GetExpression = "y＝" & K & "k＋" & B
        Else
            GetExpression = "y＝" & K
        End If
    End If
End Function
