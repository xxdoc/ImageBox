VERSION 5.00
Begin VB.UserControl ToolButton 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   360
   ScaleHeight     =   24
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   24
End
Attribute VB_Name = "ToolButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'属性变量:
Dim m_Picture As Picture
'Dim m_DisabledPicture As Picture
'事件声明:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "当用户在一个对象上按下并释放鼠标按钮时发生。"

' 0: 抬起   1: 按下     2: 无效
Dim Mouse As Integer
Dim b(0 To 15, 0 To 15) As Long
Dim t As Long
Dim bb As Boolean

Private Sub Parser()
    On Error Resume Next
    Dim i As Integer, j As Integer
    UserControl.PaintPicture m_Picture, 0, 0
    
    t = Point(0, 0)
    For i = 0 To 15
        For j = 0 To 15
            b(i, j) = Point(i, j)
        Next
    Next
    
    bb = True
End Sub

Private Sub p(ByVal x As Integer)
    On Error Resume Next
    Dim i As Integer, j As Integer
    
    For i = 0 To 15
        For j = 0 To 15
            If b(i, j) <> t Then UserControl.PSet (i + x, j + x), b(i, j)
        Next
    Next
End Sub

Private Sub d(ByVal x As Integer)
    On Error Resume Next
    Dim i As Integer, j As Integer
    
    For i = 15 To 0 Step -1
        For j = 15 To 0 Step -1
            If b(i, j) <> t Then
                UserControl.PSet (i + x + 1, j + x + 1), vb3DHighlight
            End If
        Next
    Next
    For i = 15 To 0 Step -1
        For j = 15 To 0 Step -1
            Select Case b(i, j)
                Case t, vbWhite, &HC0C0C0
                
                Case Else
                    UserControl.PSet (i + x, j + x), vbButtonShadow
            End Select
        Next
    Next
End Sub

Public Sub Refresh()
    On Error Resume Next
    Dim i As Integer, j As Integer
    
    If bb = False Then Parser
    
    UserControl.Cls
    
    'UserControl.Line (0, 0)-(23, 23), vbButtonText, B
    
    UserControl.Line (1, 0)-(23, 0), vbButtotnText
    UserControl.Line (0, 1)-(0, 23), vbButtonText
    UserControl.Line (1, 23)-(23, 23), vbButtotnText
    UserControl.Line (23, 1)-(23, 23), vbButtonText
    
    Select Case Mouse
        Case 0
            UserControl.Line (1, 1)-(22, 1), vb3DHighlight
            UserControl.Line (1, 2)-(21, 2), vb3DHighlight
            UserControl.Line (1, 1)-(1, 22), vb3DHighlight
            UserControl.Line (2, 1)-(2, 21), vb3DHighlight
            
            UserControl.Line (22, 1)-(22, 22), vbButtonShadow
            UserControl.Line (21, 2)-(21, 22), vbButtonShadow
            UserControl.Line (2, 22)-(23, 22), vbButtonShadow
            UserControl.Line (3, 21)-(22, 21), vbButtonShadow
            
            'If Not m_Picture Is Nothing Then UserControl.PaintPicture m_Picture, 4, 4
            p 4
        Case 1
            UserControl.Line (1, 1)-(22, 1), vbButtonShadow
            UserControl.Line (1, 2)-(21, 2), vbButtonShadow
            UserControl.Line (1, 1)-(1, 22), vbButtonShadow
            UserControl.Line (2, 1)-(2, 21), vbButtonShadow
            
            UserControl.Line (22, 1)-(22, 22), vb3DHighlight
            UserControl.Line (21, 2)-(21, 22), vb3DHighlight
            UserControl.Line (2, 22)-(23, 22), vb3DHighlight
            UserControl.Line (3, 21)-(22, 21), vb3DHighlight
        
            'If Not m_Picture Is Nothing Then UserControl.PaintPicture m_Picture, 5, 5
            p 5
        Case 2
            UserControl.Line (1, 1)-(22, 1), vb3DHighlight
            UserControl.Line (1, 2)-(21, 2), vb3DHighlight
            UserControl.Line (1, 1)-(1, 22), vb3DHighlight
            UserControl.Line (2, 1)-(2, 21), vb3DHighlight
            
            UserControl.Line (22, 1)-(22, 22), vbButtonShadow
            UserControl.Line (21, 2)-(21, 22), vbButtonShadow
            UserControl.Line (2, 22)-(23, 22), vbButtonShadow
            UserControl.Line (3, 21)-(22, 21), vbButtonShadow
        
            
            'If Not m_DisabledPicture Is Nothing Then UserControl.PaintPicture m_DisabledPicture, 4, 4
            'UserControl.PaintPicture m_Picture, 4, 4, , , , , , , vbDstInvert
            d 4
    End Select
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "返回/设置一个值，决定一个对象是否响应用户生成事件。"
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    If New_Enabled = False Then Mouse = 2 Else Mouse = 0
    Refresh
    PropertyChanged "Enabled"
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

'注意！不要删除或修改下列被注释的行！
'MemberInfo=11,0,0,0
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "返回/设置控件中显示的图形。"
    Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set m_Picture = New_Picture
    bb = False
    UserControl.Cls
    Parser
    Refresh
    PropertyChanged "Picture"
End Property
'
''注意！不要删除或修改下列被注释的行！
''MemberInfo=11,0,0,0
'Public Property Get DisabledPicture() As Picture
'    Set DisabledPicture = m_DisabledPicture
'End Property
'
'Public Property Set DisabledPicture(ByVal New_DisabledPicture As Picture)
'    Set m_DisabledPicture = New_DisabledPicture
'    If Mouse = 2 Then Refresh
'    PropertyChanged "DisabledPicture"
'End Property

'为用户控件初始化属性
Private Sub UserControl_InitProperties()
    Set m_Picture = LoadPicture("")
'    Set m_DisabledPicture = LoadPicture("")
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Mouse = 2 Then Exit Sub
    If Button = 1 Then Mouse = 1: Refresh
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Mouse = 2 Or Button <> 1 Then Exit Sub
    If x < 0 Or Y < 0 Or x > 24 Or Y > 24 Then
        Mouse = 0
        Refresh
    ElseIf Mouse <> 1 Then
        Mouse = 1
        Refresh
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Mouse = 2 Then Exit Sub
    Mouse = 0
    Refresh
End Sub

'从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Set m_Picture = PropBag.ReadProperty("Picture", Nothing)
'    Set m_DisabledPicture = PropBag.ReadProperty("DisabledPicture", Nothing)
    Enabled = PropBag.ReadProperty("Enabled", True)
    
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    UserControl.Width = 360
    UserControl.Height = 360
    
    Refresh
End Sub

'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Picture", m_Picture, Nothing)
'    Call PropBag.WriteProperty("DisabledPicture", m_DisabledPicture, Nothing)
End Sub

