Attribute VB_Name = "modCalculator"
Option Explicit

Const PI As Double = 3.14159265358979               ' 常数π

Private tmpChar As String

' 对字符串进行计算。
' 每个函数中的 Expression 表示需要计算的表达式，
' 参数 IsValid 如果为 True，则表示表达式有效，返回值是计算后的结果，
' 参数 XValue 表示如果表达式中出现了“x”，就把 XValue 的值代入到 x 中。
Public Function CalculateString(ByVal Expression As String, ByRef IsValid As Boolean, Optional XValue As Double = 0, Optional XLetter As String = "X") As Double
    ' 返回表达式的总长度。
    Dim ExpressionLength As Long
    ' 返回计算完成之后分析的位置。
    Dim ParserPosition As Integer
    
    If XLetter = "" Then tmpChar = "X" Else tmpChar = Left(UCase(XLetter), 1)
    
    ParserPosition = 1
    ' 首先要把圆周率替换成数值，
    ' 由于是按值传递，Expression 是不会被更改的。
    Expression = Replace(Expression, "pi", CStr(PI), , , vbTextCompare)
    ExpressionLength = Len(Expression)
    ' 由于加减法是最后被运算的，所以加减法算完了，式子也就算完了。
    CalculateString = AddMinusParser(UCase(Expression), IsValid, ParserPosition, XValue)
    ' 如果表达式没有被分析完成，就说明表达式有问题。
    If ParserPosition <= ExpressionLength Then IsValid = False
End Function

' 进行加减法的运算。
' Position 表示已经分析到的位置。
Private Function AddMinusParser(ByVal Expression As String, ByRef IsValid As Boolean, ByRef Position As Integer, ByVal XValue As Double) As Double
    Dim Tmp1 As Double, Tmp2 As Double
    On Error GoTo ErrorHandler
    ' 由于乘除法较加减法优先，所以先计算乘除法。
    ' 这种算法是将 1 * 2 / 3 ^ 4 + 2 * 3 / 4 ^ 5 变成 (1 * 2 / 3 ^ 4) + (2 * 3 / 4 ^ 5)
    Tmp1 = MulDivParser(Expression, IsValid, Position, XValue)
    If Match(Expression, "+", Position) Then
        Position = Position + 1
        ' 在第二操作数中再次进行运算。
        Tmp2 = AddMinusParser2(Expression, IsValid, Position, XValue, False)
        If IsValid = False Then Exit Function
        AddMinusParser = Tmp1 + Tmp2
    ElseIf Match(Expression, "-", Position) Then
        Position = Position + 1
        ' 在第二操作数中再次进行运算。
        '*此处出现问题：如果计算属于 1 - 2 + 3 的形式，函数会将其处理成 1 - (2 + 3)。
        ' 问题已经得到解决。
        Tmp2 = AddMinusParser2(Expression, IsValid, Position, XValue, True)
        If IsValid = False Then Exit Function
        AddMinusParser = Tmp1 - Tmp2
    Else
        ' 应该是算完了。
        AddMinusParser = Tmp1
    End If
    Exit Function
ErrorHandler:
    IsValid = False
End Function

' 在第二操作数中再次进行加减法运算。
' SignReverse 表示将每个操作数的符号变一下，
' 目的是解决 1 - 2 + 3 = 1 - (2 + 3) 的问题。
Private Function AddMinusParser2(ByVal Expression As String, ByRef IsValid As Boolean, ByRef Position As Integer, ByVal XValue As Double, ByVal SignReverse As Boolean) As Double
    Dim Tmp1 As Double, Tmp2 As Double
    ' 在第二操作数中再次寻找第一操作数。
    Tmp1 = MulDivParser2(Expression, IsValid, Position, XValue, False)
    If Match(Expression, "+", Position) Then
        Position = Position + 1
        Tmp2 = AddMinusParser2(Expression, IsValid, Position, XValue, False)
        If IsValid = False Then Exit Function
        ' 解决 1 - 2 + 3 = 1 - (2 + 3) 的问题。
        If SignReverse Then AddMinusParser2 = Tmp1 - Tmp2 Else AddMinusParser2 = Tmp1 + Tmp2
    ElseIf Match(Expression, "-", Position) Then
        Position = Position + 1
        Tmp2 = AddMinusParser2(Expression, IsValid, Position, XValue, True)
        If IsValid = False Then Exit Function
        ' 解决 1 - 2 + 3 = 1 - (2 + 3) 的问题。
        If SignReverse Then AddMinusParser2 = Tmp1 + Tmp2 Else AddMinusParser2 = Tmp1 - Tmp2
    Else
        AddMinusParser2 = Tmp1
    End If
End Function

' 进行乘除法的运算。
Private Function MulDivParser(ByVal Expression As String, ByRef IsValid As Boolean, ByRef Position As Integer, ByVal XValue As Double) As Double
    Dim Tmp1 As Double, Tmp2 As Double
    ' 先计算乘方。
    Tmp1 = PowerParser(Expression, IsValid, Position, XValue)
    If IsValid = False Then Exit Function
    If Match(Expression, "*", Position) Then
        Position = Position + 1
        ' 计算第二操作数。
        Tmp2 = MulDivParser2(Expression, IsValid, Position, XValue, False)
        If IsValid = False Then Exit Function
        MulDivParser = Tmp1 * Tmp2
    ElseIf Match(Expression, "/", Position) Then
        Position = Position + 1
        ' 计算第二操作数。
        Tmp2 = MulDivParser2(Expression, IsValid, Position, XValue, True)
        If IsValid = False Then Exit Function
        MulDivParser = Tmp1 / Tmp2
    Else
        MulDivParser = Tmp1
    End If
End Function

' 在第二操作数中再次进行乘除法运算。
' SignReverse 表示将每个操作数的符号变一下，
' 目的是解决 1 / 2 * 3 = 1 / (2 * 3) 的问题。
Private Function MulDivParser2(ByVal Expression As String, ByRef IsValid As Boolean, ByRef Position As Integer, ByVal XValue As Double, ByVal SignReverse As Boolean) As Double
    Dim Tmp1 As Double, Tmp2 As Double
    ' 先计算乘方。
    Tmp1 = PowerParser2(Expression, IsValid, Position, XValue)
    If IsValid = False Then Exit Function
    If Match(Expression, "*", Position) Then
        Position = Position + 1
        ' 计算第二操作数。
        Tmp2 = MulDivParser2(Expression, IsValid, Position, XValue, False)
        If IsValid = False Then Exit Function
        If SignReverse = True Then MulDivParser2 = Tmp1 / Tmp2 Else MulDivParser2 = Tmp1 * Tmp2
    ElseIf Match(Expression, "/", Position) Then
        Position = Position + 1
        ' 计算第二操作数。
        Tmp2 = MulDivParser2(Expression, IsValid, Position, XValue, True)
        If IsValid = False Then Exit Function
        If SignReverse = True Then MulDivParser2 = Tmp1 * Tmp2 Else MulDivParser2 = Tmp1 / Tmp2
    Else
        MulDivParser2 = Tmp1
    End If
End Function

' 进行乘方运算。
Private Function PowerParser(ByVal Expression As String, ByRef IsValid As Boolean, ByRef Position As Integer, ByVal XValue As Double) As Double
    Dim Tmp1 As Double, Tmp2 As Double
    ' 先计算函数值。
    Tmp1 = FunctionCalc(Expression, IsValid, Position, XValue)
    If IsValid = False Then Exit Function
    If Match(Expression, "^", Position) Then
        Position = Position + 1
        ' 计算第二操作数。
        Tmp2 = FunctionCalc2(Expression, IsValid, Position, XValue)
        If IsValid = False Then Exit Function
        PowerParser = Tmp1 ^ Tmp2
    Else
        PowerParser = Tmp1
    End If
End Function

' 在第二操作数中进行乘方运算。
Private Function PowerParser2(ByVal Expression As String, ByRef IsValid As Boolean, ByRef Position As Integer, ByVal XValue As Double) As Double
    Dim Tmp1 As Double, Tmp2 As Double
    ' 计算第一操作数。
    Tmp1 = FunctionCalc2(Expression, IsValid, Position, XValue)
    If IsValid = False Then Exit Function
    If Match(Expression, "^", Position) Then
        Position = Position + 1
        ' 计算第二操作数。
        Tmp2 = FunctionCalc2(Expression, IsValid, Position, XValue)
        If IsValid = False Then Exit Function
        PowerParser2 = Tmp1 ^ Tmp2
    Else
        PowerParser2 = Tmp1
    End If
End Function

'提取数的符号
Private Function FunctionCalc(ByVal Expression As String, ByRef IsValid As Boolean, ByRef Position As Integer, ByVal XValue As Double) As Double
    Dim tmp As Double
    
    tmp = SignParser(Expression, IsValid, Position, XValue)
    If IsValid = False Then tmp = FunctionParser(Expression, IsValid, Position, XValue)
    FunctionCalc = tmp
    
End Function

'提取数（正的）
Private Function FunctionCalc2(ByVal Expression As String, ByRef IsValid As Boolean, ByRef Position As Integer, ByVal XValue As Double) As Double
    Dim tmp As Double
    
    tmp = NumberParser(Expression, IsValid, Position, XValue)
    If IsValid = False Then tmp = FunctionParser(Expression, IsValid, Position, XValue)
    FunctionCalc2 = tmp
    
End Function

' 计算函数。
' 包括以下函数：
' sin, cos, tan, cot, sec, csc, sh, ch, th, cth, sch, csch,
' arctan, arcsin, arccos, arsh, arch, arth, sqrt, log, lg, ln, exp,
' abs, sgn, int, degrees, radians,
' 还需添加：Ceil（返回大于等于其数的最小整数）、Floor（返回小于等于其数的最大整数）
'           Min、Max、Round、Fac（阶乘）、Mod（取余数）、Rand（随机数）

Private Function FunctionParser(ByVal Expression As String, ByRef IsValid As Boolean, ByRef Position As Integer, ByVal XValue As Double) As Double
    Dim Tmp1 As Double, Tmp2 As Double
    
    Dim d As Double, x As Double
    Call PassBlank(Expression, Position)
    If Match(Expression, "SIN", Position) Then
        Position = Position + 3
        Call PassBlank(Expression, Position)
        FunctionParser = Sin(BracketsParser(Expression, IsValid, Position, XValue))
        Call PassBlank(Expression, Position)
        Exit Function
    End If
    If Match(Expression, "COS", Position) Then
        Position = Position + 3
        Call PassBlank(Expression, Position)
        FunctionParser = Cos(BracketsParser(Expression, IsValid, Position, XValue))
        Call PassBlank(Expression, Position)
        Exit Function
    End If
    If Match(Expression, "TAN", Position) Then
        Position = Position + 3
        Call PassBlank(Expression, Position)
        FunctionParser = Tan(BracketsParser(Expression, IsValid, Position, XValue))
        Call PassBlank(Expression, Position)
        Exit Function
    End If
    If Match(Expression, "COT", Position) Then
        Position = Position + 3
        Call PassBlank(Expression, Position)
        d = BracketsParser(Expression, IsValid, Position, XValue)
        If Abs(d / (Atn(1) * 2)) Mod 2 = 1 Then
            FunctionParser = 0
        Else
            FunctionParser = 1 / Tan(BracketsParser(Expression, IsValid, Position, XValue))
        End If
        Call PassBlank(Expression, Position)
        Exit Function
    End If
    If Match(Expression, "SEC", Position) Then
        Position = Position + 3
        Call PassBlank(Expression, Position)
        FunctionParser = 1 / Cos(BracketsParser(Expression, IsValid, Position, XValue))
        Call PassBlank(Expression, Position)
        Exit Function
    End If
    If Match(Expression, "COSEC", Position) Or Match(Expression, "CSC", Position) Then
        Position = Position + 5
        Call PassBlank(Expression, Position)
        FunctionParser = 1 / Sin(BracketsParser(Expression, IsValid, Position, XValue))
        Call PassBlank(Expression, Position)
        Exit Function
    End If
    If Match(Expression, "SH", Position) Then
        Position = Position + 2
        Call PassBlank(Expression, Position)
        d = BracketsParser(Expression, IsValid, Position, XValue)
        FunctionParser = (Exp(d) - Exp(-d)) / 2
        Call PassBlank(Expression, Position)
        Exit Function
    End If
    If Match(Expression, "CH", Position) Then
        Position = Position + 2
        Call PassBlank(Expression, Position)
        d = BracketsParser(Expression, IsValid, Position, XValue)
        FunctionParser = (Exp(d) + Exp(-d)) / 2
        Call PassBlank(Expression, Position)
        Exit Function
    End If
    If Match(Expression, "TH", Position) Then
        Position = Position + 2
        Call PassBlank(Expression, Position)
        d = BracketsParser(Expression, IsValid, Position, XValue)
        FunctionParser = (Exp(d) - Exp(-d)) / (Exp(d) + Exp(-d))
        Call PassBlank(Expression, Position)
        Exit Function
    End If
    If Match(Expression, "CTH", Position) Then
        Position = Position + 3
        Call PassBlank(Expression, Position)
        d = BracketsParser(Expression, IsValid, Position, XValue)
        FunctionParser = (Exp(d) + Exp(-d)) / (Exp(d) - Exp(-d))
        Call PassBlank(Expression, Position)
        Exit Function
    End If
    If Match(Expression, "SCH", Position) Then
        Position = Position + 3
        Call PassBlank(Expression, Position)
        d = BracketsParser(Expression, IsValid, Position, XValue)
        FunctionParser = 2 / (Exp(d) + Exp(-d))
        Call PassBlank(Expression, Position)
        Exit Function
    End If
    If Match(Expression, "CSCH", Position) Then
        Position = Position + 4
        Call PassBlank(Expression, Position)
        d = BracketsParser(Expression, IsValid, Position, XValue)
        FunctionParser = 2 / (Exp(d) - Exp(-d))
        Call PassBlank(Expression, Position)
        Exit Function
    End If
    If Match(Expression, "ARSH", Position) Then
        Position = Position + 4
        Call PassBlank(Expression, Position)
        d = BracketsParser(Expression, IsValid, Position, XValue)
        FunctionParser = Log(d + Sqr(x ^ 2 + 1))
        Call PassBlank(Expression, Position)
        Exit Function
    End If
    If Match(Expression, "ARCH", Position) Then
        Position = Position + 4
        Call PassBlank(Expression, Position)
        d = BracketsParser(Expression, IsValid, Position, XValue)
        FunctionParser = Log(x + Sqr(x ^ 2 - 1))
        Call PassBlank(Expression, Position)
        Exit Function
    End If
    If Match(Expression, "ARTH", Position) Then
        Position = Position + 4
        Call PassBlank(Expression, Position)
        d = BracketsParser(Expression, IsValid, Position, XValue)
        FunctionParser = Log((1 + d) / (1 - d)) / 2
        Call PassBlank(Expression, Position)
        Exit Function
    End If
    If Match(Expression, "SQRT", Position) Or Match(Expression, "SQR", Position) Then
        Position = Position + 4
        Call PassBlank(Expression, Position)
        d = BracketsParser(Expression, IsValid, Position, XValue)
        FunctionParser = Sqr(d)
        Call PassBlank(Expression, Position)
        Exit Function
    End If
    If Match(Expression, "LOG", Position) Then
        Position = Position + 3
        Call PassBlank(Expression, Position)
        If Match(Expression, "(", Position) Then
            Position = Position + 1
            Tmp1 = AddMinusParser(Expression, IsValid, Position, XValue)
            If Match(Expression, ",", Position) And IsValid Then
                Position = Position + 1
                Tmp2 = AddMinusParser(Expression, IsValid, Position, XValue)
                If Match(Expression, ")", Position) And IsValid Then
                    Position = Position + 1
                    Call PassBlank(Expression, Position)
                    IsValid = True
                    FunctionParser = Log(Tmp2) / Log(Tmp1)
                    Exit Function
                End If
            End If
        End If
        IsValid = False
        Exit Function
    End If
    If Match(Expression, "LG", Position) Then
        Position = Position + 3
        Call PassBlank(Expression, Position)
        FunctionParser = Log(BracketsParser(Expression, IsValid, Position, XValue)) / Log(10#)
        Call PassBlank(Expression, Position)
        Exit Function
    End If
    If Match(Expression, "LN", Position) Then
        Position = Position + 2
        Call PassBlank(Expression, Position)
        FunctionParser = Log(BracketsParser(Expression, IsValid, Position, XValue))
        Call PassBlank(Expression, Position)
        Exit Function
    End If
    If Match(Expression, "EXP", Position) Then
        Position = Position + 3
        Call PassBlank(Expression, Position)
        FunctionParser = Exp(BracketsParser(Expression, IsValid, Position, XValue))
        Call PassBlank(Expression, Position)
        Exit Function
    End If
    If Match(Expression, "ABS", Position) Then
        Position = Position + 3
        Call PassBlank(Expression, Position)
        FunctionParser = Abs(BracketsParser(Expression, IsValid, Position, XValue))
        Call PassBlank(Expression, Position)
        Exit Function
    End If
    If Match(Expression, "SGN", Position) Then
        Position = Position + 3
        Call PassBlank(Expression, Position)
        FunctionParser = Sgn(BracketsParser(Expression, IsValid, Position, XValue))
        Call PassBlank(Expression, Position)
        Exit Function
    End If
    If Match(Expression, "INT", Position) Then
        Position = Position + 3
        Call PassBlank(Expression, Position)
        FunctionParser = Int(BracketsParser(Expression, IsValid, Position, XValue))
        Call PassBlank(Expression, Position)
        Exit Function
    End If
    If Match(Expression, "DEGREES", Position) Then
        Position = Position + 7
        Call PassBlank(Expression, Position)
        FunctionParser = BracketsParser(Expression, IsValid, Position, XValue) / PI * 180
        Call PassBlank(Expression, Position)
        Exit Function
    End If
    If Match(Expression, "RADIANS", Position) Then
        Position = Position + 7
        Call PassBlank(Expression, Position)
        FunctionParser = BracketsParser(Expression, IsValid, Position, XValue) / 180 * PI
        Call PassBlank(Expression, Position)
        Exit Function
    End If
    If Match(Expression, "ARCTAN", Position) Then
        Position = Position + 6
        Call PassBlank(Expression, Position)
        FunctionParser = Atn(BracketsParser(Expression, IsValid, Position, XValue))
        Call PassBlank(Expression, Position)
        Exit Function
    End If
    If Match(Expression, "ARCSIN", Position) Then
        Position = Position + 6
        Call PassBlank(Expression, Position)
        Tmp1 = BracketsParser(Expression, IsValid, Position, XValue)
        If Tmp1 <> 1 And Tmp1 <> -1 Then
            FunctionParser = Atn(Tmp1 / Sqr(1 - Tmp1 * Tmp1))
        Else
            If Tmp1 = 1 Then FunctionParser = PI / 2 Else FunctionParser = -PI / 2
        End If
        Call PassBlank(Expression, Position)
        Exit Function
    End If
    If Match(Expression, "ARCCOS", Position) Then
        Position = Position + 6
        Call PassBlank(Expression, Position)
        Tmp1 = BracketsParser(Expression, IsValid, Position, XValue)
        If Tmp1 <> 1 And Tmp1 <> -1 Then
            FunctionParser = PI / 2 - Atn(Tmp1 / Sqr(1 - Tmp1 * Tmp1))
        Else
            If Tmp1 = 1 Then FunctionParser = 0# Else FunctionParser = PI
        End If
        Call PassBlank(Expression, Position)
        Exit Function
    End If
'    If Match(Expression, "POW", Position) Then
'        Position = Position + 3
'        Call PassBlank(Expression, Position)
'        If Match(Expression, "(", Position) Then
'            Position = Position + 1
'            Tmp1 = AddMinusParser(Expression, IsValid, Position, XValue)
'            If Match(Expression, ",", Position) And IsValid Then
'                Position = Position + 1
'                Tmp2 = AddMinusParser(Expression, IsValid, Position, XValue)
'                If Match(Expression, ")", Position) And IsValid Then
'                    Position = Position + 1
'                    Call PassBlank(Expression, Position)
'                    IsValid = True
'                    FunctionParser = Tmp1 ^ Tmp2
'                    Exit Function
'                End If
'            End If
'        End If
'        IsValid = False
'        Exit Function
'    End If
    FunctionParser = BracketsParser(Expression, IsValid, Position, XValue)
    Call PassBlank(Expression, Position)
End Function

' 计算括号内的数和表达式。
Private Function BracketsParser(ByVal Expression As String, ByRef IsValid As Boolean, ByRef Position As Integer, ByVal XValue As Double) As Double
    Dim tmp As Double
    
    If Match(Expression, "(", Position) Then
        Position = Position + 1
        Call PassBlank(Expression, Position)
        ' 在括号内又是一个表达式。
        tmp = AddMinusParser(Expression, IsValid, Position, XValue)
        If IsValid And Match(Expression, ")", Position) Then
            BracketsParser = tmp
            Position = Position + 1
            IsValid = True
        Else
            IsValid = False
        End If
    End If
End Function

' 提取表达式中的数。
Private Function NumberParser(ByVal Expression As String, ByRef IsValid As Boolean, ByRef Position As Integer, ByVal XValue As Double) As Double
    ' Tmp1 是提取出来的数，Tmp2 表示在表达式中提取的位置（相对于数而言）。
    Dim Tmp1 As Double, Tmp2 As Double
    Dim TmpStr As String
    Call PassBlank(Expression, Position)
    
    ' 如果表达式中含有未知数 x，那么就把 XValue 代入到 x 中。
    If Match(Expression, tmpChar, Position) Then
        NumberParser = XValue
        IsValid = True
        Position = Position + 1
        Call PassBlank(Expression, Position)
        Exit Function
    End If
    
    Tmp1 = 0
    Tmp2 = 0
    TmpStr = Mid(Expression, Position, 1)
    
    ' 如果没有提取到有效字符（由于是数字的第一位，所以不会有“e”），就指出运算错误。
    If (TmpStr >= "0" And TmpStr <= "9") Or TmpStr = "." Then
        ' 数分三部分：数 + 小数点 + 小数 + e + 指数
        While TmpStr >= "0" And TmpStr <= "9" Or TmpStr = " "
            Tmp2 = Tmp2 + 1
            TmpStr = Mid(Expression, Position + Tmp2, 1)
        Wend
        
        If TmpStr = "." Then Tmp2 = Tmp2 + 1
        TmpStr = Mid(Expression, Position + Tmp2, 1)
        
        While TmpStr >= "0" And TmpStr <= "9" Or TmpStr = " "
            Tmp2 = Tmp2 + 1
            TmpStr = Mid(Expression, Position + Tmp2, 1)
        Wend
        
        If TmpStr = "E" Then Tmp2 = Tmp2 + 1
        TmpStr = Mid(Expression, Position + Tmp2, 1)
        
        While TmpStr >= "0" And TmpStr <= "9" Or TmpStr = " "
            Tmp2 = Tmp2 + 1
            TmpStr = Mid(Expression, Position + Tmp2, 1)
        Wend
        
        ' 上面的一切操作都是为了寻找数值表达式的长度，而不是提取数字。
        Tmp1 = Val(Mid(Expression, Position))
        Position = Position + Tmp2
        IsValid = True
    Else
        IsValid = False
    End If
    NumberParser = Tmp1
End Function

' 处理带符号的数。
Private Function SignParser(ByVal Expression As String, ByRef IsValid As Boolean, ByRef Position As Integer, ByVal XValue As Double) As Double
    Dim tmp As Double
    Dim Sign As Integer
    Sign = 1
    Call PassBlank(Expression, Position)
    If Match(Expression, tmpChar, Position) Then
        SignParser = XValue
        IsValid = True
        Position = Position + 1
        Call PassBlank(Expression, Position)
        Exit Function
    End If
    If Match(Expression, "-", Position) Then
        Position = Position + 1
        Sign = -1
    ElseIf Match(Expression, "+", Position) Or (Mid(Expression, Position, 1) >= "0" And Mid(Expression, Position, 1) <= "9") Or Mid(Expression, Position, 1) = "." Then
        Sign = 1
    Else
        IsValid = False
        Exit Function
    End If
    tmp = FunctionCalc2(Expression, IsValid, Position, XValue)
    SignParser = tmp * Sign
End Function

' 检查 Expression 的 position 位置是否为 Expression2
Private Function Match(ByVal Expression As String, ByVal Expression2 As String, ByRef Position As Integer) As Boolean
    If Mid(Expression, Position, Len(Expression2)) = Expression2 Then Match = True Else Match = False
End Function

' 在分析表达式的时候要跳过空格
Private Sub PassBlank(ByVal Expression As String, ByRef Position As Integer)
    While Mid(Expression, Position, 1) = " "
        Position = Position + 1
    Wend
End Sub

