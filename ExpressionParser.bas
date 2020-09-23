Attribute VB_Name = "ExpressionParser"
Option Explicit

Private Function ReadVariable(VariableName As String) As Double
  ' Variable control isn't implemented, because the way to do this best
  ' depends on what you're going to use this code for.
  
  ' If you're going to use variable support, you should place the
  ' code to read a variable here. The VariableName parameter contains
  ' the name of the variable being read. Place the value for the
  ' variable as the result, like: ReadVariable = strVariables(intVar)
  
  MsgBox "Variable control not implemented, assuming 0 (zero)"
  ReadVariable = 0
End Function

Private Function BasicFunctions(Expression As String) As Boolean
  Dim intX As Integer, strFunction As String, strParam As String, dblParam As Double
  
  'Split function from parameter-list
  intX = InStr(1, Expression, "(")
  strFunction = Left(Expression, intX - 1)
  strParam = Mid(Expression, intX + 1, Len(Expression) - intX - 1)
  dblParam = Val(strParam)
  
  'Select function
  Select Case strFunction
  Case "sin"
    Expression = Str(Sin(dblParam))
  Case "asin"
    Expression = Str(Atn(dblParam / Sqr(-dblParam * dblParam + 1)))
  Case "cos"
    Expression = Str(Cos(dblParam))
  Case "acos"
    Expression = Str(Atn(-dblParam / Sqr(-dblParam * dblParam + 1)) + 2 * Atn(1))
  Case "tan"
    Expression = Str(Tan(dblParam))
  Case "atan"
    Expression = Str(Atn(dblParam))
  Case "int"
    Expression = Str(Int(dblParam))
  Case "frac"
    Expression = Str(Abs(dblParam - Int(dblParam)))
  Case "log"
    Expression = Str(Log(dblParam) / Log(10))
  Case "ln"
    Expression = Str(Log(dblParam))
  Case "abs"
    Expression = Str(Abs(dblParam))
  Case "sign"
    Expression = Str(Sgn(dblParam))
  Case "rnd"
    Randomize
    Expression = Str(Int(Rnd * dblParam))
  Case "sqrt"
    Expression = Str(Sqr(dblParam))
  Case Else
    Exit Function
  End Select
  BasicFunctions = True
End Function

Private Function ExecuteFunction(Expression As String) As Double
  Dim intX As Integer, strFunction As String, strParameters As String, strParams() As String
  
  'Split function from parameter-list
  intX = InStr(1, Expression, "(")
  strFunction = Left(Expression, intX)
  strParameters = Mid(Expression, intX + 1)
  strParameters = Left(strParameters, InStrRev(strParameters, ")") - 1)
  
  'Split parameters into array
  strParams = Split(strParameters, ",")
  
  'Parse each parameter
  For intX = 0 To UBound(strParams)
    strFunction = strFunction & Trim(Str(ParseExpression(strParams(intX)))) & ","
  Next intX
  'Call function
  strFunction = Left(strFunction, Len(strFunction) - 1) & ")"
  If BasicFunctions(strFunction) Then
    ExecuteFunction = ParseExpression(strFunction)
  Else
    MsgBox "Error handling code for ExecuteFunction not implemented"
  End If
  
  'Return result
  ExecuteFunction = ParseExpression(strFunction)
End Function

Private Function IsFunction(Expression As String) As Boolean
  Dim bytStart As Byte
  bytStart = Asc(Left(Expression, 1))
  If bytStart = 95 Or (bytStart > 64 And bytStart < 91) Or (bytStart > 96 And bytStart < 123) Then IsFunction = True
End Function

Private Function NextGroup(Expression As String) As String
  Dim intPosition As Integer, intDepth As Integer
  intPosition = 1
  
  'Check whether first character is paranthesis
  If Left(Expression, 1) = "(" Then
    'Find end of sub-expression
    intDepth = 1
    Do Until intDepth = 0
      intPosition = intPosition + 1
      If Mid(Expression, intPosition, 1) = "(" Then intDepth = intDepth + 1
      If Mid(Expression, intPosition, 1) = ")" Then intDepth = intDepth - 1
    Loop
    NextGroup = Left(Expression, intPosition)
    Expression = Mid(Expression, intPosition + 1)
    Exit Function
  End If
  
  'Check whether first character is operator
  If InStr(1, "*/+-%^", Left(Expression, 1)) Then
    NextGroup = Left(Expression, 1)
    Expression = Mid(Expression, 2)
    Exit Function
  End If
  
  'Find next operator
  Do Until InStr(1, "*/+-%^", Mid(Expression, intPosition, 1)) And intDepth = 0
    If Mid(Expression, intPosition, 1) = "(" Then intDepth = intDepth + 1
    If Mid(Expression, intPosition, 1) = ")" Then intDepth = intDepth - 1
    intPosition = intPosition + 1
  Loop
  NextGroup = Trim(Left(Expression, intPosition - 1))
  Expression = Mid(Expression, intPosition)
End Function

Public Function ParseExpression(Expression As String) As Double
  Dim strTemp As String, intGroups As Integer, intCurrent As Integer, intMoving As Integer, blnNegation As Boolean
  ReDim strGroups(0 To 0) As String
  strTemp = Expression
  
  ' 1. Split expression into groups
  Do Until strTemp = ""
    ReDim Preserve strGroups(0 To intGroups)
    strGroups(intGroups) = NextGroup(strTemp)
    intGroups = intGroups + 1
  Loop
  
  ' 2. Find functions
  For intCurrent = 0 To intGroups - 1
    If IsFunction(Left(Trim(strGroups(intCurrent)), 1)) Then
      'Check for parameter-list
      If InStr(1, strGroups(intCurrent), "(") = 0 Then
        'No parameters; this means we're dealing with a variable
        strGroups(intCurrent) = Str(ReadVariable(strGroups(intCurrent)))
      Else
        'Execute function and replace group with result
        strGroups(intCurrent) = Str(ExecuteFunction(strGroups(intCurrent)))
      End If
    End If
  Next intCurrent
  
  ' 3. Find sub-expressions
  For intCurrent = 0 To intGroups - 1
    If Left(strGroups(intCurrent), 1) = "(" Then
      'Parse sub-expression and replace with result
      strGroups(intCurrent) = ParseExpression(Mid(strGroups(intCurrent), 2, Len(strGroups(intCurrent)) - 2))
    End If
  Next intCurrent
  
  ' 4. Find exponentiations (^)
  intCurrent = 0
  Do Until intCurrent >= intGroups
    If strGroups(intCurrent) = "^" Then
      'Resolve exponentiations
      strGroups(intCurrent - 1) = Str(Val(strGroups(intCurrent - 1)) ^ Val(strGroups(intCurrent + 1)))
      'Remove two groups from array
      For intMoving = intCurrent To intGroups - 3
        strGroups(intMoving) = strGroups(intMoving + 2)
      Next intMoving
      intGroups = intGroups - 2
      ReDim Preserve strGroups(0 To intGroups - 1)
    Else
      intCurrent = intCurrent + 1
    End If
  Loop
  
  ' 5. Find multiplications (*) and divisions (/)
  intCurrent = 0
  Do Until intCurrent >= intGroups
    If strGroups(intCurrent) = "*" Or strGroups(intCurrent) = "/" Then
      If strGroups(intCurrent) = "*" Then
        'Resolve multiplication
        strGroups(intCurrent - 1) = Str(Val(strGroups(intCurrent - 1)) * Val(strGroups(intCurrent + 1)))
      Else
        'Resolve division
        strGroups(intCurrent - 1) = Str(Val(strGroups(intCurrent - 1)) / Val(strGroups(intCurrent + 1)))
      End If
      'Remove two groups from array
      For intMoving = intCurrent To intGroups - 3
        strGroups(intMoving) = strGroups(intMoving + 2)
      Next intMoving
      intGroups = intGroups - 2
      ReDim Preserve strGroups(0 To intGroups - 1)
    Else
      intCurrent = intCurrent + 1
    End If
  Loop
  
  ' 6. Find modulus (%)
  intCurrent = 0
  Do Until intCurrent >= intGroups
    If strGroups(intCurrent) = "%" Then
      'Resolve modulus
      strGroups(intCurrent - 1) = Str(Val(strGroups(intCurrent - 1)) Mod Val(strGroups(intCurrent + 1)))
      'Remove two groups from array
      For intMoving = intCurrent To intGroups - 3
        strGroups(intMoving) = strGroups(intMoving + 2)
      Next intMoving
      intGroups = intGroups - 2
      ReDim Preserve strGroups(0 To intGroups - 1)
    Else
      intCurrent = intCurrent + 1
    End If
  Loop
  
  ' 7. Find additions (+) and substractions (-)
  intCurrent = 0
  Do Until intCurrent >= intGroups
    If strGroups(intCurrent) = "+" Or strGroups(intCurrent) = "-" Then
      If strGroups(intCurrent) = "+" Then
        'Resolve addition
        strGroups(intCurrent - 1) = Str(Val(strGroups(intCurrent - 1)) + Val(strGroups(intCurrent + 1)))
      Else
        'Resolve substraction
        ' FIRST WE MUST FIND OUT WHETHER WE WANT SUSTRACTION OR NEGATION !!
        If intCurrent = 0 Then
          blnNegation = True
        ElseIf InStr("1234567890", Left(strGroups(intCurrent - 1), 1)) = 0 Then
          blnNegation = True
        Else
          blnNegation = False
        End If
        If blnNegation Then
          'Insert group with "0", because substraction from zero is the same as negation
          intGroups = intGroups + 1
          ReDim Preserve strGroups(0 To intGroups - 1)
          For intMoving = intGroups - 2 To intCurrent Step -1
            strGroups(intMoving + 1) = strGroups(intMoving)
          Next intMoving
          strGroups(intCurrent) = "0"
          intCurrent = intCurrent + 1
        End If
        strGroups(intCurrent - 1) = Str(Val(strGroups(intCurrent - 1)) - Val(strGroups(intCurrent + 1)))
      End If
      'Remove two groups from array
      For intMoving = intCurrent To intGroups - 3
        strGroups(intMoving) = strGroups(intMoving + 2)
      Next intMoving
      intGroups = intGroups - 2
      ReDim Preserve strGroups(0 To intGroups - 1)
    Else
      intCurrent = intCurrent + 1
    End If
  Loop
  
  ' 8. Return result
  ParseExpression = Val(strGroups(0))
End Function


