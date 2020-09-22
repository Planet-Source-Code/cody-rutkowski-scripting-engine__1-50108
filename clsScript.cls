VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' /* _
    Script Engine was programmed by xIso5 _
    Please do not redistribute this code _
    you can use it but dont give it out to every one!!! _
 _
    Big Thanks To Chazz For Asking me to do this, if he never _
    asked i would never of made one this good! Thank You So MUCH!


Private mCode As String
Private mName As String

Private pubVar() As typVariable

Private Func() As typFunction

Public Property Get Code() As String
    Code = mCode
End Property

Public Property Let Code(ByVal vNewValue As String)
    mCode = vNewValue
End Property

Public Property Get name() As String
    name = mName
End Property

Public Property Let name(ByVal vNewValue As String)
    mName = vNewValue
End Property




'###################################'
'# Get The Script Ready To Execute #'
'###################################'
Function PrepareScript()
Dim cFunc As Integer 'FunctionIndex
Dim ArgStr 'Stores Arguments For a Function
'Reset Function Count
ReDim Func(0)
'Split The Code into lines
Dim lines: lines = Split(mCode, vbCrLf)
'Interpret Each Line
For i = 0 To UBound(lines)
    'Prepare Line To Interpret quickly
    tmp = lines(i)
    tmp = Trim(tmp)
    If tmp = "" Then tmp = " "
    tmp = Split(tmp, " ")
    'Interpret First word of the line
    Select Case UCase(tmp(0))
        Case "FUNCTION"
            'add a function
            cFunc = UBound(Func) + 1
            ReDim Preserve Func(cFunc)
            'reset the variables
            ReDim Func(cFunc).Var(0)
            'reset lines of code
            ReDim Func(cFunc).Code(0)
            'Set the name of the function
            Func(cFunc).name = UCase(tmp(1))
            Func(cFunc).name = Left(Func(cFunc).name, InStr(Func(cFunc).name, "(") - 1)
            'reset arguments
            ReDim Func(cFunc).Args(0)
            'set temporary argument string
            ArgStr = UCase(tmp(1))
            ArgStr = Mid(ArgStr, InStr(ArgStr, "(") + 1)
            ArgStr = StrReverse(ArgStr)
            ArgStr = Mid(ArgStr, InStr(ArgStr, ")") + 1)
            ArgStr = StrReverse(ArgStr)
            ArgStr = Split(ArgStr, ",")
            For ii = 0 To UBound(ArgStr)
                ReDim Preserve Func(cFunc).Args(UBound(Func(cFunc).Args) + 1)
                Func(cFunc).Args(ii).name = ArgStr(ii)
                Func(cFunc).Args(ii).Value = ""
            Next ii
        Case "DIM"
            'Add A Variable
            Dim AllDims: AllDims = Trim(lines(i))
            AllDims = Split(AllDims, ",")
            AllDims(0) = Split(AllDims(0), " ")(1)
            For ii = 0 To UBound(AllDims)
                AllDims(ii) = Trim(AllDims(ii))
                If AllDims(ii) <> "" Then
                    ReDim Preserve Func(cFunc).Var(UBound(Func(cFunc).Var) + 1)
                    'Set The Name
                    Func(cFunc).Var(UBound(Func(cFunc).Var)).name = AllDims(ii)
                    'Set The Value
                    Func(cFunc).Var(UBound(Func(cFunc).Var)).Value = ""
                End If
            Next ii
        Case "END"
            If UCase(tmp(1)) = "FUNCTION" Then
                'The Function Is Done
                cFunc = -1
            Else
                GoTo AddCode
            End If
        Case Else
AddCode:
            If cFunc <> -1 Then
                'Add A New Line Of Code
                ReDim Preserve Func(cFunc).Code(UBound(Func(cFunc).Code) + 1)
                'Set The Line Of Code
                Func(cFunc).Code(UBound(Func(cFunc).Code)) = Trim(lines(i))
            End If
    End Select
Next i
End Function

'##################################'
'# Execute PreProgrammed Function #'
'##################################'
Function DoFunction(fName As String, FuncIndex As Integer) As String
Dim ArgStr
Dim orig_fName As String
If InStr(fName, "(") = 0 Or InStr(fName, ")") = 0 Then
    fName = fName & "()"
End If
ArgStr = Mid(fName, InStr(fName, "(") + 1)
ArgStr = StrReverse(ArgStr)
ArgStr = Mid(ArgStr, InStr(ArgStr, ")") + 1)
ArgStr = StrReverse(ArgStr)

'ArgStr = Split(ArgStr, "," )

Dim Char As String
Dim InQuotes As Boolean, InParenthisis As Integer
Dim IsQuotes As Boolean, IsParenthisis As Boolean
Dim IsEqual As Boolean
Dim ArgVar() As String
Dim LastPos As Integer
'Reset All Needed Variables Before it is used
ReDim ArgVar(0)
IsEqual = False
InQuotes = False
IsQuotes = False
IsParenthisis = False
InParenthisis = 0
LastPos = 1
'Begin Interpretting
For i = 1 To Len(ArgStr)
    'Get Letter at point ii
    Char = Mid(ArgStr, i, 1)
    Select Case Char
        Case "("
            If InQuotes = False Then
                InParenthisis = InParenthisis + 1
                IsParenthisis = True
            End If
        Case ")"
            If InQuotes = False Then _
                InParenthisis = InParenthisis - 1
        Case """": InQuotes = Not InQuotes: IsQuotes = True
        Case ","
            If InQuotes = False And InParenthisis = 0 Then
                'Add A New Argument
                ReDim Preserve ArgVar(UBound(ArgVar) + 1)
                'Set The Argument From The Portion Of String
                ArgVar(UBound(ArgVar)) = Mid(ArgStr, LastPos, i - LastPos)
                LastPos = i + 1
            End If
    End Select
Next i
'Get Last Argument
ReDim Preserve ArgVar(UBound(ArgVar) + 1)
'Set The Argument From The Portion Of String
ArgVar(UBound(ArgVar)) = Mid(ArgStr, LastPos, i - LastPos)
    
For i = 1 To UBound(ArgVar)
    'Clean Up All Of The Arguments Now
    ArgVar(i) = Interpret(CStr(ArgVar(i)), FuncIndex)
Next i
fName = Left(fName, InStr(fName, "(") - 1)
Select Case UCase(fName)
    Case "MSGBOX"
        DoFunction = MsgBox(ArgVar(1), ArgVar(2), ArgVar(3))
    Case "CHR"
        DoFunction = Chr(ArgVar(1))
    Case "ASC"
        DoFunction = Asc(ArgVar(1))
    Case "LEFT"
        DoFunction = Left(ArgVar(1), CLng(ArgVar(2)))
    Case "MID"
        DoFunction = Mid(ArgVar(1), CLng(ArgVar(2)), CLng(ArgVar(3)))
    Case "RIGHT"
        DoFunction = Right(ArgVar(1), CLng(ArgVar(2)))
    Case "LCASE"
        DoFunction = LCase(ArgVar(1))
    Case "UCASE"
        DoFunction = UCase(ArgVar(1))
    Case "LTRIM"
        DoFunction = LTrim(ArgVar(1))
    Case "TRIM"
        DoFunction = Trim(ArgVar(1))
    Case "RTRIM"
        DoFunction = RTrim(ArgVar(1))
    Case "LEN"
        DoFunction = Len(ArgVar(1))
    Case "INSTR"
        DoFunction = InStr(ArgVar(1), ArgVar(2))
    Case "STRREVERSE"
        DoFunction = StrReverse(ArgVar(1))
    Case "SPACE"
        DoFunction = Space(ArgVar(1))
    Case Else
        DoFunction = ""
End Select
End Function

'###############################'
'# Runs the Function From Here #'
'###############################'
Function Run(Optional fName As String, Optional prevFuncIndex As Integer)
Dim ArgStr
Dim orig_fName As String
'Store Original Func Name
orig_fName = fName
'Correct 'Optional' Error
If fName = vbNullString Then
    'if there is no function name _
     Set Function Name to MAIN
    fName = "MAIN"
End If
If InStr(fName, "(") = 0 Or InStr(fName, ")") = 0 Then
    fName = fName & "()"
End If
ArgStr = Mid(fName, InStr(fName, "(") + 1)
ArgStr = StrReverse(ArgStr)
ArgStr = Mid(ArgStr, InStr(ArgStr, ")") + 1)
ArgStr = StrReverse(ArgStr)

Dim Char As String
Dim InQuotes As Boolean, InParenthisis As Integer
Dim IsQuotes As Boolean, IsParenthisis As Boolean
Dim IsEqual As Boolean
Dim ArgVar() As String
Dim LastPos As Integer
'Reset All Needed Variables Before it is used
ReDim ArgVar(0)
IsEqual = False
InQuotes = False
IsQuotes = False
IsParenthisis = False
InParenthisis = 0
LastPos = 1
'Begin Interpretting
For i = 1 To Len(ArgStr)
    'Get Letter at point ii
    Char = Mid(ArgStr, i, 1)
    Select Case Char
        Case "("
            If InQuotes = False Then
                InParenthisis = InParenthisis + 1
                IsParenthisis = True
            End If
        Case ")"
            If InQuotes = False Then _
                InParenthisis = InParenthisis - 1
        Case """": InQuotes = Not InQuotes: IsQuotes = True
        Case ","
            If InQuotes = False And InParenthisis = 0 Then
                'Add A New Argument
                ReDim Preserve ArgVar(UBound(ArgVar) + 1)
                'Set The Argument From The Portion Of String
                ArgVar(UBound(ArgVar)) = Mid(ArgStr, LastPos, i - LastPos)
                LastPos = i + 1
            End If
    End Select
Next i
'Get Last Argument
ReDim Preserve ArgVar(UBound(ArgVar) + 1)
'Set The Argument From The Portion Of String
ArgVar(UBound(ArgVar)) = Mid(ArgStr, LastPos, i - LastPos)

fName = Left(fName, InStr(fName, "(") - 1)
Dim index As Integer
index = -1 'Set Function Index Default to -1(not found)
For i = 0 To UBound(Func)
    If UCase(fName) = UCase(Func(i).name) Then
        'Function Found, set the index to the _
         functions index
        index = i
        Exit For
    End If
Next i
'The Function Was not found, so exit
If index = -1 Then Run = DoFunction(orig_fName, prevFuncIndex): Exit Function

For i = 0 To UBound(ArgVar)
    If prevFuncIndex = 0 Then prevFuncIndex = -1
    ArgVar(i) = Interpret(CStr(ArgVar(i)), prevFuncIndex)
Next i

For i = 0 To UBound(ArgVar) - 1
    'Set The Value Of The Arguments Of This Function
    Func(index).Args(i).Value = ArgVar(i + 1)
Next i
Dim Part1 As String
Dim Part2 As String
Dim CodeStr As String ', Char As String
'Dim InQuotes As Boolean ', InParenthisis As Integer
'Dim IsQuotes As Boolean, IsParenthisis As Boolean
'Dim IsEqual As Boolean
Dim SkipElse As Integer
Dim IfCount As Integer
Dim curIF As Integer
IfCount = 0 'Not In An If Statement
SkipElse = -1 'Not In An If Statement
'Go Through Each Line Of The Function and use the code _
 that has been saved for execution
For i = 1 To UBound(Func(index).Code)
    'Reset Interpreting Variables
    IsEqual = False
    InQuotes = False
    IsQuotes = False
    IsParenthisis = False
    InParenthisis = 0
    'Set CodeString to the line of code
    CodeStr = Trim(Func(index).Code(i))
    'Find Out how to execute this line
    For ii = 1 To Len(CodeStr)
        'Get Letter at point ii
        Char = Mid(CodeStr, ii, 1)
        Select Case Char
            Case "("
                If InQuotes = False Then
                    InParenthisis = InParenthisis + 1
                    IsParenthisis = True
                End If
            Case ")"
                If InQuotes = False Then _
                    InParenthisis = InParenthisis - 1
            Case """": InQuotes = Not InQuotes: IsQuotes = True
            Case "="
                If InQuotes = False Then _
                    IsEqual = True
        End Select
    Next ii
    Dim toRunFunc As String
    Dim PartSep As String
    If UCase(CodeStr) = "END IF" Then
        If IfCount = 1 Then
            SkipElse = -1
        End If
        SkipElse = -1
        IfCount = IfCount - 1
    ElseIf UCase(CodeStr) = "ELSE" Then
        If IfCount = 1 Then
            If SkipElse = 0 Then
                SkipElse = -1
            Else
                SkipElse = 0
            End If
        End If
    ElseIf UCase(Left(CodeStr, 3)) = "IF " And UCase(Right(CodeStr, 5)) = " THEN" Then
        CodeStr = Mid(CodeStr, 4)
        CodeStr = StrReverse(CodeStr)
        CodeStr = Mid(CodeStr, 6)
        CodeStr = StrReverse(CodeStr)
        CodeStr = Trim(CodeStr)
        If InStr(CodeStr, " = ") Then
            PartSep = " = "
        End If
        Part1 = Split(CodeStr, PartSep)(0)
        Part2 = Split(CodeStr, PartSep)(1)
        Part1 = Interpret(Part1, index)
        Part2 = Interpret(Part2, index)
        IfCount = IfCount + 1

        Select Case Trim(PartSep)
            Case "="
                If Part1 = Part2 Then
                    SkipElse = 1
                Else
                    SkipElse = 0
                End If
            Case Else
        End Select
    ElseIf IsEqual = True Then
        If SkipElse = 0 Then GoTo NextI
        Part1 = Split(CodeStr, " = ")(0)
        Part2 = Split(CodeStr, " = ")(1)
        'We Are Going To Be Setting Something
        ''If IsParenthisis = True Then
        ''Else
            'This is not a function, it is a variable
            Part2 = Interpret(Part2, index)
            If UCase(Part1) = UCase(Func(index).name) Then
                Func(index).RetVal = Part2
            Else
                For ii = 0 To UBound(Func(index).Var)
                    'is this variable the same as part1
                    If UCase(Func(index).Var(ii).name) = UCase(Part1) Then
                        'it is now set the value
                        Func(index).Var(ii).Value = Part2
                        Exit For
                    End If
                Next ii
            End If
        ''End If
        'MsgBox CodeStr
    Else
        If SkipElse = 0 Then GoTo NextI
        'This must be a function
        If IsParenthisis = True Then
            toRunFunc = CodeStr ' Left(CodeStr, InStr(CodeStr, "(") - 1)
            Run toRunFunc, index
        End If
    End If
NextI:
Next i
If Func(index).RetVal <> "" Then
    Run = Func(index).RetVal
End If
End Function

'##################################'
'# This Function Reads In A Line  #'
'# And Replaces Variables w/ Vals #'
'# Fixes up strings, etc.         #'
'##################################'
Function Interpret(CodeStr As String, FuncIndex As Integer) As String
Dim i As Integer
Dim Char As String
Dim InQuotes As Boolean, InParenthisis As Integer
Dim IsQuotes As Boolean, IsParenthisis As Boolean
Dim IsEqual As Boolean
Dim portion() As String
Dim LastPos As Integer
'Reset All Needed Variables Before it is used
ReDim portion(0)
IsEqual = False
InQuotes = False
IsQuotes = False
IsParenthisis = False
InParenthisis = 0
LastPos = 1
'Begin Interpretting
For i = 1 To Len(CodeStr)
    'Get Letter at point ii
    Char = Mid(CodeStr, i, 1)
    Select Case Char
        Case "("
            If InQuotes = False Then
                InParenthisis = InParenthisis + 1
                IsParenthisis = True
            End If
        Case ")"
            If InQuotes = False Then _
                InParenthisis = InParenthisis - 1
        Case """": InQuotes = Not InQuotes: IsQuotes = True
        Case "&"
            If InQuotes = False And InParenthisis = 0 Then
                'Add A Seperated Portion
                ReDim Preserve portion(UBound(portion) + 1)
                portion(UBound(portion)) = Mid(CodeStr, LastPos, i - LastPos)
                LastPos = i + 1
            End If
    End Select
Next i
'Add The Last Seperated Portion
ReDim Preserve portion(UBound(portion) + 1)
portion(UBound(portion)) = Mid(CodeStr, LastPos, i - LastPos)

Dim PrepStr As String
If UBound(portion) > 1 Then
    'Brake Down the String until we can't anymore
    For i = 1 To UBound(portion)
        'Interpret Each Portion
        portion(i) = Interpret(portion(i), FuncIndex)
    Next i
    PrepStr = ""
    For i = 1 To UBound(portion)
        'Add On To Prepared String
        PrepStr = PrepStr & portion(i)
    Next i
Else
    'Now that we have 1 String Parse it
    PrepStr = portion(1)
    'Clean It Up A Bit
    PrepStr = Trim(PrepStr)
    If (InStr(PrepStr, "+") Or InStr(PrepStr, "-") Or InStr(PrepStr, "/") Or InStr(PrepStr, "*")) And InStr(PrepStr, """") = 0 Then
        'Use Math Engine
        Dim m As New clsMath
        'Add All Need Constants
        For i = 0 To UBound(Func(FuncIndex).Args)
            m.AddConstant Func(FuncIndex).Args(i).name, Func(FuncIndex).Args(i).Value
        Next i
        For i = 0 To UBound(Func(FuncIndex).Var)
            m.AddConstant Func(FuncIndex).Var(i).name, Func(FuncIndex).Var(i).Value
        Next i
        For i = 0 To UBound(pubVar)
            m.AddConstant pubVar(i).name, pubVar(i).Value
        Next i
        
        PrepStr = CStr(m.ParseExpression(PrepStr))
        
    End If
    'Make sure we are using a valid function for these variables
    If FuncIndex > -1 Then
        'Check to see if this string is a retval
        If UCase(PrepStr) = UCase(Func(FuncIndex).name) Then
            PrepStr = Func(FuncIndex).RetVal
        End If
        'Check to see if this string is an argument?
        For i = 0 To UBound(Func(FuncIndex).Args)
            If UCase(PrepStr) = UCase(Func(FuncIndex).Args(i).name) Then
                'We Found A Match, Replace it with the value
                PrepStr = Func(FuncIndex).Args(i).Value
            End If
        Next i
        'Check to see if this string is a variable?
        For i = 1 To UBound(Func(FuncIndex).Var)
            If UCase(PrepStr) = UCase(Func(FuncIndex).Var(i).name) Then
                'We Found A Match, Replace it with the value
                PrepStr = Func(FuncIndex).Var(i).Value
            End If
        Next i
        'Check to see if this string is an pubVariable?
        For i = 0 To UBound(pubVar)
            If UCase(PrepStr) = UCase(pubVar(i).name) Then
                'We Found A Match, Replace it with the value
                PrepStr = pubVar(i).Value
            End If
        Next i
    End If
    'Is this a function
    If IsParenthisis = True Then
        PrepStr = Run(PrepStr, FuncIndex)
    End If
    'Is This A String With Quotes
    If Left(PrepStr, 1) = """" And Right(PrepStr, 1) = """" Then
        'Remove the Quotes
        PrepStr = Mid(PrepStr, 2)
        PrepStr = StrReverse(PrepStr)
        PrepStr = Mid(PrepStr, 2)
        PrepStr = StrReverse(PrepStr)
    End If
End If
If PrepStr <> "" Then
    'Set Function Value To The Interpretted String
    Interpret = PrepStr
End If
End Function

Private Sub Class_Initialize()
'reset vars
ReDim pubVar(0)
'Prepare All Of The Variables at the beginning like!?
'vbCrLf -newline
ReDim Preserve pubVar(UBound(pubVar) + 1)
pubVar(UBound(pubVar)).name = "vbCrLf"
pubVar(UBound(pubVar)).Value = vbCrLf

ReDim Preserve pubVar(UBound(pubVar) + 1)
pubVar(UBound(pubVar)).name = "vbBack"
pubVar(UBound(pubVar)).Value = vbBack

ReDim Preserve pubVar(UBound(pubVar) + 1)
pubVar(UBound(pubVar)).name = "vbCr"
pubVar(UBound(pubVar)).Value = vbCr

ReDim Preserve pubVar(UBound(pubVar) + 1)
pubVar(UBound(pubVar)).name = "vbFormFeed"
pubVar(UBound(pubVar)).Value = vbFormFeed

ReDim Preserve pubVar(UBound(pubVar) + 1)
pubVar(UBound(pubVar)).name = "vbLf"
pubVar(UBound(pubVar)).Value = vbLf

ReDim Preserve pubVar(UBound(pubVar) + 1)
pubVar(UBound(pubVar)).name = "vbNewLine"
pubVar(UBound(pubVar)).Value = vbNewLine

ReDim Preserve pubVar(UBound(pubVar) + 1)
pubVar(UBound(pubVar)).name = "vbNullChar"
pubVar(UBound(pubVar)).Value = vbNullChar

ReDim Preserve pubVar(UBound(pubVar) + 1)
pubVar(UBound(pubVar)).name = "vbNullString"
pubVar(UBound(pubVar)).Value = vbNullString

ReDim Preserve pubVar(UBound(pubVar) + 1)
pubVar(UBound(pubVar)).name = "vbObjectError"
pubVar(UBound(pubVar)).Value = vbObjectError

ReDim Preserve pubVar(UBound(pubVar) + 1)
pubVar(UBound(pubVar)).name = "vbTab"
pubVar(UBound(pubVar)).Value = vbTab

ReDim Preserve pubVar(UBound(pubVar) + 1)
pubVar(UBound(pubVar)).name = "vbVerticalTab"
pubVar(UBound(pubVar)).Value = vbVerticalTab


End Sub
