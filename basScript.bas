Attribute VB_Name = "basScript"
'These Variables Are Manditory!!!!

Public Type typVariable
    name As String
    Value As String
End Type

Public Type typFunction
    name As String
    Var() As typVariable
    Args() As typVariable
    Code() As String
    RetVal As String
End Type

'this is used for a script array[Many scripts]   :-)
Private ScriptArray() As clsScript

Function LoadScript(file, Optional CallFuncNow = "OnLoad()")
Dim m As New clsMath
Dim index: index = -1
index = UBound(ScriptArray) + 1
ReDim Preserve ScriptArray(index)
Set ScriptArray(index) = New clsScript
Dim ff: ff = FreeFile
Open App.Path & "\" & file For Binary As #ff
    ScriptArray(index).Code = Input(LOF(ff), ff)
Close #ff

ScriptArray(index).name = file
ScriptArray(index).PrepareScript

If CallFuncNow <> vbNullString Then
    LoadScript = ScriptArray(index).Run(CStr(CallFuncNow))
End If
End Function

Function ScriptByName(name As String) As Integer
For i = 0 To UBound(ScriptArray)
    If UCase(ScriptArray(i).name) = UCase(name) Then
        ScriptByName = i
        Exit Function
    End If
Next i
ScriptByName = -1
End Function

Sub main()

'reset scripts
ReDim ScriptArray(0)

'Load The Script, And it will run function 'OnLoad()'
Call LoadScript("myscript.txt")
'To Stop the AutoRun Function(stop it from calling OnLoad() _
 do this _
 _
 Call LoadScript("myscript.txt", vbnullstring)
 
'If there is no AutoRun Function to run then you can do
'ScriptArray(ScriptByName("myscript.txt")).Run <function name>

'Remove All Loaded Scripts
For i = 0 To UBound(ScriptArray)
    Set ScriptArray(i) = Nothing
Next i

End Sub

