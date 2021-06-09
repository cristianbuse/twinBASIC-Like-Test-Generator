Attribute VB_Name = "TestLike"
Option Explicit
Option Private Module

Public Const DQ As String = """"
Public Const DDQ As String = """"""
Public Const STAB As String = "    "
Public Const DTAB As String = "        "

Public Sub Main()
    With New frmTBCode
        .Show
    End With
End Sub

Public Function GeneratePatterns(ByVal pCount As Long, ByVal testsPerPattern As Long) As String
    Dim p As Pattern
    Dim i As Long
    Dim j As Long
    Dim testText As String
    Dim coll
    Dim sBinaryCode As StringBuffer
    Dim sTextCode As StringBuffer
    Dim sBinaryCodeError As StringBuffer
    Dim sTextCodeError As StringBuffer
    Dim sAssert As StringBuffer
    Dim result As Boolean
    Dim patternQuoted As String
    Const errCode As String = DTAB & "Debug.Assert Err.Number <> 0" & vbNewLine & DTAB & "Err.Clear" & vbNewLine & vbNewLine
    

    Set p = New Pattern
    Set sBinaryCode = New StringBuffer
    Set sTextCode = New StringBuffer
    Set sBinaryCodeError = New StringBuffer
    Set sTextCodeError = New StringBuffer
    Set sAssert = New StringBuffer
    
    With sBinaryCode
        .Append "Module TestLikeBinary"
        .Append vbNewLine
        .Append STAB & "Option Compare Binary"
        .Append vbNewLine
        .Append STAB & "Public Sub TestLike()"
        .Append vbNewLine
    End With
    With sBinaryCodeError
        .Append DTAB & "' Check error cases"
        .Append vbNewLine
        .Append DTAB & "On Error Resume Next"
        .Append vbNewLine
        .Append vbNewLine
    End With
    With sTextCode
        .Append "Module TestLikeText"
        .Append vbNewLine
        .Append STAB & "Option Compare Text"
        .Append vbNewLine
        .Append STAB & "Public Sub TestLike()"
        .Append vbNewLine
    End With
    sTextCodeError.Append sBinaryCodeError.Value

    On Error Resume Next
    For i = 1 To pCount
        p.Randomize 20
        patternQuoted = Replace(p.ToString, DQ, DDQ)

        For j = 1 To testsPerPattern
            testText = p.RandomMatchingText(failOneIn:=100)
            With sAssert
                .Reset
                .Append DTAB
                .Append "Debug.Assert ("""
                .Append Replace(testText, DQ, DDQ)
                .Append """ Like """
                .Append patternQuoted
                .Append """) = "
            End With
            
            'Binary
            result = IsLikeBinary(testText, p.ToString)
            If Err.Number = 0 Then
                With sBinaryCode
                    .Append sAssert.Value
                    .Append CStr(result)
                    .Append vbNewLine
                End With
            Else
                With sBinaryCodeError
                    .Append sAssert.Value
                    .Append "True"
                    .Append vbNewLine
                    .Append errCode
                End With
                Err.Clear
            End If
            
            'Text
            result = IsLikeText(testText, p.ToString)
            If Err.Number = 0 Then
                With sTextCode
                    .Append sAssert.Value
                    .Append CStr(result)
                    .Append vbNewLine
                End With
            Else
                With sTextCodeError
                    .Append sAssert.Value
                    .Append "True"
                    .Append vbNewLine
                    .Append errCode
                End With
                Err.Clear
            End If
        Next j
    Next i
    On Error GoTo 0
    
    With sBinaryCode
        .Append vbNewLine
        .Append sBinaryCodeError.Value
        .Append DTAB & "On Error GoTo 0"
        .Append vbNewLine
        .Append STAB & "End Sub"
        .Append vbNewLine
        .Append "End Module"
    End With
    With sTextCode
        .Append vbNewLine
        .Append sTextCodeError.Value
        .Append DTAB & "On Error GoTo 0"
        .Append vbNewLine
        .Append STAB & "End Sub"
        .Append vbNewLine
        .Append "End Module"
    End With
    
    GeneratePatterns = sBinaryCode.Value & vbNewLine & vbNewLine & sTextCode.Value
End Function

Public Function RandBetween(ByVal lowerBound As Long, ByVal upperBound As Long) As Long
    RandBetween = VBA.Int((upperBound - lowerBound + 1) * Rnd() + lowerBound)
End Function

Public Function RandCharacter() As String
    RandCharacter = Chr(RandBetween(32, 126)) 'ChrW(RandBetween(-32768, 65535))
End Function
