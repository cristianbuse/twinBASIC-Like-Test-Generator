VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Char"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IPatternElement

Private m_value As String

Private Property Get IPatternElement_ElementType() As PatternElementType
    IPatternElement_ElementType = etChar
End Property

Public Function Self() As Char
    Set Self = Me
End Function

Public Sub Randomize()
    m_value = RandCharacter()
End Sub

Private Property Get IPatternElement_ToString() As String
    IPatternElement_ToString = Me.ToString
End Property
Public Property Get ToString() As String
    ToString = m_value
End Property

Private Function IPatternElement_RandomMatchingText(ByVal failOneIn As Long) As String
    IPatternElement_RandomMatchingText = Me.RandomMatchingText(failOneIn)
End Function
Public Function RandomMatchingText(ByVal failOneIn As Long) As String
    If RandBetween(1, failOneIn) > 1 Then
         RandomMatchingText = m_value
    Else
        'Force to not match
        Dim temp As String
            
        Do
            temp = RandCharacter()
        Loop Until AscW(temp) <> AscW(m_value)
        RandomMatchingText = temp
    End If
End Function
