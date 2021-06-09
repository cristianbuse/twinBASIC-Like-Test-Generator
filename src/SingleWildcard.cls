VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SingleWildcard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IPatternElement

Private Property Get IPatternElement_ElementType() As PatternElementType
    IPatternElement_ElementType = etSingleWildcard
End Property

Public Function Self() As SingleWildcard
    Set Self = Me
End Function

Private Property Get IPatternElement_ToString() As String
    IPatternElement_ToString = Me.ToString
End Property
Public Property Get ToString() As String
    ToString = "?"
End Property

Private Function IPatternElement_RandomMatchingText(ByVal failOneIn As Long) As String
    IPatternElement_RandomMatchingText = Me.RandomMatchingText(failOneIn)
End Function
Public Function RandomMatchingText(ByVal failOneIn As Long) As String 'Cannot fail
    RandomMatchingText = RandCharacter()
End Function
