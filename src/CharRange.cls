VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CharRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_start As String
Private m_end As String

Public Function Self() As CharRange
    Set Self = Me
End Function

Public Property Get ToString() As String
    ToString = m_start & "-" & m_end
End Property

Public Property Get StartChar() As String
    StartChar = m_start
End Property

Public Property Get EndChar() As String
    EndChar = m_end
End Property

Public Sub Randomize(Optional ByVal disallowOneElemRanges As Boolean = False)
    Do
        m_start = RandCharacter()
        m_end = RandCharacter()
    Loop Until StrComp(m_start, m_end, vbBinaryCompare) < IIf(disallowOneElemRanges, 0, 1)
    If RandBetween(0, 50) = 0 Then m_start = vbNullString
    If RandBetween(0, 50) = 0 Then m_end = vbNullString
End Sub

Public Function IsCharMatching(ByRef c As String) As Boolean
    If Me.IsValid Then
        IsCharMatching = StrComp(m_start, c, vbBinaryCompare) < 1 And StrComp(c, m_end, vbBinaryCompare) < 1
    End If
End Function

Public Function IsValid() As Boolean
    IsValid = (m_start <> vbNullString And m_end <> vbNullString)
End Function
