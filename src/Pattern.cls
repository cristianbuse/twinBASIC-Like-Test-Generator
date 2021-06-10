VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pattern"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum PatternElementType
    [_etMin] = 1
    etSingleWildcard = 1 '?
    etMultiWildcard = 2  '*
    etSingleDigit = 3    '#
    etCharList = 4       '[]
    etChar = 5           'Any other character
    [_etMax] = etChar
End Enum

Private m_value As String
Private m_elements As Collection

Public Sub Randomize(ByVal elemCount As Long)
    Dim i As Long
    Dim sb As StringBuffer
    Dim elem As IPatternElement
    Dim cList As CharList
    Dim hasMultiWildcard As Boolean 'Track to avoid "*[K-K]" quirk
    
    Set m_elements = New Collection
    Set sb = New StringBuffer
    
    For i = 1 To RandBetween(1, elemCount)
        Set elem = RandomElement()
        Select Case elem.ElementType
        Case etMultiWildcard
            hasMultiWildcard = True
        Case etCharList
            If hasMultiWildcard Then
                Set cList = elem
                cList.Randomize disallowOneElemRanges:=True
            End If
        End Select
        m_elements.Add elem
        sb.Append elem.ToString
    Next i
    m_value = sb.Value
End Sub

Public Property Get ToString() As String
    ToString = m_value
End Property

Private Function RandomElement() As IPatternElement
    Dim et As PatternElementType
    
    et = RandBetween([_etMin], [_etMax])
    Select Case et
    Case etSingleWildcard
        Set RandomElement = New SingleWildcard
    Case etMultiWildcard
        Set RandomElement = New MultiWildcard
    Case etSingleDigit
        Set RandomElement = New SingleDigit
    Case etCharList
        With New CharList
            .Randomize
            Set RandomElement = .Self
        End With
    Case etChar
        With New Char
            .Randomize
            Set RandomElement = .Self
        End With
    End Select
End Function

Public Function RandomMatchingText(ByVal failOneIn As Long) As String
    Dim sb As StringBuffer
    Dim elem As IPatternElement
    
    Set sb = New StringBuffer
    For Each elem In m_elements
        sb.Append elem.RandomMatchingText(failOneIn)
    Next elem
    RandomMatchingText = sb.Value
End Function
