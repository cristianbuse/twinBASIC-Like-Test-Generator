VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CharList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IPatternElement

Private m_elements As Collection
Private m_notIn As Boolean
Private m_value As String
Private m_hasValidElements

Private Property Get IPatternElement_ElementType() As PatternElementType
    IPatternElement_ElementType = etCharList
End Property

Public Function Self() As CharList
    Set Self = Me
End Function

Private Sub Class_Initialize()
    Set m_elements = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_elements = Nothing
End Sub

Public Property Get NotIn() As Boolean
    NotIn = m_notIn
End Property

Public Sub Randomize(Optional ByVal disallowOneElemRanges As Boolean = False)
    m_notIn = (RandBetween(0, 2) = 0)
    
    Dim i As Long
    Dim sb As StringBuffer
    Dim tempMulti As New MultiWildcard
    Dim tempText As String
    
    Set m_elements = New Collection
    Set sb = New StringBuffer
    sb.Append "["
    If m_notIn Then sb.Append "!"
    m_hasValidElements = False
    
    For i = 1 To RandBetween(0, 5)
        If RandBetween(0, 3) = 0 Then
            tempText = tempMulti.RandomMatchingText(0, 10)
            m_elements.Add tempText
            If tempText <> vbNullString Then m_hasValidElements = True
        Else
            With New CharRange
                .Randomize disallowOneElemRanges
                sb.Append .ToString
                m_elements.Add .Self
                If .IsValid Then m_hasValidElements = True
            End With
        End If
    Next i
    
    If RandBetween(1, 20) > 1 Then
        sb.Append "]"
    End If
    m_value = sb.Value
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
    Dim elem As Variant
    Dim shouldFind As Boolean
    Dim isFound As Boolean
    Dim temp As String
    Dim rng As CharRange
    Dim index As Long
    
    
    shouldFind = (RandBetween(1, failOneIn) > 1)
    If m_notIn Then shouldFind = False

    If m_elements.Count = 0 Then
        If shouldFind Then
            RandomMatchingText = vbNullString
        Else
            RandomMatchingText = RandCharacter()
        End If
        Exit Function
    ElseIf Not m_hasValidElements Then
        RandomMatchingText = RandCharacter()
        Exit Function
    ElseIf shouldFind Then
        Do
            index = RandBetween(1, m_elements.Count)
            If IsObject(m_elements(index)) Then
                Set rng = m_elements(index)
                If rng.IsValid Then
                    RandomMatchingText = ChrW$(RandBetween(AscW(rng.StartChar), AscW(rng.EndChar)))
                    Exit Function
                End If
            Else
                temp = m_elements(index)
                If temp <> vbNullString Then
                    RandomMatchingText = Mid$(temp, RandBetween(1, Len(temp)), 1)
                    Exit Function
                End If
            End If
        Loop
    Else
        Const maxIterations = 1000 'Just in case we have a range covering all characters
        Dim iterations
    
        Do
            temp = RandCharacter()
            isFound = False
            For Each elem In m_elements
                If IsObject(elem) Then
                    Set rng = elem
                    If rng.IsCharMatching(temp) Then
                        isFound = True
                        Exit For
                    Else
                        If Not rng.IsValid Then
                            Exit Do
                        End If
                    End If
                Else
                    If InStr(1, elem, temp, vbBinaryCompare) > 0 Then
                        isFound = True
                        Exit For
                    End If
                End If
            Next elem
            iterations = iterations + 1
        Loop Until Not isFound Or iterations > maxIterations
    End If

    RandomMatchingText = temp
End Function
