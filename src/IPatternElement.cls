VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPatternElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get ElementType() As PatternElementType
    Err.Raise 5, TypeName(Me) & ".ElementType(Get)", "This class is an interface"
End Property

Public Property Get ToString() As String
    Err.Raise 5, TypeName(Me) & ".ToString(Get)", "This class is an interface"
End Property

Public Function RandomMatchingText(ByVal failOneIn As Long) As String
    Err.Raise 5, TypeName(Me) & ".RandomMatchingText", "This class is an interface"
End Function
