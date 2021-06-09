Attribute VB_Name = "LikeTextCompare"
Option Explicit
Option Compare Text

Public Function IsLikeText(ByRef theString As String, ByRef Pattern As String) As Boolean
    IsLikeText = (theString Like Pattern)
End Function
