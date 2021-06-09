Attribute VB_Name = "LikeBinaryCompare"
Option Explicit
Option Compare Binary

Public Function IsLikeBinary(ByRef theString As String, ByRef Pattern As String) As Boolean
    IsLikeBinary = (theString Like Pattern)
End Function
