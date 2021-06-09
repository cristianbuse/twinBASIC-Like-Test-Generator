VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTBCode 
   Caption         =   "tB Code"
   ClientHeight    =   9435
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13095
   OleObjectBlob   =   "frmTBCode.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTBCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnClose_Click()
    Me.Hide
End Sub

Private Sub btnGenerate_Click()
    Dim pCount As Long
    Dim tCount As Long
    
    If IsNumeric(tboxPatternsCount.Value) Then pCount = CLng(tboxPatternsCount.Value)
    If IsNumeric(tboxTestsCount.Value) Then tCount = CLng(tboxTestsCount.Value)
    
    If pCount < 1 Then pCount = 10
    If tCount < 1 Then pCount = 3

    lblStatus.Caption = "Generated " & pCount & " patterns with " _
        & tCount & " tests per pattern (total: " & pCount * tCount & ")"
    
    tboxCode.Enabled = False
    tboxCode.Value = GeneratePatterns(pCount, tCount)
    tboxCode.Enabled = True
End Sub
