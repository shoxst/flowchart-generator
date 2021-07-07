Attribute VB_Name = "BlockUtil"
Option Explicit

Public Function MaxRightInBlocks(ByVal Blocks As BlockList) As Single
    Dim MaxRight As Single
    MaxRight = 0
    Dim e As BlockBase
    For Each e In Blocks.Items
        If e.Area.Right > MaxRight Then
            MaxRight = e.Area.Right
        End If
    Next
    MaxRightInBlocks = MaxRight
End Function

Public Sub WriteAllBlocks(ByVal Blocks As BlockList)
    Dim e As BlockBase
    For Each e In Blocks.Items
        e.WriteBlock
    Next
End Sub
