Attribute VB_Name = "DefaultStyle"
Option Explicit

Public Sub SetSymbolStyle(ByRef Symbol As Shape)
    With Symbol
        With .TextFrame
            .Characters.Font.Color = RGB(0, 0, 0)
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
        End With
        With .Line
            .ForeColor.RGB = RGB(0, 0, 0)
            .Weight = 0.75
        End With
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With
End Sub

Public Sub SetLineConnectorStyle(ByRef Connector As Shape)
    With Connector.Line
        .ForeColor.RGB = RGB(0, 0, 0)
        .Weight = 0.75
    End With
End Sub

Public Sub SetArrowConnectorStyle(ByRef Connector As Shape)
    With Connector.Line
        .EndArrowheadStyle = msoArrowheadOpen
        .ForeColor.RGB = RGB(0, 0, 0)
        .Weight = 0.75
    End With
End Sub

Public Sub SetTextboxStyle(ByRef Textbox As Shape)
    With Textbox
        With .TextFrame
            .Characters.Font.Color = RGB(0, 0, 0)
            .HorizontalAlignment = xlHAlignLeft
            .VerticalAlignment = xlVAlignCenter
        End With
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
    End With
End Sub

