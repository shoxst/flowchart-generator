Attribute VB_Name = "ShapeUtil"
Option Explicit

Public Function CreateStraightLineConnector() As Shape
    Dim con As Shape
    Set con = ActiveSheet.Shapes.AddConnector _
        (msoConnectorStraight, 0, 0, 0, 0)
    SetLineConnectorStyle con
    Set CreateStraightLineConnector = con
End Function

Public Function CreateElbowLineConnector() As Shape
    Dim con As Shape
    Set con = ActiveSheet.Shapes.AddConnector _
            (msoConnectorElbow, 0, 0, 0, 0)
    SetLineConnectorStyle con
    Set CreateElbowLineConnector = con
End Function

Public Function CreateElbowArrowConnector() As Shape
    Dim con As Shape
    Set con = ActiveSheet.Shapes.AddConnector _
            (msoConnectorElbow, 0, 0, 0, 0)
    SetArrowConnectorStyle con
    Set CreateElbowArrowConnector = con
End Function

Public Sub SetSymbolStyle(ByRef Symbol As Shape)
    With Symbol
        With .TextFrame
            .Characters.Font.Color = RGB(0, 0, 0)
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
        End With
        With .line
            .ForeColor.RGB = RGB(0, 0, 0)
            .Weight = 0.75
        End With
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With
End Sub

Public Sub SetLineConnectorStyle(ByRef Connector As Shape)
    With Connector.line
        .ForeColor.RGB = RGB(0, 0, 0)
        .Weight = 0.75
    End With
End Sub

Public Sub SetArrowConnectorStyle(ByRef Connector As Shape)
    With Connector.line
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
        .line.Visible = msoFalse
        .Fill.Visible = msoFalse
    End With
End Sub

