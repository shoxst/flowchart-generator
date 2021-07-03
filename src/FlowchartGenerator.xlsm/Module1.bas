Attribute VB_Name = "Module1"
Option Explicit

Sub AddShape()
    Dim sp1 As Shape, sp2 As Shape, con As Shape
    
    Call AddRectangle(sp1, 100, 100, "AAA")
    Call AddRectangle(sp2, 100, 200, "BBB")
    Call ConnectShapes(sp1, sp2, con)
    
End Sub

Private Sub AddRectangle _
    (ByRef sp As Shape, ByVal left As Single, ByVal top As Single, ByVal text As String)
    
    Set sp = ActiveSheet.Shapes.AddShape _
        (msoShapeRectangle, left, top, 100, 50)
    With sp
        With .TextFrame
            .Characters.text = text
            .Characters.Font.Color = RGB(0, 0, 0)
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
        End With
        .Line.ForeColor.RGB = RGB(0, 0, 0)
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With
End Sub

Private Sub ConnectShapes _
    (ByRef sp1 As Shape, ByRef sp2 As Shape, ByRef con As Shape)
    
    Set con = ActiveSheet.Shapes.AddConnector _
        (msoConnectorStraight, 1, 1, 1, 1)
    With con
        With .ConnectorFormat
            .BeginConnect sp1, 3
            .EndConnect sp2, 1
        End With
        .Line.ForeColor.RGB = RGB(0, 0, 0)
    End With
End Sub



