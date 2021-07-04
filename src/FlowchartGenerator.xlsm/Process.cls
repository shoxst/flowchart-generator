VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Process"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Class Process
Implements Symbol

Public InnerText As String
Public Left As Single
Public Top As Single
Public Bottom As Single
Public Connector As Shape
Public prevConnector As Shape

Public Sub AddShapes()
    Dim sp As Shape, height As Single
    height = 40
    
    Set sp = ActiveSheet.Shapes.AddShape _
        (msoShapeFlowchartProcess, Left, Top, 100, height)
    With sp
        With .TextFrame
            .Characters.text = InnerText
            .Characters.Font.Color = RGB(0, 0, 0)
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
        End With
        .Line.ForeColor.RGB = RGB(0, 0, 0)
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With
    
    prevConnector.ConnectorFormat.EndConnect sp, 1
    
    Set Connector = ActiveSheet.Shapes.AddConnector _
        (msoConnectorStraight, 0, 0, 0, 0)
    With Connector
        With .ConnectorFormat
            .BeginConnect sp, 3
        End With
        .Line.ForeColor.RGB = RGB(0, 0, 0)
    End With
    
    Bottom = Top + height + 10
End Sub

Private Sub Symbol_AddShapes()
    Me.AddShapes
End Sub





