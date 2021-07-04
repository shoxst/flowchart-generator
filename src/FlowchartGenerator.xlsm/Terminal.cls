VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Terminal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Class Terminal
Implements Symbol

Public IsStart As Boolean
Public IsEnd As Boolean
Public InnerText As String
Public Left As Single
Public Top As Single
Public Bottom As Single
Public Connector As Shape
Public prevConnector As Shape

Private Sub Class_Intiialize()
    IsStart = False
    IsEnd = False
End Sub

Public Sub AddShapes()
    Dim sp As Shape, height As Single
    
    height = 40
    
    Set sp = ActiveSheet.Shapes.AddShape _
        (msoShapeFlowchartTerminator, Left, Top, 100, height)
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
    
    If Not IsStart Then
        prevConnector.ConnectorFormat.EndConnect sp, 1
    End If
    
    If Not IsEnd Then
        Set Connector = ActiveSheet.Shapes.AddConnector _
            (msoConnectorStraight, 0, 0, 0, 0)
        With Connector
            With .ConnectorFormat
                .BeginConnect sp, 3
            End With
            .Line.ForeColor.RGB = RGB(0, 0, 0)
        End With
    End If
    
    Bottom = Top + height + 10
End Sub

Private Sub Symbol_AddShapes()
    Me.AddShapes
End Sub



