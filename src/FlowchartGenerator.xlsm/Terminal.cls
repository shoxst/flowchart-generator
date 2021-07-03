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

Public isStart As Boolean
Public isEnd As Boolean
Public innerText As String
Public left As Single
Public top As Single
Public connector As Shape
Public prevConnector As Shape

Private Sub Class_intiialize()
    isStart = False
    isEnd = False
End Sub

Public Sub AddShapes()
    Dim sp As Shape
    Set sp = ActiveSheet.Shapes.AddShape _
        (msoShapeFlowchartTerminator, left, top, 100, 40)
    With sp
        With .TextFrame
            .Characters.text = innerText
            .Characters.Font.Color = RGB(0, 0, 0)
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
        End With
        .Line.ForeColor.RGB = RGB(0, 0, 0)
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With
    
    If Not isStart Then
        prevConnector.ConnectorFormat.EndConnect sp, 1
    End If
    
    If Not isEnd Then
        Set connector = ActiveSheet.Shapes.AddConnector _
            (msoConnectorStraight, 1, 1, 1, 1)
        With connector
            With .ConnectorFormat
                .BeginConnect sp, 3
            End With
            .Line.ForeColor.RGB = RGB(0, 0, 0)
        End With
    End If
    
End Sub



