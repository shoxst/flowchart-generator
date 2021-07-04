VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BlockBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Class BlockBase
Public InnerText As String
Public Area As New BlockArea
Public Connector As Shape
Public PrevBlock As BlockBase

' Abstract Method
Public Sub WriteBlock()
End Sub

' Virtual Method
Public Sub SetDefaultShapeStyle(ByRef Shape As Shape)
    With Shape
        With .TextFrame
            .Characters.text = InnerText
            .Characters.Font.Color = RGB(0, 0, 0)
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
        End With
        .Line.ForeColor.RGB = RGB(0, 0, 0)
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With
End Sub

Public Sub SetDefaultConnectorColor(ByRef Connector As Shape)
    Connector.Line.ForeColor.RGB = RGB(0, 0, 0)
End Sub
