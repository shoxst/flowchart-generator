VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Decision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Class Decision
Implements Symbol

Public innerText As String
Public left As Single
Public top As Single
Public bottom As Single
Public connector As Shape
Public prevConnector As Shape

Public mainSymbols As New Collection
Public branchSymbols As New Collection

Private dummy As Shape

Public Sub AddShapes()
    Dim sp As Shape, height As Single
    height = 50
    
    ' Decision Symbol
    Set sp = ActiveSheet.Shapes.AddShape _
        (msoShapeFlowchartDecision, left, top, 100, height)
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
    
    prevConnector.ConnectorFormat.EndConnect sp, 1
    
    bottom = top + height + 10
    
    Set connector = ActiveSheet.Shapes.AddConnector _
        (msoConnectorStraight, 1, 1, 1, 1)
    With connector
        .ConnectorFormat.BeginConnect sp, 3
        .Line.ForeColor.RGB = RGB(0, 0, 0)
    End With
       
    ' Main Symbols
    Call AddMainSymbols
    
    ' Dummy
    Call AddDummy
    
    ' Branch Symbols
    If branchSymbols.Count = 0 Then
        With ActiveSheet.Shapes.AddConnector _
            (msoConnectorElbow, 0, 0, 0, 0)
            With .Line
                .EndArrowheadStyle = msoArrowheadTriangle
                .ForeColor.RGB = RGB(0, 0, 0)
            End With
            With .ConnectorFormat
                .BeginConnect sp, 4
                .EndConnect dummy, 4
            End With
            .Adjustments(1) = -1
        End With
    End If
    
End Sub

Private Sub AddMainSymbols()
    If mainSymbols.Count <> 0 Then
        Dim i As Integer
        For i = 1 To mainSymbols.Count
            Set mainSymbols(i).prevConnector = connector
            mainSymbols(i).left = left
            mainSymbols(i).top = bottom
            mainSymbols(i).AddShapes
            Set connector = mainSymbols(i).connector
            bottom = mainSymbols(i).bottom
        Next
    End If
End Sub

Private Sub AddDummy()
    Set dummy = ActiveSheet.Shapes.AddShape _
        (msoShapeRectangle, left + 50, bottom, 0, 10)
    With dummy
        .Visible = msoFalse
    End With
    bottom = bottom + 20
End Sub

Private Sub Symbol_AddShapes()
    Me.AddShapes
End Sub



