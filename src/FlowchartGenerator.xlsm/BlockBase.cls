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
Public Symbol As SymbolConfig
Public InnerText As String
Public Area As New BlockArea
Public Connector As Shape
Public PrevBlock As BlockBase

' Abstract Method
Public Sub WriteBlock()
End Sub

' Virtual Method
Public Function CreateSymbolShape() As Shape
    Dim shp As Shape
    Set shp = ActiveSheet.Shapes.AddShape _
        (Symbol.ShapeType, Area.Left, Area.Top, Area.SymbolWidth, Area.SymbolHeight)
    DefaultStyle.SetSymbolStyle shp
    Set CreateSymbolShape = shp
End Function

Public Sub ConnectWithPrevBlock(ByRef shp As Shape)
    PrevBlock.Connector.ConnectorFormat.EndConnect shp, Symbol.Top
End Sub

