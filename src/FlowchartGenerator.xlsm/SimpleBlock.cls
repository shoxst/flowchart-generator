VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SimpleBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' class SimpleBlock
Implements IBlockBase
Implements IInheritable

Public Super As New BlockBase

' Override Method
Private Sub IBlockBase_WriteBlock()
    Dim shp As Shape, con As Shape
    
    ' Set block start point
    Super.SetStartPoint

    ' Create process symbol
    Set shp = Super.CreateSymbolShape

    ' Connect with previous block
    Super.ConnectWithPrevBlock shp

    ' Create next connector
    Set con = ShapeUtil.CreateStraightLineConnector
    con.ConnectorFormat.BeginConnect shp, Super.Symbol.Bottom
    Set Super.Connector = con

    ' Set block end point
    Super.SetEndPoint
    
End Sub

Private Property Get IInheritable_Super() As Object
    Set IInheritable_Super = Super
End Property

