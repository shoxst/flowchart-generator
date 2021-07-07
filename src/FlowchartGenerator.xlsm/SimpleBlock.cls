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
Implements BlockBase

Private Super As New BlockBase

' Override Method
Private Sub BlockBase_WriteBlock()
    Dim shp As Shape, con As Shape
    
    ' Set block start point
    Super.Area.SetStartPoint PrevBlock.Area

    ' Create process symbol
    Set shp = BlockBase_CreateSymbolShape
    shp.TextFrame.Characters.text = Super.InnerText

    ' Connect with previous block
    BlockBase_ConnectWithPrevBlock shp

    ' Create next connector
    Set con = ActiveSheet.Shapes.AddConnector _
        (msoConnectorStraight, 0, 0, 0, 0)
    DefaultStyle.SetLineConnectorStyle con
    con.ConnectorFormat.BeginConnect shp, Super.Symbol.Bottom
    Set Super.Connector = con

    ' Set block end point
    Super.Area.SetEndPoint
    
End Sub

Private Function BlockBase_CreateSymbolShape() As Shape
    Set BlockBase_CreateSymbolShape = Super.CreateSymbolShape
End Function

Private Sub BlockBase_ConnectWithPrevBlock(ByRef shp As Shape)
    Super.ConnectWithPrevBlock shp
End Sub

' Interface
Public Sub WriteBlock()
    BlockBase_WriteBlock
End Sub

' Public Getter/Setter/Letter
Public Property Get Symbol() As SymbolConfig
    Set Symbol = BlockBase_Symbol
End Property

Public Property Set Symbol(ByVal Symbol As SymbolConfig)
    Set BlockBase_Symbol = Symbol
End Property

Public Property Get Area() As BlockArea
    Set Area = BlockBase_Area
End Property

Public Property Set Area(ByVal Area As BlockArea)
    Set BlockBase_Area = Area
End Property

Public Property Get InnerText() As String
    InnerText = BlockBase_InnerText
End Property

Public Property Let InnerText(ByVal InnerText As String)
    BlockBase_InnerText = InnerText
End Property

Public Property Get Connector() As Shape
    Set Connector = BlockBase_Connector
End Property

Public Property Set Connector(ByVal Connector As Shape)
    Set BlockBase_Connector = Connector
End Property

Public Property Get PrevBlock() As BlockBase
    Set PrevBlock = BlockBase_PrevBlock
End Property

Public Property Set PrevBlock(ByVal PrevBlock As BlockBase)
    Set BlockBase_PrevBlock = PrevBlock
End Property

' Override Getter/Setter/Letter
Private Property Get BlockBase_Symbol() As SymbolConfig
    Set BlockBase_Symbol = Super.Symbol
End Property

Private Property Set BlockBase_Symbol(ByVal Symbol As SymbolConfig)
    Set Super.Symbol = Symbol
End Property

Private Property Get BlockBase_Area() As BlockArea
    Set BlockBase_Area = Super.Area
End Property

Private Property Set BlockBase_Area(ByVal Area As BlockArea)
    Set Super.Area = Area
End Property

Private Property Get BlockBase_InnerText() As String
    BlockBase_InnerText = Super.InnerText
End Property

Private Property Let BlockBase_InnerText(ByVal InnerText As String)
    Super.InnerText = InnerText
End Property

Private Property Get BlockBase_Connector() As Shape
    Set BlockBase_Connector = Super.Connector
End Property

Private Property Set BlockBase_Connector(ByVal Connector As Shape)
    Set Super.Connector = Connector
End Property

Private Property Get BlockBase_PrevBlock() As BlockBase
    Set BlockBase_PrevBlock = Super.PrevBlock
End Property

Private Property Set BlockBase_PrevBlock(ByVal PrevBlock As BlockBase)
    Set Super.PrevBlock = PrevBlock
End Property
