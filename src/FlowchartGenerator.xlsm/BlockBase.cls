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

Public This As IBlockBase
Public Symbol As New SymbolConfig
Public Area As New BlockArea
Public Connector As Shape
Public PrevBlock As BlockBase

Public Sub Initialize(ByVal text As String, ByVal block As BlockBase)
    Symbol.InnerText = text
    Set PrevBlock = block
End Sub

Public Sub WriteBlock()
    This.WriteBlock
End Sub

Public Function CreateSymbolShape() As Shape
    Dim shp As Shape
    Set shp = ActiveSheet.Shapes.AddShape _
        (Symbol.ShapeType, Area.Left, Area.Top, Symbol.Width, Symbol.Height)
    ShapeUtil.SetSymbolStyle shp
    shp.TextFrame.Characters.text = Symbol.InnerText
    Set CreateSymbolShape = shp
End Function

Public Sub ConnectWithPrevBlock(ByRef shp As Shape)
    PrevBlock.Connector.ConnectorFormat.EndConnect shp, Symbol.Top
End Sub

Public Sub SetStartPoint()
    Area.Left = PrevBlock.HorizontalCenter - Symbol.Width / 2
    Area.Top = PrevBlock.Area.Bottom
End Sub

Public Sub SetEndPoint()
    Area.Right = Area.Left + Symbol.Width
    Area.Bottom = Area.Top + Symbol.Height + Constant.BLOCK_VERTICAL_MARGIN
End Sub

Public Property Get HorizontalCenter() As Single
    HorizontalCenter = Area.Left + Symbol.Width / 2
End Property

Public Property Get VerticalCenter() As Single
    VerticalCenter = Area.Top + Symbol.Height / 2
End Property

