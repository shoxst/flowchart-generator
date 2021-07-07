VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoopBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' class LoopBlock
Implements IBlockBase
Implements IInheritable

Public Super As New BlockBase
Public InnerBlocks As New BlockList
Public Condition As String
Public IsPreTest As Boolean

Private LoopName As String

Private Sub Class_Initialize()
    IsPreTest = True
    Set Super.Symbol = SymbolConfigLoop
End Sub

' Override Method
Private Sub IBlockBase_WriteBlock()
    WriteLoopStart
    WriteInnerBlocks
    WriteLoopEnd
End Sub

Private Sub WriteLoopStart()
    Dim shp As Shape, con As Shape
    
    If IsPreTest Then
        Super.Symbol.Height = Constant.SYMBOL_DEFAULT_HEIGHT + 10
    End If
    
    ' Set block start point
    Super.SetStartPoint

    ' Create loop start symbol
    LoopName = Super.Symbol.InnerText
    If IsPreTest Then
        Super.Symbol.InnerText = LoopName & vbCrLf & Condition
    End If
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

Private Sub WriteInnerBlocks()
    BlockUtil.WriteAllBlocks InnerBlocks
    Set Super.PrevBlock = InnerBlocks.Last
End Sub

Private Sub WriteLoopEnd()
    Dim shp As Shape, con As Shape
    
    If Not IsPreTest Then
        Super.Symbol.Height = Constant.SYMBOL_DEFAULT_HEIGHT + 10
    Else
        Super.Symbol.Height = Constant.SYMBOL_DEFAULT_HEIGHT
    End If
    
    ' Set block start point
    Super.SetStartPoint

    ' Create loop end symbol
    If Not IsPreTest Then
        Super.Symbol.InnerText = LoopName & vbCrLf & Condition
    Else
        Super.Symbol.InnerText = LoopName
    End If
    Set shp = Super.CreateSymbolShape
    shp.Adjustments(1) = 0
    shp.Adjustments(2) = 0.16667

    ' Connect with previous block
    Super.ConnectWithPrevBlock shp

    ' Create next connector
    Set con = ShapeUtil.CreateStraightLineConnector
    con.ConnectorFormat.BeginConnect shp, Super.Symbol.Bottom
    Set Super.Connector = con

    ' Set block end point
    Super.Area.Right = BlockUtil.MaxRightInBlocks(InnerBlocks)
    Super.Area.Bottom = Super.Area.Top + Super.Symbol.Height + Constant.BLOCK_VERTICAL_MARGIN
    
End Sub

Private Property Get IInheritable_Super() As Object
    Set IInheritable_Super = Super
End Property

