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
Implements BlockBase

Public InnerBlocks As New BlockList
Public Condition As String
Public IsPreTest As Boolean

Private Super As New BlockBase

Private Sub Class_Initialize()
    IsPreTest = True
End Sub

' Override Method
Private Sub BlockBase_WriteBlock()
    
    WriteLoopStart
    
    WriteInnerBlocks
    
    WriteLoopEnd
    
End Sub

Private Sub WriteLoopStart()
    Dim shp As Shape, con As Shape
    
    If IsPreTest Then
        Super.Area.SymbolHeight = Constant.SYMBOL_DEFAULT_HEIGHT + 10
    Else
        Super.Area.SymbolHeight = Constant.SYMBOL_DEFAULT_HEIGHT
    End If
    
    ' Set block start point
    Super.Area.SetStartPoint PrevBlock.Area

    ' Create loop start symbol
    Set shp = ActiveSheet.Shapes.AddShape _
        (symbolTypeLoop, Super.Area.Left, Super.Area.Top, Super.Area.SymbolWidth, Super.Area.SymbolHeight)
    DefaultStyle.SetSymbolStyle shp
    If IsPreTest Then
        shp.TextFrame.Characters.text = Super.InnerText & vbCrLf & Condition
    Else
        shp.TextFrame.Characters.text = Super.InnerText
    End If

    ' Connect with previous block
    Super.PrevBlock.Connector.ConnectorFormat.EndConnect shp, 4

    ' Create next connector
    Set con = ActiveSheet.Shapes.AddConnector _
        (msoConnectorStraight, 0, 0, 0, 0)
    DefaultStyle.SetLineConnectorStyle con
    con.ConnectorFormat.BeginConnect shp, 2
    Set Super.Connector = con

    ' Set block end point
    Super.Area.SetEndPoint
End Sub

Private Sub WriteInnerBlocks()
    Dim e As BlockBase
    For Each e In InnerBlocks.Items
        e.WriteBlock
    Next
    Set Super.PrevBlock = InnerBlocks.Last
End Sub

Private Sub WriteLoopEnd()
    Dim shp As Shape, con As Shape
    
    If Not IsPreTest Then
        Super.Area.SymbolHeight = Constant.SYMBOL_DEFAULT_HEIGHT + 10
    Else
        Super.Area.SymbolHeight = Constant.SYMBOL_DEFAULT_HEIGHT
    End If
    
    ' Set block start point
    Super.Area.SetStartPoint PrevBlock.Area

    ' Create loop start symbol
    Set shp = ActiveSheet.Shapes.AddShape _
        (symbolTypeLoop, Super.Area.Left, Super.Area.Top, Super.Area.SymbolWidth, Super.Area.SymbolHeight)
    DefaultStyle.SetSymbolStyle shp
    shp.Adjustments(1) = 0
    shp.Adjustments(2) = 0.16667
    If Not IsPreTest Then
        shp.TextFrame.Characters.text = Super.InnerText & vbCrLf & Condition
    Else
        shp.TextFrame.Characters.text = Super.InnerText
    End If

    ' Connect with previous block
    Super.PrevBlock.Connector.ConnectorFormat.EndConnect shp, 4

    ' Create next connector
    Set con = ActiveSheet.Shapes.AddConnector _
        (msoConnectorStraight, 0, 0, 0, 0)
    DefaultStyle.SetLineConnectorStyle con
    con.ConnectorFormat.BeginConnect shp, 2
    Set Super.Connector = con

    ' Set block end point
    Super.Area.Right = BlockUtil.MaxRightInBlocks(InnerBlocks)
    Super.Area.Bottom = Super.Area.Top + Super.Area.SymbolHeight + Constant.BLOCK_VERTICAL_MARGIN
    
End Sub

' Interface
Public Sub WriteBlock()
    BlockBase_WriteBlock
End Sub

' Public Getter/Setter/Letter
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

