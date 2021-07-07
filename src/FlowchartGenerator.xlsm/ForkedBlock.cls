VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ForkedBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' class ForkedBlock
Implements BlockBase

Public MainBlocks As New BlockList
Public BranchBlocks As New BlockList
Public MainLabel As String
Public BranchLabel As String

Private Super As New BlockBase
Private SymbolShape As Shape
Private EndDummy As Shape
Private ArrowConnector As Shape

' Override Method
Private Sub BlockBase_WriteBlock()
    Dim shp As Shape, con As Shape
    
    Super.Area.SymbolWidth = Constant.SYMBOL_DEFAULT_WIDTH + 30
    Super.Area.SymbolHeight = Constant.SYMBOL_DEFAULT_HEIGHT + 10
    
    ' Set block start point
    Super.Area.SetStartPoint PrevBlock.Area

    ' Create decision shape
    Set shp = BlockBase_CreateSymbolShape
    shp.TextFrame.Characters.text = Super.InnerText
    Set SymbolShape = shp

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
    
    WriteLabels
    
    ' Write main and branch blocks
    Dim MainMaxRight As Single
    
    If MainBlocks.Count <> 0 Then
        WriteMainBlocks
        MainMaxRight = BlockUtil.MaxRightInBlocks(MainBlocks)
        Super.Area.Right = MainMaxRight
    End If
    If BranchBlocks.Count = 0 Then
        WriteElbowConnector
    Else
        WriteBranchBlocks
    End If
    WriteEndDummy
    ArrowConnector.ConnectorFormat.EndConnect EndDummy, 4
    
    If BranchBlocks.Count = 0 Then
        Dim SymbolRight As Single, MaxRight As Single
        SymbolRight = Super.Area.HorizontalCenter + Super.Area.SymbolWidth / 2
        MaxRight = IIf(MainMaxRight > SymbolRight, MainMaxRight, SymbolRight)
        ArrowConnector.Adjustments(1) = _
            1 - (MaxRight - Super.Area.HorizontalCenter + Constant.BLOCK_HORIZONTAL_MARGIN) * 2 / Super.Area.SymbolWidth
    End If
    
    ' Set block end point
    If BranchBlocks.Count <> 0 Then
        Super.Area.Right = BlockUtil.MaxRightInBlocks(BranchBlocks)
    End If
    
End Sub

Private Sub WriteLabels()
    Dim mtb As Shape, btb As Shape
    
    Set mtb = ActiveSheet.Shapes.AddTextbox( _
        msoTextOrientationHorizontal, _
        Super.Area.HorizontalCenter, _
        Super.Area.Bottom - Constant.TEXTBOX_DEFAULT_HEIGHT * 0.8, _
        Constant.TEXTBOX_DEFAULT_WIDTH, _
        Constant.TEXTBOX_DEFAULT_HEIGHT)
    DefaultStyle.SetTextboxStyle mtb
    mtb.TextFrame.Characters.text = MainLabel
    
    Set btb = ActiveSheet.Shapes.AddTextbox( _
        msoTextOrientationHorizontal, _
        Super.Area.Right, _
        Super.Area.VerticalCenter - Constant.TEXTBOX_DEFAULT_HEIGHT * 0.8, _
        Constant.TEXTBOX_DEFAULT_WIDTH, _
        Constant.TEXTBOX_DEFAULT_HEIGHT)
    DefaultStyle.SetTextboxStyle btb
    btb.TextFrame.Characters.text = BranchLabel
End Sub

Private Sub WriteMainBlocks()
    Dim e As BlockBase
    For Each e In MainBlocks.Items
        e.WriteBlock
    Next
    Set Super.Connector = MainBlocks.Last.Connector
End Sub

Private Sub WriteEndDummy()
    Dim mainBottom As Single, branchBottom As Single, endDummyTop As Single
    
    If MainBlocks.Count > 0 Then
        mainBottom = MainBlocks.Last.Area.Bottom
    Else
        mainBottom = 0
    End If
    If BranchBlocks.Count > 0 Then
        branchBottom = BranchBlocks.Last.Area.Bottom
    Else
        branchBottom = 0
    End If
    endDummyTop = IIf(mainBottom > branchBottom, mainBottom, branchBottom)
    
    Set EndDummy = ActiveSheet.Shapes.AddShape _
        (msoShapeRectangle, Super.Area.HorizontalCenter, endDummyTop, 0, 0)
    EndDummy.Visible = msoFalse
    Super.Area.Bottom = endDummyTop + Constant.BLOCK_VERTICAL_MARGIN
End Sub

Private Sub WriteElbowConnector()
    Dim con As Shape
    Set con = ActiveSheet.Shapes.AddConnector _
            (msoConnectorElbow, 0, 0, 0, 0)
    DefaultStyle.SetArrowConnectorStyle con
    con.ConnectorFormat.BeginConnect SymbolShape, Super.Symbol.Right
    Set ArrowConnector = con
End Sub

Private Sub WriteBranchBlocks()
    Dim StartDummy As New BlockBase, con As Shape
    
    Set con = ActiveSheet.Shapes.AddConnector _
            (msoConnectorElbow, 0, 0, 0, 0)
    DefaultStyle.SetLineConnectorStyle con
    con.ConnectorFormat.BeginConnect SymbolShape, Super.Symbol.Right
    Set StartDummy.Connector = con
    
    StartDummy.Area.Left = Super.Area.Right + Constant.BLOCK_HORIZONTAL_MARGIN
    StartDummy.Area.Bottom = Super.Area.Bottom
    Set BranchBlocks.First.PrevBlock = StartDummy
    
    Dim e As BlockBase
    For Each e In BranchBlocks.Items
        e.WriteBlock
    Next
    
    Dim LastConnector As Shape
    Set LastConnector = BranchBlocks.Last.Connector
       
    Set con = ActiveSheet.Shapes.AddConnector _
            (msoConnectorElbow, 0, 0, 0, 0)
    DefaultStyle.SetArrowConnectorStyle con
    con.ConnectorFormat.BeginConnect _
        LastConnector.ConnectorFormat.BeginConnectedShape, _
        LastConnector.ConnectorFormat.BeginConnectionSite
    Set ArrowConnector = con
    
    LastConnector.Delete
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

