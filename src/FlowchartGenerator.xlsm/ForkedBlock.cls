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
Implements IBlockBase
Implements IInheritable

Public Super As New BlockBase
Public MainBlocks As New BlockList
Public BranchBlocks As New BlockList
Public MainLabel As String
Public BranchLabel As String

Private SymbolShape As Shape
Private EndDummy As Shape
Private ArrowConnector As Shape

' Override Method
Private Sub IBlockBase_WriteBlock()
    Dim con As Shape
    
    If Super.Symbol.ShapeType = SymbolConfigDecision.ShapeType Then
        Super.Symbol.Width = Constant.SYMBOL_DEFAULT_WIDTH + 30
        Super.Symbol.Height = Constant.SYMBOL_DEFAULT_HEIGHT + 10
    End If
    
    ' Set block start point
    Super.SetStartPoint

    ' Create decision shape
    Set SymbolShape = Super.CreateSymbolShape

    ' Connect with previous block
    Super.ConnectWithPrevBlock SymbolShape

    ' Create next connector
    Set con = ShapeUtil.CreateStraightLineConnector
    con.ConnectorFormat.BeginConnect SymbolShape, Super.Symbol.Bottom
    Set Super.Connector = con

    ' Set block end point
    Super.SetEndPoint
    
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
        SymbolRight = Super.HorizontalCenter + Super.Symbol.Width / 2
        MaxRight = IIf(MainMaxRight > SymbolRight, MainMaxRight, SymbolRight)
        ArrowConnector.Adjustments(1) = _
            1 - (MaxRight - Super.HorizontalCenter + Constant.BLOCK_HORIZONTAL_MARGIN) * 2 / Super.Symbol.Width
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
        Super.HorizontalCenter, _
        Super.Area.Bottom - Constant.TEXTBOX_DEFAULT_HEIGHT * 0.8, _
        Constant.TEXTBOX_DEFAULT_WIDTH, _
        Constant.TEXTBOX_DEFAULT_HEIGHT)
    ShapeUtil.SetTextboxStyle mtb
    mtb.TextFrame.Characters.text = MainLabel
    
    Set btb = ActiveSheet.Shapes.AddTextbox( _
        msoTextOrientationHorizontal, _
        Super.Area.Right, _
        Super.VerticalCenter - Constant.TEXTBOX_DEFAULT_HEIGHT * 0.8, _
        Constant.TEXTBOX_DEFAULT_WIDTH, _
        Constant.TEXTBOX_DEFAULT_HEIGHT)
    ShapeUtil.SetTextboxStyle btb
    btb.TextFrame.Characters.text = BranchLabel
End Sub

Private Sub WriteMainBlocks()
    BlockUtil.WriteAllBlocks MainBlocks
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
        (msoShapeRectangle, Super.HorizontalCenter, endDummyTop, 0, 0)
    EndDummy.Visible = msoFalse
    Super.Area.Bottom = endDummyTop + Constant.BLOCK_VERTICAL_MARGIN
End Sub

Private Sub WriteElbowConnector()
    Dim con As Shape
    Set con = ShapeUtil.CreateElbowArrowConnector
    con.ConnectorFormat.BeginConnect SymbolShape, Super.Symbol.Right
    Set ArrowConnector = con
End Sub

Private Sub WriteBranchBlocks()
    Dim StartDummy As New BlockBase, con As Shape
    
    Set con = ShapeUtil.CreateElbowLineConnector
    con.ConnectorFormat.BeginConnect SymbolShape, Super.Symbol.Right
    Set StartDummy.Connector = con
    
    StartDummy.Area.Left = Super.Area.Right + Constant.BLOCK_HORIZONTAL_MARGIN
    StartDummy.Area.Bottom = Super.Area.Bottom
    Set BranchBlocks.First.PrevBlock = StartDummy
    
    BlockUtil.WriteAllBlocks BranchBlocks
    
    Dim LastConnector As Shape
    Set LastConnector = BranchBlocks.Last.Connector
       
    Set con = ShapeUtil.CreateElbowArrowConnector
    con.ConnectorFormat.BeginConnect _
        LastConnector.ConnectorFormat.BeginConnectedShape, _
        LastConnector.ConnectorFormat.BeginConnectionSite
    Set ArrowConnector = con
    
    LastConnector.Delete
End Sub

Private Property Get IInheritable_Super() As Object
    Set IInheritable_Super = Super
End Property

