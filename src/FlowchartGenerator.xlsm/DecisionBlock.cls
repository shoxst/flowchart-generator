VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DecisionBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' class DecisionBlock
Implements BlockBase

Public MainBlocks As New Collection
Public BranchBlocks As New Collection
Public MainLabel As String
Public BranchLabel As String

Private Super As New BlockBase
Private DecisionShape As Shape
Private EndDummy As Shape
Private ArrowConnector As Shape

' Override Method
Private Sub BlockBase_WriteBlock()
    Dim shp As Shape, con As Shape
    
    Super.Area.SymbolHeight = Super.Area.SymbolHeight + 10
    
    ' Set block start point
    Super.Area.SetStartPoint PrevBlock.Area

    ' Create decision shape
    Set shp = ActiveSheet.Shapes.AddShape _
        (msoShapeFlowchartDecision, Super.Area.Left, Super.Area.Top, Super.Area.SymbolWidth, Super.Area.SymbolHeight)
    DefaultStyle.SetSymbolStyle shp
    shp.TextFrame.Characters.text = Super.InnerText
    Set DecisionShape = shp

    ' Connect with previous block
    Super.PrevBlock.Connector.ConnectorFormat.EndConnect shp, ConnectionSite.Top

    ' Create next connector
    Set con = ActiveSheet.Shapes.AddConnector _
        (msoConnectorStraight, 0, 0, 0, 0)
    DefaultStyle.SetLineConnectorStyle con
    con.ConnectorFormat.BeginConnect shp, ConnectionSite.Bottom
    Set Super.Connector = con

    ' Set block end point
    Super.Area.SetEndPoint
    
    WriteLabels
    
    ' Write main and branch blocks
    WriteMainBlocks
    If BranchBlocks.Count = 0 Then
        WriteElbowConnector
    Else
        WriteBranchBlocks
    End If
    WriteEndDummy
    ArrowConnector.ConnectorFormat.EndConnect EndDummy, ConnectionSite.Right
    If BranchBlocks.Count = 0 Then
        ArrowConnector.Adjustments(1) = -1
    End If
    
    ' Set block end point
    Dim MaxRight As Single
    If BranchBlocks.Count = 0 Then
        MaxRight = Super.Area.Right + BLOCK_HORIZONTAL_MARGIN
    Else
        MaxRight = 0
        Dim e As BlockBase
        For Each e In BranchBlocks
            If e.Area.Right > MaxRight Then
                MaxRight = e.Area.Right
            End If
        Next
    End If
    Super.Area.Right = MaxRight
    
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
    If MainBlocks.Count <> 0 Then
        Dim e As BlockBase
        For Each e In MainBlocks
            e.WriteBlock
        Next
        Set Super.Connector = MainBlocks(MainBlocks.Count).Connector
    End If
End Sub

Private Sub WriteEndDummy()
    Dim mainBottom As Single, branchBottom As Single, endDummyTop As Single
    
    If MainBlocks.Count > 0 Then
        mainBottom = MainBlocks(MainBlocks.Count).Area.Bottom
    Else
        mainBottom = 0
    End If
    If BranchBlocks.Count > 0 Then
        branchBottom = BranchBlocks(BranchBlocks.Count).Area.Bottom
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
    con.ConnectorFormat.BeginConnect DecisionShape, ConnectionSite.Right
    Set ArrowConnector = con
End Sub

Private Sub WriteBranchBlocks()
    Dim StartDummy As New BlockBase, con As Shape
    
    Set con = ActiveSheet.Shapes.AddConnector _
            (msoConnectorElbow, 0, 0, 0, 0)
    DefaultStyle.SetLineConnectorStyle con
    With con
        .ConnectorFormat.BeginConnect DecisionShape, ConnectionSite.Right
    End With
    Set StartDummy.Connector = con
    
    StartDummy.Area.Left = Super.Area.Right + Constant.BLOCK_HORIZONTAL_MARGIN
    StartDummy.Area.Bottom = Super.Area.Bottom
    Set BranchBlocks(1).PrevBlock = StartDummy
    
    Dim e As BlockBase
    For Each e In BranchBlocks
        e.WriteBlock
    Next
    
    Dim LastConnector As Shape
    Set LastConnector = BranchBlocks(BranchBlocks.Count).Connector
       
    Set con = ActiveSheet.Shapes.AddConnector _
            (msoConnectorElbow, 0, 0, 0, 0)
    DefaultStyle.SetArrowConnectorStyle con
    With con
        .ConnectorFormat.BeginConnect LastConnector.ConnectorFormat.BeginConnectedShape, _
            LastConnector.ConnectorFormat.BeginConnectionSite
    End With
    Set ArrowConnector = con
    
    LastConnector.Delete
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

