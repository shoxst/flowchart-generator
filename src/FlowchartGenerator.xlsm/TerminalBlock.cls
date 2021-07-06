VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TerminalBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' class TerminalBlock
Implements BlockBase

Public IsStart As Boolean
Public IsEnd As Boolean

Private Super As New BlockBase

Private Sub Class_Initialize()
    IsStart = False
    IsEnd = False
End Sub

' Override Method
Private Sub BlockBase_WriteBlock()
    Dim shp As Shape, con As Shape

    ' Set block start point
    Super.Area.SetStartPoint PrevBlock.Area

    ' Create terminal symbol
    Set shp = ActiveSheet.Shapes.AddShape _
        (symbolTypeTerminal, Super.Area.Left, Super.Area.Top, Super.Area.SymbolWidth, Super.Area.SymbolHeight)
    DefaultStyle.SetSymbolStyle shp
    shp.TextFrame.Characters.text = Super.InnerText

    ' Connect with previous block
    If Not IsStart Then
        Super.PrevBlock.Connector.ConnectorFormat.EndConnect shp, 1
    End If

    ' Create next connector
    If Not IsEnd Then
        Set con = ActiveSheet.Shapes.AddConnector _
            (msoConnectorStraight, 0, 0, 0, 0)
        DefaultStyle.SetLineConnectorStyle con
        con.ConnectorFormat.BeginConnect shp, 3
        Set Super.Connector = con
    End If

    ' Set block end point
    Super.Area.SetEndPoint
    
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

