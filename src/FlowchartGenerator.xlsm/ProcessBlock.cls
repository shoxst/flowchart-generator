VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProcessBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' class Process Block
Implements BlockBase

Private Super As New BlockBase

' Override Method
Private Sub BlockBase_WriteBlock()
    Dim sp As Shape, con As Shape
    Dim height As Single, width As Single
    
    width = 100
    height = 40

    ' Set block area start
    Super.Area.Left = Super.PrevBlock.Area.Left
    Super.Area.Top = Super.PrevBlock.Area.Bottom

    ' Create process symbols
    Set sp = ActiveSheet.Shapes.AddShape _
        (msoShapeFlowchartProcess, Super.Area.Left, Super.Area.Top, width, height)
    Call BlockBase_SetDefaultShapeStyle(sp)

    ' Connect with previous block
    Super.PrevBlock.Connector.ConnectorFormat.EndConnect sp, 1

    ' Create next connector
    Set con = ActiveSheet.Shapes.AddConnector _
        (msoConnectorStraight, 0, 0, 0, 0)
    Call BlockBase_SetDefaultConnectorColor(con)
    con.ConnectorFormat.BeginConnect sp, 3
    Set Super.Connector = con

    ' Set block area end
    Super.Area.Bottom = Super.Area.Top + height + 10
    
End Sub

Private Sub BlockBase_SetDefaultShapeStyle(ByRef Shape As Shape)
    Call Super.SetDefaultShapeStyle(Shape)
End Sub

Private Sub BlockBase_SetDefaultConnectorColor(ByRef Connector As Shape)
    Call Super.SetDefaultConnectorColor(Connector)
End Sub

' Interface
Public Sub WriteBlock()
    Call BlockBase_WriteBlock
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
