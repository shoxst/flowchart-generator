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
Implements IBlockBase
Implements IInheritable

Public Super As New BlockBase
Public IsStart As Boolean
Public IsEnd As Boolean

Private Sub Class_Initialize()
    IsStart = False
    IsEnd = False
    Set Super.Symbol = SymbolConfigTerminal.Clone
End Sub

' Override Method
Private Sub IBlockBase_WriteBlock()
    Dim shp As Shape, con As Shape

    ' Set block start point
    Super.SetStartPoint

    ' Create terminal symbol
    Set shp = Super.CreateSymbolShape

    ' Connect with previous block
    If Not IsStart Then
        Super.ConnectWithPrevBlock shp
    End If

    ' Create next connector
    If Not IsEnd Then
        Set con = ShapeUtil.CreateStraightLineConnector
        con.ConnectorFormat.BeginConnect shp, Super.Symbol.Bottom
        Set Super.Connector = con
    End If

    ' Set block end point
    Super.SetEndPoint
    
End Sub

Private Property Get IInheritable_Super() As Object
    Set IInheritable_Super = Super
End Property
