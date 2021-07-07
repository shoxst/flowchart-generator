VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SymbolConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ShapeType As MsoAutoShapeType
Public InnerText As String
Public Top As Integer
Public Left As Integer
Public Bottom As Integer
Public Right As Integer
Public Width As Single
Public Height As Single

Private Sub Class_Initialize()
    Top = 1
    Left = 2
    Bottom = 3
    Right = 4
    Width = Constant.SYMBOL_DEFAULT_WIDTH
    Height = Constant.SYMBOL_DEFAULT_HEIGHT
End Sub

Public Function Clone() As SymbolConfig
    Dim sym As New SymbolConfig
    sym.ShapeType = ShapeType
    sym.InnerText = InnerText
    sym.Top = Top
    sym.Left = Left
    sym.Bottom = Bottom
    sym.Right = Right
    sym.Width = Width
    sym.Height = Height
    Set Clone = sym
End Function

