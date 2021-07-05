VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BlockArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Class BlockArea
Public Left As Single
Public Right As Single
Public Top As Single
Public Bottom As Single
Public SymbolWidth As Single
Public SymbolHeight As Single

Private Sub Class_Initialize()
    SymbolWidth = Constant.SYMBOL_DEFAULT_WIDTH
    SymbolHeight = Constant.SYMBOL_DEFAULT_HEIGHT
End Sub

Public Sub SetStartPoint(ByVal PrevArea As BlockArea)
    Left = PrevArea.HorizontalCenter - SymbolWidth / 2
    Top = PrevArea.Bottom
End Sub

Public Sub SetEndPoint()
    Right = Left + SymbolWidth
    Bottom = Top + SymbolHeight + Constant.BLOCK_VERTICAL_MARGIN
End Sub

Public Property Get HorizontalCenter() As Single
    HorizontalCenter = Left + SymbolWidth / 2
End Property

Public Property Get VerticalCenter() As Single
    VerticalCenter = Top + SymbolHeight / 2
End Property
