VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BlockList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Blocks As New Collection

Public Property Get Count() As Long
    Count = Blocks.Count
End Property

Public Property Get Item(ByVal Index As Long)
    Set Item = Blocks(Index)
End Property

Public Property Get Items() As Collection
    Set Items = Blocks
End Property

Public Property Get First() As BlockBase
    Set First = Blocks(1)
End Property

Public Property Get Last() As BlockBase
    Set Last = Blocks(Blocks.Count)
End Property

Public Sub Add(ByVal Block As BlockBase)
    Blocks.Add Block
End Sub

Public Sub Remove(ByVal Index As Long)
    Blocks.Remove Index
End Sub
