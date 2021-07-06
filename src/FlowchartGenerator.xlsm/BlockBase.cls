VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BlockBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Class BlockBase
Public InnerText As String
Public Area As New BlockArea
Public Connector As Shape
Public PrevBlock As BlockBase

' Abstract Method
Public Sub WriteBlock()
End Sub

