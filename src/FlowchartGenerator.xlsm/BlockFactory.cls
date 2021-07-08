VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BlockFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Class BlockFactory

Public Function CreateStartDummyBlock(ByVal Left As Single, ByVal Bottom As Single)
    Dim block As New TerminalBlock
    block.Super.Area.Left = Left
    block.Super.Area.Bottom = Bottom
    Set CreateStartDummyBlock = CreateBase(block)
End Function

Public Function CreateTerminalBlock(ByVal PrevBlock As BlockBase, ByVal InnerText As String, ByVal IsStart As Boolean) As BlockBase
    Dim block As New TerminalBlock
    block.Super.Initialize InnerText, PrevBlock
    If IsStart Then
        block.IsStart = True
    Else
        block.IsEnd = True
    End If
    Set CreateTerminalBlock = CreateBase(block)
End Function

Public Function CreateProcessBlock(ByVal PrevBlock As BlockBase, ByVal cols As Collection) As BlockBase
    Dim block As New SimpleBlock
    Set block.Super.Symbol = SymbolConfigProcess.Clone
    block.Super.Initialize cols(1), PrevBlock
    Set CreateProcessBlock = CreateBase(block)
End Function

Public Function CreatePredefinedBlock(ByVal PrevBlock As BlockBase, ByVal cols As Collection) As BlockBase
    Dim block As New SimpleBlock
    Set block.Super.Symbol = SymbolConfigPredifined.Clone
    block.Super.Initialize cols(2), PrevBlock
    Set CreatePredefinedBlock = CreateBase(block)
End Function

Public Function CreateDocumentBlock(ByVal PrevBlock As BlockBase, ByVal cols As Collection) As BlockBase
    Dim block As New SimpleBlock
    Set block.Super.Symbol = SymbolConfigDocument.Clone
    block.Super.Initialize cols(2), PrevBlock
    Set CreateDocumentBlock = CreateBase(block)
End Function

Public Function CreateDisplayBlock(ByVal PrevBlock As BlockBase, ByVal cols As Collection) As BlockBase
    Dim block As New SimpleBlock
    Set block.Super.Symbol = SymbolConfigDisplay.Clone
    block.Super.Initialize cols(2), PrevBlock
    Set CreateDisplayBlock = CreateBase(block)
End Function

Public Function CreateManualInputBlock(ByVal PrevBlock As BlockBase, ByVal cols As Collection) As BlockBase
    Dim block As New SimpleBlock
    Set block.Super.Symbol = SymbolConfigManualInput.Clone
    block.Super.Initialize cols(2), PrevBlock
    Set CreateManualInputBlock = CreateBase(block)
End Function

Public Function CreateDataSimpleBlock(ByVal PrevBlock As BlockBase, ByVal cols As Collection) As BlockBase
    Dim block As New SimpleBlock
    Set block.Super.Symbol = SymbolConfigData.Clone
    block.Super.Initialize cols(2), PrevBlock
    Set CreateDataSimpleBlock = CreateBase(block)
End Function

Public Function CreateDataForkedBlock(ByVal PrevBlock As BlockBase, ByVal cols As Collection) As BlockBase
    Dim block As New ForkedBlock
    Set block.Super.Symbol = SymbolConfigData.Clone
    block.Super.Initialize cols(2), PrevBlock
    Dim tmp As String
    tmp = cols(3)
    block.BranchLabel = Mid(tmp, 2, Len(tmp) - 2)
    Set CreateDataForkedBlock = CreateBase(block)
End Function

Public Function CreateDecisionBlock(ByVal PrevBlock As BlockBase, ByVal cols As Collection) As BlockBase
    Dim block As New ForkedBlock
    Set block.Super.Symbol = SymbolConfigDecision.Clone
    block.Super.Initialize cols(2), PrevBlock
    Dim main As String, branch As String
    If cols.Count = 2 Then
        main = "Yes"
        branch = "No"
    Else
        Dim opt As Variant
        opt = Split(Mid(cols(3), 2, Len(cols(3)) - 2), ",")
        main = opt(0)
        branch = opt(1)
    End If
    block.MainLabel = main
    block.BranchLabel = branch
    Set CreateDecisionBlock = CreateBase(block)
End Function

Public Function CreateLoopBlock(ByVal PrevBlock As BlockBase, ByVal cols As Collection) As BlockBase
    Dim block As New LoopBlock
    If cols.Count = 4 Then
        block.IsPreTest = True
        block.Condition = cols(4)
    Else
        block.IsPreTest = False
        block.Condition = ""
    End If
    block.Super.Initialize cols(2), PrevBlock
    Set CreateLoopBlock = CreateBase(block)
End Function

Private Function CreateBase(ByVal SubClass As IInheritable) As BlockBase
    Set CreateBase = SubClass.Super
    Set CreateBase.This = SubClass
End Function
