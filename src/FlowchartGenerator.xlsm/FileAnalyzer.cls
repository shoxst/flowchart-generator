VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileAnalyzer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Class FileAnalyzer

Private Lines As New Collection
Private Blocks As New BlockList
Private SymbolStack As New BlockList
Private TargetStack As New Collection
Private IsMain As Boolean
Private Factory As New BlockFactory

Public Sub Read(ByVal FilePath As String)
    Dim line As String
    Open FilePath For Input As #1
        Do Until EOF(1)
            Line Input #1, line
            Lines.Add line
        Loop
    Close #1
End Sub

Public Function Analyze() As BlockList
    Dim dummyBase As BlockBase, endBase As BlockBase, InnerText As String

    Set dummyBase = Factory.CreateStartDummyBlock(100, 50)
    
    AnalyzeLine 1, Blocks, dummyBase
    
    InnerText = IIf(IsMain, "終了", "出口")
    Set endBase = Factory.CreateTerminalBlock(Blocks.Last, InnerText, False)
    Blocks.Add endBase
    
    Set Analyze = Blocks
End Function

Private Sub AnalyzeLine(ByVal index As Integer, ByRef TargetBlocks As BlockList, ByVal PrevBlock As BlockBase)
    
    If index > Lines.Count Then
        Exit Sub
    End If
    
    Dim cols As Collection
    Set cols = Parse(Lines(index))
    If cols.Count = 0 Then
        Exit Sub
    End If
    
    Dim base As BlockBase
    Dim fBlock As ForkedBlock, lBlock As LoopBlock
    Dim Target As BlockList, Self As BlockBase
    
    Select Case cols(1)
    Case "#"
        Dim InnerText As String
        If cols(2) = "Main" Then
            IsMain = True
            InnerText = "開始"
        Else
            IsMain = False
            InnerText = cols(2)
        End If
        Set base = Factory.CreateTerminalBlock(PrevBlock, InnerText, True)
        TargetBlocks.Add base
        AnalyzeLine index + 1, TargetBlocks, base
    
    Case "call"
        Set base = Factory.CreatePredefinedBlock(PrevBlock, cols)
        TargetBlocks.Add base
        AnalyzeLine index + 1, TargetBlocks, base
    
    Case "print"
        Set base = Factory.CreateDocumentBlock(PrevBlock, cols)
        TargetBlocks.Add base
        AnalyzeLine index + 1, TargetBlocks, base
    
    Case "display"
        Set base = Factory.CreateDisplayBlock(PrevBlock, cols)
        TargetBlocks.Add base
        AnalyzeLine index + 1, TargetBlocks, base
    
    Case "input"
        Set base = Factory.CreateManualInputBlock(PrevBlock, cols)
        TargetBlocks.Add base
        AnalyzeLine index + 1, TargetBlocks, base
    
    Case "write"
        Set base = Factory.CreateDataSimpleBlock(PrevBlock, cols)
        TargetBlocks.Add base
        AnalyzeLine index + 1, TargetBlocks, base
    
    Case "if"
        Set base = Factory.CreateDecisionBlock(PrevBlock, cols)
        TargetBlocks.Add base
        SymbolStack.Add base
        TargetStack.Add TargetBlocks
        Set fBlock = base.This
        Set cols = Parse(Lines(index + 1))
        If cols(1) = "continue" Then
            AnalyzeLine index + 2, Nothing, Nothing
        Else
            AnalyzeLine index + 1, fBlock.MainBlocks, base
        End If
    
    Case "else"
        Set fBlock = SymbolStack.Last.This
        AnalyzeLine index + 1, fBlock.BranchBlocks, Nothing
    
    Case "end-if"
        Set Self = SymbolStack.Last
        Set Target = TargetStack(TargetStack.Count)
        SymbolStack.Remove SymbolStack.Count
        TargetStack.Remove TargetStack.Count
        AnalyzeLine index + 1, Target, Self
    
    Case "do"
        Set base = Factory.CreateLoopBlock(PrevBlock, cols)
        TargetBlocks.Add base
        SymbolStack.Add base
        TargetStack.Add TargetBlocks
        Set lBlock = base.This
        AnalyzeLine index + 1, lBlock.InnerBlocks, base
    
    Case "loop"
        Set Self = SymbolStack.Last
        Set Target = TargetStack(TargetStack.Count)
        If cols.Count = 3 Then
            Set lBlock = Self.This
            lBlock.Condition = cols(3)
        End If
        SymbolStack.Remove SymbolStack.Count
        TargetStack.Remove TargetStack.Count
        AnalyzeLine index + 1, Target, Self
    
    Case "read"
        If cols.Count = 2 Then
            Set base = Factory.CreateDataSimpleBlock(PrevBlock, cols)
            TargetBlocks.Add base
            AnalyzeLine index + 1, TargetBlocks, base
        Else
            Set base = Factory.CreateDataForkedBlock(PrevBlock, cols)
            TargetBlocks.Add base
            SymbolStack.Add base
            TargetStack.Add TargetBlocks
            Set fBlock = base.This
            AnalyzeLine index + 1, fBlock.BranchBlocks, Nothing
        End If
    
    Case "end-read"
        Set Self = SymbolStack.Last
        Set Target = TargetStack(TargetStack.Count)
        SymbolStack.Remove SymbolStack.Count
        TargetStack.Remove TargetStack.Count
        AnalyzeLine index + 1, Target, Self
    
    Case Else
        Set base = Factory.CreateProcessBlock(PrevBlock, cols)
        TargetBlocks.Add base
        AnalyzeLine index + 1, TargetBlocks, base
        
    End Select
    
End Sub

Private Function Parse(ByVal line As String) As Collection
    Dim arr As Variant, tmp As String, str As Variant
    Dim ret As New Collection
    arr = Split(Trim(line), " ")
    tmp = ""
    
    For Each str In arr
        Dim s As String, e As String
        s = Left(str, 1)
        e = Right(str, 1)
        If s = """" Then
            If e = """" Then
                tmp = Mid(str, 2, Len(str) - 2)
                ret.Add tmp
                tmp = ""
            Else
                tmp = Right(str, Len(str) - 1)
            End If
        ElseIf e = """" Then
            tmp = tmp & " " & Left(str, Len(str) - 1)
            ret.Add tmp
            tmp = ""
        Else
            If tmp = "" Then
                ret.Add str
            Else
                tmp = tmp & " " & str
            End If
        End If
    Next
    
    Set Parse = ret
End Function
