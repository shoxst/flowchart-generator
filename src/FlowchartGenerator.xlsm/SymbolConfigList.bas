Attribute VB_Name = "SymbolConfigList"
Option Explicit

Public SymbolConfigTerminal As New SymbolConfig
Public SymbolConfigProcess As New SymbolConfig
Public SymbolConfigPredifined As New SymbolConfig
Public SymbolConfigDecision As New SymbolConfig
Public SymbolConfigDocument As New SymbolConfig
Public SymbolConfigDisplay As New SymbolConfig
Public SymbolConfigManualInput As New SymbolConfig
Public SymbolConfigData As New SymbolConfig
Public SymbolConfigLoop As New SymbolConfig

Public Sub SetSymbolConfigList()
    ' Terminal
    SymbolConfigTerminal.ShapeType = msoShapeFlowchartTerminator
    
    ' Process
    SymbolConfigProcess.ShapeType = msoShapeFlowchartProcess
    
    ' Predifined
    SymbolConfigPredifined.ShapeType = msoShapeFlowchartPredefinedProcess
    
    ' Decision
    SymbolConfigDecision.ShapeType = msoShapeFlowchartDecision
    
    ' Document
    SymbolConfigDocument.ShapeType = msoShapeFlowchartDocument
    
    ' Display
    SymbolConfigDisplay.ShapeType = msoShapeFlowchartDisplay
    
    ' ManualInput
    SymbolConfigManualInput.ShapeType = msoShapeFlowchartManualInput
    
    ' Data
    SymbolConfigData.ShapeType = msoShapeFlowchartData
    SymbolConfigData.Top = 2
    SymbolConfigData.Left = 3
    SymbolConfigData.Bottom = 5
    SymbolConfigData.Right = 6
    
    ' Loop
    SymbolConfigLoop.ShapeType = msoShapeSnip2SameRectangle
    SymbolConfigLoop.Top = 4
    SymbolConfigLoop.Left = 3
    SymbolConfigLoop.Bottom = 2
    SymbolConfigLoop.Right = 1
End Sub

