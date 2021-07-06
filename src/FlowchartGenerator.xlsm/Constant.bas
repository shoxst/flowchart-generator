Attribute VB_Name = "Constant"
Option Explicit

Public Const SYMBOL_DEFAULT_WIDTH As Single = 100
Public Const SYMBOL_DEFAULT_HEIGHT As Single = 40

Public Const TEXTBOX_DEFAULT_WIDTH As Single = 30
Public Const TEXTBOX_DEFAULT_HEIGHT As Single = 20

Public Const BLOCK_HORIZONTAL_MARGIN As Single = 30
Public Const BLOCK_VERTICAL_MARGIN As Single = 15

Public Enum SymbolType
    symbolTypeTerminal = msoShapeFlowchartTerminator
    symbolTypeProcess = msoShapeFlowchartProcess
    symbolTypePredefined = msoShapeFlowchartPredefinedProcess
    symbolTypeDecision = msoShapeFlowchartDecision
    symbolTypeLoop = msoShapeSnip2SameRectangle
    symbolTypeDocument = msoShapeFlowchartDocument
    symbolTypeDisplay = msoShapeFlowchartDisplay
End Enum

