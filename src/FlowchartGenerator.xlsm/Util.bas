Attribute VB_Name = "Util"
Option Explicit

Sub ZDeleteAllShapes()

Dim shp As Shape

For Each shp In ActiveSheet.Shapes
   If shp.Type <> msoFormControl Then
        shp.Delete
   End If
Next shp

End Sub
