' More info
' https://forums.autodesk.com/t5/inventor-ideas/create-a-class-to-handle-unexpected-dimension-text-placement/idi-p/12396222
' https://forums.autodesk.com/t5/inventor-programming-ilogic/unexpected-dimension-text-place/td-p/12227266
Public Sub Main()
	Dim Sheet_1 = ThisDrawing.Sheets.ItemByName("Sheet:1")
	Dim VIEW1 = Sheet_1.DrawingViews.ItemByName("VIEW1")
	Dim dimensions = Sheet_1.DrawingDimensions.GeneralDimensions

	' Get all intents
	Dim legIntent1 = VIEW1.GetIntent("LegMount1", PointIntentEnum.kMidPointIntent)
	Dim legIntent2 = VIEW1.GetIntent("LegMount2", PointIntentEnum.kMidPointIntent)
	Dim legIntent3 = VIEW1.GetIntent("LegMount3", PointIntentEnum.kMidPointIntent)
	Dim sideIntent1 = VIEW1.GetIntent("side1", Nothing)
	Dim sideIntent2 = VIEW1.GetIntent("side2", Nothing)
	
	' Setup default values
	Dim y1 = -0.3
	Dim horizontalDimType = DimensionTypeEnum.kHorizontalDimensionType
	
	Using manager As New DrawingComponentManager(ThisDrawing)
		' Placement logic
		manager.CenterDimensionText = False
		dimensions.AddLinear("Mounting point1", VIEW1.SheetPoint(-0.07, y1), sideIntent1, legIntent1)
		dimensions.AddLinear("Mounting point2", VIEW1.SheetPoint(1.07, y1), legIntent3, sideIntent2)
		
		manager.CenterDimensionText = True
		dimensions.AddLinear("LegMountDim1", VIEW1.SheetPoint(0.5, y1), legIntent1, legIntent2, horizontalDimType)	
		dimensions.AddLinear("LegMountDim2", VIEW1.SheetPoint(0.5, y1), legIntent2, legIntent3, horizontalDimType)
	End Using
End Sub



Public Class DrawingComponentManager
    Implements IDisposable

    Private _drawing As IManagedDrawing
    Private _inventor As Inventor.InventorServer
    Private _style As DimensionStyle

    Private _center As Boolean
    Private _horizontalOrientation As DimensionTextOrientationEnum
    Private _alignedOrientation As DimensionTextOrientationEnum
    Private _verticalOrientation As DimensionTextOrientationEnum

    Public Sub New(drawing As IManagedDrawing)
        _drawing = drawing
        _inventor = drawing.NativeEntity.Parent
        _style = drawing.Document.StylesManager.ActiveStandardStyle.ActiveObjectDefaults.LinearDimensionStyle

        _center = _inventor.DrawingOptions.CenterDimensionText
        _horizontalOrientation = _style.HorizontalDimensionTextOrientation
        _alignedOrientation = _style.AlignedDimensionTextOrientation
        _verticalOrientation = _style.VerticalDimensionTextOrientation

        _drawing.BeginManage()
        _style.HorizontalDimensionTextOrientation = DimensionTextOrientationEnum.kInlineHorizontalDimensionText
        _style.AlignedDimensionTextOrientation = DimensionTextOrientationEnum.kInlineAlignedDimensionText
        _style.VerticalDimensionTextOrientation = DimensionTextOrientationEnum.kInlineHorizontalDimensionText
    End Sub

    Public WriteOnly Property CenterDimensionText() As Boolean
        Set(ByVal value As Boolean)
            _inventor.DrawingOptions.CenterDimensionText = value
        End Set
    End Property

    Public Sub Dispose() Implements IDisposable.Dispose
        _drawing.EndManage()
        _inventor.DrawingOptions.CenterDimensionText = _center
        _style.HorizontalDimensionTextOrientation = _horizontalOrientation
        _style.AlignedDimensionTextOrientation = _alignedOrientation
        _style.VerticalDimensionTextOrientation = _verticalOrientation
    End Sub
	
	' Copyright 2024
    ' 
    ' This code was written by Jelte de Jong, and published on www.hjalte.nl/https://github.com/hjalte79
    '
    ' Permission Is hereby granted, free of charge, to any person obtaining a copy of this 
    ' software And associated documentation files (the "Software"), to deal in the Software 
    ' without restriction, including without limitation the rights to use, copy, modify, merge, 
    ' publish, distribute, sublicense, And/Or sell copies of the Software, And to permit persons 
    ' to whom the Software Is furnished to do so, subject to the following conditions:
    '
    ' The above copyright notice And this permission notice shall be included In all copies Or
    ' substantial portions Of the Software.
    ' 
    ' THE SOFTWARE Is PROVIDED "AS IS", WITHOUT WARRANTY Of ANY KIND, EXPRESS Or IMPLIED, 
    ' INCLUDING BUT Not LIMITED To THE WARRANTIES Of MERCHANTABILITY, FITNESS For A PARTICULAR 
    ' PURPOSE And NONINFRINGEMENT. In NO Event SHALL THE AUTHORS Or COPYRIGHT HOLDERS BE LIABLE 
    ' For ANY CLAIM, DAMAGES Or OTHER LIABILITY, WHETHER In AN ACTION Of CONTRACT, TORT Or 
    ' OTHERWISE, ARISING FROM, OUT Of Or In CONNECTION With THE SOFTWARE Or THE USE Or OTHER 
    ' DEALINGS In THE SOFTWARE.
End Class