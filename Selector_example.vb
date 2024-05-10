Public Class ThisRule

    Sub Main()

        Dim selector As New Selector(Of Face)(ThisApplication)
        selector.SelectionMethod = AddressOf IsPlanarFace
        ' Or
        ' selector.SelectProperties.AddSelectionFilter(SelectionFilterEnum.kPartFacePlanarFilter)

        selector.InteractProperties.SetCursor(CursorTypeEnum.kCursorBuiltInCrosshair)
        selector.InteractProperties.StatusBarText = "Select a point on a face."

        selector.Pick()

        If (selector.SelectionStatus = SelectionStatusEnum.SometingIsSelected) Then
            Dim face As Face = selector.SelectedEntities(0)
            Dim pointOnFace As Point = selector.ModelPosition

            CreateMarkFeatur(face, pointOnFace)
        End If

    End Sub

    Private Function IsPlanarFace(entity As Object) As Boolean
        If TypeOf entity Is Face Then
            If TypeOf entity.Geometry Is Plane Then
                Return True
            End If
        End If
        Return False
    End Function

    Private Sub CreateMarkFeatur(face As Face, pointOnFace As Point)
        Dim markStyleName = "Mark Surface"
        Dim markText = iProperties.Value("Project", "Part Number")
        Dim textHeight = 2 'Cm

        If (String.IsNullOrWhiteSpace(markText)) Then
            MsgBox("Mark text was not set. Ending this rule now.",, "No mark text")
            Return
        End If

        Dim doc As PartDocument = ThisDoc.Document
        Dim def As PartComponentDefinition = doc.ComponentDefinition

        Dim sketch = def.Sketches.Add(face)
        Dim modelPoint As Point2d = sketch.ModelToSketchSpace(pointOnFace)

        Dim formatedText = String.Format("<StyleOverride FontSize='{0}'>{1}</StyleOverride>", textHeight, markText)
        Dim sketchText = sketch.TextBoxes.AddFitted(modelPoint, formatedText)

        Dim markGeometry As ObjectCollection = ThisApplication.TransientObjects.CreateObjectCollection()
        markGeometry.Add(sketchText)

        Dim markFeatures As MarkFeatures = def.Features.MarkFeatures
        Dim markStyle As MarkStyle = doc.MarkStyles.Item(markStyleName)
        Dim markDef As MarkDefinition = markFeatures.CreateMarkDefinition(markGeometry, markStyle)
        If (ThisApplication.SoftwareVersion.Major >= 28) Then
            ' The following property was introduced in Inventor 2024
            ' Therefore we can't set this property before Inventor 2024
            ' If this is not set after Inventor 2023 the rule fails most of the time.
            ' (Major version 28 = Inventor 2024)
            markDef.Direction = PartFeatureExtentDirectionEnum.kNegativeExtentDirection
        End If
        markFeatures.Add(markDef)
    End Sub
End Class

Public Class Selector(Of T)

    Private _inventor As Inventor.Application

    Public Sub New(ThisApplication As Inventor.Application)
        _inventor = ThisApplication

        InteractProperties = _inventor.CommandManager.CreateInteractionEvents
        InteractProperties.InteractionDisabled = False
        AddHandler InteractProperties.OnTerminate, AddressOf oInteractEvents_OnTerminate

        SelectProperties = InteractProperties.SelectEvents
        SelectProperties.WindowSelectEnabled = False
        AddHandler SelectProperties.OnSelect, AddressOf oSelectEvents_OnSelect
        AddHandler SelectProperties.OnPreSelect, AddressOf oSelectEvents_OnPreSelect

        SelectionMethod = AddressOf DefaultSelector
    End Sub

    Public Sub Pick()
        SelectionStatus = SelectionStatusEnum.StillSelecting
        InteractProperties.Start()
        Do While SelectionStatus = SelectionStatusEnum.StillSelecting
            _inventor.UserInterfaceManager.DoEvents()
        Loop
        InteractProperties.Stop()
        _inventor.CommandManager.StopActiveCommand()
    End Sub

    Public Property InteractProperties As InteractionEvents
    Public Property SelectProperties As SelectEvents

    Public Property SelectedEntities As IEnumerable(Of T) = Nothing
    Public Property ModelPosition As Point = Nothing
    Public Property ViewPosition As Point2d = Nothing

    Public Property SelectionMethod As Func(Of Object, Boolean)
    Public Property SelectionStatus As SelectionStatusEnum = SelectionStatusEnum.NotStarted

    Private Function DefaultSelector(PreSelectEntity As Object) As Boolean
        Return (TypeOf PreSelectEntity Is T)
    End Function

    Private Sub oSelectEvents_OnPreSelect(
            ByRef PreSelectEntity As Object,
            ByRef DoHighlight As Boolean,
            ByRef MorePreSelectEntities As ObjectCollection,
            SelectionDevice As SelectionDeviceEnum,
            ModelPosition As Point,
            ViewPosition As Point2d,
            View As Inventor.View)

        DoHighlight = SelectionMethod(PreSelectEntity)
    End Sub

    Private Sub oInteractEvents_OnTerminate()
        If (SelectionStatus <> SelectionStatusEnum.SometingIsSelected) Then
            SelectionStatus = SelectionStatusEnum.Canceled
        End If
    End Sub

    Private Sub oSelectEvents_OnSelect(
            ByVal JustSelectedEntities As ObjectsEnumerator,
            ByVal SelectionDevice As SelectionDeviceEnum,
            ByVal ModelPosition As Point,
            ByVal ViewPosition As Point2d,
            ByVal View As Inventor.View)

        Dim selectedObjects As New List(Of T)
        For Each item As Object In JustSelectedEntities
            If (SelectionMethod(item)) Then
                selectedObjects.Add(item)
                Me.ModelPosition = ModelPosition
            End If
        Next

        Me.SelectedEntities = selectedObjects
        Me.ModelPosition = ModelPosition
        Me.ViewPosition = ViewPosition

        SelectionStatus = SelectionStatusEnum.SometingIsSelected
    End Sub
	
    ' Copyright 2024
    ' 
    ' This code was written by Jelte de Jong, and published on www.hjalte.nl
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
Public Enum SelectionStatusEnum : NotStarted : StillSelecting : Canceled : SometingIsSelected : End Enum