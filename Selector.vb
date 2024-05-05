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
End Class
Public Enum SelectionStatusEnum : NotStarted : StillSelecting : Canceled : SometingIsSelected : End Enum