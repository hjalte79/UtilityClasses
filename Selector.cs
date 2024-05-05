public class Selector<T>
{
	private Inventor.Application _inventor;

	public Selector(Inventor.Application ThisApplication)
	{
		_inventor = ThisApplication;

		InteractProperties = _inventor.CommandManager.CreateInteractionEvents();
		InteractProperties.InteractionDisabled = false;
		InteractProperties.OnTerminate += oInteractEvents_OnTerminate;

		SelectProperties = InteractProperties.SelectEvents;
		SelectProperties.WindowSelectEnabled = false;
		SelectProperties.OnSelect += oSelectEvents_OnSelect;
		SelectProperties.OnPreSelect += oSelectEvents_OnPreSelect;

		SelectionMethod = DefaultSelector;
	}

	public void Pick()
	{
		SelectionStatus = SelectionStatusEnum.StillSelecting;
		InteractProperties.Start();
		while (SelectionStatus == SelectionStatusEnum.StillSelecting)
			_inventor.UserInterfaceManager.DoEvents();
		InteractProperties.Stop();
		_inventor.CommandManager.StopActiveCommand();
	}

	public InteractionEvents InteractProperties { get; set; }
	public SelectEvents SelectProperties { get; set; }
	public IEnumerable<T> SelectedEntities { get; set; } = null;
	public Inventor.Point ModelPosition { get; set; } = null/* TODO Change to default(_) if this is not a reference type */;
	public Point2d ViewPosition { get; set; } = null/* TODO Change to default(_) if this is not a reference type */;
	public Func<object, bool> SelectionMethod { get; set; }
	public SelectionStatusEnum SelectionStatus { get; set; } = SelectionStatusEnum.NotStarted;

	private bool DefaultSelector(object PreSelectEntity)
	{
		return (PreSelectEntity is T);
	}

	private void oSelectEvents_OnPreSelect(ref object PreSelectEntity, out bool DoHighlight, ref ObjectCollection MorePreSelectEntities, SelectionDeviceEnum SelectionDevice, Inventor.Point ModelPosition, Point2d ViewPosition, Inventor.View View)
	{
		DoHighlight = SelectionMethod(PreSelectEntity);
	}

	private void oInteractEvents_OnTerminate()
	{
		if ((SelectionStatus != SelectionStatusEnum.SometingIsSelected))
			SelectionStatus = SelectionStatusEnum.Canceled;
	}

	private void oSelectEvents_OnSelect(ObjectsEnumerator JustSelectedEntities, SelectionDeviceEnum SelectionDevice, Inventor.Point ModelPosition, Point2d ViewPosition, Inventor.View View)
	{
		List<T> selectedObjects = new List<T>();
		foreach (object item in JustSelectedEntities)
		{
			if ((SelectionMethod(item)))
			{
				selectedObjects.Add((T)item);
				this.ModelPosition = ModelPosition;
			}
		}

		this.SelectedEntities = selectedObjects;
		this.ModelPosition = ModelPosition;
		this.ViewPosition = ViewPosition;

		SelectionStatus = SelectionStatusEnum.SometingIsSelected;
	}
}

public enum SelectionStatusEnum
{
	NotStarted,
	StillSelecting,
	Canceled,
	SometingIsSelected
}