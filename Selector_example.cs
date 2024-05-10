public partial class ThisCRule
{
	public void Main()
	{
		PartDocument doc = (PartDocument)ThisDoc.Document;

		Selector<Face> selector = new Selector<Face>(ThisApplication);
		selector.SelectionMethod = IsPlanarFace;
		// Or
		// selector.SelectProperties.AddSelectionFilter(SelectionFilterEnum.kPartFacePlanarFilter)

		selector.InteractProperties.SetCursor(CursorTypeEnum.kCursorBuiltInCrosshair);
		selector.InteractProperties.StatusBarText = "Select a point on a face.";

		// Just to show case that minitoolbars are possible
		var miniToolbar = selector.InteractProperties.CreateMiniToolbar();
		miniToolbar.EnableOK = false;
		miniToolbar.EnableApply = false;
		miniToolbar.ShowCancel = true;
		miniToolbar.ShowOptionBox = false;
		miniToolbar.Visible = true;

		selector.Pick();

		if ((selector.SelectionStatus == SelectionStatusEnum.SometingIsSelected))
		{
			Face face = selector.SelectedEntities.First();
			Inventor.Point pointOnFace = selector.ModelPosition;

			CreateMarkFeatur(doc, face, pointOnFace);
		}
	}

	private bool IsPlanarFace(object entity)
	{
		if (entity is Face)
		{
			if (((Face)entity).Geometry is Plane)
				return true;
		}
		return false;
	}

	private void CreateMarkFeatur(PartDocument doc, Face face, Inventor.Point pointOnFace)
	{
		var markStyleName = "Mark Surface";
		var markText = "TEST";
		var textHeight = 2; // Cm

		if ((string.IsNullOrWhiteSpace(markText)))
		{
			MessageBox.Show("Mark text was not set. Ending this rule now.", "No mark text");
			return;
		}

		PartComponentDefinition def = doc.ComponentDefinition;

		var sketch = def.Sketches.Add(face);
		Point2d modelPoint = sketch.ModelToSketchSpace(pointOnFace);

		var formatedText = string.Format("<StyleOverride FontSize='{0}'>{1}</StyleOverride>", textHeight, markText);
		var sketchText = sketch.TextBoxes.AddFitted(modelPoint, formatedText);

		ObjectCollection markGeometry = ThisApplication.TransientObjects.CreateObjectCollection();
		markGeometry.Add(sketchText);

		MarkFeatures markFeatures = def.Features.MarkFeatures;
		MarkStyle markStyle = doc.MarkStyles[markStyleName];
		MarkDefinition markDef = markFeatures.CreateMarkDefinition(markGeometry, markStyle);
		if ((ThisApplication.SoftwareVersion.Major >= 28))
			// The following property was introduced in Inventor 2024
			// Therefore we can't set this property before Inventor 2024
			// If this is not set after Inventor 2023 the rule fails most of the time.
			// (Major version 28 = Inventor 2024)
			markDef.Direction = PartFeatureExtentDirectionEnum.kNegativeExtentDirection;
		markFeatures.Add(markDef);
	}
}


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
	
	// Copyright 2024
    // 
    // This code was written by Jelte de Jong, and published on www.hjalte.nl/https://github.com/hjalte79
    //
    // Permission Is hereby granted, free of charge, to any person obtaining a copy of this 
    // software And associated documentation files (the "Software"), to deal in the Software 
    // without restriction, including without limitation the rights to use, copy, modify, merge, 
    // publish, distribute, sublicense, And/Or sell copies of the Software, And to permit persons 
    // to whom the Software Is furnished to do so, subject to the following conditions:
    //
    // The above copyright notice And this permission notice shall be included In all copies Or
    // substantial portions Of the Software.
    // 
    // THE SOFTWARE Is PROVIDED "AS IS", WITHOUT WARRANTY Of ANY KIND, EXPRESS Or IMPLIED, 
    // INCLUDING BUT Not LIMITED To THE WARRANTIES Of MERCHANTABILITY, FITNESS For A PARTICULAR 
    // PURPOSE And NONINFRINGEMENT. In NO Event SHALL THE AUTHORS Or COPYRIGHT HOLDERS BE LIABLE 
    // For ANY CLAIM, DAMAGES Or OTHER LIABILITY, WHETHER In AN ACTION Of CONTRACT, TORT Or 
    // OTHERWISE, ARISING FROM, OUT Of Or In CONNECTION With THE SOFTWARE Or THE USE Or OTHER 
    // DEALINGS In THE SOFTWARE.
}

public enum SelectionStatusEnum
{
	NotStarted,
	StillSelecting,
	Canceled,
	SometingIsSelected
}