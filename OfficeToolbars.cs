/*
 * $HeadURL$
 * $LastChangedBy$
 * $Date$
 * $Revision$
 */

using System;
using System.Collections;
using Microsoft.Office.Core;

namespace BlueprintIT.Office
{
	/// <summary>
	///		Represents a method to receive toolbar button click events.
	/// </summary>
	public delegate void ToolbarButtonClickHandler(ToolbarButton button, OfficeWindow window);
	/// <summary>
	///		Represents a method to receive toolbar combo box change events.
	/// </summary>
	public delegate void ToolbarComboBoxChangeHandler(ToolbarComboBox combo, OfficeWindow window);

	/// <summary>
	///		Represents a set of toolbars as would appear on a single OfficeWindow.
	/// </summary>
	/// <remarks>
	///		This is analogous to the CommandBars class.
	/// </remarks>
	public class Toolbars: IEnumerable
	{
		/// <summary>
		///		Holds the toolbars indexed by the toolbar name.
		/// </summary>
		private IDictionary toolbars;
		/// <summary>
		///		The UI manager that owns this object.
		/// </summary>
		private OfficeUIManager manager;

		/// <summary>
		///		Creates a new set of toolbars.
		/// </summary>
		/// <param name="manager">The UI manager that owns the object.</param>
		internal Toolbars(OfficeUIManager manager)
		{
			this.manager=manager;
			toolbars = new Hashtable();
		}

		/// <summary>
		///		The UI manager.
		/// </summary>
		public OfficeUIManager OfficeUIManager
		{
			get
			{
				return manager;
			}
		}

		/// <summary>
		///		An indexer that retrieves a Toolbar based on its name.
		/// </summary>
		public Toolbar this[string index]
		{
			get
			{
				return (Toolbar)toolbars[index];
			}
		}

		/// <summary>
		///		Creates a new toolbar in the collection.
		/// </summary>
		/// <param name="name">The name of the new toolbar.</param>
		/// <returns>The created toolbar.</returns>
		public Toolbar Add(string name)
		{
			Toolbar bar = new Toolbar(this,name);
			toolbars[name]=bar;
			return bar;
		}

		/// <summary>
		///		Deletes a toolbar from the collection.
		/// </summary>
		/// <param name="bar">The toolbar to be deleted.</param>
		internal void Delete(Toolbar bar)
		{
			toolbars.Remove(bar.Caption);
		}

		/// <summary>
		///		Applies the toolbars to the given window.
		/// </summary>
		/// <param name="window">The window to apply to.</param>
		internal void Apply(OfficeWindow window)
		{
			CommandBars bars = window.CommandBars;
			foreach (Toolbar toolbar in toolbars.Values)
			{
				CommandBar commandbar = null;
				try
				{
					commandbar = bars[toolbar.Caption];
				}
				catch
				{
				}
				if (commandbar==null)
				{
					manager.log("Creating new toolbar");
					commandbar = bars.Add(toolbar.Caption,toolbar.Position,false,true);
				}
				else
				{
					manager.log("Re-mapping old toolbar");
				}
				toolbar.Apply(window,commandbar);
				manager.RegisterCommandBar(commandbar,window,toolbar);
			}
		}

		/// <summary>
		///		Retrieves an Enumerator to the toolbars.
		/// </summary>
		/// <returns>The enumerator.</returns>
		public IEnumerator GetEnumerator()
		{
			return toolbars.Values.GetEnumerator();
		}
	}

	/// <summary>
	///		Represents a single Toolbar for an addin.
	/// </summary>
	/// <remarks>
	///		Analagous to a CommandBar. May map to a CommandBar created dynamically for each
	///		window or a built in CommandBar.
	/// </remarks>
	public class Toolbar: ToolbarPopup
	{
		/// <summary>
		///		The Toolbars that holds this toolbar.
		/// </summary>
		private Toolbars toolbars;
		/// <summary>
		///		The position of the toolbar.
		/// </summary>
		private MsoBarPosition position;

		/// <summary>
		///		Used to ensure that each Toolbar has a unique reference number.
		/// </summary>
		/// <remarks>
		///		Starts at 0 and is incremented with each new toolbar that is created.
		/// </remarks>
		private static int TOOLBAR_TAG = 0;

		/// <summary>
		///		Creates the new toolbar.
		/// </summary>
		/// <param name="bars">The parent container.</param>
		/// <param name="name">The name for the toolbar.</param>
		internal Toolbar(Toolbars bars, string name): base(null)
		{
			this.toolbars=bars;
			base.Caption=name;
			this.position=MsoBarPosition.msoBarFloating;
			InternalTag="UIManager_"+TOOLBAR_TAG;
			TOOLBAR_TAG++;
		}

		/// <summary>
		///		The position of the toolbar. Same semantics as <see cref="CommandBar.Position">CommandBar.Position</see>
		/// </summary>
		public MsoBarPosition Position
		{
			get
			{
				return position;
			}

			set
			{
				position=value;
			}
		}

		/// <summary>
		///		The UI manager that owns this toolbar.
		/// </summary>
		public override OfficeUIManager OfficeUIManager
		{
			get
			{
				return toolbars.OfficeUIManager;
			}
		}

		/// <summary>
		///		The name of the toolbar.
		/// </summary>
		/// <remarks>
		///		Though a set method is defined it currently does nothing.
		/// </remarks>
		public override string Caption
		{
			get
			{
				return base.Caption;
			}
			set
			{
			}
		}

		/// <summary>
		///		Finds the CommandBarControls for a given OfficeWindow that maps to this toolbar.
		/// </summary>
		/// <param name="window">The window to search through.</param>
		/// <returns>The CommandBarControls that contains controls that map to the controls that this toolbar contains.</returns>
		internal override CommandBarControls FindCommandBarControls(OfficeWindow window)
		{
			return GetCommandBar(window).Controls;
		}
		
		/// <summary>
		///		Finds the CommandBar in the windows that was created from this toolbar.
		/// </summary>
		/// <param name="window">The Window to look in.</param>
		/// <returns>The found CommandBar.</returns>
		public CommandBar GetCommandBar(OfficeWindow window)
		{
			CommandBars bars = window.CommandBars;
			try
			{
				CommandBar bar = bars[Caption];
				return bar;
			}
			catch
			{
				return null;
			}
		}

		/// <summary>
		///		Applies this toolbar to the given CommandBar.
		/// </summary>
		/// <param name="window">The window that the CommandBar belongs to.</param>
		/// <param name="commandbar"></param>
		internal void Apply(OfficeWindow window, CommandBar commandbar)
		{
			Apply(window,commandbar.Controls);
			commandbar.Visible=visible.GetValue(window);
			commandbar.Enabled=enabled.GetValue(window);
		}

		/// <summary>
		///		Deletes this toolbar from the UI.
		/// </summary>
		public override void Delete()
		{
			toolbars.Delete(this);
		}
	}

	/// <summary>
	///		Represents a popup menu on a toolbar.
	/// </summary>
	/// <remarks>
	///		Analogous to a CommandBarPopup.
	/// </remarks>
	public class ToolbarPopup: ToolbarControl, IEnumerable
	{
		/// <summary>
		///		Maps a tag to a ToolbarControl.
		/// </summary>
		private IDictionary controlMap;
		/// <summary>
		///		A list of the controls contained in the popup.
		/// </summary>
		private IList controls;

		/// <summary>
		///		Used to give each control in the popup a unique tag.
		/// </summary>
		/// <remarks>
		///		Starts at 0 and is incremented every time a new control is added to this popup instance.
		/// </remarks>
		private int NEXT_TAG = 0;

		/// <summary>
		///		Creates the popup menu.
		/// </summary>
		/// <param name="parent">The parent popup menu.</param>
		internal ToolbarPopup(ToolbarPopup parent): base(parent,MsoControlType.msoControlPopup)
		{
			controlMap = new Hashtable();
			controls = new ArrayList();
		}

		/// <summary>
		///		Returns a ToolbarControl with a given tag.
		/// </summary>
		public ToolbarControl this[string tag]
		{
			get
			{
				return (ToolbarControl)controlMap[tag];
			}
		}

		/// <summary>
		///		Finds the CommandBarControls that matches this ToolbarPopup instance on the given window.
		/// </summary>
		/// <param name="window">The window to look in.</param>
		/// <returns>The CommandBarControls instance.</returns>
		internal virtual CommandBarControls FindCommandBarControls(OfficeWindow window)
		{
			CommandBarPopup proxy = (CommandBarPopup)GetCommandBarControl(window);
			return proxy.Controls;
		}

		/// <summary>
		///		Create a new control of a particular type.
		/// </summary>
		/// <param name="type">The type of control to create.</param>
		/// <returns>One of ToolbarPopup, ToolbarButton and ToolbarComboBox depending on the type given.</returns>
		private ToolbarControl CreateControl(MsoControlType type)
		{
			ToolbarControl control = null;
			switch (type)
			{
				case MsoControlType.msoControlPopup:
					control = new ToolbarPopup(this);
					break;
				case MsoControlType.msoControlButton:
					control = new ToolbarButton(this);
					break;
				case MsoControlType.msoControlEdit:
				case MsoControlType.msoControlDropdown:
				case MsoControlType.msoControlComboBox:
					control = new ToolbarComboBox(this,type);
					break;
			}
			if (control!=null)
			{
				control.InternalTag=InternalTag+"_"+NEXT_TAG;
				NEXT_TAG++;
			}
			return control;
		}

		/// <summary>
		///		Adds a new control to the popup menu.
		/// </summary>
		/// <param name="type">The type of control to add.</param>
		/// <returns>The newly created control.</returns>
		public ToolbarControl Add(MsoControlType type)
		{
			ToolbarControl control = CreateControl(type);
			controls.Add(control);
			controlMap[control.InternalTag]=control;
			return control;
		}

		/// <summary>
		///		Creates a new control in a particular position on the menu.
		/// </summary>
		/// <param name="index">The position to add the control.</param>
		/// <param name="type">The type of control to add.</param>
		/// <returns>The added control.</returns>
		public ToolbarControl Insert(int index, MsoControlType type)
		{
			ToolbarControl control = CreateControl(type);
			controls.Insert(index,control);
			controlMap[control.InternalTag]=control;
			return control;
		}

		/// <summary>
		///		Deletes all the controls on the menu.
		/// </summary>
		public void Clear()
		{
			controls.Clear();
		}

		/// <summary>
		///		The number of controls on the menu.
		/// </summary>
		public int Count
		{
			get
			{
				return controls.Count;
			}
		}

		/// <summary>
		///		Deletes a control from the menu.
		/// </summary>
		/// <param name="control">The control to delete.</param>
		internal void Delete(ToolbarControl control)
		{
			controls.Remove(control);
		}

		/// <summary>
		///		Applies this menu to the given CommandBarControls.
		/// </summary>
		/// <param name="window">The window holding the controls.</param>
		/// <param name="controls">The control set to apply to.</param>
		protected void Apply(OfficeWindow window, CommandBarControls controls)
		{
			IList toadd = new ArrayList(this.controls);
			foreach (CommandBarControl control in controls)
			{
				if (control.Tag.StartsWith("UIManager_"))
				{
					string thetag = control.Tag.Substring(0,control.Tag.IndexOf("#"));
					ToolbarControl tcontrol = (ToolbarControl)controlMap[thetag];
					if (tcontrol!=null)
					{
						OfficeUIManager.log("Re-mapping old control");
						tcontrol.Apply(window,control);
						OfficeUIManager.RegisterCommandBarControl(control,window,tcontrol);
						toadd.Remove(tcontrol);
					}
				}
			}
			foreach (ToolbarControl tcontrol in toadd)
			{
				OfficeUIManager.log("Creating new control");
				CommandBarControl control = controls.Add(tcontrol.Type,1,System.Reflection.Missing.Value,System.Reflection.Missing.Value,true);
				control.Tag = tcontrol.InternalTag+"#"+NEXT_TAG;
				control.OnAction = "!<"+OfficeUIManager.AddinProgID+">";
				NEXT_TAG++;
				tcontrol.Apply(window,control);
				OfficeUIManager.RegisterCommandBarControl(control,window,tcontrol);
			}
		}

		/// <summary>
		///		Applys to the given control.
		/// </summary>
		/// <param name="window">The window holding the control.</param>
		/// <param name="control">The control to apply to.</param>
		internal override void Apply(OfficeWindow window, CommandBarControl control)
		{
			base.Apply(window,control);
			InternalApply(window,control as CommandBarPopup);
		}

		/// <summary>
		///		Applys any particular popup settings to the control.
		/// </summary>
		/// <param name="window">The window holding the control.</param>
		/// <param name="control">The popup control.</param>
		private void InternalApply(OfficeWindow window, CommandBarPopup control)
		{
			Apply(window,control.Controls);
		}

		/// <summary>
		///		Retrieves an IEnumerator to the controls in this popup.
		/// </summary>
		/// <returns></returns>
		public IEnumerator GetEnumerator()
		{
			return controls.GetEnumerator();
		}
	}

	/// <summary>
	///		Represents a toolbar button.
	/// </summary>
	/// <remarks>
	///		Analogous to a CommandBarButton.
	/// </remarks>
	public class ToolbarButton: ToolbarControl
	{
		/// <summary>
		///		Occurs when a CommandBarButton created from this control has been clicked.
		/// </summary>
		public event ToolbarButtonClickHandler Click;

		/// <summary>
		///		The initial state of the toolbar button.
		/// </summary>
		private MsoButtonState state = MsoButtonState.msoButtonUp;
		/// <summary>
		///		The initial style of the toolbar button.
		/// </summary>
		private MsoButtonStyle style = MsoButtonStyle.msoButtonCaption;

		/// <summary>
		///		Creates a new toolbar button.
		/// </summary>
		/// <param name="parent">The popup menu holding the button.</param>
		internal ToolbarButton(ToolbarPopup parent): base(parent,MsoControlType.msoControlButton)
		{
		}

		/// <summary>
		///		Called when a CommandBarButton created from this control is clicked. Fires
		///		the event.
		/// </summary>
		/// <param name="window">The window holding the control that was clicked.</param>
		internal void OnClick(OfficeWindow window)
		{
			if (Click!=null)
			{
				Click(this,window);
			}
		}

		/// <summary>
		///		Applys this controls settings to the proxy.
		/// </summary>
		/// <param name="window">The window containing the control.</param>
		/// <param name="control">The control to apply settings to.</param>
		internal override void Apply(OfficeWindow window, CommandBarControl control)
		{
			base.Apply(window,control);
			InternalApply(window,control as CommandBarButton);
		}

		/// <summary>
		///		Applies specific button settings to the control proxy.
		/// </summary>
		/// <param name="window">The window holding the proxy.</param>
		/// <param name="control">The proxy control.</param>
		private void InternalApply(OfficeWindow window, CommandBarButton control)
		{
			control.State=state;
			control.Style=style;
		}

		/// <summary>
		///		The state of the button.
		/// </summary>
		public MsoButtonState State
		{
			get
			{
				return state;
			}

			set
			{
				state=value;
			}
		}

		/// <summary>
		///		The style of the button.
		/// </summary>
		public MsoButtonStyle Style
		{
			get
			{
				return style;
			}

			set
			{
				style=value;
			}
		}
	}

	/// <summary>
	///		Represents a combo box on the toolbar.
	/// </summary>
	/// <remarks>
	///		Analogous to a CommandBarComboBox.
	/// </remarks>
	public class ToolbarComboBox: ToolbarControl
	{
		/// <summary>
		///		Occurs when a combo box that was created from this control is changed on any window.
		/// </summary>
		public event ToolbarComboBoxChangeHandler Change;

		/// <summary>
		///		The combo box initial style.
		/// </summary>
		private MsoComboStyle style = MsoComboStyle.msoComboNormal;
		/// <summary>
		///		The number of header lines in the combo box.
		/// </summary>
		private int headerCount = -1;
		/// <summary>
		///		The items in the combo box.
		/// </summary>
		private IList items;
		/// <summary>
		///		The number of lines to drop down.
		/// </summary>
		private int dropDownLines = 0;
		/// <summary>
		///		The width of the drop down.
		/// </summary>
		private int dropDownWidth = 0;

		/// <summary>
		///		Creates a new control.
		/// </summary>
		/// <param name="parent">The popup menu holding the control.</param>
		/// <param name="type">The type of control to create.</param>
		internal ToolbarComboBox(ToolbarPopup parent, MsoControlType type): base(parent,type)
		{
			items = new ArrayList();
		}

		/// <summary>
		///		Called when a combo box is changed that will fire the events.
		/// </summary>
		/// <param name="window">The window holding the combo box that was changed.</param>
		internal void OnChange(OfficeWindow window)
		{
			if (Change!=null)
			{
				Change(this,window);
			}
		}

		/// <summary>
		///		Applies settings to the given control.
		/// </summary>
		/// <param name="window">The window holding the control.</param>
		/// <param name="control">The proxy control.</param>
		internal override void Apply(OfficeWindow window, CommandBarControl control)
		{
			base.Apply(window,control);
			InternalApply(window,control as CommandBarComboBox);
		}

		/// <summary>
		///		Applies combo box specific settings to the control.
		/// </summary>
		/// <param name="window">The window holding the control.</param>
		/// <param name="control">The control proxy.</param>
		private void InternalApply(OfficeWindow window, CommandBarComboBox control)
		{
			control.Style=style;
			control.DropDownLines=dropDownLines;
			control.DropDownWidth=dropDownWidth;
			control.ListHeaderCount=headerCount;
			control.Clear();
			foreach (string item in items)
			{
				control.AddItem(item,System.Reflection.Missing.Value);
			}
		}

		/// <summary>
		///		Adds a new item to the drop down list.
		/// </summary>
		/// <param name="item">The item to add.</param>
		public void Add(string item)
		{
			items.Add(item);
		}

		/// <summary>
		///		Adds a new item to the drop down list in a specific position.
		/// </summary>
		/// <param name="item">The item to add.</param>
		/// <param name="index">The position to add the item to.</param>
		public void Add(string item, int index)
		{
			items.Insert(index-1,item);
		}

		/// <summary>
		///		Removes an item from the drop down list.
		/// </summary>
		/// <param name="index">The position of the item to be removed.</param>
		public void Remove(int index)
		{
			items.Remove(index-1);
		}

		/// <summary>
		///		Clears the drop down list.
		/// </summary>
		public void Clear()
		{
			items.Clear();
		}

		/// <summary>
		///		The number of items in the drop down list.
		/// </summary>
		public int Count
		{
			get
			{
				return items.Count;
			}
		}

		/// <summary>
		///		The number of header items in the drop down list.
		/// </summary>
		public int HeaderCount
		{
			get
			{
				return headerCount;
			}

			set
			{
				headerCount=value;
			}
		}

		/// <summary>
		///		 The number of lines to drop down.
		/// </summary>
		public int DropDownLines
		{
			get
			{
				return dropDownLines;
			}

			set
			{
				dropDownLines=value;
			}
		}

		/// <summary>
		///		The width of the drop down list.
		/// </summary>
		public int DropDownWidth
		{
			get
			{
				return dropDownWidth;
			}

			set
			{
				dropDownWidth=value;
			}
		}

		/// <summary>
		///		The style of the combo box.
		/// </summary>
		public MsoComboStyle Style
		{
			get
			{
				return style;
			}

			set
			{
				style=value;
			}
		}
	}

	/// <summary>
	///		The base toolbar control.
	/// </summary>
	/// <remarks>
	///		Analogous to the CommandBarControl.
	/// </remarks>
	public abstract class ToolbarControl
	{
		/// <summary>
		///		The popup holding this control.
		/// </summary>
		private ToolbarPopup parent;
		/// <summary>
		///		The caption of the control.
		/// </summary>
		private string caption;
		/// <summary>
		///		The unique tag for this control. Dynamically generated.
		/// </summary>
		protected string internalTag;
		/// <summary>
		///		The type of the control.
		/// </summary>
		private MsoControlType type;
		/// <summary>
		///		The initial visibility of the control.
		/// </summary>
		protected BooleanValue visible = true;
		/// <summary>
		///		Whether the control is enabled or not.
		/// </summary>
		protected BooleanValue enabled = true;
		/// <summary>
		///		The tooltip for the control.
		/// </summary>
		private string tip = null;
		/// <summary>
		///		The priority of the control.
		/// </summary>
		private int priority = 3;
		/// <summary>
		///		A tag to be linked to the control.
		/// </summary>
		private object tag = null;

		/// <summary>
		///		Creates a new toolbar control.
		/// </summary>
		/// <param name="parent">The popup menu holding the control.</param>
		/// <param name="type">The type of the control.</param>
		protected ToolbarControl(ToolbarPopup parent, MsoControlType type)
		{
			this.parent=parent;
			this.type=type;
		}

		/// <summary>
		///		Returns the proxy object in the given window.
		/// </summary>
		/// <param name="window">The window to search through.</param>
		/// <returns>A CommandBarControl that was created from this control on the given window.</returns>
		public CommandBarControl GetCommandBarControl(OfficeWindow window)
		{
			CommandBarControls controls = Parent.FindCommandBarControls(window);
			foreach (CommandBarControl control in controls)
			{
				if (control.Tag.StartsWith(InternalTag+"#"))
				{
					return control;
				}
			}
			return null;
		}

		/// <summary>
		///		Applies settings to the control proxy.
		/// </summary>
		/// <param name="window">The window holding the proxy.</param>
		/// <param name="control">The control proxy.</param>
		internal virtual void Apply(OfficeWindow window, CommandBarControl control)
		{
			control.Caption=caption;
			control.Visible=visible.GetValue(window);
			control.Enabled=enabled.GetValue(window);
			control.Priority=priority;
			if (tip!=null)
			{
				control.TooltipText = tip;
			}
			else
			{
				control.TooltipText = caption;
			}
		}

		/// <summary>
		///		 The UI manager.
		/// </summary>
		public virtual OfficeUIManager OfficeUIManager
		{
			get
			{
				return parent.OfficeUIManager;
			}
		}

		/// <summary>
		///		An object attachment for the control.
		/// </summary>
		/// <remarks>
		///		Not passed on the the proxy objects in any way.
		/// </remarks>
		public object Tag
		{
			get
			{
				return tag;
			}

			set
			{
				tag=value;
			}
		}

		/// <summary>
		///		The unique identifier for this control.
		/// </summary>
		internal string InternalTag
		{
			get
			{
				return internalTag;
			}

			set
			{
				internalTag=value;
			}
		}

		/// <summary>
		///		The tooltip for this control.
		/// </summary>
		/// <remarks>
		///		A null tooltip uses the caption as the tooltip.
		/// </remarks>
		public string Tooltip
		{
			get
			{
				return tip;
			}

			set
			{
				tip=value;
			}
		}

		/// <summary>
		///		The caption of the control.
		/// </summary>
		public virtual string Caption
		{
			get
			{
				return caption;
			}

			set
			{
				caption=value;
			}
		}

		/// <summary>
		///		The priority of the control
		/// </summary>
		/// <remarks>
		///		Captions with lower priority disappear first when the toolbar doesnt have enough 
		///		space to display all controls. The default priority is 3.
		/// </remarks>
		public int Priority
		{
			get
			{
				return priority;
			}

			set
			{
				priority=value;
			}
		}

		/// <summary>
		///		Whether the control is visible or not.
		/// </summary>
		public BooleanValue Visible
		{
			get
			{
				return visible;
			}

			set
			{
				visible=value;
			}
		}

		/// <summary>
		///		Whether the control is enabled or not.
		/// </summary>
		public BooleanValue Enabled
		{
			get
			{
				return enabled;
			}

			set
			{
				enabled=value;
			}
		}

		/// <summary>
		///		The type of the control.
		/// </summary>
		public MsoControlType Type
		{
			get
			{
				return type;
			}
		}

		/// <summary>
		///		 The popup menu holding this control.
		/// </summary>
		protected ToolbarPopup Parent
		{
			get
			{
				return parent;
			}
		}

		/// <summary>
		///		Deletes the control.
		/// </summary>
		public virtual void Delete()
		{
			parent.Delete(this);
		}
	}
}
