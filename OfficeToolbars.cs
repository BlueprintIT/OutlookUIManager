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
	public delegate void ToolbarButtonClickHandler(ToolbarButton button, OfficeWindow window);
	public delegate void ToolbarComboBoxChangeHandler(ToolbarComboBox combo, OfficeWindow window);

	public class Toolbars: IEnumerable
	{
		private IDictionary toolbars;
		private UIManager manager;

		internal Toolbars(UIManager manager)
		{
			this.manager=manager;
			toolbars = new Hashtable();
		}

		public UIManager UIManager
		{
			get
			{
				return manager;
			}
		}

		public Toolbar this[string index]
		{
			get
			{
				return (Toolbar)toolbars[index];
			}
		}

		public Toolbar Add(string name)
		{
			Toolbar bar = new Toolbar(this,name);
			toolbars[name]=bar;
			return bar;
		}

		internal void Delete(Toolbar bar)
		{
			toolbars.Remove(bar.Caption);
		}

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

		public IEnumerator GetEnumerator()
		{
			return toolbars.Values.GetEnumerator();
		}
	}

	public class Toolbar: ToolbarPopup
	{
		private Toolbars toolbars;
		private MsoBarPosition position;

		private static int TOOLBAR_TAG = 0;

		internal Toolbar(Toolbars bars, string name): base(null)
		{
			this.toolbars=bars;
			base.Caption=name;
			this.position=MsoBarPosition.msoBarFloating;
			InternalTag="UIManager_"+TOOLBAR_TAG;
			TOOLBAR_TAG++;
		}

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

		public override UIManager UIManager
		{
			get
			{
				return toolbars.UIManager;
			}
		}

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

		internal override CommandBarControls FindCommandBarControls(OfficeWindow window)
		{
			return GetCommandBar(window).Controls;
		}
		
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

		internal void Apply(OfficeWindow window, CommandBar commandbar)
		{
			Apply(window,commandbar.Controls);
			commandbar.Visible=visible.GetValue(window);
			commandbar.Enabled=enabled.GetValue(window);
		}

		public override void Delete()
		{
			toolbars.Delete(this);
		}
	}

	public class ToolbarPopup: ToolbarControl, IEnumerable
	{
		private IDictionary controlMap;
		private IList controls;

		private int NEXT_TAG = 0;

		internal ToolbarPopup(ToolbarPopup parent): base(parent,MsoControlType.msoControlPopup)
		{
			controlMap = new Hashtable();
			controls = new ArrayList();
		}

		public ToolbarControl this[string tag]
		{
			get
			{
				return (ToolbarControl)controlMap[tag];
			}
		}

		internal virtual CommandBarControls FindCommandBarControls(OfficeWindow window)
		{
			CommandBarPopup proxy = (CommandBarPopup)GetCommandBarControl(window);
			return proxy.Controls;
		}

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

		public ToolbarControl Add(MsoControlType type)
		{
			ToolbarControl control = CreateControl(type);
			controls.Add(control);
			controlMap[control.InternalTag]=control;
			return control;
		}

		public ToolbarControl Insert(int index, MsoControlType type)
		{
			ToolbarControl control = CreateControl(type);
			controls.Insert(index,control);
			controlMap[control.InternalTag]=control;
			return control;
		}

		public void Clear()
		{
			controls.Clear();
		}

		public int Count
		{
			get
			{
				return controls.Count;
			}
		}

		internal void Delete(ToolbarControl control)
		{
			controls.Remove(control);
		}

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
						UIManager.log("Re-mapping old control");
						tcontrol.Apply(window,control);
						UIManager.RegisterCommandBarControl(control,window,tcontrol);
						toadd.Remove(tcontrol);
					}
				}
			}
			foreach (ToolbarControl tcontrol in toadd)
			{
				UIManager.log("Creating new control");
				CommandBarControl control = controls.Add(tcontrol.Type,1,System.Reflection.Missing.Value,System.Reflection.Missing.Value,true);
				control.Tag = tcontrol.InternalTag+"#"+NEXT_TAG;
				control.OnAction = "!<"+UIManager.AddinProgID+">";
				NEXT_TAG++;
				tcontrol.Apply(window,control);
				UIManager.RegisterCommandBarControl(control,window,tcontrol);
			}
		}

		internal override void Apply(OfficeWindow window, CommandBarControl control)
		{
			base.Apply(window,control);
			InternalApply(window,control as CommandBarPopup);
		}

		private void InternalApply(OfficeWindow window, CommandBarPopup control)
		{
			Apply(window,control.Controls);
		}

		public IEnumerator GetEnumerator()
		{
			return controls.GetEnumerator();
		}
	}

	public class ToolbarButton: ToolbarControl
	{
		public event ToolbarButtonClickHandler Click;

		private MsoButtonState state = MsoButtonState.msoButtonUp;
		private MsoButtonStyle style = MsoButtonStyle.msoButtonCaption;

		internal ToolbarButton(ToolbarPopup parent): base(parent,MsoControlType.msoControlButton)
		{
		}

		internal void OnClick(OfficeWindow window)
		{
			if (Click!=null)
			{
				Click(this,window);
			}
		}

		internal override void Apply(OfficeWindow window, CommandBarControl control)
		{
			base.Apply(window,control);
			InternalApply(window,control as CommandBarButton);
		}

		private void InternalApply(OfficeWindow window, CommandBarButton control)
		{
			control.State=state;
			control.Style=style;
		}

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

	public class ToolbarComboBox: ToolbarControl
	{
		public event ToolbarComboBoxChangeHandler Change;

		private MsoComboStyle style = MsoComboStyle.msoComboNormal;
		private int headerCount = -1;
		private IList items;
		private int dropDownLines = 0;
		private int dropDownWidth = 0;

		internal ToolbarComboBox(ToolbarPopup parent, MsoControlType type): base(parent,type)
		{
			items = new ArrayList();
		}

		internal void OnChange(OfficeWindow window)
		{
			if (Change!=null)
			{
				Change(this,window);
			}
		}

		internal override void Apply(OfficeWindow window, CommandBarControl control)
		{
			base.Apply(window,control);
			InternalApply(window,control as CommandBarComboBox);
		}

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

		public void Add(string item)
		{
			items.Add(item);
		}

		public void Add(string item, int index)
		{
			items.Insert(index-1,item);
		}

		public void Remove(int index)
		{
			items.Remove(index-1);
		}

		public void Clear()
		{
			items.Clear();
		}

		public int Count
		{
			get
			{
				return items.Count;
			}
		}

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

	public abstract class ToolbarControl
	{
		private ToolbarPopup parent;
		private string caption;
		protected string internalTag;
		private MsoControlType type;
		protected BooleanValue visible = true;
		protected BooleanValue enabled = true;
		private string tip = null;
		private int priority = 3;
		private object tag = null;

		protected ToolbarControl(ToolbarPopup parent, MsoControlType type)
		{
			this.parent=parent;
			this.type=type;
		}

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

		public virtual UIManager UIManager
		{
			get
			{
				return parent.UIManager;
			}
		}

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

		public MsoControlType Type
		{
			get
			{
				return type;
			}
		}

		protected ToolbarPopup Parent
		{
			get
			{
				return parent;
			}
		}

		public virtual void Delete()
		{
			parent.Delete(this);
		}
	}
}
