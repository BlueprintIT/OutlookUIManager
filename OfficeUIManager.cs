using System;
using System.Collections;
using Microsoft.Office.Core;

namespace BlueprintIT.Office
{
	public delegate void WindowEventHandler(OfficeWindow window);

	public abstract class UIManager: IDisposable
	{
		public event WindowEventHandler WindowOpen;
		public event WindowEventHandler WindowClose;

		private IDictionary controlWindowMap;
		private IDictionary controlControlMap;
		private IDictionary controlProxyMap;

		private string progid;

		private _CommandBarComboBoxEvents_ChangeEventHandler comboBoxChange;
		private _CommandBarButtonEvents_ClickEventHandler buttonClick;

		public UIManager(string progid)
		{
			this.progid=progid;
			controlControlMap = new Hashtable();
			controlWindowMap = new Hashtable();
			controlProxyMap = new Hashtable();
			comboBoxChange = new _CommandBarComboBoxEvents_ChangeEventHandler(combo_Change);
			buttonClick = new _CommandBarButtonEvents_ClickEventHandler(button_Click);
		}

		~UIManager()
		{
			Dispose();
		}

		internal abstract void log(string text);
		
		public abstract IList Windows
		{
			get;
		}

		public string AddinProgID
		{
			get
			{
				return progid;
			}
		}

		public virtual void Dispose()
		{
			foreach (CommandBarControl control in controlProxyMap.Values)
			{
				UnbindControl(control);
#if (COMRELEASE)
				System.Runtime.InteropServices.Marshal.ReleaseComObject(control);
#endif
			}
			controlControlMap.Clear();
			controlWindowMap.Clear();
			controlProxyMap.Clear();
			GC.SuppressFinalize(this);
		}

		protected abstract Toolbars GetWindowToolbars(OfficeWindow window);

		internal void RegisterCommandBar(CommandBar bar, OfficeWindow window, Toolbar toolbar)
		{
		}

		internal void RegisterCommandBarControl(CommandBarControl control, OfficeWindow window, ToolbarControl tcontrol)
		{
			if (!controlControlMap.Contains(control.Tag))
			{
				controlControlMap[control.Tag]=tcontrol;
			}
			else
			{
				CommandBarControl oldcontrol = (CommandBarControl)controlProxyMap[control.Tag];
				UnbindControl(oldcontrol);
#if (COMRELEASE)
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oldcontrol);
#endif
			}
			BindControl(control);
			controlProxyMap[control.Tag]=control;
			controlWindowMap[control.Tag]=window;
		}

		public void ApplyInterface()
		{
			foreach (OfficeWindow window in Windows)
			{
				GetWindowToolbars(window).Apply(window);
			}
		}

		protected void OnWindowOpen(OfficeWindow window)
		{
			Toolbars bars = GetWindowToolbars(window);
			bars.Apply(window);
			if (WindowOpen!=null)
			{
				WindowOpen(window);
			}
		}

		protected void OnWindowClose(OfficeWindow window)
		{
			if (WindowClose!=null)
			{
				WindowClose(window);
			}
			foreach (string tag in controlControlMap.Keys)
			{
				if (controlControlMap[tag]==window)
				{
					controlControlMap.Remove(tag);
					controlWindowMap.Remove(tag);
					CommandBarControl control = (CommandBarControl)controlProxyMap[tag];
					UnbindControl(control);
#if (COMRELEASE)
					System.Runtime.InteropServices.Marshal.ReleaseComObject(control);
#endif
					controlProxyMap.Remove(tag);
				}
			}
		}

		private void BindControl(CommandBarControl control)
		{
			if (control is CommandBarButton)
			{
				CommandBarButton button = control as CommandBarButton;
				button.Click+=buttonClick;
			}
			else if (control is CommandBarPopup)
			{
				CommandBarPopup popup = control as CommandBarPopup;
			}
			else if (control is CommandBarComboBox)
			{
				CommandBarComboBox combo = control as CommandBarComboBox;
				combo.Change+=comboBoxChange;
			}
		}

		private void UnbindControl(CommandBarControl control)
		{
			if (control is CommandBarButton)
			{
				((CommandBarButton)control).Click-=buttonClick;
			}
			else if (control is CommandBarPopup)
			{
			}
			else if (control is CommandBarComboBox)
			{
				((CommandBarComboBox)control).Change-=comboBoxChange;
			}
		}

		private void button_Click(CommandBarButton Ctrl, ref bool CancelDefault)
		{
			ToolbarButton control = (ToolbarButton)controlControlMap[Ctrl.Tag];
			OfficeWindow window = (OfficeWindow)controlWindowMap[Ctrl.Tag];
			control.OnClick(window);
		}

		private void combo_Change(CommandBarComboBox Ctrl)
		{
			ToolbarComboBox control = (ToolbarComboBox)controlControlMap[Ctrl.Tag];
			OfficeWindow window = (OfficeWindow)controlWindowMap[Ctrl.Tag];
			control.OnChange(window);
		}
	}

	public abstract class OfficeWindow
	{
		public abstract CommandBars CommandBars
		{
			get;
		}

		public abstract void Activate();

		public abstract void Close();
 
		public abstract int Left
		{
			get;
			set;
		}
 
		public abstract int Top
		{
			get;
			set;
		}
 
		public abstract int Width
		{
			get;
			set;
		}
 
		public abstract int Height
		{
			get;
			set;
		}
	}
}
