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
	///		Represents a method to receive events about Office based windows
	/// </summary>
	public delegate void WindowEventHandler(OfficeWindow window);

	/// <summary>
	///		Manages the user interface for an Office based addin.
	/// </summary>
	/// <remarks>
	///		The manager allows addins to specify their user interface requirements 
	///		through the use of the toolbar objects. These are then translated into CommandBars when
	///		office windows are opened.
	/// </remarks>
	public abstract class OfficeUIManager: IDisposable
	{
		/// <summary>
		///		Occurs when an Office window is opened.
		/// </summary>
		public event WindowEventHandler WindowOpen;
		/// <summary>
		///		Occurs when an Office window is closed.
		/// </summary>
		public event WindowEventHandler WindowClose;

		/// <summary>
		///		Maps from the command control tag to the <see cref="OfficeWindow">OfficeWindow</see> instance.
		/// </summary>
		private IDictionary controlWindowMap;
		/// <summary>
		///		Maps from the command control tag to the ToolbarControl.
		/// </summary>
		private IDictionary controlControlMap;
		/// <summary>
		///		Maps from the command control tag to the CommandBarControl instance.
		/// </summary>
		private IDictionary controlProxyMap;

		/// <summary>
		///		Holds the COM ProgID of the addin using the UI Manager.
		/// </summary>
		private string progid;

		/// <summary>
		///		Holds the event handler for combo box events.
		/// </summary>
		private _CommandBarComboBoxEvents_ChangeEventHandler comboBoxChange;
		/// <summary>
		///		Holds the event handler for button events.
		/// </summary>
		private _CommandBarButtonEvents_ClickEventHandler buttonClick;

		/// <summary>
		///		Creates a new UI manager for the addin with the given COM ProgID.
		/// </summary>
		/// <param name="progid">The COM ProgID of the addin usin the UI manager</param>
		protected OfficeUIManager(string progid)
		{
			this.progid=progid;
			controlControlMap = new Hashtable();
			controlWindowMap = new Hashtable();
			controlProxyMap = new Hashtable();
			comboBoxChange = new _CommandBarComboBoxEvents_ChangeEventHandler(combo_Change);
			buttonClick = new _CommandBarButtonEvents_ClickEventHandler(button_Click);
		}

		/// <summary>
		///		Disposes the manager if it has not already happened.
		/// </summary>
		~OfficeUIManager()
		{
			Dispose();
		}
		
		/// <summary>
		///		Returns a list of <see cref="OfficeWindow">OfficeWindows</see> currently open
		/// </summary>
		public abstract IList Windows
		{
			get;
		}

		/// <summary>
		///		Returns the addins COM ProgID
		/// </summary>
		public string AddinProgID
		{
			get
			{
				return progid;
			}
		}

		/// <summary>
		///		Disposes of the manager cleanly, releasing all COM objects.
		/// </summary>
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

		/// <summary>
		///		This method should return the Toolbars that need to be applied to a given
		///		OfficeWindow.
		/// </summary>
		/// <remarks>
		///		Managers for particular office applications should define this method.
		/// </remarks>
		/// <param name="window">A window that has just opened.</param>
		/// <returns>The Toolbars to apply to the window.</returns>
		protected abstract Toolbars GetWindowToolbars(OfficeWindow window);

		/// <summary>
		///		Registers a newly created command bar with the UI manager.
		/// </summary>
		/// <param name="bar">The CommandBar created.</param>
		/// <param name="window">The OfficeWindow the command bar was created on.</param>
		/// <param name="toolbar">The Toolbar that the command bar was created from.</param>
		internal void RegisterCommandBar(CommandBar bar, OfficeWindow window, Toolbar toolbar)
		{
		}

		/// <summary>
		///		Registers a new command bar control.
		/// </summary>
		/// <remarks>
		///		The UI manager will begin listening out for events happening to this control
		///		and dispatch them to the correct listeners.
		/// </remarks>
		/// <param name="control">The CommandBarControl created.</param>
		/// <param name="window">The OfficeWindow holding the control.</param>
		/// <param name="tcontrol">The ToolbarControl that the control was created from.</param>
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

		/// <summary>
		///		Applies the defined interface to all currently open windows.
		/// </summary>
		public void ApplyInterface()
		{
			foreach (OfficeWindow window in Windows)
			{
				GetWindowToolbars(window).Apply(window);
			}
		}

		/// <summary>
		///		Occurs when a new window has been opened.
		/// </summary>
		/// <remarks>
		///		Applies the necessary user interface to the window.
		/// </remarks>
		/// <param name="window">The window that has just opened.</param>
		protected void OnWindowOpen(OfficeWindow window)
		{
			Toolbars bars = GetWindowToolbars(window);
			bars.Apply(window);
			if (WindowOpen!=null)
			{
				WindowOpen(window);
			}
		}

		/// <summary>
		///		Occurs when a window has been closed.
		/// </summary>
		/// <remarks>
		///		Unregisters and toolbar buttons on the window and releases their COM objects.
		/// </remarks>
		/// <param name="window"></param>
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

		/// <summary>
		///		Binds to a control to hear events.
		/// </summary>
		/// <param name="control">The control to bind to.</param>
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

		/// <summary>
		///		Removes event handlers from a control.
		/// </summary>
		/// <param name="control">The control to unbind from.</param>
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

		/// <summary>
		///		Hears button clicks for any custom CommandBarButtons.
		///		<seealso cref="Microsoft.Office._CommandBarButtonEvents_ClickEventHandler"/>
		/// </summary>
		/// <param name="Ctrl">The control that was clicked.</param>
		/// <param name="CancelDefault">Whether to cancel the default action.</param>
		private void button_Click(CommandBarButton Ctrl, ref bool CancelDefault)
		{
			ToolbarButton control = (ToolbarButton)controlControlMap[Ctrl.Tag];
			OfficeWindow window = (OfficeWindow)controlWindowMap[Ctrl.Tag];
			control.OnClick(window);
		}

		/// <summary>
		///		Hears button clicks for any custom CommandBarComboBox.
		///		<seealso cref="Microsoft.Office._CommandBarComboBoxEvents_ChangeEventHandler"/>
		/// </summary>
		/// <param name="Ctrl">The combo box that was changed.</param>
		private void combo_Change(CommandBarComboBox Ctrl)
		{
			ToolbarComboBox control = (ToolbarComboBox)controlControlMap[Ctrl.Tag];
			OfficeWindow window = (OfficeWindow)controlWindowMap[Ctrl.Tag];
			control.OnChange(window);
		}
	}

	/// <summary>
	///		An abstract OfficeWindow
	/// </summary>
	/// <remarks>
	///		UI managers for specific office applications should define their own specific
	///		classes from this.
	/// </remarks>
	public abstract class OfficeWindow
	{
		/// <summary>
		///		The CommandBars for the OfficeWindow.
		/// </summary>
		public abstract CommandBars CommandBars
		{
			get;
		}

		/// <summary>
		///		Activates the window, bringing it to the front of the display.
		/// </summary>
		public abstract void Activate();

		/// <summary>
		///		Closes the window.
		/// </summary>
		public abstract void Close();
 
		/// <summary>
		///		The x-coordinate of the left of the window in pixels.
		/// </summary>
		public abstract int Left
		{
			get;
			set;
		}
 
		/// <summary>
		///		The y-coordinate of the top of the window in pixels.
		/// </summary>
		public abstract int Top
		{
			get;
			set;
		}
 
		/// <summary>
		///		The width of the window in pixels.
		/// </summary>
		public abstract int Width
		{
			get;
			set;
		}
 
		/// <summary>
		///		The height of the window in pixels.
		/// </summary>
		public abstract int Height
		{
			get;
			set;
		}
	}
}
