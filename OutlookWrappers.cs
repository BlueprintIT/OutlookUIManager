using System;
using RlOutlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Core;
using BlueprintIT.Office.Outlook;

namespace BlueprintIT.Office.Outlook
{
	public delegate void InspectorEventHandler(OutlookInspector inspector);

	public delegate void ExplorerEventHandler(OutlookExplorer explorer);

	public class OutlookInspector: OfficeWindow
	{
		private RlOutlook.Inspector inspector;
		private OutlookItem item;

		private RlOutlook.InspectorEvents_ActivateEventHandler activateEvent;
		private RlOutlook.InspectorEvents_DeactivateEventHandler deactivateEvent;
		private RlOutlook.InspectorEvents_CloseEventHandler closeEvent;
		private RlOutlook.ItemEvents_CloseEventHandler itemCloseEvent;

		public event InspectorEventHandler Activated;
		public event InspectorEventHandler Deactivated;
		public event InspectorEventHandler Closed;

		public OutlookInspector(RlOutlook.Inspector inspector)
		{
			this.inspector=inspector;
			this.item = OutlookItem.CreateItem(inspector.CurrentItem);
			BindEvents();
		}

		public void Dispose()
		{
			UnBindEvents();
#if (COMRELEASE)
			System.Runtime.InteropServices.Marshal.ReleaseComObject(inspector);
#endif
			inspector=null;
			item.Dispose();
			item=null;
		}

		private void BindEvents()
		{
			RlOutlook.InspectorEvents_Event events = (RlOutlook.InspectorEvents_Event)inspector;
			activateEvent = new RlOutlook.InspectorEvents_ActivateEventHandler(inspector_Activate);
			deactivateEvent =	new RlOutlook.InspectorEvents_DeactivateEventHandler(inspector_Deactivate);
			closeEvent =	new RlOutlook.InspectorEvents_CloseEventHandler(inspector_Close);
			itemCloseEvent = new RlOutlook.ItemEvents_CloseEventHandler(item_Close);
			events.Activate+=activateEvent;
			events.Deactivate+=deactivateEvent;
			events.Close+=closeEvent;
			item.Events.Close+=itemCloseEvent;
		}

		private void UnBindEvents()
		{
			if (activateEvent!=null)
			{
				RlOutlook.InspectorEvents_Event events = (RlOutlook.InspectorEvents_Event)inspector;
				events.Activate-=activateEvent;
				events.Deactivate-=deactivateEvent;
				events.Close-=closeEvent;
				item.Events.Close-=itemCloseEvent;
			}
		}

		public override CommandBars CommandBars
		{
			get
			{
				return inspector.CommandBars;
			}
		}

		public RlOutlook.Inspector Inspector
		{
			get
			{
				return inspector;
			}
		}

		public OutlookItem CurrentItem
		{
			get
			{
				return item;
			}
		}

		public override void Activate()
		{
			inspector.Activate();
		}

		public override void Close()
		{
			inspector.Close(RlOutlook.OlInspectorClose.olPromptForSave);
		}
 
		public override int Left
		{
			get
			{
				return inspector.Left;
			}

			set
			{
				inspector.Left=value;
			}
		}
 
		public override int Top
		{
			get
			{
				return inspector.Top;
			}

			set
			{
				inspector.Top=value;
			}
		}
 
		public override int Width
		{
			get
			{
				return inspector.Width;
			}

			set
			{
				inspector.Width=value;
			}
		}
 
		public override int Height
		{
			get
			{
				return inspector.Height;
			}

			set
			{
				inspector.Height=value;
			}
		}

		private void inspector_Activate()
		{
			if (Activated!=null)
			{
				Activated(this);
			}
		}

		private void inspector_Deactivate()
		{
			if (Deactivated!=null)
			{
				Deactivated(this);
			}
		}

		private void inspector_Close()
		{
			if (Closed!=null)
			{
				Closed(this);
			}
			Dispose();
		}

		private void item_Close(ref bool Cancel)
		{
			if (Closed!=null)
			{
				Closed(this);
			}
			Dispose();
		}
	}

	public class OutlookExplorer: OfficeWindow
	{
		private RlOutlook.Explorer explorer;

		private RlOutlook.ExplorerEvents_ActivateEventHandler activateEvent;
		private RlOutlook.ExplorerEvents_DeactivateEventHandler deactivateEvent;
		private RlOutlook.ExplorerEvents_CloseEventHandler closeEvent;

		public event ExplorerEventHandler Activated;
		public event ExplorerEventHandler Deactivated;
		public event ExplorerEventHandler Closed;

		public OutlookExplorer(RlOutlook.Explorer explorer)
		{
			this.explorer=explorer;
			BindEvents();
		}

		public void Dispose()
		{
			UnBindEvents();
#if (COMRELEASE)
			System.Runtime.InteropServices.Marshal.ReleaseComObject(explorer);
#endif
			explorer=null;
		}

		private void BindEvents()
		{
			RlOutlook.ExplorerEvents_Event events = (RlOutlook.ExplorerEvents_Event)explorer;
			activateEvent = new Microsoft.Office.Interop.Outlook.ExplorerEvents_ActivateEventHandler(explorer_Activate);
			deactivateEvent = new Microsoft.Office.Interop.Outlook.ExplorerEvents_DeactivateEventHandler(explorer_Deactivate);
			closeEvent = new Microsoft.Office.Interop.Outlook.ExplorerEvents_CloseEventHandler(explorer_Close);
			events.Activate+=activateEvent;
			events.Deactivate+=deactivateEvent;
			events.Close+=closeEvent;
		}

		private void UnBindEvents()
		{
			if (activateEvent!=null)
			{
				RlOutlook.ExplorerEvents_Event events = (RlOutlook.ExplorerEvents_Event)explorer;
				events.Activate-=activateEvent;
				events.Deactivate-=deactivateEvent;
				events.Close-=closeEvent;
			}
		}

		public RlOutlook.MAPIFolder CurrentFolder
		{
			get
			{
				return explorer.CurrentFolder;
			}
		}

		public override CommandBars CommandBars
		{
			get
			{
				return explorer.CommandBars;
			}
		}

		public override void Activate()
		{
			explorer.Activate();
		}

		public override void Close()
		{
			explorer.Close();
		}
 
		public override int Left
		{
			get
			{
				return explorer.Left;
			}

			set
			{
				explorer.Left=value;
			}
		}
 
		public override int Top
		{
			get
			{
				return explorer.Top;
			}

			set
			{
				explorer.Top=value;
			}
		}
 
		public override int Width
		{
			get
			{
				return explorer.Width;
			}

			set
			{
				explorer.Width=value;
			}
		}
 
		public override int Height
		{
			get
			{
				return explorer.Height;
			}

			set
			{
				explorer.Height=value;
			}
		}

		private void explorer_Activate()
		{
			if (Activated!=null)
			{
				Activated(this);
			}
		}

		private void explorer_Deactivate()
		{
			if (Deactivated!=null)
			{
				Deactivated(this);
			}
		}

		private void explorer_Close()
		{
			if (Closed!=null)
			{
				Closed(this);
			}
			Dispose();
		}
	}
}
