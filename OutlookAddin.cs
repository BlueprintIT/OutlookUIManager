using System;
using RlOutlook = Microsoft.Office.Interop.Outlook;

namespace BlueprintIT.Office.Outlook
{
	public abstract class OutlookAddin: Extensibility.IDTExtensibility2, IDisposable
	{
		private OutlookUIManager manager;
		private object addInInstance;

		public RlOutlook.Application Application
		{
			get
			{
				return manager.Application;
			}
		}

		public OutlookUIManager UIManager
		{
			get
			{
				return manager;
			}
		}

		public abstract void OnOutlookOpen();

		public void OnConnection(object application, Extensibility.ext_ConnectMode connectMode, object addInInst, ref System.Array custom)
		{
			addInInstance = addInInst;
			System.Runtime.InteropServices.RegistrationServices reg = new System.Runtime.InteropServices.RegistrationServices();
			string progid = reg.GetProgIdForType(this.GetType());
			manager = new OutlookUIManager((RlOutlook.Application)application,this,progid);
			manager.OutlookClosed+=new OutlookEventHandler(OnOutlookClosed);

			if(connectMode != Extensibility.ext_ConnectMode.ext_cm_Startup)
			{
				OnStartupComplete(ref custom);
			}
		}

		public void OnDisconnection(Extensibility.ext_DisconnectMode disconnectMode, ref System.Array custom)
		{
			if(disconnectMode != Extensibility.ext_DisconnectMode.ext_dm_HostShutdown)
			{
				OnBeginShutdown(ref custom);
			}
		}

		public void OnAddInsUpdate(ref System.Array custom)
		{
		}

		public void OnStartupComplete(ref System.Array custom)
		{
			OnOutlookOpen();
		}

		public void Dispose()
		{
			if (addInInstance!=null)
			{
#if (COMRELEASE)
				System.Runtime.InteropServices.Marshal.ReleaseComObject(addInInstance);
#endif
				addInInstance=null;
				manager.Dispose();
				manager=null;
			}
			GC.Collect();
			GC.WaitForPendingFinalizers();
		}

		public void OnBeginShutdown(ref System.Array custom)
		{
			Dispose();
		}

		private void OnOutlookClosed(UIManager sender)
		{
			Dispose();
		}
	}
}
