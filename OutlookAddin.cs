/*
 * $HeadURL$
 * $LastChangedBy$
 * $Date$
 * $Revision$
 */

using System;
using RlOutlook = Microsoft.Office.Interop.Outlook;

namespace BlueprintIT.Office.Outlook
{
	/// <summary>
	///		Defines the base COM interface for an outlook addin.
	/// </summary>
	public abstract class OutlookAddin: Extensibility.IDTExtensibility2, IDisposable
	{
		/// <summary>
		///		The UI manager for this addin.
		/// </summary>
		private OutlookUIManager manager;
		/// <summary>
		///		This addin.
		/// </summary>
		private object addInInstance;

		/// <summary>
		///		The outlook application object.
		/// </summary>
		public RlOutlook.Application Application
		{
			get
			{
				return manager.Application;
			}
		}

		/// <summary>
		///		The UI manager.
		/// </summary>
		public OutlookUIManager UIManager
		{
			get
			{
				return manager;
			}
		}

		/// <summary>
		///		Called when Outlook has initialised.
		/// </summary>
		public abstract void OnOutlookOpen();

		/// <summary>
		///		Called during Outlook startup.
		/// </summary>
		/// <remarks>
		///		Should not be called from code.
		/// </remarks>
		/// <param name="application">The Outlook application.</param>
		/// <param name="connectMode">The startup type.</param>
		/// <param name="addInInst">This addin.</param>
		/// <param name="custom">Any custom arguments passed to Outlook.</param>
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

		/// <summary>
		///		Called when an addin is released from Outlook.
		/// </summary>
		/// <remarks>
		///		Should not be called from code.
		/// </remarks>
		/// <param name="disconnectMode">The shutdown mode.</param>
		/// <param name="custom">Any custom arguments.</param>
		public void OnDisconnection(Extensibility.ext_DisconnectMode disconnectMode, ref System.Array custom)
		{
			if(disconnectMode != Extensibility.ext_DisconnectMode.ext_dm_HostShutdown)
			{
				OnBeginShutdown(ref custom);
			}
		}

		/// <summary>
		///		Called when the Outlook addin list changes.
		/// </summary>
		/// <param name="custom">Custom arguments.</param>
		public virtual void OnAddInsUpdate(ref System.Array custom)
		{
		}

		/// <summary>
		///		Called by Outlook when startup is complete.
		/// </summary>
		/// <remarks>
		///		Should not be called from code.
		/// </remarks>
		/// <param name="custom">Custom arguments.</param>
		public void OnStartupComplete(ref System.Array custom)
		{
			OnOutlookOpen();
		}

		/// <summary>
		///		Disposes of the addin, the UI manager and anything else that needs releasing.
		/// </summary>
		/// <remarks>
		///		It is safe to call this multiple times.
		/// </remarks>
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

		/// <summary>
		///		Called by Outlook on shutdown.
		/// </summary>
		/// <remarks>
		/// Should not be called from code.
		/// </remarks>
		/// <param name="custom">Custom arguments</param>
		public void OnBeginShutdown(ref System.Array custom)
		{
			Dispose();
		}

		/// <summary>
		///		An event handler listening for Outlook close.
		/// </summary>
		/// <param name="sender">The UI manager that detected Outlook closing.</param>
		private void OnOutlookClosed(OfficeUIManager sender)
		{
			Dispose();
		}
	}
}
