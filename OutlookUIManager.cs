/*
 * $HeadURL$
 * $LastChangedBy$
 * $Date$
 * $Revision$
 */

using System;
using System.Collections;
using System.IO;

using RlOutlook = Microsoft.Office.Interop.Outlook;

namespace BlueprintIT.Office.Outlook
{
	public delegate void OutlookEventHandler(UIManager sender);

	public class OutlookUIManager: UIManager
	{
		private RlOutlook.Application application;
		private RlOutlook.Inspectors inspectors;
		private RlOutlook.Explorers explorers;

		private OutlookAddin addin;

		private IList inspectorCache;
		private IList explorerCache;

		private TextWriter logger;

		private RlOutlook.InspectorsEvents_NewInspectorEventHandler newInspectorEvent;
		private RlOutlook.ExplorersEvents_NewExplorerEventHandler newExplorerEvent;
		private InspectorEventHandler inspectorCloseEvent;
		private ExplorerEventHandler explorerCloseEvent;

		private Toolbars inspectorToolbars;
		private Toolbars explorerToolbars;

		public event InspectorEventHandler InspectorOpen;
		public event InspectorEventHandler InspectorClose;

		public event ExplorerEventHandler ExplorerOpen;
		public event ExplorerEventHandler ExplorerClose;

		public event OutlookEventHandler OutlookClose;
		internal event OutlookEventHandler OutlookClosed;

		public OutlookUIManager(RlOutlook.Application application, OutlookAddin addin, string progid): base(progid)
		{
			this.addin=addin;
			this.application=application;
			explorers=application.Explorers;
			inspectors=application.Inspectors;

			inspectorCache = new ArrayList();
			explorerCache = new ArrayList();

			inspectorToolbars = new Toolbars(this);
			explorerToolbars = new Toolbars(this);

			logger = new StreamWriter("c:\\log.txt",true);

			BindEvents();

			foreach (RlOutlook.Inspector inspector in inspectors)
			{
				OnInspectorOpen(new OutlookInspector(inspector));
			}

			foreach (RlOutlook.Explorer explorer in explorers)
			{
				OnExplorerOpen(new OutlookExplorer(explorer));
			}
		}

		internal override void log(string text)
		{
			logger.WriteLine(text);
			logger.Flush();
		}

		public override IList Windows
		{
			get
			{
				ArrayList list = new ArrayList(explorerCache);
				list.AddRange(inspectorCache);
				return list;
			}
		}

		public Toolbars ExplorerToolbars
		{
			get
			{
				return explorerToolbars;
			}
		}

		public Toolbars InspectorToolbars
		{
			get
			{
				return inspectorToolbars;
			}
		}

		public RlOutlook.Application Application
		{
			get
			{
				return application;
			}
		}

		public IList Inspectors
		{
			get
			{
				return ArrayList.ReadOnly(inspectorCache);
			}
		}
		
		public IList Explorers
		{
			get
			{
				return ArrayList.ReadOnly(explorerCache);
			}
		}
		
		public override void Dispose()
		{
			base.Dispose();
			if (application!=null)
			{
				UnbindEvents();
				logger.Close();
#if (COMRELEASE)
				System.Runtime.InteropServices.Marshal.ReleaseComObject(inspectors);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(explorers);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(application);
#endif
				inspectors=null;
				explorers=null;
				application=null;
			}
		}

		protected override Toolbars GetWindowToolbars(OfficeWindow window)
		{
			if (window is OutlookInspector)
			{
				return inspectorToolbars;
			}
			else if (window is OutlookExplorer)
			{
				return explorerToolbars;
			}
			return null;
		}

		private void BindEvents()
		{
			newInspectorEvent = new RlOutlook.InspectorsEvents_NewInspectorEventHandler(inspectors_NewInspector);
			newExplorerEvent = new RlOutlook.ExplorersEvents_NewExplorerEventHandler(explorers_NewExplorer);
			inspectorCloseEvent = new InspectorEventHandler(OnInspectorClose);
			explorerCloseEvent = new ExplorerEventHandler(OnExplorerClose);
			inspectors.NewInspector+=newInspectorEvent;
			explorers.NewExplorer+=newExplorerEvent;
		}

		private void UnbindEvents()
		{
			if (newInspectorEvent!=null)
			{
				inspectors.NewInspector-=newInspectorEvent;
				explorers.NewExplorer-=newExplorerEvent;
				foreach (OutlookInspector inspector in inspectorCache)
				{
					inspector.Closed-=inspectorCloseEvent;
					inspector.Dispose();
				}
				foreach (OutlookExplorer explorer in explorerCache)
				{
					explorer.Closed-=explorerCloseEvent;
					explorer.Dispose();
				}
			}
			inspectorCache.Clear();
			explorerCache.Clear();
		}

		private void inspectors_NewInspector(RlOutlook.Inspector Inspector)
		{
#if (COMRELEASE)
			System.Runtime.InteropServices.Marshal.ReleaseComObject(Inspector);
#endif
			OutlookInspector inspector = new OutlookInspector(inspectors[inspectors.Count]);
			log(inspector.CurrentItem.EntryID);
			OnInspectorOpen(inspector);
		}

		private void explorers_NewExplorer(RlOutlook.Explorer Explorer)
		{
#if (COMRELEASE)
			System.Runtime.InteropServices.Marshal.ReleaseComObject(Explorer);
#endif
			OutlookExplorer explorer = new OutlookExplorer(explorers[explorers.Count]);
			OnExplorerOpen(explorer);
		}

		private void OnInspectorOpen(OutlookInspector inspector)
		{
			log("Inspector opened");
			inspectorCache.Add(inspector);
			inspector.Closed+=inspectorCloseEvent;
			OnWindowOpen(inspector);
			if (InspectorOpen!=null)
			{
				InspectorOpen(inspector);
			}
		}

		private void OnInspectorClose(OutlookInspector inspector)
		{
			OnWindowClose(inspector);
			if (InspectorClose!=null)
			{
				InspectorClose(inspector);
			}
			inspectorCache.Remove(inspector);
			log("Inspector closed ("+inspectorCache.Count+")");
			if ((explorerCache.Count==0)&&(inspectorCache.Count==0))
			{
				OnOutlookClose();
			}
		}

		private void OnExplorerOpen(OutlookExplorer explorer)
		{
			log("Explorer opened");
			explorerCache.Add(explorer);
			explorer.Closed+=explorerCloseEvent;
			OnWindowOpen(explorer);
			if (ExplorerOpen!=null)
			{
				ExplorerOpen(explorer);
			}
		}

		private void OnExplorerClose(OutlookExplorer explorer)
		{
			OnWindowClose(explorer);
			if (ExplorerClose!=null)
			{
				ExplorerClose(explorer);
			}
			explorerCache.Remove(explorer);
			log("Explorer closed ("+explorerCache.Count+")");
			if ((explorerCache.Count==0)&&(inspectorCache.Count==0))
			{
				OnOutlookClose();
			}
		}

		private void OnOutlookClose()
		{
			log("Outlook closed");
			if (OutlookClose!=null)
			{
				OutlookClose(this);
			}
			if (OutlookClosed!=null)
			{
				OutlookClosed(this);
			}
			Dispose();
		}
	}
}
