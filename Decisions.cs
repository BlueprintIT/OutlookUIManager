using System;
using System.Collections;
using RlOutlook = Microsoft.Office.Interop.Outlook;

namespace BlueprintIT.Office
{
	public abstract class WindowBasedValue
	{
	}

	public class BooleanValue: WindowBasedValue
	{
		private static BooleanValue TRUE = new BooleanValue(true);
		private static BooleanValue FALSE = new BooleanValue(false);

		private bool normal;

		public BooleanValue(bool normal)
		{
			this.normal=normal;
		}

		public static implicit operator BooleanValue(bool value)
		{
			if (value)
			{
				return TRUE;
			}
			else
			{
				return FALSE;
			}
		}

		public virtual bool GetValue(OfficeWindow window)
		{
			return normal;
		}
	}

	namespace Outlook
	{
		public class TypeBasedBoolean: BooleanValue
		{
			private IDictionary typeMap = new Hashtable();

			public TypeBasedBoolean(bool normal): base(normal)
			{
			}

			public void SetValue(RlOutlook.OlItemType type, bool value)
			{
				typeMap[type]=value;
			}

			public void EnableForType(RlOutlook.OlItemType type)
			{
				SetValue(type,true);
			}

			public void DisableForType(RlOutlook.OlItemType type)
			{
				SetValue(type,false);
			}

			public override bool GetValue(OfficeWindow window)
			{
				RlOutlook.OlItemType type;
				if (window is OutlookInspector)
				{
					type = ((OutlookInspector)window).CurrentItem.Type;
				}
				else if (window is OutlookExplorer)
				{
					type = ((OutlookExplorer)window).CurrentFolder.DefaultItemType;
				}
				else
				{
					return base.GetValue(window);
				}
				if (typeMap.Contains(type))
				{
					return (bool)typeMap[type];
				}
				else
				{
					return base.GetValue(window);
				}
			}
		}
	}
}
