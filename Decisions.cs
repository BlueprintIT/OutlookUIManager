/*
 * $HeadURL$
 * $LastChangedBy$
 * $Date$
 * $Revision$
 */

using System;
using System.Collections;
using RlOutlook = Microsoft.Office.Interop.Outlook;

namespace BlueprintIT.Office
{
	/// <summary>
	///		Defines a value that is dependant on the OfficeWindow it is applied to.
	/// </summary>
	public abstract class WindowBasedValue
	{
	}

	/// <summary>
	///		Defines a boolean value that is dependant on the OfficeWindow it is applied to.
	/// </summary>
	public class BooleanValue: WindowBasedValue
	{
		/// <summary>
		///		A BooleanValue that is always true.
		/// </summary>
		private static BooleanValue TRUE = new BooleanValue(true);
		/// <summary>
		///		A BooleanValue that is always false.
		/// </summary>
		private static BooleanValue FALSE = new BooleanValue(false);

		/// <summary>
		///		The default state of the value.
		/// </summary>
		private bool normal;

		/// <summary>
		///		Creates a new value with a default state.
		/// </summary>
		/// <param name="normal"></param>
		protected BooleanValue(bool normal)
		{
			this.normal=normal;
		}

		/// <summary>
		///		Casts a bool value to a BooleanValue.
		/// </summary>
		/// <param name="value">The bool value to convert.</param>
		/// <returns>Either BooleanValue.TRUE or BooleanValue.FALSE.</returns>
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

		/// <summary>
		///		Returns the value for the given window.
		/// </summary>
		/// <param name="window">The window.</param>
		/// <returns>True or false depending on the window.</returns>
		public virtual bool GetValue(OfficeWindow window)
		{
			return normal;
		}
	}

	namespace Outlook
	{
		/// <summary>
		///		Defines a BooleanValue that is dependant on the type of Outlook item displayed
		///		in the window.
		/// </summary>
		public class TypeBasedBoolean: BooleanValue
		{
			/// <summary>
			///		Maps from Outlook types to bool values.
			/// </summary>
			private IDictionary typeMap = new Hashtable();

			/// <summary>
			///		Creates a new instance with a default value setting.
			/// </summary>
			/// <param name="normal"></param>
			public TypeBasedBoolean(bool normal): base(normal)
			{
			}

			/// <summary>
			///		Sets the value for a given Outlook item type.
			/// </summary>
			/// <param name="type">The item type.</param>
			/// <param name="value">The value.</param>
			public void SetValue(RlOutlook.OlItemType type, bool value)
			{
				typeMap[type]=value;
			}

			/// <summary>
			///		Sets the value to true for a given type.
			/// </summary>
			/// <param name="type">The Outlook item type.</param>
			public void EnableForType(RlOutlook.OlItemType type)
			{
				SetValue(type,true);
			}

			/// <summary>
			///		Sets the value to false for a given type.
			/// </summary>
			/// <param name="type">The Outlook item type.</param>
			public void DisableForType(RlOutlook.OlItemType type)
			{
				SetValue(type,false);
			}

			/// <summary>
			///		Returns the value for a given window.
			/// </summary>
			/// <param name="window">The window.</param>
			/// <returns>The value.</returns>
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
