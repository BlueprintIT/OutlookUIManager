/*
 * $HeadURL$
 * $LastChangedBy$
 * $Date$
 * $Revision$
 */

using System;
using System.Reflection;
using RlOutlook = Microsoft.Office.Interop.Outlook;
using MAPI33;
using MAPI33.MapiTypes;

namespace BlueprintIT.Office.Outlook
{
	#region Specific Items
	public abstract class OutlookNonNoteItem: OutlookItem
	{
		private RlOutlook.UserProperties userProperties;

		protected OutlookNonNoteItem(object item): base(item)
		{
		}

		public override void Dispose()
		{
#if (COMRELEASE)
			if (userProperties!=null)
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(userProperties);
			}
#endif
			userProperties=null;
			base.Dispose ();
		}


		public void ShowCategoriesDialog()
		{
			CallMethod("ShowCategoriesDialog");
		}

		public RlOutlook.Actions Actions
		{
			get
			{
				return (RlOutlook.Actions)GetProperty("Actions");
			}
		}

		public RlOutlook.Attachments Attachments
		{
			get
			{
				return (RlOutlook.Attachments)GetProperty("Attachments");
			}
		}

		public string BillingInformation
		{
			get
			{
				return (string)GetProperty("BillingInformation");
			}

			set
			{
				SetProperty("BillingInformation",value);
			}
		}

		public string Companies
		{
			get
			{
				return (string)GetProperty("Companies");
			}

			set
			{
				SetProperty("Companies",value);
			}
		}

		public string ConversationIndex
		{
			get
			{
				return (string)GetProperty("ConversationIndex");
			}
		}

		public string ConversationTopic
		{
			get
			{
				return (string)GetProperty("ConversationTopic");
			}
		}

		public RlOutlook.FormDescription FormDescription
		{
			get
			{
				return (RlOutlook.FormDescription)GetProperty("FormDescription");
			}
		}

		public RlOutlook.OlImportance Importance
		{
			get
			{
				return (RlOutlook.OlImportance)GetProperty("Importance");
			}

			set
			{
				SetProperty("Importance",value);
			}
		}

		public string Mileage
		{
			get
			{
				return (string)GetProperty("Mileage");
			}

			set
			{
				SetProperty("Mileage",value);
			}
		}

		public bool NoAging
		{
			get
			{
				return (bool)GetProperty("NoAging");
			}

			set
			{
				SetProperty("NoAging",value);
			}
		}

		public int OutlookInternalVersion
		{
			get
			{
				return (int)GetProperty("OutlookInternalVersion");
			}
		}

		public string OutlookVersion
		{
			get
			{
				return (string)GetProperty("OutlookVersion");
			}
		}

		public RlOutlook.OlSensitivity Sensitivity
		{
			get
			{
				return (RlOutlook.OlSensitivity)GetProperty("Sensitivity");
			}

			set
			{
				SetProperty("Sensitivity",value);
			}
		}

		public bool UnRead
		{
			get
			{
				return (bool)GetProperty("UnRead");
			}

			set
			{
				SetProperty("UnRead",value);
			}
		}

		public RlOutlook.UserProperties UserProperties
		{
			get
			{
				if (userProperties==null)
				{
					userProperties = (RlOutlook.UserProperties)GetProperty("UserProperties");
				}
				return userProperties;
			}
		}
	}

	public abstract class OutlookMailMeetingItem: OutlookNonNoteItem
	{
		protected OutlookMailMeetingItem(object item): base(item)
		{
		}

		public bool AutoForwarded
		{
			get
			{
				return (bool)GetProperty("AutoForwarded");
			}

			set
			{
				SetProperty("AutoForwarded",value);
			}
		}

		public DateTime DeferredDeliveryTime
		{
			get
			{
				return (DateTime)GetProperty("DeferredDeliveryTime");
			}

			set
			{
				SetProperty("DeferredDeliveryTime",value);
			}
		}

		public bool DeleteAfterSubmit
		{
			get
			{
				return (bool)GetProperty("DeleteAfterSubmit");
			}

			set
			{
				SetProperty("DeleteAfterSubmit",value);
			}
		}

		public DateTime FlagDueBy
		{
			get
			{
				return (DateTime)GetProperty("FlagDueBy");
			}

			set
			{
				SetProperty("FlagDueBy",value);
			}
		}

		public string FlagRequest
		{
			get
			{
				return (string)GetProperty("FlagRequest");
			}

			set
			{
				SetProperty("FlagRequest",value);
			}
		}

		public RlOutlook.OlFlagStatus FlagStatus
		{
			get
			{
				return (RlOutlook.OlFlagStatus)GetProperty("FlagStatus");
			}

			set
			{
				SetProperty("FlagStatus",value);
			}
		}

		public bool OriginatorDeliveryReportRequested
		{
			get
			{
				return (bool)GetProperty("OriginatorDeliveryReportRequested");
			}

			set
			{
				SetProperty("OriginatorDeliveryReportRequested",value);
			}
		}

		public RlOutlook.Recipients ReplyRecipients
		{
			get
			{
				return (RlOutlook.Recipients)GetProperty("ReplyRecipients");
			}
		}

		public RlOutlook.MAPIFolder SaveSentMessageFolder
		{
			get
			{
				return (RlOutlook.MAPIFolder)GetProperty("SaveSentMessageFolder");
			}
		}

		public bool Sent
		{
			get
			{
				return (bool)GetProperty("Sent");
			}
		}

		public bool Submitted
		{
			get
			{
				return (bool)GetProperty("Submitted");
			}
		}

		public DateTime ExpiryTime
		{
			get
			{
				return (DateTime)GetProperty("ExpiryTime");
			}

			set
			{
				SetProperty("ExpiryTime",value);
			}
		}

		public DateTime ReminderTime
		{
			get
			{
				return (DateTime)GetProperty("ReminderTime");
			}

			set
			{
				SetProperty("ReminderTime",value);
			}
		}

		public string SenderName
		{
			get
			{
				return (string)GetProperty("SenderName");
			}
		}

		public DateTime SentOn
		{
			get
			{
				return (DateTime)GetProperty("SentOn");
			}
		}

		public bool ReminderSet
		{
			get
			{
				return (bool)GetProperty("ReminderSet");
			}

			set
			{
				SetProperty("ReminderSet",value);
			}
		}

		public RlOutlook.Recipients Recipients
		{
			get
			{
				return (RlOutlook.Recipients)GetProperty("Recipients");
			}
		}

		public RlOutlook.MailItem Reply()
		{
			return (RlOutlook.MailItem)CallMethod("ReplyAll");
		}

		public RlOutlook.MailItem ReplyAll()
		{
			return (RlOutlook.MailItem)CallMethod("ReplyAll");
		}

		public void Send()
		{
			CallMethod("Send");
		}
	}
	public class OutlookAppointmentItem: OutlookNonNoteItem
	{
		public OutlookAppointmentItem(object item): base(item)
		{
		}

		public override RlOutlook.OlItemType Type
		{
			get
			{
				return RlOutlook.OlItemType.olAppointmentItem;
			}
		}

		public bool ReminderSet
		{
			get
			{
				return (bool)GetProperty("ReminderSet");
			}

			set
			{
				SetProperty("ReminderSet",value);
			}
		}

		public RlOutlook.Recipients Recipients
		{
			get
			{
				return (RlOutlook.Recipients)GetProperty("Recipients");
			}
		}

		public string NetMeetingServer
		{
			get
			{
				return (string)GetProperty("NetMeetingServer");
			}

			set
			{
				SetProperty("NetMeetingServer",value);
			}
		}

		public void Send()
		{
			CallMethod("Send");
		}
	}

	public class OutlookContactItem: OutlookNonNoteItem, RlOutlook._ContactItem
	{
		public OutlookContactItem(object item): base(item)
		{
		}

		protected override Tags GetMapiTag(string property, out bool found)
		{
			if (property=="Body")
			{
				found=true;
				return Tags.PR_BODY;
			}
			if (property=="ReferredBy")
			{
				found=true;
				return Tags.PR_REFERRED_BY_NAME;
			}
			return base.GetMapiTag(property,out found);
		}

		protected override MAPINAMEID GetMapiID(string property)
		{
			MAPINAMEIDInt id = new MAPINAMEIDInt();
			id.guid = CdoPropSetID3;
			if (property=="Email1Address")
			{
				id.id = 0x8083;
				return id;
			}
			if (property=="Email1AddressType")
			{
				id.id = 0x8082;
				return id;
			}
			if (property=="Email1DisplayName")
			{
				id.id = 0x8084;
				return id;
			}
			if (property=="Email1EntryID")
			{
				id.id = 0x8085;
				return id;
			}
			if (property=="Email2Address")
			{
				id.id = 0x8093;
				return id;
			}
			if (property=="Email2AddressType")
			{
				id.id = 0x8092;
				return id;
			}
			if (property=="Email2DisplayName")
			{
				id.id = 0x8094;
				return id;
			}
			if (property=="Email2EntryID")
			{
				id.id = 0x8095;
				return id;
			}
			if (property=="Email3Address")
			{
				id.id = 0x80A3;
				return id;
			}
			if (property=="Email3AddressType")
			{
				id.id = 0x80A2;
				return id;
			}
			if (property=="Email3DisplayName")
			{
				id.id = 0x80A4;
				return id;
			}
			if (property=="Email3EntryID")
			{
				id.id = 0x80A5;
				return id;
			}
			if (property=="IMAddress")
			{
				id.id = 0x8062;
				return id;
			}
			if (property=="NetMeetingAlias")
			{
				id.id = 0x0;
				return id;
			}
			return base.GetMapiID(property);
		}

		public override RlOutlook.OlItemType Type
		{
			get
			{
				return RlOutlook.OlItemType.olContactItem;
			}
		}

#if (OL2003)
		public void AddPicture(string arg0)
		{
			CallMethod("AddPicture",arg0);
		}

		public void RemovePicture()
		{
			CallMethod("RemovePicture");
		}

		public bool HasPicture
		{
			get
			{
				return (bool)GetProperty("HasPicture");
			}
		}
#endif

		public string Account
		{
			get
			{
				return (string)GetProperty("Account");
			}

			set
			{
				SetProperty("Account",value);
			}
		}

		public DateTime Anniversary
		{
			get
			{
				return (DateTime)GetProperty("Anniversary");
			}

			set
			{
				SetProperty("Anniversary",value);
			}
		}

		public string AssistantName
		{
			get
			{
				return (string)GetProperty("AssistantName");
			}

			set
			{
				SetProperty("AssistantName",value);
			}
		}

		public string AssistantTelephoneNumber
		{
			get
			{
				return (string)GetProperty("AssistantTelephoneNumber");
			}

			set
			{
				SetProperty("AssistantTelephoneNumber",value);
			}
		}

		public DateTime Birthday
		{
			get
			{
				return (DateTime)GetProperty("Birthday");
			}

			set
			{
				SetProperty("Birthday",value);
			}
		}

		public string Business2TelephoneNumber
		{
			get
			{
				return (string)GetProperty("Business2TelephoneNumber");
			}

			set
			{
				SetProperty("Business2TelephoneNumber",value);
			}
		}

		public string BusinessAddress
		{
			get
			{
				return (string)GetProperty("BusinessAddress");
			}

			set
			{
				SetProperty("BusinessAddress",value);
			}
		}

		public string BusinessAddressCity
		{
			get
			{
				return (string)GetProperty("BusinessAddressCity");
			}

			set
			{
				SetProperty("BusinessAddressCity",value);
			}
		}

		public string BusinessAddressCountry
		{
			get
			{
				return (string)GetProperty("BusinessAddressCountry");
			}

			set
			{
				SetProperty("BusinessAddressCountry",value);
			}
		}

		public string BusinessAddressPostalCode
		{
			get
			{
				return (string)GetProperty("BusinessAddressPostalCode");
			}

			set
			{
				SetProperty("BusinessAddressPostalCode",value);
			}
		}

		public string BusinessAddressPostOfficeBox
		{
			get
			{
				return (string)GetProperty("BusinessAddressPostOfficeBox");
			}

			set
			{
				SetProperty("BusinessAddressPostOfficeBox",value);
			}
		}

		public string BusinessAddressState
		{
			get
			{
				return (string)GetProperty("BusinessAddressState");
			}

			set
			{
				SetProperty("BusinessAddressState",value);
			}
		}

		public string BusinessAddressStreet
		{
			get
			{
				return (string)GetProperty("BusinessAddressStreet");
			}

			set
			{
				SetProperty("BusinessAddressStreet",value);
			}
		}

		public string BusinessFaxNumber
		{
			get
			{
				return (string)GetProperty("BusinessFaxNumber");
			}

			set
			{
				SetProperty("BusinessFaxNumber",value);
			}
		}

		public string BusinessHomePage
		{
			get
			{
				return (string)GetProperty("BusinessHomePage");
			}

			set
			{
				SetProperty("BusinessHomePage",value);
			}
		}

		public string BusinessTelephoneNumber
		{
			get
			{
				return (string)GetProperty("BusinessTelephoneNumber");
			}

			set
			{
				SetProperty("BusinessTelephoneNumber",value);
			}
		}

		public string CallbackTelephoneNumber
		{
			get
			{
				return (string)GetProperty("CallbackTelephoneNumber");
			}

			set
			{
				SetProperty("CallbackTelephoneNumber",value);
			}
		}

		public string CarTelephoneNumber
		{
			get
			{
				return (string)GetProperty("CarTelephoneNumber");
			}

			set
			{
				SetProperty("CarTelephoneNumber",value);
			}
		}

		public string Children
		{
			get
			{
				return (string)GetProperty("Children");
			}

			set
			{
				SetProperty("Children",value);
			}
		}

		public string CompanyAndFullName
		{
			get
			{
				return (string)GetProperty("CompanyAndFullName");
			}

			set
			{
				SetProperty("CompanyAndFullName",value);
			}
		}

		public string CompanyLastFirstNoSpace
		{
			get
			{
				return (string)GetProperty("CompanyLastFirstNoSpace");
			}

			set
			{
				SetProperty("CompanyLastFirstNoSpace",value);
			}
		}

		public string CompanyLastFirstSpaceOnly
		{
			get
			{
				return (string)GetProperty("CompanyLastFirstSpaceOnly");
			}

			set
			{
				SetProperty("CompanyLastFirstSpaceOnly",value);
			}
		}

		public string CompanyMainTelephoneNumber
		{
			get
			{
				return (string)GetProperty("CompanyMainTelephoneNumber");
			}

			set
			{
				SetProperty("CompanyMainTelephoneNumber",value);
			}
		}

		public string CompanyName
		{
			get
			{
				return (string)GetProperty("CompanyName");
			}

			set
			{
				SetProperty("CompanyName",value);
			}
		}

		public string ComputerNetworkName
		{
			get
			{
				return (string)GetProperty("ComputerNetworkName");
			}

			set
			{
				SetProperty("ComputerNetworkName",value);
			}
		}

		public string CustomerID
		{
			get
			{
				return (string)GetProperty("CustomerID");
			}

			set
			{
				SetProperty("CustomerID",value);
			}
		}

		public string Department
		{
			get
			{
				return (string)GetProperty("Department");
			}

			set
			{
				SetProperty("Department",value);
			}
		}

		public string Email1Address
		{
			get
			{
				return (string)GetProperty("Email1Address");
			}

			set
			{
				SetProperty("Email1Address",value);
			}
		}

		public string Email1AddressType
		{
			get
			{
				return (string)GetProperty("Email1AddressType");
			}

			set
			{
				SetProperty("Email1AddressType",value);
			}
		}

		public string Email1DisplayName
		{
			get
			{
				return (string)GetProperty("Email1DisplayName");
			}

			set
			{
				SetProperty("Email1DisplayName",value);
			}
		}

		public string Email1EntryID
		{
			get
			{
				return (string)GetProperty("Email1EntryID");
			}

			set
			{
				SetProperty("Email1EntryID",value);
			}
		}

		public string Email2Address
		{
			get
			{
				return (string)GetProperty("Email2Address");
			}

			set
			{
				SetProperty("Email2Address",value);
			}
		}

		public string Email2AddressType
		{
			get
			{
				return (string)GetProperty("Email2AddressType");
			}

			set
			{
				SetProperty("Email2AddressType",value);
			}
		}

		public string Email2DisplayName
		{
			get
			{
				return (string)GetProperty("Email2DisplayName");
			}

			set
			{
				SetProperty("Email2DisplayName",value);
			}
		}

		public string Email2EntryID
		{
			get
			{
				return (string)GetProperty("Email2EntryID");
			}

			set
			{
				SetProperty("Email2EntryID",value);
			}
		}

		public string Email3Address
		{
			get
			{
				return (string)GetProperty("Email3Address");
			}

			set
			{
				SetProperty("Email3Address",value);
			}
		}

		public string Email3AddressType
		{
			get
			{
				return (string)GetProperty("Email3AddressType");
			}

			set
			{
				SetProperty("Email3AddressType",value);
			}
		}

		public string Email3DisplayName
		{
			get
			{
				return (string)GetProperty("Email3DisplayName");
			}

			set
			{
				SetProperty("Email3DisplayName",value);
			}
		}

		public string Email3EntryID
		{
			get
			{
				return (string)GetProperty("Email3EntryID");
			}

			set
			{
				SetProperty("Email3EntryID",value);
			}
		}

		public string FileAs
		{
			get
			{
				return (string)GetProperty("FileAs");
			}

			set
			{
				SetProperty("FileAs",value);
			}
		}

		public string FirstName
		{
			get
			{
				return (string)GetProperty("FirstName");
			}

			set
			{
				SetProperty("FirstName",value);
			}
		}

		public string FTPSite
		{
			get
			{
				return (string)GetProperty("FTPSite");
			}

			set
			{
				SetProperty("FTPSite",value);
			}
		}

		public string FullName
		{
			get
			{
				return (string)GetProperty("FullName");
			}

			set
			{
				SetProperty("FullName",value);
			}
		}

		public string FullNameAndCompany
		{
			get
			{
				return (string)GetProperty("FullNameAndCompany");
			}

			set
			{
				SetProperty("FullNameAndCompany",value);
			}
		}

		public RlOutlook.OlGender Gender
		{
			get
			{
				return (RlOutlook.OlGender)GetProperty("Gender");
			}

			set
			{
				SetProperty("Gender",value);
			}
		}

		public string GovernmentIDNumber
		{
			get
			{
				return (string)GetProperty("GovernmentIDNumber");
			}

			set
			{
				SetProperty("GovernmentIDNumber",value);
			}
		}

		public string Hobby
		{
			get
			{
				return (string)GetProperty("Hobby");
			}

			set
			{
				SetProperty("Hobby",value);
			}
		}

		public string Home2TelephoneNumber
		{
			get
			{
				return (string)GetProperty("Home2TelephoneNumber");
			}

			set
			{
				SetProperty("Home2TelephoneNumber",value);
			}
		}

		public string HomeAddress
		{
			get
			{
				return (string)GetProperty("HomeAddress");
			}

			set
			{
				SetProperty("HomeAddress",value);
			}
		}

		public string HomeAddressCity
		{
			get
			{
				return (string)GetProperty("HomeAddressCity");
			}

			set
			{
				SetProperty("HomeAddressCity",value);
			}
		}

		public string HomeAddressCountry
		{
			get
			{
				return (string)GetProperty("HomeAddressCountry");
			}

			set
			{
				SetProperty("HomeAddressCountry",value);
			}
		}

		public string HomeAddressPostalCode
		{
			get
			{
				return (string)GetProperty("HomeAddressPostalCode");
			}

			set
			{
				SetProperty("HomeAddressPostalCode",value);
			}
		}

		public string HomeAddressPostOfficeBox
		{
			get
			{
				return (string)GetProperty("HomeAddressPostOfficeBox");
			}

			set
			{
				SetProperty("HomeAddressPostOfficeBox",value);
			}
		}

		public string HomeAddressState
		{
			get
			{
				return (string)GetProperty("HomeAddressState");
			}

			set
			{
				SetProperty("HomeAddressState",value);
			}
		}

		public string HomeAddressStreet
		{
			get
			{
				return (string)GetProperty("HomeAddressStreet");
			}

			set
			{
				SetProperty("HomeAddressStreet",value);
			}
		}

		public string HomeFaxNumber
		{
			get
			{
				return (string)GetProperty("HomeFaxNumber");
			}

			set
			{
				SetProperty("HomeFaxNumber",value);
			}
		}

		public string HomeTelephoneNumber
		{
			get
			{
				return (string)GetProperty("HomeTelephoneNumber");
			}

			set
			{
				SetProperty("HomeTelephoneNumber",value);
			}
		}

		public string IMAddress
		{
			get
			{
				return (string)GetProperty("IMAddress");
			}

			set
			{
				SetProperty("IMAddress",value);
			}
		}

		public string Initials
		{
			get
			{
				return (string)GetProperty("Initials");
			}

			set
			{
				SetProperty("Initials",value);
			}
		}

		public string InternetFreeBusyAddress
		{
			get
			{
				return (string)GetProperty("InternetFreeBusyAddress");
			}

			set
			{
				SetProperty("InternetFreeBusyAddress",value);
			}
		}

		public string ISDNNumber
		{
			get
			{
				return (string)GetProperty("ISDNNumber");
			}

			set
			{
				SetProperty("ISDNNumber",value);
			}
		}

		public string JobTitle
		{
			get
			{
				return (string)GetProperty("JobTitle");
			}

			set
			{
				SetProperty("JobTitle",value);
			}
		}

		public bool Journal
		{
			get
			{
				return (bool)GetProperty("Journal");
			}

			set
			{
				SetProperty("Journal",value);
			}
		}

		public string Language
		{
			get
			{
				return (string)GetProperty("Language");
			}

			set
			{
				SetProperty("Language",value);
			}
		}

		public string LastFirstAndSuffix
		{
			get
			{
				return (string)GetProperty("LastFirstAndSuffix");
			}

			set
			{
				SetProperty("LastFirstAndSuffix",value);
			}
		}

		public string LastFirstNoSpace
		{
			get
			{
				return (string)GetProperty("LastFirstNoSpace");
			}

			set
			{
				SetProperty("LastFirstNoSpace",value);
			}
		}

		public string LastFirstNoSpaceAndSuffix
		{
			get
			{
				return (string)GetProperty("LastFirstNoSpaceAndSuffix");
			}

			set
			{
				SetProperty("LastFirstNoSpaceAndSuffix",value);
			}
		}

		public string LastFirstNoSpaceCompany
		{
			get
			{
				return (string)GetProperty("LastFirstNoSpaceCompany");
			}

			set
			{
				SetProperty("LastFirstNoSpaceCompany",value);
			}
		}

		public string LastFirstSpaceOnly
		{
			get
			{
				return (string)GetProperty("LastFirstSpaceOnly");
			}

			set
			{
				SetProperty("LastFirstSpaceOnly",value);
			}
		}

		public string LastFirstSpaceOnlyCompany
		{
			get
			{
				return (string)GetProperty("LastFirstSpaceOnlyCompany");
			}

			set
			{
				SetProperty("LastFirstSpaceOnlyCompany",value);
			}
		}

		public string LastName
		{
			get
			{
				return (string)GetProperty("LastName");
			}

			set
			{
				SetProperty("LastName",value);
			}
		}

		public string LastNameAndFirstName
		{
			get
			{
				return (string)GetProperty("LastNameAndFirstName");
			}

			set
			{
				SetProperty("LastNameAndFirstName",value);
			}
		}

		public string MailingAddress
		{
			get
			{
				return (string)GetProperty("MailingAddress");
			}

			set
			{
				SetProperty("MailingAddress",value);
			}
		}

		public string MailingAddressCity
		{
			get
			{
				return (string)GetProperty("MailingAddressCity");
			}

			set
			{
				SetProperty("MailingAddressCity",value);
			}
		}

		public string MailingAddressCountry
		{
			get
			{
				return (string)GetProperty("MailingAddressCountry");
			}

			set
			{
				SetProperty("MailingAddressCountry",value);
			}
		}

		public string MailingAddressPostalCode
		{
			get
			{
				return (string)GetProperty("MailingAddressPostalCode");
			}

			set
			{
				SetProperty("MailingAddressPostalCode",value);
			}
		}

		public string MailingAddressPostOfficeBox
		{
			get
			{
				return (string)GetProperty("MailingAddressPostOfficeBox");
			}

			set
			{
				SetProperty("MailingAddressPostOfficeBox",value);
			}
		}

		public string MailingAddressState
		{
			get
			{
				return (string)GetProperty("MailingAddressState");
			}

			set
			{
				SetProperty("MailingAddressState",value);
			}
		}

		public string MailingAddressStreet
		{
			get
			{
				return (string)GetProperty("MailingAddressStreet");
			}

			set
			{
				SetProperty("MailingAddressStreet",value);
			}
		}

		public string ManagerName
		{
			get
			{
				return (string)GetProperty("ManagerName");
			}

			set
			{
				SetProperty("ManagerName",value);
			}
		}

		public string MiddleName
		{
			get
			{
				return (string)GetProperty("MiddleName");
			}

			set
			{
				SetProperty("MiddleName",value);
			}
		}

		public string MobileTelephoneNumber
		{
			get
			{
				return (string)GetProperty("MobileTelephoneNumber");
			}

			set
			{
				SetProperty("MobileTelephoneNumber",value);
			}
		}

		public string NetMeetingAlias
		{
			get
			{
				return (string)GetProperty("NetMeetingAlias");
			}

			set
			{
				SetProperty("NetMeetingAlias",value);
			}
		}

		public string NickName
		{
			get
			{
				return (string)GetProperty("NickName");
			}

			set
			{
				SetProperty("NickName",value);
			}
		}

		public string OfficeLocation
		{
			get
			{
				return (string)GetProperty("OfficeLocation");
			}

			set
			{
				SetProperty("OfficeLocation",value);
			}
		}

		public string OrganizationalIDNumber
		{
			get
			{
				return (string)GetProperty("OrganizationalIDNumber");
			}

			set
			{
				SetProperty("OrganizationalIDNumber",value);
			}
		}

		public string OtherAddress
		{
			get
			{
				return (string)GetProperty("OtherAddress");
			}

			set
			{
				SetProperty("OtherAddress",value);
			}
		}

		public string OtherAddressCity
		{
			get
			{
				return (string)GetProperty("OtherAddressCity");
			}

			set
			{
				SetProperty("OtherAddressCity",value);
			}
		}

		public string OtherAddressCountry
		{
			get
			{
				return (string)GetProperty("OtherAddressCountry");
			}

			set
			{
				SetProperty("OtherAddressCountry",value);
			}
		}

		public string OtherAddressPostalCode
		{
			get
			{
				return (string)GetProperty("OtherAddressPostalCode");
			}

			set
			{
				SetProperty("OtherAddressPostalCode",value);
			}
		}

		public string OtherAddressPostOfficeBox
		{
			get
			{
				return (string)GetProperty("OtherAddressPostOfficeBox");
			}

			set
			{
				SetProperty("OtherAddressPostOfficeBox",value);
			}
		}

		public string OtherAddressState
		{
			get
			{
				return (string)GetProperty("OtherAddressState");
			}

			set
			{
				SetProperty("OtherAddressState",value);
			}
		}

		public string OtherAddressStreet
		{
			get
			{
				return (string)GetProperty("OtherAddressStreet");
			}

			set
			{
				SetProperty("OtherAddressStreet",value);
			}
		}

		public string OtherFaxNumber
		{
			get
			{
				return (string)GetProperty("OtherFaxNumber");
			}

			set
			{
				SetProperty("OtherFaxNumber",value);
			}
		}

		public string OtherTelephoneNumber
		{
			get
			{
				return (string)GetProperty("OtherTelephoneNumber");
			}

			set
			{
				SetProperty("OtherTelephoneNumber",value);
			}
		}

		public string PagerNumber
		{
			get
			{
				return (string)GetProperty("PagerNumber");
			}

			set
			{
				SetProperty("PagerNumber",value);
			}
		}

		public string PersonalHomePage
		{
			get
			{
				return (string)GetProperty("PersonalHomePage");
			}

			set
			{
				SetProperty("PersonalHomePage",value);
			}
		}

		public string PrimaryTelephoneNumber
		{
			get
			{
				return (string)GetProperty("PrimaryTelephoneNumber");
			}

			set
			{
				SetProperty("PrimaryTelephoneNumber",value);
			}
		}

		public string Profession
		{
			get
			{
				return (string)GetProperty("Profession");
			}

			set
			{
				SetProperty("Profession",value);
			}
		}

		public string RadioTelephoneNumber
		{
			get
			{
				return (string)GetProperty("RadioTelephoneNumber");
			}

			set
			{
				SetProperty("RadioTelephoneNumber",value);
			}
		}

		public string ReferredBy
		{
			get
			{
				return (string)GetProperty("ReferredBy");
			}

			set
			{
				SetProperty("ReferredBy",value);
			}
		}

		public RlOutlook.OlMailingAddress SelectedMailingAddress
		{
			get
			{
				return (RlOutlook.OlMailingAddress)GetProperty("SelectedMailingAddress");
			}

			set
			{
				SetProperty("SelectedMailingAddress",value);
			}
		}

		public string Spouse
		{
			get
			{
				return (string)GetProperty("Spouse");
			}

			set
			{
				SetProperty("Spouse",value);
			}
		}

		public string Suffix
		{
			get
			{
				return (string)GetProperty("Suffix");
			}

			set
			{
				SetProperty("Suffix",value);
			}
		}

		public string TelexNumber
		{
			get
			{
				return (string)GetProperty("TelexNumber");
			}

			set
			{
				SetProperty("TelexNumber",value);
			}
		}

		public string Title
		{
			get
			{
				return (string)GetProperty("Title");
			}

			set
			{
				SetProperty("Title",value);
			}
		}

		public string TTYTDDTelephoneNumber
		{
			get
			{
				return (string)GetProperty("TTYTDDTelephoneNumber");
			}

			set
			{
				SetProperty("TTYTDDTelephoneNumber",value);
			}
		}

		public string User1
		{
			get
			{
				return (string)GetProperty("User1");
			}

			set
			{
				SetProperty("User1",value);
			}
		}

		public string User2
		{
			get
			{
				return (string)GetProperty("User2");
			}

			set
			{
				SetProperty("User2",value);
			}
		}

		public string User3
		{
			get
			{
				return (string)GetProperty("User3");
			}

			set
			{
				SetProperty("User3",value);
			}
		}

		public string User4
		{
			get
			{
				return (string)GetProperty("User4");
			}

			set
			{
				SetProperty("User4",value);
			}
		}

		public string UserCertificate
		{
			get
			{
				return (string)GetProperty("UserCertificate");
			}

			set
			{
				SetProperty("UserCertificate",value);
			}
		}

		public string WebPage
		{
			get
			{
				return (string)GetProperty("WebPage");
			}

			set
			{
				SetProperty("WebPage",value);
			}
		}

		public string YomiCompanyName
		{
			get
			{
				return (string)GetProperty("YomiCompanyName");
			}

			set
			{
				SetProperty("YomiCompanyName",value);
			}
		}

		public string YomiFirstName
		{
			get
			{
				return (string)GetProperty("YomiFirstName");
			}

			set
			{
				SetProperty("YomiFirstName",value);
			}
		}

		public string YomiLastName
		{
			get
			{
				return (string)GetProperty("YomiLastName");
			}

			set
			{
				SetProperty("YomiLastName",value);
			}
		}

		public string NetMeetingServer
		{
			get
			{
				return (string)GetProperty("NetMeetingServer");
			}

			set
			{
				SetProperty("NetMeetingServer",value);
			}
		}

		public RlOutlook.MailItem ForwardAsVcard()
		{
			return (RlOutlook.MailItem)CallMethod("ForwardAsVcard");
		}
	}

	public class OutlookDistListItem: OutlookNonNoteItem
	{
		public OutlookDistListItem(object item): base(item)
		{
		}

		public override RlOutlook.OlItemType Type
		{
			get
			{
				return RlOutlook.OlItemType.olDistributionListItem;
			}
		}
	}

	public class OutlookDocumentItem: OutlookNonNoteItem, RlOutlook._DocumentItem
	{
		public OutlookDocumentItem(object item): base(item)
		{
		}

		public override RlOutlook.OlItemType Type
		{
			get
			{
				return RlOutlook.OlItemType.olMailItem;
			}
		}
	}

	public class OutlookJournalItem: OutlookNonNoteItem
	{
		public OutlookJournalItem(object item): base(item)
		{
		}

		public override RlOutlook.OlItemType Type
		{
			get
			{
				return RlOutlook.OlItemType.olJournalItem;
			}
		}

		public RlOutlook.Recipients Recipients
		{
			get
			{
				return (RlOutlook.Recipients)GetProperty("Recipients");
			}
		}

		public RlOutlook.MailItem Reply()
		{
			return (RlOutlook.MailItem)CallMethod("ReplyAll");
		}

		public RlOutlook.MailItem ReplyAll()
		{
			return (RlOutlook.MailItem)CallMethod("ReplyAll");
		}

		public RlOutlook.MailItem Forward()
		{
			return (RlOutlook.MailItem)CallMethod("Forward");
		}
	}

	public class OutlookMailItem: OutlookMailMeetingItem
	{
		public OutlookMailItem(object item): base(item)
		{
		}

		public override RlOutlook.OlItemType Type
		{
			get
			{
				return RlOutlook.OlItemType.olMailItem;
			}
		}

		public RlOutlook.MailItem Forward()
		{
			return (RlOutlook.MailItem)CallMethod("Forward");
		}
	}

	public class OutlookMeetingItem: OutlookMailMeetingItem
	{
		public OutlookMeetingItem(object item): base(item)
		{
		}

		public override RlOutlook.OlItemType Type
		{
			get
			{
				return RlOutlook.OlItemType.olAppointmentItem;
			}
		}

		public RlOutlook.MeetingItem Forward()
		{
			return (RlOutlook.MeetingItem)CallMethod("Forward");
		}
	}

	public class OutlookNoteItem: OutlookItem
	{
		public OutlookNoteItem(object item): base(item)
		{
		}

		public override RlOutlook.OlItemType Type
		{
			get
			{
				return RlOutlook.OlItemType.olNoteItem;
			}
		}

		public override string Subject
		{
			get
			{
				return base.Subject;
			}

			set
			{
			}
		}

	}

	public class OutlookPostItem: OutlookNonNoteItem
	{
		public OutlookPostItem(object item): base(item)
		{
		}

		public override RlOutlook.OlItemType Type
		{
			get
			{
				return RlOutlook.OlItemType.olPostItem;
			}
		}

		public DateTime ExpiryTime
		{
			get
			{
				return (DateTime)GetProperty("ExpiryTime");
			}

			set
			{
				SetProperty("ExpiryTime",value);
			}
		}

		public string SenderName
		{
			get
			{
				return (string)GetProperty("SenderName");
			}
		}

		public DateTime SentOn
		{
			get
			{
				return (DateTime)GetProperty("SentOn");
			}
		}

		public RlOutlook.MailItem Reply()
		{
			return (RlOutlook.MailItem)CallMethod("ReplyAll");
		}

		public RlOutlook.MailItem Forward()
		{
			return (RlOutlook.MailItem)CallMethod("Forward");
		}
	}

	public class OutlookRemoteItem: OutlookNonNoteItem
	{
		public OutlookRemoteItem(object item): base(item)
		{
		}

		public override RlOutlook.OlItemType Type
		{
			get
			{
				return RlOutlook.OlItemType.olMailItem;
			}
		}
	}

	public class OutlookReportItem: OutlookNonNoteItem, RlOutlook._ReportItem
	{
		public OutlookReportItem(object item): base(item)
		{
		}

		public override RlOutlook.OlItemType Type
		{
			get
			{
				return RlOutlook.OlItemType.olMailItem;
			}
		}
	}

	public class OutlookTaskItem: OutlookNonNoteItem
	{
		public OutlookTaskItem(object item): base(item)
		{
		}

		public override RlOutlook.OlItemType Type
		{
			get
			{
				return RlOutlook.OlItemType.olTaskItem;
			}
		}

		public DateTime ReminderTime
		{
			get
			{
				return (DateTime)GetProperty("ReminderTime");
			}

			set
			{
				SetProperty("ReminderTime",value);
			}
		}

		public bool ReminderSet
		{
			get
			{
				return (bool)GetProperty("ReminderSet");
			}

			set
			{
				SetProperty("ReminderSet",value);
			}
		}

		public RlOutlook.Recipients Recipients
		{
			get
			{
				return (RlOutlook.Recipients)GetProperty("Recipients");
			}
		}

		public void Send()
		{
			CallMethod("Send");
		}
	}

	public class OutlookTaskRequestItem: OutlookNonNoteItem, RlOutlook._TaskRequestItem
	{
		public OutlookTaskRequestItem(object item): base(item)
		{
		}

		public override RlOutlook.OlItemType Type
		{
			get
			{
				return RlOutlook.OlItemType.olTaskItem;
			}
		}

		public RlOutlook.TaskItem GetAssociatedTask(bool arg0)
		{
			return (RlOutlook.TaskItem)CallMethod("GetAssociatedTask",arg0);
		}
	}

	public class OutlookTaskRequestAcceptItem: OutlookNonNoteItem, RlOutlook._TaskRequestAcceptItem
	{
		public OutlookTaskRequestAcceptItem(object item): base(item)
		{
		}

		public override RlOutlook.OlItemType Type
		{
			get
			{
				return RlOutlook.OlItemType.olTaskItem;
			}
		}

		public RlOutlook.TaskItem GetAssociatedTask(bool arg0)
		{
			return (RlOutlook.TaskItem)CallMethod("GetAssociatedTask",arg0);
		}
	}

	public class OutlookTaskRequestDeclineItem: OutlookNonNoteItem, RlOutlook._TaskRequestDeclineItem
	{
		public OutlookTaskRequestDeclineItem(object item): base(item)
		{
		}

		public override RlOutlook.OlItemType Type
		{
			get
			{
				return RlOutlook.OlItemType.olTaskItem;
			}
		}

		public RlOutlook.TaskItem GetAssociatedTask(bool arg0)
		{
			return (RlOutlook.TaskItem)CallMethod("GetAssociatedTask",arg0);
		}
	}

	public class OutlookTaskRequestUpdateItem: OutlookNonNoteItem, RlOutlook._TaskRequestUpdateItem
	{
		public OutlookTaskRequestUpdateItem(object item): base(item)
		{
		}

		public override RlOutlook.OlItemType Type
		{
			get
			{
				return RlOutlook.OlItemType.olTaskItem;
			}
		}

		public RlOutlook.TaskItem GetAssociatedTask(bool arg0)
		{
			return (RlOutlook.TaskItem)CallMethod("GetAssociatedTask",arg0);
		}
	}
	#endregion

	public abstract class OutlookItem: IDisposable
	{
		protected object oItem;
#if (OL2002)
		private RlOutlook.ItemProperties itemProperties;
#endif
		internal static bool UseMAPI = true;
		protected static Guid CdoPropSetID1 = new Guid("0006200200000000C000000000000046");
		protected static Guid CdoPropSetID2 = new Guid("0006200300000000C000000000000046");
		protected static Guid CdoPropSetID3 = new Guid("0006200400000000C000000000000046");
		protected static Guid CdoPropSetID4 = new Guid("0006200800000000C000000000000046");
		protected static Guid CdoPropSetID5 = new Guid("0002032900000000C000000000000046");
		protected static Guid CdoPropSetID6 = new Guid("0006200E00000000C000000000000046");
		protected static Guid CdoPropSetID7 = new Guid("0006200A00000000C000000000000046");

		private IntPtr pUnk = IntPtr.Zero;
		private IMAPIProp propset = null;

		#region Constructors
		protected OutlookItem(object item)
		{
			if (item==null)
			{
				throw new ArgumentNullException("item");
			}
			this.oItem=item;
		}

		~OutlookItem()
		{
			Dispose();
		}

		public static OutlookItem CreateItem(object item)
		{
			if (item is RlOutlook.AppointmentItem)
			{
				return new OutlookAppointmentItem(item);
			}
			if (item is RlOutlook.ContactItem)
			{
				return new OutlookContactItem(item);
			}
			else if (item is RlOutlook.DistListItem)
			{
				return new OutlookDistListItem(item);
			}
			else if (item is RlOutlook.DocumentItem)
			{
				return new OutlookDocumentItem(item);
			}
			else if (item is RlOutlook.JournalItem)
			{
				return new OutlookJournalItem(item);
			}
			else if (item is RlOutlook.MailItem)
			{
				return new OutlookMailItem(item);
			}
			else if (item is RlOutlook.MeetingItem)
			{
				return new OutlookMeetingItem(item);
			}
			else if (item is RlOutlook.NoteItem)
			{
				return new OutlookNoteItem(item);
			}
			else if (item is RlOutlook.PostItem)
			{
				return new OutlookPostItem(item);
			}
			else if (item is RlOutlook.RemoteItem)
			{
				return new OutlookRemoteItem(item);
			}
			else if (item is RlOutlook.ReportItem)
			{
				return new OutlookReportItem(item);
			}
			else if (item is RlOutlook.TaskItem)
			{
				return new OutlookTaskItem(item);
			}
			else if (item is RlOutlook.TaskRequestItem)
			{
				return new OutlookTaskRequestItem(item);
			}
			else if (item is RlOutlook.TaskRequestAcceptItem)
			{
				return new OutlookTaskRequestAcceptItem(item);
			}
			else if (item is RlOutlook.TaskRequestDeclineItem)
			{
				return new OutlookTaskRequestDeclineItem(item);
			}
			else if (item is RlOutlook.TaskRequestUpdateItem)
			{
				return new OutlookTaskRequestUpdateItem(item);
			}
			else
			{
				throw new ArgumentException("Not a valid Outlook item");
			}
		}
		#endregion

		public virtual void Dispose()
		{
			if (oItem!=null)
			{
#if (COMRELEASE)
#if (OL2002)
				if (itemProperties!=null)
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(itemProperties);
				}
#endif
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem);
#endif
#if (OL2002)
				itemProperties=null;
#endif
				if (propset!=null)
				{
					System.Runtime.InteropServices.Marshal.Release(pUnk);
					propset.Dispose();
					propset=null;
				}
				oItem=null;
			}
			System.GC.SuppressFinalize(this);
		}

		public RlOutlook.ItemEvents_Event Events
		{
			get
			{
				return (RlOutlook.ItemEvents_Event)oItem;
			}
		}

		public object Item
		{
			get
			{
				return oItem;
			}
		}

		public abstract RlOutlook.OlItemType Type
		{
			get;
		}

		protected virtual MAPINAMEID GetMapiID(string property)
		{
			return null;
		}

		private IMAPIProp GetPropSet()
		{
			if (propset==null)
			{
				pUnk = System.Runtime.InteropServices.Marshal.GetIUnknownForObject(MAPIOBJECT);
				propset = (IMAPIProp)pUnk;
			}
			return propset;
		}

		protected virtual Tags GetMapiTag(string property, out bool found)
		{
			try
			{
				MAPINAMEID id = GetMapiID(property);
				if (id!=null)
				{
					GetPropSet();
					Tags[] tags;
					Error error = propset.GetIDsFromNames(new MAPINAMEID[] {id},IMAPIProp.FLAGS.Default,out tags);
					if ((error==Error.Success)&&(tags.Length==1))
					{
						found=true;
						return tags[0];
					}
				}
			}
			catch
			{
			}
			found=false;
			return Tags.ptagNull;
		}

		protected object GetProperty(string property)
		{
			if (UseMAPI)
			{
				try
				{
					bool found;
					Tags tag = GetMapiTag(property, out found);
					if (found)
					{
						GetPropSet();
						Value[] values;
						Error error = propset.GetProps(new Tags[] {tag}, IMAPIProp.FLAGS.Default, out values);
						if (values.Length==1)
						{
							if ((values[0] is MapiNull)||(values[0] is MapiUnspecified)||(values[0] is MapiError))
							{
								return null;
							}
							else
							{
								return values[0].GetType().InvokeMember("Value",
									BindingFlags.Public|BindingFlags.Instance|BindingFlags.GetField,
									null,values[0],null);
							}
						}
					}
				}
				catch
				{
				}
			}
			return oItem.GetType().InvokeMember(property,
				BindingFlags.Public|BindingFlags.GetField|BindingFlags.GetProperty,
				null,oItem,null);
		}

		protected void SetProperty(string property, object value)
		{
			oItem.GetType().InvokeMember(property,
				BindingFlags.Public|BindingFlags.SetField|BindingFlags.SetProperty,
				null,oItem,new object[] {value});
		}

		protected object CallMethod(string name, params object[] args)
		{
			return oItem.GetType().InvokeMember(name,
				BindingFlags.Public|BindingFlags.InvokeMethod,
				null,oItem,args);
		}

		public object MAPIOBJECT
		{
			get
			{
				return GetProperty("MAPIOBJECT");
			}
		}

		#region Properties
		public virtual RlOutlook.Application Application
		{
			get
			{
				return (RlOutlook.Application)GetProperty("Application");
			}
		}

#if (OL2003)
		public bool AutoResolvedWinner
		{
			get
			{
				return (bool)GetProperty("AutoResolvedWinner");
			}
		}
#endif

		public virtual string Body
		{
			get
			{
				return (string)GetProperty("Body");
			}

			set
			{
				SetProperty("Body",value);
			}
		}

		public virtual string Categories
		{
			get
			{
				return (string)GetProperty("Categories");
			}

			set
			{
				SetProperty("Categories",value);
			}
		}

		public virtual RlOutlook.OlObjectClass Class
		{
			get
			{
				return (RlOutlook.OlObjectClass)GetProperty("Class");
			}
		}

#if (OL2003)
		public RlOutlook.Conflicts Conflicts
		{
			get
			{
				return (RlOutlook.Conflicts)GetProperty("Conflicts");
			}
		}
#endif

		public virtual DateTime CreationTime
		{
			get
			{
				return (DateTime)GetProperty("CreationTime");
			}
		}

#if (OL2002)
		public virtual RlOutlook.OlDownloadState DownloadState
		{
			get
			{
				return (RlOutlook.OlDownloadState)GetProperty("DownloadState");
			}
		}
#endif

		public virtual string EntryID
		{
			get
			{
				return (string)GetProperty("EntryID");
			}
		}

		public virtual RlOutlook.Inspector GetInspector
		{
			get
			{
				return (RlOutlook.Inspector)GetProperty("GetInspector");
			}
		}

		public virtual bool IsConflict
		{
			get
			{
				return (bool)GetProperty("IsConflict");
			}
		}

#if (OL2002)
		public virtual RlOutlook.ItemProperties ItemProperties
		{
			get
			{
				if (itemProperties!=null)
				{
					itemProperties = (RlOutlook.ItemProperties)GetProperty("ItemProperties");
				}
				return itemProperties;
			}
		}
#endif

		public virtual DateTime LastModificationTime
		{
			get
			{
				return (DateTime)GetProperty("LastModificationTime");
			}
		}

		public virtual RlOutlook.Links Links
		{
			get
			{
				return (RlOutlook.Links)GetProperty("Links");
			}
		}

		public virtual RlOutlook.OlRemoteStatus MarkForDownload
		{
			get
			{
				return (RlOutlook.OlRemoteStatus)GetProperty("MarkForDownload");
			}

			set
			{
				SetProperty("MarkForDownload",value);
			}
		}

		public virtual string MessageClass
		{
			get
			{
				return (string)GetProperty("MessageClass");
			}

			set
			{
				SetProperty("MessageClass",value);
			}
		}

		public virtual object Parent
		{
			get
			{
				return GetProperty("Parent");
			}
		}

		public virtual bool Saved
		{
			get
			{
				return (bool)GetProperty("Saved");
			}
		}

		public virtual RlOutlook.NameSpace Session
		{
			get
			{
				return (RlOutlook.NameSpace)GetProperty("Session");
			}
		}

		public virtual int Size
		{
			get
			{
				return (int)GetProperty("Size");
			}
		}

		public virtual string Subject
		{
			get
			{
				return (string)GetProperty("Subject");
			}

			set
			{
				SetProperty("Subject",value);
			}
		}
		#endregion

		#region Methods
		public virtual void Close(RlOutlook.OlInspectorClose arg0)
		{
			CallMethod("Close",arg0);
		}

		public virtual object Copy()
		{
			return CallMethod("Close");
		}

		public virtual void Delete()
		{
			CallMethod("Close");
		}

		public virtual void Display(object arg0)
		{
			CallMethod("Close",arg0);
		}

		public virtual object Move(RlOutlook.MAPIFolder arg0)
		{
			return CallMethod("Close",arg0);
		}

		public virtual void PrintOut()
		{
			CallMethod("Close");
		}

		public virtual void Save()
		{
			CallMethod("Close");
		}

		public virtual void SaveAs(string arg0, object arg1)
		{
			CallMethod("Close",arg0,arg1);
		}
	#endregion
	}
}
