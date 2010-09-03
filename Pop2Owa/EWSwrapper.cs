/*
 * Created by SharpDevelop.
 * User: Carlos
 * Date: 03/03/2010
 * Time: 21:11
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections;
using System.Collections.Generic;
using System.Net;
using System.Security.Cryptography;
using System.Text;

using Microsoft.Exchange.WebServices.Data;
using NLog;

namespace Pop2Owa

{
	/// <summary>
	/// Description of EWSWrapper.
	/// </summary>
	public class EWSWrapper
	{
		private static Logger logger = LogManager.GetCurrentClassLogger();
		private	long lngSize;
		private	int intCount;

		private ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);

		public ExchangeSettings EWSSettings;
		
		public string User;
		public string Password;
		private static string EndMail = Environment.NewLine+'.'+Environment.NewLine;
		private static string ReplacedEndMail = Environment.NewLine+".."+Environment.NewLine;
		internal Hashtable emails = new Hashtable();
		internal struct message{
			public string Id;
			public string Uid;
			public long Size;
			public message(string MessageId, string MessageUid, long MessageSize){
				Id=MessageId;
				Uid=MessageUid;
				Size=MessageSize;
			}
				
		}
		private string[] messages;

		public EWSWrapper(ExchangeSettings ServerSettings)
		{
			EWSSettings= ServerSettings;
			if (String.IsNullOrEmpty(EWSSettings.Server)){
				throw new System.ArgumentException("Server cannot be null", "EWSSettings.Server");
			}
		}
		public int SyncData(ref string syncState){
		
		int intReturn;
		const int pageSize = 512;
		try
	      {
	    	logger.Debug("Attemp to conect to {0}", EWSSettings.Server );

			service =Connection();

			PropertySet myPropertySet = new PropertySet(BasePropertySet.IdOnly);
			myPropertySet.Add(ItemSchema.Size);

			FolderId myFolderId= new FolderId(WellKnownFolderName.Inbox);

            ChangeCollection<ItemChange> changeCollection;
			lngSize =0;

			logger.Debug("Start findResults");
			do
			{
		    	logger.Trace("looping");
				changeCollection = service.SyncFolderItems(myFolderId,
                                                               myPropertySet, null, pageSize,
                                                            SyncFolderItemsScope.NormalItems, syncState);
				foreach (var change in changeCollection) {
					//Item newitem = Item.Bind(service, item.Id.ToString(), new PropertySet(BasePropertySet.IdOnly, new List<PropertyDefinitionBase>() { EmailMessageSchema.MimeContent }));
					if (change.ChangeType == ChangeType.Create&& ! POP3Listener.UidlCacheItems.ContainsKey(change.ItemId.ToString())){
						POP3Listener.UidlCacheItems.Add(change.Item.Id.ToString(), new message(change.Item.Id.ToString(), CalculateMD5Hash(change.Item.Id.ToString()) , change.Item.Size));
					}else if (change.ChangeType == ChangeType.Delete && POP3Listener.UidlCacheItems.ContainsKey(change.ItemId.ToString())){
						POP3Listener.UidlCacheItems.Remove(change.ItemId.ToString());
					}
				}
				//view.Offset += 50;
				syncState = changeCollection.SyncState;
			} while (changeCollection.MoreChangesAvailable);
			intReturn=0;
		    logger.Debug("End findResults");
			Array.Resize(ref messages, POP3Listener.UidlCacheItems.Count);
			intCount =0;
			foreach( DictionaryEntry ItemEntry in POP3Listener.UidlCacheItems)
	        {
				messages[intCount] = ItemEntry.Key.ToString();
	            intCount++;
	            lngSize+= ((message) ItemEntry.Value).Size;
	        }
			
	    }catch (ServiceRequestException e){
	    	intReturn=1;
	    	if (e.InnerException is WebException){
	    		WebException webexception = (WebException) e.InnerException;
	    		if (webexception.Response!=null ) {
	    			intReturn=(int) ((HttpWebResponse) webexception.Response).StatusCode;
		    		if (intReturn.Equals(HttpStatusCode.ProxyAuthenticationRequired)  && ! AppSettings.AuthRequired){
		    			logger.Warn("Proxy Authetication Required");
			    		AppSettings.AuthRequired = true; 
		    			intReturn=SyncData(ref syncState);
		    		} else if (intReturn.Equals(HttpStatusCode.Unauthorized)) {
		    			logger.Warn("Unauthorized, check your password");
		    		} else {
			    		logger.FatalException("Error conecting to the server", webexception);
		    		}
	    		}else{
	    			switch (webexception.Status){
    				case WebExceptionStatus.ConnectFailure:
		    			logger.Fatal("Error conecting to the server, check your internet connection");
		    			break;
		    		default:
		    			logger.FatalException("Error connecting to the server", webexception);
		    			break;
	    			}
	    		} 		
	    	} else {
	    		logger.ErrorException("Error conecting to the server", e);
	    	}
	    }
		return intReturn;
		}

		public byte[] GetMsg(long lngMessage){
			string strMsg;
			byte[] MimeString= Item.Bind(service, GetMsgData(lngMessage).Id, new PropertySet(BasePropertySet.IdOnly, new List<PropertyDefinitionBase>() { EmailMessageSchema.MimeContent })).MimeContent.Content;  

			strMsg = System.Text.ASCIIEncoding.ASCII.GetString(MimeString);
			if (strMsg.IndexOf(EndMail)>1){
				return System.Text.Encoding.ASCII.GetBytes(strMsg.Replace(EndMail, ReplacedEndMail));
			} else {
				return MimeString;
			}
		}
		public bool SendMsg(string msg){
			try {
				EmailMessage message = new EmailMessage(Connection());
				message.MimeContent = new MimeContent();
				message.ItemClass= "IPM.Note";

				//TODO: Check conversion between stringbuilder & byte[].
				message.MimeContent.Content=System.Text.Encoding.ASCII.GetBytes(msg);

				foreach( DictionaryEntry ItemEntry in emails)
		        {
					message.BccRecipients.Add(ItemEntry.Key.ToString());
				}

				if (EWSSettings.SaveOnSend){
					message.SendAndSaveCopy();
				} else {
					message.Send();			
				}
				
				return true;
			} catch (Exception e) {
				logger.WarnException("Exception sending mail", e);
				return false;
			}
		}
		
		public bool DeleteMsg(long lngMessage){
			service.DeleteItems(new[] {new ItemId(GetMsgData(lngMessage).Id)}, DeleteMode.HardDelete, null, null);
			return true;
		}
		/// <summary>
		/// Calculate MD5 Hast used as POP3 UID string
		/// </summary>
		/// <param name="input">EWS message id</param>
		/// <returns></returns>
		private static string CalculateMD5Hash(string input)
		{
		    // step 1, calculate MD5 hash from input
		    MD5 md5 = System.Security.Cryptography.MD5.Create();
		    byte[] inputBytes = System.Text.Encoding.ASCII.GetBytes(input);
		    byte[] hash = md5.ComputeHash(inputBytes);
		 
		    // step 2, convert byte array to hex string
		    StringBuilder sb = new StringBuilder();
		    for (int i = 0; i < hash.Length; i++)
		    {
		        sb.Append(hash[i].ToString("X2"));
		    }
		    return sb.ToString();
		}
		
		
		internal  message GetMsgData(long lngMessage){
			return (message) POP3Listener.UidlCacheItems[messages[lngMessage-1]];
		}
		
		public long TotalSize(){
			return lngSize;
			
		}
		public int TotalCount(){
			return intCount;
		}
		
		private ExchangeService Connection()
		{
			ExchangeService service = new ExchangeService(EWSSettings.ServerVersion);
			service.Credentials = new NetworkCredential(User, Password, EWSSettings.Domain);
			service.Url = new Uri(EWSSettings.Server);

			if (String.IsNullOrEmpty(EWSSettings.ProxyServer)){
				if (AppSettings.AuthRequired & !AppSettings.ProxyConfigured){
				    // Try default credentials (e.g. for ISA with NTLM integration)
				    WebRequest.DefaultWebProxy.Credentials = CredentialCache.DefaultCredentials;
					AppSettings.ProxyConfigured= true;
				}
			}else{
				WebRequest.DefaultWebProxy = new WebProxy(EWSSettings.ProxyServer,true);
				WebRequest.DefaultWebProxy.Credentials = new NetworkCredential(EWSSettings.ProxyUser, EWSSettings.ProxyPassword, EWSSettings.ProxyDomain );					
				AppSettings.ProxyConfigured= true;
			}
			return service;
		}

		
		/// <summary>
        /// Attempt to made a dummy operation to check if we have conectivity.
        /// </summary>
        public bool TestExchangeService()
        {
            try
            {
				service =Connection();
            	string DummyEntryId = "ABCDEABCDEABCDEABCDEABCDEABCDEABCDEABCDEABCDEABCDEABCDEABCDEABCDEABCDEABCDEABCDEABCDEABCDEABCDEABCDEABCDEABCDEABCDEABCDEABCDEABCDEABCDEABCDE";
                service.ConvertIds(new AlternateId[] { new AlternateId(IdFormat.HexEntryId, DummyEntryId, "pop2owa@pop2owa.com") }, IdFormat.HexEntryId);
				return true;
            }
            catch (ServiceRequestException)
            {
            	return false;
            }
        }
        ~EWSWrapper()
        {
        	logger.Trace("EWSWrapper destroyed");
        }


		
	}
}
