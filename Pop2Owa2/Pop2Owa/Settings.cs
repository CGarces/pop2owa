/*
 * Created by SharpDevelop.
 * User: Carlos
 * Date: 01/05/2010
 * Time: 20:01
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.IO;
using System.IO.IsolatedStorage;
using System.Net;
using System.Net.Sockets;
using System.Runtime.Serialization.Formatters.Soap;
using System.Windows.Forms;
using Microsoft.Exchange.WebServices.Data;

namespace Pop2Owa
{
	
	
	[Serializable()]
	public class ExchangeSettings
	{
		public string Server;
		public string Domain;
		public bool SaveOnSend;
		public ExchangeVersion ServerVersion;
		public string ProxyServer;
		public string ProxyUser;
		public string ProxyPassword;
		public string ProxyDomain;
	}
	
	/// <summary>
	/// Class to hold persisted settings
	/// </summary>
	[Serializable()]
	public class Settings: ExchangeSettings
	{
		public string HostIP;
		public int Pop3Port;
		public int SmtpPort;
	}

	/// <summary>
	/// Description of Class1.
	/// </summary>
	public static class AppSettings
	{
		internal static Settings config = new Settings();
		internal static bool AuthRequired;
		internal static bool ProxyConfigured= true;
		
		public static void ReadConfig(){
				// Get the isolated store for this assembly
				IsolatedStorageFile isf = IsolatedStorageFile.GetUserStoreForAssembly();
	
				// Open the settings file
				string path = System.IO.Path.GetDirectoryName ((new System.Uri(System.Reflection.Assembly.GetExecutingAssembly().CodeBase)).AbsolutePath);
				FileStream flStream = new FileStream(path + "\\config.xml", FileMode.Open, FileAccess.Read);
    
				// Deserialize the XML to an object
				config = new Settings();
				SoapFormatter SF = new SoapFormatter();
				config = (Settings)SF.Deserialize(flStream);
				flStream.Close();
		}
	}
}
