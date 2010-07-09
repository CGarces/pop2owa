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
using System.Runtime.CompilerServices;
using System.Runtime.Serialization.Formatters.Soap;

using Microsoft.Exchange.WebServices.Data;

[assembly: InternalsVisibleTo("Pop2Owa Test")]
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
		
		private static string m_ConfigFile;
		
		public  static string ConfigFile {
			get { 
				if (m_ConfigFile== null){
					m_ConfigFile = System.IO.Path.GetDirectoryName((new System.Uri(System.Reflection.Assembly.GetExecutingAssembly().CodeBase)).LocalPath);
					m_ConfigFile =Path.Combine(m_ConfigFile , "config.xml");
				}
				return m_ConfigFile; }
		}
		public static void ReadConfig(){
			// Open the settings file
			FileStream flStream = new FileStream(ConfigFile, FileMode.Open, FileAccess.Read);
			// Deserialize the XML to an object
			config = new Settings();
			SoapFormatter SF = new SoapFormatter();
			config = (Settings)SF.Deserialize(flStream);
			flStream.Close();
			SF=null;
			flStream=null;
		}
	}
}
