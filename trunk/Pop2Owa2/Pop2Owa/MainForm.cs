/*
 * Created by SharpDevelop.
 * User: Carlos
 * Date: 28/02/2010
 * Time: 13:19
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


using NLog;

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
	/// Description of MainForm.
	/// </summary>
	public partial class MainForm : Form
	{
		private static Logger logger = LogManager.GetCurrentClassLogger();
		internal static Settings pop2OwaSettings = new Settings();
		internal static bool AuthRequired;
		internal static bool ProxyConfigured= true;

		private POP3Listener objPOP3;
		private SMTPListener objSMTP;
		
		public MainForm()
		{
			try
			{
				logger.Trace("Calling InitializeComponent");
				InitializeComponent();
				logger.Trace("Calling LoadConfig");
				LoadConfig();			
				logger.Trace("Setting combos");	
				cboVersion.DataSource = System.Enum.GetValues(typeof(ExchangeVersion));
	            cboVersion.SelectedItem = ExchangeVersion.Exchange2007_SP1;
			}
			catch(Exception se)
			{
				logger.FatalException("Error loadding main form ", se);
			}

			try
			{
				logger.Trace("Setting sokects");
				objPOP3 = new POP3Listener(IPAddress.Parse(pop2OwaSettings.HostIP), pop2OwaSettings.Pop3Port);
				objSMTP = new SMTPListener(IPAddress.Parse(pop2OwaSettings.HostIP), pop2OwaSettings.SmtpPort);
			}
			catch(Exception se)
			{
				logger.FatalException("Error creating sockets ", se);
			}
		}

		void BntApplyClick(object sender, EventArgs e)
		{
			SaveConfig();
		}
		
		void LoadConfig()
		{
			try
			{
				// Get the isolated store for this assembly
				IsolatedStorageFile isf = IsolatedStorageFile.GetUserStoreForAssembly();
	
				// Open the settings file
				FileStream flStream = new FileStream("config.xml", FileMode.Open, FileAccess.Read);
    
				// Deserialize the XML to an object
				pop2OwaSettings = new Settings();
				SoapFormatter SF = new SoapFormatter();
				pop2OwaSettings = (Settings)SF.Deserialize(flStream);
				flStream.Close();
	
				// And apply the settings to the form
				txtServer.Text = pop2OwaSettings.Server;
				txtDomain.Text = pop2OwaSettings.Domain;
				chkSave.Checked= pop2OwaSettings.SaveOnSend;
				cboVersion.SelectedItem= pop2OwaSettings.ServerVersion;
				txtHostIP.Text=pop2OwaSettings.HostIP;
				txtPop3Port.Text= pop2OwaSettings.Pop3Port.ToString();
				txtSMTPPort.Text= pop2OwaSettings.SmtpPort.ToString();

				txtProxyDomain.Text = pop2OwaSettings.ProxyDomain;
				txtProxyServer.Text = pop2OwaSettings.ProxyServer;
				txtProxyUser.Text = pop2OwaSettings.ProxyUser;
				txtProxyPasword.Text = pop2OwaSettings.ProxyPassword;
			
			}catch(FileNotFoundException){
				logger.Warn("Config File not found");
			}catch(Exception ex){
				logger.WarnException("Error Loading Config", ex);	
			}
		}

		void SaveConfig()
		{

			// Create a settings object
			pop2OwaSettings = new Settings();
			pop2OwaSettings.Server = txtServer.Text;
			pop2OwaSettings.Domain = txtDomain.Text;
			pop2OwaSettings.ServerVersion= (ExchangeVersion) cboVersion.SelectedItem;
			pop2OwaSettings.SaveOnSend = chkSave.Checked;
			pop2OwaSettings.HostIP=txtHostIP.Text;
			pop2OwaSettings.Pop3Port= int.Parse(txtPop3Port.Text);
			pop2OwaSettings.SmtpPort= int.Parse(txtSMTPPort.Text);

			pop2OwaSettings.ProxyServer = txtProxyServer.Text;
			pop2OwaSettings.ProxyDomain = txtProxyDomain.Text;
			pop2OwaSettings.ProxyUser = txtProxyUser.Text;
			pop2OwaSettings.ProxyPassword = txtProxyPasword.Text;
			// Create or truncate the settings file
			// This will ensure that only the object we're
			// saving right now will be in the file
			FileStream flStream = new FileStream("config.xml", FileMode.Create , FileAccess.Write);

			// Serialize the object to the file
			SoapFormatter SF = new SoapFormatter();
			SF.Serialize(flStream, pop2OwaSettings);
			flStream.Close();
		}
		
		void BntResetClick(object sender, EventArgs e)
		{
			SaveConfig();
			objPOP3 = new POP3Listener(IPAddress.Parse(pop2OwaSettings.HostIP), pop2OwaSettings.Pop3Port);
			objSMTP = new SMTPListener(IPAddress.Parse(pop2OwaSettings.HostIP), pop2OwaSettings.SmtpPort);

		}
	}

}