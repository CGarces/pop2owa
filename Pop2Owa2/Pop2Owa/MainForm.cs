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
using System.Net;
using System.Runtime.Serialization.Formatters.Soap;
using System.Windows.Forms;

using Microsoft.Exchange.WebServices.Data;
using NLog;

namespace Pop2Owa

{

	/// <summary>
	/// Description of MainForm.
	/// </summary>
	public partial class MainForm : Form
	{
		private static Logger logger = LogManager.GetCurrentClassLogger();
		private POP3Listener objPOP3;
		private SMTPListener objSMTP;
		
		public MainForm()
		{
			try
			{
				logger.Trace("Calling InitializeComponent");
				InitializeComponent();
				logger.Trace("Setting combos");	
				cboVersion.DataSource = System.Enum.GetValues(typeof(ExchangeVersion));
	            cboVersion.SelectedItem = ExchangeVersion.Exchange2007_SP1;
				logger.Trace("Calling LoadConfig");
				LoadConfig();
				System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
				this.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyIcon1.Icon")));

			}
			catch(Exception se)
			{
				logger.FatalException("Error loadding main form ", se);
			}

			try
			{
				logger.Trace("Setting sokects");
				IPAddress serverIP = IPAddress.Parse(AppSettings.config.HostIP);
				objPOP3 = new POP3Listener(serverIP, AppSettings.config.Pop3Port);
				objSMTP = new SMTPListener(serverIP, AppSettings.config.SmtpPort);
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

				AppSettings.ReadConfig();
				// And apply the settings to the form
				txtServer.Text = AppSettings.config.Server;
				txtDomain.Text = AppSettings.config.Domain;
				chkSave.Checked= AppSettings.config.SaveOnSend;
				cboVersion.SelectedItem= AppSettings.config.ServerVersion;
				txtHostIP.Text=AppSettings.config.HostIP;
				txtPop3Port.Text= AppSettings.config.Pop3Port.ToString();
				txtSMTPPort.Text= AppSettings.config.SmtpPort.ToString();

				txtProxyDomain.Text = AppSettings.config.ProxyDomain;
				txtProxyServer.Text = AppSettings.config.ProxyServer;
				txtProxyUser.Text = AppSettings.config.ProxyUser;
				txtProxyPasword.Text = AppSettings.config.ProxyPassword;
			
			}catch(FileNotFoundException){
				logger.Warn("Config File not found");
			}catch(Exception ex){
				logger.WarnException("Error Loading Config", ex);	
			}
		}

		void SaveConfig()
		{

			// Create a settings object
			AppSettings.config = new Settings();
			AppSettings.config.Server = txtServer.Text;
			AppSettings.config.Domain = txtDomain.Text;
			AppSettings.config.ServerVersion= (ExchangeVersion) cboVersion.SelectedItem;
			AppSettings.config.SaveOnSend = chkSave.Checked;
			AppSettings.config.HostIP=txtHostIP.Text;
			AppSettings.config.Pop3Port= int.Parse(txtPop3Port.Text);
			AppSettings.config.SmtpPort= int.Parse(txtSMTPPort.Text);

			AppSettings.config.ProxyServer = txtProxyServer.Text;
			AppSettings.config.ProxyDomain = txtProxyDomain.Text;
			AppSettings.config.ProxyUser = txtProxyUser.Text;
			AppSettings.config.ProxyPassword = txtProxyPasword.Text;
			// Create or truncate the settings file
			// This will ensure that only the object we're
			// saving right now will be in the file
			FileStream flStream = new FileStream(AppSettings.ConfigFile, FileMode.Create , FileAccess.Write);

			// Serialize the object to the file
			SoapFormatter SF = new SoapFormatter();
			SF.Serialize(flStream, AppSettings.config);
			flStream.Close();
		}
		
		void BntResetClick(object sender, EventArgs e)
		{
			SaveConfig();
			IPAddress serverIP = IPAddress.Parse(AppSettings.config.HostIP);
			objPOP3 = new POP3Listener(serverIP, AppSettings.config.Pop3Port);
			objSMTP = new SMTPListener(serverIP, AppSettings.config.SmtpPort);

		}
		
		void MainFormResize(object sender, EventArgs e)
		{
			if (this.WindowState == FormWindowState.Minimized)
       		{
	             notifyIcon1.Visible = true;
	             notifyIcon1.BalloonTipText = "Pop2Owa minimized to tray";
	             notifyIcon1.ShowBalloonTip(1);  //show balloon tip for 2 seconds
	             notifyIcon1.Text = "Pop2Owa";
	             this.WindowState = FormWindowState.Minimized;
	             this.ShowInTaskbar = false;
	             Hide();
       		}

		}
		
		void NotifyIcon1DoubleClick(object sender, EventArgs e)
		{
			Show();
			notifyIcon1.Visible = false;
    		WindowState = FormWindowState.Normal;

		}
	}

}