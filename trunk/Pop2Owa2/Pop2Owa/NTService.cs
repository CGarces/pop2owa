/*
 * Created by SharpDevelop.
 * User: Carlos
 * Date: 01/05/2010
 * Time: 16:35
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.ServiceProcess;
using System.Text;
using System.Runtime.Serialization.Formatters.Soap;
using NLog;
using System.IO;
using System.IO.IsolatedStorage;
using System.Net;
using System.Net.Sockets;


namespace Pop2Owa
{
	public class NTService : ServiceBase
	{
		private POP3Listener objPOP3;
		private SMTPListener objSMTP;
		public const string MyServiceName = "Pop2Owa";
		private static Logger logger = LogManager.GetCurrentClassLogger();
		
		public NTService()
		{
			InitializeComponent();
		}
		
		private void InitializeComponent()
		{
			try
			{
				logger.Trace("Load Settings");
				AppSettings.ReadConfig();
			}
			catch(Exception se)
			{
				logger.FatalException("Loadding Settings ", se);
			}

		}
		
		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose(bool disposing)
		{
			base.Dispose(disposing);
		}
		
		/// <summary>
		/// Start this service.
		/// </summary>
		protected override void OnStart(string[] args)
		{
			try
			{
				logger.Trace("Setting sokects");
				objPOP3 = new POP3Listener(IPAddress.Parse(AppSettings.config.HostIP), AppSettings.config.Pop3Port);
				objSMTP = new SMTPListener(IPAddress.Parse(AppSettings.config.HostIP), AppSettings.config.SmtpPort);
			}
			catch(Exception se)
			{
				logger.FatalException("Error creating sockets ", se);
			}
		}
		
		/// <summary>
		/// Stop this service.
		/// </summary>
		protected override void OnStop()
		{
			objPOP3= null;
			objSMTP= null;
		}
	}
}
