/*
 * Created by SharpDevelop.
 * User: Carlos
 * Date: 01/05/2010
 * Time: 16:35
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Net;
using System.ServiceProcess;

using NLog;

namespace Pop2Owa
{
	public class NTService : ServiceBase
	{
		private POP3Listener objPOP3;
		private SMTPListener objSMTP;
		public const string MyServiceName = "Pop2Owa";
		private static Logger logger = LogManager.GetCurrentClassLogger();
		

		public NTService(){
			logger.Trace("Creating class");
		}
		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose(bool disposing)
		{
			logger.Trace("Dispose");
			base.Dispose(disposing);
		}
		
		/// <summary>
		/// Start this service.
		/// </summary>
		protected override void OnStart(string[] args)
		{
			try
			{
				logger.Trace("Start");
				this.RequestAdditionalTime(120000);
				logger.Trace("RequestAdditionalTime");
 				AppSettings.ReadConfig();
				logger.Trace("Setting sokects");
				IPAddress serverIP = IPAddress.Parse(AppSettings.config.HostIP);
				objPOP3 = new POP3Listener(serverIP, AppSettings.config.Pop3Port);
				objSMTP = new SMTPListener(serverIP, AppSettings.config.SmtpPort);
				GC.Collect();
      			GC.WaitForPendingFinalizers();
			}
			catch(Exception)
			{
//				logger.FatalException("Error creating sockets ", se);
			}
		}
		
		/// <summary>
		/// Stop this service.
		/// </summary>
		protected override void OnStop()
		{
			//objPOP3= null;
			//objSMTP= null;
		}
	}
}
