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
using System.Configuration.Install;
using System.ServiceProcess;

namespace Pop2Owa
{
	[RunInstaller(true)]
	public class ProjectInstaller : Installer
	{
		private ServiceProcessInstaller serviceProcessInstaller;
		private ServiceInstaller serviceInstaller;
		
		public ProjectInstaller()
		{
			serviceProcessInstaller = new ServiceProcessInstaller();
			serviceInstaller = new ServiceInstaller();
			// Here you can set properties on serviceProcessInstaller or register event handlers
			serviceProcessInstaller.Account = ServiceAccount.LocalService;
			serviceInstaller.ServiceName = NTService.MyServiceName;
			this.Installers.AddRange(new Installer[] { serviceProcessInstaller, serviceInstaller });
		}
	}
}
