/*
 * Created by SharpDevelop.
 * User: Carlos
 * Date: 28/02/2010
 * Time: 13:19
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Windows.Forms;
using System.ServiceProcess;
namespace Pop2Owa
{
	/// <summary>
	/// Class with program entry point.
	/// </summary>
	internal sealed class Program
	{
		/// <summary>
		/// Program entry point.
		/// </summary>
		[STAThread]
		private static void Main(string[] args)
		{
			if (args !=null && args.Length >0 && args[0]=="-nt"){
				// To run more than one service you have to add them here
				ServiceBase.Run(new ServiceBase[] { new NTService() });
			} else {
				Application.EnableVisualStyles();
				Application.SetCompatibleTextRenderingDefault(false);
				Application.Run(new MainForm());
			}
		}
		
	}
}
