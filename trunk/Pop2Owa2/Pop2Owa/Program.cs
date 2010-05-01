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
			Application.EnableVisualStyles();
			Application.SetCompatibleTextRenderingDefault(false);
			Application.Run(new MainForm());
		}
		
	}
}
