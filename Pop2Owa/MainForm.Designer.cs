/*
 * Created by SharpDevelop.
 * User: Carlos
 * Date: 28/02/2010
 * Time: 13:19
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
namespace Pop2Owa
{
	partial class MainForm
	{
		/// <summary>
		/// Designer variable used to keep track of non-visual components.
		/// </summary>
		private System.ComponentModel.IContainer components = null;
		
		/// <summary>
		/// Disposes resources used by the form.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing) {
				if (components != null) {
					components.Dispose();
				}
			}
			base.Dispose(disposing);
		}
		
		/// <summary>
		/// This method is required for Windows Forms designer support.
		/// Do not change the method contents inside the source code editor. The Forms designer might
		/// not be able to load this method if it was changed manually.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			this.bntApply = new System.Windows.Forms.Button();
			this.bntReset = new System.Windows.Forms.Button();
			this.tabControl1 = new System.Windows.Forms.TabControl();
			this.Exchange = new System.Windows.Forms.TabPage();
			this.cboVersion = new System.Windows.Forms.ComboBox();
			this.lblVersion = new System.Windows.Forms.Label();
			this.chkSave = new System.Windows.Forms.CheckBox();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.txtDomain = new System.Windows.Forms.TextBox();
			this.txtServer = new System.Windows.Forms.TextBox();
			this.Network = new System.Windows.Forms.TabPage();
			this.label5 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.txtSMTPPort = new System.Windows.Forms.TextBox();
			this.txtPop3Port = new System.Windows.Forms.TextBox();
			this.label3 = new System.Windows.Forms.Label();
			this.txtHostIP = new System.Windows.Forms.TextBox();
			this.tabProxy = new System.Windows.Forms.TabPage();
			this.label8 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.txtProxyPasword = new System.Windows.Forms.TextBox();
			this.txtProxyDomain = new System.Windows.Forms.TextBox();
			this.txtProxyServer = new System.Windows.Forms.TextBox();
			this.txtProxyUser = new System.Windows.Forms.TextBox();
			this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);
			this.tabControl1.SuspendLayout();
			this.Exchange.SuspendLayout();
			this.Network.SuspendLayout();
			this.tabProxy.SuspendLayout();
			this.SuspendLayout();
			// 
			// bntApply
			// 
			this.bntApply.Location = new System.Drawing.Point(226, 173);
			this.bntApply.Name = "bntApply";
			this.bntApply.Size = new System.Drawing.Size(75, 23);
			this.bntApply.TabIndex = 0;
			this.bntApply.Text = "Save";
			this.bntApply.UseVisualStyleBackColor = true;
			this.bntApply.Click += new System.EventHandler(this.BntApplyClick);
			// 
			// bntReset
			// 
			this.bntReset.Location = new System.Drawing.Point(307, 173);
			this.bntReset.Name = "bntReset";
			this.bntReset.Size = new System.Drawing.Size(75, 23);
			this.bntReset.TabIndex = 0;
			this.bntReset.Text = "Reset";
			this.bntReset.UseVisualStyleBackColor = true;
			this.bntReset.Click += new System.EventHandler(this.BntResetClick);
			// 
			// tabControl1
			// 
			this.tabControl1.Controls.Add(this.Exchange);
			this.tabControl1.Controls.Add(this.Network);
			this.tabControl1.Controls.Add(this.tabProxy);
			this.tabControl1.Location = new System.Drawing.Point(12, 12);
			this.tabControl1.Name = "tabControl1";
			this.tabControl1.SelectedIndex = 0;
			this.tabControl1.Size = new System.Drawing.Size(370, 155);
			this.tabControl1.TabIndex = 30;
			// 
			// Exchange
			// 
			this.Exchange.Controls.Add(this.cboVersion);
			this.Exchange.Controls.Add(this.lblVersion);
			this.Exchange.Controls.Add(this.chkSave);
			this.Exchange.Controls.Add(this.label2);
			this.Exchange.Controls.Add(this.label1);
			this.Exchange.Controls.Add(this.txtDomain);
			this.Exchange.Controls.Add(this.txtServer);
			this.Exchange.Location = new System.Drawing.Point(4, 22);
			this.Exchange.Name = "Exchange";
			this.Exchange.Padding = new System.Windows.Forms.Padding(3);
			this.Exchange.Size = new System.Drawing.Size(362, 129);
			this.Exchange.TabIndex = 0;
			this.Exchange.Text = "Exchange Server";
			this.Exchange.UseVisualStyleBackColor = true;
			// 
			// cboVersion
			// 
			this.cboVersion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.cboVersion.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cboVersion.FormattingEnabled = true;
			this.cboVersion.Location = new System.Drawing.Point(122, 58);
			this.cboVersion.Name = "cboVersion";
			this.cboVersion.Size = new System.Drawing.Size(232, 21);
			this.cboVersion.TabIndex = 35;
			// 
			// lblVersion
			// 
			this.lblVersion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblVersion.Location = new System.Drawing.Point(12, 61);
			this.lblVersion.Name = "lblVersion";
			this.lblVersion.Size = new System.Drawing.Size(63, 21);
			this.lblVersion.TabIndex = 36;
			this.lblVersion.Text = "Version:";
			// 
			// chkSave
			// 
			this.chkSave.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.chkSave.Location = new System.Drawing.Point(12, 85);
			this.chkSave.Name = "chkSave";
			this.chkSave.Size = new System.Drawing.Size(124, 24);
			this.chkSave.TabIndex = 34;
			this.chkSave.Text = "Save on Send";
			this.chkSave.UseVisualStyleBackColor = true;
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(12, 35);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(100, 17);
			this.label2.TabIndex = 33;
			this.label2.Text = "Domain:";
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(12, 9);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(100, 17);
			this.label1.TabIndex = 32;
			this.label1.Text = "Exchange URL";
			// 
			// txtDomain
			// 
			this.txtDomain.Location = new System.Drawing.Point(122, 32);
			this.txtDomain.Name = "txtDomain";
			this.txtDomain.Size = new System.Drawing.Size(232, 20);
			this.txtDomain.TabIndex = 30;
			// 
			// txtServer
			// 
			this.txtServer.Location = new System.Drawing.Point(122, 6);
			this.txtServer.Name = "txtServer";
			this.txtServer.Size = new System.Drawing.Size(232, 20);
			this.txtServer.TabIndex = 31;
			// 
			// Network
			// 
			this.Network.Controls.Add(this.label5);
			this.Network.Controls.Add(this.label4);
			this.Network.Controls.Add(this.txtSMTPPort);
			this.Network.Controls.Add(this.txtPop3Port);
			this.Network.Controls.Add(this.label3);
			this.Network.Controls.Add(this.txtHostIP);
			this.Network.Location = new System.Drawing.Point(4, 22);
			this.Network.Name = "Network";
			this.Network.Padding = new System.Windows.Forms.Padding(3);
			this.Network.Size = new System.Drawing.Size(362, 129);
			this.Network.TabIndex = 1;
			this.Network.Text = "Network Configuration";
			this.Network.UseVisualStyleBackColor = true;
			// 
			// label5
			// 
			this.label5.Location = new System.Drawing.Point(6, 65);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(100, 17);
			this.label5.TabIndex = 35;
			this.label5.Text = "SMTP Port";
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(6, 39);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(100, 17);
			this.label4.TabIndex = 35;
			this.label4.Text = "POP3 Port";
			// 
			// txtSMTPPort
			// 
			this.txtSMTPPort.Location = new System.Drawing.Point(112, 62);
			this.txtSMTPPort.Name = "txtSMTPPort";
			this.txtSMTPPort.Size = new System.Drawing.Size(232, 20);
			this.txtSMTPPort.TabIndex = 34;
			this.txtSMTPPort.Text = "25";
			this.txtSMTPPort.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			// 
			// txtPop3Port
			// 
			this.txtPop3Port.Location = new System.Drawing.Point(112, 36);
			this.txtPop3Port.Name = "txtPop3Port";
			this.txtPop3Port.Size = new System.Drawing.Size(232, 20);
			this.txtPop3Port.TabIndex = 34;
			this.txtPop3Port.Text = "110";
			this.txtPop3Port.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(6, 13);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(100, 17);
			this.label3.TabIndex = 9;
			this.label3.Text = "Email Server IP:";
			// 
			// txtHostIP
			// 
			this.txtHostIP.Location = new System.Drawing.Point(112, 10);
			this.txtHostIP.Name = "txtHostIP";
			this.txtHostIP.Size = new System.Drawing.Size(232, 20);
			this.txtHostIP.TabIndex = 5;
			this.txtHostIP.Text = "127.0.0.1";
			this.txtHostIP.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			// 
			// tabProxy
			// 
			this.tabProxy.Controls.Add(this.label8);
			this.tabProxy.Controls.Add(this.label6);
			this.tabProxy.Controls.Add(this.label9);
			this.tabProxy.Controls.Add(this.label7);
			this.tabProxy.Controls.Add(this.txtProxyPasword);
			this.tabProxy.Controls.Add(this.txtProxyDomain);
			this.tabProxy.Controls.Add(this.txtProxyServer);
			this.tabProxy.Controls.Add(this.txtProxyUser);
			this.tabProxy.Location = new System.Drawing.Point(4, 22);
			this.tabProxy.Name = "tabProxy";
			this.tabProxy.Size = new System.Drawing.Size(362, 129);
			this.tabProxy.TabIndex = 2;
			this.tabProxy.Text = "Proxy";
			this.tabProxy.UseVisualStyleBackColor = true;
			// 
			// label8
			// 
			this.label8.Location = new System.Drawing.Point(17, 100);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(100, 17);
			this.label8.TabIndex = 37;
			this.label8.Text = "Password";
			// 
			// label6
			// 
			this.label6.Location = new System.Drawing.Point(17, 74);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(100, 17);
			this.label6.TabIndex = 37;
			this.label6.Text = "Domain:";
			// 
			// label9
			// 
			this.label9.Location = new System.Drawing.Point(17, 22);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(100, 17);
			this.label9.TabIndex = 36;
			this.label9.Text = "Proxy Server";
			// 
			// label7
			// 
			this.label7.Location = new System.Drawing.Point(17, 48);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(100, 17);
			this.label7.TabIndex = 36;
			this.label7.Text = "User";
			// 
			// txtProxyPasword
			// 
			this.txtProxyPasword.Location = new System.Drawing.Point(127, 97);
			this.txtProxyPasword.Name = "txtProxyPasword";
			this.txtProxyPasword.PasswordChar = '*';
			this.txtProxyPasword.Size = new System.Drawing.Size(232, 20);
			this.txtProxyPasword.TabIndex = 34;
			// 
			// txtProxyDomain
			// 
			this.txtProxyDomain.Location = new System.Drawing.Point(127, 71);
			this.txtProxyDomain.Name = "txtProxyDomain";
			this.txtProxyDomain.Size = new System.Drawing.Size(232, 20);
			this.txtProxyDomain.TabIndex = 34;
			// 
			// txtProxyServer
			// 
			this.txtProxyServer.Location = new System.Drawing.Point(127, 19);
			this.txtProxyServer.Name = "txtProxyServer";
			this.txtProxyServer.Size = new System.Drawing.Size(232, 20);
			this.txtProxyServer.TabIndex = 35;
			// 
			// txtProxyUser
			// 
			this.txtProxyUser.Location = new System.Drawing.Point(127, 45);
			this.txtProxyUser.Name = "txtProxyUser";
			this.txtProxyUser.Size = new System.Drawing.Size(232, 20);
			this.txtProxyUser.TabIndex = 35;
			// 
			// notifyIcon1
			// 
			this.notifyIcon1.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info;
			this.notifyIcon1.Text = "notifyIcon1";
			this.notifyIcon1.DoubleClick += new System.EventHandler(this.NotifyIcon1DoubleClick);
			// 
			// MainForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(396, 205);
			this.Controls.Add(this.tabControl1);
			this.Controls.Add(this.bntReset);
			this.Controls.Add(this.bntApply);
			this.Name = "MainForm";
			this.Text = "Pop2Owa";
			this.Resize += new System.EventHandler(this.MainFormResize);
			this.tabControl1.ResumeLayout(false);
			this.Exchange.ResumeLayout(false);
			this.Exchange.PerformLayout();
			this.Network.ResumeLayout(false);
			this.Network.PerformLayout();
			this.tabProxy.ResumeLayout(false);
			this.tabProxy.PerformLayout();
			this.ResumeLayout(false);
		}
		private System.Windows.Forms.NotifyIcon notifyIcon1;
		private System.Windows.Forms.TextBox txtProxyServer;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.TextBox txtProxyDomain;
		private System.Windows.Forms.TextBox txtProxyUser;
		private System.Windows.Forms.TextBox txtProxyPasword;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.TabPage tabProxy;
		private System.Windows.Forms.TextBox txtHostIP;
		private System.Windows.Forms.TextBox txtSMTPPort;
		private System.Windows.Forms.TextBox txtPop3Port;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.TabPage Network;
		private System.Windows.Forms.CheckBox chkSave;
		private System.Windows.Forms.TabPage Exchange;
		private System.Windows.Forms.TabControl tabControl1;
		private System.Windows.Forms.Label lblVersion;
		private System.Windows.Forms.ComboBox cboVersion;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.TextBox txtDomain;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Button bntReset;
		private System.Windows.Forms.TextBox txtServer;
		private System.Windows.Forms.Button bntApply;
		
	}
}
