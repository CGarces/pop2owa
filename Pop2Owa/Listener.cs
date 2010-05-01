/*
 * Created by SharpDevelop.
 * User: Carlos
 * Date: 18/04/2010
 * Time: 19:38
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections;
using System.Net;
using System.Text;

using NLog;

namespace Pop2Owa
{
	/// <summary>
	/// Description of Listener.
	/// </summary>
	public abstract class Listener
	{
    	private EWSWrapper m_ObjEWS;
		protected static readonly string ENDMAIL = "." + Environment.NewLine;
		protected static Logger logger = LogManager.GetCurrentClassLogger();
    	private CSocket socket;
		public EWSWrapper ObjEWS {
			get { 
    			if (m_ObjEWS==null){
					m_ObjEWS= new EWSWrapper(MainForm.pop2OwaSettings);    				
    			} 
    			return m_ObjEWS; 
    		}
			set{m_ObjEWS=value;}
		}
    	protected Listener(IPAddress address, int port){
    		socket= new CSocket(address,port);
            socket.DataArrival += new CSocket.DataArrivalEventHandler(OnDataArrival);
            socket.ConnectionRequest += new CSocket.ConnectionRequestEventHandler(OnConnectionRequest);
            socket.ConnectionClosed += new CSocket.ConnectionClosedEventHandler(OnConnectionClosed);
        }
    	
    	protected abstract void OnDataArrival(CSocket socket);
        protected abstract void OnConnectionRequest(CSocket socket);
    	protected void OnConnectionClosed(CSocket socket){
    		socket.Reset();
    	}
	}


	public class SMTPListener : Listener
    {
		private StringBuilder Maildata;
		private string m_SMTPState="";
		
		
		public SMTPListener(IPAddress address, int port): base(address,port){
			
		}
		protected override void OnConnectionRequest(CSocket socket){
			socket.Send ("220 Simple Mail Transfer Service Ready");
    	}
		protected override void OnDataArrival(CSocket socket)
		{
			const string CODEOK  = "250 OK";
			string strSTMPCommmad;
		
			//strBuffer = strBuffer.TrimEnd();
			byte[] byData = null;
			string strDataToSend="";
				
			switch (m_SMTPState){
				case "LOGIN":
					ObjEWS.User= Decode(socket.Buffer);
					//"Password:" base64 encoded
					strDataToSend="334 UGFzc3dvcmQ6" + Environment.NewLine;
					m_SMTPState = "PASSWORD";
					break;
				case "PASSWORD":
					ObjEWS.Password= Decode(socket.Buffer);
					strDataToSend =ValidateSMTPAUTH();
					m_SMTPState = null;
					break;
				case "DATA":
					//Store Msg in the buffer
					Maildata.Append(socket.Buffer);
					if (Maildata.Length >5) {
						char[] endData= new char[5];
						Maildata.CopyTo(Maildata.Length -5, endData,0,5);
						if (new String(endData).ToString() ==("\r\n.\r\n")) {
							Maildata.Remove(Maildata.Length -5, 5);
							if(ObjEWS.SendMsg(Maildata.ToString())){
								strDataToSend = CODEOK;
							}else{
								strDataToSend = "500 Syntax error, command unrecognized";
							}
							Maildata = null;
							m_SMTPState = null;
						}
					}
					break;
				default:
					if (socket.Buffer.Length >3){
						strSTMPCommmad= socket.Buffer.Substring(0, 4);
						if (strSTMPCommmad != "PASS" && m_SMTPState != "DATA") {
							logger.Debug(socket.Buffer);
						}
					} else {
						strSTMPCommmad="";
					}
					switch (strSTMPCommmad){
						case "":
							strDataToSend = "220 Simple Mail Transfer Service Ready";
							break;
						case "HELO":
							strDataToSend = "250 ";
							break;
						case "EHLO":
							strDataToSend = "250-" + Environment.NewLine +
								"250-AUTH LOGIN PLAIN" + Environment.NewLine +
								"250 HELP";
							break;
						case "AUTH":
							if (socket.Buffer == "AUTH LOGIN" +Environment.NewLine){
								//"Username:" as base64 encoded
								strDataToSend = "334 VXNlcm5hbWU6";
								m_SMTPState = "LOGIN";
							} else {
								if (socket.Buffer.Substring(0, 10) == "AUTH PLAIN"){
									if (socket.Buffer.Length== 12){
										strDataToSend = "534 Authentication mechanism is too weak";
									}else{
										socket.Buffer  = Decode(socket.Buffer.Substring(11)) ;
										if (socket.Buffer.Length  > 1){
											ObjEWS.User = socket.Buffer.Substring(1, socket.Buffer.IndexOf("\0",1)-1);
											ObjEWS.Password = socket.Buffer.Substring(socket.Buffer.IndexOf("\0",1)+1);
											strDataToSend = ValidateSMTPAUTH();
										}else{
											strDataToSend = "534 Authentication mechanism is too weak";
										}
									}
								} else {
									strDataToSend = "504 Unrecognized authentication type.";
								}
							}
							break;
						case "MAIL":
							if (ObjEWS.Password == null ||  ObjEWS.User == null){
								strDataToSend = "530 Authentication required";
							}else{
								Maildata= new StringBuilder();
								strDataToSend = CODEOK;
							}
							break;
						case "RCPT":
							strDataToSend = CODEOK;
							break;
						case "RSET":
							m_SMTPState = "";
							Maildata = null;
							strDataToSend = CODEOK;
							break;
						case "DATA":
							strDataToSend = "354 Start mail input; end with <CRLF>.<CRLF>";
							m_SMTPState = "DATA";
							break;
						case "QUIT":
							strDataToSend = "221 Service closing transmission channel" + Environment.NewLine;
							byData = System.Text.Encoding.ASCII.GetBytes(strDataToSend);
							socket.Send (byData);
							socket.Reset();
							strDataToSend = "";
							m_SMTPState = "";
							Maildata = null;
							ObjEWS= null;
							break;
						default:
							logger.Error("Unknow command {0}", socket.Buffer);
							break;
					}
					break;
			}
			if (strDataToSend.Length>0){
				byData = System.Text.Encoding.ASCII.GetBytes(strDataToSend + Environment.NewLine);
				socket.Send (byData);
			}
		}
		private string ValidateSMTPAUTH(){
			string strReturn;
			if (ObjEWS.TestExchangeService()){
				strReturn="235 AUTHENTICATION SUCCESSFUL";
			}else{
				strReturn="535 5.7.0 Authentication failed";
			}
			return strReturn;
		}

		/// <summary>
		/// Base64 decode a string using the standard key
		/// </summary>
		/// <param name="EncodedString">The string to be decoded</param>
		/// <returns></returns>
		public static string Decode(string EncodedString)
		{
			// pad string
			EncodedString=EncodedString.TrimEnd();
			int padLength = EncodedString.Length + (EncodedString.Length % 4);
			EncodedString = EncodedString.PadRight(padLength, '=');
	
			System.Text.ASCIIEncoding enc = new System.Text.ASCIIEncoding();
			return enc.GetString(Convert.FromBase64String(EncodedString));
		}

	}

	public class POP3Listener : Listener
    {
    	private string syncState= null;
		private const string OK = "+OK ";
		private const string ERR = "-ERR ";
		
    	internal static Hashtable UidlCacheItems = new Hashtable();

		public POP3Listener(IPAddress address, int port): base(address,port){
	
		}

    	protected override  void OnConnectionRequest(CSocket socket){
			socket.Send (OK);
    	}
		protected override  void OnDataArrival(CSocket socket)
		{
			socket.Buffer = socket.Buffer.TrimEnd();
			string[] vData = socket.Buffer.Split(' ');
			byte[] byData = null;
			string strDataToSend="";
			logger.Debug(vData[0]);
			switch (vData[0]){
				case "USER":
					ObjEWS.User= vData[1];
					socket.Send (OK + "Password required for " + vData[0]);
					break;
				case "PASS":
					ObjEWS.Password=vData[1];
					switch (ObjEWS.SyncData(ref syncState) ){
						case 0:
							socket.Send (OK + "mailbox for " + ObjEWS.User + " ready");
							break;
						case 401:
							socket.Send (ERR + "Invalid Password for " + ObjEWS.User );
							break;
						default:
							socket.Send (ERR + "Unable to connect with mailbox");
							break;
					}
					break;
				case "STAT":
					socket.Send (OK + ObjEWS.TotalCount() + " " + ObjEWS.TotalSize() );
					break;
				case "RETR":
					goto case "TOP";
				case "TOP":
					int intMsg = int.Parse(vData[1]);
					socket.Send (OK + " " + ObjEWS.GetMsgData(intMsg).Size.ToString() + " octets");
					byData = ObjEWS.GetMsg(intMsg);
					socket.Send (byData);
					byData = System.Text.Encoding.ASCII.GetBytes(Environment.NewLine + "." + Environment.NewLine);
					socket.Send (byData);
					break;
				case "QUIT":
					strDataToSend = OK + " server signing off, 0 messages deleted" + Environment.NewLine;
					byData = System.Text.Encoding.ASCII.GetBytes(strDataToSend);
					socket.Send (byData);
					socket.Reset();
					ObjEWS= null;
					break;
				case "NOOP":
					socket.Send (OK);
					break;
				case "UIDL":
					if (socket.Buffer.Length > 5){
						strDataToSend=OK + vData[1] + " " + ObjEWS.GetMsgData(int.Parse(vData[1])).Uid  + Environment.NewLine;
					}else{
						strDataToSend=OK + ObjEWS.TotalCount() + " messages ("  + ObjEWS.TotalSize() +  ") octets" + Environment.NewLine;
						
						for (int i = 0; i < (ObjEWS.TotalCount()) ; i++){
							strDataToSend += (i+1).ToString() + " " + ObjEWS.GetMsgData(i+1).Uid + Environment.NewLine;
						}
						byData = System.Text.Encoding.ASCII.GetBytes(strDataToSend + "." +Environment.NewLine );
						socket.Send (byData);
					}
					break;
				case "LIST":
					if (socket.Buffer.Length > 5){
						socket.Send (OK + vData[1] + " "+ ObjEWS.GetMsgData(Int32.Parse(vData[1])).Size.ToString());
					}else{
						socket.Send (OK + ObjEWS.TotalCount() + " messages ("  + ObjEWS.TotalSize() +  ") octets");
						for (intMsg = 0; intMsg<ObjEWS.TotalCount();intMsg++){
							socket.Send((intMsg + 1).ToString() + " "+ ObjEWS.GetMsgData(intMsg+1).Size.ToString());
						}
						byData =  System.Text.Encoding.ASCII.GetBytes(Environment.NewLine + "." +Environment.NewLine);
						socket.Send (byData);
					}
					break;

				case "CAPA":
					byData =  System.Text.Encoding.ASCII.GetBytes(OK + "Capability list follows" + Environment.NewLine
					                                              + "USER" + Environment.NewLine
					                                              + "UIDL" + Environment.NewLine
					                                              + ENDMAIL);
					socket.Send (byData);
					
					break;
				case "DELE":
					socket.Send (OK + "message deleted");
					break;
					
					/*            intMsg = CInt(vData(1))
	            If objOWA.Delete(intMsg) Then
	                strDataToSend = OK & "message " & intMsg & " deleted"
	                //Additional Sleep, My Outlook 2000 fails if delete msg faster
	                MsgWaitObj 500
	            Else
	                strDataToSend = Error & "deleting message " & intMsg
	            End If*/
				case "AUTH":
					socket.Send (ERR + "unsupported feature");
					break;
				default:
					socket.Send (ERR + "Syntax error");
					break;
						/*           Debug.Assert False
        WriteLog "Unknown data: " & strDataRecived, Warning
        strDataToSend = Error & "Syntax error"
    End Select
    If LenB(strDataToSend) > 0 And skPOP3.State = sckConnected Then skPOP3.SendData strDataToSend
    Set oElement = Nothing
    Set oElements = Nothing
End If
						 */
			}
		}

	}
}
