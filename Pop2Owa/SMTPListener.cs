/*
 * Created by SharpDevelop.
 * User: Carlos
 * Date: 18/04/2010
 * Time: 19:38
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Net;
using System.Text;

namespace Pop2Owa
{
	public class SMTPListener : Listener
	    {
		private StringBuilder Maildata;
		
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
				
			switch (state){
				case State.LOGIN :
					ObjEWS.User= Decode(socket.Buffer);
					//"Password:" base64 encoded
					strDataToSend="334 UGFzc3dvcmQ6" + Environment.NewLine;
					state= State.PASSWORD;
					break;
				case State.PASSWORD:
					ObjEWS.Password= Decode(socket.Buffer);
					//TODO Reset the connection?
					strDataToSend =ValidateSMTPAUTH();
					state= State.INITIAL;
					break;
				case State.MAILDATA:
					if (socket.Buffer ==".\r\n") {
						if(ObjEWS.SendMsg(Maildata.ToString())){
							strDataToSend = CODEOK;
						}else{
							strDataToSend = "500 Syntax error, command unrecognized";
						}
						Maildata = null;
						state=State.INITIAL;
					}else{
						//Store Msg in the buffer
						Maildata.Append(socket.Buffer);						
					}
					break;
				default:
					if (socket.Buffer.Length >3){
						strSTMPCommmad= socket.Buffer.Substring(0, 4);
						if (strSTMPCommmad != "PASS" && state != State.MAILDATA) {
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
								state = State.LOGIN;
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
							state= State.INITIAL;
							Maildata = null;
							strDataToSend = CODEOK;
							break;
						case "DATA":
							strDataToSend = "354 Start mail input; end with <CRLF>.<CRLF>";
							state = State.MAILDATA;
							break;
						case "QUIT":
							//TODO Check is must be deleted on QUIT
							strDataToSend = "221 Service closing transmission channel" + Environment.NewLine;
							byData = System.Text.Encoding.ASCII.GetBytes(strDataToSend);
							socket.Send (byData);
							socket.Reset();
							strDataToSend = "";
							state=State.INITIAL;
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
}
