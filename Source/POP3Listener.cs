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

namespace Pop2Owa
{
	public class POP3Listener : Listener
	    {
	    private string syncState= null;
		private const string OK = "+OK ";
		private const string ERR = "-ERR ";
		private const string EndMail = "\r\n.\r\n";
		private static byte[] byEndMail = System.Text.Encoding.ASCII.GetBytes(EndMail);
		
	    internal static Hashtable UidlCacheItems = new Hashtable();
	
		public POP3Listener(IPAddress address, int port): base(address,port){

	    }
	
	    protected override  void OnConnectionRequest(CSocket socket){
			socket.SendCRLF (OK);
	    }

	    protected override  void OnDataArrival(CSocket socket)
		{
	    	socket.Buffer = socket.Buffer.TrimEnd();
			//TODO Check implementation RSET command.
			//TODO Maybe we have a bug if the user has a password that start with space.
			string[] vData = socket.Buffer.ToString().Split(' ');
			string strDataToSend="";
			vData[0]=vData[0].ToUpper();
			if (vData[0]== "PASS"){
				logger.Debug(vData[0]);
			}else{
				logger.Debug(socket.Buffer);		
			}
			switch (vData[0]){
				case "USER":
					ObjEWS.User= vData[1];
					socket.SendCRLF (OK + "Password required for " + vData[1]);
					break;
				case "PASS":
					ObjEWS.Password=vData[1];
					switch (ObjEWS.SyncData(ref syncState) ){
						case 0:
							socket.SendCRLF (OK + "mailbox for " + ObjEWS.User + " ready");
							break;
						case 401:
							socket.SendCRLF (ERR + "Invalid Password for " + ObjEWS.User );
							break;
						default:
							socket.SendCRLF (ERR + "Unable to connect with mailbox");
							break;
					}
					break;
				case "STAT":
					socket.SendCRLF(OK + ObjEWS.TotalCount() + " " + ObjEWS.TotalSize() );
					break;
				case "RETR":
					goto case "TOP";
				case "TOP":
					int intMsg = int.Parse(vData[1]);
					socket.SendCRLF (OK + ObjEWS.GetMsgData(intMsg).Size.ToString() + " octets");
					socket.Send (ObjEWS.GetMsg(intMsg));
					socket.Send (byEndMail);
					break;
				case "QUIT":
					socket.SendCRLF (OK + " server signing off, 0 messages deleted");
					socket.Reset();
					ObjEWS= null;
					break;
				case "NOOP":
					socket.SendCRLF (OK);
					break;
				case "UIDL":
					if (socket.Buffer.Length > 5){
						socket.SendCRLF (OK + vData[1] + " " + ObjEWS.GetMsgData(int.Parse(vData[1])).Uid );
					}else{
						strDataToSend=OK + ObjEWS.TotalCount() + " messages ("  + ObjEWS.TotalSize() +  ") octets" + Environment.NewLine;
						
						for (int i = 0; i < (ObjEWS.TotalCount()) ; i++){
							strDataToSend += (i+1).ToString() + " " + ObjEWS.GetMsgData(i+1).Uid + Environment.NewLine;
						}
						socket.SendCRLF (strDataToSend + ".");
					}
					break;
				case "LIST":
					if (socket.Buffer.Length > 5){
						socket.SendCRLF (OK + vData[1] + " "+ ObjEWS.GetMsgData(Int32.Parse(vData[1])).Size.ToString());
					}else{
						socket.SendCRLF(OK + ObjEWS.TotalCount() + " messages ("  + ObjEWS.TotalSize() +  ") octets");
						for (intMsg = 0; intMsg<ObjEWS.TotalCount();intMsg++){
							socket.SendCRLF((intMsg + 1).ToString() + " "+ ObjEWS.GetMsgData(intMsg+1).Size.ToString());
						}
						socket.Send(byEndMail);
					}
					break;
	
				case "CAPA":
					socket.SendCRLF (OK + "Capability list follows");
					socket.SendCRLF ("USER");
					socket.SendCRLF ("UIDL");
					socket.Send(byEndMail);
					break;
				case "DELE":
					ObjEWS.DeleteMsg(Int32.Parse(vData[1]));
					socket.SendCRLF (OK + "message deleted");
					break;
					
					/*            intMsg = CInt(vData(1))
	            If objOWA.Delete(intMsg) Then
	                strDataToSend = OK & "message " & intMsg & " deleted"
	                //Additional Sleep, My Outlook 2000 fails if delete msg faster
	                MsgWaitObj 500
	            Else
	                strDataToSend = Error & "deleting message " & intMsg
	            End If*/
				default:
					socket.SendCRLF (ERR + "Syntax error");
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
