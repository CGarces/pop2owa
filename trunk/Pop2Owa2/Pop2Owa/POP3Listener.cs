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
		
	    internal static Hashtable UidlCacheItems = new Hashtable();
	
		public POP3Listener(IPAddress address, int port): base(address,port){
	
		}
	
	    protected override  void OnConnectionRequest(CSocket socket){
			socket.Send (OK);
	    }

	    protected override  void OnDataArrival(CSocket socket)
		{
			socket.Buffer = socket.Buffer.TrimEnd();
			//TODO Check implementation RSET command.
			//TODO Maybe we have a bug if the user has a password that start with space.
			string[] vData = socket.Buffer.Split(' ');
			byte[] byData = null;
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
					socket.Send (OK + "Password required for " + vData[1]);
					break;
				case "PASS":
					ObjEWS.Password=vData[1];
					switch (ObjEWS.SyncData(ref syncState) ){
						case 0:
							socket.Send (OK + "mailbox for " + ObjEWS.User + " ready");
							break;
						case 401:
							socket.Send (ERR + "Invalid Password for " + ObjEWS.User );
							socket.Reset();
							break;
						default:
							socket.Send (ERR + "Unable to connect with mailbox");
							socket.Reset();
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
						socket.Send (OK + vData[1] + " " + ObjEWS.GetMsgData(int.Parse(vData[1])).Uid );
					}else{
						strDataToSend=OK + ObjEWS.TotalCount() + " messages ("  + ObjEWS.TotalSize() +  ") octets" + Environment.NewLine;
						
						for (int i = 0; i < (ObjEWS.TotalCount()) ; i++){
							strDataToSend += (i+1).ToString() + " " + ObjEWS.GetMsgData(i+1).Uid + Environment.NewLine;
						}
						socket.Send (strDataToSend + ".");
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
						socket.Send (Environment.NewLine + ".");
					}
					break;
	
				case "CAPA":
					socket.Send (OK + "Capability list follows");
					socket.Send ("USER");
					socket.Send ("UIDL");
					socket.Send (".");
					break;
				case "DELE":
					ObjEWS.DeleteMsg(Int32.Parse(vData[1]));
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
