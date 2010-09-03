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
using NLog;

namespace Pop2Owa
{
	/// <summary>
	/// Description of Listener.
	/// </summary>
	public abstract class Listener
	{
    	private EWSWrapper m_ObjEWS;
		protected static Logger logger = LogManager.GetCurrentClassLogger();
    	private CSocket socket;
    	protected enum State {
        	INITIAL, LOGIN, USER, PASSWORD, AUTHENTICATED, STARTMAIL, RECIPIENT, MAILDATA
    	};
	    // connection state
	    protected State state = State.INITIAL;

	    public int BufferSize {
			set{socket.BufferSize=value;}
	    }
	    
		public EWSWrapper ObjEWS {
			get { 
    			if (m_ObjEWS==null){
					m_ObjEWS= new EWSWrapper(AppSettings.config);    				
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
}
