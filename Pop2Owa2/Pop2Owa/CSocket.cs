/*
 * Created by SharpDevelop.
 * User: Carlos
 * Date: 18/04/2010
 * Time: 10:43
 * 
 */
using System;
using System.Net;
using System.Net.Sockets;
using System.Text;
using NLog;

namespace Pop2Owa
{
	/// <summary>
	/// Description of CSocket.
	/// </summary>
	public class CSocket: IDisposable
	{
		internal event DataArrivalEventHandler DataArrival;
		internal event ConnectionRequestEventHandler ConnectionRequest; 
		internal event ConnectionClosedEventHandler ConnectionClosed; 
        internal delegate void DataArrivalEventHandler(CSocket sender);
        internal delegate void ConnectionRequestEventHandler(CSocket sender);
        internal delegate void ConnectionClosedEventHandler(CSocket sender);

        internal StringBuilder internalBuffer= new StringBuilder();
		internal string Buffer="";
        
		private static Logger logger = LogManager.GetCurrentClassLogger();
		private Socket m_socListener;
		private AsyncCallback pfnWorkerCallBack ;
		private Socket m_socWorker;
		private bool disposed;

		private char[] EndLine = {'\n', '\0'};
	    private int m_BufferSize;
		public int BufferSize {
			set{m_BufferSize=value;}
	    }
		
		private class CSocketPacket
		{
			public System.Net.Sockets.Socket thisSocket;
			public byte[] dataBuffer = new byte[65536];
		}		
		public CSocket(IPAddress address, int port) {
			m_BufferSize=65536;
			//create the listening socket...
			m_socListener = new Socket(AddressFamily.InterNetwork,SocketType.Stream,ProtocolType.Tcp);
			m_socListener.SendBufferSize=m_BufferSize;
			m_socListener.ReceiveBufferSize=m_BufferSize;
			m_socListener.SetSocketOption(SocketOptionLevel.Socket,SocketOptionName.ReuseAddress, true);
			IPEndPoint ipLocal = new IPEndPoint (address, port);
			//bind to local IP Address...
			m_socListener.Bind( ipLocal );
			//start listening...
			m_socListener.Listen (4);
			// create the call back for any client connections...
			m_socListener.BeginAccept(new AsyncCallback ( OnClientConnect),null);
		}

		public void Reset(){
			internalBuffer= new StringBuilder();
			if (m_socWorker.Connected){
				m_socWorker.Shutdown(SocketShutdown.Both);
				m_socWorker.Close();
			}
			if (m_socListener.Connected){
				m_socListener.Shutdown(SocketShutdown.Both);
				m_socListener.Close();
			}
			//start listening...
			m_socListener.Listen (4);
			// create the call back for any client connections...
			m_socListener.BeginAccept(new AsyncCallback (OnClientConnect),null);
		}

		public void Send(byte[] byData){
			m_socWorker.Send (byData);
		}
		public void SendCRLF(string data){
			byte[] byData = System.Text.Encoding.ASCII.GetBytes(string.Concat(data, Environment.NewLine));
			m_socWorker.Send (byData);
		}
		public void Send(string data){
			byte[] byData = System.Text.Encoding.ASCII.GetBytes(data);
			m_socWorker.Send (byData);
		}
			
		private void OnClientConnect(IAsyncResult asyn)
		{
			try
			{
				m_socWorker = m_socListener.EndAccept (asyn);
				m_socWorker.SendBufferSize=m_BufferSize;
				m_socWorker.ReceiveBufferSize=m_BufferSize;
				ConnectionRequest(this);
				WaitForData(m_socWorker);
			}
			catch(ObjectDisposedException)
			{
				logger.Info("Socket has been closed");
				ConnectionClosed(this);
			}
			catch(SocketException se)
			{
				logger.FatalException("Socket error", se);
				ConnectionClosed(this);
			}
			
		}

		private void WaitForData(System.Net.Sockets.Socket soc)
		{
			try
			{
				if  ( pfnWorkerCallBack == null )
				{
					pfnWorkerCallBack = new AsyncCallback (OnDataReceived);
				}
				CSocketPacket theSocPkt = new CSocketPacket ();
				theSocPkt.thisSocket = soc;
				// now start to listen for any data...
				soc.BeginReceive (theSocPkt.dataBuffer ,0,theSocPkt.dataBuffer.Length ,SocketFlags.None,pfnWorkerCallBack,theSocPkt);
			}
			catch(SocketException se)
			{
				logger.FatalException("WaitForData exception", se);
				ConnectionClosed(this);
			}

		}
		
		private void OnDataReceived(IAsyncResult asyn)
		{
			try
			{
				CSocketPacket theSockId = (CSocketPacket)asyn.AsyncState ;
				//end receive...
				int iRx  = 0 ;
				iRx = theSockId.thisSocket.EndReceive (asyn);
				char[] chars = new char[iRx +  1];
				System.Text.Decoder d = System.Text.Encoding.UTF8.GetDecoder();
				d.GetChars(theSockId.dataBuffer, 0, iRx, chars, 0);

				internalBuffer.Append(chars, 0, iRx);
				logger.Trace("Socket {0} {1}", m_socListener.LocalEndPoint , iRx );
				if (iRx==0){
					logger.Trace("Close requested");
					internalBuffer= new StringBuilder();
					ConnectionClosed(this);
				} else {
					if (iRx>0 && chars[iRx-1]=='\n' && chars[iRx]=='\0'){
						logger.Trace("End msg defected");
						this.Buffer = internalBuffer.ToString();
						DataArrival(this);
						internalBuffer= new StringBuilder();
					}
					
					WaitForData(m_socWorker);
				}
			}
			catch (ObjectDisposedException )
			{
				logger.Info("Socket has been closed");
				ConnectionClosed(this);
			}
			catch(SocketException se)
			{
				logger.FatalException("Socket error", se);
				ConnectionClosed(this);
			}
		}
		
		#region IDisposable Code

		protected virtual void Dispose(bool disposing)
      	{
	        if (!disposed)
	        {
	            if (disposing)
	            {
	                // Free other state (managed objects).
	            }
	            // Free your own state (unmanaged objects).
	            try{
					if (m_socListener.Connected){
			            m_socListener.Shutdown(SocketShutdown.Both);
					}
			        m_socListener.Close();
	            }catch (Exception ex){
	            	logger.TraceException("Error Disposing", ex);
	            }finally{
	            	internalBuffer=null;	
		            disposed = true;
	            }
	        }
		}

		public void Dispose()
		{
		 Dispose(true);
		 GC.SuppressFinalize(this);
		}
		// Use C# destructor syntax for finalization code.
	    ~CSocket()
	    {
	        Dispose (false);
	    }
	    #endregion

	}
}
