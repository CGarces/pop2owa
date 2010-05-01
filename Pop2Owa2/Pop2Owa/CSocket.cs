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

        internal string Buffer="";
		
		private static Logger logger = LogManager.GetCurrentClassLogger();
		private Socket m_socListener;
		private AsyncCallback pfnWorkerCallBack ;
		private Socket m_socWorker;
		private bool disposed;
		
		private class CSocketPacket
		{
			public System.Net.Sockets.Socket thisSocket;
			public byte[] dataBuffer = new byte[1];
		}		
		public CSocket(IPAddress address, int port) {
			//create the listening socket...
			m_socListener = new Socket(AddressFamily.InterNetwork,SocketType.Stream,ProtocolType.Tcp);
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
			//m_socWorker.Shutdown(SocketShutdown.Both);
			if (m_socWorker.Connected){
				m_socWorker.Close();
			}
			//start listening...
			m_socListener.Listen (4);
			// create the call back for any client connections...
			m_socListener.BeginAccept(new AsyncCallback (OnClientConnect),null);
		}

		public void Send(byte[] byData){
			m_socWorker.Send (byData);
		}
		public void Send(string data){
			byte[] byData = System.Text.Encoding.ASCII.GetBytes(data + Environment.NewLine);
			m_socWorker.Send (byData);
		}
			
		private void OnClientConnect(IAsyncResult asyn)
		{
			try
			{
				m_socWorker = m_socListener.EndAccept (asyn);
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
				soc .BeginReceive (theSocPkt.dataBuffer ,0,theSocPkt.dataBuffer.Length ,SocketFlags.None,pfnWorkerCallBack,theSocPkt);
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
				System.String szData = new System.String(chars);
				
				Buffer= Buffer + szData.Trim('\0');
				if (szData.Equals("\n\0") ){
					DataArrival(this);
					Buffer= "";
				}
				WaitForData(m_socWorker);
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
					logger.Trace("Dispose Executed");
		            m_socListener.Shutdown(SocketShutdown.Both);
		            m_socListener.Close();
		            m_socWorker.Shutdown(SocketShutdown.Both);
		            m_socWorker.Close();
		            disposed = true;
	            }catch (Exception ex){
	            	logger.TraceException("Error Disposing", ex);
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
