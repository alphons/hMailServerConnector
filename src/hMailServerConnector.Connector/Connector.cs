using System;
using System.Diagnostics;
using System.Runtime.InteropServices;


// Copy ConnectorEventHandlers.vbs to EventHandlers.vbs in hMailserver/events directory
// Run _install.cmd script once, it does a 'regsvr32.exe hMailServerConnector.comhost.dll'
// Debugging: run VisualStudio as Admin (Starting/Stopping hmailserver in pre/post build)
// .. attach debugger Ctrl-Alt-P to hmailserver.exe (show processes for all users)
// For starting debugger automatically use Debugger.Launch()
// After first attach, you can use Shift-Alt-P for re-attach process for debugging


// https://www.hmailserver.com/documentation/latest/?page=scripting_onsmtpdata

namespace hMailServer
{
	[ComVisible(true)]
	[Guid("8945AE74-44E6-4833-901F-385EA68010E0")]
	[ProgId("hMailServer.Connector")]
	public class Connector : IConnector
	{
		public Connector()
		{
			Debugger.Launch(); // Remove this for production!!
		}

		public void OnClientConnect(object oClient, object oResult)
		{
			var Client = Marshal.CreateWrapperOfType<object, ClientClass>(oClient);
			var Result = Marshal.CreateWrapperOfType<object, ResultClass>(oResult);

			Debug.WriteLine($"OnClientConnect Helo:{Client.HELO} Ip:{Client.IPAddress} Port:{Client.Port} Username:{Client.Username}");

			Result.Value = 0;
			Result.Message = string.Empty;
		}

		public void OnSMTPData(object oClient, object oMessage, object oResult)
		{

			var Client = Marshal.CreateWrapperOfType<object, ClientClass>(oClient);
			var Message = Marshal.CreateWrapperOfType<object, MessageClass>(oMessage);
			var Result = Marshal.CreateWrapperOfType<object, ResultClass>(oResult);

			Debug.WriteLine($"OnSMTPData Ip:{Client.IPAddress} From:{Message.FromAddress} To:{Message.Recipients[0].Address}");

			// Some code for testing
			if (Message.FromAddress.Contains('1'))
			{
				Result.Value = 1;
				Result.Message = "Rejected 1"; // Message is not used
			}
			if (Message.FromAddress.Contains('2'))
			{
				Result.Value = 2;
				Result.Message = "Error 2";
			}
			if (Message.FromAddress.Contains('3'))
			{
				Result.Value = 3;
				Result.Message = "Error 3";
			}
		}

		public void OnAcceptMessage(object oClient, object oMessage, object oResult)
		{
			var Client = Marshal.CreateWrapperOfType<object, ClientClass>(oClient);
			var Message = Marshal.CreateWrapperOfType<object, MessageClass>(oMessage);
			var Result = Marshal.CreateWrapperOfType<object, ResultClass>(oResult);

			Debug.WriteLine($"OnAcceptMessage Ip:{Client.IPAddress} From:{Message.FromAddress} Total-Recipient:{Message.Recipients.Count}");

			Result.Value = 0;
			Result.Message = string.Empty;
		}

		private class ResultWrapper
		{
			public int Value { get; set; }
			public string Message { get; set; }
		}

		public void OnDeliveryStart(object oMessage, object oResult)
		{
			var Message = Marshal.CreateWrapperOfType<object, MessageClass>(oMessage);
			var Result = Marshal.CreateWrapperOfType<object, ResultClass>(oResult);

			Debug.WriteLine($"OnDeliveryStart Id:{Message.ID}");

			Result.Value = 0;
			Result.Message = string.Empty;
		}

		public void OnDeliverMessage(object oMessage, object oResult)
		{
			var Message = Marshal.CreateWrapperOfType<object, MessageClass>(oMessage);
			var Result = Marshal.CreateWrapperOfType<object, ResultClass>(oResult);

			Debug.WriteLine($"OnDeliverMessage Id:{Message.ID}");

			Result.Value = 0;
			Result.Message = string.Empty;
		}

		public void OnDeliveryFailed(object oMessage, string sRecipient, string sErrorMessage)
		{
			var Message = Marshal.CreateWrapperOfType<object, MessageClass>(oMessage);

			Debug.WriteLine($"OnDeliveryFailed MessageID:{Message.ID} Recipient:{sRecipient} ErrorMessage:{sErrorMessage}");
		}

		public void OnBackupFailed(string sReason)
		{
			Debug.WriteLine($"OnBackupFailed {sReason}");
		}

		public void OnBackupCompleted()
		{
			Debug.WriteLine($"OnBackupCompleted");
		}

		public void OnError(int iSeverity, int iCode, string sSource, string sDescription)
		{
			Debug.WriteLine($"OnError {iSeverity} {iCode} {sSource} {sDescription}");
		}

		public void OnExternalAccountDownload(object oFetchAccount, object oMessage, string sRemoteUID, object oResult)
		{
			var FetchAccount = Marshal.CreateWrapperOfType<object, FetchAccountClass>(oFetchAccount);
			var Message = Marshal.CreateWrapperOfType<object, MessageClass>(oMessage);
			var Result = Marshal.CreateWrapperOfType<object, ResultClass>(oResult);

			Debug.WriteLine($"OnExternalAccountDownload Username:{FetchAccount.Username} MessageId:{Message.ID} RemoteUID:{sRemoteUID}");

			Result.Value = 0;
		}


	}
}
