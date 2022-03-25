
using System;
using System.Runtime.InteropServices;

namespace hMailServer
{
	[ComVisible(true)]
	[Guid("F7AFE3CE-E4A1-40F9-B21D-CC643D53EAF1")]
	[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	interface IConnector
	{
		void OnSMTPData(object oClient, object oMessage, object oResult);
		void OnClientConnect(object oClient, object oResult);
		void OnAcceptMessage(object oClient, object oMessage, object oResult);
		void OnDeliveryStart(object oMessage, object oResult);
		void OnDeliverMessage(object oMessage, object oResult);
		void OnDeliveryFailed(object oMessage, string sRecipient, string sErrorMessage);
		void OnBackupFailed(string sReason);
		void OnBackupCompleted();
		void OnError(int iSeverity, int iCode, string sSource, string sDescription);
		void OnExternalAccountDownload(object oFetchAccount, object oMessage, string sRemoteUID, object oResult);
	}
}
