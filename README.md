# hMailServerConnector
hMailServer connector for interfacing events to .NET Core

This project connects the world famous hMailServer application to the .NET Core world.
The Connecter is build and exposes itself as a COM interface which can be called by
the events system build in in hMailServer (jscript and vbscript glue).

Copy ConnectorEventHandlers.vbs to EventHandlers.vbs in hMailserver/events directory.

```
Option Explicit

Dim connector
Set connector = CreateObject("hMailServer.Connector")

Sub OnClientConnect(oClient)
  Call connector.OnClientConnect(oClient, Result)
End Sub

Sub OnSMTPData(oClient, oMessage)
  Call connector.OnSMTPData(oClient,oMessage, Result)
End Sub

Sub OnAcceptMessage(oClient, oMessage)
  Call connector.OnAcceptMessage(oClient,oMessage, Result)
End Sub

Sub OnDeliveryStart(oMessage)
  Call connector.OnDeliveryStart(oMessage, Result)
End Sub

Sub OnDeliverMessage(oMessage)
  Call connector.OnDeliverMessage(oMessage, Result)
End Sub

Sub OnBackupFailed(sReason)
  Call connector.OnBackupFailed(sReason)
End Sub

Sub OnBackupCompleted()
  Call connector.OnBackupCompleted()
End Sub

Sub OnError(iSeverity, iCode, sSource, sDescription)
  Call connector.OnError(iSeverity, iCode, sSource, sDescription)
End Sub

Sub OnDeliveryFailed(oMessage, sRecipient, sErrorMessage)
  Call connector.OnDeliverMessage(oMessage, sRecipient, sErrorMessage)
End Sub

Sub OnExternalAccountDownload(oFetchAccount, oMessage, sRemoteUID)
  Call connector.OnExternalAccountDownload(oFetchAccount, oMessage, sRemoteUID, Result)
End Sub
```


Run _install.cmd script once, it does a 'regsvr32.exe hMailServerConnector.comhost.dll'

Because of COM nature, the hmailserver service must be stopped before re-compiling.
To make things easier, the pre en post build steps does this automatically by
calling: sc stop hmailserver and sc start hmailserver. To make this work Visual Studio must be run as ~~root~~ Administrator.

For debugging there are 2 options. 
First one is to attach debugger to the hmailserver.exe process.
Shortcut Ctrl-Alt-P (show processes for all users)
After this first attach, you can use shortcut Shift-Alt-P for re-attach process.

Second option for debugging is use Debugger.Launch(); in sourcecode, to start the jus-in-time debugger.
The COM visible .NET Core library exposes the following interface:
```
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
```

When all the magic happends we can code C# .NET Core 6

```
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
```

The project uses 32 bit architecture because of the 32 bit interop assembly of hmailserver.

Happy coding!


