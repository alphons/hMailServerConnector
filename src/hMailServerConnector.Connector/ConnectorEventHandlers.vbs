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