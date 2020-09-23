Attribute VB_Name = "Misc"
Public Function SockStatus(ByVal sock As Winsock) As String
  Dim strMessage As String
    Select Case sock.State

        Case StateConstants.sckConnected
            strMessage = "Connected to " & sock.RemoteHost
        Case StateConstants.sckClosing
            strMessage = "Closing connection to " & sock.RemoteHost
        Case StateConstants.sckClosed
            strMessage = "Not Connected"
        Case StateConstants.sckError
            strMessage = "Error in Socket"
        Case StateConstants.sckConnected
            strMessage = "Connecting to " & sock.RemoteHost
        Case StateConstants.sckHostResolved
            strMessage = sock.RemoteHost & " Resolved"
        Case StateConstants.sckOpen
            strMessage = "Opened Socket"
        Case StateConstants.sckResolvingHost
            strMessage = "Resolving " & sock.RemoteHost
        Case StateConstants.sckConnectionPending
            strMessage = "Connection is in Pending"
        Case StateConstants.sckListening
            strMessage = "Awaiting connection Request"

    End Select

SockStatus = strMessage
End Function

