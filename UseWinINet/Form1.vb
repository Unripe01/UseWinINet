Public Class Form1
    Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Int32, _
ByVal dwReserved As Int32) As Boolean

    Private Declare Function InternetDial Lib "Wininet.dll" (ByVal hwndParent As IntPtr, _
    ByVal lpszConnectoid As String, ByVal dwFlags As Int32, ByRef lpdwConnection As Int32, _
    ByVal dwReserved As Int32) As Int32

    Private Declare Function InternetHangUp Lib "Wininet.dll" _
    (ByVal lpdwConnection As Int32, ByVal dwReserved As Int32) As Int32

    Private Enum Flags As Integer
        'Local system uses a LAN to connect to the Internet.
        INTERNET_CONNECTION_LAN = &H2
        'Local system uses a modem to connect to the Internet.
        INTERNET_CONNECTION_MODEM = &H1
        'Local system uses a proxy server to connect to the Internet.
        INTERNET_CONNECTION_PROXY = &H4
        'Local system has RAS installed.
        INTERNET_RAS_INSTALLED = &H10
    End Enum

    'Declaration Used For InternetDialUp.
    Private Enum DialUpOptions As Integer
        INTERNET_DIAL_UNATTENDED = &H8000
        INTERNET_DIAL_SHOW_OFFLINE = &H4000
        INTERNET_DIAL_FORCE_PROMPT = &H2000
    End Enum

    Private Const ERROR_SUCCESS = &H0
    Private Const ERROR_INVALID_PARAMETER = &H87


    Private mlConnection As Int32

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim lngFlags As Long

        If InternetGetConnectedState(lngFlags, 0) Then
            'connected.
            If lngFlags And Flags.INTERNET_CONNECTION_LAN Then
                'LAN connection.
                MsgBox("LAN connection.")
            ElseIf lngFlags And Flags.INTERNET_CONNECTION_MODEM Then
                'Modem connection.
                MsgBox("Modem connection.")
            ElseIf lngFlags And Flags.INTERNET_CONNECTION_PROXY Then
                'Proxy connection.
                MsgBox("Proxy connection.")
            End If
        Else
            'not connected.
            MsgBox("Not connected.")
        End If
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Dim DResult As Int32

        DResult = InternetDial(Me.Handle, "My Connection", DialUpOptions.INTERNET_DIAL_FORCE_PROMPT, mlConnection, 0)

        If (DResult = ERROR_SUCCESS) Then
            MessageBox.Show("Dial Up Successful", "Dial-Up Connection")
        Else
            MessageBox.Show("UnSuccessFull Error Code" & DResult, "Dial-Up Connection")
        End If
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        Dim Result As Int32

        If Not (mlConnection = 0) Then
            Result = InternetHangUp(mlConnection, 0&)
            If Result = 0 Then
                MessageBox.Show("Hang up successful", "Hang Up Connection")
            Else
                MessageBox.Show("Hang up NOT successful", "Hang Up Connection")
            End If
        Else
            MessageBox.Show("You must dial a connection first!", "Hang Up Connection")
        End If
    End Sub
End Class
