Imports System
Imports System.IO
Imports System.Net
Imports System.Net.Sockets
Imports System.Text
Imports Microsoft.VisualBasic
Imports System.Threading


Class MyUdpClient
    Public connectDone As New ManualResetEvent(False)
    Public sendDone As New ManualResetEvent(False)
    Public receiveDone As New ManualResetEvent(False)
    Private myState As StateObject
    Private myIAsyncResult As IAsyncResult
    Private myUdpClient As UdpClient = Nothing
    Private isConnected As Boolean = False
    Private myLocalIPAddress As String = ""
    Private mGrpAddress As String = ""
    Private myLocalIPPort As Integer = 0
    Private receivedByteCount As Integer = 0
    Public response As String = String.Empty
    Public Class StateObject
        ' State object for receiving data from remote device.
        ' Client socket.
        Public workSocket As UdpClient = Nothing
        ' Size of receive buffer.
        Public Const BufferSize As Integer = 9000
        ' Receive buffer.
        Public buffer(BufferSize) As Byte
    End Class
    Public Property BytesReceived As Integer
        Get
            BytesReceived = receivedByteCount
        End Get
        Set(value As Integer)
            receivedByteCount = value
        End Set
    End Property
    ReadOnly Property LocalIPAddress As String
        Get
            LocalIPAddress = myLocalIPAddress
        End Get
    End Property
    ReadOnly Property LocalIPPort As Integer
        Get
            LocalIPPort = myLocalIPPort
        End Get
    End Property

    Public Delegate Sub UdpSocketClosedEventHandler(sender As Object)
    Public Event UdpSocketClosed As UdpSocketClosedEventHandler

    Public Delegate Sub ReceiveEventHandler(ReceiveStatus As Boolean)
    Public Event recOK As ReceiveEventHandler

    Public Delegate Sub DataReceivedEventHandler(sender As Object, e As String, ReceiveEP As System.Net.EndPoint)
    Public Event DataReceived As DataReceivedEventHandler
    Public ReceiveEP As System.Net.EndPoint = New IPEndPoint(IPAddress.Any, 0)

    Sub New(localIPAddress As String, localPort As Integer)
        MyBase.New()
        If upnpDebuglevel > DebugLevel.dlErrorsOnly Then Log("MyUdpClient.new was called with localIPAddress = " & myLocalIPAddress & ", localPort = " & myLocalIPPort, LogType.LOG_TYPE_INFO)
        myLocalIPAddress = localIPAddress
        myLocalIPPort = localPort
    End Sub

    Public Function ConnectSocket(grpAddress As String) As UdpClient
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("MyUdpClient.ConnectSocket called with localIPAddress = " & myLocalIPAddress & ", local port = " & myLocalIPPort.ToString & " and groupAddress = " & grpAddress, LogType.LOG_TYPE_INFO)
        mGrpAddress = grpAddress
        ConnectSocket = Nothing
        Try
            ' Bind And listen on port the specified local port. This constructor creates a socket 
            ' And binds it to the port on which to receive data. The family 
            ' parameter specifies that this connection uses an IPv4 address.
            myUdpClient = New UdpClient() With {
                    .ExclusiveAddressUse = False,
                    .EnableBroadcast = True,
                    .MulticastLoopback = True
            }
            myUdpClient.Client.SetSocketOption(SocketOptionLevel.IP, SocketOptionName.MulticastTimeToLive, 4)
            myUdpClient.Client.SetSocketOption(Net.Sockets.SocketOptionLevel.Socket, Net.Sockets.SocketOptionName.ReuseAddress, True)

            If grpAddress <> "" Then
                Dim localEP As System.Net.IPEndPoint = New IPEndPoint(IPAddress.Any, myLocalIPPort)
                myUdpClient.Client.Bind(localEP)
            Else
                Dim localIPAddress As IPAddress = IPAddress.Parse(myLocalIPAddress)
                Dim localIpEndPoint As IPEndPoint = New IPEndPoint(localIPAddress, myLocalIPPort)
                myUdpClient.Client.Bind(localIpEndPoint)
            End If
            Dim ListenerEndPoint As System.Net.IPEndPoint = myUdpClient.Client.LocalEndPoint
            myLocalIPPort = ListenerEndPoint.Port
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then Log("MyUdpClient.ConnectSocket had an error creating a UdpClient with localIPAddress = " & myLocalIPAddress & ", local port = " & myLocalIPPort.ToString & " and groupAddress = " & grpAddress & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        connectDone.Reset()
        Try
            ' Join the specified multicast group using one of the 
            ' JoinMulticastGroup overloaded methods.
            If mGrpAddress <> "" Then myUdpClient.JoinMulticastGroup(IPAddress.Parse(mGrpAddress))
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then Log("MyUdpClient.ConnectSocket had an error joining the multicast group groupAddress = " & mGrpAddress & ", local port = " & myLocalIPPort.ToString & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        Return myUdpClient
    End Function

    Public Function Receive() As Boolean
        Receive = False
        If myUdpClient Is Nothing Then
            If upnpDebuglevel > DebugLevel.dlOff Then Log("Error in MyUdpClient.Receive for local IP Address = " & myLocalIPAddress.ToString & ". No Socket", LogType.LOG_TYPE_ERROR)
            Exit Function
        End If
        Try
            receiveDone.Reset()
            myState = New StateObject With {
                .workSocket = myUdpClient
            }
            With myUdpClient
                myIAsyncResult = .BeginReceive(New AsyncCallback(AddressOf ReceiveCallback), myState)
                isConnected = True
                ' Wait until a connection is made and processed before  
                ' continuing.
                'connectDone.WaitOne()
                ' if local port 0 was used, a port will be dynamically assigned. Capture it and store it for potential use
                Dim ListenerEndPoint As System.Net.IPEndPoint = myUdpClient.Client.LocalEndPoint
                myLocalIPPort = ListenerEndPoint.Port
                Receive = True
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("MyUdpClient.Receive successfully opened Udp listener Port on interface = " & ListenerEndPoint.Address.ToString & " and local port = " & myLocalIPPort.ToString, LogType.LOG_TYPE_INFO)
            End With
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then Log("MyUdpClient.Receive had an error while begining to listening on interface = " & myLocalIPAddress & ", port = " & myLocalIPPort.ToString & ", multicast group groupAddress = " & mGrpAddress & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            isConnected = False
        End Try
    End Function

    Private Sub ReceiveCallback(ByVal ar As IAsyncResult)
        If piDebuglevel > DebugLevel.dlEvents Then Log("ReceiveCallback called", LogType.LOG_TYPE_INFO)
        Try
            If Not isConnected Then
                receiveDone.Set()
                Exit Sub
            End If
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in MyUdpClient.ReceiveCallback closing socket on interface = " & myLocalIPAddress & ", port = " & myLocalIPPort.ToString & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            Dim state As StateObject = CType(ar.AsyncState, StateObject)
            Dim client As UdpClient = state.workSocket

            ' Read data from the remote device.
            Dim receivedBytes As Byte() = client.EndReceive(ar, ReceiveEP)

            If receivedBytes IsNot Nothing AndAlso receivedBytes.Count > 0 Then
                ' There might be more data, so store the data received so far.
                If PIDebuglevel > DebugLevel.dlEvents Then Log("MyUdpClient.ReceiveCallback from remote address = " & ReceiveEP.ToString & ", received data = " & Encoding.UTF8.GetString(receivedBytes), LogType.LOG_TYPE_INFO)
                receivedByteCount += receivedBytes.Count
                OnReceive(Encoding.UTF8.GetString(receivedBytes), ReceiveEP)
                ' Get the rest of the data.
                If Not isConnected Or (client Is Nothing) Then
                    ' this could be if the OnReceive data forced the connection to close
                    isConnected = True
                    receiveDone.Set()
                    Exit Sub
                End If
                myIAsyncResult = myUdpClient.BeginReceive(New AsyncCallback(AddressOf ReceiveCallback), myState)
            Else
                ' All the data has arrived; put it in response.
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("MyUdpClient.ReceiveCallback on interface = " & myLocalIPAddress & ", port = " & myLocalIPPort.ToString & " received all data and connected state = " & isConnected.ToString, LogType.LOG_TYPE_WARNING)
                receiveDone.Set()
            End If
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in MyUdpClient.ReceiveCallback on interface = " & myLocalIPAddress & ", port = " & myLocalIPPort.ToString & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Try
                RaiseEvent UdpSocketClosed(Me) ' from testing, this appears to be a valid case
            Catch ex1 As Exception
            End Try
        End Try
    End Sub

    Protected Overridable Sub OnReceive(e As String, ReceiveEp As System.Net.EndPoint)
        Try
            RaiseEvent DataReceived(Me, e, ReceiveEp)
        Catch ex As Exception
            Log("Error in MyUdpClient.OnReceive on interface = " & myLocalIPAddress & ", port = " & myLocalIPPort.ToString & " and Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Function Send(ByVal data As String, remoteAddress As String, remotePort As Integer) As Boolean
        ' Convert the string data to byte data using ASCII encoding.
        If piDebuglevel > DebugLevel.dlEvents Then Log("MyUdpClient.Send called with Data = " & data & ", remote IP Address = " & remoteAddress & " and remote port = " & remotePort.ToString, LogType.LOG_TYPE_INFO)
        Send = False
        Try
            If myUdpClient Is Nothing Then
                sendDone.Set()
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in MyUdpClient.Send with remote IP Address = " & remoteAddress & " and remote port = " & remotePort.ToString & ". No Socket", LogType.LOG_TYPE_ERROR)
                Exit Function
            End If
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in MyUdpClient.Send with remote IP Address = " & remoteAddress & " and remote port = " & remotePort.ToString & " calling SendDone with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If Not isConnected Then
                sendDone.Set()
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in MyUdpClient.Send with remote IP Address = " & remoteAddress & " and remote port = " & remotePort.ToString & ". Socket is closed", LogType.LOG_TYPE_ERROR)
                Exit Function
            End If
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in MyUdpClient.Send with remote IP Address = " & remoteAddress & " and remote port = " & remotePort.ToString & " calling SendDone with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        ' Begin sending the data to the remote device.
        Dim byteData As Byte() = Encoding.UTF8.GetBytes(data)
        Try
            myIAsyncResult = myUdpClient.BeginSend(byteData, byteData.Length, remoteAddress, remotePort, New AsyncCallback(AddressOf SendCallback), myState)
            If piDebuglevel > DebugLevel.dlEvents Then Log("MyUdpClient.Send called and state = " & myIAsyncResult.IsCompleted.ToString, LogType.LOG_TYPE_INFO)
            Send = True
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in MyUdpClient.Send with remote IP Address = " & remoteAddress & " and remote port = " & remotePort.ToString & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Try
                RaiseEvent UdpSocketClosed(Me)
            Catch ex1 As Exception
            End Try
        End Try
    End Function

    Private Sub SendCallback(ByVal ar As IAsyncResult)
        ' Retrieve the socket from the state object.
        If Not isConnected Then
            sendDone.Set()
            Exit Sub
        End If
        Try
            Dim state As StateObject = CType(ar.AsyncState, StateObject)
            Dim client As UdpClient = state.workSocket
            ' Complete sending the data to the remote device.
            Dim bytesSent As Integer = client.EndSend(ar)
            If piDebuglevel > DebugLevel.dlEvents Then Log("MyUdpClient.SendCallback has sent " & bytesSent & " bytes to server.", LogType.LOG_TYPE_INFO)
            ' Signal that all bytes have been sent.
            sendDone.Set()
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in MyUdpClient.SendCallback with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Try
                RaiseEvent UdpSocketClosed(Me)
            Catch ex1 As Exception
            End Try
        End Try
    End Sub

    Public Sub CloseSocket()
        If myUdpClient Is Nothing Then Exit Sub
        If Not isConnected Then
            Try
                myUdpClient.Dispose()
            Catch ex As Exception
            End Try
            myUdpClient = Nothing
            Exit Sub
        End If
        Try
            If mGrpAddress <> "" Then myUdpClient.DropMulticastGroup(IPAddress.Parse(mGrpAddress))
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then Log("Error in MyUdpClient.CloseSocket closing a Listenener with groupAddress = " & mGrpAddress & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        isConnected = False ' move here to avoid error in receivecallback
        Try
            myUdpClient.Close()
        Catch ex As Exception
        End Try
        Try
            myUdpClient.Dispose()
        Catch ex As Exception
        End Try
        myUdpClient = Nothing
    End Sub

End Class


