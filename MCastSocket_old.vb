Imports System
Imports System.Net
Imports System.Net.Sockets
Imports System.Text
Imports Microsoft.VisualBasic
Imports System.IO
Imports System.Threading

Public Class MCastSocket


    ' ManualResetEvent instances signal completion.
    Public connectDone As New ManualResetEvent(False)
    Public sendDone As New ManualResetEvent(False)
    Public receiveDone As New ManualResetEvent(False)

    ' The response from the remote device.
    Public response As String = String.Empty
    Private MySocket As Socket = Nothing
    Private Mystate As StateObject
    Private MyIAsyncResult As IAsyncResult
    Public MySocketIsClosed As Boolean = True
    Private MyReceiveCallBack As Object = Nothing

    Private mcastAddress As IPAddress
    Private mcastPort As Integer
    Private mcastSocket As Socket
    Private mcastOption As MulticastOption
    Private MymcastAddress As String = ""
    Private LocalIEP As IPEndPoint
    Private RemoteEP As IPEndPoint


    Public Class StateObject
        ' State object for receiving data from remote device.
        ' Client socket.
        Public workSocket As Socket = Nothing
        ' Size of receive buffer.
        Public Const BufferSize As Integer = 10000
        ' Receive buffer.
        Public buffer(BufferSize) As Byte
    End Class 'StateObject

    Public Function ConnectSocket(IPLocal As IPEndPoint) As Socket 'LocalIpAddress As String, LocalIPPort As String) As Socket
        ConnectSocket = Nothing
        MySocket = Nothing
        MySocketIsClosed = True
        LocalIEP = IPLocal
        If g_bDebug Then Log("ConnectSocket called with mcastAddress = " & MymcastAddress & " and mcastPort = " & mcastPort & " and LocalIpAddress = " & IPLocal.Address.ToString & " and LocalIPPort = " & IPLocal.Port.ToString, LogType.LOG_TYPE_INFO)
        Try
            remoteEP = New IPEndPoint(mcastAddress, mcastPort)
            ' Create a multicast socket.
            MySocket = New Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp)
            MySocket.ExclusiveAddressUse = False
            MySocket.ReceiveTimeout = 0
            MySocket.SetSocketOption(Net.Sockets.SocketOptionLevel.Socket, Net.Sockets.SocketOptionName.ReuseAddress, True)
            ' Bind this endpoint to the multicast socket.
            MySocket.Bind(LocalIEP)
            ' Define a MulticastOption object specifying the multicast group address and the local IP address. 
            ' The multicast group address is the same as the address used by the listener. 
            Dim mcastOption As MulticastOption
            mcastOption = New MulticastOption(mcastAddress, IPLocal.Address)
            MySocket.SetSocketOption(SocketOptionLevel.IP, SocketOptionName.AddMembership, mcastOption)
            MySocketIsClosed = False
            Return MySocket
        Catch ex As Exception
            If g_bDebug Then Log("Error in ConnectSocket for mcastAddress = " & MymcastAddress.ToString & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            ConnectSocket = Nothing
        End Try
    End Function

    Public Delegate Sub McastSocketClosedEventHandler(sender As Object)
    Public Event MCastSocketClosed As McastSocketClosedEventHandler

    Public Sub CloseSocket()
        ' Release the socket.
        If RemoteEP IsNot Nothing And LocalIEP IsNot Nothing Then
            If g_bDebug Then Log("CloseSocket called with RemoteEP = " & RemoteEP.Address.ToString & " and RemoteEP.port = " & RemoteEP.Port.ToString & " and LocalIEP = " & LocalIEP.Address.ToString & " and LocalIEP.Port = " & LocalIEP.Port.ToString, LogType.LOG_TYPE_INFO)
        Else
            If g_bDebug Then Log("CloseSocket called for mcastAddress = " & MymcastAddress.ToString, LogType.LOG_TYPE_INFO)
        End If

        If MySocket Is Nothing Then Exit Sub
        MySocketIsClosed = True
        Try
            receiveDone.Set()
        Catch ex As Exception
            If RemoteEP IsNot Nothing And LocalIEP IsNot Nothing Then
                If g_bDebug Then Log("Error in CloseSocket setting ReceiveDone for RemoteEP = " & RemoteEP.Address.ToString & " and RemoteEP.port = " & RemoteEP.Port.ToString & " and LocalIEP = " & LocalIEP.Address.ToString & " and LocalIEP.Port = " & LocalIEP.Port.ToString & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Else
                If g_bDebug Then Log("Error in CloseSocket setting ReceiveDone for mcastAddress = " & MymcastAddress.ToString & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End If
        End Try
        Try
            Mystate.workSocket.Shutdown(SocketShutdown.Both)
        Catch ex As Exception
            If RemoteEP IsNot Nothing And LocalIEP IsNot Nothing Then
                If SuperDebug Then Log("Error in CloseSocket shutting down the worksocketin MyState for RemoteEP = " & RemoteEP.Address.ToString & " and RemoteEP.port = " & RemoteEP.Port.ToString & " and LocalIEP = " & LocalIEP.Address.ToString & " and LocalIEP.Port = " & LocalIEP.Port.ToString & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Else
                If SuperDebug Then Log("Error in CloseSocket shutting down the worksocket in MyState for mcastAddress = " & MymcastAddress.ToString & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End If
        End Try
        Try
            MySocket.Shutdown(SocketShutdown.Both)
        Catch ex As Exception
            If RemoteEP IsNot Nothing And LocalIEP IsNot Nothing Then
                If SuperDebug Then Log("Error in CloseSocket shutting down the worksocket for RemoteEP = " & RemoteEP.Address.ToString & " and RemoteEP.port = " & RemoteEP.Port.ToString & " and LocalIEP = " & LocalIEP.Address.ToString & " and LocalIEP.Port = " & LocalIEP.Port.ToString & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Else
                If SuperDebug Then Log("Error in CloseSocket shutting down the worksocket for mcastAddress = " & MymcastAddress.ToString & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End If
        End Try
        Try
            MySocket.Close()
        Catch ex As Exception
            If RemoteEP IsNot Nothing And LocalIEP IsNot Nothing Then
                If g_bDebug Then Log("Error in CloseSocket for RemoteEP = " & RemoteEP.Address.ToString & " and RemoteEP.port = " & RemoteEP.Port.ToString & " and LocalIEP = " & LocalIEP.Address.ToString & " and LocalIEP.Port = " & LocalIEP.Port.ToString & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Else
                If g_bDebug Then Log("Error in CloseSocket for mcastAddress = " & MymcastAddress.ToString & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End If
        End Try
        Try
            RaiseEvent MCastSocketClosed(Me)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub ConnectCallback(ByVal ar As IAsyncResult)
        ' Retrieve the socket from the state object.
        Dim client As Socket = CType(ar.AsyncState, Socket)
        ' Complete the connection.
        'If ar.CompletedSynchronously Then
        Try
            client.EndConnect(ar)
            MySocketIsClosed = False
            If g_bDebug Then Log("ConnectCallback connected a socket to " & client.RemoteEndPoint.ToString(), LogType.LOG_TYPE_INFO)
            connectDone.Set()
        Catch ex As Exception
            If g_bDebug Then Log("Error in ConnectCallback calling EndConnect for mcastAddress = " & MymcastAddress.ToString & " with Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    '
    Public Delegate Sub DataReceivedEventHandler(sender As Object, e As String, ReceiveEP As System.Net.EndPoint)
    Public Event DataReceived As DataReceivedEventHandler
    Public ReceiveEP As System.Net.EndPoint = New IPEndPoint(IPAddress.Any, 0)

    Protected Overridable Sub OnReceive(e As String, ReceiveEp As System.Net.EndPoint)
        Try
            'If g_bDebug Then Log("OnReceive called for mcastAddress = " & MymcastAddress.ToString & " with Data = " & e, LogType.LOG_TYPE_INFO, LogColorNavy) 
            RaiseEvent DataReceived(Me, e, ReceiveEp)
        Catch ex As Exception
            Log("Error in OnReceive for mcastAddress = " & MymcastAddress.ToString & " with Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Function Receive() As Boolean
        Receive = False
        If MySocket Is Nothing Then
            If g_bDebug Then Log("Error in Receive for mcastAddress = " & MymcastAddress.ToString & ". No Socket", LogType.LOG_TYPE_ERROR)
            Exit Function
        End If
        Try
            receiveDone.Reset()
            Mystate = New StateObject
            Mystate.workSocket = MySocket
            MyIAsyncResult = MySocket.BeginReceiveFrom(Mystate.buffer, 0, StateObject.BufferSize, SocketFlags.None, ReceiveEP, New AsyncCallback(AddressOf ReceiveCallback), Mystate)
            'If g_bDebug Then Log("Receive called for mcastAddress = " & MymcastAddress.ToString & " and state = " & MyIAsyncResult.ToString, LogType.LOG_TYPE_INFO, LogColorNavy) 
            Receive = True
        Catch ex As Exception
            Log("Error in Receive for mcastAddress = " & MymcastAddress.ToString & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Private Sub ReceiveCallback(ByVal ar As IAsyncResult)
        'If g_bDebug Then Log("ReceiveCallback called for mcastAddress = " & MymcastAddress.ToString & " with AR = " & ar.AsyncState.ToString, LogType.LOG_TYPE_INFO, LogColorNavy) 
        ' Retrieve the state object and the client socket 
        ' from the asynchronous state object.
        Try
            If MySocketIsClosed Then
                receiveDone.Set()
                If RemoteEP IsNot Nothing And LocalIEP IsNot Nothing Then
                    If g_bDebug Then Log("ReceiveCallback for RemoteEP = " & RemoteEP.Address.ToString & " and RemoteEP.port = " & RemoteEP.Port.ToString & " and LocalIEP = " & LocalIEP.Address.ToString & " and LocalIEP.Port = " & LocalIEP.Port.ToString & " has a closed Socket", LogType.LOG_TYPE_WARNING)
                Else
                    If g_bDebug Then Log("ReceiveCallback for mcastAddress = " & MymcastAddress.ToString & " has a closed Socket", LogType.LOG_TYPE_WARNING)
                End If
                Exit Sub
            End If
        Catch ex As Exception
            Log("Error in ReceiveCallback for mcastAddress = " & MymcastAddress.ToString & " closing socket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Try
                RaiseEvent MCastSocketClosed(Me)
            Catch ex1 As Exception
            End Try
        End Try
        Try
            Dim state As StateObject = CType(ar.AsyncState, StateObject)
            Dim client As Socket = state.workSocket
            ' Read data from the remote device.
            Dim bytesRead As Integer = client.EndReceiveFrom(ar, ReceiveEP)
            If bytesRead > 0 Then
                ' There might be more data, so store the data received so far.
                'If g_bDebug Then Log("ReceiveCallback received data = " & Encoding.ASCII.GetString(state.buffer, 0, bytesRead), LogType.LOG_TYPE_INFO, LogColorNavy) 
                OnReceive(Encoding.UTF8.GetString(state.buffer, 0, bytesRead), ReceiveEP)
            Else
                ' All the data has arrived; put it in response.
                ' Signal that all bytes have been received.
                'If g_bDebug Then Log("ReceiveCallback for mcastAddress = " & MymcastAddress.ToString & " received all data", LogType.LOG_TYPE_WARNING, LogColorNavy)
                receiveDone.Set()
            End If
            MyIAsyncResult = client.BeginReceiveFrom(state.buffer, 0, StateObject.BufferSize, SocketFlags.None, ReceiveEP, New AsyncCallback(AddressOf ReceiveCallback), state)
        Catch ex As Exception
            Log("Error in ReceiveCallback for mcastAddress = " & MymcastAddress.ToString & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Try
                RaiseEvent MCastSocketClosed(Me)
            Catch ex1 As Exception
            End Try
        End Try
    End Sub

    Public Function Send(ByVal data As String) As Boolean
        ' Convert the string data to byte data using ASCII encoding.
        'If g_bDebug Then log( "Send called")
        Send = False
        Try
            If MySocket Is Nothing Then
                sendDone.Set()
                If g_bDebug Then Log("Error in Send for mcastAddress = " & MymcastAddress.ToString & ". No Socket", LogType.LOG_TYPE_ERROR)
                Exit Function
            End If
        Catch ex As Exception
            If g_bDebug Then Log("Error in Send calling SendDone for mcastAddress = " & MymcastAddress.ToString & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If MySocketIsClosed Then
                sendDone.Set()
                If g_bDebug Then Log("Error in Send for mcastAddress = " & MymcastAddress.ToString & ". Socket is closed", LogType.LOG_TYPE_ERROR)
                Exit Function
            End If
        Catch ex As Exception
            If g_bDebug Then Log("Error in Send for mcastAddress = " & MymcastAddress.ToString & " calling SendDone with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Try
                RaiseEvent MCastSocketClosed(Me)
            Catch ex1 As Exception
            End Try
        End Try

        Dim byteData As Byte() = Encoding.UTF8.GetBytes(data)
        ' Begin sending the data to the remote device.
        Try
            'MySocket.BeginSend(byteData, 0, byteData.Length, SocketFlags.None, New AsyncCallback(AddressOf SendCallback), MySocket)
            Dim endPoint As IPEndPoint
            EndPoint = New IPEndPoint(mcastAddress, mcastPort)
            ' mcastSocket.SendTo(ASCIIEncoding.ASCII.GetBytes(message), endPoint)
            MySocket.BeginSendTo(byteData, 0, byteData.Length, SocketFlags.None, endPoint, New AsyncCallback(AddressOf SendCallback), MySocket)
            Send = True
        Catch ex As Exception
            If g_bDebug Then Log("Error in Send for mcastAddress = " & MymcastAddress.ToString & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Try
                RaiseEvent MCastSocketClosed(Me)
            Catch ex1 As Exception
            End Try
        End Try
    End Function

    Private Sub SendCallback(ByVal ar As IAsyncResult)
        ' Retrieve the socket from the state object.
        If MySocketIsClosed Then
            sendDone.Set()
            Exit Sub
        End If
        Try
            Dim client As Socket = CType(ar.AsyncState, Socket)
            ' Complete sending the data to the remote device.
            Dim bytesSent As Integer = client.EndSend(ar)
            If SuperDebug Then Log("SendCallback has sent " & bytesSent & " bytes to server.", LogType.LOG_TYPE_WARNING)
            ' Signal that all bytes have been sent.
            sendDone.Set()
        Catch ex As Exception
            If g_bDebug Then Log("Error in SendCallback for mcastAddress = " & MymcastAddress.ToString & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Try
                RaiseEvent MCastSocketClosed(Me)
            Catch ex1 As Exception
            End Try
        End Try
    End Sub

    Sub New(InputAddress As String, InputPort As Integer)
        MyBase.New()
        MymcastAddress = InputAddress
        mcastAddress = IPAddress.Parse(InputAddress)
        mcastPort = InputPort
        If g_bDebug Then Log("MCastSocket.new was called with MCastAddress = " & InputAddress & " and MCastPort = " & InputPort, LogType.LOG_TYPE_INFO)
    End Sub

End Class


