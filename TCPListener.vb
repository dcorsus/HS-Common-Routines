Imports System
Imports System.IO
Imports System.Net
Imports System.Net.Sockets
Imports System.Text
Imports Microsoft.VisualBasic
Imports System.Threading


Class MyTCPListener
    Public connectDone As New ManualResetEvent(False)
    Public sendDone As New ManualResetEvent(False)
    Public receiveDone As New ManualResetEvent(False)

    Private MyListenSocket As TcpListener = Nothing
    Private isConnected As Boolean = False
    Private MyRemoteIPAddress As String = ""
    Private MyRemoteIPPort As String = ""
    Private MyLocalIPAddress As String = ""
    Private MyLocalIPPort As Integer = 0
    Private receiveStatus As Boolean = False
    Private MyTCPResponse As String = ""
    Private receivedByteCount As Integer = 0
    Private debugParam As String

    Public ReadOnly Property CheckDebugParam As Boolean
        Get
            If debugParam = "" Then Return True
            Return (debugParam = MyRemoteIPAddress)
        End Get
    End Property

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
            LocalIPAddress = MyLocalIPAddress
        End Get
    End Property

    ReadOnly Property LocalIPPort As Integer
        Get
            LocalIPPort = MyLocalIPPort
        End Get
    End Property

    ReadOnly Property RemoteIPAddress As String
        Get
            RemoteIPAddress = MyRemoteIPAddress
        End Get
    End Property

    ReadOnly Property RemoteIPPort As String
        Get
            RemoteIPPort = MyRemoteIPPort
        End Get
    End Property

    WriteOnly Property TCPResponse As String
        Set(value As String)
            MyTCPResponse = value
        End Set
    End Property

    Public Delegate Sub ConnectedEventHandler(isConnected As Boolean)
    Public Event Connection As ConnectedEventHandler

    Public Delegate Sub ReceiveEventHandler(ReceiveStatus As Boolean)
    Public Event recOK As ReceiveEventHandler

    Public Delegate Sub DataEventHandler(Data As String)
    Public Event DataReceived As DataEventHandler

    Public Function Start(hostAdress As String, Optional port As Integer = 0) As Integer
        If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyTcpListener.Start called with local IP Address = " & hostAdress & " and requested local port = " & port.ToString, LogType.LOG_TYPE_INFO)
        MyLocalIPAddress = hostAdress
        Start = 0
        Try
            MyListenSocket = New TcpListener(IPAddress.Parse(hostAdress), port)
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("MyTcpListener.Start had an error creating a TCPListenener with local IP Address = " & hostAdress & " and requested local port = " & port.ToString & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        connectDone.Reset()
        Try
            With MyListenSocket
                .Start()
                .BeginAcceptTcpClient(New AsyncCallback(AddressOf DoAccept), MyListenSocket)
                isConnected = True
                ' Wait until a connection is made and processed before  
                ' continuing.
                'connectDone.WaitOne()
                Dim ListenerEndPoint As System.Net.IPEndPoint = MyListenSocket.LocalEndpoint
                MyLocalIPPort = ListenerEndPoint.Port
                Start = MyLocalIPPort
                If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyTcpListener.Start successfully opened TCP Listening Port with Local IP Address = " & hostAdress & " and local port = " & MyLocalIPPort.ToString, LogType.LOG_TYPE_INFO)
            End With
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("MyTcpListener.Start had an error while begining to listening on local IP address = " & hostAdress & " and requested local port = " & port.ToString & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            isConnected = False
        Finally
            RaiseEvent Connection(isConnected)
        End Try
    End Function

    Private Sub DoAccept(ByVal ar As IAsyncResult)

        Dim sb As New StringBuilder

        Dim listener As TcpListener
        Dim clientSocket As TcpClient
        If Not isConnected Then Exit Sub
        If ar Is Nothing Then Exit Sub

        Try
            listener = CType(ar.AsyncState, TcpListener)
            'If listener Is Nothing Then Exit Sub
            clientSocket = listener.EndAcceptTcpClient(ar)
            clientSocket.ReceiveTimeout = 10000
            'clientSocket.NoDelay = True
            Dim iremote As System.Net.IPEndPoint
            Dim pLocal As System.Net.IPEndPoint
            iremote = clientSocket.Client.RemoteEndPoint
            pLocal = clientSocket.Client.LocalEndPoint
            MyRemoteIPAddress = iremote.Address.ToString
            MyRemoteIPPort = iremote.Port.ToString
            If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MyTcpListener.DoAccept active on HostAddress = " & MyLocalIPAddress & " and HostIPPort = " & MyLocalIPPort & " and RemoteAddress = " & MyRemoteIPAddress & " and remoteIPPort = " & MyRemoteIPPort, LogType.LOG_TYPE_INFO)
            'If upnpDebuglevel > DebugLevel.dlErrorsOnly andAlso CheckDebugParam Then Log("MyTcpListener.DoAccept active on HostAddress = " & MyLocalIPAddress & " and HostIPPort = " & MyLocalIPPort & " and RemoteAddress = " & MyRemoteIPAddress & " and remoteIPPort = " & MyRemoteIPPort, LogType.LOG_TYPE_INFO)

        Catch ex As ObjectDisposedException
            'Log("MyTcpListener.DoAccept had an error while start listening on HostAddress = " & MyLocalIPAddress & " and HostIPPort = " & MyLocalIPPort & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            ' after server.stop() AsyncCallback is also active, but the object server is disposed 
            isConnected = False
            RaiseEvent Connection(isConnected)
            Exit Sub
        End Try

        ' Signal the calling thread to continue.
        'receiveDone.Set()

        ' Buffer for reading data 
        Dim bytes(32000) As Byte
        Dim data As String = Nothing
        Dim HTTPContentLength As Integer = 0
        Dim HTTPHeaderLength As Integer = 0
        Dim ChunkedTransmission As Boolean = False
        Dim ResponseSent As Boolean = False
        Dim BytesReceived As Integer = 0 ' added in v028 11/28/2018. Old interpretation had bytes converted to ascii and international characters went from 2 bytes to 1 character and caused timeouts.


        Try
            With clientSocket
                ' Get a stream object for reading and writing 
                Dim stream As NetworkStream = clientSocket.GetStream()
                Dim i As Integer ' Int32
                ' Loop to receive all the data sent by the client.
                i = stream.Read(bytes, 0, bytes.Length)
                While (i <> 0)
                    ' Translate data bytes to a ASCII string.
                    BytesReceived += i
                    receivedByteCount += i
                    data = System.Text.Encoding.UTF8.GetString(bytes, 0, i)
                    If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MyTcpListener.DoAccept Received data = " & data, LogType.LOG_TYPE_INFO)
                    'If upnpDebuglevel > DebugLevel.dlErrorsOnly andAlso CheckDebugParam Then Log("MyTcpListener.DoAccept Received data with length = " & data.Length, LogType.LOG_TYPE_WARNING)
                    sb.Append(data)
                    Try
                        If HTTPContentLength = 0 Then
                            HTTPContentLength = ParseHTTPResponse(sb.ToString, "content-length:")
                        End If
                    Catch ex As Exception
                        HTTPContentLength = 0
                    End Try
                    Try
                        If HTTPHeaderLength = 0 Then
                            HTTPHeaderLength = ParseHTTPResponseGetHeaderLength(sb.ToString)
                        End If
                    Catch ex As Exception
                        HTTPHeaderLength = 0
                    End Try
                    'If HTTPContentLength <> 0 And HTTPHeaderLength <> 0 Then
                    If HTTPHeaderLength <> 0 Then
                        'If upnpDebuglevel > DebugLevel.dlErrorsOnly andAlso CheckDebugParam Then Log("MyTcpListener.DoAccept Received data with length = " & sb.Length & " and is looking for = " & (HTTPContentLength + HTTPHeaderLength).ToString, LogType.LOG_TYPE_WARNING)
                        Dim TransferEncoding As String = ParseHTTPResponse(sb.ToString, "TRANSFER-ENCODING:")
                        If TransferEncoding <> "" Then
                            If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MyTcpListener.DoAccept Received a chunked request with TransferEncoding = " & TransferEncoding, LogType.LOG_TYPE_WARNING)
                            If TransferEncoding.ToLower = "chunked" Then
                                ChunkedTransmission = True
                            End If
                        End If
                        If (Not ChunkedTransmission) And ((HTTPContentLength + HTTPHeaderLength) = BytesReceived) Then '  changed in v028 11/28/2018 sb.Length) Then
                            ' all received
                            'If upnpDebuglevel > DebugLevel.dlEvents andAlso CheckDebugParam Then Log("MyTcpListener.DoAccept Received data with length = " & sb.Length.ToString & " and sent a successful response", LogType.LOG_TYPE_INFO)
                            If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MyTcpListener.DoAccept Received data with length = " & sb.Length.ToString & " and sent a successful response", LogType.LOG_TYPE_INFO)
                            ' we need to send a HTTP/1.1 200 OK response
                            TCPResponse = ""
                            Try
                                RaiseEvent DataReceived(sb.ToString)
                            Catch ex As Exception
                                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("MyTcpListener.DoAccept error raising DataReceived with Error = " & ex.Message, LogType.LOG_TYPE_INFO)
                            End Try
                            sendDone.WaitOne()
                            If MyTCPResponse <> "" Then
                                If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyTcpListener.DoAccept is returning Response = " & MyTCPResponse, LogType.LOG_TYPE_INFO, LogColorGreen)
                                ' Dim msg As [Byte]() = System.Text.Encoding.ASCII.GetBytes("HTTP/1.1 200 OK" & vbCrLf & "Connection: close" & vbCrLf & "Content-Length: 0" & vbCrLf & vbCrLf)
                                Dim msg As [Byte]() = System.Text.Encoding.UTF8.GetBytes(MyTCPResponse)
                                ' Send back a response.
                                stream.Write(msg, 0, msg.Length)
                                ResponseSent = True
                                MyTCPResponse = ""
                            End If
                            Exit While
                        End If
                    End If
                    i = stream.Read(bytes, 0, bytes.Length)
                End While
                If Not ResponseSent Then
                    If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MyTcpListener.DoAccept Received data but did not send a response", LogType.LOG_TYPE_WARNING)
                End If
                .Close()
            End With
            receiveStatus = True
        Catch ex As TimeoutException
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("MyTcpListener.DoAccept timeout = " & ex.Message, LogType.LOG_TYPE_ERROR)
            receiveStatus = False
            clientSocket.Close()
            'isConnected = False
            'RaiseEvent Connection(isConnected)
            'Exit Sub   removed this 12/14/2019 when I had an error and the listener went completely dead
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("MyTcpListener.DoAccept received Error = " & ex.Message, LogType.LOG_TYPE_INFO)
            receiveStatus = False
            clientSocket.Close()
            'isConnected = False
            'RaiseEvent Connection(isConnected)
            'Exit Sub   removed this 12/14/2019 when I had an error and the listener went completely dead
        Finally
            'RaiseEvent recOK(receiveStatus)
        End Try

        If MyListenSocket IsNot Nothing Then MyListenSocket.BeginAcceptTcpClient(New AsyncCallback(AddressOf DoAccept), MyListenSocket)

    End Sub

    Public Sub Close()
        Try
            If MyListenSocket IsNot Nothing Then
                If isConnected Then
                    With MyListenSocket
                        .Stop()
                    End With
                    isConnected = False
                End If
                MyListenSocket = Nothing
            End If
        Catch ex As Exception
            Log("MyTcpListener.Close had an error closing a TCPListenener with HostAddress = " & MyLocalIPAddress & " and HostIPPort = " & MyLocalIPPort & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try
    End Sub

End Class 'MyTcpListener 


