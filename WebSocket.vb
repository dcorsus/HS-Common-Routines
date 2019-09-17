Imports System.Net
Imports System.Net.Sockets
Imports System.Net.WebSockets
Imports System.Text
Imports System.Threading


Public Class MyWebSocket
    Public NotInheritable Class ClientWebSocket
        Inherits WebSocket

        Public Overrides ReadOnly Property CloseStatus As Nullable(Of WebSocketCloseStatus)
        Public Overrides ReadOnly Property CloseStatusDescription As String
        Public Overrides ReadOnly Property State As WebSocketState
        Public Overrides ReadOnly Property SubProtocol As String


        Public Overrides Sub Abort()
            Me.Abort()
        End Sub

        Public Overrides Function CloseAsync(closeStatus As WebSocketCloseStatus, statusDescription As String, cancellationToken As CancellationToken) As Task
            Return Me.CloseAsync(closeStatus, statusDescription, cancellationToken)
        End Function

        Public Overrides Function CloseOutputAsync(closeStatus As WebSocketCloseStatus, statusDescription As String, cancellationToken As CancellationToken) As Task
            Return Me.CloseOutputAsync(closeStatus, statusDescription, cancellationToken)
        End Function


        Public Function ConnectAsync(uri As Uri, cancellationToken As CancellationToken) As Task
            Return Me.ConnectAsync(uri, cancellationToken)
        End Function



        Public Overrides Function ReceiveAsync(buffer As ArraySegment(Of Byte), cancellationToken As CancellationToken) As Task(Of WebSocketReceiveResult)
            Return Me.ReceiveAsync(buffer, cancellationToken)
        End Function

        Public Overrides Function SendAsync(buffer As ArraySegment(Of Byte), messageType As WebSocketMessageType, endOfMessage As Boolean, cancellationToken As CancellationToken) As Task
            Return Me.SendAsync(buffer, messageType, endOfMessage, cancellationToken)
        End Function


        Public Overrides Sub Dispose()
            Me.Dispose()
        End Sub


        Private Function OpenWebSocket() As MyWebSocket

            OpenWebSocket = Nothing


            Dim aClient As ClientWebSocket = New ClientWebSocket





            With aClient


            End With



        End Function
    End Class



End Class



Public Class WebSocketClient

    ' ManualResetEvent instances signal completion.
    Public connectDone As New ManualResetEvent(False)
    Public sendDone As New ManualResetEvent(False)
    Public receiveDone As New ManualResetEvent(False)

    ' The response from the remote device.
    Public response As Boolean = False
    Private MySocket As Socket = Nothing
    Private Mystate As StateObject
    Private MyIAsyncResult As IAsyncResult
    Public MySocketIsClosed As Boolean = True
    Private MyReceiveCallBack As Object = Nothing
    Private MyRemoteIPAddress As String = ""
    Private MyRemoteIPPort As String = ""
    Private MyLocalIPAddress As String = ""
    Private MyLocalIPPort As String = ""
    Private MyRandomNumberGenerator As New Random()
    Private MyURL As String = ""
    Private WebSocketIsActive As Boolean = False

    ReadOnly Property LocalIPAddress As String
        Get
            LocalIPAddress = MyLocalIPAddress
        End Get
    End Property

    ReadOnly Property LocalIPPort As String
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

    Public Class StateObject
        ' State object for receiving data from remote device.
        ' Client socket.
        Public workSocket As Socket = Nothing
        ' Size of receive buffer.
        Public Const BufferSize As Integer = 9000
        ' Receive buffer.
        Public buffer(BufferSize) As Byte
        ' Received data string.
        'Public sb As New StringBuilder
    End Class 'StateObject

    Public Function UpgradeWebSocket(URL As String, SecWebSocketkey As String) As Boolean
        If g_bDebug Then Log("UpgradeWebSocket called with ipAddress = " & MyRemoteIPAddress & " and URL = " & URL, LogType.LOG_TYPE_INFO)
        Dim SocketDataString As String = "GET /" & URL & " HTTP/1.1" & vbCrLf & "Connection: Upgrade,Keep-Alive" & vbCrLf & "Upgrade: websocket" & vbCrLf & "Sec-WebSocket-Key: " & SecWebSocketkey & vbCrLf & "Sec-WebSocket-Version: 13" & vbCrLf & "Host: " & MyRemoteIPAddress & ":" & MyRemoteIPPort & vbCrLf & vbCrLf

        Dim WaitForConnection As Integer = 0
        Do While WaitForConnection < 10
            If MySocketIsClosed Then
                Wait(1)
                WaitForConnection = WaitForConnection + 1
            Else
                Exit Do
            End If
        Loop

        If WaitForConnection >= 10 Then
            ' unsuccesfull connection
            Log("Error in UpgradeWebSocket for ipAddress = " & MyRemoteIPAddress & ". Unable to open TCP connection within 10 seconds", LogType.LOG_TYPE_ERROR)
            CloseSocket()
            Return False
        End If

        Try
            Receive()
        Catch ex As Exception
            Log("Error in UpgradeWebSocket for ipAddress = " & MyRemoteIPAddress & " unable to receive data to Socket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Return False
        End Try
        response = False
        Try
            If Not Send(System.Text.ASCIIEncoding.ASCII.GetBytes(SocketDataString)) Then
                Log("Error in UpgradeWebSocket for ipAddress = " & MyRemoteIPAddress & " unable to send data to Socket", LogType.LOG_TYPE_ERROR)
                CloseSocket()
                Return False
            End If
        Catch ex As Exception
            Log("Error in UpgradeWebSocket for ipAddress = " & MyRemoteIPAddress & " unable to send data to Socket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            CloseSocket()
            Return False
        End Try
        Try
            sendDone.WaitOne()
        Catch ex As Exception
            Log("Error in UpgradeWebSocket for ipAddress = " & MyRemoteIPAddress & " unable to send data to Socket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            CloseSocket()
            Return False
        End Try
        WaitForConnection = 0
        Do While WaitForConnection < 10
            If MySocketIsClosed Then
                Wait(1)
                WaitForConnection = WaitForConnection + 1
            Else
                Exit Do
            End If
        Loop

        If WaitForConnection >= 10 Then
            ' unsuccesfull connection
            Log("Error in UpgradeWebSocket for ipAddress = " & MyRemoteIPAddress & ". Unable to open TCP connection within 10 seconds", LogType.LOG_TYPE_ERROR)
            CloseSocket()
            Return False
        End If

        Return True

    End Function

    Public Function ConnectSocket(Server As String, ipPort As String) As Socket
        ' Establish the remote endpoint for the socket.
        MyRemoteIPAddress = Server
        MyRemoteIPPort = ipPort
        ConnectSocket = Nothing
        MySocket = Nothing
        MySocketIsClosed = True
        If g_bDebug Then Log("ConnectSocket called with ipAddress = " & Server & " and ipPort = " & ipPort, LogType.LOG_TYPE_INFO)

        Try
            Dim remoteEP As New IPEndPoint(IPAddress.Parse(Server), ipPort)
            ' Create a TCP/IP socket.
            MySocket = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
            ' Connect to the remote endpoint.
            MySocket.BeginConnect(remoteEP, New AsyncCallback(AddressOf ConnectCallback), MySocket)
            ' Wait for connect.
            connectDone.WaitOne() ' I do this in the plugin itself, based on MySocketIsClosed because this runs in its own tread
            Return MySocket
        Catch ex As Exception
            If g_bDebug Then Log("Error in ConnectSocket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            ConnectSocket = Nothing
        End Try
    End Function

    Public Sub CloseSocket()
        ' Release the socket.
        If g_bDebug Then Log("CloseSocket called for IPAddress = " & MyRemoteIPAddress, LogType.LOG_TYPE_INFO)
        If MySocket Is Nothing Then Exit Sub
        MySocketIsClosed = True
        WebSocketIsActive = False
        Try
            receiveDone.Set()
            MySocket.Shutdown(SocketShutdown.Both)
            MySocket.Close()
        Catch ex As Exception
            If g_bDebug Then Log("Error in CloseSocket with ipAddress = " & MyRemoteIPAddress & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        Finally
            OnWebSocketClose()
        End Try
        MySocket = Nothing
    End Sub

    Private Sub ConnectCallback(ByVal ar As IAsyncResult)
        ' Retrieve the socket from the state object.
        Dim client As Socket = CType(ar.AsyncState, Socket)
        ' Complete the connection.
        Try
            client.EndConnect(ar)
            MySocketIsClosed = False
            Dim MylocalEndPoint As System.Net.IPEndPoint
            MylocalEndPoint = client.LocalEndPoint
            MyLocalIPAddress = MylocalEndPoint.Address.ToString
            MyLocalIPPort = MylocalEndPoint.Port.ToString
            If g_bDebug Then Log("ConnectCallback connected a socket to " & client.RemoteEndPoint.ToString(), LogType.LOG_TYPE_INFO)
            connectDone.Set()
        Catch ex As Exception
            If g_bDebug Then Log("Error in ConnectCallback calling EndConnect with ipAddress = " & MyRemoteIPAddress & " with Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    '
    Public Delegate Sub DataReceivedEventHandler(sender As Object, e As Byte())
    Public Event DataReceived As DataReceivedEventHandler

    Public Delegate Sub WebSocketClosedEventHandler(sender As Object)
    Public Event WebSocketClosed As WebSocketClosedEventHandler

    Protected Overridable Sub OnReceive(e As Byte())
        Try
            RaiseEvent DataReceived(Me, e)
        Catch ex As Exception
            Log("Error in OnReceive with ipAddress = " & MyRemoteIPAddress & " with Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Protected Overridable Sub OnWebSocketClose()
        Try
            RaiseEvent WebSocketClosed(Me)
        Catch ex As Exception
            Log("Error in OnWebSocketClose with ipAddress = " & MyRemoteIPAddress & " with Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Function Receive() As Boolean
        Receive = False
        If MySocket Is Nothing Then
            If g_bDebug Then Log("Error in Receive with ipAddress = " & MyRemoteIPAddress & ". No Socket", LogType.LOG_TYPE_ERROR)
            Exit Function
        End If
        Try
            ' Create the state object.
            Mystate = New StateObject
            Mystate.workSocket = MySocket
            ' Begin receiving the data from the remote device.
            MyIAsyncResult = MySocket.BeginReceive(Mystate.buffer, 0, StateObject.BufferSize, SocketFlags.None, New AsyncCallback(AddressOf ReceiveCallback), Mystate)
            If Superdebug Then Log("Receive called and state = " & MyIAsyncResult.IsCompleted.ToString, LogType.LOG_TYPE_WARNING)
            Receive = True
        Catch ex As Exception
            If g_bDebug Then Log("Error in Receive with ipAddress = " & MyRemoteIPAddress & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Private Sub ReceiveCallback(ByVal ar As IAsyncResult)
        'If g_bDebug Then log( "ReceiveCallback called")
        ' Retrieve the state object and the client socket 
        ' from the asynchronous state object.
        Try
            If MySocketIsClosed Then
                receiveDone.Set()
                Exit Sub
            End If
        Catch ex As Exception
            If g_bDebug Then Log("Error in ReceiveCallback closing socket with ipAddress = " & MyRemoteIPAddress & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            Dim state As StateObject = CType(ar.AsyncState, StateObject)
            Dim client As Socket = state.workSocket

            ' Read data from the remote device.
            Dim bytesRead As Integer = client.EndReceive(ar)

            If bytesRead > 0 Then
                ' There might be more data, so store the data received so far.
                If Superdebug Then Log("ReceiveCallback received data = " & Encoding.UTF8.GetString(state.buffer, 0, bytesRead), LogType.LOG_TYPE_INFO)
                'If g_bDebug Then Log("ReceiveCallback received data = " & Encoding.UTF8.GetString(state.buffer, 0, bytesRead), LogType.LOG_TYPE_INFO) 
                response = True
                Dim ByteA As Byte()
                ReDim ByteA(bytesRead)
                Array.Copy(state.buffer, ByteA, bytesRead)
                ByteA = TreatWebSocketData(ByteA)
                If ByteA IsNot Nothing Then OnReceive(ByteA)
                ByteA = Nothing
                ' Get the rest of the data.
                If MySocket IsNot Nothing Then MyIAsyncResult = client.BeginReceive(state.buffer, 0, StateObject.BufferSize, SocketFlags.None, New AsyncCallback(AddressOf ReceiveCallback), state)
            Else
                ' All the data has arrived; put it in response.
                response = True
                ' Signal that all bytes have been received.
                If g_bDebug Then Log("ReceiveCallback with ipAddress = " & MyRemoteIPAddress & " received all data", LogType.LOG_TYPE_WARNING)
                receiveDone.Set()
            End If
        Catch ex As Exception
            If g_bDebug Then Log("Error in ReceiveCallback with ipAddress = " & MyRemoteIPAddress & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Function Send(ByVal inData As Byte()) As Boolean
        ' Convert the string data to byte data using ASCII encoding.
        If Superdebug Then Log("Send called with Data = " & Encoding.ASCII.GetString(inData, 0, inData.Length), LogType.LOG_TYPE_INFO)
        'If g_bDebug Then Log("Send called with Data = " & Encoding.ASCII.GetString(inData, 0, inData.Length), LogType.LOG_TYPE_INFO)
        Send = False
        Try
            If MySocket Is Nothing Then
                sendDone.Set()
                If g_bDebug Then Log("Error in Send with ipAddress = " & MyRemoteIPAddress & ". No Socket", LogType.LOG_TYPE_ERROR)
                Exit Function
            End If
        Catch ex As Exception
            If g_bDebug Then Log("Error in Send with ipAddress = " & MyRemoteIPAddress & " calling SendDone with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If MySocketIsClosed Then
                sendDone.Set()
                If g_bDebug Then Log("Error in Send with ipAddress = " & MyRemoteIPAddress & ". Socket is closed", LogType.LOG_TYPE_ERROR)
                Exit Function
            End If
        Catch ex As Exception
            If g_bDebug Then Log("Error in Send with ipAddress = " & MyRemoteIPAddress & " calling SendDone with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        ' Begin sending the data to the remote device.
        Try
            MyIAsyncResult = MySocket.BeginSend(inData, 0, inData.Length, SocketFlags.None, New AsyncCallback(AddressOf SendCallback), MySocket)
            If Superdebug Then Log("Send called and state = " & MyIAsyncResult.IsCompleted.ToString, LogType.LOG_TYPE_INFO)
            Send = True
        Catch ex As Exception
            If g_bDebug Then Log("Error in Send with ipAddress = " & MyRemoteIPAddress & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
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
            If Superdebug Then Log("SendCallback has sent " & bytesSent & " bytes to server.", LogType.LOG_TYPE_INFO)
            ' Signal that all bytes have been sent.
            sendDone.Set()
        Catch ex As Exception
            If g_bDebug Then Log("Error in SendCallback with ipAddress = " & MyRemoteIPAddress & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Function SendDataOverWebSocket(SocketData As Byte(), UseMask As Boolean) As Boolean
        If g_bDebug Then Log("SendDataOverWebSocket called for ipAddress = " & MyRemoteIPAddress, LogType.LOG_TYPE_INFO)
        SendDataOverWebSocket = False

        If Superdebug Then Log("SendDataOverWebSocket for ipAddress - " & MyRemoteIPAddress & " will send data = " & Encoding.UTF8.GetString(SocketData, 0, SocketData.Length), LogType.LOG_TYPE_INFO)
        Dim GenerateANewIndex As Integer = MyRandomNumberGenerator.Next(1, 429496729)

        Dim Header(3) As Byte ' think of this as the header
        Dim PayLoad As Byte() = SocketData   ' and this the payload
        Header(0) = 129 ' FIN + Opcode = Text

        If SocketData.Length >= 126 Then
            Header(1) = 126 + 128 ' Mask set , 126 means next 2 bytes represent lenght (which must be >126) = FE Hex
            Header(2) = Int(SocketData.Length / 256)
            Header(3) = SocketData.Length Mod 256
        Else
            Header(1) = 128 + CByte(SocketData.Length) ' Mask set
        End If

        Dim startIndex As Integer = 4
        If SocketData.Length < 126 Then startIndex = 2

        If UseMask Then
            Dim Mask() As Byte = BitConverter.GetBytes(GenerateANewIndex)
            Array.Resize(Header, startIndex + PayLoad.Length + 4)
            Mask.CopyTo(Header, startIndex)
            For Index = 0 To UBound(PayLoad)
                PayLoad(Index) = PayLoad(Index) Xor Mask(Index Mod 4)
            Next
            startIndex += 4
        Else
            Array.Resize(Header, startIndex + PayLoad.Length)
        End If

        PayLoad.CopyTo(Header, startIndex)

        Try
            Receive()
        Catch ex As Exception
            Log("Error in SendDataOverWebSocket for ipAddress = " & MyRemoteIPAddress & " unable to receive data to Socket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        response = False
        Try
            If Not Me.Send(Header) Then
                Log("Error in SendDataOverWebSocket for ipAddress = " & MyRemoteIPAddress & " unable to send data to Socket", LogType.LOG_TYPE_ERROR)
                CloseSocket()
                Exit Function
            End If
        Catch ex As Exception
            Log("Error in SendDataOverWebSocket for ipAddress = " & MyRemoteIPAddress & " unable to send data to Socket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            CloseSocket()
            Exit Function
        End Try
        Try
            sendDone.WaitOne()
        Catch ex As Exception
            Log("Error in SendDataOverWebSocket for ipAddress = " & MyRemoteIPAddress & " unable to send data to Socket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            CloseSocket()
            Exit Function
        End Try

        Dim WaitResponse As Integer = 0
        Do While WaitResponse < 10
            If Not response Then
                Wait(1)
                WaitResponse = WaitResponse + 1
            Else
                response = False
                Exit Do
            End If
        Loop

        Try
            sendDone.WaitOne()
        Catch ex As Exception
            Log("Error in SendDataOverWebSocket for ipAddress = " & MyRemoteIPAddress & " unable to send data to Socket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            CloseSocket()
            Exit Function
        End Try

        Return True
    End Function

    Private Function TreatWebSocketData(inData As Byte()) As Byte()
        TreatWebSocketData = Nothing
        If inData Is Nothing Then Return Nothing
        If Superdebug Then Log("TreatWebSocketData called for ipAddress = " & MyRemoteIPAddress & " and Line = " & ASCIIEncoding.ASCII.GetString(inData), LogType.LOG_TYPE_INFO)
        If g_bDebug Then Log("TreatWebSocketData called for ipAddress = " & MyRemoteIPAddress & " Datasize = " & inData.Length.ToString, LogType.LOG_TYPE_INFO)
        If UBound(inData) = 0 Then Return Nothing
        If Not WebSocketIsActive Then
            ' this is most likely the response to the websocket upgrade
            If ASCIIEncoding.ASCII.GetString(inData).IndexOf("101 Switching Protocols") <> -1 Then
                WebSocketIsActive = True
                If g_bDebug Then Log("TreatWebSocketData called for ipAddress = " & MyRemoteIPAddress & " and Line = " & ASCIIEncoding.ASCII.GetString(inData, 0, inData.Length), LogType.LOG_TYPE_INFO)
                Return Nothing
            End If
        End If
        Dim FIN As Boolean = False
        Dim Opcode As Integer = 0
        Dim InfoLength As Integer = 0
        Dim InfoStartOffset As Integer = 2
        Dim MaskBit As Boolean = False
        Dim Mask(3) As Byte
        Dim TextInformation As String = ""

        If inData(0) And 128 <> 0 Then
            ' FIN flag is set, no further 
            FIN = True
        End If
        Opcode = inData(0) And 15
        InfoLength = inData(1) And 127
        MaskBit = (inData(1) And 128) <> 0
        If InfoLength = 126 Then
            ' actually the next 2 bytes represent length
            InfoStartOffset = InfoStartOffset + 2
            InfoLength = inData(2) * 256 + inData(3)
        ElseIf InfoLength = 127 Then
            ' actually the next 8 bytes represent length
            InfoStartOffset = InfoStartOffset + 8
            InfoLength = inData(2) * 256 * 256 * 256 * 256 * 256 * 256 * 256 + inData(3) * 256 * 256 * 256 * 256 * 256 * 256 + inData(4) * 256 * 256 * 256 * 256 * 256 + inData(5) * 256 * 256 * 256 * 256 + inData(6) * 256 * 256 * 256 + inData(7) * 256 * 256 + inData(8) * 256 + inData(9)
        End If
        Dim DecodeBytes As Byte() = Nothing
        If MaskBit Then
            If UBound(inData) >= InfoStartOffset + 3 Then
                For i = 0 To 3
                    Mask(i) = inData(InfoStartOffset + i)
                Next
                InfoStartOffset = InfoStartOffset + 4
                If Superdebug Then Log("TreatWebSocketData for ipAddress - " & MyRemoteIPAddress & " received Mask = " & Mask(0).ToString & " " & Mask(1).ToString & " " & Mask(2).ToString & " " & Mask(3).ToString, LogType.LOG_TYPE_INFO)
                ReDim DecodeBytes(UBound(inData) - InfoStartOffset)
                For Index = InfoStartOffset To UBound(inData)
                    DecodeBytes(Index - InfoStartOffset) = inData(Index) Xor Mask((Index - InfoStartOffset) Mod 4)
                Next
                If Superdebug Then Log("TreatWebSocketData for ipAddress - " & MyRemoteIPAddress & " decoded info = " & ASCIIEncoding.ASCII.GetChars(DecodeBytes), LogType.LOG_TYPE_INFO)
            End If
        End If

        ' opcode
        '       *  %x0 denotes a continuation frame
        '       *  %x1 denotes a text frame
        '       *  %x2 denotes a binary frame
        '       *  %x3-7 are reserved for further non-control frames
        '       *  %x8 denotes a connection close
        '       *  %x9 denotes a ping
        '       *  %xA denotes a pong
        '       *  %xB-F are reserved for further control frames
        If g_bDebug Then Log("TreatWebSocketData for ipAddress - " & MyRemoteIPAddress & " received OpCode = " & Opcode.ToString & ", FIN = " & FIN.ToString & " and Length = " & InfoLength.ToString, LogType.LOG_TYPE_INFO)
        If Superdebug Then Log("TreatWebSocketData for ipAddress - " & MyRemoteIPAddress & " received Byte(0) = " & inData(0).ToString, LogType.LOG_TYPE_INFO)
        If Superdebug Then Log("TreatWebSocketData for ipAddress - " & MyRemoteIPAddress & " received maskbit = " & MaskBit.ToString, LogType.LOG_TYPE_INFO)
        If MaskBit Then
            If Superdebug Then Log("TreatWebSocketData for device - " & MyRemoteIPAddress & " received mask = " & ASCIIEncoding.ASCII.GetChars(Mask), LogType.LOG_TYPE_INFO)
        End If

        If UBound(inData) > 0 Then If Superdebug Then Log("TreatWebSocketData for ipAddress - " & MyRemoteIPAddress & " received Byte(1) = " & inData(1).ToString, LogType.LOG_TYPE_INFO)
        If UBound(inData) > 1 Then If Superdebug Then Log("TreatWebSocketData for ipAddress - " & MyRemoteIPAddress & " received Byte(2) = " & inData(2).ToString, LogType.LOG_TYPE_INFO)
        If UBound(inData) > 2 Then If Superdebug Then Log("TreatWebSocketData for ipAddress - " & MyRemoteIPAddress & " received Byte(3) = " & inData(3).ToString, LogType.LOG_TYPE_INFO)
        If UBound(inData) > 3 Then If Superdebug Then Log("TreatWebSocketData for ipAddress - " & MyRemoteIPAddress & " received Byte(4) = " & inData(4).ToString, LogType.LOG_TYPE_INFO)

        If Opcode = 8 Then
            ' close the connection
            If Not MaskBit Then
                TextInformation = ASCIIEncoding.ASCII.GetChars(inData, InfoStartOffset, inData.Length - InfoStartOffset - 1)
            Else
                If DecodeBytes IsNot Nothing Then TextInformation = ASCIIEncoding.ASCII.GetChars(DecodeBytes)
            End If
            If g_bDebug Then Log("TreatWebSocketData for ipAddress = " & MyRemoteIPAddress & " received close connection with Info = " & TextInformation, LogType.LOG_TYPE_INFO)
            ' return a close 
            Try
                If g_bDebug Then Log("TreatWebSocketData for ipAddress = " & MyRemoteIPAddress & " is sending a close after receiving a close", LogType.LOG_TYPE_INFO)
                response = False
                If Not Send(inData) Then
                    Log("Error in TreatWebSocketData for ipAddress = " & MyRemoteIPAddress & " while sending a close", LogType.LOG_TYPE_ERROR)
                    CloseSocket()
                    Return Nothing
                End If
            Catch ex As Exception
                Log("Error in TreatWebSocketData for ipAddress = " & MyRemoteIPAddress & " while sending a close with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                CloseSocket()
                Return Nothing
            End Try
            Try
                sendDone.WaitOne()
            Catch ex As Exception
                Log("Error in TreatWebSocketData for ipAddress = " & MyRemoteIPAddress & " while sending a close with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                CloseSocket()
                Return Nothing
            End Try
            CloseSocket()
        ElseIf Opcode = 9 Then ' Ping
            ' send a Pong
            Dim GenerateANewIndex As Integer = MyRandomNumberGenerator.Next(1, 429496729)
            ' use infolength and startinfoindex
            Dim Array1(1) As Byte
            Array1(0) = 138 '128 + 10 ' set FIN and pong opcode ;
            Array1(1) = 128 + InfoLength ' set mask bit and length

            ' it appears the MASK must be set
            Dim SendMask() As Byte = BitConverter.GetBytes(GenerateANewIndex)

            Array.Resize(Array1, 2 + 4) ' 4 is for the mask
            Mask.CopyTo(Array1, 2)

            If InfoLength > 0 Then
                ' copy the rest
                Array.Resize(Array1, 6 + InfoLength) ' 6 is header + mask
                Try
                    For Index = 0 To InfoLength - 1
                        Array1(Index + 6) = inData(Index + InfoStartOffset) Xor Mask(Index Mod 4)
                    Next
                Catch ex As Exception
                    Log("Error in TreatWebSocketData while preparing a pong for ipAddress = " & MyRemoteIPAddress & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
            End If

            Try
                If g_bDebug Then Log("TreatWebSocketData for ipAddress - " & MyRemoteIPAddress & " is sending a pong after receiving a ping", LogType.LOG_TYPE_INFO)
                If Superdebug Then Log("TreatWebSocketData for ipAddress - " & MyRemoteIPAddress & " is sending a pong Byte(0) = " & Array1(0).ToString, LogType.LOG_TYPE_INFO)
                If Superdebug Then Log("TreatWebSocketData for ipAddress - " & MyRemoteIPAddress & " is sending a pong Byte(1) = " & Array1(1).ToString, LogType.LOG_TYPE_INFO)
                If UBound(Array1) > 1 Then If Superdebug Then Log("TreatWebSocketData for ipAddress - " & MyRemoteIPAddress & " is sending a pong Byte(2) = " & Array1(2).ToString, LogType.LOG_TYPE_INFO)
                If UBound(Array1) > 2 Then If Superdebug Then Log("TreatWebSocketData for ipAddress - " & MyRemoteIPAddress & " is sending a pong Byte(3) = " & Array1(3).ToString, LogType.LOG_TYPE_INFO)

                response = False
                If Not Send(Array1) Then
                    Log("Error in TreatWebSocketData for ipAddress = " & MyRemoteIPAddress & " while sending a pong", LogType.LOG_TYPE_ERROR)
                    CloseSocket()
                    Return Nothing
                End If
            Catch ex As Exception
                Log("Error in TreatWebSocketData for ipAddress = " & MyRemoteIPAddress & " while sending a pong with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                CloseSocket()
                Return Nothing
            End Try
            Try
                sendDone.WaitOne()
            Catch ex As Exception
                Log("Error in TreatWebSocketData for ipAddress = " & MyRemoteIPAddress & " while sending a pong with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                CloseSocket()
            End Try
            Return Nothing
        ElseIf Opcode = 10 Then ' Pong
        ElseIf Opcode = 1 Then ' TextFrame
            If Not MaskBit Then
                Array.Resize(DecodeBytes, inData.Length - InfoStartOffset - 1)
                Buffer.BlockCopy(inData, InfoStartOffset, DecodeBytes, 0, inData.Length - InfoStartOffset - 1)
                'TextInformation = ASCIIEncoding.ASCII.GetChars(inData, InfoStartOffset, inData.Length - InfoStartOffset - 1)
            End If
            If g_bDebug Then Log("TreatWebSocketData for ipAddress - " & MyRemoteIPAddress & " received Info =  " & ASCIIEncoding.ASCII.GetChars(DecodeBytes), LogType.LOG_TYPE_INFO)
            Return DecodeBytes
        ElseIf Opcode = 2 Then ' BinaryFrame
            ' not sure the treatment here is different from a TextFrame
            If Not MaskBit Then
                Array.Resize(DecodeBytes, inData.Length - InfoStartOffset - 1)
                Buffer.BlockCopy(inData, InfoStartOffset, DecodeBytes, 0, inData.Length - InfoStartOffset - 1)
            End If
            If g_bDebug Then Log("TreatWebSocketData for ipAddress - " & MyRemoteIPAddress & " received Info =  " & ASCIIEncoding.ASCII.GetChars(DecodeBytes), LogType.LOG_TYPE_INFO)
            Return DecodeBytes
        End If
        Return Nothing
    End Function

End Class
