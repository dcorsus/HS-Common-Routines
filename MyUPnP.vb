Imports System.Text
Imports System.Net
Imports System.Net.Sockets
Imports System.Net.NetworkInformation
Imports System.Xml
Imports System.Drawing
Imports System.IO
Imports System.Web.UI.WebControls
Imports System.Web.UI
Imports System.Web
Imports System.Web.UI.HtmlControls
Imports Scheduler.PageBuilderAndMenu
Imports Scheduler

Public Class MySSDP

    ' http://www.codeproject.com/Articles/458807/UPnP-code-for-Windows-8 for more info on how UPnP works

    Private SSDPAsyncSocket As MyUdpClient = Nothing 'MCastSocket = Nothing
    Private MySSDPClient As UdpClient 'Socket
    Private SSDPAsyncSocketLoopback As MyUdpClient = Nothing
    Private MySSDPClientLoopback As UdpClient
    Private UPNPMonitoringDevice As String = ""

    Private MulticastAsyncSocket As MyUdpClient = Nothing
    Private MyMulticastClient As UdpClient 'Socket

    Private MyUPnPMCastIPAddress As String = "239.255.255.250"
    Private MyUPnPMCastPort As Integer = 1900
    Private MyDevicesLinkedList As New MyUPnPDevices
    Private myDiscoverUdpPort As Integer = 0
    Private myActiveDiscoverUdpPort As Integer = -1

    Private MyEventListener As MyTCPListener
    Private RestartListenerFlag As Boolean = False
    Private ListenerIsActive As Boolean = False
    Private RestartMulticastListenerFlag As Boolean = False
    Private MulticastListenerIsActive As Boolean = False

    Private NotificationHandlerReEntryFlag As Boolean = False
    Private MyNotificationQueue As Queue(Of String) = New Queue(Of String)()
    Private MissedNotificationHandlerFlag As Boolean = False

    Friend WithEvents MyControllerTimer As Timers.Timer
    Friend WithEvents MyNotifyTimer As Timers.Timer
    Friend WithEvents myAuditTimer As Timers.Timer

    Private isClosing As Boolean = False
    Private overallBytesReceivedMulticast As Integer = 0
    Private isTheSameCounter As Integer = 0
    Private overallFailedRenewCounter As Integer = 0

    Public Delegate Sub NewDeviceEventHandler(UDN As String)
    Public Event NewDeviceFound As NewDeviceEventHandler
    Public Delegate Sub MCastDiedEventHandler()
    Public Event MCastDiedEvent As MCastDiedEventHandler
    Public Delegate Sub MSearchEventHandler(MSearchInfo As String)
    Public Event MSearchEvent As MSearchEventHandler
    Private debugParam As String

    Public ReadOnly Property CheckDebugParam As Boolean
        Get
            If debugParam = "" Then Return True
            Return True
        End Get
    End Property

    Public Property DiscoverUdpPort As Integer
        Get
            Return myActiveDiscoverUdpPort
        End Get
        Set(value As Integer)

        End Set
    End Property

    Public ReadOnly Property GetRXCounters As Integer
        Get
            If MulticastAsyncSocket IsNot Nothing Then
                Return MulticastAsyncSocket.BytesReceived
            Else
                Return -1
            End If
        End Get
    End Property

    Public Sub New()
        MyBase.New()
        Try
            MyControllerTimer = New Timers.Timer With {
                .Interval = 60000,  ' every minute
                .AutoReset = True,
                .Enabled = True
            }
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.New. Unable to create the overall control timer with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            MyNotifyTimer = New Timers.Timer With {
                .Interval = 500, ' 1/2 second
                .AutoReset = False,
                .Enabled = False
            }
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.New. Unable to create the notification timer with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            myAuditTimer = New Timers.Timer With {
                .Interval = 300000,  ' every 5 minutes
                .AutoReset = True,
                .Enabled = True
            }
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.New. Unable to create the audit timer with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            CreateMulticastListener()
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.New. Unable to create the multicast listener with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub Dispose()
        ' this disposes all devices and services as part of Parent SSDP
        If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MySSDP.Dispose called for PlugInIPAddress = " & PlugInIPAddress, LogType.LOG_TYPE_INFO)
        isClosing = True
        Try
            If MyControllerTimer IsNot Nothing Then MyControllerTimer.Enabled = False
            MyControllerTimer = Nothing
        Catch ex As Exception
        End Try
        Try
            If MyNotifyTimer IsNot Nothing Then MyNotifyTimer.Enabled = False
            MyNotifyTimer = Nothing
        Catch ex As Exception
        End Try
        Try
            If myAuditTimer IsNot Nothing Then myAuditTimer.Enabled = False
            myAuditTimer = Nothing
        Catch ex As Exception
        End Try
        ' 
        Try
            StopListener()
        Catch ex As Exception
        End Try
        Try
            If SSDPAsyncSocket IsNot Nothing Then
                Try
                    RemoveHandler SSDPAsyncSocket.DataReceived, AddressOf HandleSSDPDataReceived
                    RemoveHandler SSDPAsyncSocket.UdpSocketClosed, AddressOf HandleSSDPUdpSocketCloseEvent
                Catch ex As Exception
                End Try
                SSDPAsyncSocket.CloseSocket()
                SSDPAsyncSocket = Nothing
            End If
        Catch ex As Exception
        End Try
        Try
            If SSDPAsyncSocketLoopback IsNot Nothing Then
                Try
                    RemoveHandler SSDPAsyncSocketLoopback.DataReceived, AddressOf HandleSSDPDataReceived
                    RemoveHandler SSDPAsyncSocketLoopback.UdpSocketClosed, AddressOf HandleSSDPUdpSocketCloseEvent
                Catch ex As Exception
                End Try
                SSDPAsyncSocketLoopback.CloseSocket()
                SSDPAsyncSocketLoopback = Nothing
            End If
        Catch ex As Exception
        End Try
        Try
            If MulticastAsyncSocket IsNot Nothing Then
                Try
                    RemoveHandler MulticastAsyncSocket.DataReceived, AddressOf HandleSSDPDataReceived
                    RemoveHandler MulticastAsyncSocket.UdpSocketClosed, AddressOf HandleMCastSocketCloseEvent
                Catch ex As Exception
                End Try
                MulticastAsyncSocket.CloseSocket()
                MulticastAsyncSocket = Nothing
            End If
        Catch ex As Exception
        End Try
        Try
            If MyDevicesLinkedList IsNot Nothing Then
                MyDevicesLinkedList.Dispose()
                MyDevicesLinkedList = Nothing
            End If
        Catch ex As Exception
        End Try
    End Sub

    Public Sub CreateMulticastListener(Optional alsoTCPListener As Boolean = True)
        ' this creates the permanent listening port for ssdp messages such as ssdp:alive and ssdp:byebye
        If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MySSDP.CreateMulticastListener called for PlugInIPAddress = " & PlugInIPAddress, LogType.LOG_TYPE_INFO)
        If MulticastAsyncSocket Is Nothing Then
            Try
                MulticastAsyncSocket = New MyUdpClient(MyUPnPMCastIPAddress, MyUPnPMCastPort)
            Catch ex As Exception
                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.CreateMulticastListener unable to create a Multicast Socket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Sub
            End Try
            If MulticastAsyncSocket Is Nothing Then
                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.CreateMulticastListener. No AsyncSocket!", LogType.LOG_TYPE_ERROR)
                Exit Sub
            End If
            AddHandler MulticastAsyncSocket.DataReceived, AddressOf HandleSSDPDataReceived
            AddHandler MulticastAsyncSocket.UdpSocketClosed, AddressOf HandleMCastSocketCloseEvent
            Try
                MyMulticastClient = MulticastAsyncSocket.ConnectSocket(MyUPnPMCastIPAddress)
            Catch ex As Exception
                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.CreateMulticastListener unable to connect Socket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Sub
            End Try
        End If

        If MyMulticastClient Is Nothing Then
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.CreateMulticastListener. No Client!", LogType.LOG_TYPE_ERROR)
            HandleMCastSocketCloseEvent(Nothing)    ' close the socket completely           ' added 10/21/2019 in v42 to deal with proper retries when failing to open socket
            'RestartMulticastListenerFlag = True    ' this is set in HandleMCastCloseEvent
            Exit Sub
        Else
            RestartMulticastListenerFlag = False
        End If

        If alsoTCPListener Then StartEventListener(TCPListenerPort)

        Try
            MulticastAsyncSocket.Receive()
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.CreateMulticastListener for UPnPDevice = " & "" & " unable to receive data to Socket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

    End Sub

    Public Sub HandleMCastSocketCloseEvent(sender As Object)
        ' this is the event handler for the "regular" multicast listener, NOT the SSDP - discovery initiated listener. If this terminates, restart it by setting the restartlistener flag
        If isClosing Then Exit Sub
        If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.HandleMCastSocketCloseEvent received Close Socket Event. Will try to restart", LogType.LOG_TYPE_ERROR)
        ' we close the ssdpsocket here, will be restarted with next call to DoSSDPDiscovery
        Try
            ' RaiseEvent MCastDiedEvent() ' removed in v19
        Catch ex As Exception
        End Try
        Try
            If MulticastAsyncSocket IsNot Nothing Then
                Try
                    RemoveHandler MulticastAsyncSocket.DataReceived, AddressOf HandleSSDPDataReceived
                    RemoveHandler MulticastAsyncSocket.UdpSocketClosed, AddressOf HandleMCastSocketCloseEvent
                Catch ex As Exception
                End Try
                Try
                    MulticastAsyncSocket.CloseSocket()
                Catch ex As Exception
                End Try
                MulticastAsyncSocket = Nothing
            End If
        Catch ex As Exception
        End Try
        RestartMulticastListenerFlag = True
    End Sub

    Public Function RemoveDevices(UDN As String) As Boolean
        RemoveDevices = False
        If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("MySSDP.RemoveDevices called with UDN = " & UDN, LogType.LOG_TYPE_INFO, LogColorGreen)
        Dim DeviceToRemove As MyUPnPDevice = MyDevicesLinkedList.Item(UDN, True)
        If DeviceToRemove IsNot Nothing Then
            Try
                If Not MyDevicesLinkedList.RemoveDevice(DeviceToRemove, UDN) Then
                    If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.RemoveDevices. Device with UDN = " & UDN & " was not removed successfully", LogType.LOG_TYPE_ERROR)
                Else
                    RemoveDevices = True
                End If
            Catch ex As Exception
            End Try
        End If
    End Function

#Region "SSDP"

    Public Sub HandleSSDPDataReceived(sender As Object, e As String, ReceiveEP As System.Net.EndPoint)
        If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MySSDP.HandleSSDPDataReceived received Line = " & e & " and EP = " & ReceiveEP.ToString, LogType.LOG_TYPE_INFO, LogColorNavy)
        'If UPnPDebuglevel > DebugLevel.dlErrorsOnly andAlso CheckDebugParam Then Log("HandleSSDPDataReceived received Line.length = " & e.Length, LogType.LOG_TYPE_WARNING, LogColorNavy) 
        'If UPnPDebuglevel > DebugLevel.dlOff andAlso CheckDebugParam Then Log("MySSDP.HandleSSDPDataReceived received Line = " & e & " and EP = " & ReceiveEP.ToString, LogType.LOG_TYPE_INFO, LogColorNavy) 
        Try
            SyncLock (MyNotificationQueue)
                MyNotificationQueue.Enqueue(e & "receiveep:" & ReceiveEP.ToString & vbCrLf & vbCrLf)
            End SyncLock
            MyNotifyTimer.Enabled = True
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.HandleSSDPDataReceived queuing the Notification = " & e.ToString & " and Error = " & ex.Message, LogType.LOG_TYPE_INFO)
        End Try
        sender = Nothing
        e = Nothing
        ReceiveEP = Nothing
    End Sub

    Dim NbrOfMessageCounter As Integer = 0
    Dim NbrOfNotifyAliveMessageCounter As Integer = 0
    Dim NbrOfSearchMessageCounter As Integer = 0
    Dim NbrOfNotifyByeByeMessageCounter As Integer = 0
    Dim NbrOfHTTPMessageCounter As Integer = 0
    Dim NbrOfUnknownMessageCounter As Integer = 0

    Private Sub TreatNotficationQueue()
        If NotificationHandlerReEntryFlag Then
            If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MySSDP.TreatNotficationQueue has Re-Entry while processing Notification queue with # elements = " & MyNotificationQueue.Count.ToString, LogType.LOG_TYPE_WARNING)
            MissedNotificationHandlerFlag = True
            Exit Sub
        End If
        NotificationHandlerReEntryFlag = True
        If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MySSDP.TreatNotficationQueue is processing Notification queue with # elements = " & MyNotificationQueue.Count.ToString, LogType.LOG_TYPE_INFO)
        'If UPnPDebuglevel > DebugLevel.dlErrorsOnly andAlso CheckDebugParam Then Log("TreatNotficationQueue is processing Notification queue with # elements = " & MyNotificationQueue.Count.ToString, LogType.LOG_TYPE_INFO, LogColorGreen) 
        Dim NotificationEvent As String = ""

        Try
            While MyNotificationQueue.Count > 0

                SyncLock (MyNotificationQueue)
                    NotificationEvent = MyNotificationQueue.Dequeue
                    'MyNotificationQueue.TrimExcess()
                End SyncLock
                NbrOfMessageCounter += 1
                'If UPnPDebuglevel > DebugLevel.dlOff andAlso CheckDebugParam Then Log("MySSDP.TreatNotficationQueue is processing Notification Event Nbr = " & NbrOfMessageCounter.ToString, LogType.LOG_TYPE_INFO, LogColorGreen)
                If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MySSDP.TreatNotficationQueue is processing Notification = " & NotificationEvent, LogType.LOG_TYPE_INFO, LogColorGreen)
                Try
                    Dim Header As String = ParseSsdpResponse(NotificationEvent, "") ' calling this procedure with empty parameter returns the first line in the response
                    Dim URI As String = ParseSsdpResponse(NotificationEvent, "location")
                    Dim USN As String = ParseSsdpResponse(NotificationEvent, "usn")
                    Dim CacheControl As String = ParseSsdpResponse(NotificationEvent, "cache-control")
                    Dim NT As String = ParseSsdpResponse(NotificationEvent, "nt")                       ' Notification Target
                    Dim NTType As NTtype = ParseNT(NT)
                    Dim ST As String = ParseSsdpResponse(NotificationEvent, "st")                       ' Notification Target
                    Dim NTS As String = ParseSsdpResponse(NotificationEvent, "nts")                       ' Search Target
                    Dim Server As String = ParseSsdpResponse(NotificationEvent, "server")
                    Dim BootID As String = ParseSsdpResponse(NotificationEvent, "bootid.upnp.org")
                    If BootID = "" Then BootID = ParseSsdpResponse(NotificationEvent, "x-rincon-bootseq")     ' added 10/13/2019
                    Dim ConfigID As String = ParseSsdpResponse(NotificationEvent, "configid.upnp.org")
                    Dim SearchPort As String = ParseSsdpResponse(NotificationEvent, "searchport.upnp.org")  ' port the device will respond to for Unicast M-Search
                    Dim ReceiveEP As String = ParseSsdpResponse(NotificationEvent, "receiveep")
                    Dim WakeUp As String = ParseSsdpResponse(NotificationEvent, "wakeup") ' part of DIAL protocal; example WAKEUP: MAC=10:dd:b1:c9:00:e4;Timeout=10
                    Dim SSID As String = ParseSsdpResponse(NotificationEvent, "ssid") ' part of DIAL protocal; example WAKEUP: MAC=10:dd:b1:c9:00:e4;Timeout=10
                    ' Sonos has 
                    ' X-RINCON-BOOTSEQ
                    ' X-RINCON-HOUSEHOLD
                    ' X-RINCON-PROXY
                    ' X-RINCON-WIFIMODE
                    ' X-RINCON-VARIANT

                    Dim UDN As String = ""
                    If Header <> "M-SEARCH" Then
                        Dim USNParts As String() = Split(USN, "::") ' first remove UUID from Service UID
                        If USNParts IsNot Nothing Then
                            If UBound(USNParts, 1) >= 0 Then
                                UDN = Trim(USNParts(0))
                                If UDN = "" Then
                                    If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MySSDP.TreatNotficationQueue retrieved invalid UDN with Data = " & NotificationEvent, LogType.LOG_TYPE_WARNING)
                                    Exit Try   ' ignore this, wrong UDN
                                End If
                                If UDN.IndexOf("uuid:") <> 0 Then
                                    If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MySSDP.TreatNotficationQueue retrieved invalid UDN with Data = " & NotificationEvent, LogType.LOG_TYPE_WARNING)
                                    Exit Try   ' ignore this, wrong UDN
                                End If
                                Mid(UDN, 1, 5) = "     "
                                UDN = Trim(UDN)
                                If UDN.IndexOf(":upnp:rootdevice") <> -1 Then
                                    ' this is to cover the issue w/ Roku's which have a single colon instead of ::upnp:rootdevice
                                    UDN = UDN.Remove(UDN.IndexOf(":upnp:rootdevice"), 16)
                                    UDN = Trim(UDN)
                                End If
                            End If
                        End If
                        USNParts = Nothing
                    End If

                    If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MySSDP.TreatNotficationQueue retrieved URI = " & URI.ToString & ", NTS = " & NTS & ", NT = " & NT.ToString & ", ST = " & ST.ToString & " and USN = " & USN.ToString & " Cache-Control = " & CacheControl, LogType.LOG_TYPE_INFO, LogColorGreen)

                    ' If ST = upnp:rootdevice is present, then this was part of an SSDP discovery, else NTS will be present with either ssdp:alive or ssdp:byebye
                    If Header = "M-SEARCH" Then
                        NbrOfSearchMessageCounter += 1
                        ' do nothing, this is a discovery call from other UPNP controllers on the network
                        ' maybe for future use if I want to become a UPNP Server
                        If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MySSDP.TreatNotficationQueue received M-Search info = " & NotificationEvent & " From = " & ReceiveEP, LogType.LOG_TYPE_INFO, LogColorGreen)
                        Try
                            RaiseEvent MSearchEvent(NotificationEvent)
                        Catch ex As Exception
                            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.TreatNotficationQueue while generating MSearchEvent with Info = " & NotificationEvent & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                    ElseIf Header = "HTTP" And ST = UPNPMonitoringDevice Then '"upnp:rootdevice" Then
                        NbrOfHTTPMessageCounter += 1
                        If Not MyDevicesLinkedList.CheckDeviceExists(UDN) Then
                            Dim NewDevice As MyUPnPDevice
                            If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MySSDP.TreatNotficationQueue is adding rootdevice with URI = " & URI.ToString & ", NTS = " & NTS & ", NT = " & NT.ToString & ", ST = " & ST.ToString & " and USN = " & USN.ToString & " Cache-Control = " & CacheControl, LogType.LOG_TYPE_INFO, LogColorGreen)
                            NewDevice = MyDevicesLinkedList.AddDevice(UDN, URI, True, Me)
                            If NewDevice IsNot Nothing Then
                                If NewDevice.Location <> "" Then
                                    Dim IPInfo As IPAddressInfo
                                    IPInfo = ExtractIPInfo(URI)
                                    NewDevice.IPAddress = IPInfo.IPAddress
                                    NewDevice.IPPort = IPInfo.IPPort
                                    NewDevice.Server = Server
                                    NewDevice.BootID = BootID
                                    NewDevice.ConfigID = ConfigID
                                    NewDevice.CacheControl = CacheControl
                                    NewDevice.TimeoutValue = RetrieveTimeoutData(CacheControl)
                                    NewDevice.WakeUp = WakeUp
                                    NewDevice.SSID = SSID
                                    NewDevice.NotificationEvent = NotificationEvent
                                    If SearchPort <> "" Then NewDevice.SearchPort = SearchPort
                                    NewDevice.RootDevice = NewDevice
                                    If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MySSDP.TreatNotficationQueue retrieved IPInfo with IPAddress =  " & NewDevice.IPAddress & " and IPPort = " & NewDevice.IPPort & " and Location = " & NewDevice.Location, LogType.LOG_TYPE_INFO, LogColorGreen)
                                End If
                            Else
                                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.TreatNotficationQueue for device = " & NewDevice.UniqueDeviceName & ". Could not add the device to our list with Data = " & NotificationEvent, LogType.LOG_TYPE_ERROR)
                            End If
                            NewDevice = Nothing
                        Else
                            ' device exists, check whether location changed (which mostly means the IPAddress changed!
                            Dim ExistingDevice As MyUPnPDevice = MyDevicesLinkedList.Item("uuid:" & UDN, True)
                            If ExistingDevice IsNot Nothing Then
                                If ExistingDevice.CheckForUpdates(URI, "uuid:" & UDN, Me) Then
                                    ' retrieve device info is done in the checkforupdates procedure
                                End If
                                ExistingDevice.Server = Server
                                If BootID <> ExistingDevice.BootID Then
                                    If ExistingDevice.BootID <> "" And ExistingDevice.IsRootDevice And upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MySSDP.TreatEventQueue with M-Search for device = " & ExistingDevice.UniqueDeviceName & " had different bootids. Stored = " & ExistingDevice.BootID & " <-> received = " & BootID, LogType.LOG_TYPE_INFO, LogColorGreen)
                                    ExistingDevice.BootID = BootID
                                End If
                                'ExistingDevice.BootID = BootID
                                ExistingDevice.ConfigID = ConfigID
                                ExistingDevice.CacheControl = CacheControl
                                ' this is new code 10/13/2019 v37
                                If Not ExistingDevice.IsRootDevice Then
                                    ' only update the root timeout ie one timer for all related devices and services added 10/13/2019                                    
                                    If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MySSDP.TreatNotficationQueue for device = " & ExistingDevice.UniqueDeviceName & " is updating the timeout value of its parent with Value = " & RetrieveTimeoutData(CacheControl), LogType.LOG_TYPE_INFO, LogColorGreen)
                                    Dim rootDevice As MyUPnPDevice = ExistingDevice.RootDevice
                                    If rootDevice.TimeoutValue < RetrieveTimeoutData(CacheControl) Then
                                        rootDevice.TimeoutValue = RetrieveTimeoutData(CacheControl)
                                    Else
                                        rootDevice.TimeoutValue = rootDevice.TimeoutValue ' Ok this looks strange but it resets the timer
                                    End If
                                Else
                                    ExistingDevice.TimeoutValue = RetrieveTimeoutData(CacheControl)
                                End If
                                ExistingDevice.NotificationEvent = NotificationEvent
                                If SearchPort <> "" Then ExistingDevice.SearchPort = SearchPort
                            Else
                                ' this should never be
                                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.TreatNotficationQueue for discovered UUID = " & UDN & ". Could not find the device to our list with Data = " & NotificationEvent, LogType.LOG_TYPE_ERROR)
                            End If
                            ExistingDevice = Nothing
                        End If
                    ElseIf Header = "NOTIFY" And NTS = "ssdp:alive" Then
                        NbrOfNotifyAliveMessageCounter += 1
                        ' Is sent once with upnp:rootdevice for the root device, sent once in uuid:device-uuid for each device root or embedded
                        If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MySSDP.TreatNotficationQueue received ssdp:alive with URI = " & URI.ToString & ", NTS = " & NTS & ", NT = " & NT.ToString & ", NTType = " & NTType.ToString & " and USN = " & USN.ToString, LogType.LOG_TYPE_INFO, LogColorGreen)
                        Select Case NTType
                            Case MySSDP.NTtype.NTTypeRootDevice
                                If UPNPMonitoringDevice = "upnp:rootdevice" Then
                                    If Not MyDevicesLinkedList.CheckDeviceExists(UDN) Then
                                        If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MySSDP.TreatNotficationQueue received ssdp:alive and is adding device with URI = " & URI.ToString & ", NTS = " & NTS & ", NT = " & NT.ToString & ", ST = " & ST.ToString & " and USN = " & USN.ToString & " Cache-Control = " & CacheControl, LogType.LOG_TYPE_INFO, LogColorGreen)
                                        Dim NewDevice As MyUPnPDevice
                                        NewDevice = MyDevicesLinkedList.AddDevice(UDN, URI, True, Me) ' will also force device discovery and processing
                                        If NewDevice IsNot Nothing Then
                                            If NewDevice.Location <> "" Then
                                                Dim IPInfo As IPAddressInfo
                                                IPInfo = ExtractIPInfo(URI)
                                                NewDevice.IPAddress = IPInfo.IPAddress
                                                NewDevice.IPPort = IPInfo.IPPort
                                                NewDevice.Server = Server
                                                NewDevice.BootID = BootID
                                                NewDevice.ConfigID = ConfigID
                                                NewDevice.CacheControl = CacheControl
                                                NewDevice.TimeoutValue = RetrieveTimeoutData(CacheControl)
                                                NewDevice.WakeUp = WakeUp
                                                NewDevice.SSID = SSID
                                                NewDevice.NotificationEvent = NotificationEvent
                                                If SearchPort <> "" Then NewDevice.SearchPort = SearchPort
                                                NewDevice.RootDevice = NewDevice
                                                ' NewDevice.Alive = True still need to retrieve device info
                                                If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MySSDP.TreatNotficationQueue retrieved IPInfo with IPAddress =  " & NewDevice.IPAddress & " and IPPort = " & NewDevice.IPPort & " and Location = " & NewDevice.Location, LogType.LOG_TYPE_INFO, LogColorGreen)
                                            End If
                                        Else
                                            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.TreatNotficationQueue for device = " & NewDevice.UniqueDeviceName & ". Could not add the device to our list with Data = " & NotificationEvent, LogType.LOG_TYPE_ERROR)
                                        End If
                                        NewDevice = Nothing
                                    Else
                                        ' device exists, check whether location changed (which mostly means the IPAddress changed!
                                        Dim ExistingDevice As MyUPnPDevice = MyDevicesLinkedList.Item("uuid:" & UDN, True)
                                        If ExistingDevice IsNot Nothing Then
                                            If ExistingDevice.CheckForUpdates(URI, "uuid:" & UDN, Me) Then
                                                ' retrieve device info is done in the checkforupdates procedure
                                            End If
                                            ExistingDevice.Server = Server
                                            If BootID <> ExistingDevice.BootID Then
                                                If ExistingDevice.BootID <> "" And ExistingDevice.IsRootDevice And ExistingDevice.BootID <> "" And upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MySSDP.TreatEventQueue with ssdp:alive for root device = " & ExistingDevice.UniqueDeviceName & " had different bootids. Stored = " & ExistingDevice.BootID & " <-> received = " & BootID, LogType.LOG_TYPE_INFO, LogColorGreen)
                                                ExistingDevice.BootID = BootID
                                            End If
                                            'ExistingDevice.BootID = BootID
                                            ExistingDevice.ConfigID = ConfigID
                                            ExistingDevice.CacheControl = CacheControl
                                            If Not ExistingDevice.IsRootDevice Then
                                                ' only update the root timeout ie one timer for all related devices and services added 10/13/2019
                                                If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MySSDP.TreatNotficationQueue for root device = " & ExistingDevice.UniqueDeviceName & " is updating the timeout value of its parent (??) with Value = " & RetrieveTimeoutData(CacheControl), LogType.LOG_TYPE_INFO, LogColorGreen)
                                                Dim rootDevice As MyUPnPDevice = ExistingDevice.RootDevice
                                                If rootDevice.TimeoutValue < RetrieveTimeoutData(CacheControl) Then
                                                    rootDevice.TimeoutValue = RetrieveTimeoutData(CacheControl)
                                                Else
                                                    rootDevice.TimeoutValue = rootDevice.TimeoutValue ' Ok this looks strange but it resets the timer
                                                End If
                                            Else
                                                ExistingDevice.TimeoutValue = RetrieveTimeoutData(CacheControl)
                                            End If
                                            If SearchPort <> "" Then ExistingDevice.SearchPort = SearchPort
                                            ExistingDevice.NotificationEvent = NotificationEvent
                                        Else
                                            ' this should never be
                                            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.TreatNotficationQueue for alive UUID = " & UDN & ". Could not find the device to our list with Data = " & NotificationEvent, LogType.LOG_TYPE_ERROR)
                                        End If
                                        ExistingDevice = Nothing
                                    End If
                                Else
                                    ' ignore these, the URN NTType should pick up the specific ones
                                End If
                            Case MySSDP.NTtype.NTTypeDevice, MySSDP.NTtype.NTTypeURNDevice
                                'Log("TreatNotficationQueue for ssdp:alive found for type device or type urndevice = " & NT & " ST = " & ST & " with UDN = " & UDN, LogType.LOG_TYPE_WARNING)
                                If NT = UPNPMonitoringDevice And UPNPMonitoringDevice <> "upnp:rootdevice" And Not MyDevicesLinkedList.CheckDeviceExists(UDN) Then
                                    If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MySSDP.TreatNotficationQueue received ssdp:alive and is adding device with URI = " & URI.ToString & ", NTS = " & NTS & ", NT = " & NT.ToString & ", ST = " & ST.ToString & " and USN = " & USN.ToString & " Cache-Control = " & CacheControl, LogType.LOG_TYPE_INFO, LogColorGreen)
                                    Dim NewDevice As MyUPnPDevice
                                    NewDevice = MyDevicesLinkedList.AddDevice(UDN, URI, True, Me)
                                    If NewDevice IsNot Nothing Then
                                        If NewDevice.Location <> "" Then
                                            Dim IPInfo As IPAddressInfo
                                            IPInfo = ExtractIPInfo(URI)
                                            NewDevice.IPAddress = IPInfo.IPAddress
                                            NewDevice.IPPort = IPInfo.IPPort
                                            NewDevice.Server = Server
                                            NewDevice.BootID = BootID
                                            NewDevice.ConfigID = ConfigID
                                            NewDevice.CacheControl = CacheControl
                                            NewDevice.TimeoutValue = RetrieveTimeoutData(CacheControl)
                                            NewDevice.WakeUp = WakeUp
                                            NewDevice.SSID = SSID
                                            NewDevice.NotificationEvent = NotificationEvent
                                            If SearchPort <> "" Then NewDevice.SearchPort = SearchPort
                                            NewDevice.RootDevice = NewDevice
                                            ' NewDevice.Alive = True still need to retrieve device info
                                            If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MySSDP.TreatNotficationQueue retrieved IPInfo with IPAddress =  " & NewDevice.IPAddress & " and IPPort = " & NewDevice.IPPort & " and Location = " & NewDevice.Location, LogType.LOG_TYPE_INFO, LogColorGreen)
                                        End If
                                    Else
                                        If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.TreatNotficationQueue for device = " & NewDevice.UniqueDeviceName & ". Could not add the device to our list with Data = " & NotificationEvent, LogType.LOG_TYPE_ERROR)
                                    End If
                                    NewDevice = Nothing
                                Else
                                    Dim ExistingDevice As MyUPnPDevice = Nothing
                                    ExistingDevice = Me.Item(GetDeviceFromNT(NT, UDN), True)
                                    If ExistingDevice IsNot Nothing Then
                                        If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MySSDP.TreatNotficationQueue for ssdp:alive found matching NT = " & NT & " with UDN = " & UDN & " and URI = " & URI, LogType.LOG_TYPE_INFO, LogColorNavy)
                                        Dim IPInfo As IPAddressInfo
                                        IPInfo = ExtractIPInfo(URI)
                                        ExistingDevice.IPAddress = IPInfo.IPAddress
                                        ExistingDevice.IPPort = IPInfo.IPPort
                                        ExistingDevice.Server = Server
                                        If BootID <> ExistingDevice.BootID Then
                                            If ExistingDevice.BootID <> "" And ExistingDevice.IsRootDevice And upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MySSDP.TreatEventQueue with ssdp:alive for device = " & ExistingDevice.UniqueDeviceName & " had different bootids. Stored = " & ExistingDevice.BootID & " <-> received = " & BootID, LogType.LOG_TYPE_INFO, LogColorGreen)
                                            ExistingDevice.BootID = BootID
                                        End If
                                        'ExistingDevice.BootID = BootID
                                        ExistingDevice.ConfigID = ConfigID
                                        ExistingDevice.Location = URI
                                        ExistingDevice.CacheControl = CacheControl
                                        If Not ExistingDevice.IsRootDevice Then
                                            ' only update the root timeout ie one timer for all related devices and services added 10/13/2019
                                            If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MySSDP.TreatNotficationQueue for type device = " & ExistingDevice.UniqueDeviceName & " is updating the timeout value of its parent with Value = " & RetrieveTimeoutData(CacheControl), LogType.LOG_TYPE_INFO, LogColorGreen)
                                            Dim rootDevice As MyUPnPDevice = ExistingDevice.RootDevice
                                            If rootDevice.TimeoutValue < RetrieveTimeoutData(CacheControl) Then
                                                rootDevice.TimeoutValue = RetrieveTimeoutData(CacheControl)
                                            Else
                                                rootDevice.TimeoutValue = rootDevice.TimeoutValue ' Ok this looks strange but it resets the timer
                                            End If
                                        Else
                                            ExistingDevice.TimeoutValue = RetrieveTimeoutData(CacheControl)
                                        End If
                                        ExistingDevice.WakeUp = WakeUp
                                        ExistingDevice.SSID = SSID
                                        ExistingDevice.NotificationEvent = NotificationEvent
                                        If SearchPort <> "" Then ExistingDevice.SearchPort = SearchPort
                                        If Not ExistingDevice.Alive = True Then
                                            ExistingDevice.ForceNewDocumentRetrieval()
                                        End If
                                    End If
                                    ExistingDevice = Nothing
                                End If
                            Case MySSDP.NTtype.NTTypeURNService
                                If MyDevicesLinkedList.CheckDeviceExists(UDN) Then
                                    If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MySSDP.TreatNotficationQueue received ssdp:alive and is checking if service was known with URI = " & URI.ToString & ", NTS = " & NTS & ", NT = " & NT.ToString & ", ST = " & ST.ToString & " and USN = " & USN.ToString & " Info = " & NotificationEvent, LogType.LOG_TYPE_INFO, LogColorGreen)
                                    Dim ExistingDevice As MyUPnPDevice = Nothing
                                    ExistingDevice = Me.Item(GetDeviceFromNT(NT, UDN), True)
                                    If ExistingDevice IsNot Nothing Then
                                        If Not ExistingDevice.IsRootDevice Then
                                            ' only update the root timeout ie one timer for all related devices and services added 10/13/2019
                                            If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MySSDP.TreatNotficationQueue for a service = " & NT & " is updating the timeout value of its parent with Value = " & RetrieveTimeoutData(CacheControl), LogType.LOG_TYPE_INFO, LogColorGreen)
                                            Dim rootDevice As MyUPnPDevice = ExistingDevice.RootDevice
                                            If rootDevice.TimeoutValue < RetrieveTimeoutData(CacheControl) Then
                                                rootDevice.TimeoutValue = RetrieveTimeoutData(CacheControl)
                                            Else
                                                rootDevice.TimeoutValue = rootDevice.TimeoutValue ' Ok this looks strange but it resets the timer
                                            End If
                                        Else
                                            ExistingDevice.TimeoutValue = RetrieveTimeoutData(CacheControl)
                                        End If
                                        If ExistingDevice.Services IsNot Nothing Then
                                            Dim service As MyUPnPService = ExistingDevice.Services.GetServiceType(NT)
                                            If service IsNot Nothing Then
                                                If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MySSDP.TreatNotficationQueue received alive for Service  = " & NT, LogType.LOG_TYPE_INFO, LogColorGreen)
                                                ' record the notification event info
                                                service.NotificationEvent = NotificationEvent
                                            Else
                                                ' this is the interesting case, the service is unknown, do we add them?
                                                If upnpDebuglevel >= DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MySSDP.TreatNotficationQueue received ssdp:alive for an unknown Service with URI = " & URI.ToString & ", NTS = " & NTS & ", NT = " & NT.ToString & ", ST = " & ST.ToString & " and USN = " & USN.ToString & " Info = " & NotificationEvent, LogType.LOG_TYPE_INFO, LogColorGreen)
                                                Dim NewService As MyUPnPService = Nothing
                                                Try ' this is for Sony's DIAL service that returns no service info
                                                    NewService = New MyUPnPService(ExistingDevice.UniqueDeviceName, "", URI, ExistingDevice)
                                                Catch ex As Exception
                                                End Try
                                                If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MyUPnPDevice.ProcessDeviceInfo added new service for device = " & ExistingDevice.UniqueDeviceName & ". Service = " & NewService.Id, LogType.LOG_TYPE_INFO, LogColorNavy)
                                                If NewService IsNot Nothing Then ExistingDevice.Services.Add(NewService)
                                                NewService.ServiceTypeIdentifier = NT
                                                NewService.Id = NT
                                                NewService.NotificationEvent = NotificationEvent
                                                If upnpDebuglevel >= DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MySSDP.TreatNotficationQueue added a new service with URI = " & URI.ToString & ", NTS = " & NTS & ", NT = " & NT.ToString & ", ST = " & ST.ToString & " and USN = " & USN.ToString & " Info = " & NotificationEvent, LogType.LOG_TYPE_INFO, LogColorGreen)
                                            End If
                                        End If
                                    End If
                                Else
                                    ' there is no device known yet, this could be a timing thing or an issue. Generate a log entry
                                    ' 11/5/2019 actually if a device a child devices, you get here, may have to do something here 
                                    If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MySSDP.TreatNotficationQueue received ssdp:alive for a service but the device UDN is unknown at this point. URI = " & URI.ToString & ", NTS = " & NTS & ", NT = " & NT.ToString & ", ST = " & ST.ToString & " and USN = " & USN.ToString & " Info = " & NotificationEvent, LogType.LOG_TYPE_INFO, LogColorOrange)
                                End If
                                'End If
                        End Select
                    ElseIf Header = "NOTIFY" And NTS = "ssdp:byebye" Then
                        NbrOfNotifyByeByeMessageCounter += 1
                        If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MySSDP.TreatNotficationQueue received byebye with URI = " & URI.ToString & ", NTS = " & NTS & ", NT = " & NT.ToString & " and USN = " & USN.ToString, LogType.LOG_TYPE_INFO, LogColorGreen)
                        Dim ExistingDevice As MyUPnPDevice = Nothing
                        Select Case NTType
                            Case MySSDP.NTtype.NTTypeRootDevice
                                If UPNPMonitoringDevice = "upnp:rootdevice" Then
                                    ExistingDevice = Me.Item("uuid:" & UDN)
                                End If
                            Case MySSDP.NTtype.NTTypeDevice, MySSDP.NTtype.NTTypeURNDevice
                                If NT = UPNPMonitoringDevice Then
                                    ExistingDevice = Me.Item(GetDeviceFromNT(NT, UDN))
                                End If
                            Case MySSDP.NTtype.NTTypeURNService
                                'If upnpDebuglevel > DebugLevel.dlErrorsOnly andAlso CheckDebugParam Then Log("MySSDP.TreatNotficationQueue received byebye for a service with URI = " & URI.ToString & ", NTS = " & NTS & ", NT = " & NT.ToString & " and USN = " & USN.ToString, LogType.LOG_TYPE_WARNING)
                                'Dim ServiceID = GetServiceFromNT(NT)
                                'If upnpDebuglevel > DebugLevel.dlEvents andAlso CheckDebugParam Then Log("MySSDP.TreatNotficationQueue received byebye for a service  = " & ServiceID, LogType.LOG_TYPE_INFO)
                                ExistingDevice = Me.Item("uuid:" & UDN, True)
                                If ExistingDevice IsNot Nothing Then
                                    If ExistingDevice.Services IsNot Nothing Then
                                        Dim service As MyUPnPService = ExistingDevice.Services.GetServiceType(NT)
                                        If service IsNot Nothing Then
                                            If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MySSDP.TreatNotficationQueue received byebye for a Service = " & NT, LogType.LOG_TYPE_INFO, LogColorGreen)
                                            service.ServiceDiedReceived()
                                            Exit Try
                                        End If
                                    End If
                                End If
                        End Select
                        If ExistingDevice IsNot Nothing Then
                            ExistingDevice.Alive = False
                            ExistingDevice.TimeoutValue = -1 ' this will stop time AND generate event
                        Else
                            If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MySSDP.TreatNotficationQueue received byebye but device was not found or Not Alive with URI = " & URI.ToString & ", NTS = " & NTS & ", NT = " & NT.ToString & " and USN = " & USN.ToString, LogType.LOG_TYPE_WARNING)
                        End If
                        ExistingDevice = Nothing
                    Else
                        NbrOfUnknownMessageCounter += 1
                        If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MySSDP.TreatNotficationQueue received unknown Event with Header = " & Header & " and Event = " & NotificationEvent.ToString, LogType.LOG_TYPE_WARNING)
                    End If
                Catch ex As Exception
                    If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.TreatNotficationQueue for Event = " & NotificationEvent.ToString & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
            End While
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.TreatNotficationQueue with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MySSDP.TreatNotficationQueue is processing Notification Events T=" & NbrOfMessageCounter.ToString & " H=" & NbrOfHTTPMessageCounter.ToString & " A=" & NbrOfNotifyAliveMessageCounter & " B=" & NbrOfNotifyByeByeMessageCounter.ToString & " S=" & NbrOfSearchMessageCounter.ToString & " U=" & NbrOfUnknownMessageCounter.ToString, LogType.LOG_TYPE_INFO, LogColorGreen)
        NotificationEvent = ""
        If MissedNotificationHandlerFlag Then MyNotifyTimer.Enabled = True ' rearm the timer to prevent events from getting lost
        MissedNotificationHandlerFlag = False
        NotificationHandlerReEntryFlag = False
    End Sub

    Public Sub HandleSSDPUdpSocketCloseEvent(sender As Object)
        If isClosing Then Exit Sub
        If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.HandleSSDPUdpSocketCloseEvent received Close Socket Event", LogType.LOG_TYPE_ERROR)
        ' we close the ssdpsocket here, will be restarted with next call to DoSSDPDiscovery
        Try
            If SSDPAsyncSocket IsNot Nothing Then
                Try
                    RemoveHandler SSDPAsyncSocket.DataReceived, AddressOf HandleSSDPDataReceived
                    RemoveHandler SSDPAsyncSocket.UdpSocketClosed, AddressOf HandleSSDPUdpSocketCloseEvent
                Catch ex As Exception
                End Try
                SSDPAsyncSocket.CloseSocket()
                SSDPAsyncSocket = Nothing
            End If
        Catch ex As Exception
        End Try
        Try
            If SSDPAsyncSocketLoopback IsNot Nothing Then
                Try
                    RemoveHandler SSDPAsyncSocketLoopback.DataReceived, AddressOf HandleSSDPDataReceived
                    RemoveHandler SSDPAsyncSocketLoopback.UdpSocketClosed, AddressOf HandleSSDPUdpSocketCloseEvent
                Catch ex As Exception
                End Try
                SSDPAsyncSocketLoopback.CloseSocket()
                SSDPAsyncSocketLoopback = Nothing
            End If
        Catch ex As Exception
        End Try
    End Sub

    Public Function StartSSDPDiscovery(UPnPDeviceToLookFor As String, discoverPort As Integer) As MyUPnPDevices

        StartSSDPDiscovery = Nothing
        UPNPMonitoringDevice = UPnPDeviceToLookFor
        myDiscoverUdpPort = discoverPort

        If SSDPAsyncSocket Is Nothing Then
            Try
                SSDPAsyncSocket = New MyUdpClient(PlugInIPAddress, discoverPort) 'MCastSocket(MyUPnPMCastIPAddress, MyUPnPMCastPort) changed from 0 to 1901
            Catch ex As Exception
                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.StartSSDPDiscovery unable to create a Multicast Socket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Function
            End Try
            If SSDPAsyncSocket Is Nothing Then
                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.StartSSDPDiscovery. No AsyncSocket!", LogType.LOG_TYPE_ERROR)
                Exit Function
            End If
            AddHandler SSDPAsyncSocket.DataReceived, AddressOf HandleSSDPDataReceived
            AddHandler SSDPAsyncSocket.UdpSocketClosed, AddressOf HandleSSDPUdpSocketCloseEvent
            Try
                MySSDPClient = SSDPAsyncSocket.ConnectSocket("") 'by passing a zero string, I indicate no need to join any multicast groups. This is just a listener for SSDP discovery responses
            Catch ex As Exception
                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.StartSSDPDiscovery unable to connect Socket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Function
            End Try
            If MySSDPClient Is Nothing Then
                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.StartSSDPDiscovery. No Client!", LogType.LOG_TYPE_ERROR)
                Exit Function
            End If
            Try
                SSDPAsyncSocket.Receive()
            Catch ex As Exception
                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.StartSSDPDiscovery for UPnPDevice = " & "" & " unable to receive data to Socket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If
        If SSDPAsyncSocketLoopback Is Nothing Then ' I use now bind to IPV4 interface so maybe nomore need to call out the loopback interface
            Try
                SSDPAsyncSocketLoopback = New MyUdpClient(AnyIPv4Address, discoverPort)
            Catch ex As Exception
                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.StartSSDPDiscovery unable to create a Multicast LoopbackSocket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Function
            End Try
            If SSDPAsyncSocketLoopback Is Nothing Then
                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.StartSSDPDiscovery. No Loopback AsyncSocket!", LogType.LOG_TYPE_ERROR)
                Exit Function
            End If
            AddHandler SSDPAsyncSocketLoopback.DataReceived, AddressOf HandleSSDPDataReceived
            AddHandler SSDPAsyncSocketLoopback.UdpSocketClosed, AddressOf HandleSSDPUdpSocketCloseEvent
            Try
                MySSDPClientLoopback = SSDPAsyncSocketLoopback.ConnectSocket("")
            Catch ex As Exception
                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.StartSSDPDiscovery unable to connect Loopback Socket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Function
            End Try
            If MySSDPClientLoopback Is Nothing Then
                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.StartSSDPDiscovery. No Loopback Client!", LogType.LOG_TYPE_ERROR)
                Exit Function
            End If
            myActiveDiscoverUdpPort = SSDPAsyncSocket.LocalIPPort
            Try
                SSDPAsyncSocketLoopback.Receive()
            Catch ex As Exception
                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.StartSSDPDiscovery for UPnPDevice = " & "" & " unable to receive data to Loopback Socket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If


        SSDPAsyncSocket.response = ""
        Dim postData As String = ""
        postData = "M-SEARCH * HTTP/1.1" & vbCrLf
        'postData &= "HOST: 239.255.255.250:1900" & vbCrLf  ' changed from abusing a Multicast address to use unicast on 10/20/2019
        postData &= "HOST: 239.255.255.250:" & SSDPAsyncSocket.LocalIPPort.ToString & vbCrLf
        postData &= "ST: " & UPNPMonitoringDevice & vbCrLf 'upnp:rootdevice" & vbCrLf 'UPnPDeviceToLookFor & vbCrLf  ' upnp:rootdevice ' always look for the root device!
        postData &= "MAN: ""ssdp:discover""" & vbCrLf
        postData &= "MX: 4" & vbCrLf
        postData &= "" & vbCrLf

        ' Broadcast the message to the listener.
        Dim StartTime As DateTime = DateTime.Now
        If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("MySSDP.StartSSDPDiscovery, waiting for 9 seconds while discovering UPnP Devices ..... ", LogType.LOG_TYPE_INFO, LogColorGreen)

        If SSDPAsyncSocket IsNot Nothing Then SSDPAsyncSocket.Send(postData, MyUPnPMCastIPAddress, MyUPnPMCastPort)
        If SSDPAsyncSocket IsNot Nothing Then SSDPAsyncSocket.Send(postData, "255.255.255.255", MyUPnPMCastPort)        ' added to try
        If SSDPAsyncSocketLoopback IsNot Nothing Then SSDPAsyncSocketLoopback.Send(postData, MyUPnPMCastIPAddress, MyUPnPMCastPort)
        wait(9)

        Dim elapsed_time As TimeSpan = DateTime.Now.Subtract(StartTime)

        While MyNotificationQueue.Count > 0 And elapsed_time.TotalSeconds < 60
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("MySSDP.StartSSDPDiscovery waiting for an additional 5 seconds while discovering UPnP Devices. Still " & MyNotificationQueue.Count.ToString & " Entries awaiting processing ..... ", LogType.LOG_TYPE_INFO, LogColorGreen)
            wait(5)
            elapsed_time = DateTime.Now.Subtract(StartTime)
        End While

        'StartEventListener() ' is already started in myssdp.CreateMulticastListener
        ' moved this code here on 2/26/2019 because there is 30 seconds in which new devices were found but not added NOR are they evented to the PI as new devices
        If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("MySSDP.StartSSDPDiscovery found so far " & MyDevicesLinkedList.Count.ToString & " Devices. Still " & MyNotificationQueue.Count.ToString & " Entries awaiting processing ..... ", LogType.LOG_TYPE_INFO, LogColorGreen)
        If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("MySSDP.StartSSDPDiscovery waiting for an additional 30 seconds while receiving and processing the UPnP Device Info ..... ", LogType.LOG_TYPE_INFO, LogColorGreen)
        wait(30) ' wait for processing of the XML Documents

        Dim DiscoveryList As New MyUPnPDevices
        For Each Device As MyUPnPDevice In MyDevicesLinkedList
            If Device IsNot Nothing Then
                If Device.Alive Then
                    If (Device.Type = UPnPDeviceToLookFor) Or (UPnPDeviceToLookFor = "upnp:rootdevice") Then ' if we are looking with root, return all but NOT children!
                        DiscoveryList.Add(Device)
                    End If
                    If Device.HasChildren And Not (UPnPDeviceToLookFor = "upnp:rootdevice") Then ' I only do one level of embedded devices :-(
                        For Each child As MyUPnPDevice In Device.Children
                            If child IsNot Nothing Then
                                If child.Type = UPnPDeviceToLookFor Then
                                    DiscoveryList.Add(child)
                                End If
                            End If
                        Next
                    End If
                End If
            End If
        Next

        StartSSDPDiscovery = DiscoveryList
        'If UPnPDebuglevel > DebugLevel.dlOff andAlso CheckDebugParam Then Log("MySSDP.StartSSDPDiscovery found so far " & DiscoveryList.Count.ToString & " Devices. Still " & MyNotificationQueue.Count.ToString & " Entries awaiting processing ..... ", LogType.LOG_TYPE_INFO, LogColorGreen)

        Try
            If SSDPAsyncSocket IsNot Nothing Then
                Try
                    RemoveHandler SSDPAsyncSocket.DataReceived, AddressOf HandleSSDPDataReceived
                    RemoveHandler SSDPAsyncSocket.UdpSocketClosed, AddressOf HandleSSDPUdpSocketCloseEvent
                Catch ex As Exception
                End Try
                SSDPAsyncSocket.CloseSocket()
                SSDPAsyncSocket = Nothing
            End If
        Catch ex As Exception
        End Try
        Try
            If SSDPAsyncSocketLoopback IsNot Nothing Then
                Try
                    RemoveHandler SSDPAsyncSocketLoopback.DataReceived, AddressOf HandleSSDPDataReceived
                    RemoveHandler SSDPAsyncSocketLoopback.UdpSocketClosed, AddressOf HandleSSDPUdpSocketCloseEvent
                Catch ex As Exception
                End Try
                SSDPAsyncSocketLoopback.CloseSocket()
                SSDPAsyncSocketLoopback = Nothing
            End If
        Catch ex As Exception
        End Try
        'If UPnPDebuglevel > DebugLevel.dlOff andAlso CheckDebugParam Then Log("MySSDP.StartSSDPDiscovery waiting for an additional 30 seconds while receiving and processing the UPnP Device Info ..... ", LogType.LOG_TYPE_INFO, LogColorGreen)
        'wait(30) ' wait for processing of the XML Documents
        If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("MySSDP.StartSSDPDiscovery found " & DiscoveryList.Count.ToString & " Devices. Still " & MyNotificationQueue.Count.ToString & " Entries awaiting processing ..... ", LogType.LOG_TYPE_INFO, LogColorGreen)

    End Function

    Public Sub SendMSearch()
        If SSDPAsyncSocket Is Nothing Then
            Try
                SSDPAsyncSocket = New MyUdpClient(PlugInIPAddress, myDiscoverUdpPort)
            Catch ex As Exception
                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.SendMSearch unable to create a Multicast Socket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Sub
            End Try
            If SSDPAsyncSocket Is Nothing Then
                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.SendMSearch. No AsyncSocket!", LogType.LOG_TYPE_ERROR)
                Exit Sub
            End If
            AddHandler SSDPAsyncSocket.DataReceived, AddressOf HandleSSDPDataReceived
            Try
                MySSDPClient = SSDPAsyncSocket.ConnectSocket("") 'by passing a zero string, I indicate no need to join any multicast groups. This is just a listener for SSDP discovery responses
            Catch ex As Exception
                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.SendMSearch unable to connect Socket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Sub
            End Try
            If MySSDPClient Is Nothing Then
                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.SendMSearch. No Client!", LogType.LOG_TYPE_ERROR)
                Exit Sub
            End If
            myActiveDiscoverUdpPort = SSDPAsyncSocket.LocalIPPort
            Try
                SSDPAsyncSocket.Receive()
            Catch ex As Exception
                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.SendMSearch for UPnPDevice = " & "" & " unable to receive data to Socket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If

        SSDPAsyncSocket.response = ""
        Dim postData As String = ""
        postData = "M-SEARCH * HTTP/1.1" & vbCrLf
        postData &= "HOST: 239.255.255.250:1900" & vbCrLf
        postData &= "ST: " & UPNPMonitoringDevice & vbCrLf
        postData &= "MAN: ""ssdp:discover""" & vbCrLf
        postData &= "MX: 4" & vbCrLf
        postData &= "" & vbCrLf
        If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MySSDP.SendMSearch Sending M-Search for " & UPNPMonitoringDevice, LogType.LOG_TYPE_INFO, LogColorGreen)
        If SSDPAsyncSocket IsNot Nothing Then SSDPAsyncSocket.Send(postData, MyUPnPMCastIPAddress, MyUPnPMCastPort)
        'wait(30)

        Try
            If SSDPAsyncSocket IsNot Nothing Then
                Try
                    RemoveHandler SSDPAsyncSocket.DataReceived, AddressOf HandleSSDPDataReceived
                    RemoveHandler SSDPAsyncSocket.UdpSocketClosed, AddressOf HandleSSDPUdpSocketCloseEvent
                Catch ex As Exception
                End Try
                SSDPAsyncSocket.CloseSocket()
                SSDPAsyncSocket = Nothing
            End If
        Catch ex As Exception
        End Try

    End Sub

    Public Sub SendUSearch(port As String)
        ' send a Unicast search
        Try
            Dim localIPAddr As IPAddress = IPAddress.Parse(PlugInIPAddress)
            Dim postData As String = ""
            postData = "M-SEARCH * HTTP/1.1" & vbCrLf
            postData &= "HOST: " & localIPAddr.ToString & ":" & port & vbCrLf
            postData &= "ST: " & UPNPMonitoringDevice & vbCrLf
            postData &= "MAN: ""ssdp:discover""" & vbCrLf
            postData &= "" & vbCrLf
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("MySSDP.SendUSearch Sending M-Search for " & UPNPMonitoringDevice, LogType.LOG_TYPE_INFO, LogColorGreen)
            If SSDPAsyncSocket IsNot Nothing Then SSDPAsyncSocket.Send(postData, MyUPnPMCastIPAddress, MyUPnPMCastPort)
            If SSDPAsyncSocketLoopback IsNot Nothing Then SSDPAsyncSocketLoopback.Send(postData, MyUPnPMCastIPAddress, MyUPnPMCastPort)
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("MySSDP.SendUSearch Sending M-Search for " & UPNPMonitoringDevice & " but ended in Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Function ParseSsdpResponse(response As String, SearchItem As String) As String
        ParseSsdpResponse = ""
        If SearchItem <> "" Then
            Dim FuncReturn As String = (From line In response.Split({vbCr(0), vbLf(0)})
                                        Where line.ToLowerInvariant().StartsWith(SearchItem & ":")
                                        Select (line.Substring(SearchItem.Length + 1).Trim())).FirstOrDefault
            If FuncReturn IsNot Nothing Then
                Return FuncReturn
            Else
                Return ""
            End If
        Else
            ' Return the first line, which should be M-Search, Notify, HTTP
            Dim Lines As String() = Split(response, {vbCr(0), vbLf(0)})
            If Lines IsNot Nothing Then
                If Lines(0).ToLowerInvariant().StartsWith("m-search * ") Then
                    Return "M-SEARCH"
                ElseIf Lines(0).ToLowerInvariant().StartsWith("http/") Then
                    Return "HTTP"
                ElseIf Lines(0).ToLowerInvariant().StartsWith("notify * ") Then
                    Return "NOTIFY"
                End If
            End If
        End If
    End Function

    Private Function RetrieveTimeoutData(inData As String) As Integer
        RetrieveTimeoutData = 1800
        Dim Timeout As Integer = 1800
        Try
            Dim TimeoutParms As String() = Split(inData, "=") ' first part is Seconds ... or should be!! Second part could be an integer or "infinite"
            If UBound(TimeoutParms) <= 0 Then
                If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MySSDP.RetrieveTimeoutData received an Invalid TIMEOUT = " & inData, LogType.LOG_TYPE_ERROR)
                Exit Function
            End If
            If Trim(TimeoutParms(0).ToUpper) <> "MAX-AGE" Then
                If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MySSDP.RetrieveTimeoutData received an Invalid TIMEOUT = " & inData, LogType.LOG_TYPE_ERROR)
                Exit Function
            End If
            Timeout = Val(Trim(TimeoutParms(1)))
            If Timeout = 0 Then
                If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MySSDP.RetrieveTimeoutData received an invalid TIMEOUT = " & inData, LogType.LOG_TYPE_ERROR)
                Exit Function
            End If
            If Timeout < 1800 Then
                If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MySSDP.RetrieveTimeoutData received a very small timeout request, which could case a lot of network traffic. Received = " & inData, LogType.LOG_TYPE_WARNING)
            End If
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.RetrieveTimeoutData. Unable to extract the timeout info with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        Return Timeout
    End Function

    Private Sub MyNotifyTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles MyNotifyTimer.Elapsed
        Try
            TreatNotficationQueue()
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.MyNotifyTimer_Elapsed with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        e = Nothing
        sender = Nothing
    End Sub

    Public Function GetAllDevices() As MyUPnPDevices
        GetAllDevices = MyDevicesLinkedList
    End Function

    Private Function ParseNT(inNT As String) As NTtype
        If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MySSDP.ParseNT received inNT = " & inNT.ToString, LogType.LOG_TYPE_INFO, LogColorGreen)
        If inNT = "" Then Return NTtype.NTTypeUnknown
        Dim NTParts As String() = Split(inNT, ":")
        Try
            Select Case NTParts(0).ToUpper
                Case "UPNP"
                    If NTParts(1).ToUpper = "ROOTDEVICE" Then Return NTtype.NTTypeRootDevice
                Case "UUID"
                    Return NTtype.NTTypeDevice
                Case "URN"
                    Select Case NTParts(2).ToUpper
                        Case "DEVICE"
                            Return NTtype.NTTypeURNDevice
                        Case "SERVICE"
                            Return NTtype.NTTypeURNService
                    End Select
            End Select
        Catch ex As Exception
        End Try
        Return NTtype.NTTypeUnknown
    End Function

    Private Function GetDeviceFromNT(inNT As String, inUDN As String) As String
        If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MySSDP.GetDeviceFromNT received inNT = " & inNT.ToString & " and inUDN = " & inUDN, LogType.LOG_TYPE_INFO, LogColorGreen)
        If inNT = "" Then Return ""
        Dim NTParts As String() = Split(inNT, ":")
        Try
            Select Case NTParts(0).ToUpper
                Case "UUID"
                    Return inNT
                Case "URN"
                    Select Case NTParts(2).ToUpper
                        Case "DEVICE"
                            If inUDN <> "" Then
                                Return "uuid:" & inUDN
                            End If
                        Case "SERVICE"  ' added this to capture the Sonos Smartspeaker service info. May have to remove this again
                            If inUDN <> "" Then
                                Return "uuid:" & inUDN
                            End If
                    End Select
            End Select
        Catch ex As Exception
        End Try
        Return ""
    End Function

    Private Function GetServiceFromNT(inNT As String) As String
        If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MySSDP.GetServiceFromNT received inNT = " & inNT.ToString, LogType.LOG_TYPE_INFO, LogColorGreen)
        If inNT = "" Then Return ""
        Dim NTParts As String() = Split(inNT, ":")
        Try
            Select Case NTParts(0).ToUpper
                Case "URN"
                    Select Case NTParts(2).ToUpper
                        Case "SERVICE"
                            Return NTParts(3).ToString
                    End Select
            End Select
        Catch ex As Exception
        End Try
        Return ""
    End Function

    Public Sub HandleNewDeviceWasProcessed(UDN As String)
        If UDN.IndexOf("uuid:") = 0 Then
            Mid(UDN, 1, 5) = "     "
            UDN = Trim(UDN)
        End If
        Try
            RaiseEvent NewDeviceFound(UDN)
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.HandleNewDeviceWasProcessed with UDN = " & UDN & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Enum NTtype
        NTTypeRootDevice = 0
        NTTypeDevice = 1
        NTTypeURNDevice = 2
        NTTypeURNService = 3
        NTTypeUnknown = 4
    End Enum

#End Region

#Region "Notifications"

    Overloads ReadOnly Property Item(bstrUDN As String, Optional OverrideAliveFlag As Boolean = False) As MyUPnPDevice
        Get
            Item = Nothing
            If MyDevicesLinkedList IsNot Nothing Then
                If MyDevicesLinkedList.Count > 0 Then
                    Try
                        For Each Device As MyUPnPDevice In MyDevicesLinkedList
                            If Device IsNot Nothing Then
                                'Log("found = " & Device.UniqueDeviceName & " looking for " & bstrUDN, LogType.LOG_TYPE_WARNING)
                                If Device.UniqueDeviceName = bstrUDN And (Device.Alive Or OverrideAliveFlag) Then
                                    Item = Device
                                    Exit Property
                                ElseIf Device.HasChildren Then
                                    If Device.Children IsNot Nothing Then
                                        Dim ChildItem As MyUPnPDevice = Device.Children.Item(bstrUDN, OverrideAliveFlag)
                                        If ChildItem IsNot Nothing Then
                                            Item = ChildItem
                                            Exit Property
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    Catch ex As Exception
                        If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.Item retrieving a Device with UDN = " & bstrUDN & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                End If
            End If
        End Get
    End Property

    Public Sub HandleDataReceive(Data As String)
        MyEventListener.TCPResponse = "HTTP/1.1 200 OK" & vbCrLf & "Connection: close" & vbCrLf & "Content-Length: 0" & vbCrLf & vbCrLf
        MyEventListener.sendDone.Set()
        If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MySSDP.HandleDataReceive called with Length = " & Data.Length & " and Data = " & Data, LogType.LOG_TYPE_INFO, LogColorGreen)
        Dim NotifyData As String = ""
        Try
            NotifyData = System.Web.HttpUtility.UrlDecode(ParseHTTPResponse(Data, "notify"))
            If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MySSDP.HandleDataReceive called with NotifyData = " & NotifyData, LogType.LOG_TYPE_INFO, LogColorGreen) ' & " and NT = " & NTReceived & " and SID = " & SID, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.HandleDataReceive while analyzing data with error = " & ex.Message & " and Data = " & Data, LogType.LOG_TYPE_ERROR)
        End Try

        If upnpDebuglevel > DebugLevel.dlEvents Then
            Log("MySSDP.HandleDataReceive found NOTIFY = " & NotifyData, LogType.LOG_TYPE_INFO, LogColorGreen)
            Log("MySSDP.HandleDataReceive found NTS = " & ParseHTTPResponse(Data, "nts:"), LogType.LOG_TYPE_INFO, LogColorGreen)
            Log("MySSDP.HandleDataReceive found HOST = " & ParseHTTPResponse(Data, "host:"), LogType.LOG_TYPE_INFO, LogColorGreen)
            Log("MySSDP.HandleDataReceive found CONTENT-TYPE = " & ParseHTTPResponse(Data, "content-type:"), LogType.LOG_TYPE_INFO, LogColorGreen)
            Log("MySSDP.HandleDataReceive found CONTENT-LENGTH = " & ParseHTTPResponse(Data, "content-length:"), LogType.LOG_TYPE_INFO, LogColorGreen)
            Log("MySSDP.HandleDataReceive found CACHE-CONTROL = " & ParseHTTPResponse(Data, "cache-control:"), LogType.LOG_TYPE_INFO, LogColorGreen)
        End If

        Dim UDN As String = ""
        Dim ServiceId As String = ""

        NotifyData = Trim(NotifyData)
        If NotifyData = "" Then
            If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("Error in MySSDP.HandleDataReceive the Notify Data is empty and Data = " & Data, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End If

        ' now we need to take the Notify info apart, it looks like this: /RINCON_000E5859008A01400_MR/urn:upnp-org:serviceId:AVTransport HTTP/1.1
        Dim NotifyParts As String() = Split(NotifyData, " ")

        Try
            If UBound(NotifyParts, 1) > 0 Then
                Dim ServiceParts As String()
                ServiceParts = Split(NotifyParts(0), "/")
                If UBound(ServiceParts, 1) > 1 Then
                    UDN = "uuid:" & ServiceParts(1)
                    ServiceId = ServiceParts(2)
                End If
            End If
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.HandleDataReceive trying to process the Notify Data = " & NotifyData & " and Error = " & ex.Message & " and Data = " & Data, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try

        If UDN = "" Or ServiceId = "" Then
            If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("Error in MySSDP.HandleDataReceive UDN or ServiceID are empty with Notify Data = " & NotifyData & " and Data = " & Data, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End If

        Try
            ' Now go find the service this belongs to
            If MyDevicesLinkedList IsNot Nothing Then
                If MyDevicesLinkedList.Count > 0 Then
                    For Each Device As MyUPnPDevice In MyDevicesLinkedList
                        If Device IsNot Nothing Then
                            If Device.UniqueDeviceName = UDN Then
                                Try
                                    ' go retrieve the service
                                    If Device.Services IsNot Nothing Then
                                        Dim MyService As MyUPnPService = Device.Services.Item(ServiceId)
                                        If MyService IsNot Nothing Then
                                            MyService.HandleUPNPDataReceived(Data)
                                            Data = ""
                                            Exit Sub
                                        End If
                                    End If
                                Catch ex As Exception
                                    If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.HandleDataReceive (2) trying to find the Service with SID = " & ServiceId & " and Error = " & ex.Message & " and Data = " & Data, LogType.LOG_TYPE_INFO)
                                End Try
                            End If
                            If Device.HasChildren Then
                                For Each Child As MyUPnPDevice In Device.Children
                                    If Child IsNot Nothing Then
                                        If Child.UniqueDeviceName = UDN Then
                                            Try
                                                ' go retrieve the service
                                                If Child.Services IsNot Nothing Then
                                                    Dim MyService As MyUPnPService = Child.Services.Item(ServiceId)
                                                    If MyService IsNot Nothing Then
                                                        MyService.HandleUPNPDataReceived(Data)
                                                        Data = ""
                                                        Exit Sub
                                                    End If
                                                End If
                                            Catch ex As Exception
                                                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.HandleDataReceive (3) trying to find the Service with SID = " & ServiceId & " and Error = " & ex.Message & " and Data = " & Data, LogType.LOG_TYPE_INFO)
                                            End Try
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.HandleDataReceive trying to find the Service with SID = " & ServiceId & " and Error = " & ex.Message & " and Data = " & Data, LogType.LOG_TYPE_INFO)
        End Try
        If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("Error in MySSDP.HandleDataReceive Service with SID = " & ServiceId & " was not found and Data = " & Data, LogType.LOG_TYPE_INFO)
        Data = ""
    End Sub

    Private Function StartEventListener(port As Integer) As Boolean
        ' this procedure listens to the event Notifications on my own proprietary port 12291/12292
        If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MySSDP.StartEventListener called with port = " & port, LogType.LOG_TYPE_INFO, LogColorGreen)
        Dim returnPort As Integer = 0
        If MyEventListener Is Nothing Then
            Try
                MyEventListener = New MyTCPListener
                returnPort = MyEventListener.Start(PlugInIPAddress, port)
                If returnPort = 0 Then
                    ' this is an unsuccesfull start of a listener
                    StopListener()
                    RestartListenerFlag = True
                    Return False
                End If
                TCPListenerPort = returnPort
                AddHandler MyEventListener.Connection, AddressOf HandleEventConnection
                AddHandler MyEventListener.recOK, AddressOf HandleEventReceive
                AddHandler MyEventListener.DataReceived, AddressOf HandleDataReceive
            Catch ex As Exception
                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.StartEventListener on interface = " & PlugInIPAddress & " and port = " & port & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                StopListener()
                RestartListenerFlag = True
                Return False
            End Try
        Else
            Try
                returnPort = MyEventListener.Start(PlugInIPAddress, port)
                If returnPort = 0 Then
                    ' this is an unsuccesfull start of a listener
                    StopListener()
                    RestartListenerFlag = True
                    Return False
                End If
                TCPListenerPort = returnPort
            Catch ex As Exception
                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.StartEventListener on interface = " & PlugInIPAddress & " and port = " & port & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                StopListener()
                RestartListenerFlag = True
                Return False
            End Try
        End If
        If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("MySSDP.StartEventListener started listening on interface = " & PlugInIPAddress & " and port = " & TCPListenerPort, LogType.LOG_TYPE_INFO)
        Return True
    End Function

    Private Sub StopListener()
        If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MySSDP.StopListener called for Instance = " & MainInstance, LogType.LOG_TYPE_INFO, LogColorGreen)
        Try
            If MyEventListener IsNot Nothing Then
                Try
                    RemoveHandler MyEventListener.Connection, AddressOf HandleEventConnection
                    RemoveHandler MyEventListener.recOK, AddressOf HandleEventReceive
                    RemoveHandler MyEventListener.DataReceived, AddressOf HandleDataReceive
                Catch ex As Exception
                End Try
                MyEventListener.Close()
                MyEventListener = Nothing
            End If
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.StopListener for Instance = " & MainInstance & " stopping the EventListener with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub HandleEventConnection(Connection As Boolean)
        If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MySSDP.HandleEventConnection called with Connection = " & Connection, LogType.LOG_TYPE_INFO, LogColorGreen)
        If Connection Then
            RestartListenerFlag = False
            ListenerIsActive = True
        Else
            'StopListener()  ' added 10/13/2019 when this event is received, setting the timer flag will call startlistener BUT it will check on whether myEventListener is nothing and if not won't do anything
            RestartListenerFlag = True
            ListenerIsActive = False
        End If
    End Sub

    Public Sub HandleEventReceive(Receive As Boolean)
        If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MySSDP.HandleEventReceive called with Receive = " & Receive, LogType.LOG_TYPE_INFO, LogColorGreen)
    End Sub

#End Region

    Private Sub MyControllerTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles MyControllerTimer.Elapsed
        Try
            If RestartListenerFlag Then
                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("MySSDP.MyControllerTimer_Elapsed called with restart the Listener", LogType.LOG_TYPE_INFO, LogColorGreen) ' added 7/13/2019 in v3.1.0.34
                If TCPListenerPort > 0 Then TCPListenerPort += 1    ' if need to retry, take next port if specific or stay with port 0 to dynamically pick a port 10/19/2019
                StartEventListener(TCPListenerPort) ' changed in v19 StartEventListener()
            End If
            If RestartMulticastListenerFlag Then
                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("MySSDP.MyControllerTimer_Elapsed called with restart the MulicastListener", LogType.LOG_TYPE_INFO, LogColorGreen) ' added 7/13/2019 in v3.1.0.34
                CreateMulticastListener()
            End If
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MySSDP.MyControllerTimer_Elapsed with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        e = Nothing
        sender = Nothing
    End Sub

    Private Sub MyAuditTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles myAuditTimer.Elapsed

        If Not IsEthernetPortAlive(PlugInIPAddress) Then
            If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("Error in MySSDP.MyAuditTimer_Elapsed!! The Ethernet interface provide by HS = " & PlugInIPAddress & " is down !!!", LogType.LOG_TYPE_INFO, LogColorGreen)
        End If

        ' added 10/14/2019 to deal with multicast listener going out to lunch
        If MulticastAsyncSocket IsNot Nothing Then
            Try
                If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MySSDP.MyAuditTimer_Elapsed called with MC listener stats: bytesReceeived = " & MulticastAsyncSocket.BytesReceived.ToString, LogType.LOG_TYPE_INFO, LogColorGreen)
                If MulticastAsyncSocket.BytesReceived = overallBytesReceivedMulticast Then
                    isTheSameCounter += 1 ' increase counter
                    If isTheSameCounter > 6 Then
                        ' this has taken too long, let's kick the socket
                        MulticastAsyncSocket.CloseSocket()
                        MyMulticastClient = MulticastAsyncSocket.ConnectSocket(MyUPnPMCastIPAddress)
                        If MyMulticastClient Is Nothing Then
                            ' remove and retry from beginning!      ' added 10/21/2019

                            HandleMCastSocketCloseEvent(Nothing)
                            overallBytesReceivedMulticast = 0
                            isTheSameCounter = 0
                            e = Nothing
                            sender = Nothing
                            Exit Sub
                        End If
                        MulticastAsyncSocket.Receive()
                        ' reset counters
                        MulticastAsyncSocket.BytesReceived = 0
                        overallBytesReceivedMulticast = MulticastAsyncSocket.BytesReceived
                    End If
                Else
                    overallBytesReceivedMulticast = MulticastAsyncSocket.BytesReceived
                    isTheSameCounter = 0
                End If
                SendMSearch()
            Catch ex As Exception
                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("MySSDP.MyAuditTimer_Elapsed called and had error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If
        e = Nothing
        sender = Nothing
    End Sub


End Class

#Region "Devices"

<Serializable()>
Public Class MyUPnPDevice
    Public UniqueDeviceName As String = ""
    Public Location As String = ""
    Public Description As String = ""
    Public FriendlyName As String = ""
    Public ManufacturerName As String = ""
    Public ManufacturerURL As String = ""
    Public ModelName As String = ""
    Public ModelNumber As String = ""
    Public ModelURL As String = ""
    Public ParentDevice As MyUPnPDevice = Nothing
    Public RootDevice As MyUPnPDevice = Nothing
    Public PresentationURL As String = ""
    Public Type As String = ""
    Public UPC As String = ""
    Private MyIconXML As System.Xml.XmlNodeList = Nothing
    Public ApplicationURL As String = ""
    Public IsRootDevice As Boolean = False
    Public IPAddress As String = ""
    Public IPPort As String = ""
    Public Children As MyUPnPDevices
    Public HasChildren As Boolean = False
    Public Services As MyUPnPServices = Nothing
    Private AliveInLastScan As Boolean = True
    Public Server As String = ""
    Public BootID As String = ""
    Public ConfigID As String = ""
    Public SearchPort As String = "1900"
    Public CacheControl As String = ""
    Public DeviceUPnPDocument As String = ""
    Private NewDeviceFlag As Boolean = True
    Private ReferenceToSSDP As MySSDP
    'Private LocationPath As String = "" ' this is only the path part of the URL
    Private MyTimeOutValue As Integer = 0
    Public WakeUp As String = ""
    Public SSID As String = ""
    Public NotificationEvent As String = ""
    Private retrieveDeviceInfoReentrancyFlag As Boolean = False

    Friend WithEvents MySubscribeRenewTimer As Timers.Timer
    Friend WithEvents MyRetrieveDeviceInfoTimer As Timers.Timer
    Friend WithEvents MyEventTimer As Timers.Timer

    Public Delegate Sub DeviceDiedEventHandler()
    Public Event DeviceDied As DeviceDiedEventHandler
    Public Delegate Sub DeviceAliveEventHandler()           ' added in v10
    Public Event DeviceAlive As DeviceAliveEventHandler     ' added in v10
    Private DeviceDiedHandlerAddress As DeviceDiedEventHandler
    Private DeviceAliveHandlerAddress As DeviceAliveEventHandler
    Private debugParam As String


    Public ReadOnly Property CheckDebugParam As Boolean
        Get
            If debugParam = "" Then Return True
            Return (debugParam = UniqueDeviceName)
        End Get
    End Property
    Public Property Alive As Boolean
        Get
            Alive = AliveInLastScan
        End Get
        Set(value As Boolean)
            If value And (value <> AliveInLastScan) Then
                AliveInLastScan = True ' needs to be set before raising the event
                RaiseEvent DeviceAlive() ' added in v10
            End If
            AliveInLastScan = value
        End Set
    End Property

    Public Property TimeoutValue As Integer
        Get
            TimeoutValue = MyTimeOutValue
        End Get
        Set(value As Integer)
            If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MyUPnPDevice.TimeoutValue called for UDN = " & UniqueDeviceName & " with value = " & value.ToString, LogType.LOG_TYPE_INFO, LogColorGreen)
            If value = -1 Then ' this is used to indicate that the device has send a ssdp:byebye
                Try
                    If MySubscribeRenewTimer IsNot Nothing Then
                        MySubscribeRenewTimer.Enabled = False
                    End If
                    If MyEventTimer Is Nothing Then
                        MyEventTimer = New Timers.Timer
                    End If
                    MyEventTimer.Stop()
                    MyEventTimer.Interval = 100 ' 100 milliseconds
                    MyEventTimer.AutoReset = False
                    MyEventTimer.Enabled = True
                    MyEventTimer.Start()
                Catch ex As Exception
                    If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPDevice.TimeoutValue for UDN = " & UniqueDeviceName & ". Unable to create the Event timer with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
            Else
                MyTimeOutValue = value
                Try
                    If MySubscribeRenewTimer Is Nothing Then
                        MySubscribeRenewTimer = New Timers.Timer
                    End If
                    MySubscribeRenewTimer.Stop()
                    MySubscribeRenewTimer.Interval = value * 1000 * 2 '+ 5000 ' add 5 seconds to max time. If by then no renew is received, the device is dead 
                    MySubscribeRenewTimer.AutoReset = False
                    MySubscribeRenewTimer.Enabled = True
                    MySubscribeRenewTimer.Start()
                    If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MyUPnPDevice.TimeoutValue for UDN = " & UniqueDeviceName & " New Timeout = " & MySubscribeRenewTimer.Interval.ToString, LogType.LOG_TYPE_INFO)
                Catch ex As Exception
                    If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPDevice.TimeoutValue for UDN = " & UniqueDeviceName & ". Unable to create the Subscribe timer with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
            End If
        End Set
    End Property

    Public Sub New(inStrLocation As String, DoDocumentRetrieval As Boolean, inRef As MySSDP)
        MyBase.New()
        If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MyUPnPDevice.New called for UDN = " & UniqueDeviceName & " with inStrLocation = " & inStrLocation.ToString & " and DoDocumentRetrieval = " & DoDocumentRetrieval.ToString, LogType.LOG_TYPE_INFO, LogColorGreen)
        NewDeviceFlag = True
        ReferenceToSSDP = inRef
        UniqueDeviceName = ""
        ApplicationURL = ""
        Location = inStrLocation
        Description = ""
        FriendlyName = ""
        ManufacturerName = ""
        ManufacturerURL = ""
        ModelName = ""
        ModelNumber = ""
        ModelURL = ""
        ParentDevice = Nothing
        RootDevice = Nothing
        PresentationURL = ""
        Type = ""
        UPC = ""
        MyIconXML = Nothing
        IsRootDevice = False
        IPAddress = ""
        IPPort = ""
        Children = Nothing
        HasChildren = False
        Services = Nothing
        MyTimeOutValue = 1800 'TimeoutValue = 1800
        Server = ""
        BootID = ""
        ConfigID = ""
        SearchPort = "1900"
        DeviceUPnPDocument = ""
        WakeUp = ""
        SSID = ""
        If DoDocumentRetrieval Then
            Randomize()
            ' start a time to make this an asynchronous operation. Time is between 100 and 600msec.
            ' Generate random value between 100 and 600. changed this on Jan 4th 2016 to random of 1.6 sec and base of 5 sec so between 5 sec and 6.6 sec
            Dim value As Integer = CInt(Int((1600 * Rnd()) + 5000))
            'Dim value As Integer = CInt(Int((600 * Rnd()) + 100))
            Try
                MyRetrieveDeviceInfoTimer = New Timers.Timer With {
                    .Interval = value, ' in msecond
                    .AutoReset = False,
                    .Enabled = True
                }
            Catch ex As Exception
                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPDevice.New. Unable to create the RetrieveDeviceInfo timer with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If
    End Sub

    Public Sub AddHandlers(ParentClass As Object) 'HSPI)
        Try
            DeviceDiedHandlerAddress = AddressOf ParentClass.DeviceLostCallback
            DeviceAliveHandlerAddress = AddressOf ParentClass.DeviceAliveCallBack
            AddHandler DeviceDied, DeviceDiedHandlerAddress
            AddHandler DeviceAlive, DeviceAliveHandlerAddress
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPDevice.AddHandlers with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub ForceNewDocumentRetrieval()
        Try
            If MyRetrieveDeviceInfoTimer IsNot Nothing Then
                MyRetrieveDeviceInfoTimer.Enabled = False
                MyRetrieveDeviceInfoTimer.Stop()
                MyRetrieveDeviceInfoTimer.Dispose()
                MyRetrieveDeviceInfoTimer = Nothing
            End If
        Catch ex As Exception
        End Try
        Randomize()
        ' start a time to make this an asynchronous operation. Time is between 100 and 600msec.
        ' Generate random value between 100 and 600. 
        Dim value As Integer = CInt(Int((600 * Rnd()) + 100))
        Try
            MyRetrieveDeviceInfoTimer = New Timers.Timer With {
                .Interval = value, ' in msecond
                .AutoReset = False,
                .Enabled = True
            }
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPDevice.ForceNewDocumentRetrieval. Unable to create the RetrieveDeviceInfo timer with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub Dispose(RemoveSelf As Boolean)
        If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPDevice.Dispose called for UDN = " & UniqueDeviceName & " and RemoveSelf = " & RemoveSelf.ToString, LogType.LOG_TYPE_INFO, LogColorGreen)
        AliveInLastScan = False
        Location = ""
        ApplicationURL = ""
        Description = ""
        FriendlyName = ""
        ManufacturerName = ""
        ManufacturerURL = ""
        ModelName = ""
        ModelNumber = ""
        ModelURL = ""
        ParentDevice = Nothing
        RootDevice = Nothing
        PresentationURL = ""
        Type = ""
        UPC = ""
        MyIconXML = Nothing
        IPAddress = ""
        IPPort = ""
        Server = ""
        BootID = ""
        ConfigID = ""
        SearchPort = ""
        CacheControl = ""
        DeviceUPnPDocument = ""
        WakeUp = ""
        SSID = ""
        NotificationEvent = ""
        Try
            If DeviceDiedHandlerAddress IsNot Nothing Then RemoveHandler DeviceDied, DeviceDiedHandlerAddress
            If DeviceAliveHandlerAddress IsNot Nothing Then RemoveHandler DeviceAlive, DeviceAliveHandlerAddress
        Catch ex As Exception
        End Try
        DeviceDiedHandlerAddress = Nothing
        DeviceAliveHandlerAddress = Nothing
        StopTimers()
        Try
            If Children IsNot Nothing Then
                Children.Dispose()
            End If
        Catch ex As Exception
        End Try
        Children = Nothing
        HasChildren = False
        Try
            If Services IsNot Nothing Then
                Services.Dispose()
            End If
        Catch ex As Exception
        End Try
        Services = Nothing
        If ReferenceToSSDP IsNot Nothing And RemoveSelf Then
            Try
                ReferenceToSSDP.RemoveDevices(UniqueDeviceName)
            Catch ex As Exception
            End Try
        End If
        UniqueDeviceName = ""
        IsRootDevice = False
        ReferenceToSSDP = Nothing
    End Sub

    Public Sub ReleaseDeviceServiceInfo()
        If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPDevice.ReleaseDeviceServiceInfo for for UDN = " & UniqueDeviceName, LogType.LOG_TYPE_INFO, LogColorGreen)
        AliveInLastScan = False
        'ParentDevice = Nothing
        'RootDevice = Nothing
        StopTimers()
        Try
            If Children IsNot Nothing Then
                Children.Dispose()
            End If
        Catch ex As Exception
        End Try
        Children = Nothing
        HasChildren = False
        Try
            If Services IsNot Nothing Then
                Services.Dispose()
            End If
        Catch ex As Exception
        End Try
        Services = Nothing
    End Sub

    Private Sub StopTimers()
        Try
            If MySubscribeRenewTimer IsNot Nothing Then
                MySubscribeRenewTimer.Enabled = False
                MySubscribeRenewTimer.Stop()
                MySubscribeRenewTimer.Dispose()

            End If
        Catch ex As Exception
        End Try
        MySubscribeRenewTimer = Nothing
        Try
            If MyRetrieveDeviceInfoTimer IsNot Nothing Then
                MyRetrieveDeviceInfoTimer.Enabled = False
                MyRetrieveDeviceInfoTimer.Stop()
                MyRetrieveDeviceInfoTimer.Dispose()

            End If
        Catch ex As Exception
        End Try
        MyRetrieveDeviceInfoTimer = Nothing
    End Sub

    Public Function CheckForUpdates(inLocation As String, inUDN As String, inRef As MySSDP) As Boolean
        CheckForUpdates = False
        If Location = inLocation Then Exit Function
        If inLocation <> "" And Location <> "" Then
            Try
                Dim NewUri As New Uri(inLocation)
                Dim ExistingURI As New Uri(Location)
                Dim NewIPaddress As IPAddress = Nothing
                System.Net.IPAddress.TryParse(NewUri.Host, NewIPaddress)
                Dim ExistingIPaddress As IPAddress = Nothing
                System.Net.IPAddress.TryParse(ExistingURI.Host, ExistingIPaddress)
                Dim HostIPaddress As IPAddress = Nothing
                System.Net.IPAddress.TryParse(PlugInIPAddress, HostIPaddress)
                Dim LoopbackIPaddress As IPAddress = Nothing
                System.Net.IPAddress.TryParse(LoopBackIPv4Address, LoopbackIPaddress)
                If (NewIPaddress.Equals(LoopbackIPaddress) And ExistingIPaddress.Equals(HostIPaddress)) Or (NewIPaddress.Equals(HostIPaddress) And ExistingIPaddress.Equals(LoopbackIPaddress)) Then
                    ' this may be the same, just different IP address
                    If (NewUri.Port = ExistingURI.Port) And (NewUri.PathAndQuery = ExistingURI.PathAndQuery) Then
                        ' they are the same !!
                        If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPDevice.CheckForUpdates received a Loopback versus network IP address for an existing device = " & UniqueDeviceName & " and existing Location = " & Location & " and new Location = " & inLocation, LogType.LOG_TYPE_INFO, LogColorGreen)
                        'Update Location, else it may keep on spitting out warnings
                        Location = inLocation
                        NewUri = Nothing
                        ExistingURI = Nothing
                        NewIPaddress = Nothing
                        ExistingIPaddress = Nothing
                        HostIPaddress = Nothing
                        LoopbackIPaddress = Nothing
                        Exit Function
                    End If
                End If
                NewUri = Nothing
                ExistingURI = Nothing
                NewIPaddress = Nothing
                ExistingIPaddress = Nothing
                HostIPaddress = Nothing
                LoopbackIPaddress = Nothing
            Catch ex As Exception
                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPDevice.CheckForUpdates checking IP addresses for Loopback addresses an existing device = " & UniqueDeviceName & " and existing Location = " & Location & " and new Location = " & inLocation & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If
        If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPDevice.CheckForUpdates received a new location for an existing device = " & UniqueDeviceName & " and existing Location = " & Location & ", new Location = " & inLocation & ", new UDN = " & inUDN, LogType.LOG_TYPE_INFO, LogColorOrange)

        Try
            ReleaseDeviceServiceInfo()
        Catch ex As Exception
        End Try
        Try
            RaiseEvent DeviceDied()
        Catch ex As Exception
        End Try

        ReferenceToSSDP = inRef
        Location = inLocation
        UniqueDeviceName = inUDN
        Dim IPInfo As IPAddressInfo
        IPInfo = ExtractIPInfo(inLocation)
        IPAddress = IPInfo.IPAddress
        IPPort = IPInfo.IPPort
        IPInfo = Nothing
        MyTimeOutValue = 1800 'TimeoutValue = 1800
        SearchPort = "1900"
        Randomize()
        ' start a time to make this an asynchronous operation. Time is between 100 and 600msec.
        ' Generate random value between 100 and 600. 
        Dim value As Integer = CInt(Int((600 * Rnd()) + 100))
        Try
            MyRetrieveDeviceInfoTimer = New Timers.Timer With {
                .Interval = value, ' in msecond
                .AutoReset = False,
                .Enabled = True
            }
            'If UPnPDebuglevel > DebugLevel.dlErrorsOnly andAlso CheckDebugParam Then Log("MyUPnPDevice.CheckForUpdates started new RetrieveDeviceInfo timer with Interval = " & MyRetrieveDeviceInfoTimer.Interval, LogType.LOG_TYPE_WARNING)
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPDevice.CheckForUpdates. Unable to create the RetrieveDeviceInfo timer with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Private Sub RetrieveDeviceInfo()
        If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MyUPnPDevice.RetrieveDeviceInfo called for device = " & UniqueDeviceName & " with location = " & Location & " and reentrancy = " & retrieveDeviceInfoReentrancyFlag.ToString, LogType.LOG_TYPE_INFO, LogColorGreen)
        Dim PageHTML As String = ""
        If Location = "" Then
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPDevice.RetrieveDeviceInfo for device = " & UniqueDeviceName & " the location is empty", LogType.LOG_TYPE_ERROR)
            Exit Sub
        End If
        If retrieveDeviceInfoReentrancyFlag Then Exit Sub  ' added 11/12/2019
        retrieveDeviceInfoReentrancyFlag = True
        Dim beforeTime As DateTime = DateTime.Now

        Dim location_ As String = Location
        Try
            Dim RequestUri = New Uri(Location)
            'Dim p = ServicePointManager.FindServicePoint(RequestUri) ' removed 11/10/2019
            'p.Expect100Continue = False
            Dim webRequest As HttpWebRequest = DirectCast(System.Net.HttpWebRequest.Create(RequestUri), HttpWebRequest) ' HttpWebRequest.Create(RequestUri)
            webRequest.Method = "GET"
            webRequest.KeepAlive = False
            webRequest.ContentLength = 0
            webRequest.Timeout = 20000 ' 5000 ' set to max of 5 seconds changed on 9/10/2019 
            Using webResponse As WebResponse = webRequest.GetResponse   ' changed this on 2/2/2020 because I have these timeouts on Linux
                Dim afterTime As DateTime = DateTime.Now
                Dim deltaTime As TimeSpan = afterTime.Subtract(beforeTime)
                Try
                    ApplicationURL = webResponse.Headers("Application-URL")
                    If ApplicationURL <> "" Then
                        If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPDevice.RetrieveDeviceInfo called for device = " & UniqueDeviceName & " with location = " & location_ & " after " & deltaTime.TotalMilliseconds.ToString & " milliseconds retrieved Application-URL = " & ApplicationURL, LogType.LOG_TYPE_INFO, LogColorGreen)
                    End If
                Catch ex As Exception
                End Try
                Using webStream As Stream = webResponse.GetResponseStream
                    Using strmRdr As New System.IO.StreamReader(webStream)
                        PageHTML = strmRdr.ReadToEnd()
                        strmRdr.Close()
                        strmRdr.Dispose()
                        webStream.Close()
                        webResponse.Close()
                        webStream.Dispose()
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Dim afterTime As DateTime = DateTime.Now
            Dim deltaTime As TimeSpan = afterTime.Subtract(beforeTime)
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPDevice.RetrieveDeviceInfo for device = " & UniqueDeviceName & " after " & deltaTime.TotalMilliseconds.ToString & " milliseconds while retrieving document with URL = " & location_ & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            AliveInLastScan = False
            Location = "" ' reset the location which will force a rediscovery when the next ssdp:alive event is received
            retrieveDeviceInfoReentrancyFlag = False
            Exit Sub
        End Try
        PageHTML = RemoveControlCharacters(PageHTML)
        If PageHTML.ToUpper.Trim(" ") = "STATUS=OK" Then ' added 04/15/2018 
            AliveInLastScan = True
            FriendlyName = "Unknown"
            retrieveDeviceInfoReentrancyFlag = False
            Exit Sub
        End If
        If ProcessDeviceInfo(Me, PageHTML, True) Then
            AliveInLastScan = True
        Else
            If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("Warning in MyUPnPDevice.RetrieveDeviceInfo for device = " & UniqueDeviceName & " had unsuccessful ProcessDeviceInfo and reset the Location. HTML = " & PageHTML, LogType.LOG_TYPE_WARNING)
            Location = "" ' reset the location which will force a rediscovery when the next ssdp:alive event is received
            AliveInLastScan = False
        End If
        If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MyUPnPDevice.RetrieveDeviceInfo done for device = " & UniqueDeviceName & " with location = " & location_, LogType.LOG_TYPE_INFO, LogColorGreen)
        retrieveDeviceInfoReentrancyFlag = False
    End Sub

    Private Function ProcessDeviceInfo(ByRef NewDevice As MyUPnPDevice, inXML As String, IsRoot As Boolean) As Boolean
        ProcessDeviceInfo = False
        Dim xmlDoc As New XmlDocument
        inXML = Trim(inXML)
        If inXML = "" Then Exit Function
        xmlDoc.XmlResolver = Nothing
        Try
            xmlDoc.LoadXml(inXML)
            If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MyUPnPDevice.ProcessDeviceInfo for device = " & NewDevice.UniqueDeviceName & " and IsRoot = " & IsRoot.ToString & " retrieved following document = " & xmlDoc.OuterXml.ToString, LogType.LOG_TYPE_INFO, LogColorGreen)
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPDevice.ProcessDeviceInfo for device = " & NewDevice.UniqueDeviceName & " while loading the device XML with XML = " & inXML & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        DeviceUPnPDocument = inXML
        'nodename 	Selects all nodes with the name "nodename"
        '/ 	        Selects from the root node
        '// 	    Selects nodes in the document from the current node that match the selection no matter where they are
        '. 	        Selects the current node
        '.. 	    Selects the parent of the current node
        '@ 	        Selects attributes
        Dim DeviceNode_ As System.Xml.XmlNode = Nothing

        Dim DeviceNodeList As System.Xml.XmlNodeList = Nothing
        Dim DeviceListXML As String = ""
        Try
            For Each DeviceNode_ In xmlDoc.ChildNodes
                'Log("ProcessDeviceInfo for device = " & NewDevice.UniqueDeviceName & "Local Name = " & DeviceNode_.LocalName.ToString, LogType.LOG_TYPE_WARNING)
                'Log("ProcessDeviceInfo for device = " & NewDevice.UniqueDeviceName & "Name = " & DeviceNode_.Name.ToString, LogType.LOG_TYPE_WARNING)
                If DeviceNode_.Name = "root" And IsRoot Then
                    For Each Child As System.Xml.XmlNode In DeviceNode_.ChildNodes
                        'Log("ProcessDeviceInfo for device = " & NewDevice.UniqueDeviceName & "Local Child Name = " & Child.LocalName.ToString, LogType.LOG_TYPE_WARNING)
                        'Log("ProcessDeviceInfo for device = " & NewDevice.UniqueDeviceName & "Child Name = " & Child.Name.ToString, LogType.LOG_TYPE_WARNING)
                        If Child.Name = "device" Then
                            For Each GrandChild As System.Xml.XmlNode In Child.ChildNodes
                                'Log("ProcessDeviceInfo for device = " & NewDevice.UniqueDeviceName & "Local GrandChild Name = " & GrandChild.LocalName.ToString, LogType.LOG_TYPE_WARNING)
                                'Log("ProcessDeviceInfo for device = " & NewDevice.UniqueDeviceName & "GrandChild Name = " & GrandChild.Name.ToString, LogType.LOG_TYPE_WARNING)
                                If GrandChild.Name = "deviceList" Then
                                    DeviceListXML = GrandChild.OuterXml
                                    Child.RemoveChild(GrandChild)
                                    Dim RootChildXML As String = Child.OuterXml
                                    If RootChildXML <> "" Then
                                        Dim RootChildXMLDoc As New XmlDocument
                                        RootChildXMLDoc.LoadXml(RootChildXML)
                                        DeviceNodeList = RootChildXMLDoc.GetElementsByTagName("device")
                                        If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPDevice.ProcessDeviceInfo for device = " & NewDevice.UniqueDeviceName & " found " & DeviceNodeList.Count & " Root devices", LogType.LOG_TYPE_INFO, LogColorGreen)
                                    End If
                                    Exit Try
                                End If
                            Next
                        End If
                    Next
                    DeviceNodeList = xmlDoc.GetElementsByTagName("device")
                    If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPDevice.ProcessDeviceInfo for device = " & NewDevice.UniqueDeviceName & " found " & DeviceNodeList.Count & " Root devices", LogType.LOG_TYPE_INFO, LogColorGreen)
                    Exit Try
                ElseIf DeviceNode_.Name = "device" And Not IsRoot Then
                    DeviceNodeList = xmlDoc.GetElementsByTagName("device")
                    If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPDevice.ProcessDeviceInfo for device = " & NewDevice.UniqueDeviceName & " found " & DeviceNodeList.Count & " Root Child devices", LogType.LOG_TYPE_INFO, LogColorGreen)
                    Exit Try
                End If
            Next
            DeviceNodeList = xmlDoc.GetElementsByTagName("device")
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("MyUPnPDevice.ProcessDeviceInfo for device = " & NewDevice.UniqueDeviceName & " found " & DeviceNodeList.Count & " ?? devices", LogType.LOG_TYPE_WARNING)
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPDevice.ProcessDeviceInfo for device = " & NewDevice.UniqueDeviceName & " while retrieving the device from XML = " & inXML & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        Try
            Dim DeviceXML As New XmlDocument
            For Each DeviceNode As System.Xml.XmlNode In DeviceNodeList
                'Log("ProcessDeviceInfo for device = " & NewDevice.UniqueDeviceName & "Parent Name = " & DeviceNode.ParentNode.Name, LogType.LOG_TYPE_INFO)
                'Log("ProcessDeviceInfo for device = " & NewDevice.UniqueDeviceName & "My Name = " & DeviceNode.Name, LogType.LOG_TYPE_INFO)
                Try
                    DeviceXML.LoadXml(DeviceNode.OuterXml)
                    Try
                        NewDevice.Type = If(DeviceXML.GetElementsByTagName("deviceType").Item(0)?.InnerText, "")
                    Catch ex As Exception
                    End Try
                    Try
                        NewDevice.FriendlyName = If(DeviceXML.GetElementsByTagName("friendlyName").Item(0)?.InnerText, "")
                        If IsRoot Then
                            If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPDevice.ProcessDeviceInfo retrieved FriendlyName " & NewDevice.FriendlyName, LogType.LOG_TYPE_INFO, LogColorGreen)
                        Else
                            If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPDevice.ProcessDeviceInfo retrieved FriendlyName for ChildDevice " & NewDevice.FriendlyName, LogType.LOG_TYPE_INFO, LogColorGreen)
                        End If
                    Catch ex As Exception
                    End Try
                    Try
                        NewDevice.ManufacturerName = If(DeviceXML.GetElementsByTagName("manufacturer").Item(0)?.InnerText, "")
                    Catch ex As Exception
                    End Try
                    Try
                        NewDevice.ManufacturerURL = If(DeviceXML.GetElementsByTagName("manufacturerURL").Item(0)?.InnerText, "")
                    Catch ex As Exception
                    End Try
                    Try
                        NewDevice.ModelNumber = If(DeviceXML.GetElementsByTagName("modelNumber").Item(0)?.InnerText, "")
                    Catch ex As Exception
                    End Try
                    Try
                        NewDevice.ModelName = If(DeviceXML.GetElementsByTagName("modelName").Item(0)?.InnerText, "")
                    Catch ex As Exception
                    End Try
                    Try
                        NewDevice.ModelURL = If(DeviceXML.GetElementsByTagName("modelURL").Item(0)?.InnerText, "")
                    Catch ex As Exception
                    End Try
                    Try
                        If IsRoot Then
                            NewDevice.IsRootDevice = True
                        Else
                            NewDevice.IsRootDevice = False
                        End If
                        Dim NewUDN As String = If(DeviceXML.GetElementsByTagName("UDN").Item(0)?.InnerText, "")
                        If NewDevice.UniqueDeviceName <> "" And NewDevice.UniqueDeviceName <> NewUDN Then
                            If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPDevice.ProcessDeviceInfo for device = " & NewDevice.UniqueDeviceName & " has different UDN in device document UDN = " & NewUDN, LogType.LOG_TYPE_WARNING)
                        End If
                        NewDevice.UniqueDeviceName = NewUDN
                    Catch ex As Exception
                    End Try
                    Try
                        NewDevice.IconURL = DeviceXML.GetElementsByTagName("iconList")
                    Catch ex As Exception
                    End Try
                    Try
                        Dim ServicesList As System.Xml.XmlNodeList
                        ServicesList = DeviceXML.GetElementsByTagName("serviceList")
                        If ServicesList IsNot Nothing Then
                            If ServicesList.Count > 0 Then
                                For Each ServiceList As System.Xml.XmlNode In ServicesList
                                    Dim ServiceListXML As New XmlDocument
                                    ServiceListXML.LoadXml(ServiceList.OuterXml)
                                    Dim ListofServices As System.Xml.XmlNodeList = Nothing
                                    ListofServices = ServiceListXML.GetElementsByTagName("service")
                                    If ListofServices IsNot Nothing Then
                                        If ListofServices.Count > 0 Then
                                            If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MyUPnPDevice.ProcessDeviceInfo for device = " & NewDevice.UniqueDeviceName & " found " & ListofServices.Count.ToString & " Services in the ServiceList", LogType.LOG_TYPE_INFO)
                                            For Each NewServiceNode As System.Xml.XmlNode In ListofServices
                                                Try
                                                    If Trim(NewServiceNode.OuterXml) <> "" Then
                                                        If NewDevice.Services Is Nothing Then
                                                            NewDevice.Services = New MyUPnPServices
                                                        End If
                                                        Dim StartTime As DateTime
                                                        Dim elapsed_time As TimeSpan
                                                        If upnpDebuglevel > DebugLevel.dlEvents Then StartTime = DateTime.Now
                                                        Dim NewService As MyUPnPService = Nothing
                                                        Try ' this is for Sony's DIAL service that returns no service info
                                                            NewService = New MyUPnPService(NewDevice.UniqueDeviceName, NewServiceNode.OuterXml, NewDevice.Location, Me.RootDevice)
                                                        Catch ex As Exception
                                                        End Try
                                                        If upnpDebuglevel > DebugLevel.dlEvents Then
                                                            elapsed_time = DateTime.Now.Subtract(StartTime)
                                                            Log("MyUPnPDevice.ProcessDeviceInfo added new service for device = " & NewDevice.UniqueDeviceName & ". Service = " & NewService.Id & " Required = " & elapsed_time.Milliseconds.ToString & " milliseconds", LogType.LOG_TYPE_INFO, LogColorNavy)
                                                        End If
                                                        If NewService IsNot Nothing Then NewDevice.Services.Add(NewService)
                                                    End If
                                                Catch ex As Exception
                                                    If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPDevice.ProcessDeviceInfo for device = " & NewDevice.UniqueDeviceName & " processing Service XML with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                                    Exit Function
                                                End Try
                                            Next
                                        End If
                                    End If
                                    'Exit For ' only process the first one
                                Next
                            End If
                        End If
                    Catch ex As Exception
                        If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPDevice.ProcessDeviceInfo for device = " & NewDevice.UniqueDeviceName & " loading Service XML with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                        Exit Function
                    End Try
                Catch ex As Exception
                    If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPDevice.ProcessDeviceInfo for device = " & NewDevice.UniqueDeviceName & " loading device XML with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    Exit Function
                End Try
                'If IsRoot Then Exit For
            Next
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPDevice.ProcessDeviceInfo for device = " & NewDevice.UniqueDeviceName & " while processing the device node list with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try

        ' now check for Child Devices
        If DeviceListXML <> "" Then
            Dim DeviceListNodeDoc As New XmlDocument
            Try
                DeviceListNodeDoc.LoadXml(DeviceListXML)
            Catch ex As Exception
                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPDevice.ProcessDeviceInfo for device = " & NewDevice.UniqueDeviceName & " loading the Child deviceList XML = " & DeviceListXML & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Function
            End Try
            Try
                Dim ListofDevices As System.Xml.XmlNodeList = Nothing
                ListofDevices = DeviceListNodeDoc.GetElementsByTagName("device")
                If ListofDevices IsNot Nothing Then
                    If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPDevice.ProcessDeviceInfo for device = " & NewDevice.UniqueDeviceName & " found " & ListofDevices.Count.ToString & " child devices in the devicelist", LogType.LOG_TYPE_INFO)
                    If ListofDevices.Count > 0 Then
                        For Each NewChildDeviceNode As System.Xml.XmlNode In ListofDevices
                            If Trim(NewChildDeviceNode.OuterXml) <> "" Then
                                Dim NewChildDeviceXMLDocument As New XmlDocument
                                NewChildDeviceXMLDocument.LoadXml(NewChildDeviceNode.OuterXml)
                                Dim ChildUDN As String = If(NewChildDeviceXMLDocument.GetElementsByTagName("UDN").Item(0)?.InnerText, "")
                                If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPDevice.ProcessDeviceInfo for device = " & NewDevice.UniqueDeviceName & " found a child device in the devicelist with UDN = " & ChildUDN, LogType.LOG_TYPE_INFO)
                                Dim NewChildDevice As MyUPnPDevice
                                If NewDevice.Children Is Nothing Then
                                    NewDevice.Children = New MyUPnPDevices
                                End If
                                If Not NewDevice.Children.CheckDeviceExists(ChildUDN) Then
                                    NewChildDevice = NewDevice.Children.AddDevice(ChildUDN, NewDevice.Location, False, ReferenceToSSDP)
                                    NewChildDevice.ParentDevice = NewDevice
                                    NewChildDevice.RootDevice = NewDevice.RootDevice
                                    NewChildDevice.IPAddress = NewDevice.IPAddress
                                    NewChildDevice.IPPort = NewDevice.IPPort
                                    NewChildDevice.Location = NewDevice.Location
                                    If Not NewChildDevice.ProcessDeviceInfo(NewChildDevice, If(NewChildDeviceXMLDocument.GetElementsByTagName("device").Item(0)?.OuterXml, ""), False) Then Exit Function
                                End If
                                NewDevice.HasChildren = True
                            End If
                        Next
                    End If
                End If
            Catch ex As Exception
                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPDevice.ProcessDeviceInfo for device = " & NewDevice.UniqueDeviceName & " loading Child device XML with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Function
            End Try
        End If
        If NewDeviceFlag And (UniqueDeviceName <> "") And IsRoot Then ReferenceToSSDP.HandleNewDeviceWasProcessed(UniqueDeviceName)
        If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPDevice.ProcessDeviceInfo for device = " & NewDevice.UniqueDeviceName & " is done with NewDeviceFlag = " & NewDeviceFlag.ToString & " and IsRoot = " & IsRoot.ToString, LogType.LOG_TYPE_INFO, LogColorGreen)
        Return True
    End Function

    Public Sub OneChildDeviceDied(ChildUDN As String)
        If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPDevice.OneChildDeviceDied called for UDN = " & UniqueDeviceName, LogType.LOG_TYPE_WARNING)
        'If UPnPDebuglevel > DebugLevel.dlOff andAlso CheckDebugParam Then Log("MyUPnPDevice.OneChildDeviceDied called for UDN = " & UniqueDeviceName & " with ChildUDN = " & ChildUDN, LogType.LOG_TYPE_WARNING)
        Dim ChildDevice As MyUPnPDevice = Children.Item(ChildUDN, True)
        If ChildDevice IsNot Nothing Then
            Try
                Children.RemoveDevice(ChildDevice, ChildUDN)
                If Children IsNot Nothing Then
                    If Children.Count = 0 Then
                        Children.Clear()
                        Children = Nothing
                    End If
                End If
            Catch ex As Exception
            End Try
        End If
    End Sub

    Public Sub SomePartOfDeviceDied(Child As Boolean, ChildUDN As String)
        If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("MyUPnPDevice.SomePartOfDeviceDied called for UDN = " & UniqueDeviceName & " and IP = " & IPAddress & " and Name = " & FriendlyName & " and ChildFlag = " & Child.ToString & " with ChildUDN = " & ChildUDN, LogType.LOG_TYPE_WARNING)
        ' device is dead
        AliveInLastScan = False
        Try
            ReleaseDeviceServiceInfo()
        Catch ex As Exception
        End Try
        If Not IsRootDevice Then
            ' inform parent device
            Try
                If ParentDevice IsNot Nothing Then
                    ParentDevice.SomePartOfDeviceDied(IsRootDevice, UniqueDeviceName)
                    ParentDevice = Nothing
                End If
            Catch ex As Exception
            End Try
        End If
        'Try
        'Dispose(True)
        'Catch ex As Exception
        'End Try
        Try
            'If UPnPDebuglevel > DebugLevel.dlOff andAlso CheckDebugParam Then Log("MyUPnPDevice.SomePartOfDeviceDied called for UDN = " & UniqueDeviceName & " and start Event DeviceDied being raised", LogType.LOG_TYPE_WARNING)
            RaiseEvent DeviceDied()
            'If UPnPDebuglevel > DebugLevel.dlOff andAlso CheckDebugParam Then Log("MyUPnPDevice.SomePartOfDeviceDied called for UDN = " & UniqueDeviceName & " and returned from Event DeviceDied", LogType.LOG_TYPE_WARNING)
        Catch ex As Exception
        End Try

        ' changed 12/2/2018 do this after sending an event else the handle is gone v.039
        Try
            Dispose(True)
        Catch ex As Exception
        End Try



    End Sub

    Public ReadOnly Property IconURL(EncodeingFormat As String, SizeX As Integer, SizeY As Integer, BitDepth As Integer) As String
        Get
            IconURL = ""
            Dim FirstURL As String = ""
            If MyIconXML Is Nothing Then Exit Property
            If MyIconXML.Count > 0 Then
                Dim IConXML As New XmlDocument
                For Each Icon As System.Xml.XmlNode In MyIconXML
                    Try
                        If Trim(Icon.OuterXml) <> "" Then
                            IConXML.LoadXml(Icon.OuterXml)
                            Dim mimeType As String = If(IConXML.GetElementsByTagName("mimetype").Item(0)?.InnerText, "")
                            Dim Width As String = If(IConXML.GetElementsByTagName("width").Item(0)?.InnerText, "")
                            Dim Height As String = If(IConXML.GetElementsByTagName("height").Item(0)?.InnerText, "")
                            Dim depth As String = If(IConXML.GetElementsByTagName("depth").Item(0)?.InnerText, "")
                            Dim URL As String = If(IConXML.GetElementsByTagName("url").Item(0)?.InnerText, "")
                            If FirstURL = "" Then FirstURL = URL
                            If mimeType = EncodeingFormat And Width = SizeX And Height = SizeY And depth = BitDepth Then
                                IconURL = MakeURLWhole(URL)
                                'If upnpDebuglevel > DebugLevel.dlErrorsOnly andAlso CheckDebugParam Then Log("IconURL got IconURL = " & "http://" & IPAddress & ":" & IPPort & URL, LogType.LOG_TYPE_INFO)
                                Exit Property
                            End If
                            'If upnpDebuglevel > DebugLevel.dlErrorsOnly andAlso CheckDebugParam Then Log("IconURL didn't match Icon with EncodeingFormat = " & mimeType & ", SizeX = " & Width.ToString & ", SizeY = " & Height.ToString & ", BitDepth = " & depth.ToString, LogType.LOG_TYPE_WARNING)
                        End If
                    Catch ex As Exception
                        If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in get MyUPnPDevice.IconURL for UDN = " & UniqueDeviceName & " for EncodeingFormat = " & EncodeingFormat & ", SizeX = " & SizeX.ToString & ", SizeY = " & SizeY.ToString & ", BitDepth = " & BitDepth.ToString & "with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                Next
                IConXML = Nothing
            Else
                If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPDevice.IconURL for UDN = " & UniqueDeviceName & " has no Icons. Was called with EncodeingFormat = " & EncodeingFormat & ", SizeX = " & SizeX.ToString & ", SizeY = " & SizeY.ToString & ", BitDepth = " & BitDepth.ToString, LogType.LOG_TYPE_WARNING)
            End If
            If FirstURL <> "" Then
                IconURL = MakeURLWhole(FirstURL)
            End If
        End Get
    End Property

    Private Function MakeURLWhole(inURL As String) As String
        MakeURLWhole = inURL
        ' the inURL can either be complete, meaning it starts with http: urn: file: ... or it is partial and then it SHOULD start with an / so we complete it with IP address and IPPort
        inURL = Trim(inURL) ' remove blanks
        If inURL = "" Then
            If IPPort <> "" Then Return "http://" & IPAddress & ":" & IPPort Else Return "http://" & IPAddress
        End If
        Try
            Dim FullUri As New Uri(inURL)
            If FullUri.IsAbsoluteUri And Trim(FullUri.Host) <> "" Then
                If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPDevice.MakeURLWhole for UDN = " & UniqueDeviceName & " for inURL = " & inURL & " was complete with Host = " & FullUri.Host.ToString & " and type = " & FullUri.HostNameType.ToString, LogType.LOG_TYPE_INFO, LogColorGreen)
                FullUri = Nothing
                Return inURL
            End If
        Catch ex As Exception
            ' if it comes here, it is not a FullURI
        End Try
        Try
            Dim FullUri As New Uri(Location)
            Dim newURI As New Uri(FullUri, inURL)
            If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MyUPnPDevice.MakeURLWhole for UDN = " & UniqueDeviceName & " for inURL = " & inURL & " returned = " & newURI.AbsoluteUri, LogType.LOG_TYPE_INFO, LogColorGreen)
            If newURI.IsAbsoluteUri Then Return newURI.AbsoluteUri
        Catch ex As Exception
            ' if it comes here, it is not a FullURI
        End Try
        Try
            If inURL.IndexOf("/") <> 0 Then
                ' add the "/" character
                If IPPort <> "" Then Return "http://" & IPAddress & ":" & IPPort & "/" & inURL Else Return "http://" & IPAddress & "/" & inURL
            Else
                ' this URL started with the right / character
                If IPPort <> "" Then Return "http://" & IPAddress & ":" & IPPort & inURL Else Return "http://" & IPAddress & inURL
            End If
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPDevice.MakeURLWhole for UDN = " & UniqueDeviceName & " for inURL = " & inURL & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public WriteOnly Property IconURL As System.Xml.XmlNodeList
        Set(value As System.Xml.XmlNodeList)
            MyIconXML = value
        End Set
    End Property

    Private Sub MySubscribeRenewTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles MySubscribeRenewTimer.Elapsed
        If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPDevice.MySubscribeRenewTimer_Elapsed called for UDN = " & UniqueDeviceName & " and IP = " & IPAddress & " and Name = " & FriendlyName, LogType.LOG_TYPE_WARNING)
        'If UPnPDebuglevel > DebugLevel.dlOff andAlso CheckDebugParam Then Log("MyUPnPDevice.MySubscribeRenewTimer_Elapsed called for UDN = " & UniqueDeviceName & " and IP = " & IPAddress & " and Name = " & FriendlyName, LogType.LOG_TYPE_WARNING)
        SomePartOfDeviceDied(Not IsRootDevice, UniqueDeviceName)
        e = Nothing
        sender = Nothing
    End Sub

    Private Sub MyRetrieveDeviceInfoTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles MyRetrieveDeviceInfoTimer.Elapsed
        If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MyUPnPDevice.MyRetrieveDeviceInfoTimer_Elapsed called for UDN = " & UniqueDeviceName, LogType.LOG_TYPE_INFO, LogColorGreen)
        Try
            RetrieveDeviceInfo()
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPDevice.MyRetrieveDeviceInfoTimer_Elapsed called for UDN = " & UniqueDeviceName & " with Error = " & ex.Message, LogType.LOG_TYPE_INFO)
        End Try
        e = Nothing
        sender = Nothing
    End Sub

    Private Sub MyEventTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles MyEventTimer.Elapsed
        If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPDevice.MyEventTimer_Elapsed called for UDN = " & UniqueDeviceName, LogType.LOG_TYPE_WARNING)
        'If UPnPDebuglevel > DebugLevel.dlOff andAlso CheckDebugParam Then Log("MyUPnPDevice.MyEventTimer_Elapsed called for UDN = " & UniqueDeviceName, LogType.LOG_TYPE_WARNING)
        ' device is dead
        AliveInLastScan = False
        Try
            ReleaseDeviceServiceInfo()
        Catch ex As Exception
        End Try
        'Try
        'Dispose(True)
        'Catch ex As Exception
        'End Try

        'sequence changed to see if that makes a difference for Makinava
        Try
            'If UPnPDebuglevel > DebugLevel.dlOff andAlso CheckDebugParam Then Log("MyUPnPDevice.MyEventTimer_Elapsed called for UDN = " & UniqueDeviceName & " and start Event DeviceDied being raised", LogType.LOG_TYPE_WARNING)
            RaiseEvent DeviceDied()
            'If UPnPDebuglevel > DebugLevel.dlOff andAlso CheckDebugParam Then Log("MyUPnPDevice.MyEventTimer_Elapsed called for UDN = " & UniqueDeviceName & " and returned from Event DeviceDied", LogType.LOG_TYPE_WARNING)
        Catch ex As Exception
        End Try

        ' changed 12/2/2018 do this after sending an event else the handle is gone v.039
        Try
            Dispose(True)
        Catch ex As Exception
        End Try



        e = Nothing
        sender = Nothing
    End Sub
End Class



<Serializable()>
Public Class MyUPnPDevices

    Inherits List(Of MyUPnPDevice)

    Public Function AddDevice(UDN As String, Location As String, DoDocumentRetrieval As Boolean, inRef As MySSDP) As MyUPnPDevice
        If upnpDebuglevel > DebugLevel.dlEvents Then Log("MyUPnPDevices.AddDevice called with UDN = " & UDN & " and Location = " & Location & ". Devices list count = " & MyBase.Count.ToString, LogType.LOG_TYPE_INFO, LogColorGreen)

        AddDevice = Nothing
        If UDN = "" Then
            If upnpDebuglevel > DebugLevel.dlOff Then Log("Error in MyUPnPDevices.AddDevice when called with No UDN = " & UDN & " and Location = " & Location, LogType.LOG_TYPE_ERROR)
            Exit Function
        End If
        Try
            Dim Device As New MyUPnPDevice(Location, DoDocumentRetrieval, inRef)
            If UDN.IndexOf("uuid:") = -1 Or UDN.IndexOf("uuid:") <> 0 Then
                Device.UniqueDeviceName = "uuid:" & UDN
            End If
            MyBase.Add(Device)
            If upnpDebuglevel > DebugLevel.dlErrorsOnly Then Log("MyUPnPDevices.AddDevice added UDN = " & UDN & " and Location = " & Location & " to its device list. Devices list count = " & MyBase.Count.ToString, LogType.LOG_TYPE_INFO, LogColorGreen)
            AddDevice = Device
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff Then Log("Error in MyUPnPDevices.AddDevice when called with UDN = " & UDN & " and Location = " & Location & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function RemoveDevice(pDevice As MyUPnPDevice, UDN As String) As Boolean
        RemoveDevice = False
        Try
            If Not MyBase.Remove(pDevice) Then
                If upnpDebuglevel > DebugLevel.dlOff Then Log("Error in MyUPnPDevices.RemoveDevice for UDN = " & UDN & ". Unsuccessful removal ", LogType.LOG_TYPE_ERROR)
            Else
                If upnpDebuglevel > DebugLevel.dlErrorsOnly Then Log("MyUPnPDevices.RemoveDevice removed UDN = " & UDN & " successfully", LogType.LOG_TYPE_WARNING)
                Try
                    pDevice.Dispose(True)
                Catch ex As Exception
                End Try
                RemoveDevice = True
            End If
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff Then Log("Error in MyUPnPDevices.RemoveDevice for UDN = " & UDN & ". Unsuccessful removal with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Overloads ReadOnly Property Item(bstrUDN As String, Optional OverrideAliveFlag As Boolean = False) As MyUPnPDevice
        Get
            Item = Nothing
            If Me.Count > 0 Then
                For Each device As MyUPnPDevice In Me
                    If device IsNot Nothing Then
                        'If UPnPDebuglevel > DebugLevel.dlOff Then Log("MyUPnPDevices.Item called with UDN = " & bstrUDN & ", found = " & device.UniqueDeviceName & " and OverrideFlag = " & OverrideAliveFlag.ToString & " and DeviceAlive = " & device.Alive.ToString, LogType.LOG_TYPE_WARNING, LogColorNavy) 
                        If device.UniqueDeviceName = bstrUDN And (device.Alive Or OverrideAliveFlag) Then
                            Item = device
                            Exit Property
                        End If
                    End If
                Next
            End If
        End Get
    End Property

    Public Function CheckDeviceExists(UDN As String) As Boolean
        CheckDeviceExists = False
        If upnpDebuglevel > DebugLevel.dlEvents Then Log("MyUPnPDevices.CheckDeviceExists called with UDN = " & UDN, LogType.LOG_TYPE_INFO, LogColorGreen)
        Try
            For Each Device As MyUPnPDevice In Me
                'Log("MyUPnPDevices.CheckDeviceExists found USN = " & Device.UniqueDeviceName & "<->" & USNParts(1) & " and Location = " & Location & "<->" & Device.Location, LogType.LOG_TYPE_WARNING)
                If Device.UniqueDeviceName.IndexOf("uuid:" & UDN) = 0 Then 'changed on 4/10/2018 from ("uuid:" & UDN).IndexOf(Device.UniqueDeviceName) = 0) If uniqyedevicename is "" it returns zero
                    CheckDeviceExists = True
                    Exit Function
                End If
            Next
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff Then Log("Error in MyUPnPDevices.CheckDeviceExists with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Sub Dispose()
        If upnpDebuglevel > DebugLevel.dlOff Then Log("MyUPnPDevices.Dispose called", LogType.LOG_TYPE_INFO, LogColorGreen)
        Try
            For Each device As MyUPnPDevice In Me 'MyDevicesLinkedList
                Try
                    device.Dispose(False)
                    device = Nothing
                Catch ex As Exception
                End Try
            Next
            Clear()
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff Then Log("Error in MyUPnPDevices.Dispose with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

End Class

#End Region

#Region "Services"

Public Class MyUPnPService
    Private MyServiceID As String = ""
    Private MyLastTransportStatus As Integer = 0
    Private MyServiceTypeIdentifier As String = ""
    Public MycontrolURL As String = ""
    Public MyeventSubURL As String = ""
    Public MySCPDURL As String = ""
    Private MyServiceXML As String = ""
    Public ServiceStateTable As New MyServiceStateTable
    Public ActionList As New MyactionList
    Private MyUPNPClient As Socket

    Private MyIPAddress As String = ""
    Private MyIPPort As String = ""
    Private MyDeviceUDN As String = ""
    Public MyReceivedSID As String = ""
    Public MyTimeout As Integer = 600 ' changed from 300 too 600 on 10/13/2019
    Public MyLocation As String = ""
    Public NotificationEvent As String = ""
    Private mySubscribeRenewCounter As Integer = 0
    Private myBootId As String = ""

    ' this is new as of Sonosv4 to further enhance debugging
    Public hasServiceXMLRetrieved As Boolean = False
    Public hasServiceStateRetrieved As Boolean = False
    Public hasActionListRetrieved As Boolean = False
    Public isServiceEventsSubscribed As Boolean = False

    Public Delegate Sub StateChangeEventHandler(StateVarName As String, Value As String)
    Public Event StateChange As StateChangeEventHandler
    Public Delegate Sub ServiceDiedEventHandler()
    Public Event ServiceDied As ServiceDiedEventHandler
    Friend WithEvents MySubscribeRenewTimer As Timers.Timer
    Friend WithEvents MyMissedRenewTimer As Timers.Timer
    Private StateVariableChangedHandlerAddress As StateChangeEventHandler
    Private ServiceDiedChangedHandlerAddress As ServiceDiedEventHandler

    Private EventHandlerReEntryFlag As Boolean = False
    Private MyEventQueue As Queue(Of String) = New Queue(Of String)()
    Private MissedEventHandlerFlag As Boolean = False
    Private LastEventSequenceNumber As Integer = -1
    Private ReferenceToMasterUPNPDevice As MyUPnPDevice = Nothing
    Private MissedRenewCounter As Integer = 0


    Friend WithEvents MyEventTimer As Timers.Timer
    Friend WithEvents MyServiceDiedTimer As Timers.Timer
    Private debugParam As String
    Public ReadOnly Property CheckDebugParam As Boolean
        Get
            If debugParam = "" Then Return True
            Return (debugParam = MyDeviceUDN)
        End Get
    End Property

    Public Sub New(DeviceUDN As String, ServiceXML As String, DeviceURL As String, MasterdeviceReference As MyUPnPDevice)
        MyBase.New()
        'If upnpDebuglevel > DebugLevel.dlErrorsOnly andAlso CheckDebugParam Then Log("MyUPnPService.new called with Service XML = " & ServiceXML & " and DeviceURL =  " & DeviceURL, LogType.LOG_TYPE_INFO)
        If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPService.new called for DeviceUDN = " & DeviceUDN & " and DeviceURL =  " & DeviceURL, LogType.LOG_TYPE_INFO, LogColorGreen)
        If DeviceURL = "" Then
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("MyUPnPService.new called with empty DeviceURL =  " & DeviceURL, LogType.LOG_TYPE_WARNING)
            Throw New System.Exception("Error in MyUPnPService.new called with empty DeviceURL =  " & DeviceURL)
            'Exit Sub
        End If
        ReferenceToMasterUPNPDevice = MasterdeviceReference

        MyLocation = DeviceURL
        MyDeviceUDN = DeviceUDN
        If MyDeviceUDN.IndexOf("uuid:") = 0 Then
            Mid(MyDeviceUDN, 1, 5) = "     "
            MyDeviceUDN = Trim(MyDeviceUDN)
        End If
        Try
            Dim DeviceURI As Uri = New Uri(DeviceURL)
            MyIPAddress = DeviceURI.Host
            MyIPPort = DeviceURI.Port
        Catch ex As Exception
            ' not good, the URL is not good
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("MyUPnPService.new called with invalid DeviceURL =  " & DeviceURL, LogType.LOG_TYPE_WARNING)
            Throw New System.Exception("Error in MyUPnPService.new called with invalid DeviceURL =  " & DeviceURL)
            'Exit Sub
        End Try
        If ServiceXML = "" Then Exit Sub
        Dim ServiceDocument As New XmlDocument
        Try
            ServiceDocument.LoadXml(ServiceXML)
            Try
                MyServiceTypeIdentifier = If(ServiceDocument.GetElementsByTagName("serviceType").Item(0)?.InnerText, "")
                If Trim(LCase(MyServiceTypeIdentifier)) = "(null)" Then ' added this to prevent pesky HomeSeer fault in Service definition to generate error each 20 seconds or so
                    Exit Sub
                End If
            Catch ex As Exception
            End Try
            Try
                MyServiceID = If(ServiceDocument.GetElementsByTagName("serviceId").Item(0)?.InnerText, "")
            Catch ex As Exception
            End Try
            Try
                MycontrolURL = MakeURLWhole(If(ServiceDocument.GetElementsByTagName("controlURL").Item(0)?.InnerText, ""))
            Catch ex As Exception
            End Try
            Try
                MyeventSubURL = MakeURLWhole(If(ServiceDocument.GetElementsByTagName("eventSubURL").Item(0)?.InnerText, ""))
            Catch ex As Exception
            End Try
            Try
                MySCPDURL = MakeURLWhole(If(ServiceDocument.GetElementsByTagName("SCPDURL").Item(0)?.InnerText, ""))
            Catch ex As Exception
            End Try
            If MySCPDURL <> "" Then
                Dim xmlDoc As New XmlDocument With {
                    .XmlResolver = Nothing
                }
                Dim PageHTML As String = ""
                If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPService.new is retrieving the Service XML for ServiceID = " & MySCPDURL, LogType.LOG_TYPE_INFO, LogColorGreen)
                Try
                    Dim RequestUri As System.Uri = New Uri(MySCPDURL)
                    Dim p As System.Net.ServicePoint = ServicePointManager.FindServicePoint(RequestUri)
                    p.Expect100Continue = False
                    Dim webRequest As HttpWebRequest = DirectCast(System.Net.HttpWebRequest.Create(RequestUri), HttpWebRequest) 'HttpWebRequest.Create(RequestUri)
                    'Dim webRequest As WebRequest = System.Net.WebRequest.Create(RequestUri)
                    webRequest.Method = "GET"
                    'webRequest.Headers.Add("Connection", "close")
                    webRequest.KeepAlive = False
                    webRequest.Timeout = 20000  ' added 10/24/2019
                    Using webResponse As System.Net.WebResponse = webRequest.GetResponse
                        Using webStream As System.IO.Stream = webResponse.GetResponseStream
                            Using strmRdr As New System.IO.StreamReader(webStream)
                                PageHTML = strmRdr.ReadToEnd()
                                strmRdr.Close()
                                strmRdr.Dispose()
                                webStream.Close()
                                webResponse.Close()
                                webStream.Dispose()
                            End Using
                        End Using
                    End Using
                Catch ex As Exception
                    If MyServiceID = "urn:dial-multiscreen-org:serviceId:dial" Then ' added 10/24/2019
                        ' most dial devices have no service info and unfortunately (ex Google nest) have wrong info ex ssdp/notfound in the Google Nest case
                        If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("Error in MyUPnPService.new for DIAL Service = " & MyServiceID & " while retrieving document with URL = " & MySCPDURL & " for ServiceID =  " & MyServiceID & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                        Exit Sub
                    End If
                    If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPService.new for Service = " & MyServiceID & " while retrieving document with URL = " & MySCPDURL & " for ServiceID =  " & MyServiceID & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    'Throw New System.Exception("Error in MyUPnPService.new for Service = " & MyServiceID & " while retrieving document with URL = " & MySCPDURL & " for ServiceID =  " & MyServiceID & " with error = " & ex.Message)
                    Exit Sub
                End Try
                MyServiceXML = RemoveControlCharacters(PageHTML)
                hasServiceXMLRetrieved = True
            End If
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPService.new retrieving Service XML = " & ServiceXML & " for ServiceID = " & MySCPDURL & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Throw New System.Exception("Error in MyUPnPService.new retrieving Service XML = " & ServiceXML & " for ServiceID = " & MySCPDURL & " with error = " & ex.Message)
            'Exit Sub
        End Try
        Dim ServiceXMLDocument As New XmlDocument
        Dim _actionList As System.Xml.XmlNodeList = Nothing
        Dim _serviceStateTable As System.Xml.XmlNodeList = Nothing

        Try
            ' new open the servicedocument and get all variables
            ServiceXMLDocument.LoadXml(MyServiceXML)
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPService.new loading Service XML = " & MyServiceXML & " for ServiceID = " & MySCPDURL & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            ServiceXMLDocument = Nothing
            Throw New System.Exception("Error in MyUPnPService.new loading Service XML = " & MyServiceXML & " for ServiceID = " & MySCPDURL & " with error = " & ex.Message)
            'Exit Sub
        End Try
        '<serviceStateTable>
        '  <stateVariable sendEvents="no">
        '    <name>TransportState</name>
        '    <dataType>string</dataType>
        '    <allowedValueList>
        '      <allowedValue>STOPPED</allowedValue>
        '      <allowedValue>PLAYING</allowedValue>
        '      <allowedValue>PAUSED_PLAYBACK</allowedValue>
        '      <allowedValue>TRANSITIONING</allowedValue>
        '      <allowedValue>NO_MEDIA_PRESENT</allowedValue>
        '    </allowedValueList>
        '  </stateVariable>
        '  <stateVariable sendEvents="no">
        '    <name>NumberOfTracks</name>
        '    <dataType>ui4</dataType>
        '    <allowedValueRange>
        '      <minimum>0</minimum>
        '      <maximum>1</maximum>
        '    </allowedValueRange>
        '  </stateVariable>
        '<stateVariable sendEvents="no">
        '  <name>CurrentTrack</name>
        '  <dataType>ui4</dataType>
        '    <allowedValueRange>
        '      <minimum>0</minimum>
        '      <maximum>1</maximum>
        '      <step>1</step>
        '    </allowedValueRange>
        '</stateVariable>
        '
        ' the <dataType> node name can have an atribute "type" example <dataType type="xsd:byte">string</dataType>

        Try
            ' retrieve the ServiceStateTable list
            Dim ServiceStateXML As String = If(ServiceXMLDocument.GetElementsByTagName("serviceStateTable").Item(0)?.OuterXml, "")
            Dim ServiceStateXMLDocument As New XmlDocument
            ServiceStateXMLDocument.LoadXml(ServiceStateXML)
            _serviceStateTable = ServiceStateXMLDocument.GetElementsByTagName("stateVariable")
            ServiceStateXMLDocument = Nothing
            hasServiceStateRetrieved = True
            'serviceStateTable = ServiceXMLDocument.GetElementsByTagName("stateVariable")
        Catch ex As Exception
            ' could be that there is no action list
            _serviceStateTable = Nothing
            If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("Warning in MyUPnPService.new retrieving the serviceStateTable in Service XML = " & MyServiceXML & " for ServiceID = " & MySCPDURL & " with error = " & ex.Message, LogType.LOG_TYPE_WARNING)
        End Try
        '<actionList>
        '  <action>
        '    <name>SetAVTransportURI</name>
        '      <argumentList>
        '        <argument>
        '          <name>InstanceID</name>
        '          <direction>in</direction>
        '          <relatedStateVariable>A_ARG_TYPE_InstanceID</relatedStateVariable>
        '        </argument>
        '        <argument>
        '          <name>CurrentURI</name>
        '          <direction>in</direction>
        '          <relatedStateVariable>AVTransportURI</relatedStateVariable>
        '        </argument>
        '    </argumentList>
        '  </action>

        Try
            ' retrieve the action list
            Dim ActionXML As String = If(ServiceXMLDocument.GetElementsByTagName("actionList").Item(0)?.OuterXml, "")
            Dim ActionXMLDocument As New XmlDocument
            ActionXMLDocument.LoadXml(ActionXML)
            _actionList = ActionXMLDocument.GetElementsByTagName("action")
            hasActionListRetrieved = True
            ActionXMLDocument = Nothing
        Catch ex As Exception
            ' could be that there is no action list, not sure this would be valid because what would we do with this service that has no actions ??
            _actionList = Nothing
            If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("Warning in MyUPnPService.new retrieving the actionList in Service XML = " & MyServiceXML & " for ServiceID = " & MySCPDURL & " with error = " & ex.Message, LogType.LOG_TYPE_WARNING)
        End Try
        ServiceXMLDocument = Nothing

        Try
            If _serviceStateTable IsNot Nothing Then
                If _serviceStateTable.Count > 0 Then
                    For Each stateVariable As System.Xml.XmlNode In _serviceStateTable
                        Try
                            If Trim(stateVariable.OuterXml) <> "" Then
                                Dim StateVariableXMLDocument As New XmlDocument
                                StateVariableXMLDocument.LoadXml(stateVariable.OuterXml)
                                Dim NewStateVariable As New MyStateVariable
                                Try
                                    NewStateVariable.dataType = VariableDataTypes.vdtString
                                    Try
                                        If StateVariableXMLDocument IsNot Nothing Then
                                            If StateVariableXMLDocument.HasChildNodes And StateVariableXMLDocument.ChildNodes.Count > 0 Then
                                                For Each StatevariableNode As XmlNode In StateVariableXMLDocument.ChildNodes
                                                    If StatevariableNode.LocalName = "stateVariable" Then
                                                        Try
                                                            NewStateVariable.sendEvents = StatevariableNode.Attributes("sendEvents").Value.ToUpper = "YES"
                                                        Catch ex As Exception
                                                            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPService.new processing a stateVariable for sendEvents in StateVariableXML = " & stateVariable.OuterXml & " for ServiceID = " & MySCPDURL & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                                        End Try
                                                        If StatevariableNode.HasChildNodes And StatevariableNode.ChildNodes.Count > 0 Then
                                                            For Each ChildNode As XmlNode In StatevariableNode.ChildNodes
                                                                Select Case ChildNode.Name
                                                                    Case "name"
                                                                        NewStateVariable.name = ChildNode.InnerText
                                                                    Case "dataType"
                                                                        Dim DataType As String = ChildNode.InnerText
                                                                        ' allowed types: ui1, ui2, ui4, i1, i2, i4, int, r4, r8, number, fixed.14.4, float, char, string, date, dateTime, dateTime.tz, time, time.tz, boolean, bin.base64, bin.hex, uri, uuid
                                                                        Select Case DataType.ToUpper
                                                                            Case "STRING"
                                                                                NewStateVariable.dataType = VariableDataTypes.vdtString
                                                                            Case "BOOLEAN"
                                                                                NewStateVariable.dataType = VariableDataTypes.vdtBoolean
                                                                            Case "UI4", "I4", "BIN.BASE64", "INT"
                                                                                NewStateVariable.dataType = VariableDataTypes.vdtUI4
                                                                            Case "UI2", "I2", "UI1"
                                                                                NewStateVariable.dataType = VariableDataTypes.vdtUI2
                                                                            Case Else
                                                                                If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("Warning in MyUPnPService.new processing a stateVariable for dataType = " & DataType & " for ServiceID = " & MySCPDURL, LogType.LOG_TYPE_WARNING)
                                                                                NewStateVariable.dataType = VariableDataTypes.vdtString
                                                                        End Select
                                                                    Case "allowedValueRange"
                                                                        NewStateVariable.allowedValueRange = ChildNode.OuterXml
                                                                    Case "allowedValueList"
                                                                        NewStateVariable.allowedValueList = ChildNode.OuterXml
                                                                    Case "defaultValue"
                                                                        NewStateVariable.defaultValue = ChildNode.InnerText
                                                                    Case Else
                                                                        If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("Warning in MyUPnPService.new processing a stateVariable. Found unknown XMLNode = " & ChildNode.Name & " and XML = " & ChildNode.OuterXml.ToString & " for ServiceID = " & MySCPDURL, LogType.LOG_TYPE_WARNING)
                                                                End Select
                                                            Next
                                                        End If
                                                    End If
                                                Next
                                            End If
                                        End If
                                    Catch ex As Exception
                                        If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPService.new processing a stateVariable in Service XML = " & stateVariable.OuterXml & " for ServiceID = " & MySCPDURL & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                    End Try
                                    ServiceStateTable.Add(NewStateVariable)
                                Catch ex As Exception
                                    ' without a name, this is pretty useless
                                    If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPService.new processing a stateVariable with name missing in Service XML = " & stateVariable.OuterXml & " for ServiceID = " & MySCPDURL & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                                StateVariableXMLDocument = Nothing
                            End If
                        Catch ex As Exception
                            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPService.new processing a stateVariable in Service XML = " & stateVariable.OuterXml & " for ServiceID = " & MySCPDURL & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                    Next
                End If
            End If
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPService.new processing the stateVariableList in Service XML = " & MyServiceXML & " for ServiceID = " & MySCPDURL & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If _actionList IsNot Nothing Then
                If _actionList.Count > 0 Then
                    For Each Action As System.Xml.XmlNode In _actionList
                        Try
                            If Trim(Action.OuterXml) <> "" Then
                                Dim ActionListXML As New XmlDocument
                                ActionListXML.LoadXml(Action.OuterXml)
                                Dim NewAction As New MyUPnPAction
                                Try
                                    NewAction.name = If(ActionListXML.GetElementsByTagName("name").Item(0)?.InnerText, "")
                                    Try
                                        NewAction.argumentList = ActionListXML.GetElementsByTagName("argumentList").ItemOf(0)?.OuterXml
                                    Catch ex As Exception
                                    End Try
                                    ActionList.Add(NewAction)
                                Catch ex As Exception
                                    ' without a name, this is pretty useless
                                    If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPService.new processing an action with the name missing in Action XML = " & Action.OuterXml & " for ServiceID = " & MySCPDURL & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            End If
                        Catch ex As Exception
                            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPService.new processing a stateVariable in Action XML = " & Action.OuterXml & " for ServiceID = " & MySCPDURL & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                    Next
                End If
            End If
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPService.new processing the ActionList in Action XML = " & MyServiceXML & " for ServiceID = " & MySCPDURL & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Property Id As String
        Get
            Id = MyServiceID
        End Get
        Set(value As String)
            MyServiceID = value
        End Set
    End Property

    ReadOnly Property ServiceType As String
        Get
            ServiceType = MyServiceTypeIdentifier
        End Get
    End Property

    ReadOnly Property LastTransportStatus As Integer
        Get
            LastTransportStatus = MyLastTransportStatus
        End Get
    End Property

    Property ServiceTypeIdentifier As String
        Get
            ServiceTypeIdentifier = MyServiceTypeIdentifier
        End Get
        Set(Value As String)
            MyServiceTypeIdentifier = Value
        End Set
    End Property

    Public Function AddCallback(ByRef pUnkCallback As MyUPnPControlCallback) As Boolean ' what is provided is a callback address

        AddCallback = False

        'If UPnPSubscribeTimeOut <> 0 Then
        'MyTimeout = UPnPSubscribeTimeOut
        'End If

        ' this is how to subscribe to events from the Device Spy Capture
        'SUBSCRIBE /MediaRenderer/AVTransport/Event HTTP/1.1
        'NT: upnp:event
        'HOST: 192.168.1.104:1400
        'CALLBACK: <http://192.168.1.110:6062/RINCON_000E5859008A01400_MR/urn:upnp-org:serviceId:AVTransport>
        'TIMEOUT: Second-300
        'Content-Length: 0

        ' this is the response
        'HTTP/1.1 200 OK
        'SID: uuid:RINCON_000E5859008A01400_sub0000000331
        'TIMEOUT: Second-300
        'Server: Linux UPnP/1.0 Sonos/24.0-71060 (ZPS5)
        'Connection: close

        ' Here is a subscription event from Microsoft
        'SUBSCRIBE /upnphost/udhisapi.dll?event=uuid:79f84a69-0f02-4bd8-b6b0-f72f89ed86f3+urn:microsoft.com:serviceId:X_MS_MediaReceiverRegistrar HTTP/1.1
        'USER-AGENT: Linux UPnP/1.0 Sonos/24.0-71060 (ZPS5)
        'HOST: 192.168.1.110:2869
        'SID: uuid:72e42da0-ebfb-4792-b0e5-a6c06e2e4269
        'TIMEOUT: Second-3600

        'HTTP/1.1 200 OK
        'Server: Microsoft-Windows-NT/5.1 UPnP/1.0 UPnP-Device-Host/1.0 Microsoft-HTTPAPI/2.0
        'Timeout: Second-300
        'SID: uuid:72e42da0-ebfb-4792-b0e5-a6c06e2e4269
        'Date: Fri, 21 Mar 2014 04:06:45 GMT
        'Content-Length: 0

        If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPService.AddCallback called for ServiceID = " & MySCPDURL & " and UDN = " & MyDeviceUDN, LogType.LOG_TYPE_INFO, LogColorGreen)

        If MyIPAddress = "" Or MyeventSubURL = "" Then
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPService.AddCallback. IPAddress or IPPort or MyeventSubURL are missing. IPAddress = " & MyIPAddress & ", IPPort = " & MyIPPort & " MyeventSubURL = " & MyeventSubURL, LogType.LOG_TYPE_ERROR)
            Exit Function
        End If

        Dim RequestUri = New Uri(MyeventSubURL)

        Dim StatusCode As String = ""
        Dim StatusDescription As String = ""
        Dim TimeOut As String = ""
        Dim SID As String = ""

        Try
            Dim p = ServicePointManager.FindServicePoint(RequestUri)
            p.Expect100Continue = False
            Dim wRequest As HttpWebRequest = DirectCast(System.Net.HttpWebRequest.Create(RequestUri), HttpWebRequest) 'HttpWebRequest.Create(RequestUri)
            wRequest.Method = "SUBSCRIBE" '& MyeventSubURL
            wRequest.Host = RequestUri.Authority
            wRequest.ProtocolVersion = HttpVersion.Version11
            wRequest.Headers.Add("NT", "upnp:event")
            wRequest.Headers.Add("CALLBACK", "<http://" & PlugInIPAddress & ":" & TCPListenerPort.ToString & "/" & MyDeviceUDN & "/" & MyServiceID & ">")
            wRequest.Headers.Add("TIMEOUT", "Second-" & "1800") 'MyTimeout.ToString) changed 10/13/2019 set this fixed to 1/2 hour but send renews every 5 min.This guarentees eventing does go to sleep in other side
            wRequest.ContentLength = 0
            wRequest.KeepAlive = False
            wRequest.Timeout = 10000    ' added 10/24/2019
            Using webResponse As HttpWebResponse = wRequest.GetResponse

                If upnpDebuglevel > DebugLevel.dlEvents Then
                    Log("AddCallback got Method Response = " & webResponse.Method.ToString, LogType.LOG_TYPE_INFO, LogColorGreen)
                    Log("AddCallback got StatusCode Response = " & webResponse.StatusCode.ToString, LogType.LOG_TYPE_INFO, LogColorGreen)
                    Log("AddCallback got StatusDescription Response = " & webResponse.StatusDescription.ToString, LogType.LOG_TYPE_INFO, LogColorGreen)
                    Log("AddCallback got ProtocolVersion Response = " & webResponse.ProtocolVersion.ToString, LogType.LOG_TYPE_INFO, LogColorGreen)
                    Log("AddCallback got ResponseUri Response = " & webResponse.ResponseUri.ToString, LogType.LOG_TYPE_INFO, LogColorGreen)
                    Log("AddCallback got Server Response = " & webResponse.Server.ToString, LogType.LOG_TYPE_INFO, LogColorGreen)
                    Log("AddCallback got ContentLength  Response = " & webResponse.ContentLength.ToString, LogType.LOG_TYPE_INFO, LogColorGreen)
                End If

                If upnpDebuglevel > DebugLevel.dlEvents Then
                    For Each header In webResponse.Headers
                        Log("AddCallback got Header Response = " & header.ToString & " and Value = " & webResponse.Headers.Get(header).ToString, LogType.LOG_TYPE_INFO, LogColorGreen)
                    Next
                End If

                StatusCode = webResponse.StatusCode.ToString                ' OK
                StatusDescription = webResponse.StatusDescription.ToString  ' OK
                TimeOut = webResponse.Headers.Get("TIMEOUT").ToString       ' Second-300 ' Need this to set timer to renew subscriptions
                SID = webResponse.Headers.Get("SID").ToString               ' uuid:RINCON_000E5859008A01400_sub0000001133  need this for renewing the Subscription
                webResponse.Close()
                wRequest = Nothing
            End Using
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPService.AddCallback while sending URL = " & MyeventSubURL & " to = " & RequestUri.OriginalString.ToString & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Throw New System.Exception("Error in MyUPnPService.AddCallback while sending URL = " & MyeventSubURL & " to = " & RequestUri.OriginalString.ToString & " with error = " & ex.Message)
            Exit Function
        End Try

        If StatusCode.ToUpper <> "OK" Then
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("MyUPnPService.AddCallback called for ServiceID = " & MySCPDURL & " and ended unsuccessfully with StatusCode = " & StatusCode & " and Description = " & StatusDescription, LogType.LOG_TYPE_ERROR)
            Throw New System.Exception("MyUPnPService.AddCallback called for ServiceID = " & MySCPDURL & " and ended unsuccessfully with StatusCode = " & StatusCode & " and Description = " & StatusDescription)
            Exit Function
        End If
        mySubscribeRenewCounter = 0
        Try
            MyReceivedSID = SID
            Dim ReceivedTimeOutData As Integer = RetrieveTimeoutData(TimeOut)
            If ReceivedTimeOutData < MyTimeout Then
                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Warning in MyUPnPService.AddCallback for ServiceID = " & MySCPDURL & ". Timeout info is smaller from requested. Requested = " & MyTimeout.ToString & " received = " & ReceivedTimeOutData.ToString, LogType.LOG_TYPE_WARNING, LogColorNavy)
                MyTimeout = ReceivedTimeOutData
            End If
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPService.AddCallback for ServiceID = " & MySCPDURL & ". Unable to extract the timeout info with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Throw New System.Exception("Error in MyUPnPService.AddCallback for ServiceID = " & MySCPDURL & ". Unable to extract the timeout info with error = " & ex.Message)
            Exit Function
        End Try
        isServiceEventsSubscribed = True
        Randomize()
        ' Generate random value between 1 second and an 8th of the requested interval. 
        Dim value As Integer = CInt(Int((MyTimeout / 8 * Rnd()) + 1)) * 1000
        Try
            MySubscribeRenewTimer = New Timers.Timer With {
                .Interval = MyTimeout * 500 + value,  ' this is a divide by 2 !
                .AutoReset = True,
                .Enabled = True
            }
            If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MyUPnPService.AddCallback for ServiceID = " & MySCPDURL & " started Renew timeout = " & (MySubscribeRenewTimer.Interval / 1000).ToString, LogType.LOG_TYPE_INFO, LogColorGreen)
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPService.AddCallback for ServiceID = " & MySCPDURL & ". Unable to start the subscribe renew timer with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Throw New System.Exception("Error in MyUPnPService.AddCallback for ServiceID = " & MySCPDURL & ". Unable to start the subscribe renew timer with error = " & ex.Message)
            Exit Function
        End Try

        If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPService.AddCallback called for ServiceID = " & MySCPDURL & " and ended successfully with SID = " & MyReceivedSID & " and Timeout = " & MyTimeout.ToString & ", StatusCode = " & StatusCode & ", StatusDescription = " & StatusDescription, LogType.LOG_TYPE_INFO, LogColorGreen)

        If pUnkCallback IsNot Nothing Then
            StateVariableChangedHandlerAddress = AddressOf pUnkCallback.StateVariableChanged
            ServiceDiedChangedHandlerAddress = AddressOf pUnkCallback.ServiceInstanceDied
            AddHandler StateChange, StateVariableChangedHandlerAddress
            AddHandler ServiceDied, ServiceDiedChangedHandlerAddress
        End If

        Return True

    End Function

    Private Function SendRenew() As Boolean

        SendRenew = False
        If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MyUPnPService.SendRenew called for ServiceID = " & MySCPDURL, LogType.LOG_TYPE_INFO, LogColorGreen)

        If ReferenceToMasterUPNPDevice IsNot Nothing Then
            If Not ReferenceToMasterUPNPDevice.Alive Then
                StopTimers(False)
                isServiceEventsSubscribed = False
                Try
                    RaiseEvent ServiceDied()
                Catch ex As Exception
                    If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("Error in MyUPnPService.SendRenew raising the event for ServiceID = " & MySCPDURL & " with Error = " & ex.Message, LogType.LOG_TYPE_INFO)
                End Try
                Exit Function
            End If
        End If

        Dim RequestUri = New Uri(MyeventSubURL)

        Dim StatusCode As String = ""
        Dim StatusDescription As String = ""
        Dim TimeOut As String = ""
        Dim SID As String = ""
        Dim beforeTime As DateTime = DateTime.Now

        Try
            'Dim p = ServicePointManager.FindServicePoint(RequestUri)   ' removed 11/10/2019
            'p.Expect100Continue = False

            Dim wRequest As HttpWebRequest = DirectCast(System.Net.HttpWebRequest.Create(RequestUri), HttpWebRequest) 'HttpWebRequest.Create(RequestUri)
            wRequest.Method = "SUBSCRIBE"
            wRequest.Host = RequestUri.Authority
            wRequest.ProtocolVersion = HttpVersion.Version11
            wRequest.Headers.Add("SID", MyReceivedSID)
            wRequest.Headers.Add("TIMEOUT", "Second-1800")
            wRequest.KeepAlive = False
            wRequest.ContentLength = 0
            wRequest.Timeout = 10000 ' changed to 10 sec on 10/24/2019 fro max 5 seconds added back 10/13/2019. Seems to help when a player is off-line, else it times out at random (minutes!)
            Using webResponse As HttpWebResponse = wRequest.GetResponse

                If upnpDebuglevel > DebugLevel.dlEvents Then
                    Log("MyUPnPService.SendRenew got Method Response = " & webResponse.Method.ToString, LogType.LOG_TYPE_INFO, LogColorGreen)
                    Log("MyUPnPService.SendRenew got StatusCode Response = " & webResponse.StatusCode.ToString, LogType.LOG_TYPE_INFO, LogColorGreen)
                    Log("MyUPnPService.SendRenew got StatusDescription Response = " & webResponse.StatusDescription.ToString, LogType.LOG_TYPE_INFO, LogColorGreen)
                    Log("MyUPnPService.SendRenew got ProtocolVersion Response = " & webResponse.ProtocolVersion.ToString, LogType.LOG_TYPE_INFO, LogColorGreen)
                    Log("MyUPnPService.SendRenew got ResponseUri Response = " & webResponse.ResponseUri.ToString, LogType.LOG_TYPE_INFO, LogColorGreen)
                    Log("MyUPnPService.SendRenew got Server Response = " & webResponse.Server.ToString, LogType.LOG_TYPE_INFO, LogColorGreen)
                    Log("MyUPnPService.SendRenew got ContentLength  Response = " & webResponse.ContentLength.ToString, LogType.LOG_TYPE_INFO, LogColorGreen)
                    For Each header In webResponse.Headers
                        Log("MyUPnPService.SendRenew got Header Response = " & header.ToString & " and Value = " & webResponse.Headers.Get(header).ToString, LogType.LOG_TYPE_INFO, LogColorGreen)
                    Next
                End If

                StatusCode = webResponse.StatusCode.ToString                ' OK
                StatusDescription = webResponse.StatusDescription.ToString  ' OK
                TimeOut = webResponse.Headers.Get("TIMEOUT").ToString       ' Second-300 ' Need this to set timer to renew subscriptions
                SID = webResponse.Headers.Get("SID").ToString               ' uuid:RINCON_000E5859008A01400_sub0000001133  need this for renewing the Subscription
                'If UPnPDebuglevel > DebugLevel.dlErrorsOnly andAlso CheckDebugParam Then Log("Succesful MyUPnPService.SendRenew for ServiceID = " & MySCPDURL & " while sending URL = " & MyeventSubURL & " with OldSID = " & MyReceivedSID & " with NewSID = " & SID, LogType.LOG_TYPE_WARNING)
                MissedRenewCounter = 0
                webResponse.Close()
                wRequest = Nothing
                'p = Nothing
            End Using
            Dim afterTime As DateTime = DateTime.Now
            Dim deltaTime As TimeSpan = afterTime.Subtract(beforeTime)
            If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPService.SendRenew for ServiceID = " & MySCPDURL & " sent renewal with elapsed time = " & deltaTime.TotalMilliseconds.ToString, LogType.LOG_TYPE_INFO, LogColorGreen)
        Catch ex As WebException
            isServiceEventsSubscribed = False
            Dim afterTime As DateTime = DateTime.Now
            Dim deltaTime As TimeSpan = afterTime.Subtract(beforeTime)
            If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("Unsuccesfull MyUPnPService.SendRenew for ServiceID = " & MySCPDURL & " while sending URL = " & MyeventSubURL & " with SID = " & MyReceivedSID & " sent renewal with elapsed time = " & deltaTime.TotalMilliseconds.ToString & " and error = " & ex.Message, LogType.LOG_TYPE_WARNING)
            Using webResponse As HttpWebResponse = ex.Response
                If webResponse IsNot Nothing Then
                    Using webStream As Stream = webResponse.GetResponseStream
                        Dim strmRdr As New System.IO.StreamReader(webStream)
                        Dim ResponseHTML As String = strmRdr.ReadToEnd()
                        strmRdr.Dispose()
                        If webResponse.StatusCode = HttpStatusCode.PreconditionFailed Then
                            ' actually upon further study, I found it if I use a laptop, put it in sleep mode, wake it, I get this error. It means the SID is not valid anymore so we need to reconnect
                            If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("Unsuccesfull MyUPnPService.SendRenew for ServiceID = " & MySCPDURL & " because code 412. Typically means the client had a problem communicating an event and released the subscription. We will try to re-subscribe. Network issue?", LogType.LOG_TYPE_WARNING)
                            Try
                                If AddCallback(Nothing) Then
                                    Return True
                                End If
                            Catch ex1 As Exception
                                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Unsuccesfull adding new callback MyUPnPService.SendRenew for ServiceID = " & MySCPDURL & " with Error = " & ex1.Message, LogType.LOG_TYPE_ERROR)
                            End Try
                        Else
                            If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("Unsuccesfull MyUPnPService.SendRenew for ServiceID = " & MySCPDURL & " while response = " & ResponseHTML, LogType.LOG_TYPE_WARNING)
                        End If
                    End Using
                Else
                    mySubscribeRenewCounter += 1
                    If mySubscribeRenewCounter < 2 Then
                        SetRenewTimer(500)
                        Return True
                    End If
                End If
            End Using
            StopTimers(False)
            Try
                RaiseEvent ServiceDied()
            Catch ex1 As Exception
                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPService.SendRenew raising the event for ServiceID = " & MySCPDURL & " with Error = " & ex1.Message, LogType.LOG_TYPE_INFO)
            End Try
            Exit Function

        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Unsuccesfull MyUPnPService.SendRenew(1) for ServiceID = " & MySCPDURL & " while sending URL = " & MyeventSubURL & " with SID = " & MyReceivedSID & " and error = " & ex.Message, LogType.LOG_TYPE_WARNING)
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Unsuccesfull MyUPnPService.SendRenew(1) for ServiceID = " & MySCPDURL & " with StatusCode = " & StatusCode & " and Description = " & StatusDescription, LogType.LOG_TYPE_WARNING, LogColorNavy)
            isServiceEventsSubscribed = False
            StopTimers(False)
            Try
                RaiseEvent ServiceDied()
            Catch ex1 As Exception
                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPService.SendRenew raising the event for ServiceID = " & MySCPDURL & " with Error = " & ex1.Message, LogType.LOG_TYPE_INFO)
            End Try
            Exit Function
        End Try
        RequestUri = Nothing

        Try
            Dim ReceivedTimeOutData As Integer = RetrieveTimeoutData(TimeOut)
            If ReceivedTimeOutData < 1200 Then ' changed from comparision with MyTimeout to 1800 on 11/15/2019. We run the renew at 600 but ask for 1800
                If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("Warning in MyUPnPService.SendRenew for ServiceID = " & MySCPDURL & ". Timeout info is different from requested. Requested = " & MyTimeout.ToString & " received = " & ReceivedTimeOutData.ToString, LogType.LOG_TYPE_WARNING, LogColorNavy)
                Dim NewTimeout As Integer = ReceivedTimeOutData
                Randomize()
                ' Generate random value between 1 second and an 8th of the requested interval. 
                Dim value As Integer = CInt(Int((NewTimeout / 8 * Rnd()) + 1)) * 1000
                If MySubscribeRenewTimer IsNot Nothing Then
                    MySubscribeRenewTimer.Enabled = False
                    MySubscribeRenewTimer.Interval = NewTimeout * 500 + value  ' this is a divide by 2 !
                    MySubscribeRenewTimer.AutoReset = True
                    MySubscribeRenewTimer.Enabled = True
                    If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("Warning in MyUPnPService.SendRenew for ServiceID = " & MySCPDURL & ". New Timeout = " & MySubscribeRenewTimer.Interval.ToString, LogType.LOG_TYPE_WARNING, LogColorNavy)
                End If
            End If
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPService.SendRenew for ServiceID = " & MySCPDURL & ". Unable to reset timeout info with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        If StatusCode.ToUpper <> "OK" Then
            'If UPnPDebuglevel > DebugLevel.dlErrorsOnly andAlso CheckDebugParam Then Log("Unsuccesfull SendRenew for ServiceID = " & MySCPDURL & " with StatusCode = " & StatusCode & " and Description = " & StatusDescription, LogType.LOG_TYPE_WARNING)
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Unsuccesfull MyUPnPService.SendRenew for ServiceID = " & MySCPDURL & " with StatusCode = " & StatusCode & " and Description = " & StatusDescription, LogType.LOG_TYPE_WARNING, LogColorNavy)
            StopTimers(False)
            isServiceEventsSubscribed = False
            Try
                RaiseEvent ServiceDied()
            Catch ex As Exception
                If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("Error in MyUPnPService.SendRenew raising the event for ServiceID = " & MySCPDURL & " with Error = " & ex.Message, LogType.LOG_TYPE_INFO)
            End Try
            Exit Function
        End If
        SendRenew = True
    End Function

    Private Sub StopTimers(sendCancel As Boolean)
        Try
            If MyMissedRenewTimer IsNot Nothing Then
                MyMissedRenewTimer.Enabled = False
                MyMissedRenewTimer.Stop()
                MyMissedRenewTimer.Dispose()
            End If
        Catch ex As Exception
        End Try
        MyMissedRenewTimer = Nothing
        Try
            If MyEventTimer IsNot Nothing Then
                MyEventTimer.Enabled = False
                MyEventTimer.Stop()
                MyEventTimer.Dispose()
            End If
        Catch ex As Exception
        End Try
        MyEventTimer = Nothing
        If MySubscribeRenewTimer IsNot Nothing Then
            Try
                MySubscribeRenewTimer.Enabled = False
                MySubscribeRenewTimer.Stop()
                MySubscribeRenewTimer.Dispose()
            Catch ex As Exception
            End Try
            If sendCancel Then SendCancelSubscription()
        End If
        MySubscribeRenewTimer = Nothing
    End Sub

    Private Sub SetRenewTimer(newTime As Integer)
        Try
            If MyMissedRenewTimer Is Nothing Then
                MyMissedRenewTimer = New Timers.Timer With {
                .Interval = newTime,
                .AutoReset = False,
                .Enabled = False
            }
            End If
            MyMissedRenewTimer.Enabled = True
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPService.SetRenewTimer. Unable to create the renew timer for ServiceID = " & MySCPDURL & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub SendCancelSubscription()
        If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPService.SendCancelSubscription called for ServiceID = " & MySCPDURL, LogType.LOG_TYPE_INFO, LogColorGreen)
        isServiceEventsSubscribed = False
        Dim RequestUri = New Uri(MyeventSubURL)
        Try
            Dim p = ServicePointManager.FindServicePoint(RequestUri)
            p.Expect100Continue = False
        Catch ex As Exception
        End Try
        Try
            Dim wRequest As HttpWebRequest = DirectCast(System.Net.HttpWebRequest.Create(RequestUri), HttpWebRequest) 'HttpWebRequest.Create(RequestUri)
            wRequest.Method = "UNSUBSCRIBE"
            wRequest.Host = RequestUri.Authority
            wRequest.ProtocolVersion = HttpVersion.Version11
            wRequest.Headers.Add("SID", MyReceivedSID)
            wRequest.KeepAlive = False
            wRequest.ContentLength = 0
            wRequest.Timeout = 500 ' very short wait
            Using webResponse As HttpWebResponse = wRequest.GetResponse
                webResponse.Close()
            End Using
            wRequest = Nothing
        Catch ex As Exception
            'Log("Error in MyUPnPService.SendCancelSubscription while sending URL = " & MyeventSubURL & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR) ' if the serviceDied then we can't cancel
        End Try
    End Sub

    Public Sub HandleUPNPDataReceived(InXML As String)
        If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MyUPnPService.HandleUPNPDataReceived called and InXML = " & InXML, LogType.LOG_TYPE_INFO, LogColorGreen)
        'If UPnPDebuglevel > DebugLevel.dlErrorsOnly andAlso CheckDebugParam Then Log("MyUPnPService.HandleUPNPDataReceived for ServiceID = " & MySCPDURL & ", DeviceUDN = " & MyDeviceUDN & ", IPAddress = " & MyIPAddress & " and inXMLLength = " & InXML.Length.ToString, LogType.LOG_TYPE_INFO, LogColorGreen)
        If MyEventTimer Is Nothing Then
            Try
                MyEventTimer = New Timers.Timer With {
                    .Interval = 100, ' 100 msecond
                    .AutoReset = False,
                    .Enabled = False
                }
            Catch ex As Exception
                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPService.HandleUPNPDataReceived. Unable to create the Event timer for ServiceID = " & MySCPDURL & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If
        Try
            SyncLock (MyEventQueue)
                MyEventQueue.Enqueue(InXML)
            End SyncLock
            MyEventTimer.Enabled = True
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPService.HandleUPNPDataReceived for ServiceID = " & MySCPDURL & " queuing the Event = " & InXML.ToString & " and Error = " & ex.Message, LogType.LOG_TYPE_INFO)
        End Try
    End Sub

    Function InvokeAction(bstrActionName As String, vInActionArgs As Object, ByRef pvOutActionArgs As Object, Optional ByRef inCookieContainer As CookieContainer = Nothing) As Object

        ' Here's an example of a InvokeAction
        'POST /MediaServer/ConnectionManager/Control HTTP/1.1
        'HOST: 192.168.1.104:1400
        'SOAPACTION: "urn:schemas-upnp-org:service:ConnectionManager:1#GetProtocolInfo"
        'CONTENT-TYPE: text/xml ; charset="utf-8"
        'Content-Length: 294
        '<?xml version="1.0" encoding="utf-8"?>
        '<s:Envelope s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/" xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
        '   <s:Body>
        '      <u:GetProtocolInfo xmlns:u="urn:schemas-upnp-org:service:ConnectionManager:1" />
        '   </s:Body>
        '</s:Envelope>

        'HTTP/1.1 200 OK
        'CONTENT-LENGTH: 444
        'CONTENT-TYPE: text/xml; charset="utf-8"
        'EXT: 
        'Server: Linux UPnP/1.0 Sonos/24.0-71060 (ZPS5)
        'Connection: close
        '<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/" s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
        '  <s:Body>
        '    <u:GetProtocolInfoResponse xmlns:u="urn:schemas-upnp-org:service:ConnectionManager:1">
        '      <Source>file:*:audio/mpegurl:*,x-file-cifs:*:*:*,x-rincon:*:*:*,x-rincon-mp3radio:*:*:*,x-rincon-playlist:*:*:*,x-rincon-queue:*:*:*,x-rincon-stream:*:*:*</Source>
        '      <Sink></Sink>
        '    </u:GetProtocolInfoResponse>
        '  </s:Body>
        '</s:Envelope>
        If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPService.InvokeAction for ServiceID = " & MySCPDURL & " called with Action = " & bstrActionName, LogType.LOG_TYPE_INFO, LogColorGreen)
        InvokeAction = "NOK"
        Dim connectionManagerUri = New Uri(MycontrolURL)

        Dim ArgElement As String() = {}
        Dim Index As Integer = 0
        If vInActionArgs IsNot Nothing Then
            If UBound(vInActionArgs) > 0 Or vInActionArgs(0) IsNot Nothing Then
                If ActionList Is Nothing Then
                    If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPService.InvokeAction for ServiceID = " & MySCPDURL & " with Action = " & bstrActionName & ". Parameters were passed by service but action has no actionlist", LogType.LOG_TYPE_WARNING)
                    Throw New System.Exception("MyUPnPService.InvokeAction for ServiceID = " & MySCPDURL & " with Action = " & bstrActionName & ". Parameters were passed by service but action has no actionlist")
                    Exit Function
                End If
                Dim Action As MyUPnPAction = ActionList.Item(bstrActionName)
                If Action Is Nothing Then
                    If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPService.InvokeAction for ServiceID = " & MySCPDURL & " with Action = " & bstrActionName & ". Parameters were passed by service but action was not found", LogType.LOG_TYPE_WARNING)
                    Throw New System.Exception("MyUPnPService.InvokeAction for ServiceID = " & MySCPDURL & " with Action = " & bstrActionName & ". Parameters were passed by service but action was not found")
                    Exit Function
                End If
                If Action.argumentList Is Nothing Then
                    If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPService.InvokeAction for ServiceID = " & MySCPDURL & " with Action = " & bstrActionName & ". Parameters were passed by service but action has no argumentlist", LogType.LOG_TYPE_WARNING)
                    Throw New System.Exception("MyUPnPService.InvokeAction for ServiceID = " & MySCPDURL & " with Action = " & bstrActionName & ". Parameters were passed by service but action has no argumentlist")
                    Exit Function
                End If
                Index = 0
                Try
                    Dim ArgumentXML As New XmlDocument
                    ArgumentXML.LoadXml(Action.argumentList)
                    For Each Arg In vInActionArgs
                        If Arg IsNot Nothing Then
                            If ArgumentXML.GetElementsByTagName("direction").Item(Index).InnerText.ToUpper = "IN" Then  ' changed 2/26/2019 to capture a case where you have in, out, in parameters in definition (Sonos.Queue.CreateQueue)
                                'If ArgumentXML.GetElementsByTagName("direction").Item(Index).InnerText.ToUpper <> "IN" Then
                                'If UPnPDebuglevel > DebugLevel.dlErrorsOnly andAlso CheckDebugParam Then Log("MyUPnPService.InvokeAction for ServiceID = " & MySCPDURL & " with Action = " & bstrActionName & ". Parameters were passed but in/out mismatch at index = " & Index.ToString, LogType.LOG_TYPE_WARNING)
                                'Throw New System.Exception("MyUPnPService.InvokeAction for ServiceID = " & MySCPDURL & " with Action = " & bstrActionName & ". Parameters were passed but in/out mismatch at index = " & Index.ToString)
                                'Exit Function
                                'End If
                                Dim StateVariable As MyStateVariable = Nothing
                                Try
                                    Dim StateVariableType As String = ArgumentXML.GetElementsByTagName("relatedStateVariable").Item(Index).InnerText
                                    StateVariable = ServiceStateTable.Item(StateVariableType)
                                Catch ex As Exception
                                    If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPService.InvokeAction for ServiceID = " & MySCPDURL & " with Action = " & bstrActionName & ". couldn't find the relatedStateVariable at index = " & Index.ToString, LogType.LOG_TYPE_WARNING)
                                    StateVariable = Nothing
                                End Try
                                ReDim Preserve ArgElement(Index * 2 + 1)
                                ArgElement(Index * 2) = If(ArgumentXML.GetElementsByTagName("name").Item(Index)?.InnerText, "")
                                If StateVariable IsNot Nothing Then
                                    If StateVariable.dataType = VariableDataTypes.vdtBoolean Then
                                        If (Arg.ToString = "1") Or (Arg.ToString.ToLower = "true") Then
                                            ArgElement(Index * 2 + 1) = "1"
                                        Else
                                            ArgElement(Index * 2 + 1) = "0" ' anything invalid will be False
                                        End If
                                    Else
                                        ArgElement(Index * 2 + 1) = Trim(System.Web.HttpUtility.HtmlEncode(Arg.ToString))
                                    End If
                                Else
                                    ArgElement(Index * 2 + 1) = Trim(System.Web.HttpUtility.HtmlEncode(Arg.ToString))
                                End If
                                Index += 1
                            Else
                                If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPService.InvokeAction for ServiceID = " & MySCPDURL & " with Action = " & bstrActionName & ". Parameters were passed but out seen before all INs at index = " & Index.ToString, LogType.LOG_TYPE_WARNING)
                            End If
                        Else
                            If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPService.InvokeAction for ServiceID = " & MySCPDURL & " with Action = " & bstrActionName & ". Parameters were passed but where nil at index = " & Index.ToString, LogType.LOG_TYPE_WARNING)
                            Throw New System.Exception("MyUPnPService.InvokeAction for ServiceID = " & MySCPDURL & " with Action = " & bstrActionName & ". Parameters were passed but where nil at index = " & Index.ToString)
                            Exit Function
                        End If
                    Next
                Catch ex As Exception
                    If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("Error in MyUPnPService.InvokeAction for ServiceID = " & MySCPDURL & " with Action = " & bstrActionName & ". Parameters were passed by service but argumentlist matching caused error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    Throw New System.Exception("Error in MyUPnPService.InvokeAction for ServiceID = " & MySCPDURL & " with Action = " & bstrActionName & ". Parameters were passed by service but argumentlist matching caused error = " & ex.Message)
                    Exit Function
                End Try
            End If
        End If

        Dim getprotocol_request = _MakeRawSoapRequest(connectionManagerUri, MyServiceTypeIdentifier, bstrActionName, ArgElement)
        Dim ResponseHTML As String = ""
        Try
            Dim p = ServicePointManager.FindServicePoint(connectionManagerUri)
            p.Expect100Continue = False
        Catch ex As Exception
        End Try
        Dim webResponse As HttpWebResponse = Nothing
        Dim webStream As Stream = Nothing
        Dim wRequest As HttpWebRequest = Nothing
        Dim dataStream As Stream = Nothing

        Try
            wRequest = DirectCast(System.Net.HttpWebRequest.Create(connectionManagerUri), HttpWebRequest)
            wRequest.Method = "POST"
            wRequest.Host = connectionManagerUri.Authority
            wRequest.ProtocolVersion = HttpVersion.Version11
            wRequest.Headers.Add("SOAPAction", """" & MyServiceTypeIdentifier & "#" & bstrActionName & """")
            wRequest.ContentType = "text/xml; charset=""utf-8"""
            wRequest.ContentLength = getprotocol_request.Item2.Length
            If inCookieContainer IsNot Nothing Then
                If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MyUPnPService.InvokeAction has cookiecontainer", LogType.LOG_TYPE_INFO)
                wRequest.CookieContainer = inCookieContainer
            End If
            wRequest.KeepAlive = False
            ' Get the request stream.
            dataStream = wRequest.GetRequestStream()
            ' Write the data to the request stream.
            dataStream.Write(getprotocol_request.Item2, 0, getprotocol_request.Item2.Length)
            ' Close the Stream object.
            dataStream.Close()
            ' Get the response.
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPService.InvokeAction for ServiceID = " & MySCPDURL & " while preparing to send Action = " & bstrActionName & " for URI = " & connectionManagerUri.AbsoluteUri & " and Request = " & System.Text.Encoding.UTF8.GetString(getprotocol_request.Item2) & " with error = " & ex.Message, LogType.LOG_TYPE_WARNING)
            Throw New System.Exception("MyUPnPService.InvokeAction for ServiceID = " & MySCPDURL & " while preparing to send Action = " & bstrActionName & " for URI = " & connectionManagerUri.AbsoluteUri & " and Request = " & System.Text.Encoding.UTF8.GetString(getprotocol_request.Item2) & " with error = " & ex.Message)
            Exit Function
        End Try
        Try
            webResponse = wRequest.GetResponse
        Catch ex As WebException
            webResponse = ex.Response
            If webResponse IsNot Nothing Then
                webStream = webResponse.GetResponseStream
                Dim strmRdr As New System.IO.StreamReader(webStream)
                ResponseHTML = strmRdr.ReadToEnd()
                strmRdr.Dispose()
            End If
            If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPService.InvokeAction for ServiceID = " & MySCPDURL & " while sending Action = " & bstrActionName & " for URI = " & connectionManagerUri.AbsoluteUri & " and Request = " & System.Text.Encoding.UTF8.GetString(getprotocol_request.Item2) & " UPNP Error = " & TreatUPnPError(ResponseHTML) & " with error = " & ex.Message, LogType.LOG_TYPE_WARNING)
            Throw New System.Exception("MyUPnPService.InvokeAction for ServiceID = " & MySCPDURL & " while sending Action = " & bstrActionName & " for URI = " & connectionManagerUri.AbsoluteUri & " and Request = " & System.Text.Encoding.UTF8.GetString(getprotocol_request.Item2) & " UPNP Error = " & TreatUPnPError(ResponseHTML) & " with error = " & ex.Message)
            Exit Function
        End Try
        Try
            webStream = webResponse.GetResponseStream
            Dim strmRdr As New System.IO.StreamReader(webStream)
            ResponseHTML = strmRdr.ReadToEnd()
            strmRdr.Dispose()
            If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MyUPnPService.InvokeAction for ServiceID = " & MySCPDURL & " send Action = " & bstrActionName & " and received = " & ResponseHTML, LogType.LOG_TYPE_INFO, LogColorGreen)
            webStream.Close()
            webStream.Dispose()
            webResponse.Close()
            webStream = Nothing
            webResponse = Nothing
            wRequest = Nothing
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPService.InvokeAction for ServiceID = " & MySCPDURL & " while processing response for Action = " & bstrActionName & " for URI = " & connectionManagerUri.AbsoluteUri & " and Request = " & System.Text.Encoding.UTF8.GetString(getprotocol_request.Item2) & " with error = " & ex.Message, LogType.LOG_TYPE_WARNING)
            Throw New System.Exception("MyUPnPService.InvokeAction for ServiceID = " & MySCPDURL & " while processing response for Action = " & bstrActionName & " for URI = " & connectionManagerUri.AbsoluteUri & " and Request = " & System.Text.Encoding.UTF8.GetString(getprotocol_request.Item2) & " with error = " & ex.Message)
            Exit Function
        End Try
        dataStream.Dispose()
        dataStream = Nothing

        If ResponseHTML = "" Then
            If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPService.InvokeAction for ServiceID = " & MySCPDURL & " for Action = " & bstrActionName & ". Empty Response", LogType.LOG_TYPE_WARNING)
            Throw New System.Exception("MyUPnPService.InvokeAction for ServiceID = " & MySCPDURL & " for Action = " & bstrActionName & ". Empty Response")
            Exit Function
        End If

        Dim ResponseXML As New XmlDocument

        Try
            ResponseXML.LoadXml(ResponseHTML)
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("Error in MyUPnPService.InvokeAction for ServiceID = " & MySCPDURL & " for Action = " & bstrActionName & " loading XML with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Throw New System.Exception("Error in MyUPnPService.InvokeAction for ServiceID = " & MySCPDURL & " for Action = " & bstrActionName & " loading XML with Error = " & ex.Message)
            Exit Function
        End Try
        '<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/" s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
        '  <s:Body>
        '    <u:GetProtocolInfoResponse xmlns:u="urn:schemas-upnp-org:service:ConnectionManager:1">
        '      <Source>file:*:audio/mpegurl:*,x-file-cifs:*:*:*,x-rincon:*:*:*,x-rincon-mp3radio:*:*:*,x-rincon-playlist:*:*:*,x-rincon-queue:*:*:*,x-rincon-stream:*:*:*</Source>
        '      <Sink></Sink>
        '    </u:GetProtocolInfoResponse>
        '  </s:Body>
        '</s:Envelope>
        Dim ResponseNode As XmlNode = Nothing
        Try
            If ResponseXML IsNot Nothing Then
                If ResponseXML.HasChildNodes Then
                    For Each Element As XmlNode In ResponseXML.ChildNodes
                        'If upnpDebuglevel > DebugLevel.dlErrorsOnly andAlso CheckDebugParam Then Log("MyUPnPService.InvokeAction for ServiceID = " & MyFullServiceName & " for Action = " & bstrActionName & " found node with Name = " & Element.Name & " and LocalName= " & Element.LocalName, LogType.LOG_TYPE_WARNING)
                        If Element.LocalName = "Envelope" Then
                            If Element.HasChildNodes Then
                                For Each ChildElement As XmlNode In Element.ChildNodes
                                    If ChildElement.LocalName = "Body" Then
                                        If ChildElement.HasChildNodes Then
                                            For Each GrandChild As XmlNode In ChildElement.ChildNodes
                                                If GrandChild.LocalName = (bstrActionName & "Response") Then
                                                    ResponseNode = GrandChild
                                                    Exit Try
                                                End If
                                            Next
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("Error in MyUPnPService.InvokeAction for ServiceID = " & MySCPDURL & " for Action = " & bstrActionName & " Processing nodes with HTML = " & ResponseHTML & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        ResponseXML = Nothing
        Index = 0
        Try
            If ResponseNode IsNot Nothing Then
                If ResponseNode.HasChildNodes Then
                    For Each Response As XmlNode In ResponseNode.ChildNodes
                        If Index <= UBound(pvOutActionArgs) Then
                            pvOutActionArgs(Index) = Response.InnerText
                        Else
                            If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPService.InvokeAction for ServiceID = " & MySCPDURL & " for Action = " & bstrActionName & ". More response parameters than expected with response = " & ResponseNode.OuterXml, LogType.LOG_TYPE_WARNING)
                        End If
                        If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MyUPnPService.InvokeAction for ServiceID = " & MySCPDURL & " for Action = " & bstrActionName & " extracted response = " & Response.InnerText & " and with response = " & ResponseNode.OuterXml, LogType.LOG_TYPE_INFO, LogColorGreen)
                        Index += 1
                    Next
                End If
            Else
                If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MyUPnPService.InvokeAction for ServiceID = " & MySCPDURL & " for Action = " & bstrActionName & " found no responses in HTML = " & ResponseHTML, LogType.LOG_TYPE_WARNING)
            End If
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("Error in MyUPnPService.InvokeAction for ServiceID = " & MySCPDURL & " for Action = " & bstrActionName & " unable to process response with HTML = " & ResponseHTML & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Throw New System.Exception("Error in MyUPnPService.InvokeAction for ServiceID = " & MySCPDURL & " for Action = " & bstrActionName & " unable to process response with HTML = " & ResponseHTML & " and Error = " & ex.Message)
            Exit Function
        End Try
        If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MyUPnPService.InvokeAction for ServiceID = " & MySCPDURL & " called with Action = " & bstrActionName & " ended successfuly with " & Index.ToString & " arguments retrieved", LogType.LOG_TYPE_INFO, LogColorGreen)
        InvokeAction = "OK"
        ResponseNode = Nothing
    End Function

    Function QueryStateVariable(bstrVariableName As String) As String
        If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPService.QueryStateVariable called for ServiceID = " & MySCPDURL & " with VariableName = " & bstrVariableName, LogType.LOG_TYPE_INFO, LogColorGreen)
        QueryStateVariable = Nothing
        Dim StateVariable As MyStateVariable = Nothing
        Try
            StateVariable = ServiceStateTable.Item(bstrVariableName)
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPService.QueryStateVariable for ServiceID = " & MySCPDURL & " with VariableName = " & bstrVariableName & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Throw New System.Exception("Error in MyUPnPService.QueryStateVariable for ServiceID = " & MySCPDURL & " with VariableName = " & bstrVariableName & " and Error = " & ex.Message)
        End Try
        If StateVariable IsNot Nothing AndAlso StateVariable.hasValue Then ' added 11/10/2019
            QueryStateVariable = StateVariable.value
        Else
            If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("Warning in MyUPnPService.QueryStateVariable for ServiceID = " & MySCPDURL & " with VariableName = " & bstrVariableName & ". No value received", LogType.LOG_TYPE_WARNING)
            Throw New System.Exception("Warning in MyUPnPService.QueryStateVariable for ServiceID = " & MySCPDURL & " with VariableName = " & bstrVariableName & ". No Value")
        End If
    End Function

    Private Sub UpdateStateVariable(bstrVariableName As String, VariableValue As String)
        If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MyUPnPService.UpdateStateVariable called for ServiceID = " & MySCPDURL & " with VariableName = " & bstrVariableName & " and VariableValue = " & VariableValue, LogType.LOG_TYPE_INFO, LogColorGreen)
        Try
            Dim StateVariable As MyStateVariable = ServiceStateTable.Item(bstrVariableName)
            StateVariable.value = VariableValue
            StateVariable.hasValue = True
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("Error in MyUPnPService.UpdateStateVariable for ServiceID = " & MySCPDURL & " with VariableName = " & bstrVariableName & " and VariableValue = " & VariableValue & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub TreatEventQueue()
        If EventHandlerReEntryFlag Then
            MissedEventHandlerFlag = True
            Exit Sub
        End If
        EventHandlerReEntryFlag = True
        While MyEventQueue.Count > 0
            Dim NewEvent As String = ""
            SyncLock (MyEventQueue)
                NewEvent = MyEventQueue.Dequeue
            End SyncLock
            If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MyUPnPService.TreatEventQueue for ServiceID = " & MySCPDURL & " is processing Event = " & NewEvent, LogType.LOG_TYPE_INFO, LogColorGreen)
            Dim EventXMLDoc As New XmlDocument
            Try
                Dim NotifyData As String = ""
                Dim SID As String = ""
                Dim NTReceived As String = ""
                Dim NTSReceived As String = ""
                Dim BodyReceived As String = ""
                Dim BootID As String = ""
                Dim ConfigID As String = ""
                Dim SearchPort As String = ""
                Dim SEQ As String = ""
                Try
                    NotifyData = ParseHTTPResponse(NewEvent, "notify")
                    SID = System.Web.HttpUtility.HtmlDecode(ParseHTTPResponse(NewEvent, "sid:"))
                    NTReceived = ParseHTTPResponse(NewEvent, "nt:")
                    NTSReceived = ParseHTTPResponse(NewEvent, "nts:")
                    BootID = ParseHTTPResponse(NewEvent, "bootid.upnp.org:")
                    If BootID = "" Then BootID = ParseHTTPResponse(NewEvent, "x-rincon-bootseq:")     ' added 10/13/2019
                    If BootID <> myBootId Then
                        If myBootId <> "" And upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPService.TreatEventQueue for ServiceID = " & MySCPDURL & " for SID = " & NotifyData & " and NT = " & NTReceived & " and NTS = " & NTSReceived & " had different bootids. Stored = " & myBootId & " <-> received = " & BootID, LogType.LOG_TYPE_INFO, LogColorGreen)
                        myBootId = BootID
                    End If
                    ConfigID = ParseHTTPResponse(NewEvent, "configid.upnp.org:")
                    SearchPort = ParseHTTPResponse(NewEvent, "searchport.upnp.org:")
                    SEQ = ParseHTTPResponse(NewEvent, "seq:")
                    BodyReceived = ParseHTTPResponse(NewEvent, "")
                    If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPService.TreatEventQueue for ServiceID = " & MySCPDURL & " for SID = " & NotifyData & " and NT = " & NTReceived & " and NTS = " & NTSReceived, LogType.LOG_TYPE_INFO, LogColorGreen)
                Catch ex As Exception
                    If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPService.TreatEventQueue for ServiceID = " & MySCPDURL & " while analyzing data with error = " & ex.Message & " and Data = " & NewEvent, LogType.LOG_TYPE_ERROR)
                End Try

                If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MyUPnPService.TreatEventQueue for ServiceID = " & MySCPDURL & " is processing NotifyData = " & NotifyData & " with SEQ = " & SEQ.ToString, LogType.LOG_TYPE_INFO, LogColorGreen)

                If upnpDebuglevel > DebugLevel.dlEvents Then
                    'If upnpDebuglevel > DebugLevel.dlErrorsOnly andAlso CheckDebugParam Then Log("MyUPnPService.HandleDataReceive found HTTP version = " & ParseHTTPResponse(Data, "http/"), LogType.LOG_TYPE_INFO)
                    Log("MyUPnPService.TreatEventQueue found NOTIFY = " & NotifyData, LogType.LOG_TYPE_INFO, LogColorGreen)
                    Log("MyUPnPService.TreatEventQueue found HOST = " & ParseHTTPResponse(NewEvent, "host:"), LogType.LOG_TYPE_INFO, LogColorGreen)
                    Log("MyUPnPService.TreatEventQueue found CONTENT-TYPE = " & ParseHTTPResponse(NewEvent, "content-type:"), LogType.LOG_TYPE_INFO, LogColorGreen)
                    Log("MyUPnPService.TreatEventQueue found CONTENT-LENGTH = " & ParseHTTPResponse(NewEvent, "content-length:"), LogType.LOG_TYPE_INFO, LogColorGreen)
                    Log("MyUPnPService.TreatEventQueue found NT = " & NTReceived, LogType.LOG_TYPE_INFO, LogColorGreen)
                    Log("MyUPnPService.TreatEventQueue found NTS = " & NTSReceived, LogType.LOG_TYPE_INFO, LogColorGreen)
                    Log("MyUPnPService.TreatEventQueue found SID = " & SID, LogType.LOG_TYPE_INFO, LogColorGreen)
                    Log("MyUPnPService.TreatEventQueue found BootID = " & BootID, LogType.LOG_TYPE_INFO, LogColorGreen)
                    Log("MyUPnPService.TreatEventQueue found ConfigID = " & ConfigID, LogType.LOG_TYPE_INFO, LogColorGreen)
                    Log("MyUPnPService.TreatEventQueue found SearchPort = " & SearchPort, LogType.LOG_TYPE_INFO, LogColorGreen)
                    Log("MyUPnPService.TreatEventQueue found body = " & BodyReceived, LogType.LOG_TYPE_INFO, LogColorGreen)
                End If

                If SEQ <> "" Then
                    If Val(SEQ) <> LastEventSequenceNumber + 1 Then
                        If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MyUPnPService.TreatEventQueue for ServiceID = " & MySCPDURL & " received an out of sequence event with Seq = " & SEQ & " and Expected = " & (LastEventSequenceNumber + 1).ToString & " and Data = " & NewEvent, LogType.LOG_TYPE_WARNING)
                        If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPService.TreatEventQueue for ServiceID = " & MySCPDURL & " received an out of sequence event with Seq = " & SEQ & " and Expected = " & (LastEventSequenceNumber + 1).ToString, LogType.LOG_TYPE_WARNING)
                    End If
                    LastEventSequenceNumber = Val(SEQ)
                End If

                If NTReceived <> "upnp:event" Or NTSReceived <> "upnp:propchange" Then
                    If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPService.TreatEventQueue received for ServiceID = " & MySCPDURL & " the wrong NotificationEvent NT = " & NTReceived & ", NTS = " & NTSReceived & " and Data = " & NewEvent, LogType.LOG_TYPE_WARNING)
                    If MissedEventHandlerFlag Then MyEventTimer.Enabled = True ' rearm the timer to prevent events from getting lost
                    EventHandlerReEntryFlag = False
                    Exit Sub
                End If

                EventXMLDoc.LoadXml(BodyReceived)
                Dim stateVariable As String = ""
                Dim EventXML As String = ""

                Dim PropertyNodes As XmlNodeList = Nothing
                Try
                    If EventXMLDoc IsNot Nothing Then
                        If EventXMLDoc.HasChildNodes Then
                            For Each child As XmlNode In EventXMLDoc
                                If child.LocalName = "propertyset" Then
                                    If child.HasChildNodes Then
                                        For Each grandchild As XmlNode In child.ChildNodes
                                            If grandchild.LocalName = "property" Then
                                                If grandchild.HasChildNodes Then
                                                    For Each VariableName As XmlNode In grandchild.ChildNodes
                                                        If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPService.TreatEventQueue raised event for ServiceID = " & MySCPDURL & ", NotifyData = " & NotifyData & ", DeviceUDN = " & MyDeviceUDN & ", IPAddress = " & MyIPAddress & " and Event = " & VariableName.LocalName & ", Event InfoLength = " & VariableName.InnerXml.Length.ToString, LogType.LOG_TYPE_INFO, LogColorGreen)
                                                        UpdateStateVariable(VariableName.LocalName, Trim(System.Web.HttpUtility.HtmlDecode(VariableName.InnerXml)))
                                                        RaiseEvent StateChange(VariableName.LocalName, Trim(System.Web.HttpUtility.HtmlDecode(VariableName.InnerXml)))
                                                    Next
                                                End If
                                            End If
                                        Next
                                    End If
                                End If
                            Next
                        End If
                    End If
                    'PropertyNodes = EventXMLDoc.GetElementsByTagName("e:property")
                Catch ex As Exception
                    If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("MyUPnPService.TreatEventQueue for ServiceID = " & MySCPDURL & " couldn't find the property nodes in the Notify Data = " & BodyReceived & " with Error = " & ex.Message, LogType.LOG_TYPE_WARNING)
                End Try
            Catch ex As Exception
                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPService.TreatEventQueue loading for ServiceID = " & MySCPDURL & " new Event = " & NewEvent & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End While
        If MissedEventHandlerFlag Then MyEventTimer.Enabled = True ' rearm the timer to prevent events from getting lost
        EventHandlerReEntryFlag = False
    End Sub

    Public Sub RemoveCallback()
        ' probably will have to reset the callbacks here and any outstanding TCP ports/requests
        If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPService.RemoveCallback called for ServiceID = " & MySCPDURL, LogType.LOG_TYPE_INFO, LogColorGreen)
        'If UPnPDebuglevel > DebugLevel.dlOff andAlso CheckDebugParam Then Log("MyUPnPService.RemoveCallback called for ServiceID = " & MySCPDURL, LogType.LOG_TYPE_INFO, LogColorGreen)
        Try
            If StateVariableChangedHandlerAddress IsNot Nothing Then RemoveHandler StateChange, StateVariableChangedHandlerAddress
            If ServiceDiedChangedHandlerAddress IsNot Nothing Then RemoveHandler ServiceDied, ServiceDiedChangedHandlerAddress
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("Error in MyUPnPService.RemoveCallback called for ServiceID = " & MySCPDURL & " removing the EventHandlers with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        StateVariableChangedHandlerAddress = Nothing
        ServiceDiedChangedHandlerAddress = Nothing
        isServiceEventsSubscribed = False
        StopTimers(True)
    End Sub

    Public Sub Dispose()
        ' probably will have to reset the callbacks here and any outstanding TCP ports/requests
        If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MyUPnPService.Dispose called for ServiceID = " & MySCPDURL, LogType.LOG_TYPE_INFO, LogColorGreen)
        Try
            If StateVariableChangedHandlerAddress IsNot Nothing Then RemoveHandler StateChange, StateVariableChangedHandlerAddress
            If ServiceDiedChangedHandlerAddress IsNot Nothing Then RemoveHandler ServiceDied, ServiceDiedChangedHandlerAddress
            StateVariableChangedHandlerAddress = Nothing
            ServiceDiedChangedHandlerAddress = Nothing
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("Error in MyUPnPService.Dispose called for ServiceID = " & MySCPDURL & " removing the EventHandlers with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        StopTimers(True)
        If ServiceStateTable IsNot Nothing Then
            ServiceStateTable.Dispose()
        End If
    End Sub

    Private Function MakeURLWhole(inURL As String) As String
        MakeURLWhole = inURL
        ' the inURL can either be complete, meaning it starts with http: urn: file: ... or it is partial and then it SHOULD start with an / so we complete it with IP address and IPPort
        inURL = Trim(inURL) ' remove blanks
        If inURL = "" Then
            If MyIPPort <> "" Then Return "http://" & MyIPAddress & ":" & MyIPPort Else Return "http://" & MyIPAddress
        End If
        Try
            Dim FullUri As New Uri(inURL)
            If FullUri.IsAbsoluteUri And Trim(FullUri.Host) <> "" Then
                If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPService.MakeURLWhole for ServiceID = " & MySCPDURL & " for inURL = " & inURL & " was complete with Host = " & FullUri.Host.ToString & " and type = " & FullUri.HostNameType.ToString, LogType.LOG_TYPE_INFO, LogColorGreen)
                FullUri = Nothing
                Return inURL
            End If
        Catch ex As Exception
            ' if it comes here, it is not a FullURI
        End Try
        Try
            Dim FullUri As New Uri(MyLocation)
            Dim newURI As New Uri(FullUri, inURL)
            If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MyUPnPService.MakeURLWhole for ServiceID = " & MySCPDURL & " for inURL = " & inURL & " returned = " & newURI.AbsoluteUri, LogType.LOG_TYPE_INFO, LogColorGreen)
            If newURI.IsAbsoluteUri Then Return newURI.AbsoluteUri
        Catch ex As Exception
            ' if it comes here, it is not a FullURI
        End Try
        If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MyUPnPService.MakeURLWhole called for ServiceID = " & MySCPDURL & " for inURL = " & inURL & ", IPPort = " & MyIPPort & " and IPAddress = " & MyIPAddress, LogType.LOG_TYPE_INFO, LogColorGreen)
        Try
            If inURL.IndexOf("/") <> 0 Then
                ' add the "/" character
                If MyIPPort <> "" Then Return "http://" & MyIPAddress & ":" & MyIPPort & "/" & inURL Else Return "http://" & MyIPAddress & "/" & inURL
            Else
                ' this URL started with the right / character
                If MyIPPort <> "" Then Return "http://" & MyIPAddress & ":" & MyIPPort & inURL Else Return "http://" & MyIPAddress & inURL
            End If
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPService.MakeURLWhole for ServiceID = " & MySCPDURL & " for inURL = " & inURL & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Function MakeRawSoapRequest(requestUri As Uri, soapSchema As String, soapVerb As String,
                            args As String()) As Tuple(Of Uri, Byte())
        'Dim soapSchema = soapAction.Name.NamespaceName
        'Dim soapVerb = soapAction.Name.LocalName
        Dim argpairs As New List(Of Tuple(Of String, String))
        For i = 0 To args.Length - 1 Step 2
            argpairs.Add(Tuple.Create(args(i), args(i + 1)))
        Next

        Dim s = "POST " & requestUri.PathAndQuery & " HTTP/1.1" & vbCrLf &
                "Host: " & requestUri.Authority & vbCrLf &
                "Content-Length: ?" & vbCrLf &
                "Content-Type: text/xml; charset=""utf-8""" & vbCrLf &
                "SOAPAction: """ & soapSchema & "#" & soapVerb & """" & vbCrLf &
                "" & vbCrLf &
                "<?xml version=""1.0""?>" & vbCrLf &
                "<s:Envelope xmlns:s=""http://schemas.xmlsoap.org/soap/" &
                "envelope/"" s:encodingStyle=""http://schemas.xmlsoap.org/soap/encoding/"">" & vbCrLf &
                "  <s:Body>" & vbCrLf &
                "    <u:" & soapVerb & " xmlns:u=""" & soapSchema & """>" & vbCrLf &
                String.Join(vbCrLf, (From arg In argpairs Select "      <" &
                arg.Item1 & ">" & arg.Item2 & "</" & arg.Item1 & ">").Concat({""})) &
                "    </u:" & soapVerb & ">" & vbCrLf &
                "  </s:Body>" & vbCrLf &
                "</s:Envelope>" & vbCrLf
        '
        Dim len = System.Text.Encoding.UTF8.GetByteCount(s.Substring(s.IndexOf("<?xml")))
        s = s.Replace("Content-Length: ?", "Content-Length: " & len)
        Return Tuple.Create(requestUri, System.Text.Encoding.UTF8.GetBytes(s))
    End Function

    Function _MakeRawSoapRequest(requestUri As Uri, soapSchema As String, soapVerb As String,
                        args As String()) As Tuple(Of Uri, Byte())
        'Dim soapSchema = soapAction.Name.NamespaceName
        'Dim soapVerb = soapAction.Name.LocalName
        Dim argpairs As New List(Of Tuple(Of String, String))
        For i = 0 To args.Length - 1 Step 2
            argpairs.Add(Tuple.Create(args(i), args(i + 1)))
        Next

        Dim s = "<?xml version=""1.0"" encoding=""utf-8""?>" & vbCrLf &
                "<s:Envelope xmlns:s=""http://schemas.xmlsoap.org/soap/" &
                "envelope/"" s:encodingStyle=""http://schemas.xmlsoap.org/soap/encoding/"">" & vbCrLf &
                "  <s:Body>" & vbCrLf &
                "    <u:" & soapVerb & " xmlns:u=""" & soapSchema & """>" & vbCrLf &
                String.Join(vbCrLf, (From arg In argpairs Select "      <" &
                arg.Item1 & ">" & arg.Item2 & "</" & arg.Item1 & ">").Concat({""})) &
                "    </u:" & soapVerb & ">" & vbCrLf &
                "  </s:Body>" & vbCrLf &
                "</s:Envelope>" & vbCrLf
        Return Tuple.Create(requestUri, System.Text.Encoding.UTF8.GetBytes(s))
    End Function

    Function MakeSubscribeSoapRequest(requestUri As Uri, eventSubURL As String, CALLBACK As String) As Tuple(Of Uri, Byte())
        ' this is how to subscribe to events from the Device Spy Capture
        'SUBSCRIBE /MediaRenderer/AVTransport/Event HTTP/1.1
        'NT: upnp:event
        'HOST: 192.168.1.104:1400
        'CALLBACK: <http://192.168.1.110:6062/RINCON_000E5859008A01400_MR/urn:upnp-org:serviceId:AVTransport>
        'TIMEOUT: Second-300
        'Content-Length: 0

        Dim s = "SUBSCRIBE " & eventSubURL & " HTTP/1.1" & vbCrLf &
                "NT: upnp:event" & vbCrLf &
                "Host: " & requestUri.Authority & vbCrLf &
                "CALLBACK: " & CALLBACK & vbCrLf &
                "TIMEOUT: Second-300" & vbCrLf &
                "Content-Length: 0" & vbCrLf
        ' "Content-Type: text/xml; charset=""utf-8""" & vbCrLf

        Return Tuple.Create(requestUri, System.Text.Encoding.UTF8.GetBytes(s))
    End Function

    Private Sub MySubscribeRenewTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles MySubscribeRenewTimer.Elapsed
        If upnpDebuglevel > DebugLevel.dlEvents AndAlso CheckDebugParam Then Log("MyUPnPService.MySubscribeRenewTimer_Elapsed called for ServiceID = " & MySCPDURL, LogType.LOG_TYPE_INFO, LogColorGreen)
        Try
            SendRenew()
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPService.MySubscribeRenewTimer_Elapsed called for ServiceID = " & MySCPDURL & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        e = Nothing
        sender = Nothing
    End Sub

    Private Sub MyMissedRenewTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles MyMissedRenewTimer.Elapsed
        If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("MyUPnPService.MyMissedRenewTimer_Elapsed called for ServiceID = " & MySCPDURL, LogType.LOG_TYPE_INFO, LogColorNavy)
        Try
            SendRenew()
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPService.MyMissedRenewTimer_Elapsed called for ServiceID = " & MySCPDURL & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        e = Nothing
        sender = Nothing
    End Sub

    Private Sub MyEventTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles MyEventTimer.Elapsed
        Try
            TreatEventQueue()
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("Error in MyUPnPService.MyEventTimer_Elapsed for ServiceID = " & MySCPDURL & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        e = Nothing
        sender = Nothing
    End Sub

    Private Sub MyServiceDiedTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles MyServiceDiedTimer.Elapsed
        Try
            StopTimers(False)
            Try
                RaiseEvent ServiceDied()
            Catch ex1 As Exception
                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPService.MyServiceDiedTimer_Elapsed raising the event for ServiceID = " & MySCPDURL & " with Error = " & ex1.Message, LogType.LOG_TYPE_ERROR)
            End Try
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPService.MyServiceDiedTimer_Elapsed for ServiceID = " & MySCPDURL & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        e = Nothing
        sender = Nothing
    End Sub

    Private Function TreatUPnPError(ErrorInfo As String) As String
        'HTTP/1.1 500 Internal Server Error
        'CONTENT-LENGTH: 347
        'CONTENT-TYPE: text/xml; charset="utf-8"
        'EXT:
        'Server: Linux UPnP/1.0 Sonos/24.0-69180 
        'Connection: close()
        '<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/" s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
        '  <s:Body>
        '    <s:Fault>
        '      <faultcode>s:Client</faultcode>
        '      <faultstring>UPnPError</faultstring>
        '      <detail>
        '         <UPnPError xmlns="urn:schemas-upnp-org:control-1-0">
        '            <errorCode>800</errorCode>
        '         </UPnPError>
        '      </detail>
        '    </s:Fault>
        '  </s:Body>
        '</s:Envelope>
        TreatUPnPError = ErrorInfo
        If ErrorInfo = "" Then Exit Function
        Dim ResponseXML As New XmlDocument
        Try
            ResponseXML.LoadXml(ErrorInfo)
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("Error in MyUPnPService.TreatUPnPError for ServiceID = " & MySCPDURL & " loading XML with ErrorInfo = " & ErrorInfo & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        Dim ResponseNode As XmlNode = Nothing
        Dim RetErrorInfo As String = ""
        Try
            If ResponseXML IsNot Nothing Then
                If ResponseXML.HasChildNodes Then
                    For Each Element As XmlNode In ResponseXML.ChildNodes
                        'If upnpDebuglevel > DebugLevel.dlErrorsOnly andAlso CheckDebugParam Then Log("MyUPnPService.InvokeAction for ServiceID = " & MyFullServiceName & " for Action = " & bstrActionName & " found node with Name = " & Element.Name & " and LocalName= " & Element.LocalName, LogType.LOG_TYPE_WARNING)
                        If Element.LocalName = "Envelope" Then
                            If Element.HasChildNodes Then
                                For Each ChildElement As XmlNode In Element.ChildNodes
                                    If ChildElement.LocalName = "Body" Then
                                        If ChildElement.HasChildNodes Then
                                            For Each GrandChild As XmlNode In ChildElement.ChildNodes
                                                If GrandChild.LocalName = "Fault" Then
                                                    If GrandChild.HasChildNodes Then
                                                        For Each ErrorRecord As XmlNode In GrandChild.ChildNodes
                                                            If RetErrorInfo <> "" Then RetErrorInfo &= ", "
                                                            RetErrorInfo &= ErrorRecord.LocalName & " = " & ErrorRecord.InnerText
                                                        Next
                                                    End If
                                                End If
                                            Next
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPService.TreatUPnPError for ServiceID = " & MySCPDURL & " for ErrorInfo = " & ErrorInfo & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        TreatUPnPError = RetErrorInfo
    End Function

    Private Function RetrieveTimeoutData(inData As String) As Integer
        RetrieveTimeoutData = 300
        Dim Timeout As Integer = 300
        Try
            Dim TimeoutParms As String() = Split(inData, "-") ' first part is Seconds ... or should be!! Second part could be an integer or "infinite"
            If UBound(TimeoutParms) <= 0 Then
                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("MyUPnPService.RetrieveTimeoutData called for ServiceID = " & MySCPDURL & " and received an Invalid TIMEOUT = " & inData, LogType.LOG_TYPE_ERROR)
                Exit Function
            End If
            If Trim(TimeoutParms(0).ToUpper) <> "SECOND" Then
                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("MyUPnPService.RetrieveTimeoutData called for ServiceID = " & MySCPDURL & " and received an Invalid TIMEOUT = " & inData, LogType.LOG_TYPE_ERROR)
                Exit Function
            End If
            If Trim(TimeoutParms(1).ToUpper) = "INFINITE" Then
                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("MyUPnPService.RetrieveTimeoutData called for ServiceID = " & MySCPDURL & " and received an infine timeout request, changed to 300 seconds, received = " & inData, LogType.LOG_TYPE_WARNING)
                Exit Function
            Else
                Timeout = Val(Trim(TimeoutParms(1)))
                If Timeout = 0 Then
                    If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("MyUPnPService.RetrieveTimeoutData called for ServiceID = " & MySCPDURL & " and received an invalid TIMEOUT = " & inData, LogType.LOG_TYPE_ERROR)
                    Exit Function
                End If
                If Timeout <= 60 Then
                    If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("MyUPnPService.RetrieveTimeoutData called for ServiceID = " & MySCPDURL & " and received a very small timeout request, which could case a lot of network traffic. Received = " & inData, LogType.LOG_TYPE_WARNING)
                End If
            End If
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlErrorsOnly AndAlso CheckDebugParam Then Log("Error in MyUPnPService.RetrieveTimeoutData for ServiceID = " & MySCPDURL & ". Unable to extract the timeout info with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        Return Timeout
    End Function

    Public Sub ServiceDiedReceived()
        If MyServiceDiedTimer Is Nothing Then
            Try
                MyServiceDiedTimer = New Timers.Timer With {
                    .Interval = 100, ' 100 msecond
                    .AutoReset = False,
                    .Enabled = True
                }
            Catch ex As Exception
                If upnpDebuglevel > DebugLevel.dlOff AndAlso CheckDebugParam Then Log("Error in MyUPnPService.ServiceDiedReceived. Unable to create the Event timer for ServiceID = " & MySCPDURL & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If
    End Sub
End Class

<Serializable()>
Public Class MyUPnPServices

    Inherits List(Of MyUPnPService)

    Overloads ReadOnly Property Item(bstrServiceId As String) As MyUPnPService
        Get
            Item = Nothing
            Try
                If Me.Count > 0 Then
                    For Each Service As MyUPnPService In Me
                        If Service IsNot Nothing Then
                            If Service.Id = bstrServiceId Then
                                Item = Service
                                Exit Property
                            End If
                        End If
                    Next
                End If
            Catch ex As Exception
                If upnpDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in MyUPnPServices.Item retrieving a service = " & bstrServiceId & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End Get
    End Property

    Public ReadOnly Property GetServiceType(bstrServiceType As String) As MyUPnPService
        Get
            GetServiceType = Nothing
            Try
                If Me.Count > 0 Then
                    For Each Service As MyUPnPService In Me
                        If Service IsNot Nothing Then
                            If Service.ServiceType = bstrServiceType Then
                                GetServiceType = Service
                                Exit Property
                            End If
                        End If
                    Next
                End If
            Catch ex As Exception
                If upnpDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in MyUPnPServices.GetServiceType retrieving a Service = " & bstrServiceType & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End Get
    End Property

    Public Overloads Sub Add(NewService As MyUPnPService)
        MyBase.Add(NewService)
        'If UPnPDebuglevel > DebugLevel.dlEvents andAlso CheckDebugParam Then Log("MyUPnPServices.add added new service = " & NewService.Id, LogType.LOG_TYPE_INFO)
    End Sub

    Public Sub Dispose()
        If upnpDebuglevel > DebugLevel.dlEvents Then Log("MyUPnPServices.Dispose called", LogType.LOG_TYPE_INFO, LogColorGreen)
        Try
            For Each Service As MyUPnPService In Me
                Try
                    Service.Dispose()
                    Service = Nothing
                Catch ex As Exception
                End Try
            Next
            Clear()
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in MyUPnPServices.Dispose with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub
End Class

Public Class MyServiceStateTable
    Inherits List(Of MyStateVariable)

    Overloads ReadOnly Property Item(bstrVariableName As String) As MyStateVariable
        Get
            Item = Nothing
            Try
                If Me.Count > 0 Then
                    For Each ServiceState As MyStateVariable In Me
                        If ServiceState IsNot Nothing Then
                            If ServiceState.name = bstrVariableName Then
                                Item = ServiceState
                                Exit Property
                            End If
                        End If
                    Next
                End If
            Catch ex As Exception
                If upnpDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in MyServiceStateTable.Item retrieving a ServiceState = " & bstrVariableName & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End Get
    End Property

    Public Overloads Sub Add(NewStateVariable As MyStateVariable)
        'If MyServiceStateList Is Nothing Then Exit Sub
        Try
            MyBase.Add(NewStateVariable)
            'If UPnPDebuglevel > DebugLevel.dlEvents Then Log("MyServiceStateTable.Add added a new Variable = " & NewStateVariable.name, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff Then Log("Error in MyServiceStateTable.Add when called with Variable = " & NewStateVariable.name & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub Dispose()
        'If MyServiceStateList Is Nothing Then Exit Sub
        Try
            For Each StateVariable As MyStateVariable In Me
                Try
                    StateVariable.Dispose()
                    StateVariable = Nothing
                Catch ex As Exception
                End Try
            Next
            Clear()
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff Then Log("Error in MyServiceStateTable.dispose with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

End Class

Public Class MyStateVariable
    Public sendEvents As Boolean = False
    Public name As String = ""
    Public dataType As VariableDataTypes = VariableDataTypes.vdtString
    Public allowedValueList As String = "" 'System.Xml.XmlNodeList = Nothing
    Public allowedValueRange As String = "" 'System.Xml.XmlNodeList = Nothing
    Public defaultValue As String = ""
    Public value As String = ""
    Public hasValue As Boolean = False
    Public Sub Dispose()
        allowedValueList = ""
        allowedValueRange = ""
        value = ""
        name = ""
    End Sub
End Class

Public Enum VariableDataTypes
    vdtString = 0   ' can have allowedValueList
    vdtUI4 = 1
    vdtBoolean = 2
    vdtUI2 = 3      ' can have allowedValueRange
End Enum

Public Class MyactionList
    'Private MyactionList As New LinkedList(Of MyUPnPAction)

    Inherits List(Of MyUPnPAction)

    Public Overloads Sub Add(NewAction As MyUPnPAction)
        Try
            MyBase.Add(NewAction)
            If upnpDebuglevel > DebugLevel.dlEvents Then Log("MyactionList.Add added a new Action = " & NewAction.name, LogType.LOG_TYPE_INFO, LogColorGreen)
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in MyactionList.Add when called with Action = " & NewAction.name & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Overloads ReadOnly Property Item(bstrActioName As String) As MyUPnPAction
        Get
            Item = Nothing
            If Me.Count > 0 Then
                For Each Action As MyUPnPAction In Me
                    If Action IsNot Nothing Then
                        If Action.name = bstrActioName Then
                            Item = Action
                            Exit Property
                        End If
                    End If
                Next
            End If
        End Get
    End Property

    Public Sub Dispose()
        'If MyactionList Is Nothing Then Exit Sub
        Try
            For Each action As MyUPnPAction In Me
                Try
                    action.Dispose()
                    action = Nothing
                Catch ex As Exception
                End Try
            Next
            Clear()
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in MyactionList.dispose with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

End Class

Public Class MyUPnPAction

    Public name As String = ""
    Public argumentList As String = Nothing

    Public Sub Dispose()
        argumentList = Nothing
    End Sub

End Class


#End Region

Module UPnPDebug

    Public Class UPnPDebugWindow


        Inherits clsPageBuilder

        Private MyPageName As String = ""
        Private stb As New StringBuilder
        Private SSDPReference As MySSDP = Nothing


        Public Sub New(ByVal pagename As String)
            MyBase.New(pagename)
            MyPageName = pagename
        End Sub

        Public WriteOnly Property RefToSSDPn As MySSDP
            Set(value As MySSDP)
                SSDPReference = value
            End Set
        End Property

        ' test purpose
        Private Function GetPageHeader(ByVal pageName As String, title As String, extraMeta As String, HSOnLoad As String, excludeNavLinks As Boolean, noHeader As Boolean, Optional headContentOnly As Boolean = False, Optional bodyContentOnly As Boolean = False) As String
#If HS3 = "True" Then
            hs.GetPageHeader(pageName, title, extraMeta, HSOnLoad, excludeNavLinks, noHeader, headContentOnly, bodyContentOnly)
#Else
            Return ""
#End If
        End Function

        Private Function GetPageFooter(Optional noEndTags As Boolean = False) As String
#If HS3 = "True" Then
            hs.GetPageFooter(noEndTags)
#Else
            Return ""
#End If
        End Function

        ' build and return the actual page
        Public Function GetPagePlugin(ByVal pageName As String, ByVal user As String, ByVal userRights As Integer, ByVal queryString As String) As String

            If upnpDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetPagePlugin for UPnPDebugWindow called with pageName = " & pageName.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString & " and queryString = " & queryString, LogType.LOG_TYPE_INFO)

            Dim stb As New StringBuilder
            Dim stbUPnPDeviceTable As New StringBuilder

            Try
                Me.reset()

                ' handle any queries like mode=something               
                If (queryString <> "") Then
                    Return GetPageDevice(pageName, queryString)
                End If

                Me.AddHeader(GetPageHeader(pageName, "UPnP Viewer", "", "", False, True)) 'Me.AddHeader(hs.GetPageHeader(pageName, "UPnP Viewer", "", "", False, True))
                stb.Append(clsPageBuilder.DivStart("pluginpage", ""))

                ' a message area for error messages from jquery ajax postback (optional, only needed if using AJAX calls to get data)
                stb.Append(clsPageBuilder.DivStart("errormessage", "class='errormessage'"))
                stb.Append(clsPageBuilder.DivEnd)

                ' specific page starts here

                stb.Append(clsPageBuilder.DivStart("HeaderPanel", "style=""color:#0000FF"" "))
                stb.Append("<h1>UPnP Viewer</h1>" & vbCrLf)
                stb.Append(clsPageBuilder.DivEnd)

                ' create the UPnP Device table

                Dim AllDevices As MyUPnPDevices = SSDPReference.GetAllDevices()

                stb.Append("<table ID='UPnPDeviceListTable' border='1'  style='background-color:DarkGray;color:black'>")
                stb.Append("<tr ID='HeaderRow'  style='background-color:DarkGray;color:black'>")
                stb.Append("<td><h3> Image </h3></td><td><h3> Device Friendly Name </h3></td><td><h3> Alive </h3></td><td><h3> IP Address </h3></td><td><h3> IP Port </h3></td><td><h3> Location </h3></td><td><h3> DeviceType </h3></td><td><h3> UDN </h3></td></tr>")

                Dim DeviceListTableRow As Integer = 0

                Try
                    If Not AllDevices Is Nothing And AllDevices.Count > 0 Then
                        For Each DLNADevice As MyUPnPDevice In AllDevices
                            If DLNADevice.UniqueDeviceName <> "" Then
                                stb.Append("<tr ID='EntryRow'  style='background-color:LightGray;color:black'>")
                                stb.Append("<td>" & "<img src='" & DLNADevice.IconURL("image/jpeg", 200, 200, 16) & "' height='40px' width='40px' >" & "</td>") ' added ' quotes around img src in v023
                                stb.Append("<td><a href='" & UPNPVIEWPAGE & "?UDN=" & DLNADevice.UniqueDeviceName & "'>" & DLNADevice.FriendlyName & "</a></td>")
                                stb.Append("<td>" & DLNADevice.Alive & "</td>")
                                stb.Append("<td>" & DLNADevice.IPAddress & "</td>")
                                stb.Append("<td>" & DLNADevice.IPPort & "</td>")
                                stb.Append("<td><a href='" & DLNADevice.Location & "'>" & DLNADevice.Location & "</a>")
                                stb.Append("<td>" & DLNADevice.Type & "</td>")
                                stb.Append("<td>" & DLNADevice.UniqueDeviceName & "</td>")
                                'stb.Append("<td>")
                                'stb.Append("</td></tr>")
                                stb.Append("</tr>")
                                DeviceListTableRow += 1
                            End If
                        Next
                    End If
                Catch ex As Exception
                    Log("Error in Page load building the player list with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try

                'stb.Append("<hr /> ")
                stb.Append("</table>")
                stb.Append("<hr /> ")

                stb.Append(clsPageBuilder.FormEnd)

            Catch ex As Exception
                Log("Error in Page load with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try

            ' add the body html to the page
            Me.AddBody(stb.ToString)

            Me.AddFooter(GetPageFooter) 'hs.GetPageFooter)
            Me.suppressDefaultFooter = True

            ' return the full page
            Return Me.BuildPage()

        End Function

        Public Function GetPageDevice(ByVal pageName As String, ByVal queryString As String) As String
            If upnpDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetPageDevice for UPnPDebugWindow called with pageName = " & pageName.ToString & " and queryString = " & queryString, LogType.LOG_TYPE_INFO)

            Dim stb As New StringBuilder
            Dim stbUPnPDeviceTable As New StringBuilder

            Try
                Me.reset()
                ' queryString holds the UDN and ServiceID
                Dim parts As Collections.Specialized.NameValueCollection = Nothing
                If (queryString <> "") Then
                    parts = HttpUtility.ParseQueryString(queryString)
                    If parts IsNot Nothing Then
                        If parts.Count > 1 Then
                            Return GetPageService(pageName, parts.Item("UDN"), parts.Item("ServiceID"))
                        End If
                    End If
                End If

                Dim UDN As String = parts.Item("UDN")
                Dim UPnPDevice As MyUPnPDevice = SSDPReference.Item(UDN, True)

                Me.AddHeader(GetPageHeader(pageName, "UPnP Device Viewer", "", "", False, True)) 'Me.AddHeader(hs.GetPageHeader(pageName, "UPnP Device Viewer", "", "", False, True))
                stb.Append(clsPageBuilder.DivStart("pluginpage", ""))

                ' a message area for error messages from jquery ajax postback (optional, only needed if using AJAX calls to get data)
                stb.Append(clsPageBuilder.DivStart("errormessage", "class='errormessage'"))
                stb.Append(clsPageBuilder.DivEnd)

                ' specific page starts here

                stb.Append(clsPageBuilder.DivStart("HeaderPanel", "style=""color:#0000FF"" "))
                stb.Append("<h1>UPnP Device Viewer " & UPnPDevice.UniqueDeviceName & "</h1>" & vbCrLf)
                stb.Append(clsPageBuilder.DivEnd)

                stb.Append("Alive = " & UPnPDevice.Alive & "</br>")
                stb.Append("Friendly Name = " & UPnPDevice.FriendlyName & "</br>")
                stb.Append("Device UDN = " & UPnPDevice.UniqueDeviceName & "</br>")
                stb.Append("IP Address = " & UPnPDevice.IPAddress & "</br>")
                stb.Append("IP Port = " & UPnPDevice.IPPort & "</br>")
                stb.Append("Location = " & "<a href='" & UPnPDevice.Location & "'>" & UPnPDevice.Location & "</a>" & "</br>")
                stb.Append("Device Type = " & UPnPDevice.Type & "</br>")
                stb.Append("Application URL = " & UPnPDevice.ApplicationURL & "</br>")
                stb.Append("WakeUp = " & UPnPDevice.WakeUp & "</br>")
                stb.Append("SSID = " & UPnPDevice.SSID & "</br>")
                stb.Append("BootId = " & UPnPDevice.BootID & "</br>")
                stb.Append("Cache-Control = " & UPnPDevice.CacheControl & "</br>")
                stb.Append("ConfigId = " & UPnPDevice.ConfigID & "</br>")
                stb.Append("SearchPort = " & UPnPDevice.SearchPort & "</br>")
                stb.Append("Header = " & UPnPDevice.NotificationEvent & "</br>")
                stb.Append("Has Children = " & UPnPDevice.HasChildren.ToString & "</br>")
                If UPnPDevice.HasChildren Then
                    Dim Children As MyUPnPDevices = UPnPDevice.Children
                    If Children IsNot Nothing Then
                        If Children.Count > 0 Then
                            For Each Child As MyUPnPDevice In Children
                                stb.Append("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='" & UPNPVIEWPAGE & "?UDN=" & Child.UniqueDeviceName & "'>" & Child.FriendlyName & "</a></br>")
                            Next
                        End If
                    End If
                End If
                stb.Append("ManufacturerName = " & UPnPDevice.ManufacturerName & "</br>")
                stb.Append("Model Number = " & UPnPDevice.ModelNumber & "</br>")
                stb.Append("</br>")
                If UPnPDevice.Services IsNot Nothing Then
                    If UPnPDevice.Services.Count > 0 Then
                        stb.Append("Has Services = " & UPnPDevice.Services.Count.ToString & "</br>")
                        For Each Service As MyUPnPService In UPnPDevice.Services
                            stb.Append("Service ID = " & "<a href='" & UPNPVIEWPAGE & "?UDN=" & UDN & "&ServiceID=" & Service.Id & "'>" & Service.Id & "</a>" & "</br>")
                            stb.Append("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Service Type Identifier = " & Service.ServiceTypeIdentifier & "</br>")
                            stb.Append("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Service LastTransportStatus = " & Service.LastTransportStatus.ToString & "</br>")
                            stb.Append("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Service Type = " & Service.ServiceType & "</br>")
                            If Service.MyLocation <> "" Then stb.Append("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Location = " & "<a href='" & Service.MyLocation & "'>" & Service.MyLocation & "</a>" & "</br>")
                            If Service.NotificationEvent <> "" Then stb.Append("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;NotificationEvent = " & Service.NotificationEvent & "</br>")
                            stb.Append("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Service URL = " & "<a href='" & Service.MySCPDURL & "'>" & Service.MySCPDURL & "</a>" & "</br>")
                            stb.Append("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Service Control URL = " & Service.MycontrolURL & "</br>")
                            stb.Append("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Service Event URL = " & Service.MyeventSubURL & "</br>")
                            stb.Append("</br>")
                        Next
                    End If
                End If

                stb.Append("<hr /> ")

                stb.Append(clsPageBuilder.FormEnd)

            Catch ex As Exception
                Log("Error in Page load with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try

            ' add the body html to the page
            Me.AddBody(stb.ToString)

            Me.AddFooter(GetPageFooter) 'hs.GetPageFooter)
            Me.suppressDefaultFooter = True

            ' return the full page
            Return Me.BuildPage()

        End Function

        Public Function GetPageService(ByVal pageName As String, ByVal UDN As String, ByVal ServiceID As String) As String
            If upnpDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetPageService for UPnPDebugWindow called with pageName = " & pageName.ToString & " and UDN = " & UDN & " and ServiceID = " & ServiceID, LogType.LOG_TYPE_INFO)

            Dim stb As New StringBuilder
            Dim stbUPnPDeviceTable As New StringBuilder

            Try
                Me.reset()

                Dim UPnPDevice As MyUPnPDevice = SSDPReference.Item(UDN, True)
                If UPnPDevice Is Nothing Then
                    Return "Device not Found Error"
                End If
                Dim UPnPService As MyUPnPService = UPnPDevice.Services.Item(ServiceID)
                If UPnPService Is Nothing Then
                    Return "Service not Found Error"
                End If

                Me.AddHeader(GetPageHeader(pageName, "UPnP Service Viewer", "", "", False, True)) 'Me.AddHeader(hs.GetPageHeader(pageName, "UPnP Service Viewer", "", "", False, True))
                stb.Append(clsPageBuilder.DivStart("pluginpage", ""))

                ' a message area for error messages from jquery ajax postback (optional, only needed if using AJAX calls to get data)
                stb.Append(clsPageBuilder.DivStart("errormessage", "class='errormessage'"))
                stb.Append(clsPageBuilder.DivEnd)

                ' specific page starts here

                stb.Append(clsPageBuilder.DivStart("HeaderPanel", "style=""color:#0000FF"" "))
                stb.Append("<h1>UPnP Service Viewer  " & UPnPService.Id & "</h1>" & vbCrLf)
                stb.Append(clsPageBuilder.DivEnd)

                stb.Append("isSubscribed = " & UPnPService.isServiceEventsSubscribed.ToString & "</br>")
                stb.Append("Service ID = " & UPnPService.Id & "</br>")
                stb.Append("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Control URL = " & UPnPService.MycontrolURL & "</br>")
                stb.Append("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Event URL = " & UPnPService.MyeventSubURL & "</br>")
                stb.Append("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Service URL = " & "<a href='" & UPnPService.MySCPDURL & "'>" & UPnPService.MySCPDURL & "</a>" & "</br>")
                stb.Append("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Received SID = " & UPnPService.MyReceivedSID & "</br>")
                stb.Append("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;TimeOut Value SID = " & UPnPService.MyTimeout.ToString & "</br>")
                stb.Append("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Service Type = " & UPnPService.ServiceType & "</br>")
                stb.Append("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Service Type ID = " & UPnPService.ServiceTypeIdentifier & "</br>")
                stb.Append("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;hasServiceXMLRetrieved = " & UPnPService.hasServiceXMLRetrieved & "</br>")
                stb.Append("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;hasServiceStateRetrieved= " & UPnPService.hasServiceStateRetrieved & "</br>")
                stb.Append("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;hasActionListRetrieved= " & UPnPService.hasActionListRetrieved & "</br>")

                If UPnPService.ActionList IsNot Nothing Then
                    stb.Append("</br>")
                    stb.Append("Action List has = " & UPnPService.ActionList.Count.ToString & " Entries</br>")
                    For Each Action As MyUPnPAction In UPnPService.ActionList
                        stb.Append("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Action Name = <b>" & Action.name & "</b></br>")
                        If Action.argumentList <> "" Then
                            Dim ActionXML As New XmlDocument
                            ActionXML.LoadXml(Action.argumentList)
                            Dim ArgumentList As XmlNodeList = ActionXML.GetElementsByTagName("argument")
                            For Each XML_Node As XmlElement In ArgumentList
                                stb.Append("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Argument name = " & If(XML_Node.GetElementsByTagName("name").Item(0)?.InnerText, "") & "</br>")
                                stb.Append("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Argument direction = " & If(XML_Node.GetElementsByTagName("direction").Item(0)?.InnerText, "") & "</br>")
                                Dim RelatedStateVariable As String = If(XML_Node.GetElementsByTagName("relatedStateVariable").Item(0)?.InnerText, "")
                                ' <a href="#top">link to top</a>
                                stb.Append("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Argument relatedStateVariable = <a href='#" & RelatedStateVariable & "'>" & RelatedStateVariable & "</a></br>")
                            Next
                        End If
                    Next
                End If
                If UPnPService.ServiceStateTable IsNot Nothing Then
                    stb.Append("</br>")
                    stb.Append("Service State Table has = " & UPnPService.ServiceStateTable.Count.ToString & " Entries</br>")
                    For Each ServiceState In UPnPService.ServiceStateTable
                        stb.Append("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Service State Name = <b><a name='" & ServiceState.name & "'>" & ServiceState.name & "</a></b></br>")
                        stb.Append("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Service State Send Events = " & ServiceState.sendEvents & "</br>")
                        If ServiceState.sendEvents Then
                            stb.Append("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Service State hasValue = " & ServiceState.hasValue.ToString & "</br>")
                            stb.Append("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Service State Value = <textarea>" & ServiceState.value & "</textarea></br>")
                        End If
                        stb.Append("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Service State Data Type = " & ServiceState.dataType.ToString & "</br>")
                        If ServiceState.allowedValueList <> "" Then
                            Dim AllowedValueListString As String = ""
                            Dim AllowedValueListXMLDoc As New XmlDocument
                            AllowedValueListXMLDoc.LoadXml(ServiceState.allowedValueList)
                            For Each XML_Node As XmlNode In AllowedValueListXMLDoc.ChildNodes
                                If XML_Node.ChildNodes IsNot Nothing Then
                                    For Each XML_Child_node As XmlNode In XML_Node.ChildNodes
                                        If AllowedValueListString <> "" Then AllowedValueListString &= ","
                                        AllowedValueListString &= XML_Child_node.InnerText
                                    Next
                                End If
                            Next
                            If AllowedValueListString <> "" Then stb.Append("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Service State AllowedValueList = " & AllowedValueListString & "</br>")
                        End If
                        If ServiceState.allowedValueRange <> "" Then
                            Dim AllowedValueRangeString As String = ""
                            Dim AllowedValueRangeXMLDoc As New XmlDocument
                            AllowedValueRangeXMLDoc.LoadXml(ServiceState.allowedValueRange)
                            For Each XML_Node As XmlNode In AllowedValueRangeXMLDoc.ChildNodes
                                If XML_Node.ChildNodes IsNot Nothing Then
                                    For Each XML_Child_node As XmlNode In XML_Node.ChildNodes
                                        If AllowedValueRangeString <> "" Then AllowedValueRangeString &= ","
                                        AllowedValueRangeString &= XML_Child_node.LocalName & "=" & XML_Child_node.InnerText
                                    Next
                                End If
                            Next
                            If AllowedValueRangeString <> "" Then stb.Append("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Service State AllowedValueRange = " & AllowedValueRangeString & "</br>")
                        End If
                    Next
                End If
                stb.Append("<hr /> ")

                stb.Append(clsPageBuilder.FormEnd)

            Catch ex As Exception
                Log("Error in Page load with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try

            ' add the body html to the page
            Me.AddBody(stb.ToString)

            Me.AddFooter(GetPageFooter) 'hs.GetPageFooter)
            Me.suppressDefaultFooter = True

            ' return the full page
            Return Me.BuildPage()

        End Function


        Public Overrides Function postBackProc(page As String, data As String, user As String, userRights As Integer) As String
            If upnpDebuglevel > DebugLevel.dlErrorsOnly Then Log("PostBackProc for UPnPDebugWindow called with page = " & page.ToString & " and data = " & data.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString, LogType.LOG_TYPE_INFO)

            Dim parts As Collections.Specialized.NameValueCollection
            parts = HttpUtility.ParseQueryString(data)

            If parts IsNot Nothing Then
                Try
                    Dim Part As String
                    For Each Part In parts.AllKeys
                        If Part IsNot Nothing Then
                            Dim ObjectNameParts As String()
                            ObjectNameParts = Split(Part, "_")
                            If upnpDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc for UPnPDebugWindow found Key = " & ObjectNameParts(0).ToString, LogType.LOG_TYPE_INFO)
                            If upnpDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc for UPnPDebugWindow found Value = " & parts(Part).ToString, LogType.LOG_TYPE_INFO)
                            Dim ObjectValue As String = parts(Part)
                            Select Case ObjectNameParts(0).ToString
                                Case "DebugChkBox"
                                    Try
                                        WriteBooleanIniFile("Options", "Debug", ObjectValue.ToUpper = "CHECKED")
                                    Catch ex As Exception
                                        Log("Error in postBackProc for PluginControl saving Debug flag. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                    End Try
                            End Select
                        End If
                    Next
                Catch ex As Exception
                    Log("Error in postBackProc for UPnPDebugWindow processing with page = " & page.ToString & " and data = " & data.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString & " and Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
            Else
                If upnpDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc for UPnPDebugWindow found parts to be empty", LogType.LOG_TYPE_INFO)
            End If

            Return MyBase.postBackProc(page, data, user, userRights)
        End Function

        Public Enum DeviceTableItems
            ptiDeviceGivenName = 0
            ptiAddBtn = 1
            ptiRemoveBtn = 2
            ptiConfigBtn = 3
            ptiDeleteBtn = 4
        End Enum

        Private Function GetHeadContent() As String
            Try
                Return GetPageHeader(PageName, "", "", False, False, True) 'hs.GetPageHeader(sIFACE_NAME, "", "", False, False, True, False, False)
            Catch ex As Exception
            End Try
            Return ""
        End Function

        Private Function GetFooterContent() As String
            Try
                Return GetPageFooter(False) 'hs.GetPageFooter(False)
            Catch ex As Exception
            End Try
            Return ""
        End Function

        Private Function GetBodyContent() As String
            Try
                Return GetPageHeader(PageName, "", "", False, False, True, True, False) 'hs.GetPageHeader(StrConv(sIFACE_NAME, VbStrConv.ProperCase), "", "", False, False, False, True, False)
            Catch ex As Exception
            End Try
            Return ""
        End Function


    End Class

End Module

Module UPNPUtils

    Public Function ExtractIPInfo(DocumentURL As String) As IPAddressInfo
        Dim HttpIndex As Integer = DocumentURL.ToUpper.IndexOf("HTTP://")
        ExtractIPInfo.IPAddress = ""
        ExtractIPInfo.IPPort = ""
        Dim LoopBackAddress As IPAddress = Nothing
        IPAddress.TryParse(LoopBackIPv4Address, LoopBackAddress)
        Try
            If HttpIndex = -1 Then
                If upnpDebuglevel > DebugLevel.dlErrorsOnly Then Log("ERROR in MySSDP.ExtractIPInfo. Not HTTP:// found. URL = " & DocumentURL, LogType.LOG_TYPE_ERROR)
                Exit Function
            End If
            Dim SubStr As String
            SubStr = DocumentURL.Substring(HttpIndex + 7, DocumentURL.Length - HttpIndex - 7)
            ' substring should now be primed for an IP address in the form of 192.168.1.1
            ' The next forward slash marks the end of the IP address, this could include the Port!
            Dim SlashIndex As Integer = SubStr.IndexOf("/")
            If SlashIndex <> -1 Then
                SubStr = SubStr.Substring(0, SlashIndex)
            End If
            SubStr = SubStr.Trim
            If SubStr = "" Then
                If upnpDebuglevel > DebugLevel.dlErrorsOnly Then Log("ERROR in MySSDP.ExtractIPInfo. No IP address found = " & DocumentURL, LogType.LOG_TYPE_ERROR)
                Exit Function
            End If
            Dim SemiCollonIndex As Integer = SubStr.IndexOf(":")
            'log( "ExtractIPInfo has Substring = " & SubStr)
            If SemiCollonIndex <> -1 Then
                ' there is an IP address and a Port Number
                ExtractIPInfo.IPAddress = SubStr.Substring(0, SemiCollonIndex)
                ExtractIPInfo.IPPort = SubStr.Substring(SemiCollonIndex + 1, SubStr.Length - SemiCollonIndex - 1)
            Else
                ' only IP address
                ExtractIPInfo.IPAddress = SubStr
            End If
            Dim ExtractedIPAddress As IPAddress = Nothing
            IPAddress.TryParse(ExtractIPInfo.IPAddress, ExtractedIPAddress)
            Try
                If ExtractedIPAddress.Equals(LoopBackAddress) Then ExtractIPInfo.IPAddress = PlugInIPAddress
            Catch ex As Exception
            End Try
        Catch ex As Exception
            If upnpDebuglevel > DebugLevel.dlOff Then Log("ERROR in MySSDP.ExtractIPInfo URL = " & DocumentURL & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

End Module


