Imports System.Runtime.InteropServices
Imports System.Text

Partial Public Class HSPI

    Private Const MAXLEN_PHYSADDR As Integer = 6

    Private Declare Function GetIpNetTable Lib "Iphlpapi" ( _
      ByVal pIpNetTable As IntPtr, _
      ByRef pdwSize As Integer, ByVal bOrder As Boolean) As Integer

    Private Structure MIB_IPNETROW
        Dim dwIndex As Integer
        Dim dwPhysAddrLen As Integer
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=MAXLEN_PHYSADDR)> Dim dwPhysAddr As Byte()
        Dim dwAddr As Integer
        Dim dwStructure As Integer
    End Structure 'MIB_IPNETROW'

    Private Function ConvertMacAddress(ByVal byteArray() As Byte) As String
        Dim builder As New StringBuilder
        For Each byteCurrent As Byte In byteArray
            builder.Append(byteCurrent.ToString("X2")).Append("-"c)
        Next byteCurrent
        builder.Length -= 1
        Return builder.ToString
    End Function 'ConvertMacAddress'

    Private Function ConvertMacAddress(ByVal inString As String) As String
        ConvertMacAddress = ""
        inString = Trim(inString)
        If inString = "" Then Exit Function
        Dim Index As Integer = 0
        Dim OutString As String = ""
        While Index < inString.Length - 1
            OutString = OutString & inString(Index) & inString(Index + 1)
            Index = Index + 2
            If Index < inString.Length Then OutString = OutString & "-"
        End While
        ConvertMacAddress = OutString
    End Function 'ConvertMacAddress'


 
    Public Function GetMACAddress_(IPAddress As String) As String
        GetMACAddress_ = ""
        If IPAddress = "" Then Exit Function
        If SuperDebug Then Log("GetMACAddress called with IPAddress = " & IPAddress, LogType.LOG_TYPE_INFO)
        If IPAddress = hs.GetIPAddress() Or IPAddress = "127.0.0.1" Then
            ' local address
            GetMACAddress_ = ConvertMacAddress(GetLocalMacAddress())
            If SuperDebug Then Log("GetMACAddress found local IPAddress = " & IPAddress.ToString & " and MACAddress = " & GetMACAddress_.ToString, LogType.LOG_TYPE_INFO)
            Exit Function
        End If

        Const ERROR_INSUFFICIENT_BUFFER As Integer = &H7A

        ' The number of bytes needed.
        Dim bytesNeeded As Integer = 0
        ' The result from the API call.
        Dim result As Integer
        Try
            result = GetIpNetTable(IntPtr.Zero, bytesNeeded, False)
            ' Call the function, expecting an insufficient buffer.
            If result <> ERROR_INSUFFICIENT_BUFFER Then
                Log("Error in GetMACAddress getting IPNetTable with error = " & result.ToString, LogType.LOG_TYPE_ERROR)
                Exit Function
            End If
        Catch ex As Exception
            Log("Error in GetMACAddress with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try

        ' Allocate the memory, do it in a try/finally block, to ensure that it is released.
        Dim Buffer As IntPtr = IntPtr.Zero
        Try
            ' Allocate the memory.
            Buffer = Marshal.AllocCoTaskMem(bytesNeeded)
            ' Make the call again. If it did not succeed, then raise an error.
            result = GetIpNetTable(Buffer, bytesNeeded, False)
            ' If the result is not 0 (no error), then throw an exception.
            If (result <> 0) Then
                ' Throw an exception.
                Log("Error in GetMACAddress with error = " & result.ToString, LogType.LOG_TYPE_ERROR)
                Exit Function
            Else
                ' Now we have the buffer, we have to marshal it. We can read() the first 4 bytes to get the length of the buffer.
                Dim entries As Integer = Marshal.ReadInt32(Buffer)
                ' Increment the memory pointer by the size of the int.
                Dim currentBuffer As New IntPtr(Buffer.ToInt64() + Marshal.SizeOf(GetType(Integer)))
                ' Allocate an array of entries.
                Dim table(entries) As MIB_IPNETROW
                ' Cycle through the entries.
                For index As Integer = 0 To entries - 1
                    table(index) = CType(Marshal.PtrToStructure(currentBuffer, GetType(MIB_IPNETROW)), MIB_IPNETROW)
                    Dim _ipAddress As New System.Net.IPAddress(BitConverter.GetBytes(table(index).dwAddr))
                    If IPAddress = _ipAddress.ToString Then
                        Dim macAddress As String = ConvertMacAddress(table(index).dwPhysAddr)
                        If SuperDebug Then Log("GetMACAddress found entry with IPAddress = " & IPAddress.ToString & " and MACAddress = " & macAddress.ToString, LogType.LOG_TYPE_INFO)
                        GetMACAddress_ = macAddress.ToString
                        ' Release the memory.
                        Marshal.FreeCoTaskMem(Buffer)
                        currentBuffer = Nothing
                        'macAddress = Nothing
                        _ipAddress = Nothing
                        table = Nothing
                        Exit Function
                    End If
                    currentBuffer = New IntPtr(currentBuffer.ToInt64 + Marshal.SizeOf(GetType(MIB_IPNETROW)))
                Next
            End If
        Catch ex As Exception
            Log("Error in GetMACAddress 1 with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            ' Release the memory.
            Marshal.FreeCoTaskMem(Buffer)
        Catch ex As Exception
        End Try
        If g_bDebug Then Log("Warning in GetMACAddress. IPAddress = " & IPAddress & " wasn't found", LogType.LOG_TYPE_ERROR)
    End Function

    Public Function GetLocalMacAddress() As String
        If SuperDebug Then Log("GetLocalMacAddress called", LogType.LOG_TYPE_INFO)
        GetLocalMacAddress = ""
        Dim LocalMacAddress As String = ""
        Dim LocalIPAddress = hs.getIpAddress()
        If LocalIPAddress = "" Then
            Log("Error in GetLocalMacAddress trying to get own IP address", LogType.LOG_TYPE_ERROR)
            Exit Function
        End If
        For Each nic As System.Net.NetworkInformation.NetworkInterface In System.Net.NetworkInformation.NetworkInterface.GetAllNetworkInterfaces()
            If SuperDebug Then Log(String.Format("The MAC address of {0} is {1}{2}", nic.Description, Environment.NewLine, nic.GetPhysicalAddress()), LogType.LOG_TYPE_INFO)
            For Each Ipa In nic.GetIPProperties.UnicastAddresses
                If SuperDebug Then Log(String.Format("The IPaddress address of {0} is {1}{2}", nic.Description, Environment.NewLine, Ipa.Address.ToString), LogType.LOG_TYPE_INFO)
                'If g_bDebug Then log( String.Format("The IPaddress address of {0} is {1}{2}", nic.Description, Environment.NewLine, Ipa.Address.ToString))
                If Ipa.Address.ToString = LocalIPAddress Then
                    ' OK we found our IPaddress
                    LocalMacAddress = nic.GetPhysicalAddress().ToString
                    If SuperDebug Then Log("GetLocalMacAddress found local MAC address = " & LocalMacAddress, LogType.LOG_TYPE_INFO)
                    GetLocalMacAddress = LocalMacAddress
                    Exit Function
                End If
            Next
        Next
        If g_bDebug Then Log("Error in GetLocalMacAddress trying to get own MAC address, none found", LogType.LOG_TYPE_ERROR)
    End Function

    Public Function GetIPMask() As String
        If SuperDebug Then Log("GetIPMask called", LogType.LOG_TYPE_INFO)
        GetIPMask = ""
        Dim LocalIPAddress = hs.getIpAddress()
        If LocalIPAddress = "" Then
            Log("Error in GetIPMask trying to get own IP address", LogType.LOG_TYPE_ERROR)
            Exit Function
        End If
        For Each nic As System.Net.NetworkInformation.NetworkInterface In System.Net.NetworkInformation.NetworkInterface.GetAllNetworkInterfaces()
            If SuperDebug Then Log(String.Format("The MAC address of {0} is {1}{2}", nic.Description, Environment.NewLine, nic.GetPhysicalAddress()), LogType.LOG_TYPE_INFO)
            For Each Ipa In nic.GetIPProperties.UnicastAddresses
                If SuperDebug Then Log(String.Format("The IPaddress address of {0} is {1}{2}", nic.Description, Environment.NewLine, Ipa.Address.ToString), LogType.LOG_TYPE_INFO)
                'If g_bDebug Then log( String.Format("The IPaddress address of {0} is {1}{2}", nic.Description, Environment.NewLine, Ipa.Address.ToString))
                If Ipa.Address.ToString = LocalIPAddress Then
                    ' OK we found our IPaddress
                    GetIPMask = Ipa.IPv4Mask.ToString
                    If SuperDebug Then Log("GetIPMask found IP Mask = " & Ipa.IPv4Mask.ToString, LogType.LOG_TYPE_INFO)
                    Exit Function
                End If
            Next
        Next
        If g_bDebug Then Log("Error in GetIPMask, none found", LogType.LOG_TYPE_ERROR)
    End Function



End Class

