Imports System.IO
Imports System.Runtime.Serialization.Formatters
Imports System.Web.Script.Serialization
Imports System.Net
Imports System.Security.Cryptography
Imports System.Drawing
Imports System.Xml.Serialization
Imports System.Net.Sockets
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Xml
Imports System.Net.NetworkInformation
Imports System.Net.Http
Imports System.IO.Compression









#If HS3 = "True" Then
Imports HomeSeerAPI
#Else
Imports HomeSeer.PluginSdk
Imports HomeSeer.PluginSdk.Devices
#End If

Module util

    Public Const LogColorWhite = "#FFFFFF"
    Public Const LogColorRed = "#FF0000"
    Public Const LogColorBlack = "#000000"
    Public Const LogColorNavy = "#000080"
    Public Const LogColorLightBlue = "#D9F2FF"
    Public Const LogColorLightGray = "#E1E1E1"
    Public Const LogColorPink = "#FFB6C1"
    Public Const LogColorOrange = "#D58000"
    Public Const LogColorGreen = "#008000"

    'Private LogFileStreamWriter As StreamWriter = Nothing
    Private logFileTextWriter As TextWriter = Nothing
    Private MyLogFileName As String = ""

    Public Structure pair
        Dim name As String
        Dim value As String
    End Structure

    Enum LogLevel As Integer
        Normal = 1
        Debug = 2
    End Enum

    Enum MessageType
        Normal = 0
        Warning = 1
        Error_ = 2
    End Enum

    Friend Enum eTriggerType
        Unknown = 0 ' not really used
    End Enum
    Friend Enum eActionType
        Unknown = 0 ' not really used
    End Enum
    Friend Structure strTrigger
        Public WhichTrigger As eTriggerType
        Public TrigObj As Object
        Public Result As Boolean
    End Structure
    Friend Structure strAction
        Public WhichAction As eActionType
        Public ActObj As Object
        Public Result As Boolean
    End Structure



    'Sub PEDAdd(ByRef PED As clsPlugExtraData, ByVal PEDName As String, ByVal PEDValue As Object)
    'Dim ByteObject() As Byte = Nothing
    'If PED Is Nothing Then PED = New clsPlugExtraData
    '   SerializeObject(PEDValue, ByteObject)
    'If Not PED.AddNamed(PEDName, ByteObject) Then
    '       PED.RemoveNamed(PEDName)
    '      PED.AddNamed(PEDName, ByteObject)
    'End If
    'End Sub

    'Function PEDGet(ByRef PED As clsPlugExtraData, ByVal PEDName As String) As Object
    'Dim ByteObject() As Byte
    'Dim ReturnValue As New Object
    '   ByteObject = PED.GetNamed(PEDName)
    'If ByteObject Is Nothing Then Return Nothing
    '   DeSerializeObject(ByteObject, ReturnValue)
    'Return ReturnValue
    'End Function

    Public Function SerializeObject(ByRef ObjIn As Object, ByRef bteOut() As Byte) As Boolean
        If ObjIn Is Nothing Then Return False
        Dim str As New MemoryStream
        Dim sf As New Binary.BinaryFormatter

        Try
            sf.Serialize(str, ObjIn)
            ReDim bteOut(CInt(str.Length - 1))
            bteOut = str.ToArray
            Return True
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then Log(" Error: Serializing object " & ObjIn.ToString & " :" & ex.Message, LogType.LOG_TYPE_ERROR)
            Return False
        End Try

    End Function

    Public Function DeSerializeObject(ByRef bteIn() As Byte, ByRef ObjOut As Object) As Boolean
        ' Almost immediately there is a test to see if ObjOut is NOTHING.  The reason for this
        '   when the ObjOut is suppose to be where the deserialized object is stored, is that 
        '   I could find no way to test to see if the deserialized object and the variable to 
        '   hold it was of the same type.  If you try to get the type of a null object, you get
        '   only a null reference exception!  If I do not test the object type beforehand and 
        '   there is a difference, then the InvalidCastException is thrown back in the CALLING
        '   procedure, not here, because the cast is made when the ByRef object is cast when this
        '   procedure returns, not earlier.  In order to prevent a cast exception in the calling
        '   procedure that may or may not be handled, I made it so that you have to at least 
        '   provide an initialized ObjOut when you call this - ObjOut is set to nothing after it 
        '   is typed.
        If bteIn Is Nothing Then Return False
        If bteIn.Length < 1 Then Return False
        If ObjOut Is Nothing Then Return False
        Dim str As MemoryStream
        Dim sf As New Binary.BinaryFormatter
        Dim ObjTest As Object
        Dim TType As System.Type
        Dim OType As System.Type
        Try
            OType = ObjOut.GetType
            ObjOut = Nothing
            str = New MemoryStream(bteIn)
            ObjTest = sf.Deserialize(str)
            If ObjTest Is Nothing Then Return False
            TType = ObjTest.GetType
            'If Not TType.Equals(OType) Then Return False
            ObjOut = ObjTest
            If ObjOut Is Nothing Then Return False
            Return True
        Catch exIC As InvalidCastException
            Return False
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then Log(" Error: DeSerializing object: " & ex.Message, LogType.LOG_TYPE_ERROR)
            Return False
        End Try

    End Function

    Public Sub wait(ByVal secs As Decimal)
        If piDebuglevel > DebugLevel.dlEvents Then Log("wait called with value = " & secs, LogType.LOG_TYPE_INFO)
        Threading.Thread.Sleep(secs * 1000)
    End Sub

    Enum LogType
        LOG_TYPE_INFO = 0
        LOG_TYPE_ERROR = 1
        LOG_TYPE_WARNING = 2
    End Enum

    ' Member-level variables
    Private logFileWriter As StreamWriter = Nothing
    Private logFlushTimer As System.Timers.Timer = Nothing
    Private logFileLock As New Object()
    Private lastLogWrite As DateTime = DateTime.MinValue

    Public Function OpenLogFile(LogFileName As String, Optional append As Boolean = False) As Boolean
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("OpenLogFile called with LogFileName = " & LogFileName, LogType.LOG_TYPE_INFO)
        If logFileWriter IsNot Nothing Then CloseLogFile()

        Try
            If LogFileName <> "" Then
                If Not append AndAlso File.Exists(LogFileName) Then
                    If File.Exists(LogFileName & ".bak") Then File.Delete(LogFileName & ".bak")
                    File.Move(LogFileName, LogFileName & ".bak")
                End If

                Dim fs As New FileStream(LogFileName, FileMode.Append, FileAccess.Write, FileShare.Read)
                Dim writer = New StreamWriter(fs)
                writer.AutoFlush = False
                logFileWriter = writer
                MyLogFileName = LogFileName

                ' Setup flush timer
                If logFlushTimer Is Nothing Then
                    logFlushTimer = New System.Timers.Timer(10000) ' every 10 seconds
                    AddHandler logFlushTimer.Elapsed, AddressOf FlushLogIfIdle
                    logFlushTimer.AutoReset = True
                    logFlushTimer.Start()
                End If

                Return True
            End If
        Catch ex As Exception
            logFileWriter = Nothing
            MyLogFileName = ""
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in OpenLogFile: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        Return False
    End Function

    Private Sub FlushLogIfIdle(sender As Object, e As Timers.ElapsedEventArgs)
        Try
            If logFileWriter IsNot Nothing Then
                SyncLock logFileLock
                    If (DateTime.Now - lastLogWrite).TotalSeconds >= 10 Then
                        logFileWriter.Flush()
                    End If
                End SyncLock
            End If
        Catch
            ' suppress flush exceptions
        End Try
    End Sub

    Public Sub CloseLogFile()
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("CloseLogFile called for DiskFileName = " & MyLogFileName, LogType.LOG_TYPE_INFO)

        If logFlushTimer IsNot Nothing Then
            logFlushTimer.Stop()
            logFlushTimer.Dispose()
            logFlushTimer = Nothing
        End If

        If logFileWriter IsNot Nothing Then
            Try
                logFileWriter.Flush()
                logFileWriter.Close()
            Catch ex As Exception
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error closing log: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Try
                logFileWriter.Dispose()
            Catch ex As Exception
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error disposing log: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            logFileWriter = Nothing
        End If

        MyLogFileName = ""
    End Sub

#If HS3 = "True" Then
        Public Sub Log(ByVal msg As String, ByVal logType As LogType, Optional ByVal MsgColor As String = "", Optional ErrorCode As Integer = 0)
        Try
            If msg Is Nothing Then msg = ""
            If Not [Enum].IsDefined(GetType(LogType), logType) Then
                logType = util.LogType.LOG_TYPE_ERROR
            End If
            If Not ImRunningOnLinux Then Console.WriteLine(DateAndTime.Now.ToString & " : " & msg)
            Select Case logType
                Case LogType.LOG_TYPE_ERROR
                    If MsgColor <> "" Then
                        hs.WriteLogDetail(ShortIfaceName & " Error", msg, MsgColor, "1", "UPnP", ErrorCode)
                    Else
                        hs.WriteLog(ShortIfaceName & " Error", msg)
                    End If
                Case LogType.LOG_TYPE_WARNING
                    If MsgColor <> "" Then
                        hs.WriteLogDetail(ShortIfaceName & " Warning", msg, MsgColor, "0", "UPnP", ErrorCode)
                    Else
                        hs.WriteLog(ShortIfaceName & " Warning", msg)
                    End If
                Case LogType.LOG_TYPE_INFO
                    If MsgColor <> "" Then
                        hs.WriteLogDetail(ShortIfaceName, msg, MsgColor, "0", "UPnP", ErrorCode)
                    Else
                        hs.WriteLog(ShortIfaceName, msg)
                    End If
            End Select
        Catch ex As Exception
            If Not ImRunningOnLinux Then Console.WriteLine("Exception in LOG of " & IFACE_NAME & ": " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If MyLogFileName <> "" And gLogToDisk Then
                If logFileTextWriter IsNot Nothing Then
                    'LogFileStreamWriter.WriteLine(DateAndTime.Now.ToString & " : " & msg)
                    logFileTextWriter.WriteLine(DateAndTime.Now.ToString & " : " & logType.ToString & " - " & msg)
                End If
            End If
        Catch ex As Exception
            logFileTextWriter = Nothing
            MyLogFileName = ""
            If Not ImRunningOnLinux Then Console.WriteLine(DateAndTime.Now.ToString & " : " & " Exception in LOG with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            hs.WriteLog(ShortIfaceName & " Error", " Exception in LOG with Error = " & ex.Message)
        End Try
    End Sub
    Public Function GetStringIniFile(ByVal Section As String, ByVal Key As String, ByVal DefaultVal As String, Optional FileName As String = "") As String
        GetStringIniFile = ""
        If FileName = "" Then FileName = MyINIFile
        Try
            GetStringIniFile = hs.GetINISetting(Section, EncodeINIKey(Key), DefaultVal, FileName)
            'Log("GetStringIniFile called with section = " & Section & " and Key = " & Key & " read = " & GetStringIniFile, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            Log("Error in GetStringIniFile with section = " & Section & " and Key = " & EncodeINIKey(Key) & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function GetIntegerIniFile(ByVal Section As String, ByVal Key As String, ByVal DefaultVal As Integer, Optional FileName As String = "") As Integer
        GetIntegerIniFile = 0
        If FileName = "" Then FileName = MyINIFile
        Try
            GetIntegerIniFile = hs.GetINISetting(Section, EncodeINIKey(Key), DefaultVal, FileName)
            'Log("GetIntegerIniFile called with section = " & Section & " and Key = " & Key & " read = " & GetIntegerIniFile.ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            Log("Error in GetIntegerIniFile with section = " & Section & " and Key = " & EncodeINIKey(Key) & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function GetBooleanIniFile(ByVal Section As String, ByVal Key As String, ByVal DefaultVal As Boolean, Optional FileName As String = "") As Boolean
        GetBooleanIniFile = False
        If FileName = "" Then FileName = MyINIFile
        Try
            GetBooleanIniFile = hs.GetINISetting(Section, EncodeINIKey(Key), DefaultVal, FileName)
            'Log("GetBooleanIniFile called with section = " & Section & " and Key = " & Key & " read = " & GetBooleanIniFile.ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            Log("Error in GetBooleanIniFile with section = " & Section & " and Key = " & EncodeINIKey(Key) & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Sub WriteStringIniFile(ByVal Section As String, ByVal Key As String, ByVal Value As String, Optional FileName As String = "")
        If FileName = "" Then FileName = MyINIFile
        Try
            hs.SaveINISetting(Section, EncodeINIKey(Key), Value, FileName)
            'Log("WriteStringIniFile called with section = " & Section & " and Key = " & Key & " Value = " & Value.ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            Log("Error in WriteStringIniFile with section = " & Section & " and Key = " & EncodeINIKey(Key) & " and Value = " & Value.ToString & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub WriteIntegerIniFile(ByVal Section As String, ByVal Key As String, ByVal Value As Integer, Optional FileName As String = "")
        If FileName = "" Then FileName = MyINIFile
        Try
            hs.SaveINISetting(Section, EncodeINIKey(Key), Value, FileName)
            'Log("WriteIntegerIniFile called with section = " & Section & " and Key = " & Key & " Value = " & Value.ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            Log("Error in WriteIntegerIniFile writing section  " & Section & " and Key = " & EncodeINIKey(Key) & " and Value = " & Value.ToString & " with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub WriteBooleanIniFile(ByVal Section As String, ByVal Key As String, ByVal Value As Boolean, Optional FileName As String = "")
        If FileName = "" Then FileName = MyINIFile
        Try
            hs.SaveINISetting(Section, EncodeINIKey(Key), Value, FileName)
            'Log("WriteBooleanIniFile called with section = " & Section & " and Key = " & Key & " Value = " & Value.ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            Log("Error in WriteBooleanIniFile writing section  " & Section & " and Key = " & EncodeINIKey(Key) & " and Value = " & Value.ToString & " with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub DeleteEntryIniFile(Section As String, Key As String, Optional FileName As String = "")
        If FileName = "" Then FileName = MyINIFile
        Try
            hs.SaveINISetting(Section, EncodeINIKey(Key), Nothing, MyINIFile)
            If piDebuglevel > DebugLevel.dlEvents  Then Log("DeleteEntryIniFile called with section = " & Section & " and Key = " & EncodeINIKey(Key), LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            Log("Error in DeleteEntryIniFile reading " & Section & " section with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Function GetIniSection(ByVal Section As String, Optional FileName As String = "") As Dictionary(Of String, String)
        GetIniSection = Nothing
        If FileName = "" Then FileName = MyINIFile
        Try
            Dim ReturnStrings As String() = hs.GetINISectionEx(Section, FileName)
            If ReturnStrings Is Nothing Then Exit Function
            If piDebuglevel > DebugLevel.dlEvents  Then Log("GetIniSection called with section = " & Section & ", FileName = " & FileName & " and # Result = " & UBound(ReturnStrings, 1).ToString, LogType.LOG_TYPE_INFO)
            Dim KeyValues As New Dictionary(Of String, String)()
            For Each Entry In ReturnStrings
                'If piDebuglevel > DebugLevel.dlErrorsOnly  Then Log("GetIniSection called with section = " & Section & " found entry = " & Entry.ToString, LogType.LOG_TYPE_INFO)
                Dim Values As String() = Split(Entry, "=")
                If Not Entry Is Nothing And Entry <> "" Then
                    If UBound(Values, 1) > 0 Then
                        KeyValues.Add(DecodeINIKey(Values(0)), Entry.Substring(Values(0).Length + 1, Entry.Length - Values(0).Length - 1))
                    Else
                        KeyValues.Add(Values(0), "")
                    End If
                End If
            Next
            Return KeyValues
        Catch ex As Exception
            Log("Error in GetIniSection reading " & Section & " section with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Sub DeleteIniSection(ByVal Section As String, Optional FileName As String = "")
        If FileName = "" Then FileName = MyINIFile
        Try
            If piDebuglevel > DebugLevel.dlErrorsOnly  Then Log("DeleteIniSection called with section = " & Section & " and FileName = " & FileName, LogType.LOG_TYPE_INFO)
            hs.ClearINISection(Section, FileName)
        Catch ex As Exception
            Log("Error in DeleteIniSection deleting " & Section & " section with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

#Else

    Public Sub Log(ByVal msg As String, Optional ByVal logType As LogType = LogType.LOG_TYPE_INFO, Optional ByVal MsgColor As String = "", Optional ErrorCode As Integer = 0)
        Try
            If msg Is Nothing Then msg = ""
            If Not [Enum].IsDefined(GetType(LogType), logType) Then
                logType = LogType.LOG_TYPE_ERROR
            End If

            If Not ImRunningOnLinux Then Console.WriteLine(DateAndTime.Now.ToString & " : " & msg)

            Select Case logType
                Case LogType.LOG_TYPE_ERROR
                    If MsgColor <> "" Then
                        myHomeSeerSystem?.WriteLog(HomeSeer.PluginSdk.Logging.ELogType.Error, msg, shortIfaceName, MsgColor)
                    Else
                        myHomeSeerSystem?.WriteLog(HomeSeer.PluginSdk.Logging.ELogType.Error, msg, shortIfaceName)
                    End If
                Case LogType.LOG_TYPE_WARNING
                    If MsgColor <> "" Then
                        myHomeSeerSystem?.WriteLog(HomeSeer.PluginSdk.Logging.ELogType.Warning, msg, shortIfaceName, MsgColor)
                    Else
                        myHomeSeerSystem?.WriteLog(HomeSeer.PluginSdk.Logging.ELogType.Warning, msg, shortIfaceName)
                    End If
                Case LogType.LOG_TYPE_INFO
                    If MsgColor <> "" Then
                        myHomeSeerSystem?.WriteLog(HomeSeer.PluginSdk.Logging.ELogType.Info, msg, shortIfaceName, MsgColor)
                    Else
                        myHomeSeerSystem?.WriteLog(HomeSeer.PluginSdk.Logging.ELogType.Info, msg, shortIfaceName)
                    End If
            End Select
        Catch ex As Exception
            If Not ImRunningOnLinux Then Console.WriteLine("Exception in LOG: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        Try
            If MyLogFileName <> "" AndAlso gLogToDisk AndAlso logFileWriter IsNot Nothing Then
                SyncLock logFileLock
                    logFileWriter.WriteLine(DateAndTime.Now.ToString & " : " & logType.ToString & " - " & msg)
                    lastLogWrite = DateTime.Now
                    'logFileWriter.Flush() ' Force write
                End SyncLock
            End If
        Catch ex As Exception
            logFileWriter = Nothing
            MyLogFileName = ""
            If Not ImRunningOnLinux Then Console.WriteLine("Exception in LOG write: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Function GetStringIniFile(ByVal Section As String, ByVal Key As String, ByVal DefaultVal As String, Optional FileName As String = "") As String
        GetStringIniFile = ""
        If FileName = "" Then FileName = myINIFile
        Try
            GetStringIniFile = myHomeSeerSystem.GetINISetting(Section, EncodeINIKey(Key), DefaultVal, FileName)
            'Log("GetStringIniFile called with section = " & Section & " and Key = " & Key & " read = " & GetStringIniFile, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then Log("Error in GetStringIniFile with section = " & Section & " and Key = " & EncodeINIKey(Key) & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function GetIntegerIniFile(ByVal Section As String, ByVal Key As String, ByVal DefaultVal As Integer, Optional FileName As String = "") As Integer
        GetIntegerIniFile = 0
        If FileName = "" Then FileName = myINIFile
        Try
            GetIntegerIniFile = myHomeSeerSystem.GetINISetting(Section, EncodeINIKey(Key), DefaultVal, FileName)
            'Log("GetIntegerIniFile called with section = " & Section & " and Key = " & Key & " read = " & GetIntegerIniFile.ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then Log("Error in GetIntegerIniFile with section = " & Section & " and Key = " & EncodeINIKey(Key) & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function GetBooleanIniFile(ByVal Section As String, ByVal Key As String, ByVal DefaultVal As Boolean, Optional FileName As String = "") As Boolean
        GetBooleanIniFile = False
        If FileName = "" Then FileName = myINIFile
        Try
            GetBooleanIniFile = myHomeSeerSystem.GetINISetting(Section, EncodeINIKey(Key), DefaultVal, FileName)
            'Log("GetBooleanIniFile called with section = " & Section & " and Key = " & Key & " read = " & GetBooleanIniFile.ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then Log("Error in GetBooleanIniFile with section = " & Section & " and Key = " & EncodeINIKey(Key) & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Sub WriteStringIniFile(ByVal Section As String, ByVal Key As String, ByVal Value As String, Optional FileName As String = "")
        If FileName = "" Then FileName = myINIFile
        Try
            myHomeSeerSystem.SaveINISetting(Section, EncodeINIKey(Key), Value, FileName)
            'Log("WriteStringIniFile called with section = " & Section & " and Key = " & Key & " Value = " & Value.ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then Log("Error in WriteStringIniFile with section = " & Section & " and Key = " & EncodeINIKey(Key) & " and Value = " & Value.ToString & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub WriteIntegerIniFile(ByVal Section As String, ByVal Key As String, ByVal Value As Integer, Optional FileName As String = "")
        If FileName = "" Then FileName = myINIFile
        Try
            myHomeSeerSystem.SaveINISetting(Section, EncodeINIKey(Key), Value, FileName)
            'Log("WriteIntegerIniFile called with section = " & Section & " and Key = " & Key & " Value = " & Value.ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then Log("Error in WriteIntegerIniFile writing section  " & Section & " and Key = " & EncodeINIKey(Key) & " and Value = " & Value.ToString & " with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub WriteBooleanIniFile(ByVal Section As String, ByVal Key As String, ByVal Value As Boolean, Optional FileName As String = "")
        If FileName = "" Then FileName = myINIFile
        Try
            myHomeSeerSystem.SaveINISetting(Section, EncodeINIKey(Key), Value, FileName)
            'Log("WriteBooleanIniFile called with section = " & Section & " and Key = " & Key & " Value = " & Value.ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then Log("Error in WriteBooleanIniFile writing section  " & Section & " and Key = " & EncodeINIKey(Key) & " and Value = " & Value.ToString & " with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub DeleteEntryIniFile(Section As String, Key As String, Optional FileName As String = "")
        If FileName = "" Then FileName = myINIFile
        Try
            myHomeSeerSystem.SaveINISetting(Section, EncodeINIKey(Key), Nothing, myINIFile)
            If piDebuglevel > DebugLevel.dlEvents Then Log("DeleteEntryIniFile called with section = " & Section & " and Key = " & EncodeINIKey(Key), LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then Log("Error in DeleteEntryIniFile reading " & Section & " section with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Function GetIniSection(ByVal Section As String, Optional FileName As String = "") As Dictionary(Of String, String)
        GetIniSection = Nothing
        If FileName = "" Then FileName = myINIFile
        Try
            Return myHomeSeerSystem.GetIniSection(Section, FileName)
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then Log("Error in GetIniSection reading " & Section & " section with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Sub DeleteIniSection(ByVal Section As String, Optional FileName As String = "")
        If FileName = "" Then FileName = myINIFile
        Try
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("DeleteIniSection called with section = " & Section & " and FileName = " & FileName, LogType.LOG_TYPE_INFO)
            myHomeSeerSystem.ClearIniSection(Section, FileName)
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then Log("Error in DeleteIniSection deleting " & Section & " section with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub SetFeatureValueByRef(FeatureRef As Integer, NewValue As Object, Optional Update As Boolean = False)
        If piDebuglevel > DebugLevel.dlEvents Then Log("SetFeatureValueByRef called with Ref = " & FeatureRef.ToString & " and NewValue = " & NewValue.ToString, LogType.LOG_TYPE_INFO)
        Try
            'myHomeSeerSystem.UpdatePropertyByRef(FeatureRef, Devices.EProperty.Value, Convert.ToDouble(NewValue)) ' issue in v24
            myHomeSeerSystem.UpdateFeatureValueByRef(FeatureRef, Convert.ToDouble(NewValue))
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SetFeatureValueByRef called with Ref = " & FeatureRef.ToString & " and NewValue = " & NewValue.ToString & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub SetFeatureStringByRef(FeatureRef As Integer, NewValue As String, Update As Boolean)
        If FeatureRef = -1 Then Exit Sub
        Try
            If piDebuglevel > DebugLevel.dlEvents Then Log("SetFeatureStringByRef called with Ref = " & FeatureRef.ToString & " and NewValue = " & NewValue.ToString, LogType.LOG_TYPE_INFO)
            myHomeSeerSystem.UpdateFeatureValueStringByRef(FeatureRef, NewValue.ToString)
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SetFeatureStringByRef called with Ref = " & FeatureRef.ToString & " and NewValue = " & NewValue.ToString & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub SetHideFlagFeature(featureRef As Integer, flagOn As Boolean)
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SetHideFlagFeature called with Ref = " & featureRef.ToString & " and flagOn = " & flagOn.ToString, LogType.LOG_TYPE_INFO)
        If featureRef = -1 Then Exit Sub
        Try
            Dim df As Devices.HsFeature = myHomeSeerSystem.GetFeatureByRef(featureRef)
            Dim miscFlags As UInteger = df.Misc
            If flagOn Then
                miscFlags = miscFlags Or EMiscFlag.Hidden Or EMiscFlag.HideInMobile
            Else
                miscFlags = miscFlags And Not (EMiscFlag.Hidden) And Not (EMiscFlag.HideInMobile)
            End If

            Dim newMisc As New Dictionary(Of HomeSeer.PluginSdk.Devices.EProperty, Object) From {
                {EProperty.Misc, miscFlags}
            }
            myHomeSeerSystem.UpdateFeatureByRef(featureRef, newMisc)
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SetHideFlagFeature called with Ref = " & featureRef.ToString & " and flagOn = " & flagOn.ToString & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub SetHideFlagDevice(deviceRef As Integer, flagOn As Boolean)
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SetHideFlagDevice called with Ref = " & deviceRef.ToString & " and flagOn = " & flagOn.ToString, LogType.LOG_TYPE_INFO)
        If deviceRef = -1 Then Exit Sub
        Try
            Dim dv As Devices.HsDevice = myHomeSeerSystem.GetDeviceByRef(deviceRef)
            Dim miscFlags As UInteger = dv.Misc
            If flagOn Then
                miscFlags = miscFlags Or EMiscFlag.Hidden Or EMiscFlag.HideInMobile
            Else
                miscFlags = miscFlags And Not (EMiscFlag.Hidden) And Not (EMiscFlag.HideInMobile)
            End If

            Dim newMisc As New Dictionary(Of HomeSeer.PluginSdk.Devices.EProperty, Object) From {
                {EProperty.Misc, miscFlags}
            }
            myHomeSeerSystem.UpdateDeviceByRef(deviceRef, newMisc)
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SetHideFlagDevice called with Ref = " & deviceRef.ToString & " and flagOn = " & flagOn.ToString & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Function GetPicture(ByVal url As String) As System.Drawing.Image
        ' Get the picture at a given URL.
        Dim web_client As New WebClient With {
            .UseDefaultCredentials = True                         ' added 5/2/2019
            }
        GetPicture = Nothing
        Try
            url = Trim(url)
            If url = "" Then
                Return Nothing
                Exit Function
            End If
            If Not (url.ToLower().StartsWith("http://") Or url.ToLower().StartsWith("https://") Or url.ToLower().StartsWith("file:")) Then url = "http://" & url
            Dim image_stream As New MemoryStream(web_client.DownloadData(url))
            GetPicture = Image.FromStream(image_stream, True, True)
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetPicture called with url= " & url.ToString & " caused error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        Finally
            web_client.Dispose()
        End Try
    End Function

    Public Sub DumpInfoHSDevice(HSRef_ As Integer)
        If HSRef_ <> -1 Then
            Try
                Dim dv As HsDevice = myHomeSeerSystem.GetDeviceByRef(HSRef_)
                If dv IsNot Nothing Then
                    Log("Device Name = " & dv.Name, LogType.LOG_TYPE_WARNING)
                    Log("Device Ref = " & dv.Ref, LogType.LOG_TYPE_WARNING)
                    Log("--Device Image = " & dv.Image, LogType.LOG_TYPE_WARNING)
                    Log("--Device Interface = " & dv.Interface, LogType.LOG_TYPE_WARNING)
                    Log("--Device IsValueInvalid = " & dv.IsValueInvalid, LogType.LOG_TYPE_WARNING)
                    Log("--Device Relationship = " & dv.Relationship.ToString(), LogType.LOG_TYPE_WARNING)
                    Log("--Device Status = " & dv.StatusString, LogType.LOG_TYPE_WARNING)
                    Log("--Device Value = " & dv.Value.ToString(), LogType.LOG_TYPE_WARNING)
                    Log("--Device TypeInfo.Type = " & dv.TypeInfo.Type.ToString(), LogType.LOG_TYPE_WARNING)
                    Log("--Device TypeInfo.SubType = " & dv.TypeInfo.SubType.ToString(), LogType.LOG_TYPE_WARNING)
                    Dim fvList As List(Of HsFeature) = dv.Features()
                    Dim adList As HashSet(Of Integer) = dv.AssociatedDevices()
                    If fvList IsNot Nothing AndAlso fvList.Count > 0 Then
                        For Each fv As HsFeature In fvList
                            Log("----Feature Name = " & fv.Name, LogType.LOG_TYPE_WARNING)
                            Log("----Feature Ref = " & fv.Ref, LogType.LOG_TYPE_WARNING)
                            Log("------Feature Image = " & fv.Image, LogType.LOG_TYPE_WARNING)
                            Log("------Feature Interface = " & fv.Interface, LogType.LOG_TYPE_WARNING)
                            Log("------Feature IsValueInvalid = " & fv.IsValueInvalid, LogType.LOG_TYPE_WARNING)
                            Log("------Feature Relationship = " & fv.Relationship.ToString(), LogType.LOG_TYPE_WARNING)
                            Log("------Feature Status = " & fv.StatusString, LogType.LOG_TYPE_WARNING)
                            Log("------Feature Value = " & fv.Value.ToString(), LogType.LOG_TYPE_WARNING)
                            Dim scList As Controls.StatusControlCollection = fv.StatusControls()
                            Dim sgList As StatusGraphicCollection = fv.StatusGraphics()
                            If scList IsNot Nothing AndAlso scList.Count > 0 Then
                                Log("--------Status Control count = " & scList.Count.ToString(), LogType.LOG_TYPE_WARNING)
                                'For Each sc As Controls.StatusControl In scList
                                'Log("--------Status Control Name = " & sc.Name, LogType.LOG_TYPE_WARNING)
                                'sc.Column
                                'sc.ControlStates
                                'sc.ControlType
                                'sc.ControlUse
                                'sc.IsRange
                                'sc.Label
                                'sc.Location
                                'sc.Row
                                'sc.TargetRange
                                'sc.TargetValue
                                'sc.Width
                                'Next
                            End If
                            If sgList IsNot Nothing AndAlso sgList.Count > 0 Then
                                Log("--------Status Graphics count = " & sgList.Count.ToString(), LogType.LOG_TYPE_WARNING)
                                'For Each sg As StatusGraphic In sgList

                                'Next
                            End If
                        Next
                    End If
                    If adList IsNot Nothing AndAlso adList.Count > 0 Then
                        Log("----Associated devices count = " & adList.Count.ToString(), LogType.LOG_TYPE_WARNING)
                        For Each featRef_ As Integer In adList
                            Dim fv As HsFeature = myHomeSeerSystem.GetFeatureByRef(featRef_)
                            If fv IsNot Nothing Then
                                Log("----Feature Name = " & fv.Name, LogType.LOG_TYPE_WARNING)
                                Log("----Feature Ref = " & fv.Ref, LogType.LOG_TYPE_WARNING)
                                Log("------Feature Image = " & fv.Image, LogType.LOG_TYPE_WARNING)
                                Log("------Feature Interface = " & fv.Interface, LogType.LOG_TYPE_WARNING)
                                Log("------Feature IsValueInvalid = " & fv.IsValueInvalid, LogType.LOG_TYPE_WARNING)
                                Log("------Feature Relationship = " & fv.Relationship.ToString(), LogType.LOG_TYPE_WARNING)
                                Log("------Feature Status = " & fv.StatusString, LogType.LOG_TYPE_WARNING)
                                Log("------Feature Value = " & fv.Value.ToString(), LogType.LOG_TYPE_WARNING)
                                Log("------Feature TypeInfo.Type = " & fv.TypeInfo.Type.ToString(), LogType.LOG_TYPE_WARNING)
                                Log("------Feature TypeInfo.SubType = " & fv.TypeInfo.SubType.ToString(), LogType.LOG_TYPE_WARNING)
                                Dim scList As Controls.StatusControlCollection = fv.StatusControls()
                                Dim sgList As StatusGraphicCollection = fv.StatusGraphics()
                                If scList IsNot Nothing AndAlso scList.Count > 0 Then
                                    Log("--------Status Control count = " & scList.Count.ToString(), LogType.LOG_TYPE_WARNING)
                                    For Each sc As Controls.StatusControl In scList.Values
                                        Log("--------Status Control Label = " & sc.Label, LogType.LOG_TYPE_WARNING)
                                        Log("--------Status Control Column = " & sc.Column, LogType.LOG_TYPE_WARNING)

                                        Log("--------Status Control ControlType = " & sc.ControlType, LogType.LOG_TYPE_WARNING)
                                        Log("--------Status Control ControlUse = " & sc.ControlUse, LogType.LOG_TYPE_WARNING)
                                        Log("--------Status Control IsRange = " & sc.IsRange, LogType.LOG_TYPE_WARNING)
                                        Log("--------Status Control Location = " & sc.Location.ToString, LogType.LOG_TYPE_WARNING)
                                        Log("--------Status Control Row = " & sc.Row, LogType.LOG_TYPE_WARNING)
                                        'Log("--------Status Control TargetRange = " & sc.TargetRange, LogType.LOG_TYPE_WARNING)
                                        Log("--------Status Control TargetValue = " & sc.TargetValue, LogType.LOG_TYPE_WARNING)
                                        Log("--------Status Control Width = " & sc.Width, LogType.LOG_TYPE_WARNING)
                                        Dim ListCs As List(Of String) = sc.ControlStates
                                        If ListCs IsNot Nothing AndAlso ListCs.Count > 0 Then
                                            For Each scState As String In ListCs
                                                Log("----------Status Control ControlStates = " & scState, LogType.LOG_TYPE_WARNING)
                                            Next
                                        End If
                                    Next
                                End If
                                If sgList IsNot Nothing AndAlso sgList.Count > 0 Then
                                    Log("--------Status Graphics count = " & sgList.Count.ToString(), LogType.LOG_TYPE_WARNING)
                                    For Each sg As StatusGraphic In sgList.Values
                                        Log("--------Status Graphics Graphic = " & sg.Graphic, LogType.LOG_TYPE_WARNING)
                                        Log("--------Status Graphics IsRange = " & sg.IsRange, LogType.LOG_TYPE_WARNING)
                                        Log("--------Status Graphics Label = " & sg.Label, LogType.LOG_TYPE_WARNING)
                                        Log("--------Status Graphics RangeMax = " & sg.RangeMax, LogType.LOG_TYPE_WARNING)
                                        Log("--------Status Graphics RangeMin = " & sg.RangeMin, LogType.LOG_TYPE_WARNING)
                                        Log("--------Status Graphics Value = " & sg.Value, LogType.LOG_TYPE_WARNING)
                                        Log("--------Status Graphics GetType = " & sg.GetType.ToString, LogType.LOG_TYPE_WARNING)
                                    Next
                                End If
                            End If
                        Next
                    End If
                End If
            Catch ex As Exception

            End Try
        End If


    End Sub

    Public Function ExtractVolumeAttributes(value As String) As VolumeAttributes
        If value.Trim.Length = 0 Then Return Nothing
        Try
            Dim returnInfo As New VolumeAttributes
            Dim useHSRef As Integer = value.ToUpper.IndexOf("$DVR:")   ' added on 10/10/2021 in v4.0.1.4
            If useHSRef <> -1 Then
                value = value.Substring(5)  ' remove the $DVR:
                Dim tempRef As String = value
                value = If(myHomeSeerSystem.GetFeatureByRef(value)?.Value, 0)
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log(" extractVolumeAttributes retrieved device value = " & value & " for devRef = " & tempRef, LogType.LOG_TYPE_INFO)
            Else
                useHSRef = value.ToUpper.IndexOf("$DSR:")
                If useHSRef <> -1 Then
                    value = value.Substring(5)  ' remove the $DSR:
                    Dim tempRef As String = value
                    value = If(myHomeSeerSystem.GetFeatureByRef(value)?.StatusString, 0)
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log(" extractVolumeAttributes retrieved device value = " & value & " for devRef = " & tempRef, LogType.LOG_TYPE_INFO)
                End If
            End If
            Dim gPosition As Integer = value.IndexOf("G")
            If gPosition <> -1 Then
                value = value.Remove(gPosition, 1)
                returnInfo.applyToGroup = True
            End If
            gPosition = value.IndexOf("g")
            If gPosition <> -1 Then
                value = value.Remove(gPosition, 1)
                returnInfo.applyToGroup = True
            End If
            Dim plusPosition As Integer = value.IndexOf("+")
            If plusPosition <> -1 Then
                value = value.Remove(plusPosition, 1)
                returnInfo.relative = True
            End If
            Dim minusPosition As Integer = value.IndexOf("-")
            If minusPosition <> -1 Then
                value = value.Remove(minusPosition, 1)
                returnInfo.relative = True
            End If
            value = value.Trim
            If value <> "" Then
                returnInfo.volume = CInt(value)
            Else
                Return Nothing
            End If
            If minusPosition <> -1 Then returnInfo.volume = -returnInfo.volume
            Return returnInfo
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in extractVolumeAttributes for value = " & value & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Return Nothing
        End Try
    End Function
#End If

    Public Function EncodeURI(ByVal InString As String) As String
        EncodeURI = InString
        Dim InIndex As Integer = 0
        Dim Outstring As String = ""
        InString = Trim(InString)
        If InString = "" Then Exit Function
        Try
            Do While InIndex < InString.Length
                If InString(InIndex) = " " Then
                    Outstring += "%20"
                ElseIf InString(InIndex) = "!" Then
                    Outstring += "%21"
                ElseIf InString(InIndex) = """" Then
                    Outstring += "%22"
                ElseIf InString(InIndex) = "#" Then
                    Outstring += "%23"
                ElseIf InString(InIndex) = "$" Then
                    Outstring += "%24"
                ElseIf InString(InIndex) = "%" Then
                    Outstring += "%25"
                ElseIf InString(InIndex) = "&" Then
                    Outstring += "%26"
                ElseIf InString(InIndex) = "'" Then
                    Outstring += "%27"
                ElseIf InString(InIndex) = "(" Then
                    Outstring += "%28"
                ElseIf InString(InIndex) = ")" Then
                    Outstring += "%29"
                ElseIf InString(InIndex) = "*" Then
                    Outstring += "%2A"
                ElseIf InString(InIndex) = "+" Then
                    Outstring += "%2B"
                ElseIf InString(InIndex) = "," Then
                    Outstring += "%2C"
                ElseIf InString(InIndex) = "-" Then
                    Outstring += "%2D"
                ElseIf InString(InIndex) = "." Then
                    Outstring += "%2E"
                ElseIf InString(InIndex) = "/" Then
                    Outstring += "%2F"
                Else
                    Outstring &= InString(InIndex)
                End If
                InIndex += 1
            Loop
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then Log("Error in EncodeURI. URI = " & InString & " Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        EncodeURI = Outstring
    End Function

    Public Function DecodeURI(ByVal InString As String) As String
        DecodeURI = InString
        Dim InIndex As Integer = 0
        Dim Outstring As String = ""
        Dim Value As Integer
        InString = Trim(InString)
        If InString = "" Then Exit Function
        Try
            Do While InIndex < InString.Length
                If InString(InIndex) = "%" Then
                    ' this might be the start of an encode
                    ' first check length
                    If InIndex + 2 <= InString.Length Then
                        ' now check whether next two characters are hex
                        If ((InString(InIndex + 1) >= "0") And (InString(InIndex + 1) <= "9")) Or
                            ((InString(InIndex + 1) >= "A") And (InString(InIndex + 1) <= "F")) Or
                             ((InString(InIndex + 1) >= "a") And (InString(InIndex + 1) <= "f")) Then
                            ' now check the same for the second character
                            If ((InString(InIndex + 2) >= "0") And (InString(InIndex + 2) <= "9")) Or
                                ((InString(InIndex + 2) >= "A") And (InString(InIndex + 2) <= "F")) Or
                                ((InString(InIndex + 2) >= "a") And (InString(InIndex + 2) <= "f")) Then
                                ' all checks passed now convert
                                Value = 0
                                Dim Char1, Char2 As Char
                                Char1 = UCase(InString(InIndex + 1))
                                Char2 = UCase(InString(InIndex + 2))
                                If (Char1 >= "0") And (Char1 <= "9") Then
                                    Value += Val(Char1) * 16
                                Else
                                    ' convert 
                                    Char1 = ChrW(AscW(Char1) - 17)
                                    Value += (Val(Char1) + 10) * 16
                                End If
                                If (Char2 >= "0") And (Char2 <= "9") Then
                                    Value += Val(Char2)
                                Else
                                    ' convert 
                                    Char2 = ChrW(AscW(Char2) - 17)
                                    Value += (Val(Char2) + 10)
                                End If
                                Outstring &= ChrW(Value)
                                InIndex += 3
                            Else
                                Outstring &= InString(InIndex)
                                InIndex += 1
                            End If
                        Else
                            Outstring &= InString(InIndex)
                            InIndex += 1
                        End If
                    Else
                        Outstring &= InString(InIndex)
                        InIndex += 1
                    End If
                    'ElseIf InString(InIndex) = "&" Then
                    '             string = string.replace(/&amp;/g, "&");  
                    '             string = string.replace(/&quot;/g, "\"");  
                    '             string = string.replace(/&apos;/g, "'");  
                    '             string = string.replace(/&lt;/g, "<");  
                    '             string = string.replace(/&gt;/g, ">"); 
                Else
                    Outstring &= InString(InIndex)
                    InIndex += 1
                End If
            Loop
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then Log("Error in DecodeURI with URI = " & InString & " and Index= " & InIndex.ToString & " and Value = " & Value.ToString & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        DecodeURI = Outstring
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "DecodeURI: In = " & InString & " out = " & Outstring)
    End Function

    Public Function EncodeINIKey(ByVal InString As String) As String
        EncodeINIKey = InString
        Dim InIndex As Integer = 0
        Dim Outstring As String = ""
        InString = Trim(InString)
        If InString = "" Then Exit Function
        Try
            Do While InIndex < InString.Length
                If InString(InIndex) = "=" Then
                    Outstring += "%3D"
                ElseIf InString(InIndex) = "%" Then
                    Outstring += "%25"
                Else
                    Outstring &= InString(InIndex)
                End If
                InIndex += 1
            Loop
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then Log("Error in EncodeINIKey. URI = " & InString & " Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        EncodeINIKey = Outstring
    End Function

    Public Function DecodeINIKey(ByVal InString As String) As String
        DecodeINIKey = InString
        Dim InIndex As Integer = 0
        Dim Outstring As String = ""
        Dim Value As Integer
        InString = Trim(InString)
        If InString = "" Then Exit Function
        Try
            Do While InIndex < InString.Length
                If InString(InIndex) = "%" Then
                    ' this might be the start of an encode
                    ' first check length
                    If InIndex + 2 <= InString.Length Then
                        ' now check whether next two characters are hex
                        If ((InString(InIndex + 1) >= "0") And (InString(InIndex + 1) <= "9")) Or
                            ((InString(InIndex + 1) >= "A") And (InString(InIndex + 1) <= "F")) Or
                             ((InString(InIndex + 1) >= "a") And (InString(InIndex + 1) <= "f")) Then
                            ' now check the same for the second character
                            If ((InString(InIndex + 2) >= "0") And (InString(InIndex + 2) <= "9")) Or
                                ((InString(InIndex + 2) >= "A") And (InString(InIndex + 2) <= "F")) Or
                                ((InString(InIndex + 2) >= "a") And (InString(InIndex + 2) <= "f")) Then
                                ' all checks passed now convert
                                Value = 0
                                Dim Char1, Char2 As Char
                                Char1 = UCase(InString(InIndex + 1))
                                Char2 = UCase(InString(InIndex + 2))
                                If (Char1 >= "0") And (Char1 <= "9") Then
                                    Value += Val(Char1) * 16
                                Else
                                    ' convert 
                                    Char1 = ChrW(AscW(Char1) - 17)
                                    Value += (Val(Char1) + 10) * 16
                                End If
                                If (Char2 >= "0") And (Char2 <= "9") Then
                                    Value += Val(Char2)
                                Else
                                    ' convert 
                                    Char2 = ChrW(AscW(Char2) - 17)
                                    Value += (Val(Char2) + 10)
                                End If
                                Outstring &= ChrW(Value)
                                InIndex += 3
                            Else
                                Outstring &= InString(InIndex)
                                InIndex += 1
                            End If
                        Else
                            Outstring &= InString(InIndex)
                            InIndex += 1
                        End If
                    Else
                        Outstring &= InString(InIndex)
                        InIndex += 1
                    End If
                    'ElseIf InString(InIndex) = "&" Then
                    '             string = string.replace(/&amp;/g, "&");  
                    '             string = string.replace(/&quot;/g, "\"");  
                    '             string = string.replace(/&apos;/g, "'");  
                    '             string = string.replace(/&lt;/g, "<");  
                    '             string = string.replace(/&gt;/g, ">"); 
                Else
                    Outstring &= InString(InIndex)
                    InIndex += 1
                End If
            Loop
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then Log("Error in DecodeINIKey with URI = " & InString & " and Index= " & InIndex.ToString & " and Value = " & Value.ToString & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        DecodeINIKey = Outstring
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "DecodeURI: In = " & InString & " out = " & Outstring)
    End Function

    Public Function EscapeForJson(input As String) As String
        Return input.Replace("\", "\\").Replace("""", "\""")
    End Function

    Public Function DeriveIPAddress(inString As String, NextChar As String) As String
        If piDebuglevel > DebugLevel.dlEvents Then Log("DeriveIPAddress called for and inString = " & inString & " and NextChar = " & NextChar, LogType.LOG_TYPE_INFO)
        DeriveIPAddress = inString
        Dim NewURLDoc As String = Trim(inString)
        Dim httpIndex As Integer = 0
        If NextChar <> "" And (NextChar(0) <> "/" And NextChar(0) <> "\") Then
            httpIndex = NewURLDoc.LastIndexOf("/")
        ElseIf NewURLDoc.ToUpper.IndexOf("HTTP://") <> -1 Then
            NewURLDoc = NewURLDoc.Remove(0, 7)
            httpIndex = NewURLDoc.IndexOf("/") + 6
        ElseIf NewURLDoc.ToUpper.IndexOf("HTTP:\\") <> -1 Then
            NewURLDoc = NewURLDoc.Remove(0, 7)
            httpIndex = NewURLDoc.IndexOf("/") + 6
        Else
            httpIndex = NewURLDoc.LastIndexOf("/")
        End If
        If httpIndex <> -1 Then
            DeriveIPAddress = inString.Substring(0, httpIndex + 1)
        End If
    End Function

    Public Function RemoveControlCharacters(inString As String) As String
        RemoveControlCharacters = inString
        Dim OutString As String = ""
        Try
            ' remove any control characters
            Dim strIndex As Integer = inString.Length
            If piDebuglevel > DebugLevel.dlEvents Then Log("RemoveControlCharacters retrieved document with length = " & strIndex.ToString, LogType.LOG_TYPE_INFO)
            Dim SomethingGotRemoved As Boolean = False
            While strIndex > 0
                strIndex -= 1
                If inString(strIndex) < " " Then
                    inString = inString.Remove(strIndex, 1)
                    SomethingGotRemoved = True
                End If
            End While
            inString = Trim(inString)
            'If piDebuglevel > DebugLevel.dlEvents And SomethingGotRemoved Then Log("RemoveControlCharacters updated document to = " & inString.ToString, LogType.LOG_TYPE_INFO)
            RemoveControlCharacters = inString
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then Log("Error in RemoveControlCharacters while retieving document with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function ReplaceSpecialCharacters(ByVal InString As String) As String
        ReplaceSpecialCharacters = InString
        Dim InIndex As Integer = 0
        Dim Outstring As String = ""
        InString = Trim(InString)
        If InString = "" Then Exit Function
        Try
            Do While InIndex < InString.Length
                If InString(InIndex) = " " Then
                    Outstring += "_"
                ElseIf InString(InIndex) = "!" Then
                    Outstring += "_"
                ElseIf InString(InIndex) = """ Then" Then
                    Outstring += "_"
                ElseIf InString(InIndex) = "#" Then
                    Outstring += "_"
                ElseIf InString(InIndex) = "$" Then
                    Outstring += "_"
                ElseIf InString(InIndex) = "%" Then
                    Outstring += "_"
                ElseIf InString(InIndex) = "&" Then
                    Outstring += "_"
                ElseIf InString(InIndex) = "'" Then
                    Outstring += "_"
                ElseIf InString(InIndex) = "(" Then
                    Outstring += "_"
                ElseIf InString(InIndex) = ")" Then
                    Outstring += "_"
                ElseIf InString(InIndex) = "*" Then
                    Outstring += "_"
                ElseIf InString(InIndex) = "+" Then
                    Outstring += "_"
                ElseIf InString(InIndex) = "," Then
                    Outstring += "_"
                ElseIf InString(InIndex) = "-" Then
                    Outstring += "_"
                ElseIf InString(InIndex) = "." Then
                    Outstring += "_"
                ElseIf InString(InIndex) = ":" Then
                    Outstring += "_"
                ElseIf InString(InIndex) = "/" Then
                    Outstring += "_"
                Else
                    Outstring &= InString(InIndex)
                End If
                InIndex += 1
            Loop
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then Log("Error in ReplaceSpecialCharacters. URI = " & InString & " Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        ReplaceSpecialCharacters = Outstring
    End Function

    Function ParseHTTPResponse(response As String, SearchItem As String) As String
        If SearchItem <> "" Then
            Dim ReturnString As String = (From line In response.Split({vbCr(0), vbLf(0)})
                                          Where (line.ToLowerInvariant().StartsWith(SearchItem))
                                          Select (line.Substring(SearchItem.Length + 1).Trim())).FirstOrDefault
            If ReturnString IsNot Nothing Then
                Return ReturnString
            Else
                Return ""
            End If
        Else
            ' this is to return the body which is separated from the header with a blank line
            Dim Lines As String() = Strings.Split(response, {vbCr(0), vbLf(0)})
            Dim EmptyLineFound As Boolean = False
            Dim Body As String = ""
            If Lines IsNot Nothing Then
                For Each Line As String In Lines
                    If Line IsNot Nothing Then
                        If EmptyLineFound Then
                            Body &= Line
                        Else
                            If Line = "" Then
                                EmptyLineFound = True
                            End If
                        End If
                    End If
                Next
                Return Trim(Body)
            End If
        End If
        Return ""
    End Function

    Function ParseHTTPResponseGetHeaderLength(response As String) As Integer
        ParseHTTPResponseGetHeaderLength = 0
        Dim Lines As String() = Split(response, {vbCr(0), vbLf(0)})
        Dim EmptyLineFound As Boolean = False
        Dim HeaderLength As Integer = 0
        For Each Line As String In Lines
            If EmptyLineFound Then
                Return HeaderLength
                Exit Function
            Else
                If Line = "" Then
                    EmptyLineFound = True
                End If
                HeaderLength = HeaderLength + 2 + Line.Length
            End If
        Next
    End Function

    Function GetHTTPCommand(response As String) As String()
        GetHTTPCommand = {}
        Dim Lines As String() = Split(response, {vbCr(0), vbLf(0)})
        If Lines IsNot Nothing Then
            If UBound(Lines) >= 0 Then
                Dim CommandParts As String() = Split(Lines(0), " ")
                If CommandParts IsNot Nothing Then
                    If UBound(CommandParts) >= 0 Then
                        Return CommandParts
                    End If
                End If
            End If
        End If
    End Function

    Function GetIpAddressType(ipString As String) As String
        Dim ip As IPAddress = Nothing
        Try
            If IPAddress.TryParse(ipString, ip) Then
                If ip IsNot Nothing Then
                    Select Case ip.AddressFamily
                        Case Sockets.AddressFamily.InterNetwork
                            Return "IPv4"
                        Case Sockets.AddressFamily.InterNetworkV6
                            Return "IPv6"
                        Case Else
                            Return ip.AddressFamily.ToString
                    End Select
                End If
            End If
        Catch ex As Exception
            Return ex.Message
        End Try
        Return "Invalid"
    End Function

    Public Function CheckLocalIPv4Address(in_address As String) As Boolean
        Try
            Dim strHostName As String = System.Net.Dns.GetHostName()
            Dim iphe As System.Net.IPHostEntry = System.Net.Dns.GetHostEntry(strHostName)
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("CheckLocalIPv4Address found IP Host = " & strHostName, LogType.LOG_TYPE_INFO)
            For Each ipheal As System.Net.IPAddress In iphe.AddressList
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("CheckLocalIPv4Address found IP Address = " & ipheal.ToString() & " with AddressFamily = " & ipheal.AddressFamily.ToString, LogType.LOG_TYPE_INFO)
                If ipheal.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork Then
                    If ipheal.ToString = in_address Then
                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("CheckLocalIPv4Address found IP Address = " & ipheal.ToString() & " equal to HS IP address so Plugin is running local", LogType.LOG_TYPE_INFO)
                        Return True
                    End If
                End If
            Next
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then Log("Error in CheckLocalIPv4Address with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Return False
    End Function

    Public Function GetLocalIPv4Address() As String
        GetLocalIPv4Address = ""
        Try
            Dim strHostName As String = System.Net.Dns.GetHostName()
            Dim iphe As System.Net.IPHostEntry = System.Net.Dns.GetHostEntry(strHostName)
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetLocalIPv4Address found IP Host = " & strHostName, LogType.LOG_TYPE_INFO)
            Dim NbrOfIPv4Interfaces As Integer = 0
            For Each ipheal As System.Net.IPAddress In iphe.AddressList
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetLocalIPv4Address found IP Address = " & ipheal.ToString() & " with AddressFamily = " & ipheal.AddressFamily.ToString, LogType.LOG_TYPE_INFO)
                If ipheal.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork Then
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetLocalIPv4Address found Local IP Address = " & ipheal.ToString(), LogType.LOG_TYPE_INFO)
                    NbrOfIPv4Interfaces += 1
                    GetLocalIPv4Address = ipheal.ToString()
                End If
            Next
            If NbrOfIPv4Interfaces > 1 Then
                ' Houston we have a problem, we need to have the user select one. Put a warning here for time being until I can fix it
                If piDebuglevel > DebugLevel.dlOff Then Log($"Warning in GetLocalIPv4Address. Found multiple local IP Addresses. Count = {NbrOfIPv4Interfaces}. Selected = {GetLocalIPv4Address }", LogType.LOG_TYPE_WARNING)
            End If
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then Log("Error in GetLocalIPv4Address with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function GetLocalMacAddress() As String
        If piDebuglevel > DebugLevel.dlEvents Then Log("GetLocalMacAddress called", LogType.LOG_TYPE_INFO)
        GetLocalMacAddress = ""
        Dim LocalMacAddress As String = ""
        Dim LocalIPAddress = plugInIPAddress
        If LocalIPAddress = "" Then
            If piDebuglevel > DebugLevel.dlOff Then Log("Error in GetLocalMacAddress trying to get own IP address", LogType.LOG_TYPE_ERROR)
            Exit Function
        End If
        Try
            For Each nic As System.Net.NetworkInformation.NetworkInterface In System.Net.NetworkInformation.NetworkInterface.GetAllNetworkInterfaces()
                If piDebuglevel > DebugLevel.dlEvents Then Log(String.Format("The MAC address of {0} is {1}{2}", nic.Description, Environment.NewLine, nic.GetPhysicalAddress()), LogType.LOG_TYPE_INFO)
                'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log(String.Format("The MAC ID of {0} is {1}{2}", nic.Description, Environment.NewLine, nic.Id.ToString), LogType.LOG_TYPE_INFO)
                'Log(String.Format("The MAC address of {0} is {1}{2}", nic.Description, Environment.NewLine, nic.GetPhysicalAddress()), LogType.LOG_TYPE_INFO)
                For Each Ipa In nic.GetIPProperties.UnicastAddresses
                    If piDebuglevel > DebugLevel.dlEvents Then Log(String.Format("The IPaddress address of {0} is {1}{2}", nic.Description, Environment.NewLine, Ipa.Address.ToString), LogType.LOG_TYPE_INFO)
                    'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log(String.Format("The IPaddress address of {0} is {1}{2}", nic.Description, Environment.NewLine, Ipa.Address.ToString), LogType.LOG_TYPE_INFO)
                    If Ipa.Address.ToString = LocalIPAddress Then
                        ' OK we found our IPaddress
                        LocalMacAddress = nic.GetPhysicalAddress().ToString
                        If piDebuglevel > DebugLevel.dlEvents Then Log("GetLocalMacAddress found local MAC address = " & LocalMacAddress, LogType.LOG_TYPE_INFO)
                        GetLocalMacAddress = LocalMacAddress
                    End If
                Next
            Next
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then Log("Error in GetLocalMacAddress trying to get own MAC address with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetLocalMacAddress trying to get own MAC address, none found", LogType.LOG_TYPE_ERROR)
    End Function

    Public Function IsDirectlyConnected(targetIpStr As String) As Boolean
        If Not IPAddress.TryParse(targetIpStr, Nothing) Then Return False

        Dim targetIp As IPAddress = IPAddress.Parse(targetIpStr)
        If targetIp.AddressFamily <> AddressFamily.InterNetwork Then Return False ' IPv4 only

        For Each nic As NetworkInterface In NetworkInterface.GetAllNetworkInterfaces()
            If nic.OperationalStatus <> OperationalStatus.Up Then Continue For

            For Each unicast In nic.GetIPProperties().UnicastAddresses
                If unicast.Address.AddressFamily <> AddressFamily.InterNetwork Then Continue For

                Dim localIp As IPAddress = unicast.Address
                Dim subnetMask As IPAddress = unicast.IPv4Mask
                If subnetMask Is Nothing Then Continue For

                If IsInSameSubnet(localIp, targetIp, subnetMask) Then
                    Return True
                End If
            Next
        Next

        Return False
    End Function

    Private Function IsInSameSubnet(ip1 As IPAddress, ip2 As IPAddress, mask As IPAddress) As Boolean
        Dim ip1Bytes = ip1.GetAddressBytes()
        Dim ip2Bytes = ip2.GetAddressBytes()
        Dim maskBytes = mask.GetAddressBytes()
        For i = 0 To 3
            If (ip1Bytes(i) And maskBytes(i)) <> (ip2Bytes(i) And maskBytes(i)) Then
                Return False
            End If
        Next
        Return True
    End Function

    Public Function GetEthernetPorts() As IEnumerable(Of NetworkInfo)
        If piDebuglevel > DebugLevel.dlEvents Then Log("GetEthernetPorts called", LogType.LOG_TYPE_INFO)
        Dim PortInfo As New List(Of NetworkInfo)
        'Dim Ethernetports As New Dictionary(Of String, String)
        Try
            For Each nic As System.Net.NetworkInformation.NetworkInterface In System.Net.NetworkInformation.NetworkInterface.GetAllNetworkInterfaces()
                Dim netInfo As New NetworkInfo With {
                    .description = nic.Description,
                    .mac = nic.GetPhysicalAddress().ToString,
                    .operationalstate = nic.OperationalStatus.ToString,
                    .id = nic.Id.ToString,
                    .name = nic.Name,
                    .interfacetype = nic.NetworkInterfaceType.ToString
}

                If piDebuglevel > DebugLevel.dlEvents Then Log(String.Format("The MAC address of {0} is {1} with status {2}", nic.Description, nic.GetPhysicalAddress().ToString, nic.OperationalStatus.ToString), LogType.LOG_TYPE_INFO)
                If piDebuglevel > DebugLevel.dlEvents Then Log(String.Format("  The ID of {0} has name {1} and interface type {2}", nic.Id.ToString, nic.Name, nic.NetworkInterfaceType.ToString), LogType.LOG_TYPE_INFO)
                If nic.OperationalStatus = Net.NetworkInformation.OperationalStatus.Up Then
                    For Each Ipa In nic.GetIPProperties.UnicastAddresses
                        If piDebuglevel > DebugLevel.dlEvents Then Log(String.Format("    The UniCast address of {0} is {1} with mask {2} and Address family {3}", nic.Description, Ipa.Address.ToString, Ipa.IPv4Mask.ToString, Ipa.Address.AddressFamily.ToString), LogType.LOG_TYPE_INFO)
                        If Ipa.Address.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork Then
                            ' we found an IPv4 IPaddress
                            Dim addrinfo As New NetworkInfoAddrMask With {
                                .address = Ipa.Address.ToString,
                                .mask = Ipa.IPv4Mask.ToString,
                                .addrtype = "IPv4 Unicast"
                            }
                            If netInfo.addressinfo Is Nothing Then netInfo.addressinfo = New List(Of NetworkInfoAddrMask)
                            netInfo.addressinfo.Add(addrinfo)
                        ElseIf Ipa.Address.AddressFamily = System.Net.Sockets.AddressFamily.InterNetworkV6 Then
                            Dim addrinfo As New NetworkInfoAddrMask With {
                                .address = Ipa.Address.ToString,
                                .addrtype = "IPv6 Unicast"
                            }
                            If netInfo.addressinfo Is Nothing Then netInfo.addressinfo = New List(Of NetworkInfoAddrMask)
                            netInfo.addressinfo.Add(addrinfo)
                        End If
                    Next
                    For Each Ipa In nic.GetIPProperties.AnycastAddresses
                        If piDebuglevel > DebugLevel.dlEvents Then Log(String.Format("    The AnyCast address of {0} is {1} with Address family {2}", nic.Description, Ipa.Address.ToString, Ipa.Address.AddressFamily.ToString), LogType.LOG_TYPE_INFO)
                        Dim addrinfo As New NetworkInfoAddrMask With {
                            .address = Ipa.Address.ToString,
                            .mask = Ipa.Address.AddressFamily.ToString,
                            .addrtype = "Anycast"
                        }
                        If netInfo.addressinfo Is Nothing Then netInfo.addressinfo = New List(Of NetworkInfoAddrMask)
                        netInfo.addressinfo.Add(addrinfo)
                    Next
                    For Each Ipa In nic.GetIPProperties.MulticastAddresses
                        If piDebuglevel > DebugLevel.dlEvents Then Log(String.Format("    The Multicast address of {0} is {1} with Address family {2}", nic.Description, Ipa.Address.ToString, Ipa.Address.AddressFamily.ToString), LogType.LOG_TYPE_INFO)
                        Dim addrinfo As New NetworkInfoAddrMask With {
                            .address = Ipa.Address.ToString,
                            .mask = Ipa.Address.AddressFamily.ToString,
                            .addrtype = "Multicast"
                        }
                        If netInfo.addressinfo Is Nothing Then netInfo.addressinfo = New List(Of NetworkInfoAddrMask)
                        netInfo.addressinfo.Add(addrinfo)
                    Next
                Else
                    ' operational state down
                End If
                PortInfo.Add(netInfo)
            Next
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetEthernetPorts found " & PortInfo.Count.ToString & " Ethernetports with IPv4 addresses assigned", LogType.LOG_TYPE_INFO)
            If PortInfo.Count > 0 Then
                Return PortInfo
            Else
                Return Nothing
            End If
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then Log("Error in GetEthernetPorts trying to get own MAC address with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Return Nothing
        End Try
    End Function

    Public Function IsEthernetPortAlive(ipAddress As String) As Boolean
        If piDebuglevel > DebugLevel.dlEvents Then Log("IsEthernetPortAlive called with Interface = " & ipAddress, LogType.LOG_TYPE_INFO)
        Dim Ethernetports As New Dictionary(Of String, String)
        Try
            For Each nic As System.Net.NetworkInformation.NetworkInterface In System.Net.NetworkInformation.NetworkInterface.GetAllNetworkInterfaces()
                If piDebuglevel > DebugLevel.dlEvents Then Log(String.Format("The MAC address of {0} is {1} with status {2}", nic.Description, nic.GetPhysicalAddress().ToString, nic.OperationalStatus.ToString), LogType.LOG_TYPE_INFO)
                If piDebuglevel > DebugLevel.dlEvents Then Log(String.Format("  The ID of {0} has name {1} and interface type {2}", nic.Id.ToString, nic.Name, nic.NetworkInterfaceType.ToString), LogType.LOG_TYPE_INFO)
                If nic.OperationalStatus = Net.NetworkInformation.OperationalStatus.Up Then
                    For Each Ipa In nic.GetIPProperties.UnicastAddresses
                        If piDebuglevel > DebugLevel.dlEvents Then Log(String.Format("    The UniCast address of {0} is {1} with mask {2} and Address family {3}", nic.Description, Ipa.Address.ToString, Ipa.IPv4Mask.ToString, Ipa.Address.AddressFamily.ToString), LogType.LOG_TYPE_INFO)
                        If Ipa.Address.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork Then
                            If Ipa.Address.ToString = ipAddress Then Return True
                        End If
                    Next
                End If
            Next
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then Log("Error in IsEthernetPortAlive with Interface = " & ipAddress & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Return False
        End Try
        Return False
    End Function

    Public Function IsPortAvailable(port As Integer) As Boolean
        Dim listener As TcpListener = Nothing
        Try
            listener = New TcpListener(System.Net.IPAddress.Loopback, port)
            listener.Start()
            Return True ' Port is available
        Catch ex As SocketException
            Return False ' Port is in use
        Finally
            If listener IsNot Nothing Then
                Try
                    listener.Stop()
                Catch
                    ' Ignore any errors when stopping
                End Try
            End If
        End Try
    End Function

    Public Function GetAvailablePorts(startPort As Integer, endPort As Integer) As List(Of Integer)
        Dim availablePorts As New List(Of Integer)()
        Try
            For port As Integer = startPort To endPort
                If IsPortAvailable(port) Then
                    availablePorts.Add(port)
                End If
            Next
        Catch ex As Exception
        End Try
        Return availablePorts
    End Function

    Public Function EncodeTags(ByVal InString As String) As String
        EncodeTags = InString
        Dim InIndex As Integer = 0
        Dim Outstring As String = ""
        InString = Trim(InString)
        If InString = "" Then Exit Function
        Try
            Do While InIndex < InString.Length
                If InString(InIndex) = ">" Then
                    Outstring += "&gt;"
                ElseIf InString(InIndex) = "<" Then
                    Outstring += "&lt;"
                ElseIf InString(InIndex) = "'" Then
                    Outstring += "&#39;"
                Else
                    Outstring &= InString(InIndex)
                End If
                InIndex += 1
            Loop
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then Log("Error in EncodeTags. URI = " & InString & " Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        EncodeTags = Outstring
    End Function

    Public Function EncodeTagsEx(ByVal InString As String) As String
        EncodeTagsEx = InString
        Dim InIndex As Integer = 0
        Dim Outstring As String = ""
        InString = Trim(InString)
        If InString = "" Then Exit Function
        Try
            Do While InIndex < InString.Length
                If InString(InIndex) = ">" Then
                    Outstring += "&gt;"
                ElseIf InString(InIndex) = "<" Then
                    Outstring += "&lt;"
                ElseIf InString(InIndex) = "'" Then
                    Outstring += "&#39;"
                ElseIf InString(InIndex) = """" Then
                    Outstring += "&#34;"
                Else
                    Outstring &= InString(InIndex)
                End If
                InIndex += 1
            Loop
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then Log("Error in EncodeTagsEx. URI = " & InString & " Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        EncodeTagsEx = Outstring
    End Function

    Public Function RemoveDoubleQuotes(ByVal InString As String) As String
        InString = Trim(InString)
        If InString = "" Then Return ""
        Return InString.Replace("""", "&#34;")
    End Function

    Public Function FindPairInJSONString(inString As String, inName As String) As Object
        Try
            If inString = "" Or inName = "" Then Return Nothing
            If piDebuglevel > DebugLevel.dlEvents Then Log("FindPairInJSONString called with inName = " & inName, LogType.LOG_TYPE_INFO)
            Dim json As New JavaScriptSerializer
            Dim JSONdataLevel1 As Object
            JSONdataLevel1 = json.DeserializeObject(inString)
            For Each Entry As Object In JSONdataLevel1
                If Entry.Key = inName Then
                    If TypeOf (Entry.value) Is String Then
                        Return Entry.value.ToString
                    ElseIf TypeOf (Entry.value) Is Boolean Then
                        Return Entry.value.ToString
                    ElseIf TypeOf (Entry.value) Is Integer Then
                        Return Entry.value.ToString
                    Else
                        'If piDebuglevel > DebugLevel.dlEvents Then Log("FindPairInJSONString called with inString = " & inString & " did not know type for key = " & inName & " in inString = " & inString, LogType.LOG_TYPE_INFO)
                        Return json.Serialize(Entry.value)
                    End If
                End If
            Next
            If piDebuglevel > DebugLevel.dlEvents Then Log("FindPairInJSONString called with inString = " & inString & " did not find the key = " & inName, LogType.LOG_TYPE_WARNING)
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in FindPairInJSONString processing response with error = " & ex.Message & " with inString = " & inString & " and inName = " & inName, LogType.LOG_TYPE_ERROR)
        End Try
        Return ""
    End Function

    Public Function JavaScriptStringDecode(Source As String) As String
        Dim Decoded As String = Source.Replace("\'", "'")
        Decoded = Decoded.Replace("\""", """")
        Decoded = Decoded.Replace("\/", "/")
        ' Decoded = Decoded.Replace("\t", "\t")
        'Decoded = Decoded.Replace("\n", "\n")
        Return Decoded
    End Function

    Public Function HexStringToByteArray(ByVal hex As [String]) As Byte()
        Dim NumberChars As Integer = hex.Length
        Dim bytes As Byte() = New Byte((NumberChars / 2) - 1) {}
        For i As Integer = 0 To NumberChars - 1 Step 2
            bytes(i / 2) = Convert.ToByte(hex.Substring(i, 2), 16)
        Next
        Return bytes
    End Function

    Public Function ByteArrayToHexString(InBytes As Byte()) As String
        Dim sHex As String = BitConverter.ToString(InBytes).Replace("-", "")
        Return sHex
    End Function

    Public Function DigitsStringToByteArray(ByVal Digits As [String]) As Byte()
        Dim NumberChars As Integer = Digits.Length
        Dim bytes As Byte() = New Byte(NumberChars - 1) {}
        For i As Integer = 0 To NumberChars - 1
            bytes(i) = Convert.ToByte(Digits.Substring(i, 1), 10)
        Next
        Return bytes
    End Function

    Public Function ConvertArrayToList(inArray As System.Array) As List(Of String)
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ConvertArrayToList called", LogType.LOG_TYPE_INFO)
        Dim outList As New List(Of String)
        If inArray Is Nothing AndAlso inArray.Length = 0 Then
            Return outList
        End If
        For Each inString As String In inArray
            outList.Add(inString)
        Next
        Return outList
    End Function

    Public Function ConvertSecondsToTimeFormat(ByVal Seconds As Integer) As String
        ConvertSecondsToTimeFormat = "00:00:00"
        If Seconds < 0 Then Exit Function
        Dim StartTime As Date = CDate("00:00:00")
        ConvertSecondsToTimeFormat = Format(DateAdd("s", CType(Seconds, Double), StartTime), "HH:mm:ss")
    End Function

    Public Function ConvertDictionaryToQueryString(inDictionary As Dictionary(Of String, String)) As String
        If inDictionary Is Nothing Then Return ""
        Dim returnString As String = ""
        For Each entry In inDictionary
            If returnString <> "" Then returnString &= "&"
            returnString &= entry.Key & "=" & entry.Value
        Next
        Return returnString
    End Function

    Public Function PrepareForQuery(ByVal inString As String) As String
        ' this function deals with ' in query names
        PrepareForQuery = inString.Replace("'", "''")
    End Function

    Public Function EncodeBlanksinURI(ByVal InString As String) As String
        Return InString.Replace(" ", "%20")
    End Function

    Public Function RetrieveErrorCodeFromHTTPResponse(inString As String) As Integer
        inString = inString.Trim(inString)
        If inString = "" Then Return -1
        Dim errorCodeString As String = ""
        For index = 0 To inString.Length - 1
            If inString(index) >= "0" And inString(index) <= "9" Then
                errorCodeString &= inString(index)
            Else
                Exit For
            End If
        Next
        If errorCodeString <> "" Then
            Return Val(errorCodeString)
        End If
        Return -1
    End Function

    Public Sub CheckFirewallStatus(appName As String, autoAdd As Boolean, ruleName As String)
        '   This VBScript file includes sample code that enumerates
        '   Windows Firewall rules with a matching grouping string 
        '   using the Microsoft Windows Firewall APIs.
        ' https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ics/enumerating-firewall-rules

        ' Option Explicit On

        Dim CurrentProfiles
        Dim InterfaceArray
        Dim LowerBound
        Dim UpperBound
        Dim iterate
        Dim rule

        ' Profile Type
        Const NET_FW_PROFILE2_DOMAIN = 1
        Const NET_FW_PROFILE2_PRIVATE = 2
        Const NET_FW_PROFILE2_PUBLIC = 4

        ' Protocol
        Const NET_FW_IP_PROTOCOL_TCP = 6
        Const NET_FW_IP_PROTOCOL_UDP = 17
        Const NET_FW_IP_PROTOCOL_ICMPv4 = 1
        Const NET_FW_IP_PROTOCOL_ICMPv6 = 58
        Const NET_FW_IP_PROTOCOL_ANY = 256

        ' Direction
        Const NET_FW_RULE_DIR_IN = 1
        Const NET_FW_RULE_DIR_OUT = 2

        ' Action
        Const NET_FW_ACTION_BLOCK = 0
        Const NET_FW_ACTION_ALLOW = 1


        ' Create the FwPolicy2 object.
        Dim fwPolicy2 = CreateObject("HNetCfg.FwPolicy2")

        CurrentProfiles = fwPolicy2.CurrentProfileTypes

        '// The returned 'CurrentProfiles' bitmask can have more than 1 bit set if multiple profiles 
        '//   are active or current at the same time
        Dim firewallActive As Boolean = False

        If (CurrentProfiles And NET_FW_PROFILE2_DOMAIN) Then
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("CheckFirewallStatus found Domain Firewall Profile is active", LogType.LOG_TYPE_INFO)
            firewallActive = True
        End If

        If (CurrentProfiles And NET_FW_PROFILE2_PRIVATE) Then
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("CheckFirewallStatus found Private Firewall Profile is active", LogType.LOG_TYPE_INFO)
            firewallActive = True
        End If

        If (CurrentProfiles And NET_FW_PROFILE2_PUBLIC) Then
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("CheckFirewallStatus found Public Firewall Profile is active", LogType.LOG_TYPE_INFO)
            firewallActive = True
        End If

        If Not firewallActive Then
            If piDebuglevel > DebugLevel.dlOff Then Log("CheckFirewallStatus no firewall (Windows Defender) active", LogType.LOG_TYPE_INFO)
            Exit Sub
        End If

        ' Get the Rules object
        Dim RulesObject = fwPolicy2.Rules
        Dim inRuleFound As Boolean = False
        Dim outRuleFound As Boolean = False
        Dim TCPRuleFound As Boolean = False
        Dim UDPRulefound As Boolean = False

        If piDebuglevel > DebugLevel.dlEvents Then Log("CheckFirewallStatus found Rules:", LogType.LOG_TYPE_INFO)
        For Each rule In RulesObject
            If piDebuglevel > DebugLevel.dlVerbose Then Log("CheckFirewallStatus found Rule Name: " & rule.Name, LogType.LOG_TYPE_INFO)
            If rule.ApplicationName IsNot Nothing AndAlso (rule.ApplicationName.ToString.ToUpper = appName.Replace("/", "\").ToUpper) Then
                '   If rule.Grouping = "@firewallapi.dll,-23255" Then
                If piDebuglevel > DebugLevel.dlOff Then Log("CheckFirewallStatus found Rule Name: " & rule.Name & ", Description = " & rule.Description & ", App Name = " & rule.ApplicationName & ", Service Name = " & rule.ServiceName & " and Enabled = " & rule.Enabled.ToString, LogType.LOG_TYPE_INFO)
                Select Case rule.Protocol
                    Case NET_FW_IP_PROTOCOL_TCP
                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("CheckFirewallStatus ---- IP Protocol: TCP.", LogType.LOG_TYPE_INFO)
                        TCPRuleFound = True
                    Case NET_FW_IP_PROTOCOL_UDP
                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("CheckFirewallStatus ---- IP Protocol: UDP.", LogType.LOG_TYPE_INFO)
                        UDPRulefound = True
                    Case NET_FW_IP_PROTOCOL_ANY
                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("CheckFirewallStatus ---- IP Protocol: ANY.", LogType.LOG_TYPE_INFO)
                        TCPRuleFound = True
                        UDPRulefound = True
                    Case NET_FW_IP_PROTOCOL_ICMPv4
                        If piDebuglevel > DebugLevel.dlEvents Then Log("CheckFirewallStatus ---- IP Protocol: ICMPv4.", LogType.LOG_TYPE_INFO)
                    Case NET_FW_IP_PROTOCOL_ICMPv6
                        If piDebuglevel > DebugLevel.dlEvents Then Log("CheckFirewallStatus ---- IP Protocol: ICMPv6.", LogType.LOG_TYPE_INFO)
                    Case Else
                        If piDebuglevel > DebugLevel.dlEvents Then Log("CheckFirewallStatus ---- IP Protocol: " & rule.Protocol, LogType.LOG_TYPE_INFO)
                End Select
                If rule.Protocol = NET_FW_IP_PROTOCOL_TCP Or rule.Protocol = NET_FW_IP_PROTOCOL_UDP Then
                    If piDebuglevel > DebugLevel.dlEvents Then Log("CheckFirewallStatus ---- Protocol: " & rule.Protocol.ToString & " Local Ports: " & rule.LocalPorts & ", Remote Ports = " & rule.RemotePorts & ", LocalAddresses = " & rule.LocalAddresses & ", RemoteAddresses = " & rule.RemoteAddresses, LogType.LOG_TYPE_INFO)
                End If
                If rule.Protocol = NET_FW_IP_PROTOCOL_ICMPv4 Or rule.Protocol = NET_FW_IP_PROTOCOL_ICMPv6 Then
                    If piDebuglevel > DebugLevel.dlEvents Then Log("CheckFirewallStatus -- ICMP Type and Code: " & rule.IcmpTypesAndCodes, LogType.LOG_TYPE_INFO)
                End If
                Select Case rule.Direction
                    Case NET_FW_RULE_DIR_IN
                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("CheckFirewallStatus ---- Direction: In", LogType.LOG_TYPE_INFO)
                        If rule.Enabled Then inRuleFound = True
                    Case NET_FW_RULE_DIR_OUT
                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("CheckFirewallStatus ---- Direction: Out", LogType.LOG_TYPE_INFO)
                        If rule.Enabled Then outRuleFound = True
                End Select
                If piDebuglevel > DebugLevel.dlEvents Then Log("CheckFirewallStatus ----  Enabled: " & rule.Enabled, LogType.LOG_TYPE_INFO)
                If piDebuglevel > DebugLevel.dlEvents Then Log("CheckFirewallStatus ----  Edge: " & rule.EdgeTraversal, LogType.LOG_TYPE_INFO)
                If piDebuglevel > DebugLevel.dlEvents Then Log("CheckFirewallStatus ----  LocalPorts: " & rule.LocalPorts, LogType.LOG_TYPE_INFO)
                If piDebuglevel > DebugLevel.dlEvents Then Log("CheckFirewallStatus ----  Profiles: " & rule.Profiles, LogType.LOG_TYPE_INFO)

                Select Case rule.Action
                    Case NET_FW_ACTION_ALLOW
                        If piDebuglevel > DebugLevel.dlOff Then Log("CheckFirewallStatus ---- Action: Allow", LogType.LOG_TYPE_INFO, LogColorGreen, LogType.LOG_TYPE_INFO)
                    Case NET_FW_ACTION_BLOCK
                        If piDebuglevel > DebugLevel.dlOff Then Log("CheckFirewallStatus ---- Action: Block", LogType.LOG_TYPE_WARNING)
                End Select
                If piDebuglevel > DebugLevel.dlEvents Then Log("CheckFirewallStatus ---- Grouping: " & rule.Grouping & ", Interface Types: " & rule.InterfaceTypes, LogType.LOG_TYPE_INFO)
                InterfaceArray = rule.Interfaces
                If InterfaceArray Is Nothing Then
                    If piDebuglevel > DebugLevel.dlEvents Then Log("CheckFirewallStatus ---- There are no excluded interfaces", LogType.LOG_TYPE_INFO)
                Else
                    LowerBound = LBound(InterfaceArray)
                    UpperBound = UBound(InterfaceArray)
                    If piDebuglevel > DebugLevel.dlEvents Then Log("CheckFirewallStatus ---- Excluded interfaces: ", LogType.LOG_TYPE_INFO)
                    For iterate = LowerBound To UpperBound
                        If piDebuglevel > DebugLevel.dlEvents Then Log("CheckFirewallStatus ------ " & InterfaceArray(iterate), LogType.LOG_TYPE_INFO)
                    Next
                End If
            End If
        Next
        If Not (inRuleFound And UDPRulefound And TCPRuleFound) Then
            Dim success As Boolean = False
            If autoAdd Then
                success = AddApplicationFirewallRule(ruleName, appName.Replace("/", "\"))
                If success Then Log("CheckFirewallStatus added firewall rules for Application = " & appName.Replace("/", "\"), LogType.LOG_TYPE_WARNING)
            End If
            If Not success And piDebuglevel > DebugLevel.dlOff Then Log("CheckFirewallStatus did not find proper firewall rules. It was looking for Application = " & appName.Replace("/", "\") & ". Fix it as plugin will not function properly", LogType.LOG_TYPE_ERROR)
        End If
    End Sub

    Public Function AddApplicationFirewallRule(ruleName As String, AppName As String) As Boolean

        Try
            Dim CurrentProfiles
            ' Protocol
            Const NET_FW_IP_PROTOCOL_TCP = 6
            Const NET_FW_IP_PROTOCOL_UDP = 17
            'Const NET_FW_IP_PROTOCOL_ANY = 256
            Const NET_FW_RULE_DIR_IN = 1
            'Const NET_FW_RULE_DIR_OUT = 2
            ' Policy
            Const NET_FW_PROFILE2_DOMAIN = &H1
            Const NET_FW_PROFILE2_PRIVATE = &H2
            Const NET_FW_PROFILE2_PUBLIC = &H4
            Const NET_FW_PROFILE2_ALL = &H7FFFFFFF

            'Action
            Const NET_FW_ACTION_ALLOW = 1
            ' Create the FwPolicy2 object.
            Dim fwPolicy2
            fwPolicy2 = CreateObject("HNetCfg.FwPolicy2")
            ' Get the Rules object
            Dim RulesObject
            RulesObject = fwPolicy2.Rules
            CurrentProfiles = fwPolicy2.CurrentProfileTypes
            'Create a Rule Object for TCP
            Dim NewRule
            NewRule = CreateObject("HNetCfg.FWRule")
            NewRule.Name = ruleName
            NewRule.Description = "Allow HomeSeer pluging acces"
            NewRule.Applicationname = AppName
            NewRule.Protocol = NET_FW_IP_PROTOCOL_TCP
            NewRule.Direction = NET_FW_RULE_DIR_IN
            ' NewRule.LocalPorts = "*"
            NewRule.Enabled = True
            NewRule.Grouping = "@firewallapi.dll,-23255"
            NewRule.Profiles = NET_FW_PROFILE2_ALL 'CurrentProfiles
            NewRule.Action = NET_FW_ACTION_ALLOW
            'Add a new rule for UDP
            RulesObject.Add(NewRule)
            NewRule = CreateObject("HNetCfg.FWRule")
            NewRule.Name = ruleName
            NewRule.Description = "Allow HomeSeer pluging acces"
            NewRule.Applicationname = AppName
            NewRule.Protocol = NET_FW_IP_PROTOCOL_UDP
            NewRule.Direction = NET_FW_RULE_DIR_IN
            'NewRule.LocalPorts = "*"
            NewRule.Enabled = True
            NewRule.Grouping = "@firewallapi.dll,-23255"
            NewRule.Profiles = NET_FW_PROFILE2_ALL ' CurrentProfiles
            NewRule.Action = NET_FW_ACTION_ALLOW
            'Add a new rule
            RulesObject.Add(NewRule)
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then Log("Error in AddApplicationFirewallRule for Application = " & AppName & " and RuleName = " & ruleName & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Return False
        End Try
        Return True
    End Function

    Sub ExecuteNetshCommand(arguments As String)
        Try
            Dim proc As New Process()
            proc.StartInfo.FileName = "netsh"
            proc.StartInfo.Arguments = arguments
            proc.StartInfo.UseShellExecute = False
            proc.StartInfo.RedirectStandardOutput = True
            proc.StartInfo.CreateNoWindow = True
            proc.Start()
            Dim output As String = proc.StandardOutput.ReadToEnd()
            proc.WaitForExit()
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log($"ExecuteNetshCommand with arguments = {arguments} returned result = {output}", LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then Log($"Error in ExecuteNetshCommand with arguments = {arguments} and Error = {ex.Message}", LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Sub AddFirewallRule(ruleName As String, port As Integer)
        Try
            Dim command As String = $"advfirewall firewall add rule name=""{ruleName}"" dir=in action=allow protocol=TCP localport={port}"
            ExecuteNetshCommand(command)
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then Log($"Error in AddFirewallRule for ruleName = {ruleName}, port = {port} and Error = {ex.Message}", LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Sub RemoveFirewallRule(ruleName As String)
        Try
            Dim command As String = $"advfirewall firewall delete rule name=""{ruleName}"""
            ExecuteNetshCommand(command)
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then Log($"Error in RemoveFirewallRule for ruleName = {ruleName} and Error = {ex.Message}", LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Function GetOperatingSystem() As String
        If RuntimeInformation.IsOSPlatform(OSPlatform.Windows) Then
            Return "windows"
        ElseIf RuntimeInformation.IsOSPlatform(OSPlatform.OSX) Then
            Return "mac"
        ElseIf RuntimeInformation.IsOSPlatform(OSPlatform.Linux) Then
            Dim uname = RunShellCommand("uname", Nothing)
            If uname.ToLower().Contains("bsd") Then
                Return "bsd"
            Else
                Return "linux"
            End If
        Else
            Return "unknown"
        End If
    End Function

    Function GetArchitecture(os As String) As String
        Dim arch = RuntimeInformation.OSArchitecture.ToString().ToLower()
        Select Case arch
            Case "x64"
                If os = "windows" Then Return "win64" Else Return "amd64"
            Case "x86"
                If os = "windows" Then Return "win32" Else Return "386i"
            Case "arm64"
                If os = "windows" Then Return "winarm64" Else Return "arm64"
            Case "arm"
                Dim cpuinfo = RunShellCommand("uname -m", Nothing)
                If cpuinfo.Contains("v6") Then Return "armv6"
                If cpuinfo.Contains("v7") Then Return "armv7"
                If cpuinfo.Contains("v8") Then Return "armv8"
                Return "arm"
            Case "mipsel"
                Return "mipsel"
            Case Else
                Return arch
        End Select
    End Function

    Function ExecutableChecker(exeName As String, requiredVersion As String, path As String) As Boolean
        Dim isInstalled As Boolean = False
        Dim installedVersion As String = ""

        If RuntimeInformation.IsOSPlatform(OSPlatform.Windows) Then
            Dim exePath As String = $"{path}{exeName}.exe" ' Adjust path as needed

            If File.Exists(exePath) Then
                Dim versionInfo = FileVersionInfo.GetVersionInfo(exePath)
                If versionInfo.FileVersion IsNot Nothing Then installedVersion = versionInfo.FileVersion
                isInstalled = True
            End If

        ElseIf RuntimeInformation.IsOSPlatform(OSPlatform.Linux) OrElse RuntimeInformation.IsOSPlatform(OSPlatform.OSX) Then
            Dim versionOutput = RunShellCommand($"{exeName} -version", Nothing)

            If Not String.IsNullOrEmpty(versionOutput) Then
                ' Try to extract version from output (customize for your target executable)
                Dim match = System.Text.RegularExpressions.Regex.Match(versionOutput, "\d+(\.\d+)+")
                If match.Success Then
                    installedVersion = match.Value
                    isInstalled = True
                End If
            End If
        End If

        If isInstalled Then
            If installedVersion.StartsWith(requiredVersion) Then
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log($"ExecutableChecker found {exeName} installed with required Version: {installedVersion}", LogType.LOG_TYPE_INFO)
                Return True
            Else
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log($"ExecutableChecker found {exeName} installed with versionInstalled: {installedVersion} but needs {requiredVersion }", LogType.LOG_TYPE_INFO)
            End If
        Else
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log($"ExecutableChecker did find {exeName} installed", LogType.LOG_TYPE_INFO)
        End If
        Return False
    End Function

    Function GetDownloadPath() As String
        If RuntimeInformation.IsOSPlatform(OSPlatform.Windows) Then
            Return Path.GetTempPath()
        Else
            Dim home As String = Environment.GetEnvironmentVariable("HOME")
            Return Path.Combine(home, "Downloads")
        End If
    End Function

    Sub RunProcess(filePath As String, maxWait As Integer)
        Dim psi As New ProcessStartInfo() With {
            .FileName = filePath,
            .UseShellExecute = False,
            .RedirectStandardOutput = True,
            .RedirectStandardError = True
        }
        Using proc As New Process()
            Process.Start(psi)
            ' Optional timeout (x seconds)
            If Not proc.WaitForExit(maxWait * 1000) Then
                proc.Kill()
                Exit Sub
            End If
        End Using
    End Sub

    Public Function SerializeToXmlString(nodeList As XmlNodeList) As String
        Try
            Dim result As New StringBuilder()
            For Each node As XmlNode In nodeList
                result.Append(node.OuterXml)
            Next
            Return result.ToString()
        Catch ex As Exception
        End Try
        Return ""
    End Function

    Function GetMacFromIP(ipAddress As String, ByRef macAddress As String) As String
        macAddress = ""
        Try
            ' Step 1: Ping the IP to populate ARP cache
            Dim pingSuccess = My.Computer.Network.Ping(ipAddress, 1000)

            If Not pingSuccess Then
                Return "IP unreachable or not responding"
            End If

            ' Step 2: Determine platform
            Dim isWindows = RuntimeInformation.IsOSPlatform(OSPlatform.Windows)

            ' Step 3: First try ARP
            Dim arpResult = RunCommand("arp", If(isWindows, "-a " & ipAddress, "-n " & ipAddress))
            If Not String.IsNullOrEmpty(arpResult) Then
                Dim mac = ParseMacFromOutput(arpResult, ipAddress, isWindows)
                If mac IsNot Nothing Then
                    macAddress = mac
                    Return ""
                End If
            End If

            ' Step 4: On non-Windows, try ip neighbor
            If Not isWindows Then
                Dim ipResult = RunCommand("ip", "neighbor show " & ipAddress)
                If Not String.IsNullOrEmpty(ipResult) Then
                    Dim parts = ipResult.Split({" "c}, StringSplitOptions.RemoveEmptyEntries)
                    Dim lladdrIndex = Array.IndexOf(parts, "lladdr")
                    If lladdrIndex <> -1 AndAlso lladdrIndex + 1 < parts.Length Then
                        macAddress = parts(lladdrIndex + 1)
                        Return ""
                    End If
                End If
            End If

        Catch ex As Exception
            Return "Error: " & ex.Message
        End Try

        Return "MAC not found"
    End Function

    Private Function RunCommand(command As String, args As String) As String
        Try
            Dim p As New Process()
            p.StartInfo.FileName = command
            p.StartInfo.Arguments = args
            p.StartInfo.RedirectStandardOutput = True
            p.StartInfo.UseShellExecute = False
            p.StartInfo.CreateNoWindow = True
            p.Start()
            Dim output As String = p.StandardOutput.ReadToEnd()
            p.WaitForExit()
            Return output
        Catch
            Return Nothing
        End Try
    End Function

    Private Function ParseMacFromOutput(output As String, ipAddress As String, isWindows As Boolean) As String
        Dim lines = output.Split({vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)
        For Each line In lines
            If line.Contains(ipAddress) Then
                Dim parts = line.Split({" ", vbTab}, StringSplitOptions.RemoveEmptyEntries)
                If isWindows Then
                    ' Windows format: Interface Address Physical Type
                    If parts.Length >= 3 Then Return parts(1)
                Else
                    ' Linux/macOS format: IP HW-addr Flags Mask Device
                    If parts.Length >= 3 Then Return parts(2)
                End If
            End If
        Next
        Return Nothing
    End Function

    Function GetIPFromMac(macAddress As String, ByRef ipAddress As String) As String
        ipAddress = ""
        Try
            Dim isWindows = RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
            Dim output As String = Nothing

            ' Try ARP on Windows or Linux
            output = RunCommand("arp", If(isWindows, "-a", "-n"))

            If Not String.IsNullOrEmpty(output) Then
                Dim ip = ParseIPFromArpOutput(output, macAddress, isWindows)
                If ip IsNot Nothing Then
                    ipAddress = ip
                    Return ""
                End If
            End If

            ' On Linux, try 'ip neighbor show'
            If Not isWindows Then
                output = RunCommand("ip", "neighbor show")
                If Not String.IsNullOrEmpty(output) Then
                    Dim lines = output.Split({vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)
                    For Each line In lines
                        If line.IndexOf(macAddress, StringComparison.OrdinalIgnoreCase) >= 0 Then
                            Dim parts = line.Split({" "c}, StringSplitOptions.RemoveEmptyEntries)
                            If parts.Length >= 1 Then
                                ipAddress = parts(0) ' IP address is first
                                Return ""
                            End If
                        End If
                    Next
                End If
            End If

        Catch ex As Exception
            Return "Error: " & ex.Message
        End Try

        Return "IP not found"
    End Function

    Private Function ParseIPFromArpOutput(output As String, macAddress As String, isWindows As Boolean) As String
        Dim lines = output.Split({vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)
        For Each line In lines
            If line.IndexOf(macAddress, StringComparison.OrdinalIgnoreCase) >= 0 Then
                Dim parts = line.Split({" ", vbTab}, StringSplitOptions.RemoveEmptyEntries)
                If isWindows Then
                    ' Format: Interface Address Physical Type
                    If parts.Length >= 2 Then Return parts(0)
                Else
                    ' Format: IP HW-addr Flags Mask Device
                    If parts.Length >= 1 Then Return parts(0)
                End If
            End If
        Next
        Return Nothing
    End Function

    Public Function GitHubCompressedDownloader(githubUrl As String, downloadFileName As String, extractFilenames As String(), downloadPath As String) As Boolean
        Try
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log($"GitHubCompressedDownloader is downloading from: {githubUrl} to: {downloadFileName} with targetNames: { String.Join(", ", extractFilenames)}", LogType.LOG_TYPE_INFO)

            Dim fullDownloadFilename As String = Path.Combine(downloadPath, downloadFileName)

            ' Delete existing ZIP if it exists
            Try
                If File.Exists(fullDownloadFilename) Then
                    File.Delete(fullDownloadFilename)
                End If
            Catch ex As Exception
                If piDebuglevel > DebugLevel.dlOff Then Log($"GitHubCompressedDownloader has error deleting existing download file = {fullDownloadFilename} with error = {ex.Message}", LogType.LOG_TYPE_ERROR)
            End Try

            ' Download ZIP
            Try
                Using client As New HttpClient()
                    client.DefaultRequestHeaders.UserAgent.ParseAdd("StreamCamPlugin/1.0")

                    Dim response = client.GetAsync(githubUrl).GetAwaiter().GetResult()

                    If Not response.IsSuccessStatusCode Then
                        If piDebuglevel > DebugLevel.dlOff Then Log($"GitHubCompressedDownloader has error downloading file: HTTP {response.StatusCode}", LogType.LOG_TYPE_ERROR)
                        Return False
                    End If

                    Dim contentType = response.Content.Headers.ContentType?.MediaType
                    If contentType IsNot Nothing AndAlso contentType.Contains("html") Then
                        If piDebuglevel > DebugLevel.dlOff Then Log($"GitHubCompressedDownloader downloaded content is HTML, likely an error or redirect page. Aborting.", LogType.LOG_TYPE_ERROR)
                        Return False
                    End If

                    Dim zipBytes = response.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult()
                    File.WriteAllBytes(fullDownloadFilename, zipBytes)
                End Using
            Catch ex As Exception
                If piDebuglevel > DebugLevel.dlOff Then Log($"GitHubCompressedDownloader has error downloading a new file with error = {ex.Message}", LogType.LOG_TYPE_ERROR)
                Return False
            End Try

            ' dcor for testing if I want to check hash for tampering
            'Dim expectedHash As String = "3259a04c402cd22b39092119151140d91b7d9898591c8d94422c6f84eddf7380"
            'Dim actualHash As String = ComputeSha256(fullDownloadFilename)

            'If String.Equals(actualHash, expectedHash, StringComparison.OrdinalIgnoreCase) Then
            'Log("ZIP hash verified successfully.", LogType.LOG_TYPE_INFO)
            'Else
            'Log($"ZIP hash mismatch! Expected: {expectedHash}, Got: {actualHash}", LogType.LOG_TYPE_ERROR)
            'Return False
            'End If

            ' Extract based on platform
            If RuntimeInformation.IsOSPlatform(OSPlatform.Windows) OrElse
                (RuntimeInformation.IsOSPlatform(OSPlatform.OSX) AndAlso fullDownloadFilename.EndsWith(".zip", StringComparison.OrdinalIgnoreCase)) Then

                For Each extractFilename In extractFilenames
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log($"GitHubCompressedDownloader is deleting file = {Path.Combine(downloadPath, extractFilename)}...", LogType.LOG_TYPE_INFO)
                    Try
                        If File.Exists(Path.Combine(downloadPath, extractFilename)) Then
                            File.Delete(Path.Combine(downloadPath, extractFilename))
                        End If
                    Catch ex As Exception
                        If piDebuglevel > DebugLevel.dlOff Then Log($"GitHubCompressedDownloader has error deleting existing target file = {Path.Combine(downloadPath, extractFilename)} with error = {ex.Message}", LogType.LOG_TYPE_ERROR)
                    End Try
                Next

                ZipFile.ExtractToDirectory(fullDownloadFilename, downloadPath)

            ElseIf RuntimeInformation.IsOSPlatform(OSPlatform.Linux) OrElse RuntimeInformation.IsOSPlatform(OSPlatform.OSX) Then
                If fullDownloadFilename.EndsWith(".tar.gz", StringComparison.OrdinalIgnoreCase) Then
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log($"GitHubCompressedDownloader is extracting  {fullDownloadFilename} using system 'tar' to {downloadPath}", LogType.LOG_TYPE_INFO)
                    Dim result As String = RunShellCommand($"tar -xvzf ""{fullDownloadFilename}"" -C ""{downloadPath}""", Nothing)
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log($"GitHubCompressedDownloader extracted {fullDownloadFilename} with result = {result}", LogType.LOG_TYPE_INFO)
                Else
                    ChangeLinuxFileAccessRights("777", fullDownloadFilename)
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("GitHubCompressedDownloader updated exec rights on Unix.", LogType.LOG_TYPE_INFO)
                End If
            Else
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("GitHubCompressedDownloader unsupported operating system.", LogType.LOG_TYPE_WARNING)
                Return False
            End If

            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("GitHubCompressedDownloader download & decompression complete.", LogType.LOG_TYPE_INFO)
            Return True

        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then Log($"Error in GitHubCompressedDownloader with error = {ex.Message}", LogType.LOG_TYPE_ERROR)
            Return False
        End Try
        Return True
    End Function

    Public Function ComputeSha256(filePath As String) As String
        ' computes a hash for the downloaded (zip) file
        Using sha As SHA256 = SHA256.Create()
            Using stream = File.OpenRead(filePath)
                Dim hashBytes = sha.ComputeHash(stream)
                Return BitConverter.ToString(hashBytes).Replace("-", "").ToLowerInvariant()
            End Using
        End Using
    End Function

    Public NotInheritable Class Simple3Des
        Private TripleDes As New TripleDESCryptoServiceProvider

        Sub New(ByVal key As String)
            ' Initialize the crypto provider.
            TripleDes.Key = TruncateHash(key, TripleDes.KeySize \ 8)
            TripleDes.IV = TruncateHash("", TripleDes.BlockSize \ 8)
        End Sub

        Private Function TruncateHash(ByVal key As String, ByVal length As Integer) As Byte()
            Try
                Dim sha1 As New SHA1CryptoServiceProvider
                ' Hash the key.
                Dim keyBytes() As Byte =
                    System.Text.Encoding.Unicode.GetBytes(key)
                Dim hash() As Byte = sha1.ComputeHash(keyBytes)

                ' Truncate or pad the hash.
                ReDim Preserve hash(length - 1)
                Return hash
            Catch ex As Exception
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in TruncateHash with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Return Nothing
        End Function

        Public Function EncryptData(ByVal plaintext As String) As String
            ' Convert the plaintext string to a byte array.
            If plaintext = "" Then
                If piDebuglevel > DebugLevel.dlEvents Then Log("EncryptData , no data", LogType.LOG_TYPE_INFO)
                Return ""
            End If
            Try
                Dim plaintextBytes() As Byte = System.Text.Encoding.Unicode.GetBytes(plaintext)
                ' Create the stream.
                Dim ms As New System.IO.MemoryStream
                ' Create the encoder to write to the stream.
                Dim encStream As New CryptoStream(ms,
                    TripleDes.CreateEncryptor(),
                    System.Security.Cryptography.CryptoStreamMode.Write)
                ' Use the crypto stream to write the byte array to the stream.
                encStream.Write(plaintextBytes, 0, plaintextBytes.Length)
                encStream.FlushFinalBlock()
                ' Convert the encrypted stream to a printable string.
                Return Convert.ToBase64String(ms.ToArray)
            Catch ex As Exception
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in EncryptData with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Return ""
        End Function

        Public Function DecryptData(ByVal encryptedtext As String) As String
            ' Convert the encrypted text string to a byte array.
            If encryptedtext = "" Then
                If piDebuglevel > DebugLevel.dlEvents Then Log("DecryptData called without data", LogType.LOG_TYPE_INFO)
                Return ""
            End If
            Try
                Dim encryptedBytes() As Byte = Convert.FromBase64String(encryptedtext)
                ' Create the stream.
                Dim ms As New System.IO.MemoryStream
                ' Create the decoder to write to the stream.
                Dim decStream As New CryptoStream(ms,
                    TripleDes.CreateDecryptor(),
                    System.Security.Cryptography.CryptoStreamMode.Write)
                ' Use the crypto stream to write the byte array to the stream.
                decStream.Write(encryptedBytes, 0, encryptedBytes.Length)
                decStream.FlushFinalBlock()
                ' Convert the plaintext stream to a string.
                Return System.Text.Encoding.Unicode.GetString(ms.ToArray)
            Catch ex As Exception
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log($"Error in DecryptData with txt = {encryptedtext} and Error = {ex.Message}", LogType.LOG_TYPE_ERROR)
            End Try
            Return ""
        End Function
    End Class



#Region "Linux Functions"
    Function RunShellCommand(command As String, Optional dummy As Object = Nothing) As String
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log($"RunShellCommand called with cmd = {command}", LogType.LOG_TYPE_INFO)

        Try
            Dim psi As New ProcessStartInfo("/bin/bash", $"-c ""{command}""") With {
            .RedirectStandardOutput = True,
            .RedirectStandardError = True,
            .UseShellExecute = False,
            .CreateNoWindow = True
        }

            Dim output As New System.Text.StringBuilder()
            Dim errorOutput As New System.Text.StringBuilder()

            Using proc As New Process()
                proc.StartInfo = psi
                AddHandler proc.OutputDataReceived, Sub(sender, e)
                                                        If e.Data IsNot Nothing Then output.AppendLine(e.Data)
                                                    End Sub
                AddHandler proc.ErrorDataReceived, Sub(sender, e)
                                                       If e.Data IsNot Nothing Then errorOutput.AppendLine(e.Data)
                                                   End Sub

                proc.Start()
                proc.BeginOutputReadLine()
                proc.BeginErrorReadLine()
                proc.WaitForExit()

                Dim result As String = $"[STDOUT]{Environment.NewLine}{output.ToString().Trim()}" &
                   $"{Environment.NewLine}[STDERR]{Environment.NewLine}{errorOutput.ToString().Trim()}" &
                   $"{Environment.NewLine}[ExitCode] {proc.ExitCode}"
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log($"RunShellCommand called with cmd = {command} and result = {result}", LogType.LOG_TYPE_INFO)

                Return output.ToString.Trim()

            End Using

        Catch ex As Exception
            Return $"[EXCEPTION] {ex.Message}"
        End Try
    End Function

    Public Function IsNodeInstalled() As String
        Try
            Dim startInfo As New ProcessStartInfo() With {
            .FileName = "/bin/bash",
            .Arguments = "-c ""node -v""",
            .UseShellExecute = False,
            .RedirectStandardOutput = True,
            .RedirectStandardError = True,
            .CreateNoWindow = True
        }

            Using proc As New Process()
                proc.StartInfo = startInfo
                proc.Start()

                Dim output As String = proc.StandardOutput.ReadToEnd().Trim()
                Dim errorOutput As String = proc.StandardError.ReadToEnd().Trim()

                ' Timeout safeguard
                If Not proc.WaitForExit(5000) Then
                    proc.Kill()
                    If piDebuglevel > DebugLevel.dlOff Then Log("IsNodeInstalled timed out", LogType.LOG_TYPE_WARNING)
                    Return ""
                End If

                If piDebuglevel > DebugLevel.dlEvents Then
                    Log($"IsNodeInstalled stdout = {output}", LogType.LOG_TYPE_INFO)
                    If Not String.IsNullOrWhiteSpace(errorOutput) Then
                        Log($"IsNodeInstalled stderr = {errorOutput}", LogType.LOG_TYPE_WARNING)
                    End If
                End If

                If output.StartsWith("v") AndAlso output.Length > 1 Then
                    Return output.Substring(1) ' Return version number without 'v'
                End If
            End Using

        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then
                Log("Error in IsNodeInstalled: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End If
        End Try

        Return ""
    End Function

    Public Function IsArmArchitecture() As Boolean
        Try
            Dim startInfo As New ProcessStartInfo() With {
            .FileName = "uname",
            .Arguments = "-m",
            .UseShellExecute = False,
            .RedirectStandardOutput = True,
            .RedirectStandardError = True,
            .CreateNoWindow = True
        }

            Using proc As New Process()
                proc.StartInfo = startInfo
                proc.Start()

                Dim output As String = proc.StandardOutput.ReadToEnd()
                Dim errorOutput As String = proc.StandardError.ReadToEnd()

                ' Timeout safeguard
                If Not proc.WaitForExit(5000) Then
                    proc.Kill()
                    If piDebuglevel > DebugLevel.dlOff Then Log("IsArmArchitecture timeout calling uname", LogType.LOG_TYPE_ERROR)
                    Return False
                End If

                output = output.Trim().ToUpperInvariant()

                If piDebuglevel > DebugLevel.dlErrorsOnly Then
                    Log("IsArmArchitecture received uname = " & output, LogType.LOG_TYPE_INFO)
                    If Not String.IsNullOrWhiteSpace(errorOutput) Then
                        Log("IsArmArchitecture stderr = " & errorOutput.Trim(), LogType.LOG_TYPE_WARNING)
                    End If
                End If

                If output.Contains("ARM") OrElse output.Contains("AARCH") Then
                    Return True
                End If

            End Using

        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then
                Log("Error in IsArmArchitecture: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End If
            Return False
        End Try

        Return False
    End Function

    Public Sub ChangeLinuxFileAccessRights(newAccessrights As String, fileName As String)
        ' first set exe priviliges Execute the chmod command
        Try
            Dim command As String = "sudo chmod " & newAccessrights & " " & fileName
            Dim SetRightsProcess = New Process()
            SetRightsProcess.StartInfo.FileName = "/bin/bash"
            SetRightsProcess.StartInfo.Arguments = " -c """ & command & """"
            SetRightsProcess.StartInfo.RedirectStandardOutput = True
            SetRightsProcess.StartInfo.RedirectStandardError = True
            SetRightsProcess.StartInfo.UseShellExecute = False
            SetRightsProcess.Start()
            Dim output As String = SetRightsProcess.StandardOutput.ReadToEnd()
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ChangeLinuxFileAccessRights set accessrights = " & newAccessrights & " for file = " & fileName & " with result = " & output, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then Log("Error in ChangeLinuxFileAccessRights setting accessrights = " & newAccessrights & " for file = " & fileName & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Function RunBashCmd(cmd As String) As String
        Try
            Dim startInfo As New ProcessStartInfo()
            startInfo.FileName = "/bin/bash"
            startInfo.Arguments = cmd '"-c """ & cmd & """"
            startInfo.UseShellExecute = False
            startInfo.RedirectStandardOutput = True
            startInfo.RedirectStandardError = True
            Dim Process = New Process()
            Process.StartInfo = startInfo
            Process.Start()
            Dim output As String = Process.StandardOutput.ReadToEnd()
            If piDebuglevel > DebugLevel.dlEvents Then Log("RunBashCmd received response = " & output, LogType.LOG_TYPE_INFO)
            Return output
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then Log("Error in RunBashCmd with cmd = " & cmd & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Return ""
        End Try
        Return ""
    End Function

    Public Function SendLinuxCmd(cmd As String, argument As String) As String
        Try
            Dim startInfo As New ProcessStartInfo() With {
            .FileName = cmd,
            .Arguments = argument,
            .UseShellExecute = False,
            .RedirectStandardOutput = True,
            .RedirectStandardError = True,
            .CreateNoWindow = True
        }

            Using proc As New Process()
                proc.StartInfo = startInfo
                proc.Start()

                Dim output As String = proc.StandardOutput.ReadToEnd().Trim()
                Dim errorOutput As String = proc.StandardError.ReadToEnd().Trim()

                If Not proc.WaitForExit(5000) Then
                    proc.Kill()
                    If piDebuglevel > DebugLevel.dlOff Then Log($"SendLinuxCmd timed out for: {cmd} {argument}", LogType.LOG_TYPE_WARNING)
                    Return ""
                End If

                If piDebuglevel > DebugLevel.dlEvents Then
                    Log($"SendLinuxCmd response: {output}", LogType.LOG_TYPE_INFO)
                    If Not String.IsNullOrWhiteSpace(errorOutput) Then
                        Log($"SendLinuxCmd stderr: {errorOutput}", LogType.LOG_TYPE_WARNING)
                    End If
                End If

                Return output
            End Using

        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then
                Log($"Error in SendLinuxCmd: cmd = {cmd}, args = {argument}, err = {ex.Message}", LogType.LOG_TYPE_ERROR)
            End If
            Return ""
        End Try
    End Function

    Public Function StartLinuxProcess(Command As String, arguments As String) As String
        Try
            Dim psi As New ProcessStartInfo() With {
                .FileName = Command,
                .Arguments = arguments,
                .RedirectStandardOutput = True,
                .RedirectStandardError = True,
                .UseShellExecute = False,
                .CreateNoWindow = True
            }
            Using process As Process = Process.Start(psi)
                Dim output As String = process.StandardOutput.ReadToEnd()
                Dim errors As String = process.StandardError.ReadToEnd()
                process.WaitForExit()
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log($"StartLinuxProcess received response = {output}", LogType.LOG_TYPE_INFO)
                If Not String.IsNullOrWhiteSpace(errors) Then
                    Console.WriteLine("Errors
                " & vbCrLf & errors)
                    If piDebuglevel > DebugLevel.dlOff Then Log($"Warning in StartLinuxProcess received and Error response = {errors}", LogType.LOG_TYPE_WARNING)
                    Return errors
                End If
                Return output
            End Using
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlOff Then Log($"Error in StartLinuxProcess with command = {Command}, arguments = {arguments} and Error = {ex.Message }", LogType.LOG_TYPE_ERROR)
            Return ex.Message
        End Try

    End Function

    Public Function FindLinuxProgramIsRunning(programName As String) As String
        Dim nodeProcesses As String = SendLinuxCmd("ps", " -ef")
        If nodeProcesses = "" Then Return ""
        Dim nodeProcessStrings As String() = nodeProcesses.Split(vbLf)
        If nodeProcessStrings IsNot Nothing Then
            If piDebuglevel > DebugLevel.dlEvents Then Log("FindLinuxProgramIsRunning found " & nodeProcessStrings.Count.ToString & " processes", LogType.LOG_TYPE_INFO)
            If nodeProcesses.Count < 2 Then Return ""
            ' the first line will be the header use it to find the location of the collumns
            Dim locationPID As Integer = nodeProcessStrings(0).IndexOf("PID")
            Dim locationCMD As Integer = nodeProcessStrings(0).IndexOf("CMD")
            If piDebuglevel > DebugLevel.dlEvents Then Log("FindLinuxProgramIsRunning found PID offset = " & locationPID.ToString & " and CMD offset = " & locationCMD.ToString & " in Entry = " & nodeProcessStrings(0), LogType.LOG_TYPE_INFO)
            For Each entry As String In nodeProcessStrings
                If entry.ToUpper.IndexOf(programName.ToUpper) <> -1 Then
                    If piDebuglevel > DebugLevel.dlEvents Then Log("FindLinuxProgramIsRunning found process with name = " & programName & " and info = " & entry, LogType.LOG_TYPE_INFO)
                    Dim index As Integer = locationPID + 2  'the column is lined up with the right of the word PID
                    Dim entryParts As String() = entry.Split(CType(" ", Char()), StringSplitOptions.RemoveEmptyEntries)
                    If entryParts.Count > 1 Then
                        If piDebuglevel > DebugLevel.dlEvents Then Log("FindLinuxProgramIsRunning found process with name = " & programName & " and Pid = " & entryParts(1), LogType.LOG_TYPE_INFO)
                        Return entryParts(1).Trim
                    End If
                End If
            Next
        End If
        Return ""
    End Function

#End Region


End Module
