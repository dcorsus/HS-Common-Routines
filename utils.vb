Imports System.IO
Imports System.Runtime.Serialization.Formatters
Imports System.Web.Script.Serialization
Imports System.Net

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

    Private LogFileStreamWriter As StreamWriter = Nothing
    'Private LogFileFileStream As FileStream = Nothing
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
            Log(" Error: Serializing object " & ObjIn.ToString & " :" & ex.Message, LogType.LOG_TYPE_ERROR)
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
            Log(" Error: DeSerializing object: " & ex.Message, LogType.LOG_TYPE_ERROR)
            Return False
        End Try

    End Function

    Public Sub wait(ByVal secs As Decimal)
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("wait called with value = " & secs, LogType.LOG_TYPE_INFO)
        Threading.Thread.Sleep(secs * 1000)
    End Sub

    Enum LogType
        LOG_TYPE_INFO = 0
        LOG_TYPE_ERROR = 1
        LOG_TYPE_WARNING = 2
    End Enum

    Public Function OpenLogFile(LogFileName As String) As Boolean
        OpenLogFile = False
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("OpenLogFile called with LogFileName = " & LogFileName, LogType.LOG_TYPE_INFO)
        If LogFileStreamWriter IsNot Nothing Then
            CloseLogFile()
        End If
        Try
            If LogFileName <> "" Then
                If File.Exists(LogFileName) Then
                    If File.Exists(LogFileName & ".bak") Then
                        File.Delete(LogFileName & ".bak")
                    End If
                    File.Move(LogFileName, LogFileName & ".bak")
                End If
                Try
                    LogFileStreamWriter = File.AppendText(LogFileName)
                    LogFileStreamWriter.AutoFlush = True
                Catch ex1 As Exception
                    If Not ImRunningOnLinux Then Console.WriteLine("Exception creating log file with Error = " & ex1.Message, LogType.LOG_TYPE_ERROR)
                End Try
            End If
            OpenLogFile = True
            MyLogFileName = LogFileName
        Catch ex As Exception
            LogFileStreamWriter = Nothing
            MyLogFileName = ""
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in OpenLogFile with Error = " & ex.Message & " and DiskFileName = " & LogFileName, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Sub CloseLogFile()
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("CloseLogFile called for DiskFileName = " & MyLogFileName, LogType.LOG_TYPE_INFO)
        If LogFileStreamWriter IsNot Nothing Then
            Try
                LogFileStreamWriter.Flush()
                LogFileStreamWriter.Close()
            Catch ex As Exception
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error closing debug disk Log with Error = " & ex.Message & " and DiskFileName = " & MyLogFileName, LogType.LOG_TYPE_ERROR)
            End Try
            Try
                LogFileStreamWriter.Dispose()
            Catch ex As Exception
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error disposing debug disk Log with Error = " & ex.Message & " and DiskFileName = " & MyLogFileName, LogType.LOG_TYPE_ERROR)
            End Try
            LogFileStreamWriter = Nothing
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
                If LogFileStreamWriter IsNot Nothing Then
                    LogFileStreamWriter.WriteLine(DateAndTime.Now.ToString & " : " & msg)
                End If
            End If
        Catch ex As Exception
            LogFileStreamWriter = Nothing
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
                logType = util.LogType.LOG_TYPE_ERROR
            End If
            If Not ImRunningOnLinux Then Console.WriteLine(DateAndTime.Now.ToString & " : " & msg)
            Select Case logType
                Case LogType.LOG_TYPE_ERROR
                    If MsgColor <> "" Then
                        If myHomeSeerSystem IsNot Nothing Then myHomeSeerSystem.WriteLog(HomeSeer.PluginSdk.Logging.ELogType.Error, msg, shortIfaceName, MsgColor)
                        'If Not gLogErrorsOnly Then hs.WriteLogDetail(ShortIfaceName & " Error", msg, MsgColor, "1", "UPnP", ErrorCode)
                    Else
                        If myHomeSeerSystem IsNot Nothing Then myHomeSeerSystem.WriteLog(HomeSeer.PluginSdk.Logging.ELogType.Error, msg, shortIfaceName)
                        'hs.WriteLog(ShortIfaceName & " Error", msg)
                    End If
                Case LogType.LOG_TYPE_WARNING
                    If MsgColor <> "" Then
                        If myHomeSeerSystem IsNot Nothing Then myHomeSeerSystem.WriteLog(HomeSeer.PluginSdk.Logging.ELogType.Warning, msg, shortIfaceName, MsgColor)
                        'If Not gLogErrorsOnly Then hs.WriteLogDetail(ShortIfaceName & " Warning", msg, MsgColor, "0", "UPnP", ErrorCode)
                    Else
                        If myHomeSeerSystem IsNot Nothing Then myHomeSeerSystem.WriteLog(HomeSeer.PluginSdk.Logging.ELogType.Warning, msg, shortIfaceName)
                        'If Not gLogErrorsOnly Then hs.WriteLog(ShortIfaceName & " Warning", msg)
                    End If
                Case LogType.LOG_TYPE_INFO
                    If MsgColor <> "" Then
                        If myHomeSeerSystem IsNot Nothing Then myHomeSeerSystem.WriteLog(HomeSeer.PluginSdk.Logging.ELogType.Info, msg, shortIfaceName, MsgColor)
                        'If Not gLogErrorsOnly Then hs.WriteLogDetail(ShortIfaceName, msg, MsgColor, "0", "UPnP", ErrorCode)
                    Else
                        If myHomeSeerSystem IsNot Nothing Then myHomeSeerSystem.WriteLog(HomeSeer.PluginSdk.Logging.ELogType.Info, msg, shortIfaceName)
                        'If Not gLogErrorsOnly Then hs.WriteLog(ShortIfaceName, msg)
                    End If
            End Select
        Catch ex As Exception
            If Not ImRunningOnLinux Then Console.WriteLine("Exception in LOG of " & IFACE_NAME & ": " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If MyLogFileName <> "" And gLogToDisk Then
                If LogFileStreamWriter IsNot Nothing Then
                    LogFileStreamWriter.WriteLine(DateAndTime.Now.ToString & " : " & msg)
                End If
            End If
        Catch ex As Exception
            LogFileStreamWriter = Nothing
            MyLogFileName = ""
            If Not ImRunningOnLinux Then Console.WriteLine(DateAndTime.Now.ToString & " : " & " Exception in LOG with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            If myHomeSeerSystem IsNot Nothing Then myHomeSeerSystem.WriteLog(HomeSeer.PluginSdk.Logging.ELogType.Error, " Exception in LOG with Error = " & ex.Message, shortIfaceName)
            'hs.WriteLog(ShortIfaceName & " Error", " Exception in LOG with Error = " & ex.Message)
        End Try
    End Sub

    Public Function GetStringIniFile(ByVal Section As String, ByVal Key As String, ByVal DefaultVal As String, Optional FileName As String = "") As String
        GetStringIniFile = ""
        If FileName = "" Then FileName = myINIFile
        Try
            GetStringIniFile = myHomeSeerSystem.GetINISetting(Section, EncodeINIKey(Key), DefaultVal, FileName)
            'Log("GetStringIniFile called with section = " & Section & " and Key = " & Key & " read = " & GetStringIniFile, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            Log("Error in GetStringIniFile with section = " & Section & " and Key = " & EncodeINIKey(Key) & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function GetIntegerIniFile(ByVal Section As String, ByVal Key As String, ByVal DefaultVal As Integer, Optional FileName As String = "") As Integer
        GetIntegerIniFile = 0
        If FileName = "" Then FileName = myINIFile
        Try
            GetIntegerIniFile = myHomeSeerSystem.GetINISetting(Section, EncodeINIKey(Key), DefaultVal, FileName)
            'Log("GetIntegerIniFile called with section = " & Section & " and Key = " & Key & " read = " & GetIntegerIniFile.ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            Log("Error in GetIntegerIniFile with section = " & Section & " and Key = " & EncodeINIKey(Key) & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function GetBooleanIniFile(ByVal Section As String, ByVal Key As String, ByVal DefaultVal As Boolean, Optional FileName As String = "") As Boolean
        GetBooleanIniFile = False
        If FileName = "" Then FileName = myINIFile
        Try
            GetBooleanIniFile = myHomeSeerSystem.GetINISetting(Section, EncodeINIKey(Key), DefaultVal, FileName)
            'Log("GetBooleanIniFile called with section = " & Section & " and Key = " & Key & " read = " & GetBooleanIniFile.ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            Log("Error in GetBooleanIniFile with section = " & Section & " and Key = " & EncodeINIKey(Key) & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Sub WriteStringIniFile(ByVal Section As String, ByVal Key As String, ByVal Value As String, Optional FileName As String = "")
        If FileName = "" Then FileName = myINIFile
        Try
            myHomeSeerSystem.SaveINISetting(Section, EncodeINIKey(Key), Value, FileName)
            'Log("WriteStringIniFile called with section = " & Section & " and Key = " & Key & " Value = " & Value.ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            Log("Error in WriteStringIniFile with section = " & Section & " and Key = " & EncodeINIKey(Key) & " and Value = " & Value.ToString & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub WriteIntegerIniFile(ByVal Section As String, ByVal Key As String, ByVal Value As Integer, Optional FileName As String = "")
        If FileName = "" Then FileName = myINIFile
        Try
            myHomeSeerSystem.SaveINISetting(Section, EncodeINIKey(Key), Value, FileName)
            'Log("WriteIntegerIniFile called with section = " & Section & " and Key = " & Key & " Value = " & Value.ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            Log("Error in WriteIntegerIniFile writing section  " & Section & " and Key = " & EncodeINIKey(Key) & " and Value = " & Value.ToString & " with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub WriteBooleanIniFile(ByVal Section As String, ByVal Key As String, ByVal Value As Boolean, Optional FileName As String = "")
        If FileName = "" Then FileName = myINIFile
        Try
            myHomeSeerSystem.SaveINISetting(Section, EncodeINIKey(Key), Value, FileName)
            'Log("WriteBooleanIniFile called with section = " & Section & " and Key = " & Key & " Value = " & Value.ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            Log("Error in WriteBooleanIniFile writing section  " & Section & " and Key = " & EncodeINIKey(Key) & " and Value = " & Value.ToString & " with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub DeleteEntryIniFile(Section As String, Key As String, Optional FileName As String = "")
        If FileName = "" Then FileName = myINIFile
        Try
            myHomeSeerSystem.SaveINISetting(Section, EncodeINIKey(Key), Nothing, myINIFile)
            If piDebuglevel > DebugLevel.dlEvents Then Log("DeleteEntryIniFile called with section = " & Section & " and Key = " & EncodeINIKey(Key), LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            Log("Error in DeleteEntryIniFile reading " & Section & " section with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Function GetIniSection(ByVal Section As String, Optional FileName As String = "") As Dictionary(Of String, String)
        GetIniSection = Nothing
        If FileName = "" Then FileName = myINIFile
        Try
            Return myHomeSeerSystem.GetIniSection(Section, FileName)
        Catch ex As Exception
            Log("Error in GetIniSection reading " & Section & " section with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Sub DeleteIniSection(ByVal Section As String, Optional FileName As String = "")
        If FileName = "" Then FileName = myINIFile
        Try
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("DeleteIniSection called with section = " & Section & " and FileName = " & FileName, LogType.LOG_TYPE_INFO)
            myHomeSeerSystem.ClearIniSection(Section, FileName)
        Catch ex As Exception
            Log("Error in DeleteIniSection deleting " & Section & " section with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
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

    Public Sub SetFeatureStringByRef(FeatureRef As Integer, NewValue As Object, Update As Boolean)
        If piDebuglevel > DebugLevel.dlEvents Then Log("SetFeatureStringByRef called with Ref = " & FeatureRef.ToString & " and NewValue = " & NewValue.ToString, LogType.LOG_TYPE_INFO)
        If FeatureRef = -1 Then Exit Sub
        Try
            'Dim df As Devices.HsFeature = myHomeSeerSystem.GetFeatureByRef(FeatureRef)
            'Dim status As Dictionary(Of HomeSeer.PluginSdk.Devices.EProperty, Object) = New Dictionary(Of HomeSeer.PluginSdk.Devices.EProperty, Object)()
            'status.Add(HomeSeer.PluginSdk.Devices.EProperty.Status, NewValue)
            'myHomeSeerSystem.UpdateFeatureByRef(FeatureRef, status) ' issue in v.24
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
            Dim newMisc As Dictionary(Of HomeSeer.PluginSdk.Devices.EProperty, Object) = New Dictionary(Of HomeSeer.PluginSdk.Devices.EProperty, Object)()
            newMisc.Add(EProperty.Misc, miscFlags)
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
            Dim newMisc As Dictionary(Of HomeSeer.PluginSdk.Devices.EProperty, Object) = New Dictionary(Of HomeSeer.PluginSdk.Devices.EProperty, Object)()
            newMisc.Add(EProperty.Misc, miscFlags)
            myHomeSeerSystem.UpdateDeviceByRef(deviceRef, newMisc)
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SetHideFlagDevice called with Ref = " & deviceRef.ToString & " and flagOn = " & flagOn.ToString & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Function GetPicture(ByVal url As String) As Image
        ' Get the picture at a given URL.
        Dim web_client As New WebClient With {
            .UseDefaultCredentials = True                         ' added 5/2/2019
            }
        ' web_client.Credentials = CredentialCache.DefaultCredentials     ' added 5/2/2019
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
                    Log("--Device Status = " & dv.Status, LogType.LOG_TYPE_WARNING)
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
                            Log("------Feature Status = " & fv.Status, LogType.LOG_TYPE_WARNING)
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
                                Log("------Feature Status = " & fv.Status, LogType.LOG_TYPE_WARNING)
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
                ElseIf InString(InIndex) = """ Then" Then
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
            Log("Error in EncodeURI. URI = " & InString & " Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
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
            Log("Error in DecodeURI with URI = " & InString & " and Index= " & InIndex.ToString & " and Value = " & Value.ToString & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
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
            Log("Error in EncodeINIKey. URI = " & InString & " Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
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
            Log("Error in DecodeINIKey with URI = " & InString & " and Index= " & InIndex.ToString & " and Value = " & Value.ToString & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        DecodeINIKey = Outstring
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "DecodeURI: In = " & InString & " out = " & Outstring)
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
            If piDebuglevel > DebugLevel.dlEvents And SomethingGotRemoved Then Log("RemoveControlCharacters updated document to = " & inString.ToString, LogType.LOG_TYPE_INFO)
            RemoveControlCharacters = inString
        Catch ex As Exception
            Log("Error in RemoveControlCharacters while retieving document with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
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
            Log("Error in ReplaceSpecialCharacters. URI = " & InString & " Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
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
            Dim Lines As String() = Split(response, {vbCr(0), vbLf(0)})
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

    Public Function CheckLocalIPv4Address(in_address As String) As Boolean
        CheckLocalIPv4Address = False
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
            Log("Error in CheckLocalIPv4Address with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
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
                Log("Warning in GetLocalIPv4Address. Found multiple local IP Addresses. Last one selected = " & GetLocalIPv4Address, LogType.LOG_TYPE_WARNING)
            End If
        Catch ex As Exception
            Log("Error in GetLocalIPv4Address with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function GetLocalMacAddress() As String
        If piDebuglevel > DebugLevel.dlEvents Then Log("GetLocalMacAddress called", LogType.LOG_TYPE_INFO)
        GetLocalMacAddress = ""
        Dim LocalMacAddress As String = ""
        Dim LocalIPAddress = PlugInIPAddress
        If LocalIPAddress = "" Then
            Log("Error in GetLocalMacAddress trying to get own IP address", LogType.LOG_TYPE_ERROR)
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
            Log("Error in GetLocalMacAddress trying to get own MAC address with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetLocalMacAddress trying to get own MAC address, none found", LogType.LOG_TYPE_ERROR)
    End Function

    Public Function GetEthernetPorts() As IEnumerable(Of NetworkInfo)
        If piDebuglevel > DebugLevel.dlEvents Then Log("GetEthernetPorts called", LogType.LOG_TYPE_INFO)
        Dim PortInfo As List(Of NetworkInfo) = New List(Of NetworkInfo)
        Dim Ethernetports As New Dictionary(Of String, String)
        Try
            For Each nic As System.Net.NetworkInformation.NetworkInterface In System.Net.NetworkInformation.NetworkInterface.GetAllNetworkInterfaces()
                Dim netInfo As New NetworkInfo
                netInfo.description = nic.Description
                netInfo.mac = nic.GetPhysicalAddress().ToString
                netInfo.operationalstate = nic.OperationalStatus.ToString
                netInfo.id = nic.Id.ToString
                netInfo.name = nic.Name
                netInfo.interfacetype = nic.NetworkInterfaceType.ToString

                If piDebuglevel > DebugLevel.dlEvents Then Log(String.Format("The MAC address of {0} is {1} with status {2}", nic.Description, nic.GetPhysicalAddress().ToString, nic.OperationalStatus.ToString), LogType.LOG_TYPE_INFO) ' dcor changed level
                If piDebuglevel > DebugLevel.dlEvents Then Log(String.Format("  The ID of {0} has name {1} and interface type {2}", nic.Id.ToString, nic.Name, nic.NetworkInterfaceType.ToString), LogType.LOG_TYPE_INFO)
                If nic.OperationalStatus = Net.NetworkInformation.OperationalStatus.Up Then
                    For Each Ipa In nic.GetIPProperties.UnicastAddresses
                        If piDebuglevel > DebugLevel.dlEvents Then Log(String.Format("    The UniCast address of {0} is {1} with mask {2} and Address family {3}", nic.Description, Ipa.Address.ToString, Ipa.IPv4Mask.ToString, Ipa.Address.AddressFamily.ToString), LogType.LOG_TYPE_INFO) ' dcor level
                        If Ipa.Address.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork Then
                            ' we found an IPv4 IPaddress
                            Dim addrinfo As New NetworkInfoAddrMask
                            addrinfo.address = Ipa.Address.ToString
                            addrinfo.mask = Ipa.IPv4Mask.ToString
                            addrinfo.addrtype = "IPv4 Unicast"
                            If netInfo.addressinfo Is Nothing Then netInfo.addressinfo = New List(Of NetworkInfoAddrMask)
                            netInfo.addressinfo.Add(addrinfo)
                        ElseIf Ipa.Address.AddressFamily = System.Net.Sockets.AddressFamily.InterNetworkV6 Then
                            Dim addrinfo As New NetworkInfoAddrMask
                            addrinfo.address = Ipa.Address.ToString
                            addrinfo.addrtype = "IPv6 Unicast"
                            If netInfo.addressinfo Is Nothing Then netInfo.addressinfo = New List(Of NetworkInfoAddrMask)
                            netInfo.addressinfo.Add(addrinfo)
                        End If
                    Next
                    For Each Ipa In nic.GetIPProperties.AnycastAddresses
                        If piDebuglevel > DebugLevel.dlEvents Then Log(String.Format("    The AnyCast address of {0} is {1} with Address family {2}", nic.Description, Ipa.Address.ToString, Ipa.Address.AddressFamily.ToString), LogType.LOG_TYPE_INFO) ' dcor level
                        Dim addrinfo As New NetworkInfoAddrMask
                        addrinfo.address = Ipa.Address.ToString
                        addrinfo.mask = Ipa.Address.AddressFamily.ToString
                        addrinfo.addrtype = "Anycast"
                        If netInfo.addressinfo Is Nothing Then netInfo.addressinfo = New List(Of NetworkInfoAddrMask)
                        netInfo.addressinfo.add(addrinfo)
                    Next
                    For Each Ipa In nic.GetIPProperties.MulticastAddresses
                        If piDebuglevel > DebugLevel.dlEvents Then Log(String.Format("    The Multicast address of {0} is {1} with Address family {2}", nic.Description, Ipa.Address.ToString, Ipa.Address.AddressFamily.ToString), LogType.LOG_TYPE_INFO) ' dcor level
                        Dim addrinfo As New NetworkInfoAddrMask
                        addrinfo.address = Ipa.Address.ToString
                        addrinfo.mask = Ipa.Address.AddressFamily.ToString
                        addrinfo.addrtype = "Multicast"
                        If netInfo.addressinfo Is Nothing Then netInfo.addressinfo = New List(Of NetworkInfoAddrMask)
                        netInfo.addressinfo.add(addrinfo)
                    Next
                Else
                    ' operational state down
                End If
                PortInfo.Add(netInfo)
            Next
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetEthernetPorts found " & Ethernetports.Count.ToString & " Ethernetports with IPv4 addresses assigned", LogType.LOG_TYPE_INFO)
            If PortInfo.Count > 0 Then
                Return PortInfo
            Else
                Return Nothing
            End If
        Catch ex As Exception
            Log("Error in GetEthernetPorts trying to get own MAC address with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
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
                        If piDebuglevel > DebugLevel.dlEvents Then Log(String.Format("    The UniCast address of {0} is {1} with mask {2} and Address family {3}", nic.Description, Ipa.Address.ToString, Ipa.IPv4Mask.ToString, Ipa.Address.AddressFamily.ToString), LogType.LOG_TYPE_INFO) ' dcor level
                        If Ipa.Address.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork Then
                            If Ipa.Address.ToString = ipAddress Then Return True
                        End If
                    Next
                End If
            Next
        Catch ex As Exception
            Log("Error in IsEthernetPortAlive with Interface = " & ipAddress & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Return False
        End Try
        Return False
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
            Log("Error in EncodeTags. URI = " & InString & " Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        EncodeTags = Outstring
    End Function

    Public Function FindPairInJSONString(inString As String, inName As String) As Object
        Try
            If inString = "" Or inName = "" Then Return Nothing
            If piDebuglevel > DebugLevel.dlEvents Then Log("FindPairInJSONString called with inString = " & inString & " and inName = " & inName, LogType.LOG_TYPE_INFO)
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
                        If piDebuglevel > DebugLevel.dlEvents Then Log("FindPairInJSONString called with inString = " & inString & " did not know type for key = " & inName & " in inString = " & inString, LogType.LOG_TYPE_INFO)
                        Return json.Serialize(Entry.value)
                    End If
                End If
            Next
            If piDebuglevel > DebugLevel.dlEvents Then Log("FindPairInJSONString called with inString = " & inString & " did not find the key = " & inName & " in inString = " & inString, LogType.LOG_TYPE_WARNING)
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

End Module