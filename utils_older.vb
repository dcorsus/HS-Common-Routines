Imports System.IO
Imports System.Runtime.Serialization.Formatters
Imports HomeSeerAPI
Imports System.Web.Script.Serialization

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



    Sub PEDAdd(ByRef PED As clsPlugExtraData, ByVal PEDName As String, ByVal PEDValue As Object)
        Dim ByteObject() As Byte = Nothing
        If PED Is Nothing Then PED = New clsPlugExtraData
        SerializeObject(PEDValue, ByteObject)
        If Not PED.AddNamed(PEDName, ByteObject) Then
            PED.RemoveNamed(PEDName)
            PED.AddNamed(PEDName, ByteObject)
        End If
    End Sub

    Function PEDGet(ByRef PED As clsPlugExtraData, ByVal PEDName As String) As Object
        Dim ByteObject() As Byte
        Dim ReturnValue As New Object
        ByteObject = PED.GetNamed(PEDName)
        If ByteObject Is Nothing Then Return Nothing
        DeSerializeObject(ByteObject, ReturnValue)
        Return ReturnValue
    End Function

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
        Threading.Thread.Sleep(secs * 1000)
    End Sub

    Enum LogType
        LOG_TYPE_INFO = 0
        LOG_TYPE_ERROR = 1
        LOG_TYPE_WARNING = 2
    End Enum

    Public Sub Log(ByVal msg As String, ByVal logType As LogType, Optional ByVal MsgColor As String = "", Optional ErrorCode As Integer = 0)
        Try
            If msg Is Nothing Then msg = ""
            If Not [Enum].IsDefined(GetType(LogType), logType) Then
                logType = util.LogType.LOG_TYPE_ERROR
            End If
            If Not ImRunningOnLinux Then Console.WriteLine(DateAndTime.Now.ToString & " : " & msg)
            Select Case logType
                Case logType.LOG_TYPE_ERROR
                    If MsgColor <> "" Then
                        If Not gLogErrorsOnly Then hs.WriteLogDetail(ShortIfaceName & " Error", msg, MsgColor, "1", "UPnP", ErrorCode)
                    Else
                        hs.WriteLog(ShortIfaceName & " Error", msg)
                    End If
                Case logType.LOG_TYPE_WARNING
                    If MsgColor <> "" Then
                        If Not gLogErrorsOnly Then hs.WriteLogDetail(ShortIfaceName & " Warning", msg, MsgColor, "0", "UPnP", ErrorCode)
                    Else
                        If Not gLogErrorsOnly Then hs.WriteLog(ShortIfaceName & " Warning", msg)
                    End If
                Case logType.LOG_TYPE_INFO
                    If MsgColor <> "" Then
                        If Not gLogErrorsOnly Then hs.WriteLogDetail(ShortIfaceName, msg, MsgColor, "0", "UPnP", ErrorCode)
                    Else
                        If Not gLogErrorsOnly Then hs.WriteLog(ShortIfaceName, msg)
                    End If
            End Select
        Catch ex As Exception
            If Not ImRunningOnLinux Then Console.WriteLine("Exception in LOG of " & IFACE_NAME & ": " & ex.Message, logType.LOG_TYPE_ERROR)
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
            If Not ImRunningOnLinux Then Console.WriteLine(DateAndTime.Now.ToString & " : " & " Exception in LOG with Error = " & ex.Message, logType.LOG_TYPE_ERROR)
            hs.WriteLog(ShortIfaceName & " Error", " Exception in LOG with Error = " & ex.Message)
        End Try
    End Sub

    Public Function OpenLogFile(LogFileName As String) As Boolean
        OpenLogFile = False
        If g_bDebug Then Log("OpenLogFile called with LogFileName = " & LogFileName, LogType.LOG_TYPE_INFO)
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
            If g_bDebug Then Log("Error in OpenLogFile with Error = " & ex.Message & " and DiskFileName = " & LogFileName, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Sub CloseLogFile()
        If g_bDebug Then Log("CloseLogFile called for DiskFileName = " & MyLogFileName, LogType.LOG_TYPE_INFO)
        If LogFileStreamWriter IsNot Nothing Then
            Try
                LogFileStreamWriter.Flush()
                LogFileStreamWriter.Close()
            Catch ex As Exception
                If g_bDebug Then Log("Error closing debug disk Log with Error = " & ex.Message & " and DiskFileName = " & MyLogFileName, LogType.LOG_TYPE_ERROR)
            End Try
            Try
                LogFileStreamWriter.Dispose()
            Catch ex As Exception
                If g_bDebug Then Log("Error disposing debug disk Log with Error = " & ex.Message & " and DiskFileName = " & MyLogFileName, LogType.LOG_TYPE_ERROR)
            End Try
            LogFileStreamWriter = Nothing
        End If
        MyLogFileName = ""
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
            If SuperDebug Then Log("DeleteEntryIniFile called with section = " & Section & " and Key = " & EncodeINIKey(Key), LogType.LOG_TYPE_INFO)
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
            If SuperDebug Then Log("GetIniSection called with section = " & Section & ", FileName = " & FileName & " and # Result = " & UBound(ReturnStrings, 1).ToString, LogType.LOG_TYPE_INFO)
            Dim KeyValues As New Dictionary(Of String, String)()
            For Each Entry In ReturnStrings
                'If g_bDebug Then Log("GetIniSection called with section = " & Section & " found entry = " & Entry.ToString, LogType.LOG_TYPE_INFO)
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
            If g_bDebug Then Log("DeleteIniSection called with section = " & Section & " and FileName = " & FileName, LogType.LOG_TYPE_INFO)
            hs.ClearINISection(Section, FileName)
        Catch ex As Exception
            Log("Error in DeleteIniSection deleting " & Section & " section with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Function EncodeURI(ByVal InString As String) As String
        EncodeURI = InString
        Dim InIndex As Integer = 0
        Dim Outstring As String = ""
        InString = Trim(InString)
        If InString = "" Then Exit Function
        Try
            Do While InIndex < InString.Length
                If InString(InIndex) = " " Then
                    Outstring = Outstring + "%20"
                ElseIf InString(InIndex) = "!" Then
                    Outstring = Outstring + "%21"
                ElseIf InString(InIndex) = """ Then" Then
                    Outstring = Outstring + "%22"
                ElseIf InString(InIndex) = "#" Then
                    Outstring = Outstring + "%23"
                ElseIf InString(InIndex) = "$" Then
                    Outstring = Outstring + "%24"
                ElseIf InString(InIndex) = "%" Then
                    Outstring = Outstring + "%25"
                ElseIf InString(InIndex) = "&" Then
                    Outstring = Outstring + "%26"
                ElseIf InString(InIndex) = "'" Then
                    Outstring = Outstring + "%27"
                ElseIf InString(InIndex) = "(" Then
                    Outstring = Outstring + "%28"
                ElseIf InString(InIndex) = ")" Then
                    Outstring = Outstring + "%29"
                ElseIf InString(InIndex) = "*" Then
                    Outstring = Outstring + "%2A"
                ElseIf InString(InIndex) = "+" Then
                    Outstring = Outstring + "%2B"
                ElseIf InString(InIndex) = "," Then
                    Outstring = Outstring + "%2C"
                ElseIf InString(InIndex) = "-" Then
                    Outstring = Outstring + "%2D"
                ElseIf InString(InIndex) = "." Then
                    Outstring = Outstring + "%2E"
                ElseIf InString(InIndex) = "/" Then
                    Outstring = Outstring + "%2F"
                Else
                    Outstring = Outstring & InString(InIndex)
                End If
                InIndex = InIndex + 1
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
                        If ((InString(InIndex + 1) >= "0") And (InString(InIndex + 1) <= "9")) Or _
                            ((InString(InIndex + 1) >= "A") And (InString(InIndex + 1) <= "F")) Or _
                             ((InString(InIndex + 1) >= "a") And (InString(InIndex + 1) <= "f")) Then
                            ' now check the same for the second character
                            If ((InString(InIndex + 2) >= "0") And (InString(InIndex + 2) <= "9")) Or _
                                ((InString(InIndex + 2) >= "A") And (InString(InIndex + 2) <= "F")) Or _
                                ((InString(InIndex + 2) >= "a") And (InString(InIndex + 2) <= "f")) Then
                                ' all checks passed now convert
                                Value = 0
                                Dim Char1, Char2 As Char
                                Char1 = UCase(InString(InIndex + 1))
                                Char2 = UCase(InString(InIndex + 2))
                                If (Char1 >= "0") And (Char1 <= "9") Then
                                    Value = Value + Val(Char1) * 16
                                Else
                                    ' convert 
                                    Char1 = ChrW(AscW(Char1) - 17)
                                    Value = Value + (Val(Char1) + 10) * 16
                                End If
                                If (Char2 >= "0") And (Char2 <= "9") Then
                                    Value = Value + Val(Char2)
                                Else
                                    ' convert 
                                    Char2 = ChrW(AscW(Char2) - 17)
                                    Value = Value + (Val(Char2) + 10)
                                End If
                                Outstring = Outstring & ChrW(Value)
                                InIndex = InIndex + 3
                            Else
                                Outstring = Outstring & InString(InIndex)
                                InIndex = InIndex + 1
                            End If
                        Else
                            Outstring = Outstring & InString(InIndex)
                            InIndex = InIndex + 1
                        End If
                    Else
                        Outstring = Outstring & InString(InIndex)
                        InIndex = InIndex + 1
                    End If
                    'ElseIf InString(InIndex) = "&" Then
                    '             string = string.replace(/&amp;/g, "&");  
                    '             string = string.replace(/&quot;/g, "\"");  
                    '             string = string.replace(/&apos;/g, "'");  
                    '             string = string.replace(/&lt;/g, "<");  
                    '             string = string.replace(/&gt;/g, ">"); 
                Else
                    Outstring = Outstring & InString(InIndex)
                    InIndex = InIndex + 1
                End If
            Loop
        Catch ex As Exception
            Log("Error in DecodeURI with URI = " & InString & " and Index= " & InIndex.ToString & " and Value = " & Value.ToString & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        DecodeURI = Outstring
        'If g_bDebug Then log( "DecodeURI: In = " & InString & " out = " & Outstring)
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
                    Outstring = Outstring + "%3D"
                ElseIf InString(InIndex) = "%" Then
                    Outstring = Outstring + "%25"
                Else
                    Outstring = Outstring & InString(InIndex)
                End If
                InIndex = InIndex + 1
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
                        If ((InString(InIndex + 1) >= "0") And (InString(InIndex + 1) <= "9")) Or _
                            ((InString(InIndex + 1) >= "A") And (InString(InIndex + 1) <= "F")) Or _
                             ((InString(InIndex + 1) >= "a") And (InString(InIndex + 1) <= "f")) Then
                            ' now check the same for the second character
                            If ((InString(InIndex + 2) >= "0") And (InString(InIndex + 2) <= "9")) Or _
                                ((InString(InIndex + 2) >= "A") And (InString(InIndex + 2) <= "F")) Or _
                                ((InString(InIndex + 2) >= "a") And (InString(InIndex + 2) <= "f")) Then
                                ' all checks passed now convert
                                Value = 0
                                Dim Char1, Char2 As Char
                                Char1 = UCase(InString(InIndex + 1))
                                Char2 = UCase(InString(InIndex + 2))
                                If (Char1 >= "0") And (Char1 <= "9") Then
                                    Value = Value + Val(Char1) * 16
                                Else
                                    ' convert 
                                    Char1 = ChrW(AscW(Char1) - 17)
                                    Value = Value + (Val(Char1) + 10) * 16
                                End If
                                If (Char2 >= "0") And (Char2 <= "9") Then
                                    Value = Value + Val(Char2)
                                Else
                                    ' convert 
                                    Char2 = ChrW(AscW(Char2) - 17)
                                    Value = Value + (Val(Char2) + 10)
                                End If
                                Outstring = Outstring & ChrW(Value)
                                InIndex = InIndex + 3
                            Else
                                Outstring = Outstring & InString(InIndex)
                                InIndex = InIndex + 1
                            End If
                        Else
                            Outstring = Outstring & InString(InIndex)
                            InIndex = InIndex + 1
                        End If
                    Else
                        Outstring = Outstring & InString(InIndex)
                        InIndex = InIndex + 1
                    End If
                    'ElseIf InString(InIndex) = "&" Then
                    '             string = string.replace(/&amp;/g, "&");  
                    '             string = string.replace(/&quot;/g, "\"");  
                    '             string = string.replace(/&apos;/g, "'");  
                    '             string = string.replace(/&lt;/g, "<");  
                    '             string = string.replace(/&gt;/g, ">"); 
                Else
                    Outstring = Outstring & InString(InIndex)
                    InIndex = InIndex + 1
                End If
            Loop
        Catch ex As Exception
            Log("Error in DecodeINIKey with URI = " & InString & " and Index= " & InIndex.ToString & " and Value = " & Value.ToString & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        DecodeINIKey = Outstring
        'If g_bDebug Then log( "DecodeURI: In = " & InString & " out = " & Outstring)
    End Function

    Public Function DeriveIPAddress(inString As String, NextChar As String) As String
        If SuperDebug Then Log("DeriveIPAddress called for and inString = " & inString & " and NextChar = " & NextChar, LogType.LOG_TYPE_INFO)
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
            If SuperDebug Then Log("RemoveControlCharacters retrieved document with length = " & strIndex.ToString, LogType.LOG_TYPE_INFO)
            Dim SomethingGotRemoved As Boolean = False
            While strIndex > 0
                strIndex = strIndex - 1
                If inString(strIndex) < " " Then
                    inString = inString.Remove(strIndex, 1)
                    SomethingGotRemoved = True
                End If
            End While
            inString = Trim(inString)
            If SuperDebug And SomethingGotRemoved Then Log("RemoveControlCharacters updated document to = " & inString.ToString, LogType.LOG_TYPE_INFO)
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
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "!" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = """ Then" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "#" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "$" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "%" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "&" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "'" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "(" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = ")" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "*" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "+" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "," Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "-" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "." Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = ":" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "/" Then
                    Outstring = Outstring + "_"
                Else
                    Outstring = Outstring & InString(InIndex)
                End If
                InIndex = InIndex + 1
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
                            Body = Body & Line
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
            If g_bDebug Then Log("CheckLocalIPv4Address found IP Host = " & strHostName, LogType.LOG_TYPE_INFO)
            For Each ipheal As System.Net.IPAddress In iphe.AddressList
                If g_bDebug Then Log("CheckLocalIPv4Address found IP Address = " & ipheal.ToString() & " with AddressFamily = " & ipheal.AddressFamily.ToString, LogType.LOG_TYPE_INFO)
                If ipheal.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork Then
                    If ipheal.ToString = in_address Then
                        If g_bDebug Then Log("CheckLocalIPv4Address found IP Address = " & ipheal.ToString() & " equal to HS IP address so Plugin is running local", LogType.LOG_TYPE_INFO)
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
            If g_bDebug Then Log("GetLocalIPv4Address found IP Host = " & strHostName, LogType.LOG_TYPE_INFO)
            Dim NbrOfIPv4Interfaces As Integer = 0
            For Each ipheal As System.Net.IPAddress In iphe.AddressList
                If g_bDebug Then Log("GetLocalIPv4Address found IP Address = " & ipheal.ToString() & " with AddressFamily = " & ipheal.AddressFamily.ToString, LogType.LOG_TYPE_INFO)
                If ipheal.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork Then
                    If g_bDebug Then Log("GetLocalIPv4Address found Local IP Address = " & ipheal.ToString(), LogType.LOG_TYPE_INFO)
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
        If SuperDebug Then Log("GetLocalMacAddress called", LogType.LOG_TYPE_INFO)
        GetLocalMacAddress = ""
        Dim LocalMacAddress As String = ""
        Dim LocalIPAddress = PlugInIPAddress
        If LocalIPAddress = "" Then
            Log("Error in GetLocalMacAddress trying to get own IP address", LogType.LOG_TYPE_ERROR)
            Exit Function
        End If
        Try
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
        Catch ex As Exception
            Log("Error in GetLocalMacAddress trying to get own MAC address with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If g_bDebug Then Log("Error in GetLocalMacAddress trying to get own MAC address, none found", LogType.LOG_TYPE_ERROR)
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
                    Outstring = Outstring + "&gt;"
                ElseIf InString(InIndex) = "<" Then
                    Outstring = Outstring + "&lt;"
                ElseIf InString(InIndex) = "'" Then
                    Outstring = Outstring + "&#39;"
                Else
                    Outstring = Outstring & InString(InIndex)
                End If
                InIndex = InIndex + 1
            Loop
        Catch ex As Exception
            Log("Error in EncodeTags. URI = " & InString & " Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        EncodeTags = Outstring
    End Function

    Public Function FindPairInJSONString(inString As String, inName As String) As String
        FindPairInJSONString = ""
        If SuperDebug Then Log("FindPairInJSONString called with inString = " & inString & " and inName = " & inName, LogType.LOG_TYPE_INFO)
        Try
            Dim json As New JavaScriptSerializer
            Dim JSONdataLevel1 As Object
            JSONdataLevel1 = json.DeserializeObject(inString)
            For Each Entry As Object In JSONdataLevel1
                If Entry.Key = inName Then
                    Return json.Serialize(Entry.value)
                End If
            Next
        Catch ex As Exception
            If g_bDebug Then Log("Error in FindPairInJSONString processing response with error = " & ex.Message & " with inString = " & inString & " and inName = " & inName, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

End Module
