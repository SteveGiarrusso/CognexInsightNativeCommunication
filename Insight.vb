Imports System.IO
Imports System.Net
Imports System.Net.Sockets
Imports System.Threading
Imports System.Drawing
Imports System.Xml
Imports System.IO.MemoryMappedFiles
Imports Cognex.InSight
Imports Cognex.InSight.NativeMode
Public Class InsightCamera
    '8087 for HMI ports 
    '23 for Telnet Ports

    Private Delegate Sub DlgSub(Arg As Object) 'generic delegate sub used to invoke events

    Dim WithEvents Native As CvsNativeModeClient
    'Dim WithEvents Camera As CvsInSight
    ''' <summary>
    ''' Native Mode Communication Client
    ''' </summary>
    ''' <remarks></remarks>
    Public Event StringRecieved(ByVal Response As String, ByVal LastCommand As String)
    ''' <summary>
    ''' A command was sent
    ''' </summary>
    ''' <param name="LastCommand">The last command sent</param>
    ''' <remarks></remarks>
    Public Event CommandSent(ByVal LastCommand As String)
    ''' <summary>
    ''' Connection Status of Camera has changed
    ''' </summary>
    ''' <param name="status">Is Camera Connected</param>
    ''' <remarks></remarks>
    Public Event ConnectionStatusChanged(ByVal status As Boolean)
    ''' <summary>
    ''' An Error has Occured
    ''' </summary>
    ''' <param name="CustomMessage">Message defining the error </param>
    ''' <param name="Exception">Exception Message</param>
    ''' <param name="Command">The command the camera failed on</param>
    ''' <remarks></remarks>
    Public Event RaiseError(ByVal CustomMessage As String, ByVal Exception As String, ByVal Command As String)
    ''' <summary>
    ''' Online Mode has Changed
    ''' </summary>
    ''' <param name="Online">Is the camera online</param>
    ''' <remarks></remarks>
    Public Event OnlineMode(Online As Boolean)
    ''' <summary>
    ''' A log Event has occured
    ''' </summary>
    ''' <param name="Message">Message to be Logged</param>
    ''' <param name="_LastCommand">The Last COmmand Sent</param>
    ''' <remarks></remarks>
    Public Event Log(ByVal Message As String, _LastCommand As String)


    ''' <summary>
    ''' Camera IP
    ''' </summary>
    ''' <remarks></remarks>
    Private _IP As String
    ''' <summary>
    ''' Telnet Port
    ''' </summary>
    ''' <remarks></remarks>
    Private _Port As Integer = 23
    ''' <summary>
    ''' Port for Web HMI
    ''' </summary>
    ''' <remarks></remarks>
    Private _WebPort As Integer = 8087
    ''' <summary>
    ''' Is Camera in Live Mode
    ''' </summary>
    ''' <remarks></remarks>
    Private _isLive As Boolean
    ''' <summary>
    ''' Current Job Loaded on the Camera
    ''' </summary>
    ''' <remarks></remarks>
    Private _CurrentJob As String
    ''' <summary>
    ''' Is the Camera Online
    ''' </summary>
    ''' <remarks></remarks>
    Private _Online As Boolean
    ''' <summary>
    ''' Is the Camera Connected
    ''' </summary>
    ''' <remarks></remarks>
    Private _ConnectionStatus As Boolean
    ''' <summary>
    ''' Hostname of Camera
    ''' </summary>
    ''' <remarks></remarks>
    Private _Hostname As String


    ' Local Consts used for building command strings
    Private Const cLoadFile As String = "LF"
    Private Const cGetFile As String = "GF"
    Private Const cReadImage As String = "RI"
    Private Const cGetValue As String = "GV"
    Private Const cSetInetger As String = "SI"
    Private Const cSetFloat As String = "SF"
    Private Const cSetString As String = "SS"
    Private Const cGetInfo As String = "GI"
    Private Const cSetOnline As String = "SO1"
    Private Const cSetOffline As String = "SO0"
    Private Const cGetOnline As String = "GO"
    Private Const cSetEvent As String = "SE"
    Private Const cSetEventWait As String = "SW"
    Private Const cResetSystem As String = "RT"
    Private Const cSendMessage As String = "SM"
    Private Const cGet As String = "Get "
    Private Const cGetConnections As String = "Get Connections"
    Private Const cGetExpr As String = "Get Expr "
    Private Const cGetFilelist As String = "Get Filelist"
    Private Const cPutLive As String = "Put Live "
    Private Const cPutUpdate As String = "Put Update 1"
    Private Const cPutWatch As String = "Put Watch "
    Private Const cEvaluate As String = "EV "
    Private Const cGetCellName As String = "EV GetCellName"
    Private Const cGetCellValue As String = "EV GetCellValue"
    Private Const cGetDiagnosticLog As String = "EV GetDiagnosticLog()"

    ''' <summary>
    ''' Creates New Instance of Camera
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        Native = New CvsNativeModeClient
    End Sub

#Region "Properties"
    ''PROPERTIES
    Public Property IP As String
        Get
            Return _IP
        End Get
        Set(value As String)
            _IP = value
        End Set
    End Property
    Public Property Port As Integer
        Get
            Return _Port
        End Get
        Set(value As Integer)
            _Port = value
        End Set
    End Property

    Public ReadOnly Property CurrentJob As String
        Get
            _CurrentJob = GetFile()
            Return _CurrentJob
        End Get
    End Property
    Public Property Hostname As String
        Get
            Return _Hostname
        End Get
        Set(value As String)
            _Hostname = value
        End Set
    End Property
    Public Property Online As Boolean
        Get
            Return GetOnline()
        End Get
        Set(value As Boolean)
            If value = True Then
                SetOnline()
            ElseIf value = False Then
                SetOffline()
            End If
        End Set
    End Property
    Public Property IsLive As Boolean
        Get
            Return _IsLive
        End Get
        Set(value As Boolean)
            If value = True Then
                PutLive(True)
            ElseIf value = False Then
                PutLive(False)
            End If
        End Set
    End Property
    Public ReadOnly Property ConnectionStatus As Boolean
        Get
            updateConnectionStatus()
            Return _ConnectionStatus
        End Get
    End Property
    
#End Region
#Region "Connections"
    ''' <summary>
    ''' Pulls Device info from CognexDBM and Starts TCP Client/Server
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Connect()
        Try
            Native.Connect(_IP, "admin", "")

            updateConnectionStatus()
            RaiseEvent Log("Connected", "Connect")
        Catch ex As Exception
            Throw New Exception("Connection not established to Camera" & vbCrLf & ex.Message & vbCrLf & "Connect")
            'RaiseEvent GenericError()
            Exit Sub
        End Try



    End Sub
    ''' <summary>
    ''' Stops TCP Client and Server'
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Disconnect()
        Try
            Native.Disconnect()
            RaiseEvent Log("Disconnected", "Disconnect")
        Catch ex As Exception
            Throw New Exception("Error Disconnecting & vbCrLf & ex.Message & vbCrLf" & "Disconnect")
        End Try
        updateConnectionStatus()
    End Sub
    ''' <summary>
    ''' Updates Camera Connection Status
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub updateConnectionStatus()
        Try
            _ConnectionStatus = Native.Connected
        Catch ex As Exception
            _ConnectionStatus = False
        End Try

        RaiseEvent ConnectionStatusChanged(_ConnectionStatus)
    End Sub

#End Region

#Region "SendCognexCommands"

    ''' <summary>
    ''' Send a Given Command over TCP/IP to Camera
    ''' </summary>
    ''' <param name="TelnetCommand">The command as a string, with no parameters</param>
    ''' <param name="Parameters">All parameters concatnated as a string</param>
    ''' <remarks></remarks>
    Private Function SendCommand(XMLElement As String, TelnetCommand As String, Optional Parameters As String = "") As String
        Dim SentCommand As String
        updateConnectionStatus()
        If _ConnectionStatus = False Then Throw New Exception("Not Connected to Camera!!!!")
        Dim pResponse As String
        'Try
        SentCommand = TelnetCommand & Parameters
        pResponse = Native.SendCommand(SentCommand & vbCrLf)
        RaiseEvent Log("Sent Command: " & SentCommand, TelnetCommand & Parameters)
        Dim pStatus As Integer
        pStatus = ReadXML("Status", pResponse)
        Select Case pStatus
            Case Response.CommandSuccsess
                'do nothing!
            Case Response.CellIDorTagInvalid
                Throw New Exception("Command " & SentCommand & " is invalid.")
            Case Response.CommandCouldNotBeExecuted
                Throw New Exception("Command " & SentCommand & " failed.")
            Case Response.UnRecognizedCommand
                Throw New Exception(SentCommand & " is an unrecognized command.")
            Case Response.SystemOutOfMemory
                Throw New Exception("Camera is out of memory.")
            Case Response.AccessDenied
                Throw New Exception("User does not have Full Access to execute the command, " & SentCommand)
            Case Response.CannotGoOnlineInManualMode
                Throw New Exception("Cannot go Live in Manual Mode " & SentCommand)
            Case vbNull
                Throw New Exception("Unable to find <Status> from response, " & pResponse & ", to command, " & SentCommand)
            Case Else
                Throw New Exception("Response, " & pStatus & ", is not recognized but indicates an error.")
        End Select
        If XMLElement <> "" Then
            Dim str As String = ReadXML(XMLElement, pResponse)
            RaiseEvent Log("Recieved: " & str, TelnetCommand & Parameters)
            Return str
        Else
            RaiseEvent Log("Recieved: " & pResponse, TelnetCommand & Parameters)
            Return pResponse
        End If

    End Function
    ''' <summary>
    ''' Adds leading zeros so all ints will be in format of 000. All spreadsheet refrences must be made like this
    ''' </summary>
    ''' <param name="int">Value to add zeros to</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function AddLeadingZeros(ByVal int As Integer) As String
        Dim RowString As String
        If int.ToString.Length = 1 Then
            RowString = "00" & int
        ElseIf int.ToString.Length = 2 Then
            RowString = "0" & int
        Else
            RowString = int
        End If
        Return RowString
    End Function

    'Enums
    ''' <summary>
    ''' Generic Responses from sending camera a commmand
    ''' </summary>
    ''' <remarks></remarks>
    Enum Response
        CommandSuccsess = 1
        UnRecognizedCommand = 0
        CellIDorTagInvalid = -1
        CommandCouldNotBeExecuted = -2
        SystemOutOfMemory = -4
        CannotGoOnlineInManualMode = -5
        AccessDenied = -6
    End Enum
    ''' <summary>
    ''' Responses to Get Online
    ''' </summary>
    ''' <remarks></remarks>
    Private Enum GetOnlineResponses
        Offline = 0
        Online = 1
    End Enum

    'Load Job
    ''' <summary>
    ''' Loads the specified jobfrom flash memory on the In-Sight vision system, RAM Disk or SD Card, making it the active job.
    ''' </summary>
    ''' <param name="Filename">Name of Job to be loaded, including .job</param>
    ''' <remarks></remarks>
    Function LoadFile(ByVal Filename As String) As Integer
        Dim _CommandResponse As String
        _CommandResponse = SendCommand("Status", cLoadFile, Filename)
        Return CInt(_CommandResponse)
    End Function


    'Get File
    ''' <summary>
    ''' Returns the filename of the active jobon the In-Sight vision system, RAM Disk or SD Card.
    ''' </summary>
    ''' <remarks></remarks>
    Function GetFile() As String
        Dim _CommandResponse As String
        _CommandResponse = SendCommand("FileName", cGetFile)
        Try
            Return Path.GetFileName(_CommandResponse)
        Catch
            Return _CommandResponse
        End Try

    End Function


    'Read Image
    ''' <summary>
    ''' Sends the current image, in ASCII hexadecimal format (formatted to 80characters per line), from an In-Sight sensor out to a remote device.When converted to binary, the resulting data is in standard BMP imageformat.
    ''' </summary>
    ''' <remarks></remarks>
    Function ReadImage() As String
        Dim _CommandResponse As String
        _CommandResponse = SendCommand("Data", cReadImage)
        Return _CommandResponse
    End Function
    'Line 1 - Status Code
    'Line 2 - Inetger Val for Bytes
    'Lines - Image Data
    'Line Last - Checksum '61EE'

    'GetValue
    ''' <summary>
    ''' Returns the value contained in the specified spreadsheet cell
    ''' </summary>
    ''' <param name="ColumnLetter">The spreadsheet Column Letter of the Cell value (A to Z)</param>
    ''' <param name="RowNumber">The spreadsheet Row Number as Integer</param>
    ''' <remarks></remarks>
    Function GetValue(ColumnLetter As String, RowNumber As Integer) As Double
        Dim _CommandResponse As String
        _CommandResponse = SendCommand("CellOutput", cGetValue, ColumnLetter & AddLeadingZeros(RowNumber))
        Try
            Return CDbl(_CommandResponse)
        Catch ex As Exception
            Return _CommandResponse
        End Try
    End Function
    ''' <summary>
    ''' 'Returns the contents of a specified symbolic tag, such as an EasyBuilder Location or Inspection Tool result or job data
    ''' </summary>
    ''' <param name="SymbolicTag">The name of the symbolic tag [such as a Location or Inspection Tool parameter</param>
    ''' <remarks></remarks>
    Function GetValue(SymbolicTag As String) As Double
        Dim _CommandResponse As String
        _CommandResponse = SendCommand("CellOutput", cGetValue, SymbolicTag)
        Try
            Return CDbl(_CommandResponse)
        Catch ex As Exception
            Return _CommandResponse
        End Try
    End Function
    'GetString
    ''' <summary>
    ''' Returns the string contained in the specified spreadsheet cell
    ''' </summary>
    ''' <param name="ColumnLetter">The spreadsheet Column Letter of the Cell value (A to Z)</param>
    ''' <param name="RowNumber">The spreadsheet Row Number as Integer</param>
    ''' <remarks></remarks>
    Function GetString(ColumnLetter As String, RowNumber As Integer) As String
        Dim _CommandResponse As String
        _CommandResponse = SendCommand("CellOutput", cGetValue, ColumnLetter & AddLeadingZeros(RowNumber))
        Try
            Return _CommandResponse
        Catch ex As Exception
            Return _CommandResponse
        End Try
    End Function
    ''' <summary>
    ''' 'Returns the contents of a specified symbolic tag, such as an EasyBuilder Location or Inspection Tool result or job data
    ''' </summary>
    ''' <param name="SymbolicTag">The name of the symbolic tag [such as a Location or Inspection Tool parameter</param>
    ''' <remarks></remarks>
    Function GetString(SymbolicTag As String) As String
        Dim _CommandResponse As String
        _CommandResponse = SendCommand("CellOutput", cGetValue, SymbolicTag)
        Try
            Return _CommandResponse
        Catch ex As Exception
            Return _CommandResponse.ToString
        End Try
    End Function


    'SetInteger
    ''' <summary>
    ''' Sets the control contained in a cell to the specified integer value. The control must be of the types EditInt, Checkbox,or ListBox
    ''' </summary>
    ''' <param name="Column">The column letter of the cell value to set(A to Z).</param>
    ''' <param name="Row">The row number of the cell value to set.</param>
    ''' <param name="int"> The integer value to set. </param>
    ''' <remarks></remarks>
    Function SetInteger(ByVal Column As String, Row As Integer, int As Integer) As Integer
        Dim _CommandResponse As String
        _CommandResponse = SendCommand("Status", cSetInetger, Column & AddLeadingZeros(Row) & int)
        Return CInt(_CommandResponse)
    End Function
    ''' <summary>
    ''' Sets the integer value of a symbolic tag, such as a parameter contained in an EasyBuilder Location or Inspection Tool, or job data, to the specified integer value. The parameter or job data must be an Integer Data Type. 
    ''' </summary>
    ''' <param name="Name"> The name of the Location or Inspection Tool parameter ("Pattern_1.Model_Type", for example) or EasyBuilder job data ("Job.External_Reset_Counters", for example) to be set. </param>
    ''' <param name="int">The integer value to set.</param>
    ''' <remarks></remarks>
    Function SetInteger(ByVal Name As String, int As Integer) As Integer
        Dim _CommandResponse As String
        _CommandResponse = SendCommand("Status", cSetInetger, Name & " " & int)
        Return CInt(_CommandResponse)
    End Function


    'SetFloat
    ''' <summary>
    ''' Sets an edit box control contained in a cell to a specified floating-point value. The edit box control must be of theEditFloat type.
    ''' </summary>
    ''' <param name="Column">The column letter of the cell value to set(A to Z).</param>
    ''' <param name="Row">The row number of the cell value to set</param>
    ''' <param name="float">The floating-point value to set, includingthe decimal point (.) character.</param>
    ''' <remarks></remarks>
    Function SetFloat(ByVal Column As String, Row As Integer, float As Double) As Integer
        Dim _CommandResponse As String
        _CommandResponse = SendCommand("Status", cSetFloat, Column & AddLeadingZeros(Row) & float)
        Return CInt(_CommandResponse)
    End Function
    ''' <summary>
    ''' Sets the floating-point value of a symbolic tag, such as an EasyBuilder Location or Inspection Tool parameter, or job data, to the specified floating-point value. The symbolic tag must be a Floating Point Data Type.
    ''' </summary>
    ''' <param name="SymbolicTag">The name of the symbolic tag [such as a Location or Inspection Tool parameter ("Pattern_1.Horizontal_Offset", for example) or EasyBuilder job data ("Acquistion.Exposure_Time", for example)] to be set.</param>
    ''' <param name="Float">The floating-point value to set, includingthe decimal point (.) character</param>
    ''' <remarks></remarks>
    Function SetFloat(ByVal SymbolicTag As String, ByVal Float As Double) As Integer
        Dim _CommandResponse As String
        _CommandResponse = SendCommand("Status", cSetFloat, SymbolicTag & " " & Float)
        Return CInt(_CommandResponse)
    End Function


    'SetString
    ''' <summary>
    ''' Sets an edit box control contained in a cellto a specified string. The edit box must be of the type EditString.
    ''' </summary>
    ''' <param name="Column">The column letter of the cell value to set(A to Z).</param>
    ''' <param name="Row">The row number of the cell value to set</param>
    ''' <param name="str">The string to set</param>
    ''' <remarks></remarks>
    Function SetString(ByVal Column As String, Row As Integer, str As String) As Integer
        Dim _CommandResponse As String
        _CommandResponse = SendCommand("Status", cSetString, Column & AddLeadingZeros(Row) & str)
        Return CInt(_CommandResponse)
    End Function
    ''' <summary>
    ''' Sets an edit box control contained in a cellto a specified string. The edit box must be of the type EditString
    ''' </summary>
    ''' <param name="SymbolicTag">The name of the symbolic tag [such as a Location or Inspection Tool parameter ("Pattern_1.Horizontal_Offset", for example) or EasyBuilder job data ("Acquistion.Exposure_Time", for example)] to be set.</param>
    ''' <param name="str">The string to set</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function SetString(ByVal SymbolicTag As String, str As String) As Integer
        Dim _CommandResponse As String
        _CommandResponse = SendCommand("Status", cSetString, SymbolicTag & " " & str)
        Return CInt(_CommandResponse)
    End Function

    'GetInfo
    ''' <summary>
    ''' Returns a status code, followed by the system information
    ''' </summary>
    ''' <remarks></remarks>
    Function GetInfo(Optional InfoXMLElement As String = "") As String
        Dim _CommandResponse As String
        _CommandResponse = SendCommand(InfoXMLElement, cGetInfo)
        Return _CommandResponse
    End Function


    'SetOnline/Offline
    ''' <summary>
    ''' Sets the Camera in Online Mode
    ''' </summary>
    ''' <remarks></remarks>
    Function SetOnline() As Integer
        Dim _CommandResponse As String
        _CommandResponse = SendCommand("Status", cSetOnline)
        If _CommandResponse = 1 Then RaiseEvent OnlineMode(1)
        Return CInt(_CommandResponse)
    End Function
    ''' <summary>
    ''' Sets the Camera in Offline Mode
    ''' </summary>
    ''' <remarks></remarks>
    Function SetOffline() As Integer
        Dim _CommandResponse As String
        _CommandResponse = SendCommand("Status", cSetOffline)
        If _CommandResponse = 1 Then RaiseEvent OnlineMode(0)
        Return CInt(_CommandResponse)
    End Function


    'GetOnline
    ''' <summary>
    ''' Returns the Onlinestate of the In-Sight vision system
    ''' </summary>
    ''' <remarks></remarks>
    Function GetOnline() As Integer
        Dim _CommandResponse As String
        _CommandResponse = SendCommand("Online", cGetOnline)
        RaiseEvent OnlineMode(CInt(_CommandResponse))
        Return CInt(_CommandResponse)
    End Function


    'SetEVent/SetEventWait

    ''' <summary>
    ''' Triggers a specified event in the spreadsheet through a Native Mode command.
    ''' </summary>
    ''' <param name="int">The Event code to set. •8 = Acquirean image and update the spreadsheet. This option requires the AcquireImagefunction's Trigger parameter tobe set to External, Manual or Network.</param>
    ''' <remarks>•0 to 7 = Specifies a soft trigger </remarks>
    Function SetEvent(int As Integer) As Integer
        Dim _CommandResponse As String
        _CommandResponse = SendCommand("Status", cSetEvent, int)
        Return CInt(_CommandResponse)
    End Function
    ''' <summary>
    ''' Runs SetEvent to Aquire Image and Update Spreadsheet
    ''' </summary>
    ''' <remarks></remarks>
    Function AcquireImage() As Integer
        Dim _CommandResponse As String
        _CommandResponse = SendCommand("Status", cSetEvent, 8)
        Return CInt(_CommandResponse)
        '8 Is is the canned soft event for aquire Image and update spreadsheet
    End Function
    ''' <summary>
    ''' Triggers a specified event and waits until the commandis completed to return a response..
    ''' </summary>
    ''' <param name="int">The Event code to set. •8 = Acquirean image and update the spreadsheet. This option requires the AcquireImagefunction's Trigger parameter tobe set to External, Manual or Network.</param>
    ''' <remarks>•0 to 7 = Specifies a soft trigger </remarks>
    Function SetEventandWait(int As Integer) As Integer
        Dim _CommandResponse As String
        _CommandResponse = SendCommand("Status", cSetEventWait, int)
        Return CInt(_CommandResponse)
    End Function
    ''' <summary>
    ''' Runs SetEvent to Aquire Image and Update Spreadsheet, and waits until the commandis completed to return a response
    ''' </summary>
    ''' <remarks></remarks>
    Function AcquireImageandWait() As Integer
        Dim _CommandResponse As String
        _CommandResponse = SendCommand("Status", cSetEventWait, 8)
        Return CInt(_CommandResponse)
        '8 Is is the canned soft event for aquire Image and update spreadsheet
    End Function


    'Reset
    ''' <summary>
    ''' Resets the In-Sight sensor. This command is similarto physically cycling power on the sensor.
    ''' </summary>
    ''' <remarks></remarks>
    Sub ResetSystem()
        RaiseEvent Log("Reset System", "RT")
        updateConnectionStatus()
        Native.SendCommand(cResetSystem)
        updateConnectionStatus()
    End Sub

    'Send Message
    ''' <summary>
    ''' Sends a string to an In-Sight spreadsheetover a Native Mode connection,and optionally, triggers a spreadsheet Event. 
    ''' </summary>
    ''' <param name="str">Thestring to set</param>
    ''' <param name="int">The Event code to set. This is anoptional parameter.</param>
    ''' <remarks></remarks>
    Function SendMessage(ByVal str As String, Optional int As Integer = -1) As Integer
        Dim intstr As String = ""
        If int >= 0 Then intstr = CStr(int)
        Dim _CommandResponse As String
        _CommandResponse = SendCommand("Status", cSendMessage, Chr(34) & str & Chr(34) & intstr)
        Return CInt(_CommandResponse)

    End Function


    'GetConnections
    ''' <summary>
    ''' Returns current connection information for the In-Sightvision system.
    ''' </summary>
    ''' <remarks></remarks>
    Function GetConnections() As String
        Dim _CommandResponse As String
        _CommandResponse = SendCommand("User", cGetConnections)
        Return _CommandResponse
    End Function

    ''Get
    ' ''' <summary>
    ' ''' Three commands are used in conjunction with the Get extended Native Mode command to receive information from the In-Sightvision system and its spreadsheet
    ' ''' </summary>
    ' ''' <param name="Command"></param>
    ' ''' <remarks></remarks>
    'Sub GetCommand(Command As String)
    '    SendCommand(cGet, Command)
    'End Sub

    'GetExpr
    ''' <summary>
    ''' Returns the parameters or value stored in the cellspecified by the column and row address, as well as the state of that cell.
    ''' </summary>
    ''' <param name="Column">The column letter of the cell value to set(A to Z)</param>
    ''' <param name="Row">The row number of the cell value to set.</param>
    ''' <remarks></remarks>
    Function GetExpr(Column As String, Row As Integer) As String
        Dim _CommandResponse As String
        _CommandResponse = SendCommand("Expr", cGetExpr, Column & AddLeadingZeros(Row))
        Return _CommandResponse
    End Function

    'Get Filelist
    ''' <summary>
    ''' Returns the number of files stored in memory, andthe name of each file in memory on the In-Sight vision system.
    ''' </summary>
    ''' <remarks></remarks>
    Function GetFilelist() As String
        Dim _CommandResponse As String
        _CommandResponse = SendCommand("", cGetFilelist)
        Return _CommandResponse
    End Function

    'Put Live
    ''' <summary>
    ''' 'Turns liveacquisition mode on or off.
    ''' </summary>
    ''' <param name="int">Enable or Disable, (1 or 0) </param>
    ''' <remarks></remarks>
    Function PutLive(int As Integer) As Integer
        Dim _CommandResponse As String
        _isLive = int
        _CommandResponse = SendCommand("Status", cPutLive, int)
        Return CInt(_CommandResponse)
    End Function

    'Put Update
    ''' <summary>
    ''' Updates the GUI (spreadsheet, cell graphics and image display).
    ''' </summary>
    ''' <remarks></remarks>
    Function PutUpdate() As Integer
        Dim _CommandResponse As String
        _CommandResponse = SendCommand("Status", cPutUpdate)
        Return CInt(_CommandResponse)
    End Function

    'Put Watch
    Enum PutWatchInputs As Integer
        Disable = 0
        EnableOnChange = 1
        EnableEachAquisition = 2
    End Enum
    ''' <summary>
    ''' Returns the value contained in the specified cell each time the cellis updated. The Put Watch command can be used to specify output cellsand send data using the DataChannel.
    ''' </summary>
    ''' <param name="Column">The column letter of the cell value to set(A to Z</param>
    ''' <param name="Row"> The row number of the cell value to set</param>
    ''' <param name="int">•0= disable output of the cell value. 1 = enable output of thecell value only when it changes. 2 = enableoutput of the cell value on every acquisition </param>
    ''' <remarks></remarks>
    Function PutWatch(Column As String, Row As Integer, int As PutWatchInputs) As Integer
        Dim _CommandResponse As String
        _CommandResponse = SendCommand("Status", cPutWatch, Column & AddLeadingZeros(Row) & " " & int)
        Return CInt(_CommandResponse)
    End Function

    'Evalute
    ''' <summary>
    ''' Executes In-Sight functions, as well as inserting formulas into the spreadsheet. Evaluate executes commands for retrieving informationfrom, and making changes to, vision systems
    ''' </summary>
    ''' <param name="Command">Any supported In-Sight function which returns a float, or a legalstring of functions, as well as general commands.Ex) GetString(A2) </param>
    ''' <remarks></remarks>
    Function EvaluateFloat(Command As String) As Double
        Dim _CommandResponse As String
        _CommandResponse = SendCommand("Float", cEvaluate, Command)
        Return CDbl(_CommandResponse)
    End Function
    ''' <summary>
    ''' Executes In-Sight functions, as well as inserting formulas into the spreadsheet. Evaluate executes commands for retrieving informationfrom, and making changes to, vision systems
    ''' </summary>
    ''' <param name="Command">Any supported In-Sight function which returns a float, or a legalstring of functions, as well as general commands.Ex) GetString(A2) </param>
    ''' <remarks></remarks>
    Function Evaluate(Command As String) As Integer
        Dim _CommandResponse As String
        _CommandResponse = SendCommand("Status", cEvaluate, Command)
        Return CInt(_CommandResponse)
    End Function
    ''' <summary>
    ''' The general syntax to remotely place an In-Sightformula into a spreadsheet cell using the Evaluate command is as follows
    ''' </summary>
    ''' <param name="Column">The column letter of the cell value to set(A to Z).</param>
    ''' <param name="Row">The row number of the cell value to set (0 to 399).</param>
    ''' <param name="CellState">Cell Enabled 0 = Disabled 1 = Enabled</param>
    ''' <param name="Formula">A combination of values, functions, arguments,and operators used to create a formula</param>
    ''' <remarks></remarks>
    Function PlaceFormula(Column As String, Row As Integer, CellState As Boolean, Formula As String) As Integer
        Dim _CommandResponse As String
        _CommandResponse = SendCommand("Status", cEvaluate, Column & Row & " " & CellState & " " & Formula)
        Return CInt(_CommandResponse)
    End Function

    'GetCellName
    ''' <summary>
    ''' Returns the cell location of a specified symbolic tag name, or the symbolic tag name of a specified cell
    ''' </summary>
    ''' <param name="Name">The name of the symbolic tag (such as Distance_1.Distance, for example) or the cell location (A4, for example). </param>
    ''' <remarks></remarks>
    Function GetCellName(Name As String) As String
        Dim _CommandResponse As String
        _CommandResponse = SendCommand("String", cGetCellName, "(" & Chr(34) & Name & Chr(34) & ")")
        Return (_CommandResponse)
    End Function

    'Get Cell Value
    ''' <summary>
    ''' Returns the contents of a specified symbolic tag, such as an EasyBuilder fixture output, EasyBuilder Location or Inspection Tool result or Spreadsheet cell, in XML format. 
    ''' </summary>
    ''' <param name="Tag">The name of the symbolic tag, such as a Location or Inspection Tool result ("Distance_1.Distance", for example) or the cell reference (A4, for example). For the automatically generated EasyBuilder fixture output data, the name must be either "Job.Robot.FormatString" or "Job.FormatString", depending upon the selected Device and Protocol selected in the Communications Application Step</param>
    ''' <remarks></remarks>
    Function GetCellValue(Tag As String) As Double
        Dim _CommandResponse As String
        _CommandResponse = SendCommand("Float", cGetCellValue, "(" & Chr(34) & Tag & Chr(34) & ")")
        Return CDbl(_CommandResponse)
    End Function

    'Get Diagnostic Log
    ''' <summary>
    ''' Returns a Diagnostic Log from the Camera
    ''' </summary>
    ''' <remarks></remarks>
    Function GetDiagnosticLog() As String
        Dim _CommandResponse As String
        _CommandResponse = SendCommand("", cGetDiagnosticLog)
        Return (_CommandResponse)
    End Function

    ''' <summary>
    ''' Send a custom command to the camera
    ''' </summary>
    ''' <param name="str">Command to be sent</param>
    ''' <remarks></remarks>
    Function SendMessage(ByVal str As String) As String
        Dim _CommandResponse As String
        _CommandResponse = SendCommand("", str)
        Return (_CommandResponse)
    End Function
    ''' <summary>
    ''' Directly Send a Command Via the Native Client
    ''' </summary>
    ''' <param name="Command">Command to be Send</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function SendDirectNative(ByVal Command As String) As String
        Dim _CommandResponse As String
        _CommandResponse = SendCommand("", Command)
        Return (_CommandResponse)
    End Function
#End Region
#Region "Results"
    ''' <summary>
    ''' Parses Out XML
    ''' </summary>
    ''' <param name="Element">The Element to Extract</param>
    ''' <param name="XMLString">The String of XML to Parse</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ReadXML(ByVal Element As String, ByVal XMLString As String) As String
        Dim DecodedString As String
        Dim MyDoc As New XmlDocument
        MyDoc.LoadXml(XMLString)
        Dim Node As XmlNode
        Node = MyDoc.SelectSingleNode("Response/" & Element)
        If Node Is Nothing Then
            Throw New Exception("XML node, " & Element & ", not found.")
        Else
            DecodedString = Node.InnerText
            Return DecodedString
        End If
    End Function

#End Region

End Class