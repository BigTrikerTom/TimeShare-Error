' ######################################################################
' ## Copyright (c) 2021 TimeShareIt GdbR
' ## by Thomas Steger
' ## File creation Date: 2021-1-29 04:37
' ## File update Date: 2021-8-4 09:38
' ## Filename: Error_Logger.vb (F:\ConDrop\ConDrop_Server\Error_Logger.vb)
' ## Project: ConDrop_Server
' ## Last User: stegert
' ######################################################################
'
'

Option Strict On

Imports log4net
'Imports log4net.Repository
Imports log4net.Repository.Hierarchy
Imports log4net.Core
Imports log4net.Appender
Imports log4net.Layout
Imports System.IO
'Imports Microsoft.Win32
'Imports System.Reflection
'Imports System.Runtime.CompilerServices
Imports log4net.Filter
Imports System.Windows.Forms

' --------------------------------------------------------------------------------------------------------
' in modMain -> sub Main() einfügen und anpassen:
' 
'' ********************* Logging starten *********************
'Error_Logger.DbTableSettings = db_table_condrop_Settings
'Error_Logger.DatabaseName = Helper_DBconnectLocal.SelectDatabase.ConDrop
'Call Error_Logger.InitializeLogger()
'
'
'
' in app.config nach <configSections> einfügen:
'
'<appSettings>
'   <add key = "log4net.Internal.Debug" value="false"/>
'</appSettings>
'
' --------------------------------------------------------------------------------------------------------
''' <summary>
''' Klasse für Logging
''' </summary>
''' <remarks></remarks>
Public Class Error_Logger
    Private Shared ReadOnly log As log4net.ILog = LogManager.GetLogger(Application.ProductName & "." & Environment.MachineName.ToUpper)
    'Private Shared ReadOnly log As log4net.ILog = LogManager.GetLogger(ApplicationProductName & "." & Environment.MachineName.ToUpper)
#Region "Enum"
    Public Enum MinLogLevel
        Log_ALL
        Log_DEBUG
        Log_INFO
        Log_WARN
        Log_ERROR
        Log_FATAL
        Log_OFF
    End Enum

#End Region
#Region "Properties"
    Private Shared _ApplicationProductName As String
    Public Shared Property ApplicationProductName() As String
        Get
            Return _ApplicationProductName
        End Get
        Set(ByVal value As String)
            _ApplicationProductName = value
        End Set
    End Property
    Private Shared _dbTableSettings As String
    Public Shared Property DbTableSettings() As String
        Get
            Return _dbTableSettings
        End Get
        Set(ByVal value As String)
            _dbTableSettings = value
        End Set
    End Property
    Private Shared _databaseName As Error_DbConnectLocal.SelectDatabase
    Public Shared Property DatabaseName() As Error_DbConnectLocal.SelectDatabase
        Get
            Return _databaseName
        End Get
        Set(ByVal value As Error_DbConnectLocal.SelectDatabase)
            _databaseName = value
        End Set
    End Property

#End Region

    Public Shared Sub InitializeLogger(ByVal LogLevel As Level)
        Dim SetLogLevel As String = "DEBUG"
        Dim LogSize As String = "10MB"
        Dim Log2Text As Boolean = False
        Dim Log2XML As Boolean = False
        Dim Log2MySQL As Boolean = True
        ''Dim LogLevel As New Level(0, "All")

        'SetLogLevel = Convert.ConvertToString(HelperDB.ReadSettingsFromDB(ReadWriteSettingForApp, "Log4Net", "MinLogLevel", DatabaseName, DbTableSettings).Value)
        'LogSize = Convert.ConvertToString(HelperDB.ReadSettingsFromDB(ReadWriteSettingForApp, "Log4Net", "LogSize", DatabaseName, DbTableSettings).Value)
        'Log2Text = Convert.ConvertToBoolean(HelperDB.ReadSettingsFromDB(ReadWriteSettingForApp, "Log4Net", "Log2Text", DatabaseName, DbTableSettings).Value)
        'Log2XML = Convert.ConvertToBoolean(HelperDB.ReadSettingsFromDB(ReadWriteSettingForApp, "Log4Net", "Log2XML", DatabaseName, DbTableSettings).Value)
        'Log2MySQL = Convert.ConvertToBoolean(HelperDB.ReadSettingsFromDB(ReadWriteSettingForApp, "Log4Net", "Log2MySQL", DatabaseName, DbTableSettings).Value)
        Log2MySQL = True


        If Log2MySQL Then
            Error_Logger.ConfigureWithDb("Server=rt01.logonme.de; Database=logging; Uid=loggingdbu; Pwd=retCEHECKTIR20A;IgnorePrepare=true;", False)
        End If
        'If Log2XML Then
        '    Error_Logger.ConfigureWithXml(LogSize, Level.Info)
        'End If
        'If Log2Text Then
        '    Error_Logger.ConfigureWithFile(LogSize, Level.Error)
        'End If
        LogLevel = GetLogLevelInteger(SetLogLevel)
        Call Error_Logger.SetLogingLevel(LogLevel)



        Dim architecture As String = ""
        If CInt(IntPtr.Size) > 4 Then
            architecture = "64 Bit"
        Else
            architecture = "32 Bit"
        End If
        Call Error_Logger.writelog(Level.Info, "Application " & Application.ProductName & " Start")
        Call Error_Logger.writelog(Level.Info, "Programmversion: " & String.Format("{0}", My.Application.Info.Version.ToString) & " (" & architecture & ")")
        Call Error_Logger.writelog(Level.Notice, "Der LogLevel wurde initial auf " & LogLevel.DisplayName & " gesetzt.")

    End Sub
    Public Shared Function GetLogLevelInteger(SetLogLevel As Level) As Level
        Dim LogLevel As New Level(0, "All")
        If SetLogLevel = Level.All Then
            LogLevel = Level.All
            LogManager.GetRepository().Threshold = Level.All
        ElseIf SetLogLevel = Level.Debug Then
            LogLevel = Level.Debug
            LogManager.GetRepository().Threshold = Level.Debug
        ElseIf SetLogLevel = Level.Info Then
            LogLevel = Level.Info
            LogManager.GetRepository().Threshold = Level.Info
        ElseIf SetLogLevel = Level.Warn Then
            LogLevel = Level.Warn
            LogManager.GetRepository().Threshold = Level.Warn
        ElseIf SetLogLevel = Level.Error Then
            LogLevel = Level.Error
            LogManager.GetRepository().Threshold = Level.Error
        ElseIf SetLogLevel = Level.Fatal Then
            LogLevel = Level.Fatal
            LogManager.GetRepository().Threshold = Level.Fatal
        ElseIf SetLogLevel = Level.Off Then
            LogLevel = Level.Off
            LogManager.GetRepository().Threshold = Level.Off
        Else
            LogLevel = Level.All
            LogManager.GetRepository().Threshold = Level.All
        End If
        Return LogLevel

    End Function
    Public Shared Function GetLogLevelInteger(SetLogLevel As String) As Level
        Dim LogLevel As New Level(0, "All")
        If SetLogLevel = "ALL" Then
            LogLevel = Level.All
            LogManager.GetRepository().Threshold = Level.All
        ElseIf SetLogLevel = "DEBUG" Then
            LogLevel = Level.Debug
            LogManager.GetRepository().Threshold = Level.Debug
        ElseIf SetLogLevel = "INFO" Then
            LogLevel = Level.Info
            LogManager.GetRepository().Threshold = Level.Info
        ElseIf SetLogLevel = "WARN" Then
            LogLevel = Level.Warn
            LogManager.GetRepository().Threshold = Level.Warn
        ElseIf SetLogLevel = "ERROR" Then
            LogLevel = Level.Error
            LogManager.GetRepository().Threshold = Level.Error
        ElseIf SetLogLevel = "FATAL" Then
            LogLevel = Level.Fatal
            LogManager.GetRepository().Threshold = Level.Fatal
        ElseIf SetLogLevel = "OFF" Then
            LogLevel = Level.Off
            LogManager.GetRepository().Threshold = Level.Off
        Else
            LogLevel = Level.All
            LogManager.GetRepository().Threshold = Level.All
        End If
        Return LogLevel
    End Function
    Public Shared Sub SetUpDbConnection(ByVal connectionString As String, ByVal logConfig As String)
        Dim hier As Hierarchy = TryCast(LogManager.GetRepository(), Hierarchy)
        log4net.Config.XmlConfigurator.ConfigureAndWatch(New FileInfo(logConfig))

        If hier IsNot Nothing Then
            Dim adoNetAppenders As IEnumerable(Of AdoNetAppender) = hier.GetAppenders().OfType(Of AdoNetAppender)()
            For Each AdoNetAppender As AdoNetAppender In adoNetAppenders
                AdoNetAppender.ConnectionString = connectionString
                AdoNetAppender.ActivateOptions()
            Next
        End If
    End Sub
    Public Shared Sub SetLogingLevel(ByVal LogLevel As Level)
        Dim repositories As log4net.Repository.ILoggerRepository() = log4net.LogManager.GetAllRepositories()
        For Each repository As log4net.Repository.ILoggerRepository In repositories
            repository.Threshold = LogLevel

            Dim hier As log4net.Repository.Hierarchy.Hierarchy = CType(repository, log4net.Repository.Hierarchy.Hierarchy)
            Dim loggers As log4net.Core.ILogger() = hier.GetCurrentLoggers()
            For Each logger As log4net.Core.ILogger In loggers
                logger.Repository.Threshold = LogLevel
            Next
        Next

        Dim h As log4net.Repository.Hierarchy.Hierarchy = CType(log4net.LogManager.GetRepository(), log4net.Repository.Hierarchy.Hierarchy)
        Dim rootLogger As log4net.Repository.Hierarchy.Logger = h.Root
        rootLogger.Level = LogLevel
    End Sub


    Public Shared Sub TestLog4Net()
        If Not log4net.LogManager.GetRepository().Configured Then

            For Each LogMessage As log4net.Util.LogLog In log4net.LogManager.GetRepository().ConfigurationMessages
                '.Cast < log4net.Util.LogLog()
                'Debug.Print(LogMessage.Message)
            Next
        End If
    End Sub

    Public Shared Sub writelog(ByVal loglevel As Level,
                               ByVal LogText As String,
                               Optional ByVal Ex As Exception = Nothing,
                               Optional ByVal WriteDebug As Boolean = False)
        'log.Debug(loglevel.Value & " - " & loglevel.DisplayName & " - " & loglevel.Name & " - " & LogText)
        If loglevel.Value = Level.Info.Value Then
            log.Info(LogText)
        ElseIf loglevel.Value = Level.Error.Value Then
            If Ex Is Nothing Then
                log.Error(LogText)
            Else
                log.Error(LogText, Ex)
            End If
        ElseIf loglevel.Value = Level.Fatal.Value Then
            If Ex Is Nothing Then
                log.Fatal(LogText)
            Else
                log.Fatal(LogText, Ex)
            End If
        ElseIf loglevel.Value = Level.Debug.Value Then
            log.Debug(LogText)
        ElseIf loglevel.Value = Level.Warn.Value Then
            log.Warn(LogText)
        ElseIf loglevel.Value = Level.Notice.Value Then
            log.Notice(LogText)
        Else
            Call Error_Logger.writelog(Level.Info, LogText)
        End If
        If WriteDebug Then
            Debug.Write(LogText)
        End If
    End Sub

#Region "Appender"
    Public Shared Function CreateRollingXmlAppender(ByVal MaxLogSize As String) As IAppender
        Dim XmlLayout As New XmlLayoutSchemaLog4j
        XmlLayout.LocationInfo = True
        Dim xmlroller As New RollingFileAppender()
        xmlroller.AppendToFile = True
        xmlroller.File = "Logs\Logging.xml"
        xmlroller.Layout = XmlLayout
        xmlroller.Encoding = System.Text.Encoding.UTF8
        xmlroller.DatePattern = "yyyyMMdd"
        xmlroller.MaxSizeRollBackups = 10
        xmlroller.MaximumFileSize = MaxLogSize
        xmlroller.RollingStyle = RollingFileAppender.RollingMode.Date
        xmlroller.StaticLogFileName = True
        xmlroller.LockingModel = New FileAppender.MinimalLock()
        xmlroller.ActivateOptions()
        Return xmlroller
    End Function
    Public Shared Function CreateFileAppender(ByVal MaxLogSize As String) As IAppender
        Dim patternLayout As New PatternLayout()
        patternLayout.ConversionPattern = "%date [%thread] %-5level %logger - %message%newline"
        patternLayout.ActivateOptions()
        Dim roller As New RollingFileAppender()
        roller.AppendToFile = True
        roller.File = "TextLogs\Logging.txt"
        roller.Layout = patternLayout
        roller.Encoding = System.Text.Encoding.UTF8
        roller.MaxSizeRollBackups = 10
        roller.MaximumFileSize = MaxLogSize
        roller.RollingStyle = RollingFileAppender.RollingMode.Size
        roller.StaticLogFileName = True
        roller.LockingModel = New FileAppender.MinimalLock()
        roller.ActivateOptions()
        Return roller
    End Function
    Public Shared Function CreateConsoleAppender() As IAppender
        Dim appender As ConsoleAppender = New ConsoleAppender()
        appender.Name = "ConsoleAppender"
        Dim layout As PatternLayout = New PatternLayout()
        layout.ConversionPattern = "%newline%date %-5level %logger – %message – %property%newline"
        layout.ActivateOptions()
        appender.Layout = layout
        appender.ActivateOptions()
        Return appender
    End Function
    Public Shared Function CreateAdoNetAppender(ByVal ConnectionString As String) As IAppender
        Dim architecture As String = ""
        Dim Appender As AdoNetAppender = New AdoNetAppender()
        Appender.Name = "AdoNetAppender"
        Appender.BufferSize = 1
        Appender.ConnectionType = "MySql.Data.MySqlClient.MySqlConnection, MySql.Data"
        Appender.ConnectionString = ConnectionString
        Appender.ReconnectOnError = True

        Appender.CommandText = "INSERT INTO errorlog (Date,hostname,Thread,Level,Logger,Message,Method,Exception,stacktrace,Context,line,appdomain,username,location,appname,appversion,architecture) VALUES (?log_date?, ?hostname?, ?thread?, ?log_level?, ?logger?, ?message?, ?method_name?, ?exception?, ?stacktrace?, ?context?, ?line?, ?appdomain?, ?username?, ?location?,?appname?,?appversion?,?architecture?)"

        Error_Extensions.AddDateTimeParameterToAppender(Appender, "?log_date?")
        Error_Extensions.AddStringParameterToAppender(Appender, "?hostname?", 255, "%property{log4net:HostName}")
        Error_Extensions.AddStringParameterToAppender(Appender, "?thread?", 32, "%t")
        Error_Extensions.AddStringParameterToAppender(Appender, "?log_level?", 512, "%p")
        Error_Extensions.AddStringParameterToAppender(Appender, "?logger?", 512, "%c")
        Error_Extensions.AddStringParameterToAppender(Appender, "?method_name?", 200, "%method")
        Error_Extensions.AddStringParameterToAppender(Appender, "?message?", 1000, "%m")
        Error_Extensions.AddErrorParameterToAppender(Appender, "?exception?", 4000)
        Error_Extensions.AddStringParameterToAppender(Appender, "?stacktrace?", 4000, "%stacktrace")
        Error_Extensions.AddStringParameterToAppender(Appender, "?context?", 512, "%x")
        Error_Extensions.AddStringParameterToAppender(Appender, "?appdomain?", 512, "%appdomain")
        Error_Extensions.AddInt32ParameterToAppender(Appender, "?line?", "%L")
        Error_Extensions.AddStringParameterToAppender(Appender, "?username?", 75, "%username")
        Error_Extensions.AddStringParameterToAppender(Appender, "?location?", 512, "%location")

        If CInt(IntPtr.Size) > 4 Then
            architecture = "64 Bit"
        Else
            architecture = "32 Bit"
        End If
        Error_Extensions.AddStringParameterToAppender(Appender, "?appname?", 512, My.Application.Info.AssemblyName & " (" & architecture & ")")
        Error_Extensions.AddStringParameterToAppender(Appender, "?appversion?", 512, My.Application.Info.Version.ToString)
        Error_Extensions.AddStringParameterToAppender(Appender, "?architecture?", 512, architecture)

        Appender.ActivateOptions()
        Return Appender
    End Function

#End Region

#Region "ConfigureLogger"
    Public Shared Sub ConfigureWithFile(ByVal MaxLogSize As String, ByVal MinLogLevel As Level)
        Dim h As Hierarchy = CType(LogManager.GetRepository(), Hierarchy)
        h.Root.Level = MinLogLevel
        h.Root.AddAppender(CreateFileAppender(MaxLogSize))
        h.Configured = True
    End Sub
    Public Shared Sub ConfigureWithXml(ByVal MaxLogSize As String, ByVal MinLogLevel As Level)
        Dim h As Hierarchy = CType(LogManager.GetRepository(), Hierarchy)
        h.Root.Level = MinLogLevel
        h.Root.AddAppender(CreateRollingXmlAppender(MaxLogSize))
        h.Configured = True
    End Sub
    Public Shared Sub ConfigureWithDb(ByVal ConnectionString As String, ByVal onlyErrors As Boolean)
        Dim h As Hierarchy = CType(LogManager.GetRepository(), Hierarchy)
        h.Root.Level = Level.All
        Dim ado As IAppender = CreateAdoNetAppender(ConnectionString)
        h.Root.AddAppender(ado)

        If onlyErrors Then
            Dim filter As New LevelRangeFilter()
            filter.LevelMin = Level.Error
            CType(ado, AppenderSkeleton).AddFilter(filter)
        End If

        h.Configured = True
    End Sub

#End Region


End Class

