#Region "Imports"
Imports System.Net
Imports System.IO
Imports System.Text
Imports Newtonsoft.Json
Imports PowerwallService.PWJson
Imports PowerwallService.SolCast
Imports PowerwallService.PVOutput
Imports PowerwallService.SunriseSunset
Imports PowerwallService.PowerBIStreaming
Imports PowerwallService.WeatherUndergroundForecast
Imports PowerwallService.WillyWeatherForecast
Imports PowerwallService.TariffDefinition
Imports PowerwallService.Maps
#End Region
Public Class PowerwallService
#Region "Variables"
    Private ObsTA As New PWHistoryDataSetTableAdapters.observationsTableAdapter
    Private SolarTA As New PWHistoryDataSetTableAdapters.solarTableAdapter
    Private SOCTA As New PWHistoryDataSetTableAdapters.socTableAdapter
    Private BatteryTA As New PWHistoryDataSetTableAdapters.batteryTableAdapter
    Private LoadTA As New PWHistoryDataSetTableAdapters.loadTableAdapter
    Private SiteTA As New PWHistoryDataSetTableAdapters.siteTableAdapter
    Private CompactTA As New PWHistoryDataSetTableAdapters.CompactObsTableAdapter
    Private CompactTALocal As New PWHistoryDataSetTableAdapters.CompactObsTALocal
    Private Get5MinuteAveragesTA As New PWHistoryDataSetTableAdapters.spGet5MinuteAveragesTableAdapter
    Private SPs As New PWHistoryDataSetTableAdapters.SPs
    Private ObservationTimer As New Timers.Timer
    Private OneMinuteTimer As New Timers.Timer
    Private ReportingTimer As New Timers.Timer
    Private ForecastTimer As New Timers.Timer
    Private FiveBeforeTheHourTimer As New Timers.Timer
    Private PWToken As String = String.Empty
    Shared NextDayForecastGeneration As Single = 0
    Shared NextDayMorningGeneration As Single = 0
    Shared MeterReading As MeterAggregates
    Shared SOC As SOC
    Shared CurrentDayForecast As DayForecast
    Shared NextDayForecast As DayForecast
    Shared SecondDayForecast As DayForecast
    Shared ForecastsRetrieved As Date = DateAdd(DateInterval.Hour, -2, Now)
    Shared FirstReadingsAvailable As Boolean = False
    Shared PreCharging As Boolean = False
    Shared OnStandby As Boolean = False
    Shared AboveMinBackup As Boolean = False
    Shared AutonomousMode As Boolean = False
    Shared LastTarget As Integer = 0
    Shared PWTarget As Decimal = 0
    Shared OffPeakStart As DateTime
    Shared PeakStart As DateTime
    Shared OffPeakStartHour As Integer
    Shared PeakStartHour As Integer
    Shared OffPeakHours As Double
    Shared OperationLockout As DateTime = DateAdd(DateInterval.Hour, -2, Now)
    Shared CurrentDayAllOffPeak As Boolean
    Shared NextDayAllDayOffPeak As Boolean
    Shared IsCharging As Boolean = False
    Shared LastPeriodForecast As Forecast
    Shared CurrentPeriodForecast As Forecast
    Shared PVForecast As OutputForecast
    Shared CurrentDOW As DayOfWeek
    Shared CivilTwilightSunrise As DateTime
    Shared CivilTwilightSunset As DateTime
    Shared Sunrise As DateTime
    Shared Sunset As DateTime
    Shared AsAtSunrise As Result
    Shared DBLock As New Object
    Shared PWLock As New Object
    Shared SkipObservation As Boolean = False
    Shared TariffDefinition As Tariff
    Shared TariffMap As List(Of TariffPart)
    Shared NearTermPeriods As List(Of NextDayPart)
    Shared TempFCasts As List(Of TemperaturePart)
    Shared CurrentTariff As NextDayPart
    Shared NextTariff As NextDayPart
#End Region
#Region "Timer Handlers"
    Protected Async Sub OnObservationTimer(Sender As Object, Args As Timers.ElapsedEventArgs)
        Await Task.Run(Sub()
                           GetObservationAndStore()
                       End Sub)
    End Sub
    Protected Async Sub OnOneMinuteTime(Sender As Object, Args As Timers.ElapsedEventArgs)
        Await Task.Run(Sub()
                           DoPerMinuteTasks()
                       End Sub)
    End Sub
    Protected Async Sub OnReportingTimer(Sender As Object, Args As Timers.ElapsedEventArgs)
        Await Task.Run(Sub()
                           If My.Settings.PVReportingEnabled Then
                               If My.Settings.PVSendPowerwall Then
                                   SendPowerwallData(Now)
                                   If Now.Minute < 6 Then
                                       If Now.Hour = 0 Then
                                           DoBackFill(DateAdd(DateInterval.Day, -1, Now))
                                       Else
                                           DoBackFill(Now)
                                       End If
                                   End If
                               ElseIf My.Settings.PVSendForecast Then
                                   SendForecast()
                               End If
                           End If
                       End Sub)
    End Sub
    Protected Async Sub OnForecastTimer(Sender As Object, Args As Timers.ElapsedEventArgs)
        Await Task.Run(Sub()
                           CheckSOCLevel()
                       End Sub)
    End Sub
    Protected Async Sub OnFiveBeforeTheHourTimer(Sender As Object, Args As Timers.ElapsedEventArgs)
        Await Task.Run(Sub()
                           DoHourlyTasks(Now)
                       End Sub)
    End Sub
    Protected Async Sub DebugTask()
        Await Task.Run(Sub()
                           LoadTariffConfig()
                           DoHourlyTasks(Now)
                           DoAsyncStartupProcesses()
                       End Sub)
    End Sub
#End Region
#Region "Service Methods"
    Public Sub New()
        MyBase.New()
        CanPauseAndContinue = True
        AutoLog = True
        InitializeComponent()
    End Sub
    Protected Overrides Sub OnStart(ByVal args() As String)
        EventLog.WriteEntry("Powerwall Service Starting", EventLogEntryType.Information, 100)

        ObservationTimer.Interval = 6 * 1000 ' Every Six Seconds
        ObservationTimer.AutoReset = True
        AddHandler ObservationTimer.Elapsed, AddressOf OnObservationTimer

        OneMinuteTimer.Interval = 60 * 1000 ' Every Minute
        OneMinuteTimer.AutoReset = True
        AddHandler OneMinuteTimer.Elapsed, AddressOf OnOneMinuteTime

        If My.Settings.PVReportingEnabled Then
            ReportingTimer.Interval = 5 * 60 * 1000 ' Every Five Minutes
            ReportingTimer.AutoReset = True
            AddHandler ReportingTimer.Elapsed, AddressOf OnReportingTimer
        End If

        If My.Settings.PWControlEnabled Then
            ForecastTimer.Interval = 60 * 10 * 1000 ' Every 10 Minutes
            ForecastTimer.AutoReset = True
            AddHandler ForecastTimer.Elapsed, AddressOf OnForecastTimer
        End If

        FiveBeforeTheHourTimer.Interval = 60 * 60 * 1000 ' Every Hour
        FiveBeforeTheHourTimer.AutoReset = True
        AddHandler FiveBeforeTheHourTimer.Elapsed, AddressOf OnFiveBeforeTheHourTimer

        Task.Run(Sub()
                     DoAsyncStartupProcesses()
                 End Sub)

        EventLog.WriteEntry("Powerwall Service Started", EventLogEntryType.Information, 101)
    End Sub
    Private Sub DoAsyncStartupProcesses()
        If My.Settings.PWForceModeOnStartup Then
            Dim RunningResult As Integer
            Try
                Dim ChargeSettings As Operation
                Dim NewChargeSettings As Operation
                Dim APIResult As Integer
                SyncLock PWLock
                    ChargeSettings = New Operation With {.backup_reserve_percent = My.Settings.PWForceBackupPercentage, .real_mode = My.Settings.PWForceMode}
                    NewChargeSettings = PostPWSecureAPISettings(Of Operation)("operation", ChargeSettings, ForceReLogin:=True)
                    APIResult = GetPWSecure("config/completed")
                    RunningResult = GetPWRunning()
                End SyncLock
                If APIResult = 202 Then
                    EventLog.WriteEntry(String.Format("Set PW Mode: Mode={0}, BackupPercentage={1}, APIResult = {2}", NewChargeSettings.real_mode, NewChargeSettings.backup_reserve_percent, APIResult), EventLogEntryType.Information, 600)
                Else
                    EventLog.WriteEntry(String.Format("Failed to Set PW Mode: Mode={0}, BackupPercentage={1}, APIResult = {2}", NewChargeSettings.real_mode, NewChargeSettings.backup_reserve_percent, APIResult), EventLogEntryType.Warning, 601)
                End If
            Catch ex As Exception
                SyncLock PWLock
                    RunningResult = GetPWRunning()
                End SyncLock
            End Try
        End If
        LoadTariffConfig()
        BuildTariffMap()
        If My.Settings.WWGetWeather Then GetWWTemperature()
        If My.Settings.WUGetWeather Then GetWUTemperature()
        SetOffPeakHours(Now)
        'SetTariffHours(Now)
        GetForecasts()
        GetPWMode()
        Task.Run(Sub()
                     SleepUntilSecBoundary(6)
                     ObservationTimer.Start()
                     EventLog.WriteEntry("Observation Timer Started", EventLogEntryType.Information, 108)
                 End Sub)

        Task.Run(Sub()
                     SleepUntilSecBoundary(60)
                     OneMinuteTimer.Start()
                     DoPerMinuteTasks()
                     EventLog.WriteEntry("One Minute Timer Started", EventLogEntryType.Information, 109)
                 End Sub)
    End Sub
    Protected Overrides Sub OnContinue()
        EventLog.WriteEntry("Powerwall Service Resuming", EventLogEntryType.Information, 102)
        ObservationTimer.Start()
        OneMinuteTimer.Start()
        Task.Run(Sub()
                     DebugTask()
                 End Sub
                 )
        EventLog.WriteEntry("Powerwall Service Running", EventLogEntryType.Information, 103)
    End Sub
    Protected Overrides Sub OnPause()
        EventLog.WriteEntry("Powerwall Service Pausing", EventLogEntryType.Information, 104)
        ObservationTimer.Stop()
        OneMinuteTimer.Stop()
        If My.Settings.PVReportingEnabled Then ReportingTimer.Stop()
        If My.Settings.PWControlEnabled Then ForecastTimer.Stop()
        EventLog.WriteEntry("Powerwall Service Paused", EventLogEntryType.Information, 105)
    End Sub
    Protected Overrides Sub OnStop()
        EventLog.WriteEntry("Powerwall Service Stopping", EventLogEntryType.Information, 106)
        ObservationTimer.Stop()
        ObservationTimer.Dispose()
        OneMinuteTimer.Stop()
        OneMinuteTimer.Dispose()
        If My.Settings.PVReportingEnabled Then
            ReportingTimer.Stop()
            ReportingTimer.Dispose()
        End If
        ForecastTimer.Stop()
        ForecastTimer.Dispose()
        EventLog.WriteEntry("Powerwall Service Stopped", EventLogEntryType.Information, 107)
    End Sub
#End Region
#Region "Miscellaneous Helpers"
    Sub SleepUntilSecBoundary(Boundary As Integer)
        Dim LastObs As Date = Now
        Dim MilliSeconds As Integer = LastObs.Millisecond
        Dim SecondOffset As Integer = Boundary - LastObs.Second Mod Boundary
        Threading.Thread.Sleep((SecondOffset * 1000) - MilliSeconds)
    End Sub
    Sub AggregateToMinute()
        If My.Settings.DebugLogging Then EventLog.WriteEntry(String.Format("Powerwall Service Running at {0:yyyy-MM-dd HH:mm}", Now), EventLogEntryType.Information, 200)
        If My.Settings.LogData Then
            Try
                SyncLock DBLock
                    SPs.spAggregateToMinute()
                End SyncLock
            Catch Ex As Exception
                EventLog.WriteEntry(Ex.Message & vbCrLf & vbCrLf & Ex.StackTrace, EventLogEntryType.Error)
            End Try
        End If
    End Sub
    Public Sub SetOffPeakHours(InvokedTime As DateTime)
        Dim TZI As TimeZoneInfo = TimeZoneInfo.Local
        If My.Settings.DebugLogging Then EventLog.WriteEntry(String.Format("SetOffPeakHours Invoked at {0:yyyy-MM-dd HH:mm}", InvokedTime), EventLogEntryType.Information, 700)
        CurrentDOW = InvokedTime.DayOfWeek
        If My.Settings.DebugLogging Then EventLog.WriteEntry(String.Format("SetOffPeakHours DOW {0}", CurrentDOW), EventLogEntryType.Information, 701)
        CurrentDayAllOffPeak = False
        If (CurrentDOW = DayOfWeek.Saturday Or CurrentDOW = DayOfWeek.Sunday) Then
            If My.Settings.DebugLogging Then EventLog.WriteEntry("In Weekend", EventLogEntryType.Information, 702)
            If My.Settings.TariffPeakOnWeekends Then
                OffPeakStartHour = My.Settings.TariffPeakEndWeekend
                PeakStartHour = My.Settings.TariffPeakStartWeekend
            Else
                OffPeakStartHour = 0
                PeakStartHour = 0
                CurrentDayAllOffPeak = True
            End If
        Else
            OffPeakStartHour = My.Settings.TariffPeakEndWeekday
            PeakStartHour = My.Settings.TariffPeakStartWeekday
        End If
        If My.Settings.TariffIgnoresDST And TZI.IsDaylightSavingTime(InvokedTime) And (Not (CurrentDOW = DayOfWeek.Saturday Or CurrentDOW = DayOfWeek.Sunday) Or My.Settings.TariffPeakOnWeekends) Then
            OffPeakStartHour += CByte(1)
            If OffPeakStartHour > 23 Then OffPeakStartHour -= CByte(24)
            PeakStartHour += CByte(1)
            If PeakStartHour > 23 Then PeakStartHour -= CByte(24)
        End If
        If My.Settings.DebugLogging Then EventLog.WriteEntry(String.Format("Off Peak Hours From {0} To {1}", OffPeakStartHour, PeakStartHour), EventLogEntryType.Information, 703)
        Dim OffPeakStartsBeforeMidnight As Boolean = ((OffPeakStartHour > PeakStartHour) Or (OffPeakStartHour = 0 And PeakStartHour = 0))
        OffPeakHours = PeakStartHour - OffPeakStartHour
        If OffPeakHours <= 0 Then OffPeakHours += 24
        If My.Settings.DebugLogging Then EventLog.WriteEntry(String.Format("Adjusted Off Peak Duration {0}", OffPeakHours), EventLogEntryType.Information, 706)
        OffPeakStart = New DateTime(InvokedTime.Year, InvokedTime.Month, InvokedTime.Day, OffPeakStartHour, 0, 0)
        If My.Settings.DebugLogging Then EventLog.WriteEntry(String.Format("Initial Off Peak Start {0:yyyy-MM-dd HH:mm}", OffPeakStart), EventLogEntryType.Information, 707)
        If InvokedTime.Hour > PeakStartHour And Not OffPeakStartsBeforeMidnight Then
            OffPeakStart = DateAdd(DateInterval.Day, 1, OffPeakStart)
            If My.Settings.DebugLogging Then EventLog.WriteEntry(String.Format("Adjusted Off Peak Start {0:yyyy-MM-dd HH:mm}", OffPeakStart), EventLogEntryType.Information, 708)
        End If
        If InvokedTime.Hour >= 0 And InvokedTime.Hour < PeakStartHour + 1 And OffPeakStartsBeforeMidnight Then
            OffPeakStart = DateAdd(DateInterval.Day, -1, OffPeakStart)
            If My.Settings.DebugLogging Then EventLog.WriteEntry(String.Format("Adjusted Off Peak Start {0:yyyy-MM-dd HH:mm}", OffPeakStart), EventLogEntryType.Information, 709)
        End If
        PeakStart = New DateTime(OffPeakStart.Year, OffPeakStart.Month, OffPeakStart.Day, PeakStartHour, 0, 0)
        If My.Settings.DebugLogging Then EventLog.WriteEntry(String.Format("Initial Off Peak End {0:yyyy-MM-dd HH:mm}", PeakStart), EventLogEntryType.Information, 710)
        If OffPeakStartsBeforeMidnight Then
            PeakStart = DateAdd(DateInterval.Day, 1, PeakStart)
            If My.Settings.DebugLogging Then EventLog.WriteEntry(String.Format("Adjusted Off Peak End {0:yyyy-MM-dd HH:mm}", PeakStart), EventLogEntryType.Information, 711)
        End If
        If (PeakStart.DayOfWeek = DayOfWeek.Saturday Or PeakStart.DayOfWeek = DayOfWeek.Sunday) And Not My.Settings.TariffPeakOnWeekends Then
            If My.Settings.DebugLogging Then EventLog.WriteEntry("End of Off Peak is All Off Peak Weekend", EventLogEntryType.Information, 714)
            NextDayAllDayOffPeak = True
        Else
            NextDayAllDayOffPeak = False
        End If
        If Sunrise.Date <> InvokedTime.Date Then
            Dim SunriseSunsetData As Result = GetSunriseSunset(Of Result)(InvokedTime)
            With SunriseSunsetData.results
                CivilTwilightSunrise = .civil_twilight_begin.ToLocalTime
                CivilTwilightSunset = .civil_twilight_end.ToLocalTime
                Sunrise = .sunrise.ToLocalTime
                Sunset = .sunset.ToLocalTime
            End With
        End If
        If My.Settings.DebugLogging Then EventLog.WriteEntry(String.Format("Start: {0:yyyy-MM-dd HH:mm} End: {1:yyyy-MM-dd HH:mm}", OffPeakStart, PeakStart), EventLogEntryType.Information, 713)
    End Sub
    Public Sub SetTariffHours(InvokedTime As DateTime)
        Dim TodayDOW As DayOfWeek = InvokedTime.DayOfWeek
        Dim Tomorrow As Date = DateAdd(DateInterval.Day, 1, InvokedTime)
        Dim TomorrowDOW As DayOfWeek = Tomorrow.DayOfWeek
        Dim TZI As TimeZoneInfo = TimeZoneInfo.Local
        Dim DSTOffset As Integer = 0
        If My.Settings.TariffIgnoresDST And TZI.IsDaylightSavingTime(InvokedTime) Then
            DSTOffset = 1
        End If

        CurrentDOW = InvokedTime.DayOfWeek

        NearTermPeriods = New List(Of NextDayPart)

        ' Get the data for the remainder of today
        Dim TMEs = From TMs In TariffMap
                   Join TFs In TempFCasts
                       On TFs.PartHour Equals TMs.Hour
                   Where TMs.DOW = CurrentDOW And
                       TFs.PartDate = InvokedTime.Date And
                       TMs.Hour >= InvokedTime.Hour - 1
                   Order By TMs.Hour
                   Select TFs.PartDate, TMs.DOW, TMs.Hour, TMs.Cooling, TMs.Heating, TMs.ConsumptionRate, TMs.LoadPercentage, TMs.FITPriority, TMs.FITRate, TMs.OffsetPriority, TMs.StandbyPreferred, TMs.Tariff, TFs.Temperature, TMs.IsOffPeak
        For Each TME In TMEs
            NearTermPeriods.Add(New NextDayPart With {.PartDate = TME.PartDate, .PartHour = TME.Hour + DSTOffset, .PartDOW = TME.DOW, .Cooling = TME.Cooling, .Heating = TME.Heating, .ConsumptionRate = TME.ConsumptionRate, .LoadPercentage = TME.LoadPercentage, .FITPriority = TME.FITPriority, .FITRate = TME.FITRate, .OffsetPriority = TME.OffsetPriority, .StandbyPreferred = TME.StandbyPreferred, .Tariff = TME.Tariff, .Temperature = TME.Temperature, .CDD = CDec(IIf(TME.Temperature < My.Settings.CDDBase, (My.Settings.CDDBase - TME.Temperature) / 24.0, 0)), .HDD = CDec(IIf(TME.Temperature > My.Settings.HDDBase, (TME.Temperature - My.Settings.HDDBase) / 24.0, 0)), .IsOffPeak = TME.IsOffPeak})
        Next

        ' Get the data for tomorrow
        TMEs = From TMs In TariffMap
               Join TFs In TempFCasts
                       On TFs.PartHour Equals TMs.Hour
               Where TMs.DOW = TomorrowDOW And
                     TFs.PartDate = Tomorrow.Date
               Order By TMs.Hour
               Select TFs.PartDate, TMs.DOW, TMs.Hour, TMs.Cooling, TMs.Heating, TMs.ConsumptionRate, TMs.LoadPercentage, TMs.FITPriority, TMs.FITRate, TMs.OffsetPriority, TMs.StandbyPreferred, TMs.Tariff, TFs.Temperature, TMs.IsOffPeak
        For Each TME In TMEs
            NearTermPeriods.Add(New NextDayPart With {.PartDate = TME.PartDate, .PartHour = TME.Hour + DSTOffset, .PartDOW = TME.DOW, .Cooling = TME.Cooling, .Heating = TME.Heating, .ConsumptionRate = TME.ConsumptionRate, .LoadPercentage = TME.LoadPercentage, .FITPriority = TME.FITPriority, .FITRate = TME.FITRate, .OffsetPriority = TME.OffsetPriority, .StandbyPreferred = TME.StandbyPreferred, .Tariff = TME.Tariff, .Temperature = TME.Temperature, .CDD = CDec(IIf(TME.Temperature < My.Settings.CDDBase, (My.Settings.CDDBase - TME.Temperature) / 24.0, 0)), .HDD = CDec(IIf(TME.Temperature > My.Settings.HDDBase, (TME.Temperature - My.Settings.HDDBase) / 24.0, 0)), .IsOffPeak = TME.IsOffPeak})
        Next

        ' Adjust any hours that were pushed forward to 24 due to DST to the beginning of the next day
        Dim AdjustPeriods = From NTPs In NearTermPeriods
                            Where NTPs.PartHour = 24
        For Each AdjustPeriod In AdjustPeriods
            AdjustPeriod.PartHour = 0
            AdjustPeriod.PartDate = DateAdd(DateInterval.Day, 1, AdjustPeriod.PartDate)
        Next

        ' Check if current day is all off peak
        Dim AOP = From NTPs In NearTermPeriods
                  Where NTPs.PartDate = InvokedTime.Date And
                        NTPs.IsOffPeak = False
        CurrentDayAllOffPeak = AOP.Count = 0

        ' Check if next day is all off peak
        AOP = From NTPs In NearTermPeriods
              Where NTPs.PartDate = Tomorrow.Date And
                    NTPs.IsOffPeak = False
        NextDayAllDayOffPeak = AOP.Count = 0

        ' Get Current Tariff
        Dim CT = From NTPs In NearTermPeriods
                 Where NTPs.PartDate = InvokedTime.Date And
                       NTPs.PartHour = InvokedTime.Hour
                 Order By NTPs.PartDate, NTPs.PartHour
                 Take 1
        CurrentTariff = CT(0)

        ' Get Next Tariff
        Dim NT = From NTPs In NearTermPeriods
                 Where NTPs.Tariff <> CurrentTariff.Tariff
                 Order By NTPs.PartDate, NTPs.PartHour
                 Take 1
        NextTariff = NT(0)

        If CurrentTariff.IsOffPeak = True Then
            ' Off Peak started on or before current interval.
            OffPeakStart = DateAdd(DateInterval.Hour, CurrentTariff.PartHour, CurrentTariff.PartDate)
            If NextTariff.IsOffPeak = True Then
                ' Next Tariff is also off peak = assume Peak is next day
                PeakStart = DateAdd(DateInterval.Day, 1, NextTariff.PartDate)
            Else
                ' Set start of Peak Period
                PeakStart = DateAdd(DateInterval.Hour, NextTariff.PartHour, NextTariff.PartDate)
            End If
        Else
            ' Peak started on or before current interval.
            PeakStart = DateAdd(DateInterval.Hour, CurrentTariff.PartHour, CurrentTariff.PartDate)
            If NextTariff.IsOffPeak = True Then
                ' Next Tariff is off peak 
                OffPeakStart = DateAdd(DateInterval.Hour, NextTariff.PartHour, NextTariff.PartDate)
            Else
                ' next tariff is not off peak, assume off peak is next day
                OffPeakStart = DateAdd(DateInterval.Day, 1, NextTariff.PartDate)
            End If

        End If

        If Sunrise.Date <> InvokedTime.Date Then
            Dim SunriseSunsetData As Result = GetSunriseSunset(Of Result)(InvokedTime)
            With SunriseSunsetData.results
                CivilTwilightSunrise = .civil_twilight_begin.ToLocalTime
                CivilTwilightSunset = .civil_twilight_end.ToLocalTime
                Sunrise = .sunrise.ToLocalTime
                Sunset = .sunset.ToLocalTime
            End With
        End If

    End Sub
    Sub DoPerMinuteTasks()
        Dim Minute As Integer = Now.Minute
        If Minute Mod 10 = 6 Or Minute Mod 10 = 1 Then
            If My.Settings.PVReportingEnabled Then
                If ReportingTimer.Enabled = False Then
                    ReportingTimer.Start()
                    EventLog.WriteEntry("PVOutput Timer Started", EventLogEntryType.Information, 110)
                    SendPowerwallData(Now)
                End If
            End If
        End If
        If Minute Mod 5 = 0 And Minute Mod 10 <> 0 Then
            If ForecastTimer.Enabled = False Then
                ForecastTimer.Start()
                EventLog.WriteEntry("Solar Forecast and Charge Monitoring Timer Started", EventLogEntryType.Information, 111)
                CheckSOCLevel()
            End If
        End If
        If Minute Mod 55 = 0 Then
            If FiveBeforeTheHourTimer.Enabled = False Then
                FiveBeforeTheHourTimer.Start()
                EventLog.WriteEntry("Hourly Timer Started", EventLogEntryType.Information, 112)
                DoHourlyTasks(Now)
            End If
        End If
        AggregateToMinute()
    End Sub
    Sub DoHourlyTasks(InvokedTime As DateTime)
        If InvokedTime.Hour = 23 Then
            If My.Settings.WWGetWeather Then GetWWTemperature()
            If My.Settings.WUGetWeather Then GetWUTemperature()
        End If
        'BuildTariffMap()
        SetOffPeakHours(Now)
        'SetTariffHours(Now)
    End Sub
    Function GetUnsecuredJSONResult(Of JSONType)(URL As String) As JSONType
        Dim response As HttpWebResponse
        Try
            Dim request As WebRequest = WebRequest.Create(URL)
            Try
                response = CType(request.GetResponse(), HttpWebResponse)
            Catch ex As WebException
                Dim resp As HttpWebResponse = CType(ex.Response, HttpWebResponse)
                If resp.StatusCode <> 502 Then
                    EventLog.WriteEntry(ex.Message & vbCrLf & vbCrLf & ex.StackTrace, EventLogEntryType.Error)
                End If
                Return Nothing
            Catch ex As Exception
                Return Nothing
            End Try
            If Not response Is Nothing Then
                Dim dataStream As Stream = response.GetResponseStream()
                Dim reader As StreamReader = New StreamReader(dataStream)
                Dim responseFromServer As String = reader.ReadToEnd()
                reader.Close()
                response.Close()
                Return JsonConvert.DeserializeObject(Of JSONType)(responseFromServer)
            End If
        Catch Ex As Exception
            EventLog.WriteEntry(Ex.Message & vbCrLf & vbCrLf & Ex.StackTrace, EventLogEntryType.Error)
            Return Nothing
        End Try
    End Function
    Function GetUnsecured(URI As String) As Integer
        Try
            Dim request As WebRequest = WebRequest.Create(URI)
            Dim response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
            GetUnsecured = response.StatusCode
            response.Close()
        Catch Ex As Exception
            EventLog.WriteEntry(Ex.Message & vbCrLf & vbCrLf & Ex.StackTrace, EventLogEntryType.Error)
            Return 0
        End Try
    End Function
    Function GetSunriseSunset(Of JSONType)(AsAt As Date) As JSONType
        Return GetUnsecuredJSONResult(Of JSONType)(String.Format("https://api.sunrise-sunset.org/json?lat={0}&lng={1}&formatted=0&date={2:yyyy-MM-dd}", My.Settings.PVSystemLattitude, My.Settings.PVSystemLongitude, AsAt))
    End Function
    Function CheckSunIsUp(AsAt As Date) As Boolean
        If AsAt.Date = Now.Date Then
            If AsAt.Date = CivilTwilightSunrise.Date Then
                If AsAt > CivilTwilightSunrise And AsAt < CivilTwilightSunset Then Return True
                If CivilTwilightSunrise.Date <> AsAt.Date Then
                    Dim SunriseSunsetData As Result = GetSunriseSunset(Of Result)(AsAt)
                    With SunriseSunsetData.results
                        CivilTwilightSunrise = .civil_twilight_begin.ToLocalTime
                        CivilTwilightSunset = .civil_twilight_end.ToLocalTime
                        Sunrise = .sunrise.ToLocalTime
                        Sunset = .sunset.ToLocalTime
                    End With
                End If
                If AsAt > CivilTwilightSunrise And AsAt < CivilTwilightSunset Then Return True
            End If
        Else
            If AsAtSunrise Is Nothing Then
                AsAtSunrise = GetSunriseSunset(Of Result)(AsAt)
            End If
            If AsAtSunrise.results.sunrise.ToLocalTime.Date <> AsAt.Date Then
                AsAtSunrise = GetSunriseSunset(Of Result)(AsAt)
            End If
            If AsAt > AsAtSunrise.results.civil_twilight_begin.ToLocalTime And AsAt < AsAtSunrise.results.civil_twilight_end.ToLocalTime Then Return True
        End If
        Return False
    End Function
#End Region
#Region "Insolation Forecasts and Charge Targets"
    'Sub OperateBasedOnTariff()
    '    Dim InvokedTime As DateTime = Now
    '    SetTariffHours(InvokedTime)
    '    If Not FirstReadingsAvailable Then
    '        GetObservationAndStore()
    '    End If
    '    GetForecasts()
    '    Dim RawTargetSOC As Integer
    '    Dim ShortfallInsolation As Single = 0
    '    Dim NoStandbyTargetSOC As Single = 0
    '    Dim StandbyTargetSOC As Single = 0
    '    Dim RemainingOvernightRatio As Single
    '    Dim RemainingInsolationToday As Single
    '    Dim ForecastInsolationTomorrow As Single
    '    Dim PWPeakConsumption As Integer = CInt(IIf(CurrentDOW = DayOfWeek.Saturday Or CurrentDOW = DayOfWeek.Sunday, My.Settings.PWPeakConsumptionWeekend, My.Settings.PWPeakConsumption))
    '    Dim RemainingOffPeak As Single
    '    Dim Intent As String = "Thinking"
    '    Dim NewTarget As Single = 0
    '    If InvokedTime > Sunrise And InvokedTime < Sunset Then
    '        RemainingOvernightRatio = 1
    '        RemainingInsolationToday = CurrentDayForecast.PVEstimate
    '        ForecastInsolationTomorrow = NextDayForecastGeneration
    '        ShortfallInsolation = 0
    '        Intent = "Sun is Up"
    '    ElseIf InvokedTime > PeakStart And InvokedTime < Sunrise Then
    '        RemainingOvernightRatio = 0
    '        RemainingInsolationToday = CurrentDayForecast.PVEstimate
    '        ForecastInsolationTomorrow = NextDayForecastGeneration
    '        ShortfallInsolation = PWPeakConsumption - NextDayForecastGeneration
    '        Intent = "Waiting for Sunrise"
    '    ElseIf InvokedTime > Sunset And InvokedTime < OffPeakStart Then
    '        RemainingOvernightRatio = 1
    '        RemainingInsolationToday = 0
    '        ForecastInsolationTomorrow = NextDayForecastGeneration
    '        ShortfallInsolation = PWPeakConsumption - NextDayForecastGeneration
    '        Intent = "Waiting for Off Peak"
    '    ElseIf InvokedTime > OffPeakStart And InvokedTime < PeakStart Then
    '        RemainingOvernightRatio = CSng((DateDiff(DateInterval.Hour, InvokedTime, PeakStart) + 1) / OffPeakHours)
    '        If RemainingOvernightRatio < 0 Then RemainingOvernightRatio = 0
    '        If RemainingOvernightRatio > 1 Then RemainingOvernightRatio = 1
    '        RemainingInsolationToday = NextDayForecastGeneration
    '        ForecastInsolationTomorrow = 0
    '        ShortfallInsolation = PWPeakConsumption - NextDayForecastGeneration
    '        Intent = "Monitoring"
    '    End If
    '    RemainingOffPeak = My.Settings.PWOvernightLoad * RemainingOvernightRatio
    '    RawTargetSOC = My.Settings.PWMorningBuffer + CInt(RemainingOffPeak)
    '    If ShortfallInsolation < 0 Then ShortfallInsolation = 0
    '    NoStandbyTargetSOC = RawTargetSOC + (ShortfallInsolation / My.Settings.PWCapacity * 100)
    '    If NoStandbyTargetSOC > 100 Then NoStandbyTargetSOC = 100
    '    StandbyTargetSOC = My.Settings.PWMorningBuffer + (ShortfallInsolation / My.Settings.PWCapacity * 100)
    '    If StandbyTargetSOC > 100 Then StandbyTargetSOC = 100
    '    NewTarget = CSng(IIf(My.Settings.PWOvernightStandby, StandbyTargetSOC, NoStandbyTargetSOC))
    '    If InvokedTime > Sunset And InvokedTime < OffPeakStart Then
    '        If ShortfallInsolation > 0 Or NewTarget > SOC.percentage Then Intent = "Planning to Charge" Else Intent = "No Charging Required"
    '    End If
    '    If My.Settings.PWControlEnabled Then
    '        Try
    '            If InvokedTime > DateAdd(DateInterval.Minute, -10, PeakStart) And Not CurrentDayAllOffPeak And (PreCharging Or AboveMinBackup Or OnStandby) Then
    '                OperationLockout = PeakStart
    '                EventLog.WriteEntry(String.Format("Reaching end of off-peak period with SOC={0}, was aiming for Target={1}", SOC.percentage, StandbyTargetSOC), EventLogEntryType.Information, 504)
    '                DoExitCharging(Intent)
    '            ElseIf (InvokedTime >= OffPeakStart And InvokedTime < PeakStart And InvokedTime > OperationLockout) Then
    '                If My.Settings.VerboseLogging Then EventLog.WriteEntry(String.Format("In Off Peak Period: Current SOC={0}, Minimum required at end of Off-Peak={1}, Shortfall Generation Tomorrow={2}, As at now, Charge Target={3}", SOC.percentage, RawTargetSOC, ShortfallInsolation, NoStandbyTargetSOC), EventLogEntryType.Information, 500)
    '                If My.Settings.DebugLogging Then EventLog.WriteEntry(String.Format("In Off Peak Period: Invoked={0:yyyy-MM-dd HH:mm}, OperationStart={1:yyyy-MM-dd HH:mm}, OperationEnd={2:yyyy-MM-dd HH:mm}", InvokedTime, OffPeakStart, PeakStart), EventLogEntryType.Information, 714)
    '                If NextDayAllDayOffPeak And (Not OnStandby Or SOC.percentage > LastTarget) And Not AutonomousMode Then
    '                    If SetPWMode("Switching to Standby for Off Peak, Standby Mode Enabled", "Standby", SOC.percentage, "self_consumption", Intent) = 202 Then
    '                        OnStandby = True
    '                        PreCharging = False
    '                    End If
    '                ElseIf My.Settings.PWOvernightStandby And SOC.percentage >= StandbyTargetSOC And Not PreCharging And (Not OnStandby Or SOC.percentage > LastTarget) Then
    '                    If SetPWMode("Current SOC above required morning SOC, Standby Mode Enabled", "Standby", SOC.percentage, "self_consumption", Intent) = 202 Then
    '                        OnStandby = True
    '                        PreCharging = False
    '                    End If
    '                ElseIf My.Settings.PWWeekendStandbyOnTarget And SOC.percentage >= My.Settings.PWWeekendStandbyTarget And CurrentDayAllOffPeak And Not AutonomousMode And Not PreCharging And (Not OnStandby Or SOC.percentage > LastTarget) Then
    '                    If SetPWMode("Current SOC above weekend target, standby on target enabled", "Standby", SOC.percentage, "self_consumption", Intent) = 202 Then
    '                        OnStandby = True
    '                        PreCharging = False
    '                    End If
    '                ElseIf (SOC.percentage < StandbyTargetSOC And OnStandby And Not CurrentDayAllOffPeak And Not NextDayAllDayOffPeak) Or (SOC.percentage < NoStandbyTargetSOC And Not OnStandby And Not PreCharging) Then
    '                    If My.Settings.VerboseLogging Then EventLog.WriteEntry(String.Format("Current SOC below required setting: Current SOC={0}, Required at end of Off-Peak={1}, Shortfall Generation Tomorrow={2}, As at now, Charge Target={3}", SOC.percentage, StandbyTargetSOC, ShortfallInsolation, NewTarget), EventLogEntryType.Information, 501)
    '                    If SetPWMode("Current SOC below required morning SOC", IIf(NewTarget > (SOC.percentage + 5), "Charging", "Standby").ToString, NewTarget, IIf(My.Settings.PWChargeModeBackup, "backup", "self_consumption").ToString, Intent) = 202 Then
    '                        PreCharging = True
    '                        OnStandby = False
    '                    End If
    '                ElseIf SOC.percentage >= NoStandbyTargetSOC And PreCharging And Not OnStandby And My.Settings.PWOvernightStandby Then
    '                    EventLog.WriteEntry(String.Format("Current SOC above required setting and Standby Mode Enabled: Current SOC={0}, Required at end of Off-Peak={1}, Shortfall Generation Tomorrow={2}, As at now, Charge Target={3}", SOC.percentage, RawTargetSOC, ShortfallInsolation, NoStandbyTargetSOC), EventLogEntryType.Information, 505)
    '                    If SetPWMode("Switching to Standby for Off Peak, Standby Mode Enabled", "Standby", SOC.percentage, "self_consumption", Intent) = 202 Then
    '                        OnStandby = True
    '                        PreCharging = False
    '                    End If
    '                ElseIf (LastTarget < NewTarget And OnStandby And Not CurrentDayAllOffPeak And Not NextDayAllDayOffPeak) Or (LastTarget < NewTarget And PreCharging) Then
    '                    If My.Settings.VerboseLogging Then EventLog.WriteEntry(String.Format("Charge Target Increased & SOC below required setting: Current SOC={0}, Required at end of Off-Peak={1}, Shortfall Generation Tomorrow={2}, As at now, Charge Target={3}", SOC.percentage, StandbyTargetSOC, ShortfallInsolation, NewTarget), EventLogEntryType.Information, 514)
    '                    If SetPWMode("Current SOC below required morning SOC", IIf(NewTarget > (SOC.percentage + 5), "Charging", "Standby").ToString, NewTarget, IIf(My.Settings.PWChargeModeBackup, "backup", "self_consumption").ToString, Intent) = 202 Then
    '                        PreCharging = True
    '                        OnStandby = False
    '                    End If
    '                ElseIf SOC.percentage >= NoStandbyTargetSOC And PreCharging And Not OnStandby Then
    '                    EventLog.WriteEntry(String.Format("Current SOC above required setting: Current SOC={0}, Required at end of Off-Peak={1}, Shortfall Generation Tomorrow={2}, As at now, Charge Target={3}", SOC.percentage, RawTargetSOC, ShortfallInsolation, NoStandbyTargetSOC), EventLogEntryType.Information, 502)
    '                    DoExitCharging(Intent)
    '                End If
    '            Else
    '                If My.Settings.VerboseLogging Then EventLog.WriteEntry(String.Format("Outside Operation Period: SOC={0}", SOC.percentage), EventLogEntryType.Information, 503)
    '            End If
    '        Catch Ex As Exception
    '            EventLog.WriteEntry(Ex.Message & vbCrLf & vbCrLf & Ex.StackTrace, EventLogEntryType.Error, 510)
    '        End Try
    '    End If
    '    If My.Settings.PBIChargeIntentEndpoint <> String.Empty Then
    '        Try
    '            Dim PBIRows As New PBIChargeLogging With {.Rows = New List(Of ChargePlan)}
    '            PBIRows.Rows.Add(New ChargePlan With {.AsAt = InvokedTime, .CurrentSOC = SOC.percentage, .RemainingInsolation = RemainingInsolationToday, .ForecastGeneration = ForecastInsolationTomorrow, .MorningBuffer = My.Settings.PWMorningBuffer, .OperatingIntent = Intent, .RequiredSOC = NewTarget, .RemainingOffPeak = RemainingOffPeak * My.Settings.PWCapacity / 100, .Shortfall = ShortfallInsolation})
    '            Dim PowerBIPostResult As Integer = PostPowerBIStreamingData(My.Settings.PBIChargeIntentEndpoint, PBIRows)
    '        Catch ex As Exception

    '        End Try
    '    End If
    'End Sub
    Sub CheckSOCLevel()
        Dim InvokedTime As DateTime = Now
        SetOffPeakHours(InvokedTime)
        If Not FirstReadingsAvailable Then
            GetObservationAndStore()
        End If
        GetForecasts()
        Dim RawTargetSOC As Integer
        Dim ShortfallInsolation As Single = 0
        Dim NoStandbyTargetSOC As Single = 0
        Dim StandbyTargetSOC As Single = 0
        Dim RemainingOvernightRatio As Single
        Dim RemainingInsolationToday As Single
        Dim ForecastInsolationTomorrow As Single
        Dim PWPeakConsumption As Integer = CInt(IIf(CurrentDOW = DayOfWeek.Saturday Or CurrentDOW = DayOfWeek.Sunday, My.Settings.PWPeakConsumptionWeekend, My.Settings.PWPeakConsumption))
        Dim RemainingOffPeak As Single
        Dim Intent As String = "Thinking"
        Dim NewTarget As Single = 0
        If InvokedTime > Sunrise And InvokedTime < Sunset Then
            RemainingOvernightRatio = 1
            RemainingInsolationToday = CurrentDayForecast.PVEstimate
            ForecastInsolationTomorrow = NextDayForecastGeneration
            ShortfallInsolation = 0
            Intent = "Sun is Up"
        ElseIf InvokedTime > PeakStart And InvokedTime < Sunrise Then
            RemainingOvernightRatio = 0
            RemainingInsolationToday = CurrentDayForecast.PVEstimate
            ForecastInsolationTomorrow = NextDayForecastGeneration
            ShortfallInsolation = PWPeakConsumption - NextDayForecastGeneration
            Intent = "Waiting for Sunrise"
        ElseIf InvokedTime > Sunset And InvokedTime < OffPeakStart Then
            RemainingOvernightRatio = 1
            RemainingInsolationToday = 0
            ForecastInsolationTomorrow = NextDayForecastGeneration
            ShortfallInsolation = PWPeakConsumption - NextDayForecastGeneration
            Intent = "Waiting for Off Peak"
        ElseIf InvokedTime > OffPeakStart And InvokedTime < PeakStart Then
            RemainingOvernightRatio = CSng((DateDiff(DateInterval.Hour, InvokedTime, PeakStart) + 1) / OffPeakHours)
            If RemainingOvernightRatio < 0 Then RemainingOvernightRatio = 0
            If RemainingOvernightRatio > 1 Then RemainingOvernightRatio = 1
            RemainingInsolationToday = NextDayForecastGeneration
            ForecastInsolationTomorrow = 0
            ShortfallInsolation = PWPeakConsumption - NextDayForecastGeneration
            Intent = "Monitoring"
        End If
        RemainingOffPeak = My.Settings.PWOvernightLoad * RemainingOvernightRatio
        RawTargetSOC = My.Settings.PWMorningBuffer + CInt(RemainingOffPeak)
        If ShortfallInsolation < 0 Then ShortfallInsolation = 0
        NoStandbyTargetSOC = RawTargetSOC + (ShortfallInsolation / My.Settings.PWCapacity * 100)
        If NoStandbyTargetSOC > 100 Then NoStandbyTargetSOC = 100
        StandbyTargetSOC = My.Settings.PWMorningBuffer + (ShortfallInsolation / My.Settings.PWCapacity * 100)
        If StandbyTargetSOC > 100 Then StandbyTargetSOC = 100
        NewTarget = CSng(IIf(My.Settings.PWOvernightStandby, StandbyTargetSOC, NoStandbyTargetSOC))
        If InvokedTime > Sunset And InvokedTime < OffPeakStart Then
            If ShortfallInsolation > 0 Or NewTarget > SOC.percentage Then Intent = "Planning to Charge" Else Intent = "No Charging Required"
        End If
        If My.Settings.PWControlEnabled Then
            Try
                If InvokedTime > DateAdd(DateInterval.Minute, -10, PeakStart) And Not CurrentDayAllOffPeak And (PreCharging Or AboveMinBackup Or OnStandby) Then
                    OperationLockout = PeakStart
                    EventLog.WriteEntry(String.Format("Reaching end of off-peak period with SOC={0}, was aiming for Target={1}", SOC.percentage, StandbyTargetSOC), EventLogEntryType.Information, 504)
                    DoExitCharging(Intent)
                ElseIf (InvokedTime >= OffPeakStart And InvokedTime < PeakStart And InvokedTime > OperationLockout) Then
                    If My.Settings.VerboseLogging Then EventLog.WriteEntry(String.Format("In Operation Period: Current SOC={0}, Minimum required at end of Off-Peak={1}, Shortfall Generation Tomorrow={2}, As at now, Charge Target={3}", SOC.percentage, RawTargetSOC, ShortfallInsolation, NoStandbyTargetSOC), EventLogEntryType.Information, 500)
                    If My.Settings.DebugLogging Then EventLog.WriteEntry(String.Format("In Operation Period: Invoked={0:yyyy-MM-dd HH:mm}, OperationStart={1:yyyy-MM-dd HH:mm}, OperationEnd={2:yyyy-MM-dd HH:mm}", InvokedTime, OffPeakStart, PeakStart), EventLogEntryType.Information, 714)
                    If NextDayAllDayOffPeak And (Not OnStandby Or SOC.percentage > LastTarget) And SOC.percentage > PWTarget Then
                        If SetPWMode("Switching to Standby for Off Peak, Standby Mode Enabled", "Standby", SOC.percentage, "self_consumption", Intent) = 202 Then
                            OnStandby = True
                            PreCharging = False
                        End If
                    ElseIf My.Settings.PWOvernightStandby And SOC.percentage >= StandbyTargetSOC And Not PreCharging And (Not OnStandby Or SOC.percentage > LastTarget) And SOC.percentage > PWTarget Then
                        If SetPWMode("Current SOC above required morning SOC, Standby Mode Enabled", "Standby", SOC.percentage, "self_consumption", Intent) = 202 Then
                            OnStandby = True
                            PreCharging = False
                        End If
                    ElseIf My.Settings.PWWeekendStandbyOnTarget And SOC.percentage >= My.Settings.PWWeekendStandbyTarget And CurrentDayAllOffPeak And Not PreCharging And (Not OnStandby Or SOC.percentage > LastTarget) And SOC.percentage > PWTarget Then
                        If SetPWMode("Current SOC above weekend target, standby on target enabled", "Standby", SOC.percentage, "self_consumption", Intent) = 202 Then
                            OnStandby = True
                            PreCharging = False
                        End If
                    ElseIf (SOC.percentage < StandbyTargetSOC And OnStandby And Not CurrentDayAllOffPeak And Not NextDayAllDayOffPeak) Or (SOC.percentage < NoStandbyTargetSOC And Not OnStandby And Not PreCharging) Then
                        If My.Settings.VerboseLogging Then EventLog.WriteEntry(String.Format("Current SOC below required setting: Current SOC={0}, Required at end of Off-Peak={1}, Shortfall Generation Tomorrow={2}, As at now, Charge Target={3}", SOC.percentage, StandbyTargetSOC, ShortfallInsolation, NewTarget), EventLogEntryType.Information, 501)
                        If SetPWMode("Current SOC below required morning SOC", IIf(NewTarget > (SOC.percentage + 5), "Charging", "Standby").ToString, NewTarget, IIf(My.Settings.PWChargeModeBackup, "backup", "self_consumption").ToString, Intent) = 202 Then
                            PreCharging = True
                            OnStandby = False
                        End If
                    ElseIf SOC.percentage >= NoStandbyTargetSOC And PreCharging And Not OnStandby And My.Settings.PWOvernightStandby And SOC.percentage > PWTarget Then
                        EventLog.WriteEntry(String.Format("Current SOC above required setting and Standby Mode Enabled: Current SOC={0}, Required at end of Off-Peak={1}, Shortfall Generation Tomorrow={2}, As at now, Charge Target={3}", SOC.percentage, RawTargetSOC, ShortfallInsolation, NoStandbyTargetSOC), EventLogEntryType.Information, 505)
                        If SetPWMode("Switching to Standby for Off Peak, Standby Mode Enabled", "Standby", SOC.percentage, "self_consumption", Intent) = 202 Then
                            OnStandby = True
                            PreCharging = False
                        End If
                    ElseIf (LastTarget < NewTarget And OnStandby And Not CurrentDayAllOffPeak And Not NextDayAllDayOffPeak) Or (LastTarget < NewTarget And PreCharging) Then
                        If My.Settings.VerboseLogging Then EventLog.WriteEntry(String.Format("Charge Target Increased & SOC below required setting: Current SOC={0}, Required at end of Off-Peak={1}, Shortfall Generation Tomorrow={2}, As at now, Charge Target={3}", SOC.percentage, StandbyTargetSOC, ShortfallInsolation, NewTarget), EventLogEntryType.Information, 514)
                        If SetPWMode("Current SOC below required morning SOC", IIf(NewTarget > (SOC.percentage + 5), "Charging", "Standby").ToString, NewTarget, IIf(My.Settings.PWChargeModeBackup, "backup", "self_consumption").ToString, Intent) = 202 Then
                            PreCharging = True
                            OnStandby = False
                        End If
                    ElseIf SOC.percentage >= NoStandbyTargetSOC And PreCharging And Not OnStandby And SOC.percentage > PWTarget Then
                        EventLog.WriteEntry(String.Format("Current SOC above required setting: Current SOC={0}, Required at end of Off-Peak={1}, Shortfall Generation Tomorrow={2}, As at now, Charge Target={3}", SOC.percentage, RawTargetSOC, ShortfallInsolation, NoStandbyTargetSOC), EventLogEntryType.Information, 502)
                        DoExitCharging(Intent)
                    End If
                Else
                    If My.Settings.VerboseLogging Then EventLog.WriteEntry(String.Format("Outside Operation Period: SOC={0}", SOC.percentage), EventLogEntryType.Information, 503)
                End If
            Catch Ex As Exception
                EventLog.WriteEntry(Ex.Message & vbCrLf & vbCrLf & Ex.StackTrace, EventLogEntryType.Error, 510)
            End Try
        End If
        If My.Settings.PBIChargeIntentEndpoint <> String.Empty Then
            Try
                Dim PBIRows As New PBIChargeLogging With {.Rows = New List(Of ChargePlan)}
                PBIRows.Rows.Add(New ChargePlan With {.AsAt = InvokedTime, .CurrentSOC = SOC.percentage, .RemainingInsolation = RemainingInsolationToday, .ForecastGeneration = ForecastInsolationTomorrow, .MorningBuffer = My.Settings.PWMorningBuffer, .OperatingIntent = Intent, .RequiredSOC = NewTarget, .RemainingOffPeak = RemainingOffPeak * My.Settings.PWCapacity / 100, .Shortfall = ShortfallInsolation})
                Dim PowerBIPostResult As Integer = PostPowerBIStreamingData(My.Settings.PBIChargeIntentEndpoint, PBIRows)
            Catch ex As Exception

            End Try
        End If
    End Sub
    Sub GetForecasts()
        Try
            Dim NewForecastsRetrieved As Boolean = False
            Dim InvokedTime As DateTime = Now
            If DateAdd(DateInterval.Hour, 1, ForecastsRetrieved) < InvokedTime And InvokedTime.Hour < Sunset.Hour Then
                PVForecast = GetSolCastResult(Of OutputForecast)()
                CurrentDayForecast = New DayForecast With {.ForecastDate = DateAdd(DateInterval.Day, 0, Now.Date), .PVEstimate = 0, .MorningForecast = 0}
                NextDayForecast = New DayForecast With {.ForecastDate = DateAdd(DateInterval.Day, 1, Now.Date), .PVEstimate = 0, .MorningForecast = 0}
                SecondDayForecast = New DayForecast With {.ForecastDate = DateAdd(DateInterval.Day, 2, Now.Date), .PVEstimate = 0, .MorningForecast = 0}
                ForecastsRetrieved = InvokedTime
                NewForecastsRetrieved = True
            End If
            If My.Settings.PVSendForecast Then
                If LastPeriodForecast Is Nothing Then
                    LastPeriodForecast = New Forecast With {.period = PVForecast.forecasts(0).period, .period_end = PVForecast.forecasts(0).period_end, .pv_estimate = PVForecast.forecasts(0).pv_estimate}
                    CurrentPeriodForecast = New Forecast With {.period = PVForecast.forecasts(1).period, .period_end = PVForecast.forecasts(1).period_end, .pv_estimate = PVForecast.forecasts(1).pv_estimate}
                ElseIf CurrentPeriodForecast.period_end.ToLocalTime < InvokedTime Then
                    LastPeriodForecast = New Forecast With {.period = CurrentPeriodForecast.period, .period_end = CurrentPeriodForecast.period_end, .pv_estimate = CurrentPeriodForecast.pv_estimate}
                    CurrentPeriodForecast = New Forecast With {.period = PVForecast.forecasts(0).period, .period_end = PVForecast.forecasts(0).period_end, .pv_estimate = PVForecast.forecasts(0).pv_estimate}
                End If
            End If
            If NewForecastsRetrieved Then
                For Each PeriodForecast As Forecast In PVForecast.forecasts()
                    With PeriodForecast
                        Select Case DateDiff(DateInterval.Day, InvokedTime.Date, .period_end.ToLocalTime.Date)
                            Case 0
                                CurrentDayForecast.PVEstimate += .pv_estimate / 2
                                If .period_end.ToLocalTime.Hour <= 9 Then CurrentDayForecast.MorningForecast += .pv_estimate / 2
                            Case 1
                                NextDayForecast.PVEstimate += .pv_estimate / 2
                                If .period_end.ToLocalTime.Hour <= 9 Then NextDayForecast.MorningForecast += .pv_estimate / 2
                            Case 2
                                SecondDayForecast.PVEstimate += .pv_estimate / 2
                                If .period_end.ToLocalTime.Hour <= 9 Then SecondDayForecast.MorningForecast += .pv_estimate / 2
                            Case Else
                                Exit For
                        End Select
                    End With
                Next
                Dim ForecastLogEntry As String = "Three Day Forecast" & vbCrLf
                With CurrentDayForecast
                    ForecastLogEntry += vbCrLf & String.Format("Date: {0:yyyy-MM-dd} Total: {1} Morning: {2}", .ForecastDate, .PVEstimate, .MorningForecast)
                End With
                With NextDayForecast
                    ForecastLogEntry += vbCrLf & String.Format("Date: {0:yyyy-MM-dd} Total: {1} Morning: {2}", .ForecastDate, .PVEstimate, .MorningForecast)
                End With
                With SecondDayForecast
                    ForecastLogEntry += vbCrLf & String.Format("Date: {0:yyyy-MM-dd} Total: {1} Morning: {2}", .ForecastDate, .PVEstimate, .MorningForecast)
                End With
                EventLog.WriteEntry(ForecastLogEntry, EventLogEntryType.Information, 1000)
            End If
            If InvokedTime.Hour >= 0 And InvokedTime.Hour < PeakStartHour Then
                With CurrentDayForecast
                    NextDayForecastGeneration = .PVEstimate
                    NextDayMorningGeneration = .MorningForecast
                End With
            Else
                With NextDayForecast
                    NextDayForecastGeneration = .PVEstimate
                    NextDayMorningGeneration = .MorningForecast
                End With
            End If
        Catch Ex As Exception
            EventLog.WriteEntry(Ex.Message & vbCrLf & vbCrLf & Ex.StackTrace, EventLogEntryType.Error)
        End Try
    End Sub
    Private Sub SendForecast()
        Dim InvokedTime As DateTime = Now
        If CurrentPeriodForecast Is Nothing Or LastPeriodForecast Is Nothing Then
            GetForecasts()
        End If
        If Not CurrentPeriodForecast Is Nothing And Not LastPeriodForecast Is Nothing Then
            If CurrentPeriodForecast.period_end.ToLocalTime < InvokedTime Then
                GetForecasts()
            End If
            If CurrentPeriodForecast.pv_estimate > 0 Or LastPeriodForecast.pv_estimate > 0 Then
                If LastPeriodForecast.period_end.ToLocalTime <= InvokedTime And CurrentPeriodForecast.period_end.ToLocalTime >= DateAdd(DateInterval.Minute, 4, InvokedTime) Then
                    Dim EndPower As Single = CurrentPeriodForecast.pv_estimate
                    Dim DiffPower As Single = EndPower - LastPeriodForecast.pv_estimate
                    Dim ElapsedMinutes As Double = DateDiff(DateInterval.Minute, LastPeriodForecast.period_end.ToLocalTime, InvokedTime)
                    Dim ElapsedMinutesMod5 As Double = ElapsedMinutes - ElapsedMinutes Mod 5
                    Dim PeriodRatio As Double
                    Select Case ElapsedMinutesMod5
                        Case 0
                            PeriodRatio = 1 - (0 / 24)
                        Case 5
                            PeriodRatio = 1 - (1 / 24)
                        Case 10
                            PeriodRatio = 1 - (5 / 24)
                        Case 15
                            PeriodRatio = 1 - (12 / 24)
                        Case 20
                            PeriodRatio = 1 - (19 / 24)
                        Case 25
                            PeriodRatio = 1 - (23 / 24)
                        Case 30
                            PeriodRatio = 1 - (24 / 24)
                    End Select
                    Dim ForecastSolar As String = (EndPower - (DiffPower * PeriodRatio)).ToString
                    PVSaveExtendedData(InvokedTime, My.Settings.PVSendForecastAs, ForecastSolar)
                End If
            End If
        End If
    End Sub
    Function GetSolCastResult(Of JSONType)() As JSONType
        Dim responseFromServer As String = String.Empty
        Try
            Dim request As WebRequest = WebRequest.Create(String.Format(My.Settings.SolcastAddress, My.Settings.PVSystemLongitude, My.Settings.PVSystemLattitude, My.Settings.PVSystemCapacity, My.Settings.PVSystemTilt, My.Settings.PVSystemAzimuth, My.Settings.PVSystemInstallDate, My.Settings.SolcastAPIKey))
            Dim response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
            Dim dataStream As Stream = response.GetResponseStream()
            Dim reader As StreamReader = New StreamReader(dataStream)
            responseFromServer = reader.ReadToEnd()

            If My.Settings.DualPVSystem Then
                ' Deserialise the results of the first array
                Dim Results As OutputForecast = JsonConvert.DeserializeObject(Of OutputForecast)(responseFromServer)


                ' Request PV Data for the second array
                Dim responseFromServer2 As String = String.Empty
                Dim request2 As WebRequest = WebRequest.Create(String.Format(My.Settings.SolcastAddress, My.Settings.PVSystemLongitude, My.Settings.PVSystemLattitude, My.Settings.PVSystem2Capacity, My.Settings.PVSystem2Tilt, My.Settings.PVSystem2Azimuth, My.Settings.PVSystemInstallDate, My.Settings.SolcastAPIKey))
                Dim response2 As HttpWebResponse = CType(request2.GetResponse(), HttpWebResponse)
                Dim dataStream2 As Stream = response2.GetResponseStream()
                Dim reader2 As StreamReader = New StreamReader(dataStream2)
                responseFromServer2 = reader2.ReadToEnd()
                Dim Results2 As OutputForecast = JsonConvert.DeserializeObject(Of OutputForecast)(responseFromServer2)

                ' Add the results for the second array to that of the first
                For i As Integer = 0 To Results.forecasts.Count - 1
                    Dim newEstimate As Single = Convert.ToSingle(Results.forecasts.Item(i).pv_estimate) + Convert.ToSingle(Results2.forecasts.Item(i).pv_estimate)
                    Results.forecasts.Item(i).pv_estimate = newEstimate
                Next

                ' Convert the combined PV Estimates back to JSON so the rest of the code will work as before
                responseFromServer = JsonConvert.SerializeObject(Results, Formatting.None)

                reader2.Close()
                response2.Close()
            End If

            If My.Settings.LogData Then
                Try
                    SyncLock DBLock
                        SPs.spStoreForecast(responseFromServer)
                    End SyncLock
                Catch ex As Exception
                    EventLog.WriteEntry(ex.Message & vbCrLf & vbCrLf & ex.StackTrace, EventLogEntryType.Error)
                End Try
            End If
            GetSolCastResult = JsonConvert.DeserializeObject(Of JSONType)(responseFromServer)
            reader.Close()
            response.Close()
        Catch Ex As Exception
            EventLog.WriteEntry(Ex.Message & vbCrLf & vbCrLf & Ex.StackTrace & vbCrLf & vbCrLf & responseFromServer, EventLogEntryType.Error)
        End Try
    End Function
#End Region
#Region "Six Second Logger"
    Sub GetObservationAndStore()
        Dim ObservationTime As DateTime
        Dim LastObservationTime As DateTime
        If Not MeterReading Is Nothing Then
            LastObservationTime = MeterReading.site.last_communication_time
        End If
        Dim GotResults As Boolean = False
        Try
            SyncLock PWLock
                Try
                    If Not SkipObservation Then
                        MeterReading = GetPWAPIResult(Of MeterAggregates)("meters/aggregates")
                        SOC = GetPWAPIResult(Of SOC)("system_status/soe")
                    End If
                    If Not MeterReading Is Nothing Then
                        ObservationTime = MeterReading.site.last_communication_time
                    End If
                    SkipObservation = False
                Catch Ex As Exception
                Finally
                    SkipObservation = False
                End Try
            End SyncLock
            If ObservationTime > LastObservationTime Then
                FirstReadingsAvailable = True
                GotResults = True
            End If
            If My.Settings.LogData And GotResults Then
                With MeterReading
                    SyncLock DBLock
                        Try
                            CompactTA.Insert(ObservationTime.ToUniversalTime, ObservationTime, SOC.percentage, .battery.instant_average_voltage, .site.instant_average_voltage, .battery.instant_power, .site.instant_power, .solar.instant_power, .load.instant_power)
                        Catch Ex As Exception
                            EventLog.WriteEntry(Ex.Message & vbCrLf & vbCrLf & Ex.StackTrace, EventLogEntryType.Error)
                        End Try
                        If Not My.Settings.LogAzureOnly Then
                            Try
                                CompactTALocal.Insert(ObservationTime.ToUniversalTime, ObservationTime, SOC.percentage, .battery.instant_average_voltage, .site.instant_average_voltage, .battery.instant_power, .site.instant_power, .solar.instant_power, .load.instant_power)
                            Catch Ex As Exception
                                EventLog.WriteEntry(Ex.Message & vbCrLf & vbCrLf & Ex.StackTrace, EventLogEntryType.Error)
                            End Try
                        End If
                        If Not My.Settings.LogCompactOnly Then
                            Try
                                Dim ObservationID As Integer = CInt(ObsTA.InsertAndReturnIdentity(ObservationTime.ToUniversalTime, ObservationTime))
                                'ObsTA.Transaction.Commit()
                                SolarTA.Insert(.solar.last_communication_time, .solar.instant_power, .solar.instant_reactive_power, .solar.instant_apparent_power, .solar.frequency, .solar.energy_exported, .solar.energy_imported, .solar.instant_average_voltage, .solar.instant_total_current, .solar.i_a_current, .solar.i_b_current, .solar.i_c_current, ObservationID)
                                BatteryTA.Insert(.battery.last_communication_time, .battery.instant_power, .battery.instant_reactive_power, .battery.instant_apparent_power, .battery.frequency, .battery.energy_exported, .battery.energy_imported, .battery.instant_average_voltage, .battery.instant_total_current, .battery.i_a_current, .battery.i_b_current, .battery.i_c_current, ObservationID)
                                SiteTA.Insert(.site.last_communication_time, .site.instant_power, .site.instant_reactive_power, .site.instant_apparent_power, .site.frequency, .site.energy_exported, .site.energy_imported, .site.instant_average_voltage, .site.instant_total_current, .site.i_a_current, .site.i_b_current, .site.i_c_current, ObservationID)
                                LoadTA.Insert(.load.last_communication_time, .load.instant_power, .load.instant_reactive_power, .load.instant_apparent_power, .load.frequency, .load.energy_exported, .load.energy_imported, .load.instant_average_voltage, .load.instant_total_current, .load.i_a_current, .load.i_b_current, .load.i_c_current, ObservationID)
                                SOCTA.Insert(SOC.percentage, ObservationID)
                            Catch Ex As Exception
                                EventLog.WriteEntry(Ex.Message & vbCrLf & vbCrLf & Ex.StackTrace, EventLogEntryType.Error)
                            End Try
                        End If
                    End SyncLock
                End With
            End If
            If My.Settings.PBILiveLoggingEndpoint <> String.Empty Then
                Try
                    Dim PBIRows As New PBILiveLogging With {.Rows = New List(Of SixSecondOb)}
                    PBIRows.Rows.Add(New SixSecondOb With {.AsAt = ObservationTime, .Battery = MeterReading.battery.instant_power, .Grid = MeterReading.site.instant_power, .Load = MeterReading.load.instant_power, .SOC = SOC.percentage, .Solar = CSng(IIf(MeterReading.solar.instant_power < 0, 0, MeterReading.solar.instant_power)), .Voltage = MeterReading.battery.instant_average_voltage})
                    Dim PowerBIPostResult As Integer = PostPowerBIStreamingData(My.Settings.PBILiveLoggingEndpoint, PBIRows)
                Catch ex As Exception

                End Try
            End If
        Catch Ex As Exception
            EventLog.WriteEntry(Ex.Message & vbCrLf & vbCrLf & Ex.StackTrace, EventLogEntryType.Error)
        End Try
    End Sub
    Function GetPWAPIResult(Of JSONType)(API As String) As JSONType
        Return GetUnsecuredJSONResult(Of JSONType)(My.Settings.PWGatewayAddress & "/api/" & API)
    End Function
#End Region
#Region "Powerwall Control"
    Private Function SetPWMode(ActionMessage As String, ActionType As String, Target As Double, Mode As String, ByRef Intent As String) As Integer
        Dim RunningResult As Integer
        SkipObservation = True
        If ActionType = "Standby" Then
            GetPWMode()
            If Target < PWTarget Then Target = PWTarget
        Else
            If Target > My.Settings.PWMinBackupPercentage Then
                Target = Math.Round(Target) + 2
            End If
        End If
        If Target > 100 Then Target = 100
        LastTarget = CInt(Target)
        PWTarget = CDec(Target)
        Try
            If My.Settings.VerboseLogging Then EventLog.WriteEntry(String.Format(ActionMessage & " Current SOC={0}, Current Target={1}", SOC.percentage, Target), EventLogEntryType.Information, 511)
            Intent = ActionType
            Dim ChargeSettings As New Operation With {.backup_reserve_percent = LastTarget, .real_mode = Mode}
            Dim NewChargeSettings As Operation
            Dim APIResult As Integer
            SyncLock PWLock
                NewChargeSettings = PostPWSecureAPISettings(Of Operation)("operation", ChargeSettings, ForceReLogin:=True)
                APIResult = GetPWSecure("config/completed")
                RunningResult = GetPWRunning()
            End SyncLock
            If APIResult = 202 Then
                EventLog.WriteEntry(String.Format("Entered {5} Mode: Current SOC={0}, Current Target={1}, Set Mode={2}, Set Backup Percentage={3}, APIResult = {4}, Reason = {6}", SOC.percentage, Target, NewChargeSettings.real_mode, NewChargeSettings.backup_reserve_percent, APIResult, ActionType, ActionMessage), EventLogEntryType.Information, 512)
                AboveMinBackup = (NewChargeSettings.backup_reserve_percent > My.Settings.PWMinBackupPercentage)
                Intent = ActionType
            Else
                EventLog.WriteEntry(String.Format("Failed to Enter {5} Mode: Current SOC={0}, Attempted Target={1}, Mode={2}, BackupPercentage={3}, APIResult = {4}, Reason = {6}", SOC.percentage, Target, NewChargeSettings.real_mode, NewChargeSettings.backup_reserve_percent, APIResult, ActionType, ActionMessage), EventLogEntryType.Warning, 513)
                Intent = String.Format("Trying to Enter {0}", ActionType)
            End If
            SetPWMode = APIResult
        Catch ex As Exception
            SyncLock PWLock
                RunningResult = GetPWRunning()
            End SyncLock
            SetPWMode = 0
        End Try
        SkipObservation = False
    End Function
    Private Function GetPWRunning() As Integer
        GetPWRunning = GetPWSecure("sitemaster/run")
    End Function
    Private Sub GetPWMode()
        If My.Settings.PWControlEnabled Then
            Dim RunningResult As Integer
            Try
                If Not PreCharging Then
                    Dim APIResult As Integer
                    Dim CurrentChargeSettings As Operation
                    SyncLock PWLock
                        CurrentChargeSettings = GetPWSecureAPIResult(Of Operation)("operation", ForceReLogin:=True)
                        APIResult = GetPWSecure("config/completed")
                        RunningResult = GetPWRunning()
                    End SyncLock
                    If APIResult = 202 Then
                        With CurrentChargeSettings
                            PWTarget = .backup_reserve_percent
                            EventLog.WriteEntry(String.Format("Current PW Mode={0}, BackupPercentage={1}, APIResult = {2}", .real_mode, PWTarget, APIResult), EventLogEntryType.Information, 602)
                            AboveMinBackup = (PWTarget > My.Settings.PWMinBackupPercentage)
                            AutonomousMode = (.real_mode = "autonomous")
                        End With
                    Else
                        EventLog.WriteEntry("Failed to obtain current operation mode", EventLogEntryType.Warning, 513)
                    End If
                End If
            Catch ex As Exception
                SyncLock PWLock
                    RunningResult = GetPWRunning()
                End SyncLock
            End Try
        End If
    End Sub
    Function GetPWSecureAPIResult(Of JSONType)(API As String, Optional ForceReLogin As Boolean = False) As JSONType
        Try
            PWToken = LoginPWLocal(ForceReLogin:=ForceReLogin)
            Dim request As WebRequest = GetPWRequest(API)
            request.Headers.Add("Authorization", "Bearer " & PWToken)
            Dim response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
            Dim dataStream As Stream = response.GetResponseStream()
            Dim reader As StreamReader = New StreamReader(dataStream)
            Dim responseFromServer As String = reader.ReadToEnd()
            GetPWSecureAPIResult = JsonConvert.DeserializeObject(Of JSONType)(responseFromServer)
            reader.Close()
            response.Close()
        Catch Ex As Exception
            EventLog.WriteEntry(Ex.Message & vbCrLf & vbCrLf & Ex.StackTrace, EventLogEntryType.Error)
        End Try
    End Function
    Private Shared Function GetPWRequest(API As String) As WebRequest
        Dim wr As HttpWebRequest
        wr = CType(WebRequest.Create(My.Settings.PWGatewayAddress & "/api/" & API), HttpWebRequest)
        wr.ServerCertificateValidationCallback = Function()
                                                     Return True
                                                 End Function
        Return wr
    End Function
    Function GetPWSecure(API As String, Optional ForceReLogin As Boolean = False) As Integer
        Try
            PWToken = LoginPWLocal(ForceReLogin:=ForceReLogin)
            Dim request As WebRequest = GetPWRequest(API)
            request.Headers.Add("Authorization", "Bearer " & PWToken)
            Dim response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
            Dim dataStream As Stream = response.GetResponseStream()
            Dim reader As StreamReader = New StreamReader(dataStream)
            Dim responseFromServer As String = reader.ReadToEnd()
            GetPWSecure = response.StatusCode
            reader.Close()
            response.Close()
        Catch Ex As Exception
            Return -1
            EventLog.WriteEntry(Ex.Message & vbCrLf & vbCrLf & Ex.StackTrace, EventLogEntryType.Error)
        End Try
    End Function
    Function PostPWSecureAPISettings(Of JSONType)(API As String, Settings As JSONType, Optional ForceReLogin As Boolean = False) As JSONType
        Try
            PWToken = LoginPWLocal(ForceReLogin:=ForceReLogin)
            Dim BodyPostData As String = JsonConvert.SerializeObject(Settings).ToString
            Dim BodyByteStream As Byte() = Encoding.UTF8.GetBytes(BodyPostData)
            Dim request As WebRequest = GetPWRequest(API)
            request.Headers.Add("Authorization", "Bearer " & PWToken)
            request.Method = "POST"
            request.ContentType = "application/json"
            request.ContentLength = BodyByteStream.Length
            Dim BodyStream As Stream = request.GetRequestStream()
            BodyStream.Write(BodyByteStream, 0, BodyByteStream.Length)
            BodyStream.Close()
            Dim response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
            Dim dataStream As Stream = response.GetResponseStream()
            Dim reader As StreamReader = New StreamReader(dataStream)
            Dim responseFromServer As String = reader.ReadToEnd()
            PostPWSecureAPISettings = JsonConvert.DeserializeObject(Of JSONType)(responseFromServer)
            reader.Close()
            response.Close()
        Catch Ex As Exception
            EventLog.WriteEntry(Ex.Message & vbCrLf & vbCrLf & Ex.StackTrace, EventLogEntryType.Error)
        End Try
    End Function
    Function LoginPWLocal(Optional ForceReLogin As Boolean = False) As String
        If PWToken = String.Empty Or ForceReLogin = True Then
            Try
                Dim LoginRequest As New LoginRequest With {
                    .username = My.Settings.PWGatewayUsername,
                    .password = My.Settings.PWGatewayPassword,
                    .email = String.Empty,
                    .force_sm_off = True
                }
                Dim BodyPostData As String = JsonConvert.SerializeObject(LoginRequest).ToString
                Dim BodyByteStream As Byte() = Encoding.ASCII.GetBytes(BodyPostData)
                Dim request As WebRequest = GetPWRequest("login/Basic")
                request.Method = "POST"
                request.ContentType = "application/json"
                request.ContentLength = BodyByteStream.Length
                Dim BodyStream As Stream = request.GetRequestStream()
                BodyStream.Write(BodyByteStream, 0, BodyByteStream.Length)
                BodyStream.Close()
                Dim response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
                Dim dataStream As Stream = response.GetResponseStream()
                Dim reader As StreamReader = New StreamReader(dataStream)
                Dim responseFromServer As String = reader.ReadToEnd()
                Dim LoginResult As LoginResult = JsonConvert.DeserializeObject(Of LoginResult)(responseFromServer)
                PWToken = LoginResult.token
                reader.Close()
                response.Close()
            Catch ex As Exception
                EventLog.WriteEntry(ex.Message & vbCrLf & vbCrLf & ex.StackTrace, EventLogEntryType.Error)
            End Try
        End If
        Return PWToken
    End Function
    Private Sub DoExitCharging(ByRef Intent As String)
        If SetPWMode("Exit Charge or Standby Mode", "Self Consumption", My.Settings.PWMinBackupPercentage, "self_consumption", Intent) = 202 Then
            PreCharging = False
            OnStandby = False
            AboveMinBackup = False
        End If
    End Sub
#End Region
#Region "PVOutput"
    Private Sub DoBackFill(AsAt As DateTime)
        Dim recs As List(Of StatusRecord) = PVGetStatusHistory(AsAt)
        If recs Is Nothing Then Exit Sub
        Dim BeginRecTimeStamp As DateTime = AsAt.Date
        Dim EndRecTimeStamp As DateTime = DateAdd(DateInterval.Minute, -5, DateAdd(DateInterval.Day, 1, BeginRecTimeStamp))
        Dim CurrRecTimeStamp As DateTime = BeginRecTimeStamp
        Dim RecOK As Boolean
        Dim FiveMinData As PWHistoryDataSet.spGet5MinuteAveragesRow = Nothing
        For Each rec As StatusRecord In recs
            RecOK = True
            Dim FoundRecTimeStamp As New DateTime(CInt(Left(rec.d, 4)), CInt(Mid(rec.d, 5, 2)), CInt(Right(rec.d, 2)), CInt(Left(rec.t, 2)), CInt(Right(rec.t, 2)), 0)
            If FoundRecTimeStamp = CurrRecTimeStamp Then
                Try
                    FiveMinData = GetFiveMinuteData(CurrRecTimeStamp)
                    If Not FiveMinData Is Nothing Then
                        If CheckSunIsUp(CurrRecTimeStamp) Then
                            If My.Settings.PVSendPV AndAlso rec.v2 = "NaN" AndAlso FiveMinData.solar_instant_power > 0 Then RecOK = False
                            If My.Settings.PVSendForecast AndAlso rec(My.Settings.PVSendForecastAs).ToString = "NaN" AndAlso FiveMinData.solcast_forecast > 0 Then RecOK = False
                        End If
                        If My.Settings.PVSendLoad AndAlso rec.v4 = "NaN" AndAlso FiveMinData.load_instant_power <> 0 Then RecOK = False
                        If My.Settings.PVSendVoltage AndAlso rec.v6 = "NaN" AndAlso FiveMinData.site_instant_average_voltage <> 0 Then RecOK = False
                        If My.Settings.PVv7 <> String.Empty AndAlso rec.v7 = "NaN" AndAlso CDbl(FiveMinData(My.Settings.PVv7)) <> 0 Then RecOK = False
                        If My.Settings.PVv8 <> String.Empty AndAlso rec.v8 = "NaN" AndAlso CDbl(FiveMinData(My.Settings.PVv8)) <> 0 Then RecOK = False
                        If My.Settings.PVv9 <> String.Empty AndAlso rec.v9 = "NaN" AndAlso CDbl(FiveMinData(My.Settings.PVv9)) <> 0 Then RecOK = False
                        If My.Settings.PVv10 <> String.Empty AndAlso rec.v10 = "NaN" AndAlso CDbl(FiveMinData(My.Settings.PVv10)) <> 0 Then RecOK = False
                        If My.Settings.PVv11 <> String.Empty AndAlso rec.v11 = "NaN" AndAlso CDbl(FiveMinData(My.Settings.PVv11)) <> 0 Then RecOK = False
                        If My.Settings.PVv12 <> String.Empty AndAlso rec.v12 = "NaN" AndAlso CDbl(FiveMinData(My.Settings.PVv12)) <> 0 Then RecOK = False
                    End If
                Catch ex As Exception
                    FiveMinData = Nothing
                    RecOK = False
                End Try
            Else
                RecOK = False
                While CurrRecTimeStamp < FoundRecTimeStamp
                    SendPowerwallData(CurrRecTimeStamp)
                    CurrRecTimeStamp = DateAdd(DateInterval.Minute, 5, CurrRecTimeStamp)
                End While
            End If
            If Not RecOK Then
                SendPowerwallData(CurrRecTimeStamp, FiveMinData)
            End If
            CurrRecTimeStamp = DateAdd(DateInterval.Minute, 5, CurrRecTimeStamp)
        Next
    End Sub
    Private Function GetFiveMinuteData(AsAt As Date) As PWHistoryDataSet.spGet5MinuteAveragesRow
        Try
            SyncLock DBLock
                Return CType(Get5MinuteAveragesTA.GetData(AsAt).Rows(0), PWHistoryDataSet.spGet5MinuteAveragesRow)
            End SyncLock
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Private Sub SendPowerwallData(AsAt As DateTime, Optional FiveMinData As PWHistoryDataSet.spGet5MinuteAveragesRow = Nothing)
        Try
            If FiveMinData Is Nothing Then FiveMinData = GetFiveMinuteData(AsAt)
            If Not FiveMinData Is Nothing Then
                Dim Params As New List(Of ParamData)
                If CheckSunIsUp(AsAt) Then
                    If My.Settings.PVSendPV Then
                        Dim PV As Object = FiveMinData("solar_instant_power")
                        If Not IsDBNull(PV) AndAlso CSng(PV) > 0 Then
                            Params.Add(New ParamData With {.ParamName = "v2", .Data = PV.ToString})
                        Else
                            Params.Add(New ParamData With {.ParamName = "v2", .Data = 0.ToString})
                        End If
                    End If
                    If My.Settings.PVSendForecastAs <> String.Empty Then
                        Dim FC As Object = FiveMinData("solcast_forecast")
                        If Not IsDBNull(FC) AndAlso CSng(FC) > 0 Then
                            Params.Add(New ParamData With {.ParamName = My.Settings.PVSendForecastAs, .Data = FC.ToString})
                        Else
                            Params.Add(New ParamData With {.ParamName = My.Settings.PVSendForecastAs, .Data = 0.ToString})
                        End If
                    End If
                End If
                If My.Settings.PVSendLoad Then Params.Add(New ParamData With {.ParamName = "v4", .Data = FiveMinData("load_instant_power").ToString})
                If My.Settings.PVSendVoltage Then Params.Add(New ParamData With {.ParamName = "v6", .Data = FiveMinData("site_instant_average_voltage").ToString})
                If My.Settings.PVv7 <> String.Empty Then Params.Add(New ParamData With {.ParamName = "v7", .Data = FiveMinData(My.Settings.PVv7).ToString})
                If My.Settings.PVv8 <> String.Empty Then Params.Add(New ParamData With {.ParamName = "v8", .Data = FiveMinData(My.Settings.PVv8).ToString})
                If My.Settings.PVv9 <> String.Empty Then Params.Add(New ParamData With {.ParamName = "v9", .Data = FiveMinData(My.Settings.PVv9).ToString})
                If My.Settings.PVv10 <> String.Empty Then Params.Add(New ParamData With {.ParamName = "v10", .Data = FiveMinData(My.Settings.PVv10).ToString})
                If My.Settings.PVv11 <> String.Empty Then Params.Add(New ParamData With {.ParamName = "v11", .Data = FiveMinData(My.Settings.PVv11).ToString})
                If My.Settings.PVv12 <> String.Empty Then Params.Add(New ParamData With {.ParamName = "v12", .Data = FiveMinData(My.Settings.PVv12).ToString})
                PVSaveMultipleData(AsAt, Params)
            End If
        Catch ex As Exception
            EventLog.WriteEntry(ex.Message & vbCrLf & vbCrLf & ex.StackTrace, EventLogEntryType.Error)
        End Try
    End Sub
    Function PVSaveExtendedData(OutputDate As DateTime, ExParam As String, ExData As String) As Integer
        Try
            Dim request As WebRequest = WebRequest.Create(String.Format("https://pvoutput.org/service/r2/addstatus.jsp?d={0:yyyyMMdd}&t={0:HH:mm}&{1}={2}&sid={3}&key={4}", OutputDate, ExParam, ExData, My.Settings.PVOutputSID, My.Settings.PVOutputAPIKey))
            Dim response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
            PVSaveExtendedData = response.StatusCode
            response.Close()
        Catch Ex As Exception
            EventLog.WriteEntry(Ex.Message & vbCrLf & vbCrLf & Ex.StackTrace, EventLogEntryType.Error)
            Return 0
        End Try
    End Function
    Function PVGetStatusHistory(OutputDate As DateTime) As List(Of StatusRecord)
        Try
            Dim request As WebRequest = WebRequest.Create(String.Format("https://pvoutput.org/service/r2/getstatus.jsp?d={0:yyyyMMdd}&h=1&ext=1&asc=1&limit=288&sid={1}&key={2}", OutputDate, My.Settings.PVOutputSID, My.Settings.PVOutputAPIKey))
            Dim response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
            Dim dataStream As Stream = response.GetResponseStream()
            Dim reader As StreamReader = New StreamReader(dataStream)
            Dim responseFromServer As String = reader.ReadToEnd()
            reader.Close()
            response.Close()
            Dim records() As String = responseFromServer.Split(CChar(";"))
            Dim StatusRecords As New List(Of StatusRecord)
            For i = 0 To records.Count - 1
                StatusRecords.Add(New StatusRecord(records(i).Split(CChar(","))))
            Next
            PVGetStatusHistory = StatusRecords
        Catch Ex As Exception
            EventLog.WriteEntry(Ex.Message & vbCrLf & vbCrLf & Ex.StackTrace, EventLogEntryType.Error)
            Return Nothing
        End Try
    End Function
    Function PVSaveMultipleData(OutputDate As DateTime, Params As List(Of ParamData)) As Integer
        If Params.Count > 0 Then
            Try
                Dim BaseReqString As String = String.Format("https://pvoutput.org/service/r2/addstatus.jsp?d={0:yyyyMMdd}&t={0:HH:mm}&sid={1}&key={2}", OutputDate, My.Settings.PVOutputSID, My.Settings.PVOutputAPIKey)
                For Each Param As ParamData In Params
                    BaseReqString += String.Format("&{0}={1}", Param.ParamName, Param.Data)
                Next
                Dim request As WebRequest = WebRequest.Create(BaseReqString)
                Dim response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
                PVSaveMultipleData = response.StatusCode
                response.Close()
            Catch Ex As Exception
                EventLog.WriteEntry(Ex.Message & vbCrLf & vbCrLf & Ex.StackTrace, EventLogEntryType.Error)
                Return 0
            End Try
        Else
            Return 0
        End If
    End Function
#End Region
#Region "PowerBI Streaming"
    Function PostPowerBIStreamingData(Of JSONType)(Target As String, Data As JSONType) As Integer
        Try
            Dim BodyPostData As String = JsonConvert.SerializeObject(Data).ToString
            Dim BodyByteStream As Byte() = Encoding.UTF8.GetBytes(BodyPostData)
            Dim request As WebRequest = WebRequest.Create(Target)
            request.Method = "POST"
            request.ContentType = "application/json"
            request.ContentLength = BodyByteStream.Length
            Dim BodyStream As Stream = request.GetRequestStream()
            BodyStream.Write(BodyByteStream, 0, BodyByteStream.Length)
            BodyStream.Close()
            Dim response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
            PostPowerBIStreamingData = response.StatusCode
            response.Close()
        Catch Ex As Exception
            EventLog.WriteEntry(Ex.Message & vbCrLf & vbCrLf & Ex.StackTrace, EventLogEntryType.Error)
            Return 0
        End Try
    End Function
#End Region
#Region "Temperature Forecasts"
    Function GetWUTemperatureForecast(Of JSONType)() As JSONType
        Return GetUnsecuredJSONResult(Of JSONType)(String.Format(My.Settings.WUAddress, My.Settings.WUAPIKey, My.Settings.WULocation))
    End Function
    Function GetWWTemperatureForecast(Of JSONType)() As JSONType
        Return GetUnsecuredJSONResult(Of JSONType)(String.Format(My.Settings.WWAddress, My.Settings.WWAPIKey, My.Settings.WWLocation))
    End Function
    Sub GetWUTemperature()
        TempFCasts = New List(Of TemperaturePart)
        Dim WUForecast As WUForecast = GetWUTemperatureForecast(Of WUForecast)()
        For Each HF As Hourly_Forecast In WUForecast.hourly_forecast()
            TempFCasts.Add(New TemperaturePart With {.PartDate = New Date(CInt(HF.FCTTIME.year), CInt(HF.FCTTIME.mon), CInt(HF.FCTTIME.mday)), .PartHour = CInt(HF.FCTTIME.hour), .Temperature = CDec(HF.temp.metric)})
        Next
    End Sub
    Sub GetWWTemperature()
        TempFCasts = New List(Of TemperaturePart)
        Dim WWForecast As WWForecast = GetWWTemperatureForecast(Of WWForecast)()
        For Each Day As Day In WWForecast.forecasts.temperature.days()
            For Each Entry As Entry In Day.entries()
                TempFCasts.Add(New TemperaturePart With {.PartDate = Day.dateTime.Date, .PartHour = Entry.dateTime.Hour, .Temperature = Entry.temperature})
            Next
        Next
    End Sub
#End Region
#Region "Tariffs"
    Sub LoadTariffConfig()
        Try
            Dim TariffConfigFile As String = New StreamReader(AppContext.BaseDirectory & My.Settings.TariffConfigFile).ReadToEnd
            TariffDefinition = JsonConvert.DeserializeObject(Of Tariff)(TariffConfigFile)
            Dim MaxFIT As Decimal = 0
            Dim MaxCost As Decimal = 0
            For Each TariffItem As TariffItem In TariffDefinition.TariffItems()
                With TariffItem
                    If .FITRate > MaxFIT Then MaxFIT = .FITRate
                    If .ConsumptionRate > MaxCost Then MaxCost = .ConsumptionRate
                End With
            Next
            For Each TariffItem As TariffItem In TariffDefinition.TariffItems()
                With TariffItem
                    .MaxCost = MaxCost
                    .MaxFIT = MaxFIT
                End With
            Next
        Catch ex As Exception

        End Try
    End Sub
    Sub BuildTariffMap()
        TariffMap = New List(Of TariffPart)
        Dim TMD = From TIs In TariffDefinition.TariffItems
                  Where TIs.IsDefault = True
                  Select TIs.ConsumptionRate, TIs.FITRate, TIs.OffsetPriority, TIs.FITPriority, TIs.StandbyPreferred, TIs.IsDefault, TIs.Name, TIs.IsOffPeak
        For i As DayOfWeek = DayOfWeek.Sunday To DayOfWeek.Saturday
            For j = 0 To 23
                Dim il = i
                Dim jl = j
                Dim TMs = From TPs In TariffDefinition.TariffPeriods
                          Join TIs In TariffDefinition.TariffItems
                            On TPs.Tariff Equals TIs.Name
                          Where TPs.StartHour <= jl And
                               TPs.EndHour >= jl And
                               ((TPs.Sun And il = DayOfWeek.Sunday) Or
                                (TPs.Mon And il = DayOfWeek.Monday) Or
                                (TPs.Tue And il = DayOfWeek.Tuesday) Or
                                (TPs.Wed And il = DayOfWeek.Wednesday) Or
                                (TPs.Thu And il = DayOfWeek.Thursday) Or
                                (TPs.Fri And il = DayOfWeek.Friday) Or
                                (TPs.Sat And il = DayOfWeek.Saturday))
                          Order By ((jl - TPs.StartHour) + (TPs.EndHour - jl))
                          Select TPs.Cooling, TPs.Heating, TPs.LoadPercentage, TPs.Tariff, TIs.ConsumptionRate, TIs.FITRate, TIs.OffsetPriority, TIs.FITPriority, TIs.StandbyPreferred, TIs.IsDefault, TIs.IsOffPeak
                If TMs.Count > 0 Then
                    TariffMap.Add(New TariffPart With {.DOW = il, .Hour = jl, .Cooling = TMs(0).Cooling, .Heating = TMs(0).Heating, .LoadPercentage = TMs(0).LoadPercentage, .ConsumptionRate = TMs(0).ConsumptionRate, .FITRate = TMs(0).FITRate, .FITPriority = TMs(0).FITPriority, .OffsetPriority = TMs(0).OffsetPriority, .StandbyPreferred = TMs(0).StandbyPreferred, .Tariff = TMs(0).Tariff, .IsOffPeak = TMs(0).IsDefault})
                Else
                    TariffMap.Add(New TariffPart With {.DOW = il, .Hour = jl, .Cooling = False, .Heating = False, .LoadPercentage = 0, .ConsumptionRate = TMD(0).ConsumptionRate, .FITRate = TMD(0).FITRate, .FITPriority = TMD(0).FITPriority, .OffsetPriority = TMD(0).OffsetPriority, .StandbyPreferred = TMD(0).StandbyPreferred, .Tariff = TMD(0).Name, .IsOffPeak = TMD(0).IsOffPeak})
                End If
            Next j
        Next i
    End Sub
#End Region
End Class
