﻿#Region "Imports"
Imports System.Net
Imports System.IO
Imports System.Text
Imports Newtonsoft.Json
Imports PowerwallService.PWJson
Imports PowerwallService.SolCast
Imports PowerwallService.PVOutput
Imports PowerwallService.SunriseSunset
Imports PowerwallService.PowerBIStreaming
Imports TeslaAuth
#End Region
Public Class PowerwallService
#Region "Variables"
    Private Enum PWStatusEnum
        Charging
        Discharging
        Standby
    End Enum
    Const self_consumption As String = "self_consumption"
    Const backup As String = "backup"
    Const autonomous As String = "autonomous"
    Const AppMinCharge As Decimal = 5
    Const AppToLocalRatio As Decimal = CDec(95 / 100)
    Private DischargeMode As String = self_consumption
    Private ReadOnly ObsTA As New PWHistoryDataSetTableAdapters.observationsTableAdapter
    Private ReadOnly SolarTA As New PWHistoryDataSetTableAdapters.solarTableAdapter
    Private ReadOnly SOCTA As New PWHistoryDataSetTableAdapters.socTableAdapter
    Private ReadOnly BatteryTA As New PWHistoryDataSetTableAdapters.batteryTableAdapter
    Private ReadOnly LoadTA As New PWHistoryDataSetTableAdapters.loadTableAdapter
    Private ReadOnly SiteTA As New PWHistoryDataSetTableAdapters.siteTableAdapter
    Private ReadOnly CompactTA As New PWHistoryDataSetTableAdapters.CompactObsTableAdapter
    Private ReadOnly CompactTALocal As New PWHistoryDataSetTableAdapters.CompactObsTALocal
    Private ReadOnly Get5MinuteAveragesTA As New PWHistoryDataSetTableAdapters.spGet5MinuteAveragesTableAdapter
    Private ReadOnly SPs As New PWHistoryDataSetTableAdapters.SPs
    Private ReadOnly SixSecondTimer As New Timers.Timer
    Private ReadOnly OneMinuteTime As New Timers.Timer
    Private ReadOnly FiveMinuteTimer As New Timers.Timer
    Private ReadOnly TenMinuteTimer As New Timers.Timer
    Private ReadOnly DailyTimer As New Timers.Timer
    Shared PWLocalCookies As CookieCollection
    Shared PWLocalToken As String = String.Empty
    Shared PWCloudToken As String = String.Empty
    Shared PWCloudRefreshToken As String = String.Empty
    Shared NextDayForecastGeneration As Single = 0
    Shared NextDayMorningGeneration As Single = 0
    Shared MeterReading As MeterAggregates
    Shared SOC As New SOC With {.percentage = 5}
    Shared CurrentDayForecast As DayForecast
    Shared NextDayForecast As DayForecast
    Shared SecondDayForecast As DayForecast
    Shared ForecastsRetrieved As Date = DateAdd(DateInterval.Hour, -2, Now)
    Shared FirstReadingsAvailable As Boolean = False
    Shared PreCharging As Boolean = False
    Shared OnStandby As Boolean = False
    Shared AboveMinBackup As Boolean = False
    Shared LastTarget As Decimal = 0
    Shared OffPeakStart As DateTime
    Shared PeakStart As DateTime
    Shared OffPeakStartHour As Integer
    Shared PeakStartHour As Integer
    Shared OffPeakHours As Double
    Shared OperationLockout As DateTime = DateAdd(DateInterval.Hour, -2, Now)
    Shared CurrentDayAllOffPeak As Boolean
    Shared NextDayAllDayOffPeak As Boolean
    Shared LastPeriodForecast As Forecast
    Shared CurrentPeriodForecast As Forecast
    Shared PVForecast As OutputForecast
    Shared CurrentDOW As DayOfWeek
    Shared CivilTwilightSunrise As DateTime
    Shared CivilTwilightSunset As DateTime
    Shared Sunrise As DateTime
    Shared Sunset As DateTime
    Shared AsAtSunrise As Result
    Shared ReadOnly DBLock As New Object
    Shared ReadOnly PWLock As New Object
    Shared SkipObservation As Boolean = False
    Shared PWStatus As PWStatusEnum = PWStatusEnum.Standby
    Shared ReadOnly CurrentChargeSettings As New Operation With {.real_mode = self_consumption, .backup_reserve_percent = 0}
    Shared PWCloudEnergyID As Long
    Shared PWCloudSiteID As String
#End Region
#Region "Timer Handlers"
    Protected Async Sub OnSixSecondTimer(Sender As Object, Args As System.Timers.ElapsedEventArgs)
        Await Task.Run(Sub()
                           GetObservationAndStore()
                       End Sub)
    End Sub
    Protected Async Sub OnOneMinuteTime(Sender As Object, Args As System.Timers.ElapsedEventArgs)
        Await Task.Run(Sub()
                           DoPerMinuteTasks()
                       End Sub)
    End Sub
    Protected Async Sub OnFiveMinuteTimer(Sender As Object, Args As System.Timers.ElapsedEventArgs)
        Await Task.Run(Sub()
                           DoFiveMinuteTasks()
                       End Sub)
    End Sub
    Protected Async Sub OnTenMinuteTimer(Sender As Object, Args As System.Timers.ElapsedEventArgs)
        Await Task.Run(Sub()
                           DoTenMinuteTasks()
                       End Sub)
    End Sub
    Protected Async Sub OnDailyTimer(Sender As Object, Args As System.Timers.ElapsedEventArgs)
        Await Task.Run(Sub()
                           DoDailyTasks()
                       End Sub)
    End Sub
    Protected Async Sub DebugTask()
        Await Task.Run(Sub()
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

        SixSecondTimer.Interval = 6 * 1000 ' Every Six Seconds
        SixSecondTimer.AutoReset = True
        AddHandler SixSecondTimer.Elapsed, AddressOf OnSixSecondTimer

        OneMinuteTime.Interval = 60 * 1000 ' Every Minute
        OneMinuteTime.AutoReset = True
        AddHandler OneMinuteTime.Elapsed, AddressOf OnOneMinuteTime

        If My.Settings.PVReportingEnabled Then
            FiveMinuteTimer.Interval = 5 * 60 * 1000 ' Every Five Minutes
            FiveMinuteTimer.AutoReset = True
            AddHandler FiveMinuteTimer.Elapsed, AddressOf OnFiveMinuteTimer
        End If

        TenMinuteTimer.Interval = 60 * 10 * 1000 ' Every 10 Minutes
        TenMinuteTimer.AutoReset = True
        AddHandler TenMinuteTimer.Elapsed, AddressOf OnTenMinuteTimer

        DailyTimer.Interval = 60 * 1000 * 60 * 24 ' Every 24 Hours
        DailyTimer.AutoReset = True
        AddHandler DailyTimer.Elapsed, AddressOf OnDailyTimer

        Task.Run(Sub()
                     DoAsyncStartupProcesses()
                 End Sub)

        EventLog.WriteEntry(String.Format("Powerwall Service version {0} started", My.Application.Info.Version.ToString), EventLogEntryType.Information, 101)
    End Sub
    Private Sub DoAsyncStartupProcesses()
        Threading.Thread.Sleep(10000)
        SetOffPeakHours(Now)
        If My.Settings.PWUseAutonomous Then
            DischargeMode = autonomous
        End If
        PWCloudToken = LoginPWCloud()
        PWLocalToken = LoginPWLocalUser(ForceReLogin:=True)
        GetCloudProducts()
        GetCloudPWMode()
        If My.Settings.PWForceModeOnStartup Then
            Dim Intent As String = "Thinking"
            Dim APIResult As Integer = SetPWMode("Execute Force Startup Mode", "Enter", My.Settings.PWForceMode.ToString, My.Settings.PWForceBackupPercentage, My.Settings.PWForceMode, Intent)
            If APIResult = 202 Then
                EventLog.WriteEntry(String.Format("Forced PW Mode on Startup to: Mode={0}, BackupPercentage={1}, APIResult = {2}", My.Settings.PWForceMode, My.Settings.PWForceBackupPercentage, APIResult), EventLogEntryType.Information, 800)
            Else
                EventLog.WriteEntry(String.Format("Failed to Force PW Mode on Startup: Mode={0}, BackupPercentage={1}, APIResult = {2}", My.Settings.PWForceMode, My.Settings.PWForceBackupPercentage, APIResult), EventLogEntryType.Warning, 801)
            End If
        End If

        Try
            GetForecasts()
        Catch ex As Exception
            EventLog.WriteEntry(String.Format("Failed to get forecasts: Exception: {0}, Stack Trace: {1}", ex.Message, ex.StackTrace), EventLogEntryType.Error, 802)
        End Try


        Task.Run(Sub()
                     SleepUntilSecBoundary(6)
                     SixSecondTimer.Start()
                     EventLog.WriteEntry("Six Second (Observation) Timer Started", EventLogEntryType.Information, 108)
                 End Sub)

        Task.Run(Sub()
                     SleepUntilSecBoundary(60)
                     OneMinuteTime.Start()
                     DoPerMinuteTasks()
                     EventLog.WriteEntry("One Minute Timer Started", EventLogEntryType.Information, 109)
                 End Sub)
    End Sub
    Protected Overrides Sub OnContinue()
        EventLog.WriteEntry("Powerwall Service Resuming", EventLogEntryType.Information, 102)
        SixSecondTimer.Start()
        OneMinuteTime.Start()
        Task.Run(Sub()
                     DebugTask()
                 End Sub
                 )
        EventLog.WriteEntry("Powerwall Service Running", EventLogEntryType.Information, 103)
    End Sub
    Protected Overrides Sub OnPause()
        EventLog.WriteEntry("Powerwall Service Pausing", EventLogEntryType.Information, 104)
        SixSecondTimer.Stop()
        OneMinuteTime.Stop()
        If My.Settings.PVReportingEnabled Then FiveMinuteTimer.Stop()
        TenMinuteTimer.Stop()
        EventLog.WriteEntry("Powerwall Service Paused", EventLogEntryType.Information, 105)
    End Sub
    Protected Overrides Sub OnStop()
        EventLog.WriteEntry("Powerwall Service Stopping", EventLogEntryType.Information, 106)
        SixSecondTimer.Stop()
        SixSecondTimer.Dispose()
        OneMinuteTime.Stop()
        OneMinuteTime.Dispose()
        If My.Settings.PVReportingEnabled Then
            FiveMinuteTimer.Stop()
            FiveMinuteTimer.Dispose()
        End If
        TenMinuteTimer.Stop()
        TenMinuteTimer.Dispose()
        DailyTimer.Stop()
        DailyTimer.Dispose()
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
        If My.Settings.TariffIgnoresDST And TZI.IsDaylightSavingTime(InvokedTime) And Not CurrentDayAllOffPeak Then
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
        If InvokedTime.Hour >= 0 And InvokedTime.Hour < PeakStartHour + 1 And OffPeakStartsBeforeMidnight And Not CurrentDayAllOffPeak Then
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
    Sub DoPerMinuteTasks()
        Dim Minute As Integer = Now.Minute
        If Minute Mod 10 = 6 Or Minute Mod 10 = 1 Then
            If My.Settings.PVReportingEnabled Then
                If FiveMinuteTimer.Enabled = False Then
                    FiveMinuteTimer.Start()
                    EventLog.WriteEntry("Five Minute (PVOutput & Mode Check) Timer Started", EventLogEntryType.Information, 110)
                    DoFiveMinuteTasks()
                End If
            End If
        End If
        If Minute Mod 10 = 1 Then
            If TenMinuteTimer.Enabled = False Then
                TenMinuteTimer.Start()
                EventLog.WriteEntry("Ten Minute (Solar Forecast & Charge Monitoring) Timer Started", EventLogEntryType.Information, 111)
                DoTenMinuteTasks()
            End If
            If DailyTimer.Enabled = False Then
                DailyTimer.Start()
                EventLog.WriteEntry("Daily Timer Started", EventLogEntryType.Information, 112)
            End If
        End If
        AggregateToMinute()
    End Sub
    Private Sub DoFiveMinuteTasks()
        If My.Settings.PVReportingEnabled Then
            If My.Settings.PVSendPowerwall Then
                SendPowerwallData(Now)
                If Now.Minute < 6 Then
                    If Now.Hour < 3 Then
                        DoBackFill(DateAdd(DateInterval.Day, -1, Now))
                    Else
                        DoBackFill(Now)
                    End If
                End If
            ElseIf My.Settings.PVSendForecast Then
                SendForecast()
            End If
        End If
    End Sub
    Private Sub DoTenMinuteTasks()
        SetOffPeakHours(Now)
        If Not FirstReadingsAvailable Then
            GetObservationAndStore()
        End If
        GetForecasts()
        If My.Settings.PWControlEnabled Then
            CheckSOCLevel()
        End If
    End Sub
    Private Sub DoDailyTasks()
        RefreshTokens()
    End Sub
    Function GetUnsecuredJSONResult(Of JSONType)(URL As String) As JSONType
        Dim response As HttpWebResponse
        Try
            Dim request As WebRequest = WebRequest.Create(URL)
            Try
                response = CType(request.GetResponse(), HttpWebResponse)
            Catch ex As WebException
                EventLog.WriteEntry(ex.Message & vbCrLf & vbCrLf & ex.StackTrace, EventLogEntryType.Error)
                Return Nothing
            Catch ex As Exception
                Return Nothing
            End Try
            If response IsNot Nothing Then
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
            If AsAt > CivilTwilightSunrise And AsAt < CivilTwilightSunset Then Return True
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
#Region "Forecasts and Targets"
    Sub CheckSOCLevel()
        Dim InvokedTime As DateTime = Now
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
        Dim NewTarget As Decimal = 0
        If InvokedTime > Sunrise And InvokedTime < Sunset And InvokedTime < PeakStart Then
            RemainingOvernightRatio = 0
            RemainingInsolationToday = CurrentDayForecast.PVEstimate
            ForecastInsolationTomorrow = NextDayForecastGeneration
            ShortfallInsolation = 0
            Intent = "Sun is Up, Waiting for Peak"
        ElseIf InvokedTime > Sunrise And InvokedTime < Sunset Then
            RemainingOvernightRatio = 1
            RemainingInsolationToday = CurrentDayForecast.PVEstimate
            ForecastInsolationTomorrow = NextDayForecastGeneration
            ShortfallInsolation = 0
            Intent = "Sun is Up, In Peak"
        ElseIf InvokedTime > PeakStart And InvokedTime < Sunrise Then
            RemainingOvernightRatio = 0
            RemainingInsolationToday = CurrentDayForecast.PVEstimate
            ForecastInsolationTomorrow = NextDayForecastGeneration
            ShortfallInsolation = PWPeakConsumption - NextDayForecastGeneration
            Intent = "Waiting for Sunrise, In Peak"
        ElseIf InvokedTime > Sunset And InvokedTime < OffPeakStart Then
            RemainingOvernightRatio = 1
            RemainingInsolationToday = 0
            ForecastInsolationTomorrow = NextDayForecastGeneration
            ShortfallInsolation = PWPeakConsumption - NextDayForecastGeneration
            Intent = "Sun is Down, Waiting for Off Peak"
            'ElseIf InvokedTime > Sunset And InvokedTime > OffPeakStart Then ' Suspect
            '    RemainingOvernightRatio = 1
            '    RemainingInsolationToday = 0
            '    ForecastInsolationTomorrow = NextDayForecastGeneration
            '    ShortfallInsolation = PWPeakConsumption - NextDayForecastGeneration
            '    Intent = "Sun is Down, Off Peak"
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
        NewTarget = CDec(IIf(My.Settings.PWOvernightStandby, StandbyTargetSOC, NoStandbyTargetSOC))
        If InvokedTime > Sunset And InvokedTime < OffPeakStart Then
            If ShortfallInsolation > 0 Or NewTarget > SOC.percentage Then Intent = "Planning to Charge" Else Intent = "No Charging Required"
        End If
        Try
            If InvokedTime > DateAdd(DateInterval.Minute, -15, PeakStart) And Not CurrentDayAllOffPeak And (PreCharging Or AboveMinBackup Or OnStandby) Then
                OperationLockout = PeakStart
                EventLog.WriteEntry(String.Format("Reaching end of off-peak period with SOC={0}, was aiming for Target={1}", SOC.percentage, StandbyTargetSOC), EventLogEntryType.Information, 504)
                DoExitCharging(Intent)
            ElseIf (InvokedTime >= OffPeakStart And InvokedTime < PeakStart And InvokedTime > OperationLockout) Then
                If My.Settings.VerboseLogging Then EventLog.WriteEntry(String.Format("In Operation Period: Current SOC={0}, Minimum required at end of Off-Peak={1}, Shortfall Generation Tomorrow={2}, As at now, Charge Target={3}", SOC.percentage, RawTargetSOC, ShortfallInsolation, NoStandbyTargetSOC), EventLogEntryType.Information, 500)
                If My.Settings.DebugLogging Then EventLog.WriteEntry(String.Format("In Operation Period: Invoked={0:yyyy-MM-dd HH:mm}, OperationStart={1:yyyy-MM-dd HH:mm}, OperationEnd={2:yyyy-MM-dd HH:mm}", InvokedTime, OffPeakStart, PeakStart), EventLogEntryType.Information, 714)
                If NextDayAllDayOffPeak And (My.Settings.PWOvernightStandby Or My.Settings.PWWeekendStandbyOnTarget) And (Not OnStandby Or SOC.percentage > LastTarget) Then
                    If SetPWMode("Switching to Standby for Off Peak, Standby Mode Enabled", "Enter", "Standby", SOC.percentage, DischargeMode, Intent) = 202 Then
                        OnStandby = True
                        PreCharging = False
                    End If
                ElseIf Not My.Settings.PWOvernightStandby And SOC.percentage >= StandbyTargetSOC And SOC.percentage <= NoStandbyTargetSOC And Not PreCharging And (Not OnStandby Or SOC.percentage > LastTarget) Then
                    If SetPWMode("Morning SOC would be below required morning SOC, Standby Mode Not Enabled", "Enter", "Standby", SOC.percentage, DischargeMode, Intent) = 202 Then
                        OnStandby = True
                        PreCharging = False
                    End If
                ElseIf My.Settings.PWOvernightStandby And SOC.percentage >= StandbyTargetSOC And Not PreCharging And (Not OnStandby Or SOC.percentage > LastTarget) Then
                    If SetPWMode("Current SOC above required morning SOC, Standby Mode Enabled", "Enter", "Standby", SOC.percentage, DischargeMode, Intent) = 202 Then
                        OnStandby = True
                        PreCharging = False
                    End If
                ElseIf My.Settings.PWWeekendStandbyOnTarget And SOC.percentage >= My.Settings.PWWeekendStandbyTarget And CurrentDayAllOffPeak And Not PreCharging And (Not OnStandby Or SOC.percentage > LastTarget) Then
                    If SetPWMode("Current SOC above weekend target, standby on target enabled", "Enter", "Standby", SOC.percentage, DischargeMode, Intent) = 202 Then
                        OnStandby = True
                        PreCharging = False
                    End If
                ElseIf (SOC.percentage < StandbyTargetSOC And OnStandby And Not (CurrentDayAllOffPeak Or NextDayAllDayOffPeak)) Or (SOC.percentage < NoStandbyTargetSOC And Not OnStandby And Not PreCharging And Not (CurrentDayAllOffPeak Or NextDayAllDayOffPeak)) Then
                    If My.Settings.VerboseLogging Then EventLog.WriteEntry(String.Format("Current SOC below required setting: Current SOC={0}, Required at end of Off-Peak={1}, Shortfall Generation Tomorrow={2}, As at now, Charge Target={3}", SOC.percentage, StandbyTargetSOC, ShortfallInsolation, NewTarget), EventLogEntryType.Information, 501)
                    If SetPWMode("Current SOC below required morning SOC", "Enter", IIf(NewTarget > (SOC.percentage + 5), "Charging", "Standby").ToString, NewTarget, IIf(My.Settings.PWChargeModeBackup, backup, DischargeMode).ToString, Intent) = 202 Then
                        PreCharging = True
                        OnStandby = False
                    End If
                ElseIf SOC.percentage >= NoStandbyTargetSOC And PreCharging And Not OnStandby And My.Settings.PWOvernightStandby Then
                    EventLog.WriteEntry(String.Format("Current SOC above required setting and Standby Mode Enabled: Current SOC={0}, Required at end of Off-Peak={1}, Shortfall Generation Tomorrow={2}, As at now, Charge Target={3}", SOC.percentage, RawTargetSOC, ShortfallInsolation, NoStandbyTargetSOC), EventLogEntryType.Information, 505)
                    If SetPWMode("Switching to Standby for Off Peak, Standby Mode Enabled", "Enter", "Standby", SOC.percentage, DischargeMode, Intent) = 202 Then
                        OnStandby = True
                        PreCharging = False
                    End If
                ElseIf (LastTarget < NewTarget And OnStandby And Not (CurrentDayAllOffPeak Or NextDayAllDayOffPeak)) Or (LastTarget < NewTarget And PreCharging) Then
                    If My.Settings.VerboseLogging Then EventLog.WriteEntry(String.Format("Charge Target Increased & SOC below required setting: Current SOC={0}, Required at end of Off-Peak={1}, Shortfall Generation Tomorrow={2}, As at now, Charge Target={3}", SOC.percentage, StandbyTargetSOC, ShortfallInsolation, NewTarget), EventLogEntryType.Information, 514)
                    If SetPWMode("Current SOC below required morning SOC", "Enter", IIf(NewTarget > (SOC.percentage + 5), "Charging", "Standby").ToString, NewTarget, IIf(My.Settings.PWChargeModeBackup, backup, DischargeMode).ToString, Intent) = 202 Then
                        PreCharging = True
                        OnStandby = False
                    End If
                ElseIf SOC.percentage >= NoStandbyTargetSOC And PreCharging And Not OnStandby Then
                    EventLog.WriteEntry(String.Format("Current SOC above required setting: Current SOC={0}, Required at end of Off-Peak={1}, Shortfall Generation Tomorrow={2}, As at now, Charge Target={3}", SOC.percentage, RawTargetSOC, ShortfallInsolation, NoStandbyTargetSOC), EventLogEntryType.Information, 502)
                    DoExitCharging(Intent)
                End If
            Else
                If My.Settings.VerboseLogging Then EventLog.WriteEntry(String.Format("Outside Operation Period: SOC={0}", SOC.percentage), EventLogEntryType.Information, 503)
            End If
        Catch Ex As Exception
            EventLog.WriteEntry(Ex.Message & vbCrLf & vbCrLf & Ex.StackTrace, EventLogEntryType.Error, 510)
        End Try
        If My.Settings.PBIChargeIntentEndpoint <> String.Empty Then
            Try
                Dim PBIRows As New PBIChargeLogging With {.Rows = New List(Of ChargePlan)}
                PBIRows.Rows.Add(New ChargePlan With {.AsAt = InvokedTime, .CurrentSOC = SOC.percentage, .RemainingInsolation = RemainingInsolationToday, .ForecastGeneration = ForecastInsolationTomorrow, .MorningBuffer = My.Settings.PWMorningBuffer, .OperatingIntent = Intent, .RequiredSOC = NewTarget, .RemainingOffPeak = RemainingOffPeak * My.Settings.PWCapacity / 100, .Shortfall = ShortfallInsolation})
                Dim PowerBIPostResult As Integer = PostPowerBIStreamingData(My.Settings.PBIChargeIntentEndpoint, PBIRows)
            Catch ex As Exception
                EventLog.WriteEntry(String.Format("Failed to write Charge Plan to Power BI: Ex:{0} ({1})", ex.GetType, ex.Message), EventLogEntryType.Warning, 911)
            End Try
        End If
    End Sub
    Sub GetForecasts()
        Dim InvokedTime As DateTime = Now
        Try
            If CurrentDayForecast Is Nothing Then ' No Forecast retrieved - let's initialise with empty.
                CurrentDayForecast = New DayForecast With {.ForecastDate = DateAdd(DateInterval.Day, 0, Now.Date), .PVEstimate = 0, .MorningForecast = 0}
                NextDayForecast = New DayForecast With {.ForecastDate = DateAdd(DateInterval.Day, 1, Now.Date), .PVEstimate = 0, .MorningForecast = 0}
                SecondDayForecast = New DayForecast With {.ForecastDate = DateAdd(DateInterval.Day, 2, Now.Date), .PVEstimate = 0, .MorningForecast = 0}
            End If
            Dim NewForecastsRetrieved As Boolean = False
            If DateAdd(DateInterval.Hour, 1, ForecastsRetrieved) < InvokedTime And InvokedTime.Hour < (Sunset.Hour + 1) Then
                PVForecast = GetSolCastResult(Of OutputForecast)()
                If PVForecast IsNot Nothing Then
                    ForecastsRetrieved = InvokedTime
                    NewForecastsRetrieved = True
                    CurrentDayForecast = New DayForecast With {.ForecastDate = DateAdd(DateInterval.Day, 0, Now.Date), .PVEstimate = 0, .MorningForecast = 0}
                    NextDayForecast = New DayForecast With {.ForecastDate = DateAdd(DateInterval.Day, 1, Now.Date), .PVEstimate = 0, .MorningForecast = 0}
                    SecondDayForecast = New DayForecast With {.ForecastDate = DateAdd(DateInterval.Day, 2, Now.Date), .PVEstimate = 0, .MorningForecast = 0}
                End If
            End If
            If My.Settings.PVSendForecast And NewForecastsRetrieved Then
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
            End If
        Catch Ex As Exception
            EventLog.WriteEntry(Ex.Message & vbCrLf & vbCrLf & Ex.StackTrace, EventLogEntryType.Error)
            NextDayForecastGeneration = 0
            NextDayMorningGeneration = 0
        End Try
    End Sub
    Private Sub SendForecast()
        Dim InvokedTime As DateTime = Now
        If CurrentPeriodForecast Is Nothing Or LastPeriodForecast Is Nothing Then
            GetForecasts()
        End If
        If CurrentPeriodForecast IsNot Nothing And LastPeriodForecast IsNot Nothing Then
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
        If MeterReading IsNot Nothing Then
            LastObservationTime = MeterReading.site.last_communication_time
        End If
        Dim GotResults As Boolean = False
        Try
            SyncLock PWLock
                Try
                    If Not SkipObservation Then
                        MeterReading = GetLocalPWAPIResult(Of MeterAggregates)("meters/aggregates")
                        SOC = GetLocalPWAPIResult(Of SOC)("system_status/soe")
                    End If
                    If MeterReading IsNot Nothing Then
                        ObservationTime = MeterReading.site.last_communication_time
                        Select Case MeterReading.battery.instant_power
                            Case < -25
                                PWStatus = PWStatusEnum.Charging
                            Case > 25
                                PWStatus = PWStatusEnum.Discharging
                            Case Else
                                PWStatus = PWStatusEnum.Standby
                        End Select
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
                    EventLog.WriteEntry(String.Format("Failed to write Six Second Observations to Power BI: Ex:{0} ({1})", ex.GetType, ex.Message), EventLogEntryType.Warning, 910)
                End Try
            End If
        Catch Ex As Exception
            EventLog.WriteEntry(Ex.Message & vbCrLf & vbCrLf & Ex.StackTrace, EventLogEntryType.Error)
        End Try
    End Sub
    Function GetLocalPWAPIResult(Of JSONType)(API As String) As JSONType
        Return GetLocalLoggedInJSONResult(Of JSONType)(My.Settings.PWGatewayAddress & "/api/" & API)
    End Function
#End Region
#Region "Powerwall Control"
    Private Function SetPWMode(ActionMessage As String, ActionMode As String, ActionType As String, Target As Decimal, Mode As String, ByRef Intent As String) As Integer
        SkipObservation = True
        LastTarget = Target
        Target = (Target - AppMinCharge) / AppToLocalRatio ' Convert to Cloud Target from local SOC target calcuted by charge planning routine
        If Target > 100 Then Target = 100
        If Target < 0 Then Target = 0
        Try
            If My.Settings.VerboseLogging Then EventLog.WriteEntry(String.Format(ActionMessage & " Current SOC={0}, Current Target={1}", SOC.percentage, Target), EventLogEntryType.Information, 511)
            Intent = ActionType
            Dim ChargeSettings As New Operation With {.backup_reserve_percent = Target, .real_mode = Mode}
            Dim APIResult As Integer = DoSetPWModeCloudAPICalls(ChargeSettings)
            If APIResult = 202 Or APIResult = 200 Then
                EventLog.WriteEntry(String.Format("{5}ed {6} Mode: Current SOC={0}, Raw Target={1}, Set Mode={2}, API Call Target={3}, APIResult = {4}", SOC.percentage, LastTarget, ChargeSettings.real_mode, ChargeSettings.backup_reserve_percent, APIResult, ActionMode, ActionType), EventLogEntryType.Information, 512)
                AboveMinBackup = (ChargeSettings.backup_reserve_percent > My.Settings.PWMinBackupPercentage)
                Intent = ActionType
                SetPWMode = 202 ' Calls to SetPWMode expect APIResult of 202 as per behaviour for FW 1.42 and earlier
            Else
                EventLog.WriteEntry(String.Format("Failed to {5} {6} Mode: Current SOC={0}, Raw Target={1}, Mode={2}, API Call Target={3}, APIResult = {4}", SOC.percentage, LastTarget, ChargeSettings.real_mode, ChargeSettings.backup_reserve_percent, APIResult, ActionMode, ActionType), EventLogEntryType.Warning, 513)
                Intent = String.Format("Trying to {0} {1}", ActionMode, ActionType)
                SetPWMode = APIResult
            End If
        Catch ex As Exception
            SetPWMode = 0
        End Try
        SkipObservation = False
    End Function
    Private Function DoSetPWModeCloudAPICalls(ChargeSettings As Operation) As Integer
        Dim BackupSetting As New CloudBackup With {.backup_reserve_percent = ChargeSettings.backup_reserve_percent}
        Dim OperationSetting As New CloudOperation With {.default_real_mode = ChargeSettings.real_mode}
        Dim ModeResponse As CloudAPIResponse = PostPWCloudAPISettings("energy_sites/" & PWCloudEnergyID.ToString.Trim & "/backup", BackupSetting)
        Dim PercentageResponse As CloudAPIResponse = PostPWCloudAPISettings("energy_sites/" & PWCloudEnergyID.ToString.Trim & "/operation", OperationSetting)
        If ModeResponse.response.code >= 200 And ModeResponse.response.code <= 202 And PercentageResponse.response.code >= 200 And PercentageResponse.response.code <= 202 Then
            DoSetPWModeCloudAPICalls = 202
        Else
            DoSetPWModeCloudAPICalls = 500
        End If
    End Function
    Private Shared Function GetPWRequest(API As String) As HttpWebRequest
        Dim wr As HttpWebRequest
        wr = CType(WebRequest.Create(My.Settings.PWGatewayAddress & "/api/" & API), HttpWebRequest)
        wr.ServerCertificateValidationCallback = Function()
                                                     Return True
                                                 End Function
        Return wr
    End Function
    Private Shared Function GetCloudPWRequest(API As String) As HttpWebRequest
        Dim wr As HttpWebRequest
        wr = CType(WebRequest.Create(My.Settings.PWCloudAPI & API), HttpWebRequest)
        Return wr
    End Function
    Private Sub GetCloudPWMode()
        Dim SiteInfoResults As CloudSiteInfo
        Try
            If Not PreCharging Then
                SiteInfoResults = GetPWCloudAPIResult(Of CloudSiteInfo)("energy_sites/" & PWCloudEnergyID & "/site_info")
                CurrentChargeSettings.real_mode = SiteInfoResults.response.default_real_mode
                CurrentChargeSettings.backup_reserve_percent = SiteInfoResults.response.backup_reserve_percent
                EventLog.WriteEntry(String.Format("Current PW Mode={0}, BackupPercentage={1}", CurrentChargeSettings.real_mode, CurrentChargeSettings.backup_reserve_percent), EventLogEntryType.Information, 602)
                AboveMinBackup = (CurrentChargeSettings.backup_reserve_percent > My.Settings.PWMinBackupPercentage)
            End If
        Catch ex As Exception
            EventLog.WriteEntry(String.Format("Error Getting Current PW Mode: {0}, {1}", ex.Message, ex.StackTrace), EventLogEntryType.Warning, 1602)
        End Try
    End Sub
    Private Sub GetCloudProducts()
        Dim ListProductResult As ListProducts
        Dim FoundEnergyID As Long
        Dim ConfigEnergyID As Long = My.Settings.PWEnergySiteID
        Try
            ListProductResult = GetPWCloudAPIResult(Of ListProducts)("products")
            FoundEnergyID = ListProductResult.response(0).energy_site_id
            If ConfigEnergyID = 0 And FoundEnergyID <> 0 Then
                PWCloudEnergyID = FoundEnergyID
            ElseIf ConfigEnergyID = FoundEnergyID And ConfigEnergyID <> 0 Then
                PWCloudEnergyID = FoundEnergyID
            ElseIf ConfigEnergyID <> 0 Then
                PWCloudEnergyID = ConfigEnergyID
            Else
                PWCloudEnergyID = FoundEnergyID
            End If
            PWCloudSiteID = ListProductResult.response(0).id
            EventLog.WriteEntry(String.Format("Site ID={0}, Powerwall ID={1}", PWCloudEnergyID, PWCloudSiteID), EventLogEntryType.Information, 603)
        Catch ex As Exception
            EventLog.WriteEntry(String.Format("Error Getting Site ID: {0}, {1}", ex.Message, ex.StackTrace), EventLogEntryType.Warning, 1603)
        End Try
    End Sub
    Function GetPWCloudAPIResult(Of JSONType)(API As String, Optional ForceReLogin As Boolean = False) As JSONType
        Try
            PWCloudToken = LoginPWCloud(ForceReLogin:=ForceReLogin)
            Dim request As WebRequest = WebRequest.Create(My.Settings.PWCloudAPI & "api/1/" & API)
            request.Headers.Add("Authorization", "Bearer " & PWCloudToken)
            Dim response As HttpWebResponse
            Try
                response = CType(request.GetResponse(), HttpWebResponse)
            Catch WebEx As WebException
                Dim ExResponseCode As HttpStatusCode = CType(WebEx.Response, HttpWebResponse).StatusCode
                If ExResponseCode = 401 Or ExResponseCode = 403 Then
                    PWCloudToken = LoginPWCloud(ForceReLogin:=True)
                    request.Headers.Set("Authorization", "Bearer " & PWCloudToken)
                    response = CType(request.GetResponse(), HttpWebResponse)
                Else
                    EventLog.WriteEntry(String.Format("Unexpected error calling API {0} with response status code: {1}", API, ExResponseCode), EventLogEntryType.Error, 903)
                    EventLog.WriteEntry(WebEx.Message & vbCrLf & vbCrLf & WebEx.StackTrace, EventLogEntryType.Error)
                    Throw WebEx
                End If
            End Try
            Dim dataStream As Stream = response.GetResponseStream()
            Dim reader As StreamReader = New StreamReader(dataStream)
            Dim responseFromServer As String = reader.ReadToEnd()
            GetPWCloudAPIResult = JsonConvert.DeserializeObject(Of JSONType)(responseFromServer)
            reader.Close()
            response.Close()
        Catch Ex As Exception
            EventLog.WriteEntry(Ex.Message & vbCrLf & vbCrLf & Ex.StackTrace, EventLogEntryType.Error)
        End Try
    End Function
    Function PostPWCloudAPISettings(Of JSONType)(API As String, Settings As JSONType, Optional ForceReLogin As Boolean = False) As CloudAPIResponse
        Try
            PWCloudToken = LoginPWCloud(ForceReLogin:=ForceReLogin)
            Dim BodyPostData As String = JsonConvert.SerializeObject(Settings).ToString
            Dim BodyByteStream As Byte() = Encoding.UTF8.GetBytes(BodyPostData)
            Dim request As WebRequest = GetCloudPWRequest("api/1/" & API)
            request.Headers.Add("Authorization", "Bearer " & PWCloudToken)
            request.Method = "POST"
            request.ContentType = "application/json"
            request.ContentLength = BodyByteStream.Length
            Dim BodyStream As Stream = request.GetRequestStream()
            BodyStream.Write(BodyByteStream, 0, BodyByteStream.Length)
            BodyStream.Close()
            Dim response As HttpWebResponse
            Try
                response = CType(request.GetResponse(), HttpWebResponse)
            Catch WebEx As WebException
                Dim ExResponseCode As HttpStatusCode = CType(WebEx.Response, HttpWebResponse).StatusCode
                If ExResponseCode = 401 Or ExResponseCode = 403 Then
                    PWCloudToken = LoginPWCloud(ForceReLogin:=True)
                    request.Headers.Set("Authorization", "Bearer " & PWCloudToken)
                    response = CType(request.GetResponse(), HttpWebResponse)
                Else
                    EventLog.WriteEntry(String.Format("Unexpected error calling API {0} with settings {1} with response status code: {2}", API, JsonConvert.SerializeObject(Settings).ToString, ExResponseCode), EventLogEntryType.Error, 904)
                    EventLog.WriteEntry(WebEx.Message & vbCrLf & vbCrLf & WebEx.StackTrace, EventLogEntryType.Error)
                    Throw WebEx
                End If
            End Try
            Dim dataStream As Stream = response.GetResponseStream()
            Dim reader As StreamReader = New StreamReader(dataStream)
            Dim responseFromServer As String = reader.ReadToEnd()
            PostPWCloudAPISettings = JsonConvert.DeserializeObject(Of CloudAPIResponse)(responseFromServer)
            reader.Close()
            response.Close()
        Catch Ex As Exception
            EventLog.WriteEntry(Ex.Message & vbCrLf & vbCrLf & Ex.StackTrace, EventLogEntryType.Error)
            Dim FunctionReturn As New CloudAPIResponse
            FunctionReturn.response.code = 500
            FunctionReturn.response.message = Ex.Message
            PostPWCloudAPISettings = FunctionReturn
        End Try
    End Function
    Function LoginPWLocalUser(Optional ForceReLogin As Boolean = False) As String
        If PWLocalToken = String.Empty Or ForceReLogin = True Then
            Try
                Dim LoginRequest As New LoginRequest With {
                    .username = "customer",
                    .password = My.Settings.PWLocalUserPassword,
                    .email = My.Settings.PWLocalUserUsername,
                    .force_sm_off = True
                }
                Dim BodyPostData As String = JsonConvert.SerializeObject(LoginRequest).ToString
                Dim BodyByteStream As Byte() = Encoding.ASCII.GetBytes(BodyPostData)
                Dim request As HttpWebRequest = GetPWRequest("login/Basic")
                request.Method = "POST"
                request.ContentType = "application/json"
                request.ContentLength = BodyByteStream.Length
                request.CookieContainer = New CookieContainer()
                Dim BodyStream As Stream = request.GetRequestStream()
                BodyStream.Write(BodyByteStream, 0, BodyByteStream.Length)
                BodyStream.Close()
                Dim response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
                PWLocalCookies = New CookieCollection
                PWLocalCookies = response.Cookies
                Dim dataStream As Stream = response.GetResponseStream()
                Dim reader As StreamReader = New StreamReader(dataStream)
                Dim responseFromServer As String = reader.ReadToEnd()
                Dim LoginResult As LoginResult = JsonConvert.DeserializeObject(Of LoginResult)(responseFromServer)
                PWLocalToken = LoginResult.token
                reader.Close()
                response.Close()
            Catch ex As Exception
                EventLog.WriteEntry(ex.Message & vbCrLf & vbCrLf & ex.StackTrace, EventLogEntryType.Error)
            End Try
        End If
        Return PWLocalToken
    End Function
    Function LoginPWCloud(Optional ForceReLogin As Boolean = False) As String
        If PWCloudToken = String.Empty Then
            If My.Settings.PWCloudToken <> String.Empty Then
                PWCloudToken = My.Settings.PWCloudToken
                EventLog.WriteEntry(String.Format("Skipping Cloud Login: Found Settings Access Token: {0}", PWCloudToken), EventLogEntryType.Information, 904)
            End If
        End If
        If PWCloudToken = String.Empty Or ForceReLogin = True Then
            Try
                Dim AuthHelper As New TeslaAuthHelper(String.Format("PowerwallService/{0}", My.Application.Info.Version.ToString))
                PWCloudRefreshToken = AuthHelper.AuthenticateAsync(My.Settings.PWCloudEmail, My.Settings.PWCloudPassword, My.Settings.PWCloudMFARecoveryToken).Result.RefreshToken
                With AuthHelper.RefreshTokenAsync(PWCloudRefreshToken, TeslaAccountRegion.Unknown).Result
                    PWCloudToken = .AccessToken
                    PWCloudRefreshToken = .RefreshToken
                End With
                EventLog.WriteEntry(String.Format("Initial Access Token: {0}", PWCloudToken), EventLogEntryType.Information, 900)
                EventLog.WriteEntry(String.Format("Initial Refresh Token: {0}", PWCloudRefreshToken), EventLogEntryType.Information, 901)
            Catch ex As Exception
                EventLog.WriteEntry(ex.Message & vbCrLf & vbCrLf & ex.StackTrace, EventLogEntryType.Error)
            End Try
        End If
        Return PWCloudToken
    End Function
    Private Sub DoExitCharging(ByRef Intent As String)
        If SetPWMode("Exit Charge or Standby Mode", "Enter", DischargeMode, My.Settings.PWMinBackupPercentage, DischargeMode, Intent) = 202 Then
            PreCharging = False
            OnStandby = False
            AboveMinBackup = False
        End If
    End Sub
    Function GetLocalLoggedInJSONResult(Of JSONType)(URL As String) As JSONType
        If PWLocalCookies Is Nothing Or PWLocalToken = String.Empty Then
            PWLocalToken = LoginPWLocalUser()
        End If
        Dim response As HttpWebResponse
        Try
            Dim request As HttpWebRequest = CType(WebRequest.Create(URL), HttpWebRequest)
            request.CookieContainer = New CookieContainer()
            request.CookieContainer.Add(PWLocalCookies)
            Try
                response = CType(request.GetResponse(), HttpWebResponse)
            Catch ex As WebException
                If ex.Status <> 502 Then
                    EventLog.WriteEntry(ex.Message & vbCrLf & vbCrLf & ex.StackTrace, EventLogEntryType.Error)
                End If
                Return Nothing
            Catch ex As Exception
                Return Nothing
            End Try
            If response IsNot Nothing Then
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
    Private Sub RefreshTokens()
        If PWCloudRefreshToken <> String.Empty Then
            Dim AuthHelper As New TeslaAuthHelper(String.Format("PowerwallService/{0}", My.Application.Info.Version.ToString))
            With AuthHelper.RefreshTokenAsync(PWCloudRefreshToken, TeslaAccountRegion.Unknown).Result
                PWCloudToken = .AccessToken
                PWCloudRefreshToken = .RefreshToken
            End With
            EventLog.WriteEntry(String.Format("Refreshed Access Token: {0}", PWCloudToken), EventLogEntryType.Information, 902)
            EventLog.WriteEntry(String.Format("Refreshed Refresh Token: {0}", PWCloudRefreshToken), EventLogEntryType.Information, 903)
        Else
            EventLog.WriteEntry(String.Format("No Refresh Token available, using existing Access Token", PWCloudToken), EventLogEntryType.Information, 905)
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
                    If FiveMinData IsNot Nothing Then
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
            If FiveMinData IsNot Nothing Then
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
End Class
