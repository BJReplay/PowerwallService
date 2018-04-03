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
    Private ObservationTimer As New System.Timers.Timer
    Private OneMinuteTime As New System.Timers.Timer
    Private ReportingTimer As New System.Timers.Timer
    Private ForecastTimer As New System.Timers.Timer
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
    Shared OperationStart As DateTime
    Shared OperationEnd As DateTime
    Shared OperationStartHour As Integer
    Shared OperationEndHour As Integer
    Shared OperationHours As Double
    Shared IsCharging As Boolean = False
    Shared LastPeriodForecast As Forecast
    Shared CurrentPeriodForecast As Forecast
    Shared PVForecast As OutputForecast
    Shared Sunrise As DateTime
    Shared Sunset As DateTime
    Shared AsAtSunrise As Result
    Shared DBLock As New Object
#End Region
#Region "Timer Handlers"
    Protected Async Sub OnObservationTimer(Sender As Object, Args As System.Timers.ElapsedEventArgs)
        Await Task.Run(Sub()
                           GetObservationAndStore()
                       End Sub)
    End Sub
    Protected Async Sub OnOneMinuteTime(Sender As Object, Args As System.Timers.ElapsedEventArgs)
        Await Task.Run(Sub()
                           DoPerMinuteTasks()
                       End Sub)
    End Sub
    Protected Async Sub OnReportingTimer(Sender As Object, Args As System.Timers.ElapsedEventArgs)
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
    Protected Async Sub OnForecastTimer(Sender As Object, Args As System.Timers.ElapsedEventArgs)
        Await Task.Run(Sub()
                           CheckSOCLevel()
                       End Sub)
    End Sub
    Protected Async Sub DebugTask()
        Await Task.Run(Sub()
                           CheckSOCLevel()
                           If My.Settings.PVSendPowerwall Then DoBackFill(Now)
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

        OneMinuteTime.Interval = 60 * 1000 ' Every Minute
        OneMinuteTime.AutoReset = True
        AddHandler OneMinuteTime.Elapsed, AddressOf OnOneMinuteTime

        If My.Settings.PVReportingEnabled Then
            ReportingTimer.Interval = 5 * 60 * 1000 ' Every Five Minutes
            ReportingTimer.AutoReset = True
            AddHandler ReportingTimer.Elapsed, AddressOf OnReportingTimer
        End If

        ForecastTimer.Interval = 60 * 10 * 1000 ' Every 10 Minutes
        ForecastTimer.AutoReset = True
        AddHandler ForecastTimer.Elapsed, AddressOf OnForecastTimer

        Task.Run(Sub()
                     DoAsyncStartupProcesses()
                 End Sub)

        EventLog.WriteEntry("Powerwall Service Started", EventLogEntryType.Information, 101)
    End Sub
    Private Sub DoAsyncStartupProcesses()
        SetOperationHours()
        GetForecasts()
        Task.Run(Sub()
                     SleepUntilSecBoundary(6)
                     ObservationTimer.Start()
                     EventLog.WriteEntry("Observation Timer Started", EventLogEntryType.Information, 108)
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
        ObservationTimer.Start()
        OneMinuteTime.Start()
        Task.Run(Sub()
                     DebugTask()
                 End Sub
                 )
        EventLog.WriteEntry("Powerwall Service Running", EventLogEntryType.Information, 103)
    End Sub
    Protected Overrides Sub OnPause()
        EventLog.WriteEntry("Powerwall Service Pausing", EventLogEntryType.Information, 104)
        ObservationTimer.Stop()
        OneMinuteTime.Stop()
        If My.Settings.PVReportingEnabled Then ReportingTimer.Stop()
        ForecastTimer.Stop()
        EventLog.WriteEntry("Powerwall Service Paused", EventLogEntryType.Information, 105)
    End Sub
    Protected Overrides Sub OnStop()
        EventLog.WriteEntry("Powerwall Service Stoping", EventLogEntryType.Information, 106)
        ObservationTimer.Stop()
        ObservationTimer.Dispose()
        OneMinuteTime.Stop()
        OneMinuteTime.Dispose()
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
        EventLog.WriteEntry("Powerwall Service Running at " & Now.ToString, EventLogEntryType.Information, 200)
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
    Private Sub SetOperationHours()
        Dim InvokedTime As DateTime = Now
        Dim OperationDOW As DayOfWeek = InvokedTime.DayOfWeek
        If (OperationDOW = DayOfWeek.Saturday Or OperationDOW = DayOfWeek.Sunday) Then
            If My.Settings.TariffPeakOnWeekends Then
                OperationStartHour = My.Settings.TariffPeakEndWeekend
                OperationEndHour = My.Settings.TariffPeakStartWeekend
            Else
                OperationStartHour = 21
                OperationEndHour = 9
            End If
        Else
            OperationStartHour = My.Settings.TariffPeakEndWeekday
            OperationEndHour = My.Settings.TariffPeakStartWeekday
        End If
        If My.Settings.TariffIgnoresDST And System.TimeZone.CurrentTimeZone.IsDaylightSavingTime(InvokedTime) Then
            OperationStartHour += CByte(1)
            If OperationStartHour > 23 Then OperationStartHour -= CByte(24)
            OperationEndHour += CByte(1)
            If OperationEndHour > 23 Then OperationEndHour -= CByte(24)
        End If
        OperationHours = OperationEndHour - OperationStartHour
        If OperationHours <= 0 Then OperationHours += 24
        OperationStart = New DateTime(InvokedTime.Year, InvokedTime.Month, InvokedTime.Day, OperationStartHour, 0, 0)
        If OperationStartHour = 0 Then
            OperationStart = DateAdd(DateInterval.Day, 1, OperationStart)
        End If
        If InvokedTime.Hour >= 0 And InvokedTime.Hour < OperationEndHour + 1 Then
            OperationStart = DateAdd(DateInterval.Day, -1, OperationStart)
        End If
        OperationEnd = New DateTime(OperationStart.Year, OperationStart.Month, OperationStart.Day, OperationEndHour, 0, 0)
        If OperationStartHour <> 0 Then
            OperationEnd = DateAdd(DateInterval.Day, 1, OperationEnd)
        End If
        If Sunrise.Date <> InvokedTime.Date Then
            Dim SunriseSunsetData As Result = GetSunriseSunset(Of Result)(InvokedTime)
            Sunrise = DateAdd(DateInterval.Minute, -10, SunriseSunsetData.results.sunrise.ToLocalTime)
            Sunset = DateAdd(DateInterval.Minute, 10, SunriseSunsetData.results.sunset.ToLocalTime)
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
        If Minute Mod 10 = 1 Then
            If ForecastTimer.Enabled = False Then
                ForecastTimer.Start()
                EventLog.WriteEntry("Solar Forecast and Charge Monitoring Timer Started", EventLogEntryType.Information, 111)
                CheckSOCLevel()
            End If
        End If
        AggregateToMinute()
    End Sub
    Function GetUnsecuredJSONResult(Of JSONType)(URL As String) As JSONType
        Try
            Dim request As WebRequest = WebRequest.Create(URL)
            Dim response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
            Dim dataStream As Stream = response.GetResponseStream()
            Dim reader As StreamReader = New StreamReader(dataStream)
            Dim responseFromServer As String = reader.ReadToEnd()
            GetUnsecuredJSONResult = JsonConvert.DeserializeObject(Of JSONType)(responseFromServer)
            reader.Close()
            response.Close()
        Catch Ex As Exception
            EventLog.WriteEntry(Ex.Message & vbCrLf & vbCrLf & Ex.StackTrace, EventLogEntryType.Error)
        End Try
    End Function
    Function GetSunriseSunset(Of JSONType)(AsAt As Date) As JSONType
        Return GetUnsecuredJSONResult(Of JSONType)(String.Format("https://api.sunrise-sunset.org/json?lat={0}&lng={1}&formatted=0&date={2:yyyy-MM-dd}", My.Settings.PVSystemLattitude, My.Settings.PVSystemLongitude, AsAt))
    End Function
    Function CheckSunIsUp(AsAt As Date) As Boolean
        If AsAt.Date = Now.Date Then
            If AsAt > Sunrise And AsAt < Sunset Then Return True
        Else
            If AsAtSunrise Is Nothing Then
                AsAtSunrise = GetSunriseSunset(Of Result)(AsAt)
            End If
            If AsAtSunrise.results.sunrise.ToLocalTime.Date <> AsAt.Date Then
                AsAtSunrise = GetSunriseSunset(Of Result)(AsAt)
            End If
            If AsAt > DateAdd(DateInterval.Minute, -10, AsAtSunrise.results.sunrise.ToLocalTime) And AsAt < DateAdd(DateInterval.Minute, 10, AsAtSunrise.results.sunset.ToLocalTime) Then Return True
        End If
        Return False
    End Function
#End Region
#Region "Forecasts and Targets"
    Sub CheckSOCLevel()
        Dim InvokedTime As DateTime = Now
        SetOperationHours()
        If Not FirstReadingsAvailable Then
            GetObservationAndStore()
        End If
        GetForecasts()
        Dim DoExitSoc As Boolean = False
        Dim RawTargetSOC As Integer
        Dim ShortfallInsolation As Single = 0
        Dim PreChargeTargetSOC As Single = 0
        Dim RemainingOvernightRatio As Single
        Dim RemainingInsolationToday As Single
        Dim ForecastInsolationTomorrow As Single
        Dim RemainingOffPeak As Single
        Dim Intent As String = "Thinking"
        If InvokedTime > OperationStart And InvokedTime < OperationEnd Then
            RemainingOvernightRatio = CSng(DateDiff(DateInterval.Hour, InvokedTime, OperationEnd) / OperationHours)
            If RemainingOvernightRatio < 0 Then RemainingOvernightRatio = 0
            If RemainingOvernightRatio > 1 Then RemainingOvernightRatio = 1
            RemainingInsolationToday = NextDayForecastGeneration
            ForecastInsolationTomorrow = 0
            ShortfallInsolation = My.Settings.PWPeakConsumption - NextDayForecastGeneration
            Intent = "Thinking"
        ElseIf InvokedTime > Sunrise And InvokedTime < Sunset Then
            RemainingOvernightRatio = 1
            RemainingInsolationToday = CurrentDayForecast.PVEstimate
            ForecastInsolationTomorrow = NextDayForecastGeneration
            ShortfallInsolation = 0
            Intent = "Sun is Up"
        ElseIf InvokedTime > OperationEnd And InvokedTime < Sunrise Then
            RemainingOvernightRatio = 0
            RemainingInsolationToday = CurrentDayForecast.PVEstimate
            ForecastInsolationTomorrow = NextDayForecastGeneration
            ShortfallInsolation = My.Settings.PWPeakConsumption - NextDayForecastGeneration
            Intent = "Waiting for Sunrise"
        ElseIf InvokedTime > Sunset And InvokedTime < OperationStart Then
            RemainingOvernightRatio = 1
            RemainingInsolationToday = 0
            ForecastInsolationTomorrow = NextDayForecastGeneration
            ShortfallInsolation = My.Settings.PWPeakConsumption - NextDayForecastGeneration
            Intent = "Waiting for Off Peak"
        End If
        RemainingOffPeak = My.Settings.PWOvernightLoad * RemainingOvernightRatio
        RawTargetSOC = My.Settings.PWMorningBuffer + CInt(RemainingOffPeak)
        If ShortfallInsolation < 0 Then ShortfallInsolation = 0
        PreChargeTargetSOC = RawTargetSOC + (ShortfallInsolation / My.Settings.PWCapacity * 100)
        If PreChargeTargetSOC > 100 Then PreChargeTargetSOC = 100
        If InvokedTime > Sunset And InvokedTime < OperationStart Then
            If PreChargeTargetSOC > RawTargetSOC Or RawTargetSOC > SOC.percentage Then Intent = "Planning to Charge" Else Intent = "No Charging Required"
        End If
        Try
            If InvokedTime > DateAdd(DateInterval.Minute, -20, OperationEnd) And PreCharging Then
                DoExitSoc = True
                EventLog.WriteEntry(String.Format("Reached end of pre-charging period: SOC={0}, Required={1}, Shortfall={2}, NewTarget={3}", SOC.percentage.ToString, RawTargetSOC.ToString, ShortfallInsolation.ToString, PreChargeTargetSOC.ToString), EventLogEntryType.Information, 504)
                Intent = "Stop Charging"
            End If
            If (InvokedTime >= OperationStart And InvokedTime < OperationEnd) Or DoExitSoc Then
                EventLog.WriteEntry(String.Format("In Operation Period: SOC={0}, Required={1}, Shortfall={2}, NewTarget={3}", SOC.percentage.ToString, RawTargetSOC.ToString, ShortfallInsolation.ToString, PreChargeTargetSOC.ToString), EventLogEntryType.Information, 500)
                If SOC.percentage < PreChargeTargetSOC And Not DoExitSoc Then
                    EventLog.WriteEntry(String.Format("Current SOC below required setting: SOC={0}, Required={1}, Shortfall={2}, NewTarget={3}", SOC.percentage.ToString, RawTargetSOC.ToString, ShortfallInsolation.ToString, PreChargeTargetSOC.ToString), EventLogEntryType.Information, 501)
                    Intent = "Start Charging"
                    If Not PreCharging Then
                        Dim ChargeSettings As New Operation With {.backup_reserve_percent = CInt(PreChargeTargetSOC), .mode = IIf(My.Settings.PWChargeModeBackup, "backup", "self_consumption").ToString}
                        Dim NewChargeSettings As Operation = PostPWSecureAPISettings(Of Operation)("operation", ChargeSettings, ForceReLogin:=True)
                        Dim APIResult As Integer = GetPWSecureConfigCompleted("config/completed")
                        If APIResult = 202 Then
                            EventLog.WriteEntry(String.Format("Entered Charge Mode: SOC={0}, Required={1}, Shortfall={2}, NewTarget={3}, Mode={4}, BackupPercentage={5}, APIResult = {6}", SOC.percentage.ToString, RawTargetSOC.ToString, ShortfallInsolation.ToString, PreChargeTargetSOC.ToString, NewChargeSettings.mode, NewChargeSettings.backup_reserve_percent, APIResult.ToString), EventLogEntryType.Information, 506)
                            PreCharging = True
                            If PreChargeTargetSOC > (SOC.percentage + 5) Then
                                Intent = "Charging"
                            Else
                                Intent = "Standby"
                            End If
                        Else
                            EventLog.WriteEntry(String.Format("Failed to enter Charge Mode: SOC={0}, Required={1}, Shortfall={2}, NewTarget={3}, Mode={4}, BackupPercentage={5}, APIResult = {6}", SOC.percentage.ToString, RawTargetSOC.ToString, ShortfallInsolation.ToString, PreChargeTargetSOC.ToString, NewChargeSettings.mode, NewChargeSettings.backup_reserve_percent, APIResult.ToString), EventLogEntryType.Information, 507)
                            Intent = "Trying To Charge"
                        End If
                    End If
                End If
                If SOC.percentage > (PreChargeTargetSOC + 5) And PreCharging Then
                    DoExitSoc = True
                    EventLog.WriteEntry(String.Format("Current SOC above required setting: SOC={0}, Required={1}, Shortfall={2}, NewTarget={3}", SOC.percentage.ToString, RawTargetSOC.ToString, ShortfallInsolation.ToString, PreChargeTargetSOC.ToString), EventLogEntryType.Information, 502)
                    Intent = "Exit Charging"
                End If
                If DoExitSoc Then
                    EventLog.WriteEntry(String.Format("Exiting Charge Mode: SOC={0}, Required={1}, Shortfall={2}, NewTarget={3}", SOC.percentage.ToString, RawTargetSOC.ToString, ShortfallInsolation.ToString, PreChargeTargetSOC.ToString), EventLogEntryType.Information, 505)
                    Dim ChargeSettings As New Operation With {.backup_reserve_percent = My.Settings.PWMinBackupPercentage, .mode = "self_consumption"}
                    Dim NewChargeSettings As Operation = PostPWSecureAPISettings(Of Operation)("operation", ChargeSettings, ForceReLogin:=True)
                    Dim APIResult As Integer = GetPWSecureConfigCompleted("config/completed")
                    If APIResult = 202 Then
                        PreCharging = False
                        EventLog.WriteEntry(String.Format("Exited Charge Mode: SOC={0}, Required={1}, Shortfall={2}, NewTarget={3}, Mode={4}, BackupPercentage={5}, APIResult = {6}", SOC.percentage.ToString, RawTargetSOC.ToString, ShortfallInsolation.ToString, PreChargeTargetSOC.ToString, NewChargeSettings.mode, NewChargeSettings.backup_reserve_percent, APIResult.ToString), EventLogEntryType.Information, 508)
                        Intent = "Self Consumption"
                    Else
                        EventLog.WriteEntry(String.Format("Failed to exit Charge Mode: SOC={0}, Required={1}, Shortfall={2}, NewTarget={3}, Mode={4}, BackupPercentage={5}, APIResult = {6}", SOC.percentage.ToString, RawTargetSOC.ToString, ShortfallInsolation.ToString, PreChargeTargetSOC.ToString, NewChargeSettings.mode, NewChargeSettings.backup_reserve_percent, APIResult.ToString), EventLogEntryType.Information, 509)
                        Intent = "Trying to Exit Charging"
                    End If
                End If
            Else
                EventLog.WriteEntry(String.Format("Outside Operation Period: SOC={0}", SOC.percentage.ToString), EventLogEntryType.Information, 503)
            End If
        Catch Ex As Exception
            EventLog.WriteEntry(Ex.Message & vbCrLf & vbCrLf & Ex.StackTrace, EventLogEntryType.Error, 510)
        End Try
        If My.Settings.PBIChargeIntentEndpoint <> String.Empty Then
            Try
                Dim PBIRows As New PBIChargeLogging With {.Rows = New List(Of ChargePlan)}
                PBIRows.Rows.Add(New ChargePlan With {.AsAt = InvokedTime, .CurrentSOC = SOC.percentage, .RemainingInsolation = RemainingInsolationToday, .ForecastGeneration = ForecastInsolationTomorrow, .MorningBuffer = My.Settings.PWMorningBuffer, .OperatingIntent = Intent, .RequiredSOC = PreChargeTargetSOC, .RemainingOffPeak = RemainingOffPeak * My.Settings.PWCapacity / 100, .Shortfall = ShortfallInsolation})
                Dim PowerBIPostResult As Integer = PostPowerBIStreamingData(My.Settings.PBIChargeIntentEndpoint, PBIRows)
            Catch ex As Exception

            End Try
        End If
    End Sub
    Sub GetForecasts()
        Try
            Dim NewForecastsRetrieved As Boolean = False
            Dim InvokedTime As DateTime = Now
            If DateAdd(DateInterval.Minute, 10, ForecastsRetrieved) < InvokedTime Then
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
                    ForecastLogEntry += vbCrLf & String.Format("Date: {0} Total: {1} Morning: {2}", .ForecastDate.ToString, .PVEstimate.ToString, .MorningForecast.ToString)
                End With
                With NextDayForecast
                    ForecastLogEntry += vbCrLf & String.Format("Date: {0} Total: {1} Morning: {2}", .ForecastDate.ToString, .PVEstimate.ToString, .MorningForecast.ToString)
                End With
                With SecondDayForecast
                    ForecastLogEntry += vbCrLf & String.Format("Date: {0} Total: {1} Morning: {2}", .ForecastDate.ToString, .PVEstimate.ToString, .MorningForecast.ToString)
                End With
                EventLog.WriteEntry(ForecastLogEntry, EventLogEntryType.Information, 1000)
            End If
            If InvokedTime.Hour >= 0 And InvokedTime.Hour < OperationEndHour Then
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
        Try
            Dim request As WebRequest = WebRequest.Create(String.Format(My.Settings.SolcastAddress, My.Settings.PVSystemLongitude, My.Settings.PVSystemLattitude, My.Settings.PVSystemCapacity, My.Settings.PVSystemTilt, My.Settings.PVSystemAzimuth, My.Settings.PVSystemInstallDate, My.Settings.SolcastAPIKey))
            Dim response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
            Dim dataStream As Stream = response.GetResponseStream()
            Dim reader As StreamReader = New StreamReader(dataStream)
            Dim responseFromServer As String = reader.ReadToEnd()
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
            EventLog.WriteEntry(Ex.Message & vbCrLf & vbCrLf & Ex.StackTrace, EventLogEntryType.Error)
        End Try
    End Function
#End Region
#Region "Six Second Logger"
    Sub GetObservationAndStore()
        Try
            MeterReading = GetPWAPIResult(Of MeterAggregates)("meters/aggregates")
            Dim ObservationTime As Date = MeterReading.site.last_communication_time
            SOC = GetPWAPIResult(Of SOC)("system_status/soe")
            FirstReadingsAvailable = True
            If My.Settings.LogData Then
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
                If My.Settings.PBILiveLoggingEndpoint <> String.Empty Then
                    Try
                        Dim PBIRows As New PBILiveLogging With {.Rows = New List(Of SixSecondOb)}
                        PBIRows.Rows.Add(New SixSecondOb With {.AsAt = ObservationTime, .Battery = MeterReading.battery.instant_power, .Grid = MeterReading.site.instant_power, .Load = MeterReading.load.instant_power, .SOC = SOC.percentage, .Solar = CSng(IIf(MeterReading.solar.instant_power < 0, 0, MeterReading.solar.instant_power)), .Voltage = MeterReading.battery.instant_average_voltage})
                        Dim PowerBIPostResult As Integer = PostPowerBIStreamingData(My.Settings.PBILiveLoggingEndpoint, PBIRows)
                    Catch ex As Exception

                    End Try
                End If
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
    Function GetPWSecureAPIResult(Of JSONType)(API As String, Optional ForceReLogin As Boolean = False) As JSONType
        Try
            PWToken = LoginPWLocal(ForceReLogin:=ForceReLogin)
            Dim request As WebRequest = WebRequest.Create(My.Settings.PWGatewayAddress & "/api/" & API)
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
    Function GetPWSecureConfigCompleted(API As String, Optional ForceReLogin As Boolean = False) As Integer
        Try
            PWToken = LoginPWLocal(ForceReLogin:=ForceReLogin)
            Dim request As WebRequest = WebRequest.Create(My.Settings.PWGatewayAddress & "/api/" & API)
            request.Headers.Add("Authorization", "Bearer " & PWToken)
            Dim response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
            Dim dataStream As Stream = response.GetResponseStream()
            Dim reader As StreamReader = New StreamReader(dataStream)
            Dim responseFromServer As String = reader.ReadToEnd()
            GetPWSecureConfigCompleted = response.StatusCode
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
            Dim request As WebRequest = WebRequest.Create(My.Settings.PWGatewayAddress & "/api/" & API)
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
                    .force_sm_off = False
                }
                Dim BodyPostData As String = JsonConvert.SerializeObject(LoginRequest).ToString
                Dim BodyByteStream As Byte() = Encoding.ASCII.GetBytes(BodyPostData)
                Dim request As WebRequest = WebRequest.Create(My.Settings.PWGatewayAddress & "/api/login/Basic")
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
End Class
