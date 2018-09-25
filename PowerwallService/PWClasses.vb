Imports System.Reflection
Public Class PWJson
    Public Class MeterAggregates
        Public Property site As Site
        Public Property battery As Battery
        Public Property load As Load
        Public Property solar As Solar
        Public Property busway As Busway
        Public Property frequency As Frequency
    End Class
    Public Class Status
        Public Property start_time As String
        Public Property up_time_seconds As String
        Public Property is_new As Boolean
        Public Property version As String
        Public Property git_hash As String
    End Class
    Public Class Site
        Public Property last_communication_time As Date
        Public Property instant_power As Decimal
        Public Property instant_reactive_power As Decimal
        Public Property instant_apparent_power As Decimal
        Public Property frequency As Decimal
        Public Property energy_exported As Decimal
        Public Property energy_imported As Decimal
        Public Property instant_average_voltage As Decimal
        Public Property instant_total_current As Decimal
        Public Property i_a_current As Integer
        Public Property i_b_current As Integer
        Public Property i_c_current As Integer
    End Class
    Public Class Battery
        Public Property last_communication_time As Date
        Public Property instant_power As Decimal
        Public Property instant_reactive_power As Decimal
        Public Property instant_apparent_power As Decimal
        Public Property frequency As Decimal
        Public Property energy_exported As Decimal
        Public Property energy_imported As Decimal
        Public Property instant_average_voltage As Decimal
        Public Property instant_total_current As Decimal
        Public Property i_a_current As Integer
        Public Property i_b_current As Integer
        Public Property i_c_current As Integer
    End Class
    Public Class Load
        Public Property last_communication_time As Date
        Public Property instant_power As Decimal
        Public Property instant_reactive_power As Decimal
        Public Property instant_apparent_power As Decimal
        Public Property frequency As Decimal
        Public Property energy_exported As Decimal
        Public Property energy_imported As Decimal
        Public Property instant_average_voltage As Decimal
        Public Property instant_total_current As Decimal
        Public Property i_a_current As Integer
        Public Property i_b_current As Integer
        Public Property i_c_current As Integer
    End Class
    Public Class Solar
        Public Property last_communication_time As Date
        Public Property instant_power As Decimal
        Public Property instant_reactive_power As Decimal
        Public Property instant_apparent_power As Decimal
        Public Property frequency As Decimal
        Public Property energy_exported As Decimal
        Public Property energy_imported As Decimal
        Public Property instant_average_voltage As Decimal
        Public Property instant_total_current As Decimal
        Public Property i_a_current As Integer
        Public Property i_b_current As Integer
        Public Property i_c_current As Integer
    End Class
    Public Class Busway
        Public Property last_communication_time As Date
        Public Property instant_power As Integer
        Public Property instant_reactive_power As Integer
        Public Property instant_apparent_power As Integer
        Public Property frequency As Integer
        Public Property energy_exported As Decimal
        Public Property energy_imported As Decimal
        Public Property instant_average_voltage As Integer
        Public Property instant_total_current As Integer
        Public Property i_a_current As Integer
        Public Property i_b_current As Integer
        Public Property i_c_current As Integer
    End Class
    Public Class Frequency
        Public Property last_communication_time As Date
        Public Property instant_power As Integer
        Public Property instant_reactive_power As Integer
        Public Property instant_apparent_power As Integer
        Public Property frequency As Integer
        Public Property energy_exported As Decimal
        Public Property energy_imported As Decimal
        Public Property instant_average_voltage As Integer
        Public Property instant_total_current As Integer
        Public Property i_a_current As Integer
        Public Property i_b_current As Integer
        Public Property i_c_current As Integer
    End Class
    Public Class Powerwalls
        Public Property powerwalls() As List(Of Powerwall)
        Public Property hasSync As Boolean
    End Class
    Public Class Powerwall
        Public Property PackagePartNumber As String
        Public Property PackageSerialNumber As String
    End Class
    Public Class SiteMaster
        Public Property running As Boolean
        Public Property uptime As String
        Public Property connected_to_tesla As Boolean
    End Class
    Public Class SiteInfo
        Public Property site_name As String
        Public Property timezone As String
        Public Property measured_frequency As Decimal
        Public Property min_site_meter_power_kW As Integer
        Public Property max_site_meter_power_kW As Integer
        Public Property nominal_system_energy_kWh As Decimal
        Public Property grid_code As String
        Public Property grid_voltage_setting As Integer
        Public Property grid_freq_setting As Integer
        Public Property grid_phase_setting As String
        Public Property country As String
        Public Property state As String
        Public Property region As String
    End Class
    Public Class SOC
        Public Property percentage As Decimal
    End Class
    Public Class CustomerRegistration
        Public Property privacy_notice As Boolean
        Public Property limited_warranty As Boolean
        Public Property grid_services As Boolean
        Public Property marketing As Boolean
        Public Property registered As Boolean
        Public Property emailed_registration As Boolean
        Public Property skipped_registration As Boolean
        Public Property timed_out_registration As Boolean
    End Class
    Public Class Operation
        Public Property backup_reserve_percent As Decimal
        Public Property real_mode As String
    End Class
    Public Class LoginRequest
        Public Property username As String
        Public Property password As String
        Public Property email As String
        Public Property force_sm_off As Boolean
    End Class
    Public Class LoginResult
        Public Property email As Object
        Public Property firstname As String
        Public Property lastname As String
        Public Property roles() As List(Of String)
        Public Property token As String
        Public Property provider As String
    End Class
End Class
Public Class SolCast
    Public Class OutputForecast
        Public Property forecasts() As List(Of Forecast)
    End Class
    Public Class Forecast
        Public Property period_end As Date
        Public Property period As String
        Public Property pv_estimate As Single
    End Class
    Public Class DayForecast
        Public Property ForecastDate As Date
        Public Property PVEstimate As Single
        Public Property MorningForecast As Single
    End Class
End Class
Public Class PVOutput
    Public Class ParamData
        Public Property ParamName As String
        Public Property Data As String
    End Class
    Public Class StatusRecord
        Public Sub New(field() As String)
            d = field(0)
            t = field(1)
            v1 = field(2)
            Efficieny = field(3)
            v2 = field(4)
            AvPower = field(5)
            Output = field(6)
            v3 = field(7)
            v4 = field(8)
            v5 = field(9)
            v6 = field(10)
            v7 = field(11)
            v8 = field(12)
            v9 = field(13)
            v10 = field(14)
            v11 = field(15)
            v12 = field(16)
        End Sub
        Default ReadOnly Property Item(PropertyName As String) As Object
            Get
                Return [GetType].InvokeMember(PropertyName, BindingFlags.GetProperty, Nothing, Me, Nothing)
            End Get
        End Property
        Public Property d As String
        Public Property t As String
        Public Property v1 As String
        Public Property Efficieny As String
        Public Property v2 As String
        Public Property AvPower As String
        Public Property Output As String
        Public Property v3 As String
        Public Property v4 As String
        Public Property v5 As String
        Public Property v6 As String
        Public Property v7 As String
        Public Property v8 As String
        Public Property v9 As String
        Public Property v10 As String
        Public Property v11 As String
        Public Property v12 As String
    End Class
End Class
Public Class SunriseSunset
    Public Class Result
        Public Property results As Results
        Public Property status As String
    End Class
    Public Class Results
        Public Property sunrise As Date
        Public Property sunset As Date
        Public Property solar_noon As Date
        Public Property day_length As Integer
        Public Property civil_twilight_begin As Date
        Public Property civil_twilight_end As Date
        Public Property nautical_twilight_begin As Date
        Public Property nautical_twilight_end As Date
        Public Property astronomical_twilight_begin As Date
        Public Property astronomical_twilight_end As Date
    End Class
End Class
Public Class PowerBIStreaming
    Public Class PBILiveLogging
        Public Property Rows() As List(Of SixSecondOb)
    End Class
    Public Class SixSecondOb
        Public Property Load As Single
        Public Property Solar As Single
        Public Property Grid As Single
        Public Property Battery As Single
        Public Property Voltage As Single
        Public Property SOC As Single
        Public Property AsAt As DateTime
        Public Property kWMax As Single = My.Settings.PBIkWChartMinMax
        Public Property kWMin As Single = -My.Settings.PBIkWChartMinMax
        Public Property SOCMax As Single = 100
        Public Property SOCMin As Single = 0
        Public Property VoltageMax As Single = CSng(My.Settings.PBIVoltageNominal * 1.1)
        Public Property VoltageMin As Single = CSng(My.Settings.PBIVoltageNominal * 0.9)
        Public Property VoltageTarget As Single = My.Settings.PBIVoltageNominal
        Public Property SOCTarget As Single = 90
    End Class
    Public Class PBIChargeLogging
        Public Property Rows() As List(Of ChargePlan)
    End Class
    Public Class ChargePlan
        Public Property CurrentSOC As Single
        Public Property MorningBuffer As Single
        Public Property RemainingOffPeak As Single
        Public Property ForecastGeneration As Single
        Public Property Shortfall As Single
        Public Property RequiredSOC As Single
        Public Property AsAt As DateTime
        Public Property OperatingIntent As String
        Public Property RemainingInsolation As Single
    End Class
End Class

