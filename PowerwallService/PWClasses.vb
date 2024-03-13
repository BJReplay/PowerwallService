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
    Public Class ListProducts
        Public Property response() As List(Of Products)
        Public Property count As Integer
    End Class
    Public Class Products
        Public Property energy_site_id As Long
        Public Property resource_type As String
        Public Property site_name As String
        Public Property id As String
        Public Property gateway_id As String
        Public Property asset_site_id As String
        Public Property energy_left As Single
        Public Property total_pack_energy As Integer
        Public Property percentage_charged As Single
        Public Property battery_type As String
        Public Property backup_capable As Boolean
        Public Property battery_power As Decimal
        Public Property sync_grid_alert_enabled As Boolean
        Public Property breaker_alert_enabled As Boolean
        Public Property components As ProductComponents
    End Class
    Public Class ProductComponents
        Public Property battery As Boolean
        Public Property battery_type As String
        Public Property solar As Boolean
        Public Property solar_type As String
        Public Property grid As Boolean
        Public Property load_meter As Boolean
        Public Property market_type As String
    End Class
    Public Class CloudLoginRequest
        Public Property grant_type As String = "password"
        Public Property client_id As String = "81527cff06843c8634fdc09e8ac0abefb46ac849f38fe1e431c2ef2106796384"
        Public Property client_secret As String = "c7257eb71a564034f9419ee651c7d0e5f7aa6bfbd18bafb5c5c033b093bb2fa3"
        Public Property email As String
        Public Property password As String
    End Class
    Public Class CloudLoginResponse
        Public Property access_token As String
        Public Property token_type As String
        Public Property expires_in As Integer
        Public Property refresh_token As String
        Public Property created_at As Integer
    End Class
    Public Class PowerwallStatus
        Public Property response As PowerwallStatusResponse
    End Class
    Public Class PowerwallStatusResponse
        Public Property site_name As String
        Public Property id As String
        Public Property energy_left As Single
        Public Property total_pack_energy As Integer
        Public Property percentage_charged As Single
        Public Property battery_power As Integer
    End Class
    Public Class SiteStatus
        Public Property response As SiteStatusResponse
    End Class
    Public Class SiteStatusResponse
        Public Property resource_type As String
        Public Property site_name As String
        Public Property gateway_id As String
        Public Property energy_left As Single
        Public Property total_pack_energy As Integer
        Public Property percentage_charged As Single
        Public Property battery_type As String
        Public Property backup_capable As Boolean
        Public Property battery_power As Integer
        Public Property sync_grid_alert_enabled As Boolean
        Public Property breaker_alert_enabled As Boolean
    End Class
    Public Class CloudBackup
        Public Property backup_reserve_percent As Decimal
    End Class
    Public Class CloudOperation
        Public Property default_real_mode As String
    End Class
    Public Class CloudAPIResponse
        Public Property response As APIResponse
    End Class
    Public Class APIResponse
        Public Property code As Integer
        Public Property message As String
    End Class
    Public Class CloudSiteInfo
        Public Property response As SiteInfoResponse
    End Class
    Public Class SiteInfoResponse
        Public Property id As String
        Public Property site_name As String
        Public Property backup_reserve_percent As Decimal
        Public Property default_real_mode As String
        Public Property installation_date As Date
        Public Property user_settings As User_Settings
        Public Property components As Components
        Public Property version As String
        Public Property battery_count As Integer
        Public Property tou_settings As Tou_Settings
        Public Property nameplate_power As Decimal
        Public Property nameplate_energy As Decimal
        Public Property installation_time_zone As String
    End Class
    Public Class User_Settings
        Public Property storm_mode_enabled As Boolean
        Public Property sync_grid_alert_enabled As Boolean
        Public Property breaker_alert_enabled As Boolean
    End Class
    Public Class Components
        Public Property solar As Boolean
        Public Property solar_type As String
        Public Property battery As Boolean
        Public Property grid As Boolean
        Public Property backup As Boolean
        Public Property gateway As String
        Public Property load_meter As Boolean
        Public Property tou_capable As Boolean
        Public Property storm_mode_capable As Boolean
        Public Property flex_energy_request_capable As Boolean
        Public Property car_charging_data_supported As Boolean
        Public Property off_grid_vehicle_charging_reserve_supported As Boolean
        Public Property vehicle_charging_performance_view_enabled As Boolean
        Public Property vehicle_charging_solar_offset_view_enabled As Boolean
        Public Property battery_solar_offset_view_enabled As Boolean
        Public Property battery_type As String
        Public Property configurable As Boolean
        Public Property grid_services_enabled As Boolean
    End Class
    Public Class Tou_Settings
        Public Property optimization_strategy As String
        Public Property schedule() As List(Of Schedule)
    End Class
    Public Class Schedule
        Public Property target As String
        Public Property week_days() As List(Of Integer)
        Public Property start_seconds As Integer
        Public Property end_seconds As Integer
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
        Public Sub New(field() As String, len As Integer)
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
            If len > 10 Then
                v7 = field(11)
                v8 = field(12)
                v9 = field(13)
                v10 = field(14)
                v11 = field(15)
                v12 = field(16)
            End If
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
        Public Property Frequency As Single
        Public Property FrequencyMin As Single = 45.0
        Public Property FrequencyMax As Single = 55.0
    End Class
    Public Class PBIChargeLogging
        Public Property Rows() As List(Of ChargePlan)
    End Class
    Public Class ChargePlan
        Public Property CurrentSOC As Single
        Public Property OvernightConsumption As Single
        Public Property SunriseToPeak As Single
        Public Property ForecastGeneration As Single
        Public Property Shortfall As Single
        Public Property RequiredSOC As Single
        Public Property AsAt As DateTime
        Public Property OperatingIntent As String
        Public Property RemainingInsolation As Single
        Public Property PeakConsumption As Single
    End Class
End Class
Public Class HashiCorp
    Public Class HashiToken
        Public Property access_token As String
        Public Property expires_in As Integer
        Public Property token_type As String
    End Class

    Public Class SecretsList
        Public Property secrets() As List(Of Secret)
    End Class

    Public Class Secret
        Public Property name As String
        Public Property version As Version
        Public Property created_at As Date
        Public Property latest_version As String
        Public Property created_by As Created_By1
        Public Property sync_status As Sync_Status
    End Class

    Public Class Version
        Public Property version As String
        Public Property type As String
        Public Property created_at As Date
        Public Property value As String
        Public Property created_by As Created_By
    End Class

    Public Class Created_By
        Public Property name As String
        Public Property type As String
        Public Property email As String
    End Class

    Public Class Created_By1
        Public Property name As String
        Public Property type As String
        Public Property email As String
    End Class

    Public Class Sync_Status
    End Class

End Class
Public Class HomeAssistant
    Public Class EntityState
        Public Class Entity
            Public Property entity_id As String
            Public Property state As String
            Public Property attributes As Object
            Public Property last_changed As Date
            Public Property last_updated As Date
            Public Property context As Context
        End Class
        Public Class Context
            Public Property id As String
            Public Property parent_id As Object
            Public Property user_id As String
        End Class
    End Class
End Class
