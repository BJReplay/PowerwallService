﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:4.0.30319.42000
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On


Namespace My
    
    <Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "17.9.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
    Partial Friend NotInheritable Class MySettings
        Inherits Global.System.Configuration.ApplicationSettingsBase
        
        Private Shared defaultInstance As MySettings = CType(Global.System.Configuration.ApplicationSettingsBase.Synchronized(New MySettings()),MySettings)
        
#Region "My.Settings Auto-Save Functionality"
#If _MyType = "WindowsForms" Then
    Private Shared addedHandler As Boolean

    Private Shared addedHandlerLockObject As New Object

    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(), Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)> _
    Private Shared Sub AutoSaveSettings(sender As Global.System.Object, e As Global.System.EventArgs)
        If My.Application.SaveMySettingsOnExit Then
            My.Settings.Save()
        End If
    End Sub
#End If
#End Region
        
        Public Shared ReadOnly Property [Default]() As MySettings
            Get
                
#If _MyType = "WindowsForms" Then
               If Not addedHandler Then
                    SyncLock addedHandlerLockObject
                        If Not addedHandler Then
                            AddHandler My.Application.Shutdown, AddressOf AutoSaveSettings
                            addedHandler = True
                        End If
                    End SyncLock
                End If
#End If
                Return defaultInstance
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public ReadOnly Property LogCompactOnly() As Boolean
            Get
                Return CType(Me("LogCompactOnly"),Boolean)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public ReadOnly Property PVSendLoad() As Boolean
            Get
                Return CType(Me("PVSendLoad"),Boolean)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public ReadOnly Property PVSendPV() As Boolean
            Get
                Return CType(Me("PVSendPV"),Boolean)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("battery_instant_power")>  _
        Public ReadOnly Property PVv7() As String
            Get
                Return CType(Me("PVv7"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("load_instant_power")>  _
        Public ReadOnly Property PVv8() As String
            Get
                Return CType(Me("PVv8"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("soc_state_of_charge")>  _
        Public ReadOnly Property PVv9() As String
            Get
                Return CType(Me("PVv9"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("site_instant_power")>  _
        Public ReadOnly Property PVv10() As String
            Get
                Return CType(Me("PVv10"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("battery_instant_average_voltage")>  _
        Public ReadOnly Property PVv11() As String
            Get
                Return CType(Me("PVv11"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("solar_instant_power")>  _
        Public ReadOnly Property PVv12() As String
            Get
                Return CType(Me("PVv12"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("https://api.solcast.com.au/pv_power/forecasts?longitude={0}&latitude={1}&capacity"& _ 
            "={2}&tilt={3}&azimuth={4}&install_date={5:yyyyMMdd}&api_key={6}&format=json")>  _
        Public ReadOnly Property SolcastAddress() As String
            Get
                Return CType(Me("SolcastAddress"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("35")>  _
        Public ReadOnly Property PWOvernightLoad() As Integer
            Get
                Return CType(Me("PWOvernightLoad"),Integer)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("10")>  _
        Public ReadOnly Property PWMorningBuffer() As Integer
            Get
                Return CType(Me("PWMorningBuffer"),Integer)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public ReadOnly Property PVReportingEnabled() As Boolean
            Get
                Return CType(Me("PVReportingEnabled"),Boolean)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("5000")>  _
        Public ReadOnly Property PWPeakConsumption() As Integer
            Get
                Return CType(Me("PWPeakConsumption"),Integer)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("13500")>  _
        Public ReadOnly Property PWCapacity() As Integer
            Get
                Return CType(Me("PWCapacity"),Integer)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("13000")>  _
        Public ReadOnly Property PWSelfConsumption() As Integer
            Get
                Return CType(Me("PWSelfConsumption"),Integer)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public ReadOnly Property LogData() As Boolean
            Get
                Return CType(Me("LogData"),Boolean)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public ReadOnly Property TariffIgnoresDST() As Boolean
            Get
                Return CType(Me("TariffIgnoresDST"),Boolean)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("15")>  _
        Public ReadOnly Property TariffPeakStart() As Integer
            Get
                Return CType(Me("TariffPeakStart"),Integer)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("22")>  _
        Public ReadOnly Property TariffPeakEnd() As Integer
            Get
                Return CType(Me("TariffPeakEnd"),Integer)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("36")>  _
        Public ReadOnly Property TariffPeakRate() As Single
            Get
                Return CType(Me("TariffPeakRate"),Single)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("23")>  _
        Public ReadOnly Property TariffOffPeakRate() As Single
            Get
                Return CType(Me("TariffOffPeakRate"),Single)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("8.2")>  _
        Public ReadOnly Property TariffFeedInRate() As Single
            Get
                Return CType(Me("TariffFeedInRate"),Single)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("0.85")>  _
        Public ReadOnly Property PWRoundTripEfficiency() As Decimal
            Get
                Return CType(Me("PWRoundTripEfficiency"),Decimal)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public ReadOnly Property LogAzureOnly() As Boolean
            Get
                Return CType(Me("LogAzureOnly"),Boolean)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public ReadOnly Property PVSendForecast() As Boolean
            Get
                Return CType(Me("PVSendForecast"),Boolean)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("v12")>  _
        Public ReadOnly Property PVSendForecastAs() As String
            Get
                Return CType(Me("PVSendForecastAs"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public ReadOnly Property PVSendPowerwall() As Boolean
            Get
                Return CType(Me("PVSendPowerwall"),Boolean)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public ReadOnly Property PVSendVoltage() As Boolean
            Get
                Return CType(Me("PVSendVoltage"),Boolean)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public ReadOnly Property PWChargeModeBackup() As Boolean
            Get
                Return CType(Me("PWChargeModeBackup"),Boolean)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("5")>  _
        Public ReadOnly Property PWMinBackupPercentage() As Integer
            Get
                Return CType(Me("PWMinBackupPercentage"),Integer)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("TODO")>  _
        Public ReadOnly Property PWGatewayAddress() As String
            Get
                Return CType(Me("PWGatewayAddress"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public ReadOnly Property PVOutputSID() As Integer
            Get
                Return CType(Me("PVOutputSID"),Integer)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("TODO")>  _
        Public ReadOnly Property PVOutputAPIKey() As String
            Get
                Return CType(Me("PVOutputAPIKey"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("installer")>  _
        Public ReadOnly Property PWLocalInstallerUsername() As String
            Get
                Return CType(Me("PWLocalInstallerUsername"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("STTODO")>  _
        Public ReadOnly Property PWLocalInstallerPassword() As String
            Get
                Return CType(Me("PWLocalInstallerPassword"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("TODO")>  _
        Public ReadOnly Property SolcastAPIKey() As String
            Get
                Return CType(Me("SolcastAPIKey"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public ReadOnly Property PVSystemLongitude() As Single
            Get
                Return CType(Me("PVSystemLongitude"),Single)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public ReadOnly Property PVSystemLattitude() As Single
            Get
                Return CType(Me("PVSystemLattitude"),Single)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public ReadOnly Property PVSystemCapacity() As Integer
            Get
                Return CType(Me("PVSystemCapacity"),Integer)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public ReadOnly Property PVSystemTilt() As Integer
            Get
                Return CType(Me("PVSystemTilt"),Integer)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public ReadOnly Property PVSystemAzimuth() As Integer
            Get
                Return CType(Me("PVSystemAzimuth"),Integer)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public ReadOnly Property PVSystemInstallDate() As Date
            Get
                Return CType(Me("PVSystemInstallDate"),Date)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.SpecialSettingAttribute(Global.System.Configuration.SpecialSetting.WebServiceUrl)>  _
        Public ReadOnly Property PBILiveLoggingEndpoint() As String
            Get
                Return CType(Me("PBILiveLoggingEndpoint"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.SpecialSettingAttribute(Global.System.Configuration.SpecialSetting.WebServiceUrl)>  _
        Public ReadOnly Property PBIChargeIntentEndpoint() As String
            Get
                Return CType(Me("PBIChargeIntentEndpoint"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("5000")>  _
        Public ReadOnly Property PBIkWChartMinMax() As Single
            Get
                Return CType(Me("PBIkWChartMinMax"),Single)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("230")>  _
        Public ReadOnly Property PBIVoltageNominal() As Single
            Get
                Return CType(Me("PBIVoltageNominal"),Single)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public ReadOnly Property PWForceModeOnStartup() As Boolean
            Get
                Return CType(Me("PWForceModeOnStartup"),Boolean)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("self_consumption")>  _
        Public ReadOnly Property PWForceMode() As String
            Get
                Return CType(Me("PWForceMode"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("5")>  _
        Public ReadOnly Property PWForceBackupPercentage() As Byte
            Get
                Return CType(Me("PWForceBackupPercentage"),Byte)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public ReadOnly Property VerboseLogging() As Boolean
            Get
                Return CType(Me("VerboseLogging"),Boolean)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public ReadOnly Property PWControlEnabled() As Boolean
            Get
                Return CType(Me("PWControlEnabled"),Boolean)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public ReadOnly Property DebugLogging() As Boolean
            Get
                Return CType(Me("DebugLogging"),Boolean)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public ReadOnly Property PWOvernightStandby() As Boolean
            Get
                Return CType(Me("PWOvernightStandby"),Boolean)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("5000")>  _
        Public ReadOnly Property PWPeakConsumptionWeekend() As String
            Get
                Return CType(Me("PWPeakConsumptionWeekend"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public ReadOnly Property DualPVSystem() As Boolean
            Get
                Return CType(Me("DualPVSystem"),Boolean)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("0")>  _
        Public ReadOnly Property PVSystem2Azimuth() As Integer
            Get
                Return CType(Me("PVSystem2Azimuth"),Integer)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("0")>  _
        Public ReadOnly Property PVSystem2Capacity() As Integer
            Get
                Return CType(Me("PVSystem2Capacity"),Integer)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("0")>  _
        Public ReadOnly Property PVSystem2Tilt() As Integer
            Get
                Return CType(Me("PVSystem2Tilt"),Integer)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public ReadOnly Property PWCloudEmail() As String
            Get
                Return CType(Me("PWCloudEmail"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public ReadOnly Property PWCloudPassword() As String
            Get
                Return CType(Me("PWCloudPassword"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("https://owner-api.teslamotors.com/")>  _
        Public ReadOnly Property PWCloudAPI() As String
            Get
                Return CType(Me("PWCloudAPI"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("0")>  _
        Public ReadOnly Property PWEnergySiteID() As Long
            Get
                Return CType(Me("PWEnergySiteID"),Long)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public ReadOnly Property PWLocalUserUsername() As String
            Get
                Return CType(Me("PWLocalUserUsername"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public ReadOnly Property PWLocalUserPassword() As String
            Get
                Return CType(Me("PWLocalUserPassword"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public ReadOnly Property PWCloudMFARecoveryToken() As String
            Get
                Return CType(Me("PWCloudMFARecoveryToken"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public ReadOnly Property PWUseAutonomous() As Boolean
            Get
                Return CType(Me("PWUseAutonomous"),Boolean)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public ReadOnly Property PWCloudToken() As String
            Get
                Return CType(Me("PWCloudToken"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public ReadOnly Property PWPeakConsumptionUseHistory() As Boolean
            Get
                Return CType(Me("PWPeakConsumptionUseHistory"),Boolean)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public ReadOnly Property PWCloudRefreshToken() As String
            Get
                Return CType(Me("PWCloudRefreshToken"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public ReadOnly Property IgnoreCertificates() As Boolean
            Get
                Return CType(Me("IgnoreCertificates"),Boolean)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public ReadOnly Property PWConsumptionToPeakStartUseHistory() As Boolean
            Get
                Return CType(Me("PWConsumptionToPeakStartUseHistory"),Boolean)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public ReadOnly Property PWOvernightConsumptionUseHistory() As Boolean
            Get
                Return CType(Me("PWOvernightConsumptionUseHistory"),Boolean)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("0")>  _
        Public ReadOnly Property PWChargeBuffer() As Integer
            Get
                Return CType(Me("PWChargeBuffer"),Integer)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public ReadOnly Property PWSkipControl() As Boolean
            Get
                Return CType(Me("PWSkipControl"),Boolean)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("10")>  _
        Public ReadOnly Property TariffSuperOffPeakStart() As Integer
            Get
                Return CType(Me("TariffSuperOffPeakStart"),Integer)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("15")>  _
        Public ReadOnly Property TariffSuperOffPeakEnd() As Integer
            Get
                Return CType(Me("TariffSuperOffPeakEnd"),Integer)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public ReadOnly Property TariffSuperOffPeakActive() As Boolean
            Get
                Return CType(Me("TariffSuperOffPeakActive"),Boolean)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public ReadOnly Property PWAutoChargeMode() As Boolean
            Get
                Return CType(Me("PWAutoChargeMode"),Boolean)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public ReadOnly Property UseTeslaFleetAPI() As Boolean
            Get
                Return CType(Me("UseTeslaFleetAPI"),Boolean)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("openid offline_access user_data vehicle_device_data energy_device_data energy_cmd"& _ 
            "s")>  _
        Public ReadOnly Property TeslaFleetScope() As String
            Get
                Return CType(Me("TeslaFleetScope"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("https://fleet-api.prd.na.vn.cloud.tesla.com/")>  _
        Public ReadOnly Property TeslaFleetAPI() As String
            Get
                Return CType(Me("TeslaFleetAPI"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.SpecialSettingAttribute(Global.System.Configuration.SpecialSetting.ConnectionString),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Server=tcp:TODO.database.windows.net,1433;Initial Catalog=PWHistory;Persist Secur"& _ 
            "ity Info=True;User ID=TODO;Password=TODO;MultipleActiveResultSets=False;Encrypt="& _ 
            "True;TrustServerCertificate=False;Connection Timeout=30;")>  _
        Public ReadOnly Property PWHistory() As String
            Get
                Return CType(Me("PWHistory"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.SpecialSettingAttribute(Global.System.Configuration.SpecialSetting.ConnectionString),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Data Source=TODO;Initial Catalog=PWHistory;Persist Security Info=True;User ID=TOD"& _ 
            "O;Password=TODO")>  _
        Public ReadOnly Property PWHistoryLocal() As String
            Get
                Return CType(Me("PWHistoryLocal"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("http://wsl2:8123/api/")>  _
        Public ReadOnly Property HAHost() As String
            Get
                Return CType(Me("HAHost"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("states/input_boolean.override_forecast")>  _
        Public ReadOnly Property HAOverrideForecast() As String
            Get
                Return CType(Me("HAOverrideForecast"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("TODO")>  _
        Public ReadOnly Property HABearer() As String
            Get
                Return CType(Me("HABearer"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public ReadOnly Property UseHA() As Boolean
            Get
                Return CType(Me("UseHA"),Boolean)
            End Get
        End Property
    End Class
End Namespace

Namespace My
    
    <Global.Microsoft.VisualBasic.HideModuleNameAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute()>  _
    Friend Module MySettingsProperty
        
        <Global.System.ComponentModel.Design.HelpKeywordAttribute("My.Settings")>  _
        Friend ReadOnly Property Settings() As Global.PowerwallService.My.MySettings
            Get
                Return Global.PowerwallService.My.MySettings.Default
            End Get
        End Property
    End Module
End Namespace
