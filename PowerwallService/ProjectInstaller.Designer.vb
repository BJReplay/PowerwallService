<System.ComponentModel.RunInstaller(True)> Partial Class ProjectInstaller
    Inherits System.Configuration.Install.Installer

    'Installer overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.PowerwallServiceProcessInstaller = New System.ServiceProcess.ServiceProcessInstaller()
        Me.PowerwallServiceInstaller = New System.ServiceProcess.ServiceInstaller()
        '
        'PowerwallServiceProcessInstaller
        '
        Me.PowerwallServiceProcessInstaller.Account = System.ServiceProcess.ServiceAccount.NetworkService
        Me.PowerwallServiceProcessInstaller.Password = Nothing
        Me.PowerwallServiceProcessInstaller.Username = Nothing
        '
        'PowerwallServiceInstaller
        '
        Me.PowerwallServiceInstaller.DelayedAutoStart = True
        Me.PowerwallServiceInstaller.Description = "This service can log data from a local Powerwall to a database, can control pre-c" &
    "harging during off-peak, and can upload data to PVOutput."
        Me.PowerwallServiceInstaller.DisplayName = "Powerwall Logging and Control Service"
        Me.PowerwallServiceInstaller.ServiceName = "PowerwallService"
        '
        'ProjectInstaller
        '
        Me.Installers.AddRange(New System.Configuration.Install.Installer() {Me.PowerwallServiceProcessInstaller, Me.PowerwallServiceInstaller})

    End Sub

    Friend WithEvents PowerwallServiceProcessInstaller As ServiceProcess.ServiceProcessInstaller
    Friend WithEvents PowerwallServiceInstaller As ServiceProcess.ServiceInstaller
End Class
