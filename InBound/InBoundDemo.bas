Attribute VB_Name = "mmMain"
Sub Main()
    ' Activación de Presence Agent
    frmInBoundDemo.PresenceInterfaceX1.Active
    'Conexión a los servicios
    frmInBoundDemo.PresenceInterfaceX1.ConnectToService (2001)
    frmInBoundDemo.PresenceInterfaceX1.ConnectToService (2051)
End Sub
