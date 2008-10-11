Attribute VB_Name = "mmMain"
Sub Main()
    ' Activación Presence Agent
    frmOutBoundDemo.PresenceInterfaceX1.Active
    ' Conexión a lo servicios
    frmOutBoundDemo.PresenceInterfaceX1.ConnectToService (2001)
    frmOutBoundDemo.PresenceInterfaceX1.ConnectToService (2051)
End Sub
