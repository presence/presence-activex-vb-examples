Attribute VB_Name = "mmMain"
Sub Main()
    ' Activaci�n Presence Agent
    frmOutBoundDemo.PresenceInterfaceX1.Active
    ' Conexi�n a lo servicios
    frmOutBoundDemo.PresenceInterfaceX1.ConnectToService (2001)
    frmOutBoundDemo.PresenceInterfaceX1.ConnectToService (2051)
End Sub
