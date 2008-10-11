VERSION 5.00
Object = "{45A0E20C-D21B-11D5-B730-00B0D039C0EF}#1.0#0"; "PresenceInterfaceX.ocx"
Begin VB.Form frmActivar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Activar Presence Producción"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin PresenceInterfaceXControl1.PresenceInterfaceX PresenceX 
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   2400
      Width           =   615
      LineActive      =   -1
      ContactCode     =   -1
      Phone2          =   ""
      ScheduledDate   =   0
      Comments        =   ""
      ContactName     =   ""
      CaptureCall     =   -1
      CaptureCallDateLimit=   0
      EMailOutFrom    =   ""
      EMailOutTo      =   ""
      EMailOutSubject =   ""
      EMailOutMessage =   ""
      DoubleBuffered  =   0   'False
      Enabled         =   -1  'True
      Object.Visible         =   -1  'True
      Cursor          =   0
      ClientInfo      =   ""
   End
   Begin VB.CommandButton btAnadir 
      Caption         =   "Anadir"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox edServicio 
      Height          =   285
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton btEliminar 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
   Begin VB.ListBox lServicios 
      Height          =   1035
      ItemData        =   "frmActivar.frx":0000
      Left            =   360
      List            =   "frmActivar.frx":0002
      TabIndex        =   2
      Top             =   1080
      Width           =   2655
   End
   Begin VB.CommandButton btActivar 
      Caption         =   "Activar"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Servicios a conectar:"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Servicio:"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmActivar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub PropiedadesContacto()
    frmDatosContacto.edCodigoServicio.Text = PresenceX.ServiceId
    frmDatosContacto.edVDN.Text = PresenceX.VDN
    frmDatosContacto.edTelefono.Text = PresenceX.Phone2
    frmDatosContacto.edContactID.Text = PresenceX.ContactId
    frmDatosContacto.edClientID.Text = PresenceX.ClientId
End Sub
Sub LimpiarPropiedadesContacto()
    frmDatosContacto.edCodigoServicio.Text = ""
    frmDatosContacto.edVDN.Text = ""
    frmDatosContacto.edTelefono.Text = ""
    frmDatosContacto.edContactID.Text = ""
    frmDatosContacto.edClientID.Text = ""
    frmDatosContacto.edTipoEvento.Text = ""
    frmDatosContacto.edFinal.Text = ""
    frmDatosContacto.cbFinalizar.Value = 0
End Sub
Private Sub btActivar_Click()
    Dim i As Integer
    
    PresenceX.Active
    
    For i = 0 To lServicios.ListCount - 1
        PresenceX.ConnectToService lServicios.List(i)
    Next
    
    frmActivar.Hide
    
End Sub

Private Sub btAnadir_Click()
    lServicios.AddItem edServicio.Text
End Sub

Private Sub btEliminar_Click()
    If lServicios.ListIndex >= 0 Then
        lServicios.RemoveItem lServicios.ListIndex
    End If
End Sub

Private Sub PresenceX_CloseEvent()
    Unload frmDatosContacto
    Unload Me
End Sub

Private Sub PresenceX_EndContactEvent(EndContact As Boolean)
    If frmDatosContacto.cbFinalizar.Value = 1 Then
        EndContact = True
        LimpiarPropiedadesContacto
    Else
        EndContact = False
    End If
End Sub

Private Sub PresenceX_InboundCallEvent()
    frmDatosContacto.edTipoEvento.Text = "InboundCallEvent"
    
    PropiedadesContacto
End Sub

Private Sub PresenceX_LoginEvent()
    Load frmDatosContacto
    frmDatosContacto.Show 0
    frmDatosContacto.edAgente.Text = PresenceX.AgentId
    frmDatosContacto.edExtension.Text = PresenceX.AgentStation
End Sub

Private Sub PresenceX_NewEndCodeEvent(ByVal EndCode As Long)
    frmDatosContacto.edFinal.Text = EndCode
End Sub

Private Sub PresenceX_OutboundCallEvent()
    frmDatosContacto.edTipoEvento.Text = "OutboundCallEvent"
    
    PropiedadesContacto
End Sub
