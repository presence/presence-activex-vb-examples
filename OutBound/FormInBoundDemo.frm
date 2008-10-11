VERSION 5.00
Object = "{45A0E20C-D21B-11D5-B730-00B0D039C0EF}#1.0#0"; "PresenceInterfaceX.ocx"
Begin VB.Form frmInBoundDemo 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contacto"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7665
   FillColor       =   &H00E0E0E0&
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin PresenceInterfaceXControl1.PresenceInterfaceX PresenceInterfaceX1 
      Height          =   615
      Left            =   7080
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   375
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
      AllowEndContact =   0   'False
      QueuedContactsEventTimer=   -1
      ClientId        =   -1
   End
   Begin VB.CommandButton btModificar 
      Caption         =   "Modificar comrpra"
      Height          =   375
      Left            =   6000
      TabIndex        =   25
      Top             =   5400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton btAceptar 
      Caption         =   "Validar comrpra"
      Height          =   375
      Left            =   6000
      TabIndex        =   24
      Top             =   4920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   2160
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   5400
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2160
      TabIndex        =   22
      Text            =   "Combo1"
      Top             =   5036
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   21
      Top             =   5404
      Width           =   885
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Producto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   20
      Top             =   5040
      Width           =   885
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Información de la compra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   4515
      Width           =   3615
   End
   Begin VB.Label lbMovil2 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2160
      TabIndex        =   17
      Tag             =   "2400"
      Top             =   3645
      Width           =   1935
   End
   Begin VB.Label lbPais2 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2160
      TabIndex        =   16
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label lbCiudad2 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2160
      TabIndex        =   15
      Tag             =   "2400"
      Top             =   2835
      Width           =   2895
   End
   Begin VB.Label lbDireccion2 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2160
      TabIndex        =   14
      Top             =   2445
      Width           =   4095
   End
   Begin VB.Label lbNombre2 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2160
      TabIndex        =   13
      Tag             =   "2400"
      Top             =   2040
      Width           =   4095
   End
   Begin VB.Label lbCliente2 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2160
      TabIndex        =   12
      Top             =   4515
      Width           =   4095
   End
   Begin VB.Label lbCiudad 
      BackStyle       =   0  'Transparent
      Caption         =   "Ciudad:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   375
      TabIndex        =   11
      Top             =   2880
      Width           =   885
   End
   Begin VB.Label lbMovil 
      BackStyle       =   0  'Transparent
      Caption         =   "Móvil:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   375
      TabIndex        =   10
      Top             =   3675
      Width           =   885
   End
   Begin VB.Label lbPais 
      BackStyle       =   0  'Transparent
      Caption         =   "País:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   375
      TabIndex        =   9
      Top             =   3285
      Width           =   885
   End
   Begin VB.Label lbDireccion 
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   375
      TabIndex        =   8
      Top             =   2445
      Width           =   885
   End
   Begin VB.Label lbNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   375
      TabIndex        =   7
      Top             =   2040
      Width           =   720
   End
   Begin VB.Label lbCliente 
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   375
      TabIndex        =   6
      Top             =   4515
      Width           =   660
   End
   Begin VB.Label lbTitulo3 
      BackStyle       =   0  'Transparent
      Caption         =   "Información del cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1515
      Width           =   3615
   End
   Begin VB.Label lbTitulo2 
      BackStyle       =   0  'Transparent
      Caption         =   "Información del contacto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label lbTelefono2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   1080
      Width           =   1500
   End
   Begin VB.Label lbServicio2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   675
      Width           =   1500
   End
   Begin VB.Label lbTelefono 
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   375
      TabIndex        =   1
      Top             =   1080
      Width           =   825
   End
   Begin VB.Label lbServicio 
      BackStyle       =   0  'Transparent
      Caption         =   "Servicio:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   375
      TabIndex        =   0
      Top             =   675
      Width           =   765
   End
   Begin VB.Shape spTitulo2 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   120
      Top             =   165
      Width           =   7395
   End
   Begin VB.Shape spTitulo3 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   120
      Top             =   1440
      Width           =   7395
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   120
      Top             =   4440
      Width           =   7395
   End
End
Attribute VB_Name = "frmInBoundDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub RellenarCampos()
    lbNombre2.Caption = deDBCliente.rsSQLDBCliente.Fields("Nombre").Value
    lbDireccion2.Caption = deDBCliente.rsSQLDBCliente.Fields("Direccion").Value
    lbCiudad2.Caption = deDBCliente.rsSQLDBCliente.Fields("Ciudad").Value
    lbPais2.Caption = deDBCliente.rsSQLDBCliente.Fields("Pais").Value
    lbMovil2.Caption = deDBCliente.rsSQLDBCliente.Fields("Movil").Value
End Sub
' Busqueda clientes por nº teléfono
Public Sub BuscarClienteTelefono(ClienteTelefono As String)
    deDBCliente.rsSQLDBCliente.Open
    deDBCliente.rsSQLDBCliente.Find ("Telefono = '" & ClienteTelefono & "'")
    RellenarCampos
    deDBCliente.rsSQLDBCliente.Close
End Sub

Private Sub Form_Terminate()
    Unload Me
End Sub

Private Sub PresenceInterfaceX1_EndContactEvent(EndContact As Boolean)
    EndContact = True
    Hide
End Sub

Private Sub PresenceInterfaceX1_InboundCallEvent()
   ' Obtenemos el servicio
    lbServicio2.Caption = CLng(PresenceInterfaceX1.ServiceId)
    ' Obtenemos el teléfono que nos llama o al que hemos llamado
    lbTelefono2.Caption = PresenceInterfaceX1.Phone2
    ' Obtenemos el identifiador del cliente
    lbCliente2.Caption = PresenceInterfaceX1.ClientId
    BuscarClienteTelefono (PresenceInterfaceX1.Phone2)
    Show
End Sub

