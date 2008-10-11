VERSION 5.00
Begin VB.Form frmDatosContacto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos contacto"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox edFinal 
      Height          =   285
      Left            =   2160
      TabIndex        =   19
      Top             =   4080
      Width           =   975
   End
   Begin VB.CheckBox cbFinalizar 
      Caption         =   "Finalizar contacto"
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   4560
      Width           =   1815
   End
   Begin VB.TextBox edExtension 
      Height          =   285
      Left            =   1440
      TabIndex        =   10
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox edAgente 
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      Top             =   360
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Información contaco:"
      Height          =   2655
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   3975
      Begin VB.TextBox edClientID 
         Height          =   285
         Left            =   1680
         TabIndex        =   16
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox edContactID 
         Height          =   285
         Left            =   1680
         TabIndex        =   15
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox edTelefono 
         Height          =   285
         Left            =   1680
         TabIndex        =   14
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox edVDN 
         Height          =   285
         Left            =   1680
         TabIndex        =   13
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox edCodigoServicio 
         Height          =   285
         Left            =   1680
         TabIndex        =   12
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox edTipoEvento 
         Height          =   285
         Left            =   1680
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Código Servicio:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Client ID:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Contact ID:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo evento:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Telefóno:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "VDN/CDN:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   855
      End
   End
   Begin VB.Label Label9 
      Caption         =   "Motivo de cierre (final):"
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Extensión:"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Id. Agente:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "frmDatosContacto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
