VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmDetailedPatientSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4800
   Icon            =   "frmDetailedPatientSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   4800
   Begin VB.Frame frameSearchPatient 
      Caption         =   "Search Patient"
      Height          =   7095
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   4335
      Begin VB.TextBox txtSearchSurname 
         Height          =   345
         Left            =   1320
         TabIndex        =   2
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox txtSearchFirstName 
         Height          =   345
         Left            =   1320
         TabIndex        =   1
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txtSearchID 
         Height          =   345
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   2775
      End
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   4815
         Left            =   120
         TabIndex        =   7
         Top             =   2160
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   8493
         _Version        =   393216
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
      End
      Begin btButtonEx.ButtonEx bttnSearch 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Se&arch"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "&Surname"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "&First Name"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "&ID"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2055
      End
   End
   Begin btButtonEx.ButtonEx bttnSelect 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   7320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Sele&ct"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx bttnCancel 
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   7320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Ca&ncel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmDetailedPatientSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
