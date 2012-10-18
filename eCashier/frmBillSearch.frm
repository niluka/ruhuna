VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmBillSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bill Search"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10650
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   10650
   Begin VB.TextBox txtBillDetails 
      Height          =   5175
      Left            =   120
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   960
      Width           =   10215
   End
   Begin btButtonEx.ButtonEx btnSearch 
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Search"
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
   Begin VB.TextBox txtID 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin btButtonEx.ButtonEx ButtonEx1 
      Height          =   615
      Left            =   8520
      TabIndex        =   4
      Top             =   6240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      Appearance      =   3
      Caption         =   "&Search"
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
   Begin VB.Label Label1 
      Caption         =   "Bill ID"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmBillSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
