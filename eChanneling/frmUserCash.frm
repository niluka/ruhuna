VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmUserCash 
   Caption         =   "User Cash"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10380
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
   ScaleHeight     =   8340
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List5 
      Height          =   780
      Left            =   5640
      TabIndex        =   6
      Top             =   3240
      Width           =   3015
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2820
      Left            =   5640
      TabIndex        =   4
      Top             =   360
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4974
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   70320129
      CurrentDate     =   39485
   End
   Begin VB.ListBox List4 
      Height          =   3660
      Left            =   5040
      TabIndex        =   3
      Top             =   360
      Width           =   375
   End
   Begin VB.ListBox List3 
      Height          =   3660
      Left            =   2880
      TabIndex        =   2
      Top             =   360
      Width           =   2535
   End
   Begin VB.ListBox List2 
      Height          =   3660
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Width           =   495
   End
   Begin VB.ListBox List1 
      Height          =   3660
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Settling Credit Bookings Value"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   4920
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Total Cash Bookings Value"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   4560
      Width           =   2415
   End
End
Attribute VB_Name = "frmUserCash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

