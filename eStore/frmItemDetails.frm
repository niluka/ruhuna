VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmItemDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Details"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3240
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
   ScaleHeight     =   3225
   ScaleWidth      =   3240
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      Begin btButtonEx.ButtonEx ButtonEx2 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   1500
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Arrange By Display Names"
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
      Begin btButtonEx.ButtonEx ButtonEx3 
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   930
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Arrange By Codes"
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
      Begin btButtonEx.ButtonEx ButtonEx4 
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Arrange By Categoty"
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
      Begin btButtonEx.ButtonEx ButtonEx5 
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   2070
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Arrange By Trade Names"
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
      Begin btButtonEx.ButtonEx btnPrices 
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   2640
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Prices && Distributors"
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
End
Attribute VB_Name = "frmItemDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ButtonEx1_Click()
    dtrItemView.Show
End Sub

Private Sub ButtonEx2_Click()
    dtrItemMaster.Show
End Sub

Private Sub ButtonEx3_Click()
    dtrItemCodeWise.Show
End Sub

Private Sub ButtonEx4_Click()
    dtrCategoryWise.Show
End Sub

Private Sub ButtonEx5_Click()
    dtrItemTradeNameWise.Show
End Sub
