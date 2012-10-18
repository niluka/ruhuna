VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBatch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Batch Details"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6855
   ControlBox      =   0   'False
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
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   5400
      TabIndex        =   8
      Top             =   2160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Close"
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
   Begin VB.TextBox txtBatch 
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   720
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker dtpDOM 
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   1200
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   20774915
      CurrentDate     =   39545
   End
   Begin MSComCtl2.DTPicker dtpDOE 
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   1680
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   20774915
      CurrentDate     =   39545
   End
   Begin VB.Label lblItem 
      Caption         =   "Label5"
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label Label4 
      Caption         =   "Item"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Date of Expiary"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Date of Manufacture"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Batch"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2055
   End
End
Attribute VB_Name = "frmBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public DOM As Date
Public DOE As Date
Public Batch As String

Private Sub bttnClose_Click()
    Dim TR As Integer
    If dtpDOE.Value = Date Then
        TR = MsgBox("You have not entered a Date of Expiary", vbCritical, "Expiary Date")
        dtpDOE.SetFocus
        Exit Sub
    End If
    If Trim(txtBatch.Text) = Empty Then
        TR = MsgBox("You have not entered a Batch number", vbCritical, "Expiary Date")
        dtpDOE.SetFocus
        Exit Sub
    End If
    DOE = dtpDOE.Value
    DOM = dtpDOM.Value
    Batch = txtBatch.Text
    Unload Me
End Sub

Private Sub Form_Load()
    lblItem.Caption = frmGoodReceive.Item
    If IsDate(frmGoodReceive.DOE) Then
        dtpDOE.Value = frmGoodReceive.DOE
    Else
        dtpDOE.Value = Date
    End If
    If IsDate(frmGoodReceive.DOM) Then
        dtpDOM.Value = frmGoodReceive.DOM
    Else
        dtpDOM.Value = Date
    End If
    txtBatch.Text = frmGoodReceive.Batch
    dtpDOE.MinDate = Date
    dtpDOM.MaxDate = Date
End Sub
