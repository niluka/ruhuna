VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmIxDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Investigation Format"
   ClientHeight    =   11010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
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
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.Frame FrameIx 
      Height          =   10455
      Left            =   4440
      TabIndex        =   1
      Top             =   0
      Width           =   10695
      Begin VB.TextBox txtRef15 
         Height          =   285
         Left            =   6960
         TabIndex        =   86
         Top             =   5760
         Width           =   3615
      End
      Begin VB.TextBox txtRef14 
         Height          =   285
         Left            =   6960
         TabIndex        =   85
         Top             =   5400
         Width           =   3615
      End
      Begin VB.TextBox txtRef13 
         Height          =   285
         Left            =   6960
         TabIndex        =   84
         Top             =   5040
         Width           =   3615
      End
      Begin VB.TextBox txtRef12 
         Height          =   285
         Left            =   6960
         TabIndex        =   83
         Top             =   4680
         Width           =   3615
      End
      Begin VB.TextBox txtRef11 
         Height          =   285
         Left            =   6960
         TabIndex        =   82
         Top             =   4320
         Width           =   3615
      End
      Begin VB.TextBox txtRef10 
         Height          =   285
         Left            =   6960
         TabIndex        =   81
         Top             =   3960
         Width           =   3615
      End
      Begin VB.TextBox txtRef9 
         Height          =   285
         Left            =   6960
         TabIndex        =   80
         Top             =   3600
         Width           =   3615
      End
      Begin VB.TextBox txtRef8 
         Height          =   285
         Left            =   6960
         TabIndex        =   79
         Top             =   3240
         Width           =   3615
      End
      Begin VB.TextBox txtRef7 
         Height          =   285
         Left            =   6960
         TabIndex        =   78
         Top             =   2880
         Width           =   3615
      End
      Begin VB.TextBox txtRef6 
         Height          =   285
         Left            =   6960
         TabIndex        =   77
         Top             =   2520
         Width           =   3615
      End
      Begin VB.TextBox txtRef5 
         Height          =   285
         Left            =   6960
         TabIndex        =   76
         Top             =   2160
         Width           =   3615
      End
      Begin VB.TextBox txtRef4 
         Height          =   285
         Left            =   6960
         TabIndex        =   75
         Top             =   1800
         Width           =   3615
      End
      Begin VB.TextBox txtRef3 
         Height          =   285
         Left            =   6960
         TabIndex        =   74
         Top             =   1440
         Width           =   3615
      End
      Begin VB.TextBox txtRef2 
         Height          =   285
         Left            =   6960
         TabIndex        =   73
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox txtRef1 
         Height          =   285
         Left            =   6960
         TabIndex        =   72
         Top             =   720
         Width           =   3615
      End
      Begin VB.TextBox txtUnit15 
         Height          =   285
         Left            =   4440
         TabIndex        =   71
         Top             =   5760
         Width           =   2295
      End
      Begin VB.TextBox txtUnit14 
         Height          =   285
         Left            =   4440
         TabIndex        =   70
         Top             =   5400
         Width           =   2295
      End
      Begin VB.TextBox txtUnit13 
         Height          =   285
         Left            =   4440
         TabIndex        =   69
         Top             =   5040
         Width           =   2295
      End
      Begin VB.TextBox txtUnit12 
         Height          =   285
         Left            =   4440
         TabIndex        =   68
         Top             =   4680
         Width           =   2295
      End
      Begin VB.TextBox txtUnit11 
         Height          =   285
         Left            =   4440
         TabIndex        =   67
         Top             =   4320
         Width           =   2295
      End
      Begin VB.TextBox txtUnit10 
         Height          =   285
         Left            =   4440
         TabIndex        =   66
         Top             =   3960
         Width           =   2295
      End
      Begin VB.TextBox txtUnit9 
         Height          =   285
         Left            =   4440
         TabIndex        =   65
         Top             =   3600
         Width           =   2295
      End
      Begin VB.TextBox txtUnit8 
         Height          =   285
         Left            =   4440
         TabIndex        =   64
         Top             =   3240
         Width           =   2295
      End
      Begin VB.TextBox txtUnit7 
         Height          =   285
         Left            =   4440
         TabIndex        =   63
         Top             =   2880
         Width           =   2295
      End
      Begin VB.TextBox txtUnit6 
         Height          =   285
         Left            =   4440
         TabIndex        =   62
         Top             =   2520
         Width           =   2295
      End
      Begin VB.TextBox txtUnit5 
         Height          =   285
         Left            =   4440
         TabIndex        =   61
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox txtUnit4 
         Height          =   285
         Left            =   4440
         TabIndex        =   60
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox txtUnit3 
         Height          =   285
         Left            =   4440
         TabIndex        =   59
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox txtUnit2 
         Height          =   285
         Left            =   4440
         TabIndex        =   58
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txtUnit1 
         Height          =   285
         Left            =   4440
         TabIndex        =   57
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtField15 
         Height          =   285
         Left            =   360
         TabIndex        =   56
         Top             =   5760
         Width           =   3735
      End
      Begin VB.TextBox txtField14 
         Height          =   285
         Left            =   360
         TabIndex        =   55
         Top             =   5400
         Width           =   3735
      End
      Begin VB.TextBox txtField13 
         Height          =   285
         Left            =   360
         TabIndex        =   54
         Top             =   5040
         Width           =   3735
      End
      Begin VB.TextBox txtField12 
         Height          =   285
         Left            =   360
         TabIndex        =   53
         Top             =   4680
         Width           =   3735
      End
      Begin VB.TextBox txtField11 
         Height          =   285
         Left            =   360
         TabIndex        =   52
         Top             =   4320
         Width           =   3735
      End
      Begin VB.TextBox txtField10 
         Height          =   285
         Left            =   360
         TabIndex        =   51
         Top             =   3960
         Width           =   3735
      End
      Begin VB.TextBox txtField9 
         Height          =   285
         Left            =   360
         TabIndex        =   50
         Top             =   3600
         Width           =   3735
      End
      Begin VB.TextBox txtField8 
         Height          =   285
         Left            =   360
         TabIndex        =   49
         Top             =   3240
         Width           =   3735
      End
      Begin VB.TextBox txtField7 
         Height          =   285
         Left            =   360
         TabIndex        =   48
         Top             =   2880
         Width           =   3735
      End
      Begin VB.TextBox txtField6 
         Height          =   285
         Left            =   360
         TabIndex        =   47
         Top             =   2520
         Width           =   3735
      End
      Begin VB.TextBox txtField5 
         Height          =   285
         Left            =   360
         TabIndex        =   46
         Top             =   2160
         Width           =   3735
      End
      Begin VB.TextBox txtField4 
         Height          =   285
         Left            =   360
         TabIndex        =   45
         Top             =   1800
         Width           =   3735
      End
      Begin VB.TextBox txtField3 
         Height          =   285
         Left            =   360
         TabIndex        =   44
         Top             =   1440
         Width           =   3735
      End
      Begin VB.TextBox txtField2 
         Height          =   285
         Left            =   360
         TabIndex        =   43
         Top             =   1080
         Width           =   3735
      End
      Begin VB.TextBox txtField1 
         Height          =   285
         Left            =   360
         TabIndex        =   42
         Top             =   720
         Width           =   3735
      End
      Begin VB.TextBox txtIx 
         Height          =   285
         Left            =   2280
         TabIndex        =   41
         Top             =   120
         Width           =   6735
      End
      Begin VB.TextBox txtComments 
         Height          =   615
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   40
         Top             =   9720
         Width           =   7815
      End
      Begin VB.TextBox txtField16 
         Height          =   285
         Left            =   360
         TabIndex        =   39
         Top             =   6120
         Width           =   3735
      End
      Begin VB.TextBox txtUnit16 
         Height          =   285
         Left            =   4440
         TabIndex        =   38
         Top             =   6120
         Width           =   2295
      End
      Begin VB.TextBox txtRef16 
         Height          =   285
         Left            =   6960
         TabIndex        =   37
         Top             =   6120
         Width           =   3615
      End
      Begin VB.TextBox txtField17 
         Height          =   285
         Left            =   360
         TabIndex        =   36
         Top             =   6480
         Width           =   3735
      End
      Begin VB.TextBox txtUnit17 
         Height          =   285
         Left            =   4440
         TabIndex        =   35
         Top             =   6480
         Width           =   2295
      End
      Begin VB.TextBox txtRef17 
         Height          =   285
         Left            =   6960
         TabIndex        =   34
         Top             =   6480
         Width           =   3615
      End
      Begin VB.TextBox txtField18 
         Height          =   285
         Left            =   360
         TabIndex        =   33
         Top             =   6840
         Width           =   3735
      End
      Begin VB.TextBox txtUnit18 
         Height          =   285
         Left            =   4440
         TabIndex        =   32
         Top             =   6840
         Width           =   2295
      End
      Begin VB.TextBox txtRef18 
         Height          =   285
         Left            =   6960
         TabIndex        =   31
         Top             =   6840
         Width           =   3615
      End
      Begin VB.TextBox txtField19 
         Height          =   285
         Left            =   360
         TabIndex        =   30
         Top             =   7200
         Width           =   3735
      End
      Begin VB.TextBox txtUnit19 
         Height          =   285
         Left            =   4440
         TabIndex        =   29
         Top             =   7200
         Width           =   2295
      End
      Begin VB.TextBox txtRef19 
         Height          =   285
         Left            =   6960
         TabIndex        =   28
         Top             =   7200
         Width           =   3615
      End
      Begin VB.TextBox txtField20 
         Height          =   285
         Left            =   360
         TabIndex        =   27
         Top             =   7560
         Width           =   3735
      End
      Begin VB.TextBox txtUnit20 
         Height          =   285
         Left            =   4440
         TabIndex        =   26
         Top             =   7560
         Width           =   2295
      End
      Begin VB.TextBox txtRef20 
         Height          =   285
         Left            =   6960
         TabIndex        =   25
         Top             =   7560
         Width           =   3615
      End
      Begin VB.TextBox txtField21 
         Height          =   285
         Left            =   360
         TabIndex        =   24
         Top             =   7920
         Width           =   3735
      End
      Begin VB.TextBox txtUnit21 
         Height          =   285
         Left            =   4440
         TabIndex        =   23
         Top             =   7920
         Width           =   2295
      End
      Begin VB.TextBox txtRef21 
         Height          =   285
         Left            =   6960
         TabIndex        =   22
         Top             =   7920
         Width           =   3615
      End
      Begin VB.TextBox txtField22 
         Height          =   285
         Left            =   360
         TabIndex        =   21
         Top             =   8280
         Width           =   3735
      End
      Begin VB.TextBox txtUnit22 
         Height          =   285
         Left            =   4440
         TabIndex        =   20
         Top             =   8280
         Width           =   2295
      End
      Begin VB.TextBox txtRef22 
         Height          =   285
         Left            =   6960
         TabIndex        =   19
         Top             =   8280
         Width           =   3615
      End
      Begin VB.TextBox txtField23 
         Height          =   285
         Left            =   360
         TabIndex        =   18
         Top             =   8640
         Width           =   3735
      End
      Begin VB.TextBox txtUnit23 
         Height          =   285
         Left            =   4440
         TabIndex        =   17
         Top             =   8640
         Width           =   2295
      End
      Begin VB.TextBox txtRef23 
         Height          =   285
         Left            =   6960
         TabIndex        =   16
         Top             =   8640
         Width           =   3615
      End
      Begin VB.TextBox txtField24 
         Height          =   285
         Left            =   360
         TabIndex        =   15
         Top             =   9000
         Width           =   3735
      End
      Begin VB.TextBox txtUnit24 
         Height          =   285
         Left            =   4440
         TabIndex        =   14
         Top             =   9000
         Width           =   2295
      End
      Begin VB.TextBox txtRef24 
         Height          =   285
         Left            =   6960
         TabIndex        =   13
         Top             =   9000
         Width           =   3615
      End
      Begin VB.TextBox txtRef25 
         Height          =   285
         Left            =   6960
         TabIndex        =   12
         Top             =   9360
         Width           =   3615
      End
      Begin VB.TextBox txtUnit25 
         Height          =   285
         Left            =   4440
         TabIndex        =   11
         Top             =   9360
         Width           =   2295
      End
      Begin VB.TextBox txtField25 
         Height          =   285
         Left            =   360
         TabIndex        =   10
         Top             =   9360
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "Feild Name"
         Height          =   255
         Left            =   720
         TabIndex        =   91
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Unit Value"
         Height          =   255
         Left            =   3840
         TabIndex        =   90
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Ref Range"
         Height          =   255
         Left            =   8040
         TabIndex        =   89
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Investigation Name"
         Height          =   255
         Left            =   480
         TabIndex        =   88
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "Comments"
         Height          =   375
         Left            =   360
         TabIndex        =   87
         Top             =   9720
         Width           =   2055
      End
   End
   Begin VB.TextBox txtSearch 
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3840
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin btButtonEx.ButtonEx bttnAdd 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   6240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      Appearance      =   3
      Caption         =   "A&dd"
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
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   5655
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   9975
      _Version        =   393216
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
   End
   Begin btButtonEx.ButtonEx bttnEdit 
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   6240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      Appearance      =   3
      Caption         =   "Ed&it"
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
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   13560
      TabIndex        =   5
      Top             =   10560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Cl&ose"
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
   Begin btButtonEx.ButtonEx bttnDelete 
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   6240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      Appearance      =   3
      Caption         =   "&Delete"
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
      Height          =   255
      Left            =   10080
      TabIndex        =   7
      Top             =   10560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
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
   Begin btButtonEx.ButtonEx bttnSave 
      Height          =   255
      Left            =   6120
      TabIndex        =   8
      Top             =   10560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      Appearance      =   3
      Caption         =   "Sa&ve"
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
   Begin btButtonEx.ButtonEx bttnChange 
      Height          =   255
      Left            =   6840
      TabIndex        =   9
      Top             =   10560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      Appearance      =   3
      Caption         =   "C&hange"
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
Attribute VB_Name = "frmIxDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim FromGrid As Boolean
    Dim TemIxId As Long
    Dim BorderMargin As Long

Private Sub bttnAdd_Click()
    Call AfterAdd
    Call ClearValues
    txtIx.SetFocus
End Sub

Private Sub bttnCancel_Click()
    Call ClearValues
    Call BeforeAddEdit
End Sub

Private Sub bttnChange_Click()
Dim TemResponce As Byte

If Trim(txtIx.Text) = "" Then
    TemResponce = MsgBox("You must enter a name for the investigation", vbCritical, "No Name")
    txtIx.SetFocus
    Exit Sub
End If

Call EditData
Call ClearValues
Call BeforeAddEdit

End Sub

Private Sub EditData()
With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    .Source = "SELECT ix , ixid from tblinvcestigationdetails where ixid = " & TemIxId
    If .State = 0 Then .Open
    If .RecordCount = 0 Then .Close: Exit Sub
    !ix = Trim(txtIx.Text)
    !field1 = Trim(txtField1.Text)
    !field2 = Trim(txtField2.Text)
    !field3 = Trim(txtField3.Text)
    !field4 = Trim(txtField4.Text)
    !field5 = Trim(txtField5.Text)
    !field6 = Trim(txtField6.Text)
    !field7 = Trim(txtField7.Text)
    !field8 = Trim(txtField8.Text)
    !field9 = Trim(txtField9.Text)
    !field10 = Trim(txtField10.Text)
    !field11 = Trim(txtField11.Text)
    !field12 = Trim(txtField12.Text)
    !field13 = Trim(txtField13.Text)
    !field14 = Trim(txtField14.Text)
    !Field15 = Trim(txtField15.Text)
    !field16 = Trim(txtField16.Text)
    !field17 = Trim(txtField17.Text)
    !field18 = Trim(txtField18.Text)
    !field19 = Trim(txtField19.Text)
    !field20 = Trim(txtField20.Text)
    !field21 = Trim(txtField21.Text)
    !field22 = Trim(txtField22.Text)
    !field23 = Trim(txtField23.Text)
    !field24 = Trim(txtField24.Text)
    !Field25 = Trim(txtField25.Text)
    
   
    !fieldunit1 = Trim(txtUnit1.Text)
    !fieldunit2 = Trim(txtUnit2.Text)
    !fieldunit3 = Trim(txtUnit3.Text)
    !fieldunit4 = Trim(txtUnit4.Text)
    !fieldunit5 = Trim(txtUnit5.Text)
    !fieldunit6 = Trim(txtUnit6.Text)
    !fieldunit7 = Trim(txtUnit7.Text)
    !fieldunit8 = Trim(txtUnit8.Text)
    !fieldunit9 = Trim(txtUnit9.Text)
    !fieldunit10 = Trim(txtUnit10.Text)
    !fieldunit11 = Trim(txtUnit11.Text)
    !fieldunit12 = Trim(txtUnit12.Text)
    !fieldunit13 = Trim(txtUnit13.Text)
    !fieldunit14 = Trim(txtUnit14.Text)
    !Fieldunit15 = Trim(txtUnit15.Text)
    !fieldunit16 = Trim(txtUnit16.Text)
    !fieldunit17 = Trim(txtUnit17.Text)
    !fieldunit18 = Trim(txtUnit18.Text)
    !fieldunit19 = Trim(txtUnit19.Text)
    !fieldunit20 = Trim(txtUnit20.Text)
    !fieldunit21 = Trim(txtUnit21.Text)
    !fieldunit22 = Trim(txtUnit22.Text)
    !fieldunit23 = Trim(txtUnit23.Text)
    !fieldunit24 = Trim(txtUnit24.Text)
    !Fieldunit25 = Trim(txtUnit25.Text)
    
    
    !fieldref1 = Trim(txtRef1.Text)
    !fieldref2 = Trim(txtRef2.Text)
    !fieldref3 = Trim(txtRef3.Text)
    !fieldref4 = Trim(txtRef4.Text)
    !fieldref5 = Trim(txtRef5.Text)
    !fieldref6 = Trim(txtRef6.Text)
    !fieldref7 = Trim(txtRef7.Text)
    !fieldref8 = Trim(txtRef8.Text)
    !fieldref9 = Trim(txtRef9.Text)
    !fieldref10 = Trim(txtRef10.Text)
    !fieldref11 = Trim(txtRef11.Text)
    !fieldref12 = Trim(txtRef12.Text)
    !fieldref13 = Trim(txtRef13.Text)
    !fieldref14 = Trim(txtRef14.Text)
    !Fieldref15 = Trim(txtRef15.Text)
    !fieldref16 = Trim(txtRef16.Text)
    !fieldref17 = Trim(txtRef17.Text)
    !fieldref18 = Trim(txtRef18.Text)
    !fieldref19 = Trim(txtRef19.Text)
    !fieldref20 = Trim(txtRef20.Text)
    !fieldref21 = Trim(txtRef21.Text)
    !fieldref22 = Trim(txtRef22.Text)
    !fieldref23 = Trim(txtRef23.Text)
    !fieldref24 = Trim(txtRef24.Text)
    !Fieldref25 = Trim(txtRef25.Text)
    
    !Comments = Trim(txtComments.Text)
    
    
    .Update
    Grid1.Col = 1
    Grid1.Text = Trim(txtIx.Text)
    Grid1.Col = 2
    Grid1.Text = !ixid
    TemIxId = !ixid
End With
End Sub

Private Sub FormatGrid()

Dim BorderMargin As Long
BorderMargin = 100


With Grid1
    .Clear
    
    .Rows = 1
    .Cols = 3
    
    .ColWidth(0) = 600
    .ColWidth(2) = 1
    .ColWidth(1) = .Width - (.ColWidth(0) + .ColWidth(2) + BorderMargin)
    
    .Col = 0
    .CellAlignment = 4
    .Text = "No."
    
    .Col = 1
    .CellAlignment = 4
    .Text = "Investigation Name"
    
End With
End Sub

Private Sub FillGrid()
Dim NowRow As Long
With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    .Source = "SELECT ix , ixid from tblinvestigationdetails order by ix"
    If .State = 0 Then .Open
    If .RecordCount = 0 Then .Close: Exit Sub
    .MoveFirst
    NowRow = 0
    While .EOF = False
        NowRow = NowRow + 1
        Grid1.Rows = NowRow + 1
        Grid1.Row = NowRow
        Grid1.Col = 0
        Grid1.Text = NowRow
        Grid1.Col = 1
        Grid1.Text = !ix
        Grid1.Col = 2
        Grid1.Text = !ixid
        .MoveNext
    Wend

End With
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnDelete_Click()
    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        .Source = "SELECT ixid from tblinvestigationdetails where ixid = " & TemIxId
        If .State = 0 Then .Open
        If .RecordCount = 0 Then .Close:   Exit Sub
        .Delete adAffectCurrent
        .Close
    End With
    Dim TemNum As Long
    With Grid1
        .RemoveItem (Grid1.Row)
        .Col = 0
        
        .Sort = 1
        
        For TemNum = 1 To .Rows - 1
            .Row = TemNum
            .Text = TemNum
        Next
    End With
    
End Sub



Private Sub bttnEdit_Click()
    FromGrid = True
    Call AfterEdit
End Sub

Private Sub bttnSave_Click()
Dim TemResponce As Byte

If Trim(txtIx.Text) = "" Then
    TemResponce = MsgBox("You must enter a name for the investigation", vbCritical, "No Name")
    txtIx.SetFocus
    Exit Sub
End If

With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    .Source = "Select Ixid , ix from tblinvestigationdetails where ( ix = '" & Trim(txtIx.Text) & "')"
    If .State = 0 Then .Open
    If .RecordCount <> 0 Then
        TemResponce = MsgBox("The Investigation you entered already exist.", vbCritical, "Name Exists")
        txtIx.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    End If
End With

Call SaveDetails
Call ClearValues
Call BeforeAddEdit

End Sub


Private Sub SaveDetails()
Dim TemNum As Long

With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    .Source = "SELECT * from tblinvestigationdetails "
    If .State = 0 Then .Open
    .AddNew
    !ix = Trim(txtIx.Text)
    !field1 = Trim(txtField1.Text)
    !field2 = Trim(txtField2.Text)
    !field3 = Trim(txtField3.Text)
    !field4 = Trim(txtField4.Text)
    !field5 = Trim(txtField5.Text)
    !field6 = Trim(txtField6.Text)
    !field7 = Trim(txtField7.Text)
    !field8 = Trim(txtField8.Text)
    !field9 = Trim(txtField9.Text)
    !field10 = Trim(txtField10.Text)
    !field11 = Trim(txtField11.Text)
    !field12 = Trim(txtField12.Text)
    !field13 = Trim(txtField13.Text)
    !field14 = Trim(txtField14.Text)
    !Field15 = Trim(txtField15.Text)
    !field16 = Trim(txtField16.Text)
    !field17 = Trim(txtField17.Text)
    !field18 = Trim(txtField18.Text)
    !field19 = Trim(txtField19.Text)
    !field20 = Trim(txtField20.Text)
    !field21 = Trim(txtField21.Text)
    !field22 = Trim(txtField22.Text)
    !field23 = Trim(txtField23.Text)
    !field24 = Trim(txtField24.Text)
    !Field25 = Trim(txtField25.Text)
    
   
    !fieldunit1 = Trim(txtUnit1.Text)
    !fieldunit2 = Trim(txtUnit2.Text)
    !fieldunit3 = Trim(txtUnit3.Text)
    !fieldunit4 = Trim(txtUnit4.Text)
    !fieldunit5 = Trim(txtUnit5.Text)
    !fieldunit6 = Trim(txtUnit6.Text)
    !fieldunit7 = Trim(txtUnit7.Text)
    !fieldunit8 = Trim(txtUnit8.Text)
    !fieldunit9 = Trim(txtUnit9.Text)
    !fieldunit10 = Trim(txtUnit10.Text)
    !fieldunit11 = Trim(txtUnit11.Text)
    !fieldunit12 = Trim(txtUnit12.Text)
    !fieldunit13 = Trim(txtUnit13.Text)
    !fieldunit14 = Trim(txtUnit14.Text)
    !Fieldunit15 = Trim(txtUnit15.Text)
    !fieldunit16 = Trim(txtUnit16.Text)
    !fieldunit17 = Trim(txtUnit17.Text)
    !fieldunit18 = Trim(txtUnit18.Text)
    !fieldunit19 = Trim(txtUnit19.Text)
    !fieldunit20 = Trim(txtUnit20.Text)
    !fieldunit21 = Trim(txtUnit21.Text)
    !fieldunit22 = Trim(txtUnit22.Text)
    !fieldunit23 = Trim(txtUnit23.Text)
    !fieldunit24 = Trim(txtUnit24.Text)
    !Fieldunit25 = Trim(txtUnit25.Text)
    
    
    !fieldref1 = Trim(txtRef1.Text)
    !fieldref2 = Trim(txtRef2.Text)
    !fieldref3 = Trim(txtRef3.Text)
    !fieldref4 = Trim(txtRef4.Text)
    !fieldref5 = Trim(txtRef5.Text)
    !fieldref6 = Trim(txtRef6.Text)
    !fieldref7 = Trim(txtRef7.Text)
    !fieldref8 = Trim(txtRef8.Text)
    !fieldref9 = Trim(txtRef9.Text)
    !fieldref10 = Trim(txtRef10.Text)
    !fieldref11 = Trim(txtRef11.Text)
    !fieldref12 = Trim(txtRef12.Text)
    !fieldref13 = Trim(txtRef13.Text)
    !fieldref14 = Trim(txtRef14.Text)
    !Fieldref15 = Trim(txtRef15.Text)
    !fieldref16 = Trim(txtRef16.Text)
    !fieldref17 = Trim(txtRef17.Text)
    !fieldref18 = Trim(txtRef18.Text)
    !fieldref19 = Trim(txtRef19.Text)
    !fieldref20 = Trim(txtRef20.Text)
    !fieldref21 = Trim(txtRef21.Text)
    !fieldref22 = Trim(txtRef22.Text)
    !fieldref23 = Trim(txtRef23.Text)
    !fieldref24 = Trim(txtRef24.Text)
    !Fieldref25 = Trim(txtRef25.Text)
    
    !Comments = Trim(txtComments.Text)
    
    
    .Update
    
    TemIxId = !ixid
    Grid1.Rows = Grid1.Rows + 1
    Grid1.Row = Grid1.Rows - 1
    Grid1.Col = 1
    Grid1.Text = Trim(txtIx.Text)
    Grid1.Col = 2
    Grid1.Text = TemIxId
    Grid1.Col = 1
    Grid1.Sort = 1
    Grid1.Col = 0
    For TemNum = 1 To Grid1.Rows - 1
        Grid1.Row = TemNum
        Grid1.Text = TemNum
    Next

End With
End Sub

Private Sub Form_Load()
    Call FormatGrid
    Call FillGrid
    Call ClearValues
    Call BeforeAddEdit
    
End Sub

Private Sub Grid1_Click()
    
    FromGrid = True
            
            bttnAdd.Enabled = True
            
            With Grid1
                If .Row < 1 Then FromGrid = False: Exit Sub
                .Col = 2
                If Not IsNumeric(.Text) Then FromGrid = False: Exit Sub
                TemIxId = Val(.Text)
                .Col = 1
                txtSearch.Text = .Text
                
                Call GetData
                
                .Col = 0
                .ColSel = .Cols - 1
                
                txtSearch.SetFocus
                SendKeys "{home}+{end}"
            FromGrid = False
            bttnAdd.Enabled = False
            bttnEdit.Enabled = True
            bttnDelete.Enabled = True
        End With
End Sub

Private Sub ClearValues()
    txtIx.Text = Empty
    txtField1.Text = Empty
    txtField2.Text = Empty
    txtField3.Text = Empty
    txtField4.Text = Empty
    txtField5.Text = Empty
    txtField6.Text = Empty
    txtField7.Text = Empty
    txtField8.Text = Empty
    txtField9.Text = Empty
    txtField10.Text = Empty
    txtField11.Text = Empty
    txtField12.Text = Empty
    txtField13.Text = Empty
    txtField14.Text = Empty
    txtField15.Text = Empty
    txtField16.Text = Empty
    txtField17.Text = Empty
    txtField18.Text = Empty
    txtField19.Text = Empty
    txtField20.Text = Empty
    txtField21.Text = Empty
    txtField22.Text = Empty
    txtField23.Text = Empty
    txtField24.Text = Empty
    txtField25.Text = Empty
    txtUnit1.Text = Empty
    txtUnit2.Text = Empty
    txtUnit3.Text = Empty
    txtUnit4.Text = Empty
    txtUnit5.Text = Empty
    txtUnit6.Text = Empty
    txtUnit7.Text = Empty
    txtUnit8.Text = Empty
    txtUnit9.Text = Empty
    txtUnit10.Text = Empty
    txtUnit11.Text = Empty
    txtUnit12.Text = Empty
    txtUnit13.Text = Empty
    txtUnit14.Text = Empty
    txtUnit15.Text = Empty
    txtUnit16.Text = Empty
    txtUnit17.Text = Empty
    txtUnit18.Text = Empty
    txtUnit19.Text = Empty
    txtUnit20.Text = Empty
    txtUnit21.Text = Empty
    txtUnit22.Text = Empty
    txtUnit23.Text = Empty
    txtUnit24.Text = Empty
    txtUnit25.Text = Empty
    txtRef1.Text = Empty
    txtRef2.Text = Empty
    txtRef3.Text = Empty
    txtRef4.Text = Empty
    txtRef5.Text = Empty
    txtRef6.Text = Empty
    txtRef7.Text = Empty
    txtRef8.Text = Empty
    txtRef9.Text = Empty
    txtRef10.Text = Empty
    txtRef11.Text = Empty
    txtRef12.Text = Empty
    txtRef13.Text = Empty
    txtRef14.Text = Empty
    txtRef15.Text = Empty
    txtRef16.Text = Empty
    txtRef17.Text = Empty
    txtRef18.Text = Empty
    txtRef19.Text = Empty
    txtRef20.Text = Empty
    txtRef21.Text = Empty
    txtRef22.Text = Empty
    txtRef23.Text = Empty
    txtRef24.Text = Empty
    txtRef25.Text = Empty
    txtComments.Text = Empty
End Sub

Private Sub GetData()
With DataEnvironment1.rssqlInvestigations
    If .State = 1 Then .Close
    .Source = "SELECT * from tblinvestigationdetails where ixid = " & TemIxId
    If .State = 0 Then .Open
    If .RecordCount = 0 Then Exit Sub
        
    If Not IsNull(!ix) Then
        txtIx.Text = !ix
    End If
    
    If Not IsNull(!field1) Then
        txtField1.Text = !field1
    End If
    If Not IsNull(!field2) Then
        txtField2.Text = !field2
    End If
    If Not IsNull(!field3) Then
        txtField3.Text = !field3
    End If
    If Not IsNull(!field4) Then
        txtField4.Text = !field4
    End If
    If Not IsNull(!field5) Then
        txtField5.Text = !field5
    End If
    If Not IsNull(!field6) Then
        txtField6.Text = !field6
    End If
    If Not IsNull(!field7) Then
        txtField7.Text = !field7
    End If
    If Not IsNull(!field8) Then
        txtField8.Text = !field8
    End If
    If Not IsNull(!field9) Then
        txtField9.Text = !field9
    End If
    If Not IsNull(!field10) Then
        txtField10.Text = !field10
    End If
    If Not IsNull(!field11) Then
        txtField11.Text = !field11
    End If
    If Not IsNull(!field12) Then
        txtField12.Text = !field12
    End If
    If Not IsNull(!field13) Then
        txtField13.Text = !field13
    End If
    If Not IsNull(!field14) Then
        txtField14.Text = !field14
    End If
    If Not IsNull(!Field15) Then
        txtField15.Text = !Field15
    End If
    If Not IsNull(!field16) Then
        txtField16.Text = !field16
    End If
    If Not IsNull(!field17) Then
        txtField17.Text = !field17
    End If
    If Not IsNull(!field18) Then
        txtField18.Text = !field18
    End If
    If Not IsNull(!field19) Then
        txtField19.Text = !field19
    End If
    If Not IsNull(!field20) Then
        txtField20.Text = !field20
    End If
    If Not IsNull(!field21) Then
        txtField21.Text = !field21
    End If
    If Not IsNull(!field22) Then
        txtField22.Text = !field22
    End If
    If Not IsNull(!field23) Then
        txtField23.Text = !field23
    End If
    If Not IsNull(!field24) Then
        txtField24.Text = !field24
    End If
    If Not IsNull(!Field25) Then
        txtField25.Text = !Field25
    End If
    
    
    
    
    
    
    
    If Not IsNull(!fieldunit1) Then
        txtUnit1.Text = !fieldunit1
    End If
    If Not IsNull(!fieldunit2) Then
        txtUnit2.Text = !fieldunit2
    End If
    If Not IsNull(!fieldunit3) Then
        txtUnit3.Text = !fieldunit3
    End If
    If Not IsNull(!fieldunit4) Then
        txtUnit4.Text = !fieldunit4
    End If
    If Not IsNull(!fieldunit5) Then
        txtUnit5.Text = !fieldunit5
    End If
    If Not IsNull(!fieldunit6) Then
        txtUnit6.Text = !fieldunit6
    End If
    If Not IsNull(!fieldunit7) Then
        txtUnit7.Text = !fieldunit7
    End If
    If Not IsNull(!fieldunit8) Then
        txtUnit8.Text = !fieldunit8
    End If
    If Not IsNull(!fieldunit9) Then
        txtUnit9.Text = !fieldunit9
    End If
    If Not IsNull(!fieldunit10) Then
        txtUnit10.Text = !fieldunit10
    End If
    If Not IsNull(!fieldunit11) Then
        txtUnit11.Text = !fieldunit11
    End If
    If Not IsNull(!fieldunit12) Then
        txtUnit12.Text = !fieldunit12
    End If
    If Not IsNull(!fieldunit13) Then
        txtUnit13.Text = !fieldunit13
    End If
    If Not IsNull(!fieldunit14) Then
        txtUnit14.Text = !fieldunit14
    End If
    If Not IsNull(!Fieldunit15) Then
        txtUnit15.Text = !Fieldunit15
    End If
    If Not IsNull(!fieldunit16) Then
        txtUnit16.Text = !fieldunit16
    End If
    If Not IsNull(!fieldunit17) Then
        txtUnit17.Text = !fieldunit17
    End If
    If Not IsNull(!fieldunit18) Then
        txtUnit18.Text = !fieldunit18
    End If
    If Not IsNull(!fieldunit19) Then
        txtUnit19.Text = !fieldunit19
    End If
    If Not IsNull(!fieldunit20) Then
        txtUnit20.Text = !fieldunit20
    End If
    If Not IsNull(!fieldunit21) Then
        txtUnit21.Text = !fieldunit21
    End If
    If Not IsNull(!fieldunit22) Then
        txtUnit22.Text = !fieldunit22
    End If
    If Not IsNull(!fieldunit23) Then
        txtUnit23.Text = !fieldunit23
    End If
    If Not IsNull(!fieldunit24) Then
        txtUnit24.Text = !fieldunit24
    End If
    If Not IsNull(!Fieldunit25) Then
        txtUnit25.Text = !Fieldunit25
    End If
    
    
    
    
    
    
    
    If Not IsNull(!fieldref1) Then
        txtRef1.Text = !fieldref1
    End If
    If Not IsNull(!fieldref2) Then
        txtRef2.Text = !fieldref2
    End If
    If Not IsNull(!fieldref3) Then
        txtRef3.Text = !fieldref3
    End If
    If Not IsNull(!fieldref4) Then
        txtRef4.Text = !fieldref4
    End If
    If Not IsNull(!fieldref5) Then
        txtRef5.Text = !fieldref5
    End If
    If Not IsNull(!fieldref6) Then
        txtRef6.Text = !fieldref6
    End If
    If Not IsNull(!fieldref7) Then
        txtRef7.Text = !fieldref7
    End If
    If Not IsNull(!fieldref8) Then
        txtRef8.Text = !fieldref8
    End If
    If Not IsNull(!fieldref9) Then
        txtRef9.Text = !fieldref9
    End If
    If Not IsNull(!fieldref10) Then
        txtRef10.Text = !fieldref10
    End If
    If Not IsNull(!fieldref11) Then
        txtRef11.Text = !fieldref11
    End If
    If Not IsNull(!fieldref12) Then
        txtRef12.Text = !fieldref12
    End If
    If Not IsNull(!fieldref13) Then
        txtRef13.Text = !fieldref13
    End If
    If Not IsNull(!fieldref14) Then
        txtRef14.Text = !fieldref14
    End If
    If Not IsNull(!Fieldref15) Then
        txtRef15.Text = !Fieldref15
    End If
    
    If Not IsNull(!fieldref16) Then
        txtRef16.Text = !fieldref16
    End If
    If Not IsNull(!fieldref17) Then
        txtRef17.Text = !fieldref17
    End If
    If Not IsNull(!fieldref18) Then
        txtRef18.Text = !fieldref18
    End If
    If Not IsNull(!fieldref19) Then
        txtRef19.Text = !fieldref19
    End If
    If Not IsNull(!fieldref20) Then
        txtRef20.Text = !fieldref20
    End If
    If Not IsNull(!fieldref21) Then
        txtRef21.Text = !fieldref21
    End If
    If Not IsNull(!fieldref22) Then
        txtRef22.Text = !fieldref22
    End If
    If Not IsNull(!fieldref23) Then
        txtRef23.Text = !fieldref23
    End If
    If Not IsNull(!fieldref24) Then
        txtRef24.Text = !fieldref24
    End If
    If Not IsNull(!Fieldref25) Then
        txtRef25.Text = !Fieldref25
    End If
    
    If Not IsNull(!Comments) Then
        txtComments.Text = !Comments
    End If
    


    
    
    If .State = 1 Then .Close
End With
End Sub


Private Sub BeforeAddEdit()

txtSearch.Text = Empty

Grid1.Enabled = True
FrameIx.Enabled = False
bttnAdd.Enabled = True
bttnEdit.Enabled = False
bttnDelete.Enabled = False

bttnSave.Visible = False
bttnChange.Visible = False
bttnCancel.Visible = False
    FromGrid = False


End Sub


Private Sub AfterAdd()

txtSearch.Text = Empty

Call ClearValues

Grid1.Enabled = False
FrameIx.Enabled = True
bttnAdd.Enabled = False
bttnEdit.Enabled = False
bttnDelete.Enabled = False

bttnSave.Visible = True
bttnChange.Visible = False
bttnCancel.Visible = True

End Sub

Private Sub AfterEdit()

txtSearch.Text = Empty

Grid1.Enabled = True
FrameIx.Enabled = False
bttnAdd.Enabled = True
bttnEdit.Enabled = False
bttnDelete.Enabled = False

bttnSave.Visible = False
bttnChange.Visible = True
bttnCancel.Visible = True

End Sub


Private Sub txtSearch_Change()

    
' **************************************

    If FromGrid = True Then Exit Sub
    Dim TemFRows As Long
    Dim TemNowRow As Long
    Dim TemArray As Long
    Dim SearchSuccess As Boolean
    Dim TemLength As Single
    TemFRows = Grid1.Rows
    Grid1.Col = 1
    SearchSuccess = False
    If Len(txtSearch.Text) = 0 Then GoTo MeasureSuccess
    For TemArray = 1 To (TemFRows - 1)
        Grid1.Row = TemArray
        If Len(txtSearch.Text) > Len(Grid1.Text) Then
            GoTo FinishLoop
        Else
            TemLength = Len(txtSearch.Text)
        End If
        If UCase(Left((Grid1.Text), TemLength)) = UCase(txtSearch.Text) Then
            SearchSuccess = True
            Exit For
        Else
            SearchSuccess = False
        End If
FinishLoop:
    Next
    
MeasureSuccess:
    
    If SearchSuccess = True Then
        Grid1.TopRow = TemArray
        Grid1.Row = TemArray
        Grid1.Col = 0
        Grid1.ColSel = (Grid1.Cols - 1)
        bttnEdit.Enabled = True
        bttnAdd.Enabled = False
        Grid1.Col = 2
        TemIxId = Grid1.Text
        Call GetData
        Grid1.Col = 0
        Grid1.ColSel = Grid1.Cols - 1
    Else
        Grid1.TopRow = 1
        Grid1.Row = 0
        Grid1.Col = 0
        Grid1.ColSel = 0
        bttnAdd.Enabled = True
        bttnEdit.Enabled = False
    End If
'**************************************



End Sub

