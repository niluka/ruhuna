VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmBackUpOnExit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5820
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBackUpOnExit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   5820
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPath 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   720
      Width           =   5535
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "C&lose"
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
   Begin btButtonEx.ButtonEx bttnBackUP 
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Back Up"
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
   Begin btButtonEx.ButtonEx bttnSelect 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Select Directory"
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
      Caption         =   "Select the directory to store the backup files"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   3855
   End
End
Attribute VB_Name = "frmBackUpOnExit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim FSys As New Scripting.FileSystemObject
    
    Private Const BIF_RETURNONLYFSDIRS = 1
    Private Const BIF_DONTGOBELOWDOMAIN = 2
    Private Const MAX_PATH = 260
    
    Private Declare Function SHBrowseForFolder Lib "shell32" _
                                      (lpbi As BrowseInfo) As Long
    
    Private Declare Function SHGetPathFromIDList Lib "shell32" _
                                      (ByVal pidList As Long, _
                                      ByVal lpBuffer As String) As Long
    
    Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
                                      (ByVal lpString1 As String, ByVal _
                                      lpString2 As String) As Long
    
    Private Type BrowseInfo
       hWndOwner      As Long
       pIDLRoot       As Long
       pszDisplayName As Long
       lpszTitle      As Long
       ulFlags        As Long
       lpfnCallback   As Long
       lparam         As Long
       iImage         As Long
    End Type

Private Sub Setcolours()
    bttnSelect.BackColor = BttnBackColour
    bttnSelect.ForeColor = BttnForeColour
    bttnClose.BackColor = BttnBackColour
    bttnClose.ForeColor = BttnForeColour
    bttnBackUP.BackColor = BttnBackColour
    bttnBackUP.ForeColor = BttnForeColour
    Me.BackColor = FrmBackColour
    Me.ForeColor = FrmForeColour
    Label1.BackColor = LblBackColour
    Label1.ForeColor = LblForeColour
End Sub

Private Sub bttnBackUP_Click()
    Dim TemResponce  As Integer
    On Error GoTo ErrorHandler
    If FSys.FolderExists(txtPath.Text) = True Then
        Me.MousePointer = vbHourglass
        FSys.CopyFile App.Path & "\hospital.mdb", txtPath.Text & "\BUeHos" & Format(Date, "ddmmyy") & ".mdb", True
        TemResponce = MsgBox("Backup Successful", vbInformation, "Success")
        Me.MousePointer = vbDefault
    Else
        TemResponce = MsgBox("The path you selected is not valid. Please select a valid path", vbCritical, "Path not valid")
        Exit Sub
    End If
    Exit Sub
ErrorHandler:
        TemResponce = MsgBox("An unknown error occured. Please contact lakmedipro with following details." & vbNewLine & App.EXEName & vbNewLine & Me.Caption & vbNewLine & Err.Number & vbNewLine & Err.Description, vbInformation, "Error")
        Me.MousePointer = vbDefault
        Exit Sub
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnSelect_Click()
         Dim lpIDList As Long
         Dim sBuffer As String
         Dim szTitle As String
         Dim tBrowseInfo As BrowseInfo
         szTitle = "Select Backup Directory"
         With tBrowseInfo
            .hWndOwner = Me.hwnd
            .lpszTitle = lstrcat(szTitle, "")
            .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
         End With
         lpIDList = SHBrowseForFolder(tBrowseInfo)
         If (lpIDList) Then
            sBuffer = Space(MAX_PATH)
            SHGetPathFromIDList lpIDList, sBuffer
            sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
            txtPath.Text = sBuffer
         End If
End Sub

Private Sub Form_Load()
    Call Setcolours
    If FSys.FolderExists(BackUpPath) Then
        txtPath.Text = BackUpPath
    Else
        txtPath.Text = App.Path
    End If
End Sub

