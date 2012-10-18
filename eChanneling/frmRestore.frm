VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmRestore 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Restore"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6405
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRestore.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   6405
   Begin VB.TextBox txtPath 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   5895
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2535
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   4471
      _Version        =   393216
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
      Left            =   4800
      TabIndex        =   0
      Top             =   4320
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
   Begin btButtonEx.ButtonEx bttnRestore 
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   4320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Restore"
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
   Begin btButtonEx.ButtonEx bttnSelectPath 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Select Path"
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
      Caption         =   "Select the directory from which to restore"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmRestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim FSys As New Scripting.FileSystemObject
    Dim TemResponce  As Integer
    Dim ThisFolder As Folder
    Dim AllFiles As Files
    Dim ThisFile As File
    Dim TemString1 As String
    Dim TemString2 As String
    Dim TemString3 As String
    Dim TemString4 As String
    Dim TemDate As Date
    Private Const BIF_RETURNONLYFSDIRS = 1
    Private Const BIF_DONTGOBELOWDOMAIN = 2
    Private Const MAX_PATH = 260
    Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
    Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
    Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
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
    bttnSelectPath.BackColor = BttnBackColour
    bttnSelectPath.ForeColor = BttnForeColour
    bttnClose.BackColor = BttnBackColour
    bttnClose.ForeColor = BttnForeColour
    bttnRestore.BackColor = BttnBackColour
    bttnRestore.ForeColor = BttnForeColour
    frmRestore.BackColor = FrmBackColour
    frmRestore.ForeColor = FrmForeColour
    Grid1.BackColor = GridBackColor
    Grid1.ForeColor = GridForeColor
    Grid1.BackColorBkg = GridBackColorBkg
    Grid1.BackColorFixed = GridBackColorFixed
    Grid1.BackColorSel = GridBackColorSel
    Grid1.ForeColor = GridForeColor
    Grid1.ForeColorFixed = GridForeColorFixed
    Grid1.ForeColorSel = GridForeColorSel
    Label1.BackColor = LblBackColour
    Label1.ForeColor = LblForeColour
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnRestore_Click()
    Grid1.Col = 0
    TemResponce = MsgBox("Are you sure you want replace the current database with a database backed up on " & Grid1.Text, vbInformation + vbYesNo, "Restore?")
    If TemResponce = vbNo Then Exit Sub
    On Error GoTo ErrorHandler
    Me.MousePointer = vbHourglass
    DoEvents
    FSys.CopyFile App.Path & "/hospital.mdb", txtPath.Text & "\BUeHos" & Format(Date, "ddmmyy") & ".mdb", True
    FSys.CopyFile Grid1.Text, DatabasePath, True
    'FSys.CopyFile Grid1.Text, App.Path & "/hospital.mdb", True
    
    TemResponce = MsgBox("Restore Successful", vbInformation, "Success")
    Me.MousePointer = vbDefault
    Exit Sub
ErrorHandler:
    TemResponce = MsgBox("An unknown error occured. Please contact lakmedipro with following details." & vbNewLine & App.EXEName & vbNewLine & Me.Caption & vbNewLine & Err.Number & vbNewLine & Err.Description, vbInformation, "Error")
    Exit Sub
End Sub

Private Sub ListDatabases()
    Dim NowROw As Long
    Grid1.Clear
    Grid1.Rows = 1
    Grid1.Cols = 2
    Grid1.ColWidth(0) = 1
    Grid1.ColWidth(1) = Grid1.Width - 101
    Grid1.Row = 0
    Grid1.Col = 1
    Grid1.Text = "Backed Up Date"
    If FSys.FolderExists(txtPath.Text) = True Then
        
        Set ThisFolder = FSys.GetFolder(txtPath.Text)
        Set AllFiles = ThisFolder.Files
        
        NowROw = 0
        For Each ThisFile In AllFiles
            If UCase(Left(ThisFile.Name, 6)) = UCase("buehos") Then
                NowROw = NowROw + 1
                Grid1.Rows = NowROw + 1
                Grid1.Row = NowROw
                TemString1 = Mid(ThisFile.Name, 7, 2)
                TemString2 = Mid(ThisFile.Name, 9, 2)
                TemString3 = Mid(ThisFile.Name, 11, 2)
                TemDate = DateSerial(TemString3, TemString2, TemString1)
                Grid1.Col = 0
                Grid1.Text = ThisFile.Path
                Grid1.Col = 1
                Grid1.CellAlignment = 4
                Grid1.Text = Format(TemDate, DefaultLongDate)
            End If
        Next
    End If
    bttnRestore.Enabled = False
End Sub

Private Sub bttnSelectPath_Click()
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
            Call ListDatabases
         End If
End Sub

Private Sub Form_Load()
    Call Setcolours
    bttnRestore.Enabled = False
    txtPath.Text = BackUpPath
    Call ListDatabases
End Sub

Private Sub Grid1_Click()
    If Grid1.Row < 1 Then
        bttnRestore.Enabled = False
        Exit Sub
    Else
        bttnRestore.Enabled = True
    End If
End Sub
