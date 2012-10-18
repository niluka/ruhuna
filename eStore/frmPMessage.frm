VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPMessage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prescreption Messages"
   ClientHeight    =   6375
   ClientLeft      =   2130
   ClientTop       =   1635
   ClientWidth     =   10920
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   10920
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5280
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   5520
      Width           =   3615
      Begin btButtonEx.ButtonEx bttnAdd 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "&Add"
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
      Begin btButtonEx.ButtonEx bttnEdit 
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "&Edit"
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
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   3960
      TabIndex        =   13
      Top             =   4920
      Width           =   6855
      Begin btButtonEx.ButtonEx bttnSave 
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "&Save"
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
         Left            =   4800
         TabIndex        =   9
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "&Cancel"
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
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "&Save"
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
   Begin VB.Frame Frame2 
      Height          =   5415
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   3615
      Begin MSDataListLib.DataCombo dtcPMessage 
         Height          =   4980
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   8784
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   1
         Text            =   ""
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4815
      Left            =   3960
      TabIndex        =   11
      Top             =   120
      Width           =   6855
      Begin VB.TextBox txtSFont 
         Height          =   360
         Left            =   3240
         TabIndex        =   29
         Top             =   2640
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtSText 
         Height          =   360
         Left            =   1440
         TabIndex        =   28
         Top             =   1200
         Width           =   5295
      End
      Begin VB.TextBox txtTText 
         Height          =   360
         Left            =   1440
         TabIndex        =   27
         Top             =   1680
         Width           =   5295
      End
      Begin VB.TextBox txtEText 
         Height          =   360
         Left            =   1440
         TabIndex        =   26
         Top             =   2160
         Width           =   5295
      End
      Begin VB.TextBox txtTFont 
         Height          =   360
         Left            =   3240
         TabIndex        =   25
         Top             =   3000
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtEFont 
         Height          =   360
         Left            =   3240
         TabIndex        =   24
         Top             =   3360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtSFontSize 
         Height          =   360
         Left            =   3720
         TabIndex        =   23
         Top             =   2640
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtTFontSize 
         Height          =   360
         Left            =   3720
         TabIndex        =   22
         Top             =   3000
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtEFontSize 
         Height          =   360
         Left            =   3720
         TabIndex        =   21
         Top             =   3360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox chkSB 
         Height          =   375
         Left            =   4320
         TabIndex        =   20
         Top             =   2640
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chkSI 
         Height          =   375
         Left            =   4680
         TabIndex        =   19
         Top             =   2640
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chkTB 
         Height          =   375
         Left            =   4320
         TabIndex        =   18
         Top             =   3000
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chkTI 
         Height          =   375
         Left            =   4680
         TabIndex        =   17
         Top             =   3000
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chkEB 
         Height          =   375
         Left            =   4320
         TabIndex        =   16
         Top             =   3360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chkEI 
         Height          =   375
         Left            =   4680
         TabIndex        =   15
         Top             =   3360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtPMessage 
         Height          =   360
         Left            =   1440
         TabIndex        =   4
         Top             =   240
         Width           =   5295
      End
      Begin VB.TextBox txtcomment 
         Height          =   1935
         Left            =   1440
         TabIndex        =   6
         Top             =   2760
         Width           =   5295
      End
      Begin btButtonEx.ButtonEx bttnSFont 
         Height          =   375
         Left            =   1440
         TabIndex        =   30
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Sinhala"
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
      Begin btButtonEx.ButtonEx bttnEFont 
         Height          =   375
         Left            =   5040
         TabIndex        =   31
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "English"
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
      Begin btButtonEx.ButtonEx bttnTFont 
         Height          =   375
         Left            =   3240
         TabIndex        =   32
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Tamil"
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
      Begin VB.Label Label8 
         Caption         =   "Sinhala Text"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "&Font"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Tamil Text"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "English Text"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "&Message"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "C&omments"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2760
         Width           =   1575
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   8880
      TabIndex        =   10
      Top             =   5880
      Width           =   1815
      _ExtentX        =   3201
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
End
Attribute VB_Name = "frmPMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsPMessage As New ADODB.Recordset
    Dim rsViewPMessage As New ADODB.Recordset
    Dim TemPMessageID As Long

Private Sub bttnCancel_Click()
    Call BeforeAddEdit
    Call ClearValues
    dtcPMessage.SetFocus
    SendKeys "{home}+{end}"
End Sub


Private Sub bttnEFont_Click()
    If chkEB.Value = 1 Then
        CommonDialog1.FontBold = True
    Else
        CommonDialog1.FontBold = False
    End If
    If chkEI.Value = 1 Then
        CommonDialog1.FontItalic = True
    Else
        CommonDialog1.FontItalic = False
    End If
    CommonDialog1.FontName = txtEFont.Text
    CommonDialog1.FontSize = Val(txtEFontSize.Text)
    CommonDialog1.Flags = cdlCFBoth
    CommonDialog1.ShowFont
    txtEFont.Text = CommonDialog1.FontName
    txtEFontSize.Text = CommonDialog1.FontSize
    If CommonDialog1.FontBold = True Then
        chkEB.Value = 1
    Else
        chkEB.Value = 0
    End If
    If CommonDialog1.FontItalic = True Then
        chkEI.Value = 1
    Else
        chkEI.Value = 0
    End If
    txtEText.FontName = txtEFont.Text
    txtEText.FontSize = Val(txtEFontSize.Text)
    If chkEB.Value = 1 Then
        txtEText.FontBold = True
    Else
        txtEText.FontBold = False
    End If
    If chkEI.Value = 1 Then
        txtEText.FontItalic = True
    Else
        txtEText.FontItalic = False
    End If
End Sub

Private Sub bttnSFont_Click()
    On Error Resume Next
    If chkSB.Value = 1 Then
        CommonDialog1.FontBold = True
    Else
        CommonDialog1.FontBold = False
    End If
    If chkSI.Value = 1 Then
        CommonDialog1.FontItalic = True
    Else
        CommonDialog1.FontItalic = False
    End If
    CommonDialog1.FontName = txtSFont.Text
    CommonDialog1.FontSize = Val(txtSFontSize.Text)
    CommonDialog1.Flags = cdlCFBoth
    CommonDialog1.ShowFont
    txtSFont.Text = CommonDialog1.FontName
    txtSFontSize.Text = CommonDialog1.FontSize
    If CommonDialog1.FontBold = True Then
        chkSB.Value = 1
    Else
        chkSB.Value = 0
    End If
    If CommonDialog1.FontItalic = True Then
        chkSI.Value = 1
    Else
        chkSI.Value = 0
    End If
    txtSText.FontName = txtSFont.Text
    txtSText.FontSize = Val(txtSFontSize.Text)
    If chkSB.Value = 1 Then
        txtSText.FontBold = True
    Else
        txtSText.FontBold = False
    End If
    If chkSI.Value = 1 Then
        txtSText.FontItalic = True
    Else
        txtSText.FontItalic = False
    End If
End Sub

Private Sub bttnTFont_Click()
    If chkTB.Value = 1 Then
        CommonDialog1.FontBold = True
    Else
        CommonDialog1.FontBold = False
    End If
    If chkTI.Value = 1 Then
        CommonDialog1.FontItalic = True
    Else
        CommonDialog1.FontItalic = False
    End If
    CommonDialog1.FontName = txtTFont.Text
    CommonDialog1.FontSize = Val(txtTFontSize.Text)
    CommonDialog1.Flags = cdlCFBoth
    CommonDialog1.ShowFont
    txtTFont.Text = CommonDialog1.FontName
    txtTFontSize.Text = CommonDialog1.FontSize
    If CommonDialog1.FontBold = True Then
        chkTB.Value = 1
    Else
        chkTB.Value = 0
    End If
    If CommonDialog1.FontItalic = True Then
        chkTI.Value = 1
    Else
        chkTI.Value = 0
    End If
    txtTText.FontName = txtTFont.Text
    txtTText.FontSize = Val(txtTFontSize.Text)
    If chkTB.Value = 1 Then
        txtTText.FontBold = True
    Else
        txtTText.FontBold = False
    End If
    If chkTI.Value = 1 Then
        txtTText.FontItalic = True
    Else
        txtTText.FontItalic = False
    End If
End Sub

Private Sub dtcPMessage_Click(Area As Integer)
    If IsNumeric(dtcPMessage.BoundText) = False Then Exit Sub
    Call DisplaySelected
End Sub

Private Sub Form_Load()
    FillGenaricCombo
    BeforeAddEdit
    ClearValues
End Sub

Private Sub bttnAdd_Click()
    ClearValues
    AfterAdd
    txtPMessage.Text = dtcPMessage.Text
    txtPMessage.SetFocus
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnEdit_Click()
    AfterEdit
    txtPMessage.SetFocus
    SendKeys "{Home}+{end}"
End Sub

Private Sub bttnChange_Click()
    Dim TemResponce As Integer
    If Trim(txtPMessage.Text) = "" Then NoName: Exit Sub
    With rsPMessage
    On Error Resume Next
        If .State = 1 Then .Close
        .Open "Select* From tblPMessage Where (PMessageID = " & TemPMessageID & ")", cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !PMessage = Trim(txtPMessage.Text)
            !sFont = txtSFont.Text
            !Sfontsize = Val(txtSFontSize.Text)
            If chkSB.Value = 1 Then
                !SBold = True
            Else
                !SBold = False
            End If
            If chkSI.Value = 1 Then
                !SItalic = True
            Else
                !SItalic = False
            End If
            !Stext = txtSText.Text
            !tFont = txtTFont.Text
            !tfontsize = Val(txtTFontSize.Text)
            If chkTB.Value = 1 Then
                !tBold = True
            Else
                !tBold = False
            End If
            If chkTI.Value = 1 Then
                !tItalic = True
            Else
                !tItalic = False
            End If
            !TText = txtTText.Text
            
            !eFont = txtEFont.Text
            !efontsize = Val(txtEFontSize.Text)
            If chkEB.Value = 1 Then
                !eBold = True
            Else
                !eBold = False
            End If
            If chkEI.Value = 1 Then
                !eItalic = True
            Else
                !eItalic = False
            End If
            !EText = txtEText.Text
            !Comments = txtcomment.Text
            .Update
        End If
        If .State = 1 Then .Close
        FillGenaricCombo
        BeforeAddEdit
        ClearValues
        Exit Sub
   
ErrorHandler:
        TemResponce = MsgBox(Err.Number & vbNewLine & Err.Description & Me.Caption, vbCritical + vbOKOnly, "Save Error")
        If .State = 1 Then .CancelUpdate
        ClearValues
        BeforeAddEdit
        If .State = 1 Then .Close
    End With
End Sub

Private Sub bttnSave_Click()
    Dim TemResponce As Integer
    If Trim(txtPMessage.Text) = "" Then NoName: Exit Sub
    With rsPMessage
    On Error Resume Next
        If .State = 1 Then .Close
        .Open "Select* From tblPMessage", cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
            !PMessage = Trim(txtPMessage.Text)
            !sFont = txtSFont.Text
            !Sfontsize = Val(txtSFontSize.Text)
            If chkSB.Value = 1 Then
                !SBold = True
            Else
                !SBold = False
            End If
            If chkSI.Value = 1 Then
                !SItalic = True
            Else
                !SItalic = False
            End If
            !Stext = txtSText.Text
            !tFont = txtTFont.Text
            !tfontsize = Val(txtTFontSize.Text)
            If chkTB.Value = 1 Then
                !tBold = True
            Else
                !tBold = False
            End If
            If chkTI.Value = 1 Then
                !tItalic = True
            Else
                !tItalic = False
            End If
            !TText = txtTText.Text
            
            !eFont = txtEFont.Text
            !efontsize = Val(txtEFontSize.Text)
            If chkEB.Value = 1 Then
                !eBold = True
            Else
                !eBold = False
            End If
            If chkEI.Value = 1 Then
                !eItalic = True
            Else
                !eItalic = False
            End If
            !EText = txtEText.Text
            !Comments = txtcomment.Text
            .Update
        If .State = 1 Then .Close
        FillGenaricCombo
        BeforeAddEdit
        ClearValues
        Exit Sub
ErrorHandler:
        TemResponce = MsgBox(Err.Number & vbNewLine & Err.Description & Me.Caption, vbCritical + vbOKOnly, "Save Error")
        .CancelUpdate
        ClearValues
        BeforeAddEdit
        If .State = 1 Then .Close
    End With
End Sub

Private Sub NoName()
    Dim TemResponce As Integer
    TemResponce = MsgBox("You have not entered a Generic Name to save", vbCritical, "No Name")
    txtPMessage.SetFocus
End Sub


Private Sub FillGenaricCombo()
    With rsViewPMessage
        If .State = 1 Then .Close
        .Open "Select* From tblPMessage Order By PMessage", cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount = 0 Then Exit Sub
        Set dtcPMessage.RowSource = rsViewPMessage
        dtcPMessage.ListField = "PMessage"
        dtcPMessage.BoundColumn = "PMessageID"
    End With
End Sub

Private Sub AfterAdd()
    bttnSave.Visible = True
    bttnChange.Visible = False
    bttnCancel.Visible = True
    bttnAdd.Enabled = False
    bttnEdit.Enabled = False
    bttnSave.Enabled = True
    bttnCancel.Enabled = True
    bttnChange.Enabled = False
    Frame2.Enabled = False
    Frame1.Enabled = True
End Sub

Private Sub AfterEdit()
    bttnSave.Visible = False
    bttnChange.Visible = True
    bttnCancel.Visible = True
    bttnAdd.Enabled = False
    bttnEdit.Enabled = False
    bttnSave.Enabled = False
    bttnCancel.Enabled = True
    bttnChange.Enabled = True
    Frame2.Enabled = False
    Frame1.Enabled = True
End Sub



Private Sub BeforeAddEdit()
    bttnAdd.Visible = True
    bttnEdit.Visible = True
    bttnSave.Visible = False
    bttnCancel.Visible = False
    bttnChange.Visible = False
    bttnAdd.Enabled = True
    bttnEdit.Enabled = True
    bttnSave.Enabled = False
    bttnCancel.Enabled = False
    Frame2.Enabled = True
    Frame1.Enabled = False
End Sub

Private Sub ClearValues()
    txtPMessage.Text = Empty
    txtcomment.Text = Empty
    txtSFont.Text = Empty
    txtSFontSize.Text = Empty
    txtTFont.Text = Empty
    txtTFontSize.Text = Empty
    txtEFont.Text = Empty
    txtEFontSize.Text = Empty
    chkEB.Value = 0
    chkEI.Value = 0
    chkSB.Value = 0
    chkSI.Value = 0
    chkTB.Value = 0
    chkTI.Value = 0
    txtSText.Text = Empty
    txtTText.Text = Empty
    txtEText.Text = Empty
    TemPMessageID = Empty
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsPMessage.State = 1 Then rsPMessage.Close: Set rsPMessage = Nothing
    If rsViewPMessage.State = 1 Then rsViewPMessage.Close: Set rsViewPMessage = Nothing
End Sub

Private Sub DisplaySelected()
    If Not IsNumeric(dtcPMessage.BoundText) Then Exit Sub
    With rsPMessage
        If .State = 1 Then .Close
        .Open "Select* From tblPMessage Where (PMessageID = " & dtcPMessage.BoundText & ")", cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount = 0 Then Exit Sub
        Call ClearValues
        txtPMessage.Text = !PMessage
        If Not IsNull(!Comments) Then txtcomment.Text = !Comments
        TemPMessageID = !PMessageID
        If Not IsNull(!Comments) Then txtcomment.Text = !Comments
        If Not IsNull(!sFont) Then txtSFont.Text = !sFont
        If Not IsNull(!Sfontsize) Then txtSFontSize.Text = !Sfontsize
        If !SBold = True Then chkSB.Value = 1
        If !SItalic = True Then chkSI.Value = 1
        If Not IsNull(!tFont) Then txtTFont.Text = !tFont
        If Not IsNull(!tfontsize) Then txtTFontSize.Text = !tfontsize
        If !tBold = True Then chkTB.Value = 1
        If !tItalic = True Then chkTI.Value = 1
        If Not IsNull(!eFont) Then txtEFont.Text = !eFont
        If Not IsNull(!efontsize) Then txtEFontSize.Text = !efontsize
        If !eBold = True Then chkEB.Value = 1
        If !eItalic = True Then chkEI.Value = 1
        If Not IsNull(!Stext) Then txtSText.Text = !Stext
        If Not IsNull(!TText) Then txtTText.Text = !TText
        If Not IsNull(!EText) Then txtEText.Text = !EText
    End With
    On Error Resume Next
    If chkEB.Value = 1 Then
        txtEText.FontBold = True
    Else
        txtEText.FontBold = False
    End If
    If chkEI.Value = 1 Then
        txtEText.FontItalic = True
    Else
        txtEText.FontItalic = False
    End If
    txtSText.FontName = txtSFont.Text
    txtSText.FontSize = Val(txtSFontSize.Text)
    If chkSB.Value = 1 Then
        txtSText.FontBold = True
    Else
        txtSText.FontBold = False
    End If
    If chkSI.Value = 1 Then
        txtSText.FontItalic = True
    Else
        txtSText.FontItalic = False
    End If
    txtTText.FontName = txtTFont.Text
    txtTText.FontSize = Val(txtTFontSize.Text)
    If chkTB.Value = 1 Then
        txtTText.FontBold = True
    Else
        txtTText.FontBold = False
    End If
    If chkTI.Value = 1 Then
        txtTText.FontItalic = True
    Else
        txtTText.FontItalic = False
    End If
End Sub

