VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmEditRoomDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Room Details"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11040
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
   ScaleHeight     =   6870
   ScaleWidth      =   11040
   Begin VB.TextBox txtID 
      Height          =   375
      Left            =   4800
      TabIndex        =   14
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox txtDetails 
      Height          =   1815
      Left            =   6240
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   240
      Width           =   4695
   End
   Begin btButtonEx.ButtonEx btnSave 
      Height          =   495
      Left            =   4800
      TabIndex        =   11
      Top             =   1560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
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
   Begin MSFlexGridLib.MSFlexGrid gridRoom 
      Height          =   3975
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   7011
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker dtpfromTime 
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   1200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   67108866
      CurrentDate     =   40182
   End
   Begin MSComCtl2.DTPicker dtpFromDate 
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   1200
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   67108867
      CurrentDate     =   40182
   End
   Begin MSDataListLib.DataCombo cmbBHT 
      Height          =   360
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbRoom 
      Height          =   360
      Left            =   960
      TabIndex        =   3
      Top             =   720
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpToDate 
      Height          =   375
      Left            =   960
      TabIndex        =   8
      Top             =   1680
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   67108867
      CurrentDate     =   40182
   End
   Begin MSComCtl2.DTPicker dtpToTime 
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   1680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   67108866
      CurrentDate     =   40182
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   9720
      TabIndex        =   12
      Top             =   6240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
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
   Begin VB.Label Label4 
      Caption         =   "To"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "From"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Room"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "BHT"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmEditRoomDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsBHT As New ADODB.Recordset
    Dim temSql As String
    Dim MyBHT As New clsBHT

Private Sub FillCombos()
    With rsBHT
        If .State = 1 Then .Close
        temSql = "Select * from tblBHT where IsBHT = 1 And Discharge = 0 order by BHT"
        'temSql = "Select * from tblBHT where IsBHT = 1 Order by BHT"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbBHT
        Set .RowSource = rsBHT
        .ListField = "BHT"
        .BoundColumn = "BHTID"
    End With
    Dim CR As New clsFillCombos
    CR.FillAnyCombo cmbRoom, "Room", False
End Sub

Private Sub btnSave_Click()
    If IsNumeric(cmbBHT.BoundText) = False Then
        MsgBox "Please select the BHT"
        cmbBHT.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(cmbRoom.BoundText) = False Then
        MsgBox "Please select the room"
        cmbRoom.SetFocus
        Exit Sub
    End If
    
    If gridRoom.Rows <= 1 Then
        MsgBox "No Rooms Listed"
        cmbBHT.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(txtID.Text) = False Then
        MsgBox "Error in Update"
        cmbBHT.SetFocus
        Exit Sub
    End If
    
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblRoomPatient where RoomPatientID = " & Val(txtID.Text)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !FromDate = dtpFromDate.Value
            !FromTime = Format(dtpFromDate.Value, "dd MMMM yyyy") & " " & dtpfromTime.Value
            !RoomID = Val(cmbRoom.BoundText)
            
            If dtpToDate.Visible = True Then !ToDate = dtpToDate.Value
            If dtpToTime.Visible = True Then !ToTime = Format(dtpToDate.Value, "dd MMMM yyyy") & " " & dtpToTime.Value
            
            .Update
        End If
        .Close
    End With
    Call FormatGrid
    Call FillGrid
    
End Sub

Private Sub cmbBHT_Change()
    MyBHT.BHTID = Val(cmbBHT.BoundText)
    Call ClearValues
    Call DisplayDetails
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub ClearValues()
    txtDetails.Text = Empty
    txtID.Text = Empty
End Sub

Private Sub DisplayDetails()
    'On Error Resume Next
    Dim temText As String
    Dim r As Long
    temText = "Patient Name : " & MyBHT.FirstName & vbNewLine
    temText = temText & "Guardian : " & MyBHT.GuardianName & vbNewLine
    temText = temText & "Address : " & MyBHT.PtAddress & vbNewLine
    temText = temText & "BHT : " & MyBHT.BHT & vbNewLine
    temText = temText & "Age : " & MyBHT.AgeInWords & vbNewLine
    temText = temText & "Admitted : " & Format(MyBHT.DOA, "dd MMMM yyyy") & " at " & Format(MyBHT.TOA, "HH:MM AMPM") & vbNewLine
    If MyBHT.Discharge = True Then
        temText = temText & "Discharged :" & Format(MyBHT.DOD, "dd MMMM yyyy") & " at " & Format(MyBHT.TOD, "HH:MM AMPM") & vbNewLine
    Else
        temText = temText & "Not yet discharged" & vbNewLine
    End If
    temText = temText & "Payment Method : " & MyBHT.PaymentMethod
    If MyBHT.HealthSchemeSupplier <> "" Then
        temText = temText & " (" & MyBHT.HealthSchemeSupplier & ")" & vbNewLine
    Else
        temText = temText & vbNewLine
    End If
    If MyBHT.Comments <> "" Then
        temText = temText & MyBHT.Comments & vbNewLine
    End If
    
    txtDetails.Text = temText
    
End Sub

Private Sub Form_Load()
    Call FormatGrid
    Call GetSettings
    Call FillCombos
End Sub

Private Sub GetSettings(): On Error Resume Next
    GetCommonSettings Me
End Sub


Private Sub SaveSettings()
    SaveCommonSettings Me
End Sub

Private Sub FormatGrid()
    With gridRoom
        .Clear
        
        .Cols = 8
        .Rows = 1
        
        .Row = 0
        
        .Col = 0
        .Text = "ID"
        
        .Col = 1
        .Text = "No"
        
        .Col = 2
        .Text = "Room"
        
        .Col = 3
        .Text = "From Date"
        
        .Col = 4
        .Text = "From Time"
        
        .Col = 5
        .Text = "To Date"
        
        .Col = 6
        .Text = "To Time"
        
        .Col = 7
        .Text = "Room ID"
    End With
End Sub

Private Sub FillGrid()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT     TOP 100 PERCENT dbo.tblRoomPatient.RoomPatientID, dbo.tblRoom.RoomID, dbo.tblRoom.Room, dbo.tblRoomPatient.FromDate, dbo.tblRoomPatient.FromTime, dbo.tblRoomPatient.ToDate , dbo.tblRoomPatient.ToTime " & _
                    "FROM         dbo.tblRoomPatient LEFT OUTER JOIN " & _
                      "dbo.tblRoom ON dbo.tblRoomPatient.RoomID = dbo.tblRoom.RoomID " & _
                        "Where (dbo.tblRoomPatient.BHTID = " & Val(cmbBHT.BoundText) & ") " & _
                        "ORDER BY dbo.tblRoomPatient.FromTime "
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        While .EOF = False
            gridRoom.Rows = gridRoom.Rows + 1
            gridRoom.Row = gridRoom.Rows - 1
            
            gridRoom.Col = 0
            gridRoom.Text = !RoomPatientID
            
            gridRoom.Col = 1
            gridRoom.Text = gridRoom.Row
            
            gridRoom.Col = 2
            gridRoom.Text = !Room
            
            gridRoom.Col = 3
            gridRoom.Text = Format(!FromDate, "dd MMMM yyyy")
            
            gridRoom.Col = 4
            gridRoom.Text = Format(!FromTime, "hh:mm:ss AMPM")
            
            gridRoom.Col = 5
            gridRoom.Text = Format(!ToDate, "dd MMMM yyyy")
            
            gridRoom.Col = 6
            gridRoom.Text = Format(!ToTime, "hh:mm:ss AMPM")
            
            gridRoom.Col = 7
            gridRoom.Text = !RoomID
            
            
            .MoveNext
        Wend
    End With
    gridRoom.ColWidth(0) = 0
    gridRoom.ColWidth(1) = 0
    gridRoom.ColWidth(7) = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveSettings
End Sub

Private Sub gridRoom_DblClick()
    txtID.Text = Empty
    With gridRoom
        If IsNumeric(.TextMatrix(.Row, 0)) = False Then Exit Sub
        cmbRoom.BoundText = Val(.TextMatrix(.Row, 7))
        dtpFromDate.Value = CDate(.TextMatrix(.Row, 3))
        dtpfromTime.Value = CDate(.TextMatrix(.Row, 4))
        If IsDate(.TextMatrix(.Row, 5)) = True Then
            dtpToDate.Visible = True
            dtpToDate.Value = CDate(.TextMatrix(.Row, 5))
        Else
            dtpToDate.Visible = False
        End If
        If IsDate(.TextMatrix(.Row, 6)) = True Then
            dtpToTime.Visible = True
            dtpToTime.Value = CDate(.TextMatrix(.Row, 6))
        Else
            dtpToTime.Visible = False
        End If
        txtID.Text = Val(.TextMatrix(.Row, 0))
    End With
End Sub
