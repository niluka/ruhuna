VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmReportsItemSale 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Issue to All Agents - Time Wise -  Items "
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6915
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
   ScaleHeight     =   8040
   ScaleWidth      =   6915
   Begin TabDlg.SSTab SSTab1 
      Height          =   7215
      Left            =   120
      TabIndex        =   48
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   12726
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Analysis Details"
      TabPicture(0)   =   "frmReportsItemSale.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "dtpFrom"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "dtpTO"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Graph Details"
      TabPicture(1)   =   "frmReportsItemSale.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "Frame5"
      Tab(1).Control(2)=   "Frame6"
      Tab(1).Control(3)=   "Frame7"
      Tab(1).Control(4)=   "Frame8"
      Tab(1).Control(5)=   "Frame9"
      Tab(1).Control(6)=   "Frame10"
      Tab(1).Control(7)=   "Frame11"
      Tab(1).ControlCount=   8
      Begin VB.Frame Frame11 
         Height          =   1095
         Left            =   -72120
         TabIndex        =   45
         Top             =   4320
         Width           =   3615
         Begin VB.OptionButton optDisplayLegend 
            Caption         =   "Display Ligend"
            Height          =   375
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Value           =   -1  'True
            Width           =   2175
         End
         Begin VB.OptionButton optNoLegend 
            Caption         =   "Do not display Legend"
            Height          =   375
            Left            =   120
            TabIndex        =   47
            Top             =   600
            Width           =   2175
         End
      End
      Begin VB.Frame Frame10 
         Height          =   1095
         Left            =   -72120
         TabIndex        =   42
         Top             =   3240
         Width           =   3615
         Begin VB.OptionButton optYAxis 
            Caption         =   "Plot By Rows"
            Height          =   375
            Left            =   120
            TabIndex        =   44
            Top             =   600
            Width           =   3255
         End
         Begin VB.OptionButton optXAxis 
            Caption         =   "Plot By Colmns"
            Height          =   375
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Value           =   -1  'True
            Width           =   2895
         End
      End
      Begin VB.Frame Frame9 
         Height          =   1695
         Left            =   -72120
         TabIndex        =   38
         Top             =   1560
         Width           =   3615
         Begin VB.ComboBox cmbChartType 
            Height          =   360
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   1080
            Width           =   3135
         End
         Begin VB.OptionButton optOtherCharts 
            Caption         =   "Other Chart Types"
            Height          =   375
            Left            =   120
            TabIndex        =   40
            Top             =   600
            Width           =   3375
         End
         Begin VB.OptionButton optStandardChart 
            Caption         =   "Standared Chart type"
            Height          =   375
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Value           =   -1  'True
            Width           =   3375
         End
      End
      Begin VB.Frame Frame8 
         Height          =   1095
         Left            =   -72120
         TabIndex        =   35
         Top             =   360
         Width           =   3615
         Begin VB.OptionButton optDisplayZero 
            Caption         =   "Display Zero Values"
            Height          =   375
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Value           =   -1  'True
            Width           =   2415
         End
         Begin VB.OptionButton optNoZero 
            Caption         =   "Don't Display Zero Values"
            Height          =   375
            Left            =   120
            TabIndex        =   37
            Top             =   600
            Width           =   3375
         End
      End
      Begin VB.Frame Frame7 
         Height          =   1095
         Left            =   -74880
         TabIndex        =   32
         Top             =   4320
         Width           =   2535
         Begin VB.OptionButton optDisplayValues 
            Caption         =   "Display values"
            Height          =   375
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Value           =   -1  'True
            Width           =   2175
         End
         Begin VB.OptionButton optDoNotDisplayValues 
            Caption         =   "Do not display values"
            Height          =   375
            Left            =   120
            TabIndex        =   34
            Top             =   600
            Width           =   2175
         End
      End
      Begin VB.Frame Frame6 
         Height          =   1095
         Left            =   -74880
         TabIndex        =   29
         Top             =   3120
         Width           =   2535
         Begin VB.OptionButton optNoTitle 
            Caption         =   "No title"
            Height          =   375
            Left            =   120
            TabIndex        =   31
            Top             =   600
            Width           =   2175
         End
         Begin VB.OptionButton optDisplayTitle 
            Caption         =   "Display title"
            Height          =   375
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Value           =   -1  'True
            Width           =   2175
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1095
         Left            =   -74880
         TabIndex        =   26
         Top             =   1920
         Width           =   2535
         Begin VB.OptionButton opt2D 
            Caption         =   "2 D"
            Height          =   375
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton opt3D 
            Caption         =   "3 D"
            Height          =   375
            Left            =   120
            TabIndex        =   28
            Top             =   600
            Width           =   1455
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1455
         Left            =   -74880
         TabIndex        =   22
         Top             =   360
         Width           =   2535
         Begin VB.OptionButton optPie 
            Caption         =   "Pie Chart"
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Top             =   960
            Width           =   1455
         End
         Begin VB.OptionButton optLine 
            Caption         =   "Line Chart"
            Height          =   375
            Left            =   120
            TabIndex        =   24
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton optBar 
            Caption         =   "Bar Chart"
            Height          =   375
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1095
         Left            =   1560
         TabIndex        =   5
         Top             =   1320
         Width           =   3855
         Begin VB.OptionButton optMonthly 
            Caption         =   "Monthly"
            Height          =   375
            Left            =   2160
            TabIndex        =   7
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optYearly 
            Caption         =   "Yearly"
            Height          =   375
            Left            =   2160
            TabIndex        =   9
            Top             =   600
            Width           =   1335
         End
         Begin VB.OptionButton optWeekly 
            Caption         =   "Weekly"
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Top             =   600
            Width           =   1215
         End
         Begin VB.OptionButton optDaily 
            Caption         =   "Daily"
            Height          =   375
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   1560
         TabIndex        =   11
         Top             =   2400
         Width           =   3855
         Begin VB.OptionButton optByQty 
            Caption         =   "Quentity"
            Height          =   375
            Left            =   2160
            TabIndex        =   13
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optByVal 
            Caption         =   "Value"
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3975
         Left            =   1560
         TabIndex        =   15
         Top             =   3120
         Width           =   3855
         Begin MSDataListLib.DataCombo cmbGeneric 
            Height          =   360
            Left            =   120
            TabIndex        =   18
            Top             =   600
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   635
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            Text            =   ""
         End
         Begin VB.ListBox lstItems 
            Height          =   2220
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   19
            Top             =   1680
            Width           =   3615
         End
         Begin VB.ListBox lstItemIDs 
            Height          =   600
            Left            =   3240
            Style           =   1  'Checkbox
            TabIndex        =   49
            Top             =   1680
            Width           =   495
         End
         Begin VB.OptionButton optSelectdeItem 
            Caption         =   "Selected Item"
            Height          =   255
            Left            =   1800
            TabIndex        =   17
            Top             =   240
            Width           =   1935
         End
         Begin VB.OptionButton optAllItems 
            Caption         =   "All Items"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin MSDataListLib.DataCombo cmbCategory 
            Height          =   360
            Left            =   120
            TabIndex        =   50
            Top             =   1080
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   635
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            Text            =   ""
         End
      End
      Begin MSComCtl2.DTPicker dtpTO 
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   960
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dddd, dd MMMM yyyy"
         Format          =   20709379
         CurrentDate     =   39576
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   480
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dddd, dd MMMM yyyy"
         Format          =   20709379
         CurrentDate     =   39576
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Interval"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Calculate"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Items"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   3240
         Width           =   1215
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   5520
      TabIndex        =   21
      Top             =   7440
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
   Begin btButtonEx.ButtonEx bttnCreate 
      Height          =   495
      Left            =   4200
      TabIndex        =   20
      Top             =   7440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Graph"
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
Attribute VB_Name = "frmReportsItemSale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim myworkbook As Excel.Workbook
    Dim myworksheet As Excel.Worksheet
    Dim mychart As Excel.Chart
    Dim TemPath As String
    Dim FSys As New Scripting.FileSystemObject
    Dim rsViewDriver As New ADODB.Recordset
    Dim rsItem As New ADODB.Recordset
    Dim rsShape As New ADODB.Recordset
    Dim rsGeneric As New ADODB.Recordset
    Dim rsCategory As New ADODB.Recordset
    
    
    Dim temTopic As String
    Dim temSubTopic As String
    
    Dim rsTem As New ADODB.Recordset
        
    Dim rsTemReport As New ADODB.Recordset

    Dim temSql As String
    Dim temSelect As String
    Dim temWhere As String
    Dim temFrom As String
    Dim temOrderBY As String
    Dim temGroupBy As String
    
    Dim rsProduction As New ADODB.Recordset
    Dim rsViewItem As New ADODB.Recordset


Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnCreate_Click()
    Dim temDays As Integer
    Dim temDay1 As Date
    Dim temDay2 As Date
    Dim temValue As Double
    Dim i As Integer
    Dim tr As Integer
    Dim Flag As Boolean
    Dim ItemCount As Long
    Dim ArrayItems() As String
    Dim ArrayItemIDs() As Long
    Dim ii As Long
    Dim RowCount As Long
    
    If optAllItems.Value = False Then
        Flag = False
        ItemCount = 0
        For i = 0 To lstItemIDs.ListCount - 1
            If lstItemIDs.Selected(i) = True Then
                Flag = True
                ItemCount = ItemCount + 1
            End If
        Next
        If Flag = False Then
            tr = MsgBox("You have not selected an item", vbCritical, "Select Item")
            lstItems.SetFocus
            Exit Sub
        End If
    End If
    ReDim ArrayItemIDs(ItemCount) As Long
    ReDim ArrayItems(ItemCount) As String
    ii = 0
    For i = 0 To lstItemIDs.ListCount - 1
        If lstItemIDs.Selected(i) = True Then
            lstItemIDs.ListIndex = i
            lstItems.ListIndex = i
            ArrayItemIDs(ii) = Val(lstItemIDs.Text)
            ArrayItems(ii) = lstItems.Text
            ii = ii + 1
        End If
    Next
    
    
    
    If dtpFrom.Value > dtpTo.Value Then
        temDay1 = dtpTo.Value
        dtpTo.Value = dtpFrom.Value
        dtpFrom.Value = temDay1
    Else
        temDay1 = dtpFrom.Value
        temDay2 = dtpTo.Value
    End If
    
    TemPath = FSys.GetParentFolderName(Database)
    If FSys.FileExists(TemPath & "\Lucky1.xls") = False Then
        tr = MsgBox("There are no graphs on the specified location")
        Exit Sub
    End If
    
    frmPleaseWait.Show
    DoEvents
    
    Set myworkbook = GetObject(TemPath & "\Lucky1.xls")
    Set myworksheet = myworkbook.Worksheets.Item(1)
    Set mychart = myworkbook.Charts.Item(1)
    
    myworksheet.UsedRange.Clear
    myworksheet.Cells(1, 1) = "From"
    myworksheet.Cells(1, 2) = "To"
    myworksheet.Cells(1, 3) = "Period"
    
    If optSelectdeItem.Value = True Then
        For i = 0 To ItemCount - 1
            myworksheet.Cells(1, i + 4) = ArrayItems(i)
        Next
    Else
        If optByVal.Value = True Then
            myworksheet.Cells(1, 4) = "Total Value"
        ElseIf optByQty.Value = True Then
            myworksheet.Cells(1, 4) = "Total Quentity"
        End If
    End If
    
    RowCount = 0
    If optDaily.Value = True Then
        temDays = DateDiff("d", dtpFrom.Value, dtpTo.Value)
        If temDays < 0 Then
            tr = MsgBox("You have not selected a valid time period or interval. Please adjust", vbCritical, "Wrong settings")
            Unload frmPleaseWait
            dtpFrom.SetFocus
            Exit Sub
        End If
        For i = 0 To temDays
            RowCount = RowCount + 1
            myworksheet.Cells(i + 2, 1) = Format(dtpFrom.Value + i, "dd MMMM yyyy")
            myworksheet.Cells(i + 2, 2) = Format(dtpFrom.Value + i, "dd MMMM yyyy")
            myworksheet.Cells(i + 2, 3) = Format(dtpFrom.Value + i, LongDateFormat)
        Next
    ElseIf optWeekly.Value = True Then
        temDays = DateDiff("d", dtpFrom.Value, dtpTo.Value)
        If temDays < 21 Then
            tr = MsgBox("You have not selected a valid time period or interval. Please adjust", vbCritical, "Wrong settings")
            Unload frmPleaseWait
            dtpFrom.SetFocus
            Exit Sub
        End If
        For i = 0 To temDays Step 7
            RowCount = RowCount + 1
            temDay1 = dtpFrom.Value + i
            temDay2 = dtpFrom.Value + i + 7
            myworksheet.Cells((i \ 7) + 2, 1) = Format(dtpFrom.Value + i, "dd MMMM yyyy")
            myworksheet.Cells((i \ 7) + 2, 2) = Format(dtpFrom.Value + i + 7, "dd MMMM yyyy")
            myworksheet.Cells((i \ 7) + 2, 3) = "Week from " & Format(dtpFrom.Value + i, LongDateFormat)
        Next
    ElseIf optMonthly.Value = True Then
        temDays = DateDiff("m", dtpFrom.Value, dtpTo.Value)
        If temDays < 3 Then
            tr = MsgBox("You have not selected a valid time period or interval. Please adjust", vbCritical, "Wrong settings")
            Unload frmPleaseWait
            dtpFrom.SetFocus
            Exit Sub
        End If
        For i = 0 To temDays + 1
            RowCount = RowCount + 1
            temDay1 = DateSerial(Year(dtpFrom.Value), Month(dtpFrom.Value) + i, 1)
            temDay2 = DateSerial(Year(dtpFrom.Value), Month(dtpFrom.Value) + i + 1, 1) - 1
            myworksheet.Cells(i + 2, 1) = Format(temDay1, "DD MMMM yyyy")
            myworksheet.Cells(i + 2, 2) = Format(temDay2, "DD MMMM yyyy")
            myworksheet.Cells(i + 2, 3) = Format(temDay1, "MMMM yyyy")
        Next
    ElseIf optYearly.Value = True Then
        temDays = DateDiff("yyyy", dtpFrom.Value, dtpTo.Value)
        If temDays < 2 Then
            tr = MsgBox("You have not selected a valid time period or interval. Please adjust", vbCritical, "Wrong settings")
            Unload frmPleaseWait
            dtpFrom.SetFocus
            Exit Sub
        End If
        DoEvents
        For i = 0 To temDays
            RowCount = RowCount + 1
            temDay1 = DateSerial(Year(dtpFrom.Value) + i, 1, 1)
            temDay2 = DateSerial(Year(dtpFrom.Value) + i, 12, 31)
            myworksheet.Cells(i + 2, 1) = Format(temDay1, "dd MMMM yyyy")
            myworksheet.Cells(i + 2, 2) = Format(temDay2, "dd MMMM yyyy")
            myworksheet.Cells(i + 2, 3) = "Year " & Format(temDay1, "yyyy")
        Next
    End If
    
    
'SELECT Sum(tblSale.Amount - tblReturn.Amount) AS Display
'FROM (tblSale RIGHT JOIN tblSaleBill ON tblSale.SaleBillID = tblSaleBill.SaleBillID) LEFT JOIN tblReturn ON tblSale.SaleBillID = tblReturn.SaleBillID
'WHERE (((tblSaleBill.Date) Between #1/1/2008# And #12/31/2008#) AND ((tblSale.ItemID)=1));
        
'SELECT Sum([tblSale].[GrossPrice]-[tblReturn].[GrossPrice]) AS Display
'FROM (tblSale RIGHT JOIN tblSaleBill ON tblSale.SaleBillID = tblSaleBill.SaleBillID) LEFT JOIN tblReturn ON tblSale.SaleBillID = tblReturn.SaleBillID
'WHERE (((tblSaleBill.Date) Between #1/1/2008# And #12/31/2008#) AND ((tblSale.ItemID)=1));
'
    
    Dim TemSale As Double
    Dim TemReturn As Double
    
    
    If optSelectdeItem.Value = True Then
        For i = 0 To ItemCount - 1
            For ii = 0 To RowCount - 1
                With rsTem
                    If .State = 1 Then .Close
                    If optByVal.Value = True Then
                        temSelect = "SELECT Sum(tblSale.Amount) AS Sale "
                    Else
                        temSelect = "SELECT Sum(tblSale.Price) AS Sale "
                    End If
                    temFrom = "From tblSale "
                    temWhere = "WHERE (((tblSale.Date) Between '" & Format(myworksheet.Cells((ii + 2), 1), "dd MMMM yyyy") & "' And '" & Format(myworksheet.Cells(ii + 2, 2), "dd MMMM yyyy") & "')  AND ((tblSale.ItemID)=" & ArrayItemIDs(i) & "))"
                    temSql = temSelect & temFrom & temWhere
                    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                    If IsNull(!Sale) = False Then
                        TemSale = !Sale
                    Else
                        TemSale = 0
                    End If
                    If .State = 1 Then .Close
                    If optByVal.Value = True Then
                        temSelect = "SELECT Sum(tblReturn.Amount) AS Return "
                    Else
                        temSelect = "SELECT Sum(tblReturn.Price) AS Return "
                    End If
                    temFrom = "From tblReturn "
                    temWhere = "WHERE (((tblReturn.Date) Between '" & Format(myworksheet.Cells((ii + 2), 1), "dd MMMM yyyy") & "' And '" & Format(myworksheet.Cells(ii + 2, 2), "dd MMMM yyyy") & "')  AND ((tblReturn.ItemID)=" & ArrayItemIDs(i) & "))"
                    temSql = temSelect & temFrom & temWhere
                    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                    If IsNull(!Return) = False Then
                        TemReturn = !Return
                    Else
                        TemReturn = 0
                    End If
                    myworksheet.Cells(ii + 2, i + 4) = TemSale - TemReturn
                    If .State = 1 Then .Close
                End With
            Next ii
            DoEvents
        Next i
        mychart.SetSourceData myworksheet.Range("c1:" & GetColumnName(ItemCount + 3) & RowCount + 1)
    Else
        For ii = 0 To RowCount - 1
            With rsTem
                If .State = 1 Then .Close
                If optByVal.Value = True Then
                    temSelect = "SELECT Sum(tblSale.Amount) AS Sale "
                Else
                    temSelect = "SELECT Sum(tblSale.Price) AS Sale "
                End If
                temFrom = "From tblSale "
                temWhere = "WHERE (((tblSale.Date) Between '" & Format(myworksheet.Cells((ii + 2), 1), "dd MMMM yyyy") & "' And '" & Format(myworksheet.Cells(ii + 2, 2), "dd MMMM yyyy") & "'))"
                temSql = temSelect & temFrom & temWhere
                .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                If IsNull(!Sale) = False Then
                    TemSale = !Sale
                Else
                    TemSale = 0
                End If
                If .State = 1 Then .Close
                If optByVal.Value = True Then
                    temSelect = "SELECT Sum(tblReturn.Amount) AS Return "
                Else
                    temSelect = "SELECT Sum(tblReturn.Price) AS Return "
                End If
                temFrom = "From tblReturn "
                temWhere = "WHERE (((tblReturn.Date) Between '" & Format(myworksheet.Cells((ii + 2), 1), "dd MMMM yyyy") & "' And '" & Format(myworksheet.Cells(ii + 2, 2), "dd MMMM yyyy") & "'))"
                temSql = temSelect & temFrom & temWhere
                .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                If IsNull(!Return) = False Then
                    TemReturn = !Return
                Else
                    TemReturn = 0
                End If
                If .State = 1 Then .Close
            End With
            myworksheet.Cells(ii + 2, 4) = TemSale - TemReturn
            DoEvents
        Next ii
        mychart.SetSourceData myworksheet.Range("c1:D" & RowCount + 1)
    End If
        
    
    
    
    If optStandardChart.Value = True Or cmbChartType.ListIndex > 0 Then
        If optBar.Value = True Then
            If opt2D.Value = True Then
                mychart.ChartType = xlColumnClustered
            ElseIf opt3D.Value = True Then
                mychart.ChartType = xl3DColumn
            End If
        ElseIf optLine.Value = True Then
            If opt2D.Value = True Then
                mychart.ChartType = xlLine
            ElseIf opt3D.Value = True Then
                mychart.ChartType = xl3DLine
            End If
        ElseIf optPie.Value = True Then
            If opt2D.Value = True Then
                mychart.ChartType = xlPie
            Else
                mychart.ChartType = xl3DPie
            End If
        End If
    Else
        mychart.ChartType = cmbChartType.ItemData(cmbChartType.ListIndex)
    End If
    
    If optDisplayTitle.Value = True Then
        temTopic = ""
        If optDaily.Value = True Then
            temTopic = "Daily "
        ElseIf optWeekly.Value = True Then
            temTopic = "Weekly "
        ElseIf optMonthly.Value = True Then
            temTopic = "Monthly "
        ElseIf optYearly.Value = True Then
            temTopic = "Yearly "
        End If
        If optByQty.Value = True Then
            temTopic = temTopic & " Quentity-wise "
        ElseIf optByQty.Value = True Then
            temTopic = temTopic & " Value-wise "
        End If
        temTopic = temTopic & "Sale "
        If optAllItems.Value = True Then
            temTopic = temTopic & "of all Items "
        Else
            temTopic = temTopic & "of selected Items "
        End If
        If dtpFrom.Value = dtpTo.Value Then
            temSubTopic = "On " & Format(dtpFrom.Value, LongDateFormat)
        Else
            temSubTopic = "From " & Format(dtpFrom.Value, LongDateFormat) & " to " & Format(dtpTo.Value, LongDateFormat)
        End If
        mychart.HasTitle = True
        mychart.ChartTitle.Caption = temTopic & vbNewLine & temSubTopic
    Else
        mychart.HasTitle = False
    End If
    If optDisplayLegend.Value = True Then
        mychart.HasLegend = True
    Else
        mychart.HasLegend = False
    End If
    If optDisplayValues.Value = True Then
        mychart.ApplyDataLabels xlDataLabelsShowValue
    Else
        mychart.ApplyDataLabels xlDataLabelsShowNone
    End If
    
    mychart.HasLegend = True
    myworkbook.Save
    mychart.Activate
    Unload frmPleaseWait
    frmGraph.Show
    frmGraph.Caption = temTopic & " - " & temSubTopic
End Sub


Private Sub cmbCategory_Change()
    Call FillLists
End Sub

Private Sub FillLists()
    Screen.MousePointer = vbHourglass
    With rsItem
        If .State = 1 Then .Close
        If IsNumeric(cmbGeneric.BoundText) = True And IsNumeric(cmbCategory.BoundText) = True Then
            temSql = "Select * From tblItem where GenericNameID = " & Val(cmbGeneric.BoundText) & " AND ItemCategoryID = " & Val(cmbCategory.BoundText) & " Order By Display"
        ElseIf IsNumeric(cmbGeneric.BoundText) = True And IsNumeric(cmbCategory.BoundText) = False Then
            temSql = "Select * From tblItem where GenericNameID = " & Val(cmbGeneric.BoundText) & " Order By Display"
        ElseIf IsNumeric(cmbGeneric.BoundText) = False And IsNumeric(cmbCategory.BoundText) = True Then
            temSql = "Select * From tblItem where ItemCategoryID = " & Val(cmbCategory.BoundText) & " Order By Display"
        Else
            temSql = "Select * From tblItem Order By Display"
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        lstItemIDs.Clear
        lstItems.Clear
        If .RecordCount > 0 Then
            While .EOF = False
                lstItemIDs.AddItem !ItemID
                lstItems.AddItem !Display
                .MoveNext
            Wend
        End If
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmbCategory_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cmbCategory.Text = Empty
    End If
End Sub

Private Sub cmbChartType_Change()
    On Error Resume Next
    If cmbChartType.ListIndex > 0 Then
        mychart.ChartType = cmbChartType.ItemData(cmbChartType.ListIndex)
    End If
End Sub

Private Sub cmbChartType_Click()
    On Error Resume Next
    If cmbChartType.ListIndex > 0 Then
        mychart.ChartType = cmbChartType.ItemData(cmbChartType.ListIndex)
    End If
End Sub

Private Sub cmbChartType_Scroll()
    On Error Resume Next
    If cmbChartType.ListIndex > 0 Then
        mychart.ChartType = cmbChartType.ItemData(cmbChartType.ListIndex)
    End If
End Sub

Private Sub cmbGeneric_Change()
    Call FillLists
End Sub

Private Sub cmbGeneric_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cmbGeneric.Text = Empty
    End If
End Sub

Private Sub Form_Load()
    Call FillCombos
    lstItemIDs.Visible = False
    lstItems.Enabled = False
    cmbChartType.Enabled = False
    dtpFrom.Value = Date
    dtpTo.Value = Date
    SSTab1.Tab = 0
    With cmbChartType
        .AddItem "3D Area"
        .AddItem "3D Area Stacked"
        .AddItem "3D Area Stacked 100"
        .AddItem "xl3DBar"
        .AddItem "3D Bar Clustered"
        .AddItem "3DBarStacked"
        .AddItem "3DBarStacked100"
        .AddItem "3DColumn"
        .AddItem "3DColumnClustered"
        .AddItem "3DColumnStacked"
        .AddItem "3DColumnStacked100"
        .AddItem "3DLine"
        .AddItem "3DPie"
        .AddItem "3DPieExploded"
        .AddItem "Area"
        .AddItem "AreaStacked"
        .AddItem "AreaStacked100"
        .AddItem "BarClustered"
        .AddItem "BarOfPie"
        .AddItem "BarStacked"
        .AddItem "BarStacked"
        .AddItem "BarStacked100"
        .AddItem "Bubble"
        .AddItem "Bubble3DEffect"
        .AddItem "Column"
        .AddItem "ColumnClustered"
        .AddItem "ColumnStacked"
        .AddItem "ColumnStacked"
        .AddItem "ColumnStacked100"
        .AddItem "ConeBarClustered"
        .AddItem "ConeBarStacked"
        .AddItem "ConeBarStacked100"
        .AddItem "ConeCol"
        .AddItem "ConeColClustered"
        .AddItem "ConeColStacked"
        .AddItem "ConeColStacked100"
        .AddItem "Cylinder"
        .AddItem "CylinderBarClustered"
        .AddItem "CylinderBarStacked"
        .AddItem "CylinderBarStacked100"
        .AddItem "CylinderCol"
        .AddItem "CylinderColClustered"
        .AddItem "CylinderColStacked"
        .AddItem "CylinderColStacked100"
        .AddItem "Doughnut"
        .AddItem "DoughnutExploded"
        .AddItem "Line"
        .AddItem "LineMarkers"
        .AddItem "LineMarkersStacked"
        .AddItem "LineMarkersStacked100"
        .AddItem "LineStacked"
        .AddItem "LineStacked100"
        .AddItem "Pie"
        .AddItem "PieExploded"
        .AddItem "PieOfPie"
        .AddItem "PyramidBarClustered"
        .AddItem "PyramidBarStacked"
        .AddItem "PyramidBarStacked100"
        .AddItem "PyramidCol"
        .AddItem "PyramidColClustered"
        .AddItem "PyramidColStacked"
        .AddItem "PyramidColStacked100"
        .AddItem "Radar"
        .AddItem "RadarFilled"
        .AddItem "RadarMarkers"
        .AddItem "Surface"
        .AddItem "SurfaceTopView"
        .AddItem "SurfaceTopViewWireframe"
        .AddItem "SurfaceWireframe"
        .AddItem "XYScatter"
        .AddItem "XYScatterLines"
        .AddItem "XYScatterLinesNoMarkers"
        .AddItem "XYScatterSmooth"
        .AddItem "XYScatterSmoothNoMarkers"
        
        .ItemData(0) = xl3DArea
        .ItemData(1) = xl3DAreaStacked
        .ItemData(2) = xl3DAreaStacked
        .ItemData(3) = xl3DBarClustered
        .ItemData(4) = xl3DBarClustered
        .ItemData(5) = xl3DBarStacked
        .ItemData(6) = xl3DBarStacked100
        .ItemData(7) = xl3DColumn
        .ItemData(8) = xl3DColumnClustered
        .ItemData(9) = xl3DColumnStacked
        .ItemData(10) = xl3DColumnStacked100
        .ItemData(11) = xl3DLine
        .ItemData(12) = xl3DPie
        .ItemData(13) = xl3DPieExploded
        .ItemData(14) = xlArea
        .ItemData(15) = xlAreaStacked
        .ItemData(16) = xlAreaStacked100
        .ItemData(17) = xlBarClustered
        .ItemData(18) = xlBarOfPie
        .ItemData(19) = xlBarStacked
        .ItemData(20) = xlBarStacked
        .ItemData(21) = xlBarStacked100
        .ItemData(22) = xlBubble
        .ItemData(23) = xlBubble3DEffect
        .ItemData(24) = xlColumnClustered
        .ItemData(25) = xlColumnClustered
        .ItemData(26) = xlColumnStacked
        .ItemData(27) = xlColumnStacked
        .ItemData(28) = xlColumnStacked100
        .ItemData(29) = xlConeBarClustered
        .ItemData(30) = xlConeBarStacked
        .ItemData(31) = xlConeBarStacked100
        .ItemData(32) = xlConeCol
        .ItemData(33) = xlConeColClustered
        .ItemData(34) = xlConeColStacked
        .ItemData(35) = xlConeColStacked100
        .ItemData(36) = xlCylinderBarClustered
        .ItemData(37) = xlCylinderBarClustered
        .ItemData(38) = xlCylinderBarStacked
        .ItemData(39) = xlCylinderBarStacked100
        .ItemData(40) = xlCylinderCol
        .ItemData(41) = xlCylinderColClustered
        .ItemData(42) = xlCylinderColStacked
        .ItemData(43) = xlCylinderColStacked100
        .ItemData(44) = xlDoughnut
        .ItemData(45) = xlDoughnutExploded
        .ItemData(46) = xlLine
        .ItemData(47) = xlLineMarkers
        .ItemData(48) = xlLineMarkersStacked
        .ItemData(49) = xlLineMarkersStacked100
        .ItemData(50) = xlLineStacked
        .ItemData(51) = xlLineStacked100
        .ItemData(52) = xlPie
        .ItemData(53) = xlPieExploded
        .ItemData(54) = xlPieOfPie
        .ItemData(55) = xlPyramidBarClustered
        .ItemData(56) = xlPyramidBarStacked
        .ItemData(57) = xlPyramidBarStacked100
        .ItemData(58) = xlPyramidCol
        .ItemData(59) = xlPyramidColClustered
        .ItemData(60) = xlPyramidColStacked
        .ItemData(61) = xlPyramidColStacked100
        .ItemData(62) = xlRadar
        .ItemData(63) = xlRadarFilled
        .ItemData(64) = xlRadarMarkers
        .ItemData(65) = xlSurface
        .ItemData(66) = xlSurfaceTopView
        .ItemData(67) = xlSurfaceTopViewWireframe
        .ItemData(68) = xlSurfaceWireframe
        .ItemData(69) = xlXYScatter
        .ItemData(70) = xlXYScatterLines
        .ItemData(71) = xlXYScatterLinesNoMarkers
        .ItemData(72) = xlXYScatterSmooth
        .ItemData(73) = xlXYScatterSmoothNoMarkers
        
    
    
    End With
End Sub

Private Sub FillCombos()
    With rsGeneric
        If .State = 1 Then .Close
        temSql = "Select * from tblGenericName order by GenericName"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbGeneric
        Set .RowSource = rsGeneric
        .ListField = "GenericName"
        .BoundColumn = "GenericNameID"
    End With
    With rsCategory
        If .State = 1 Then .Close
        temSql = "Select * from tblItemCategory order by ItemCategory"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbCategory
        Set .RowSource = rsCategory
        .ListField = "ItemCategory"
        .BoundColumn = "ItemCategoryID"
    End With
End Sub


Private Sub lstItems_Click()
    lstItemIDs.ListIndex = lstItems.ListIndex
    lstItemIDs.Selected(lstItems.ListIndex) = lstItems.Selected(lstItems.ListIndex)
End Sub

Private Function GetColumnName(ColumnNo As Long) As String
    Dim TemNum As Integer
    Dim temnum1 As Integer
    
    If ColumnNo < 27 Then
        GetColumnName = Chr(ColumnNo + 64)
    Else
        TemNum = ColumnNo \ 26
        temnum1 = ColumnNo Mod 26
        GetColumnName = Chr(TemNum + 64) & Chr(temnum1 + 64)
    End If
End Function

Private Sub opt2D_Click()
    Call SetGraph
End Sub

Private Sub opt3D_Click()
    Call SetGraph
End Sub

Private Sub OptAllItems_Click()
    If optAllItems.Value = True Then
        lstItems.Enabled = False
        Dim i As Integer
        For i = 0 To lstItems.ListCount - 1
            lstItemIDs.Selected(i) = False
            lstItems.Selected(i) = False
        Next i
    Else
        lstItems.Enabled = True
    End If
End Sub

Private Sub optBar_Click()
    Call SetGraph
End Sub

Private Sub optDisplayLegend_Click()
    Call SetGraph
End Sub

Private Sub optDisplayTitle_Click()
    Call SetGraph
End Sub

Private Sub optDisplayValues_Click()
    Call SetGraph
End Sub

Private Sub optDoNotDisplayValues_Click()
    Call SetGraph
End Sub


Private Sub optLine_Click()
    Call SetGraph
End Sub

Private Sub optNoLegend_Click()
    Call SetGraph
End Sub

Private Sub optPie_Click()
    Call SetGraph
End Sub

Private Sub optNoTitle_Click()
    Call SetGraph
End Sub

Private Sub optOtherCharts_Click()
    If optOtherCharts.Value = True Then
        cmbChartType.Enabled = True
    Else
        cmbChartType.Enabled = False
    End If
End Sub

Private Sub optSelectdeItem_Click()
    If optSelectdeItem.Value = True Then
        lstItems.Enabled = True
    Else
        lstItems.Enabled = False
        Dim i As Integer
        For i = 0 To lstItems.ListCount - 1
            lstItemIDs.Selected(i) = False
            lstItems.Selected(i) = False
        Next i
    End If
End Sub

Private Sub optStandardChart_Click()
    If optStandardChart.Value = True Then
        cmbChartType.Enabled = False
    Else
        cmbChartType.Enabled = True
    End If
End Sub

Private Sub optXAxis_Click()
    If optXAxis.Value = True Then
        mychart.PlotBy = xlColumns
    ElseIf optYAxis.Value = True Then
        mychart.PlotBy = xlRows
    End If
End Sub

Private Sub optYAxis_Click()
    If optXAxis.Value = True Then
        mychart.PlotBy = xlColumns
    ElseIf optYAxis.Value = True Then
        mychart.PlotBy = xlRows
    End If
End Sub

Private Sub SetGraph()
    If optBar.Value = True Then
        If opt2D.Value = True Then
            mychart.ChartType = xlColumnClustered
        ElseIf opt3D.Value = True Then
            mychart.ChartType = xl3DColumn
        End If
    ElseIf optLine.Value = True Then
        If opt2D.Value = True Then
            mychart.ChartType = xlLine
        ElseIf opt3D.Value = True Then
            mychart.ChartType = xl3DLine
        End If
    ElseIf optPie.Value = True Then
        If opt2D.Value = True Then
            mychart.ChartType = xlPie
        Else
            mychart.ChartType = xl3DPie
        End If
    End If
    optStandardChart.Value = True
    optOtherCharts.Value = False
    cmbChartType.Enabled = False
    If optDisplayTitle.Value = True Then
        mychart.HasTitle = True
    Else
        mychart.HasTitle = False
    End If
    If optDisplayLegend.Value = True Then
        mychart.HasLegend = True
    Else
        mychart.HasLegend = False
    End If
    If optDisplayValues.Value = True Then
        mychart.ApplyDataLabels xlDataLabelsShowValue
    Else
        mychart.ApplyDataLabels xlDataLabelsShowNone
    End If
End Sub
