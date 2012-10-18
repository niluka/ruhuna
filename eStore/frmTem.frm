VERSION 5.00
Begin VB.Form frmTem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   7395
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txtPrint 
      Height          =   1935
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   720
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   4200
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "frmTem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim CsetPrinter As New cSetDfltPrinter
    Dim NumForms As Long, i As Long
    Dim FI1 As FORM_INFO_1
    Dim aFI1() As FORM_INFO_1
    Dim Temp() As Byte
    Dim BytesNeeded As Long
    Dim PrinterName As String
    Dim PrinterHandle As Long
    Dim FormItem As String
    Dim RetVal As Long
    Dim FormSize As SIZEL
    Dim SetPrinter As Boolean
    
    
    
Private Sub Command1_Click()
   
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text)

    PrinterName = Printer.DeviceName
    
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        ClosePrinter (PrinterHandle)
    End If
    
    
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
    
        
    Dim MyPrinter As VB.Printer
    For Each MyPrinter In VB.Printers
        If MyPrinter.DeviceName = cmbPrinter.Text Then
            Set Printer = MyPrinter
        End If
    Next
    
    
    Printer.Print txtPrint.Text
    Printer.EndDoc

End Sub

Private Sub Command2_Click()
    Dim rsTem As New ADODB.Recordset
    Dim rsNew As New ADODB.Recordset
    Dim i As Integer
    Dim cols As Integer
    Dim temsql As String
    With rsTem
        temsql = "select * from tblItem order by ItemID"
        .Open temsql, cnnStores, adOpenStatic, adLockReadOnly
        cols = rsTem.Fields.Count
        While .EOF = False
            temsql = "select * from tblItemNew where ItemID = " & !ItemID
            rsNew.Open temsql, cnnStores, adOpenStatic, adLockOptimistic
            If rsNew.RecordCount > 0 Then
                For i = 1 To cols - 1
                    rsNew.Fields(i).Value = .Fields(i).Value
                Next
                rsNew.Update
            End If
            rsNew.Close
            .MoveNext
        Wend
        .Close
    End With
End Sub

Private Sub Form_Load()
    Call FillPrinters
    txtPrint.Text = "sfdpjsdfsd" & vbNewLine & "asd asdas asdasd" & "sfdpjsdfsd" & vbNewLine & "asd asdas asdasd" & "sfdpjsdfsd" & vbNewLine & "asd asdas asdasd" & "sfdpjsdfsd" & vbNewLine & "asd asdas asdasd" & "sfdpjsdfsd" & vbNewLine & "asd asdas asdasd"
End Sub

Private Sub FillPrinters()
    
    Dim MyPrinter As VB.Printer
    For Each MyPrinter In VB.Printers
        cmbPrinter.AddItem MyPrinter.DeviceName
    Next
End Sub
