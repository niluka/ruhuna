Attribute VB_Name = "modBillPrint"
Option Explicit
    Public Type MyBillPoints
        DX As Double
        DY As Double
        VX As Double
        VY As Double
        CX As Double
        CY As Double
        CenterX As Double
    End Type

Public Function PrintThisBill(Optional ReceiptNo As String, Optional PaymentType As String, Optional PatientName As String, Optional BillDate As String, Optional BillTime As String, Optional Bill As String, Optional Comment As String) As MyBillPoints

    Dim X1 As Double
    Dim X2 As Double
    Dim X3 As Double
    Dim Y1 As Double
    Dim Y2 As Double
    Dim Y3 As Double
    Dim Y4 As Double
    Dim Y5 As Double
    Dim Y6 As Double
    Dim C As Double
    
    Dim LeftMargin As Double
    Dim TopMargin As Double
    
    
    LeftMargin = 0 - (1440 * 0.1)
    TopMargin = 0 + (1440 * 0.05)
    
    Dim TipsPerInch As Double
    
    TipsPerInch = 1440

     X1 = 0.7 * TipsPerInch + LeftMargin
     X2 = 1.75 * TipsPerInch + LeftMargin
     X3 = 3.85 * TipsPerInch + LeftMargin
     
     Y1 = 0.5 * TipsPerInch + TopMargin
     Y2 = 0.68 * TipsPerInch + TopMargin
     Y3 = 0.85 * TipsPerInch + TopMargin
     Y4 = 1.6 * TipsPerInch + TopMargin
     Y5 = 4.7 * TipsPerInch + TopMargin
     Y6 = 1.1 * TipsPerInch + TopMargin
     
     C = 2.4 * TipsPerInch + LeftMargin

    With Printer


'        Printer.Line (X1, Y1)-(X3, Y5), , B
        
        .CurrentX = X2
        .CurrentY = Y1
        .Font.Name = "Tahoma"
        .Font.Size = 9
        .Font.Bold = False
        .Font.Italic = False
        If ReceiptNo <> "" Then Printer.Print ReceiptNo


        .CurrentX = X2
        .CurrentY = Y2
        .Font.Name = "Tahoma"
        .Font.Size = 9
        .Font.Bold = False
        .Font.Italic = False
        If PaymentType <> "" Then Printer.Print PaymentType


        .CurrentX = X2
        .CurrentY = Y3
        .Font.Name = "Tahoma"
        .Font.Size = 9
        .Font.Bold = False
        .Font.Italic = False
        If PatientName <> "" Then Printer.Print PatientName


        .CurrentX = X3
        .CurrentY = Y1
        .Font.Name = "Tahoma"
        .Font.Size = 9
        .Font.Bold = False
        .Font.Italic = False
        If BillDate <> "" Then Printer.Print BillDate


        .CurrentX = X3
        .CurrentY = Y2
        .Font.Name = "Tahoma"
        .Font.Size = 9
        .Font.Bold = False
        .Font.Italic = False
        If BillTime <> "" Then Printer.Print BillTime


        .CurrentX = C - (.TextWidth(Bill) / 2)
        .CurrentY = Y6
        .Font.Name = "Tahoma"
        .Font.Size = 11
        .Font.Bold = False
        .Font.Italic = False
        If Bill <> "" Then Printer.Print Bill


        .CurrentX = X1
        .CurrentY = Y5
        .Font.Name = "Tahoma"
        .Font.Size = 9
        .Font.Bold = False
        .Font.Italic = False
        If Comment <> "" Then Printer.Print Comment


        PrintThisBill.DX = X1
        PrintThisBill.DY = Y4
        PrintThisBill.VX = X3 + 720
        PrintThisBill.VY = Y4
        PrintThisBill.CX = X3
        PrintThisBill.CY = Y5
        PrintThisBill.CenterX = C

    End With

End Function
