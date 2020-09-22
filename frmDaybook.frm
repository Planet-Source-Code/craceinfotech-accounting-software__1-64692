VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Daybook 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daybook"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   10755
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   300
      Left            =   5978
      TabIndex        =   2
      Top             =   6300
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   300
      Left            =   4298
      TabIndex        =   1
      Top             =   6300
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   5835
      Left            =   255
      TabIndex        =   0
      Top             =   180
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   10292
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmDaybook.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Daybook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LineCount As Integer
Dim HeaderLength As Integer
Dim FooterLength As Integer
Dim DetailLength As Integer
Dim PageLength As Integer
Dim PageWidth As Integer
Dim PageCount As Integer
'Dim CashInHand As Currency
Dim DaybookDetailRS As Recordset

Dim NumberOfRecords As Long
Dim ExcessNarration As Boolean
Dim DebitBalance As Currency
Dim CreditBalance As Currency
Dim NewDate As Boolean
Dim DaybookDate As Date

Dim NewPage As Boolean
Dim PrintedDate As Date

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
Screen.MousePointer = vbHourglass
PrepareDaybook
If NumberOfRecords > 0 Then
    RichTextBox1.LoadFile "c:\vbprog\vba\rpt\daybook.txt", rtfText
    cmdPrint.Enabled = True
Else
    Unload Me
End If
Screen.MousePointer = vbDefault

End Sub

Private Sub DaybookHeader()
    'IMPORTANT:10+2+36+2+14+2+14=80
    Print #1,
    Print #1, PADC(CompanyName, PageWidth); Tab(PageWidth - (7 + 4)); "Page : "; Format(PageCount, "@@@@")
    Print #1, PADC("DAY BOOK", PageWidth)
    Print #1, String(PageWidth, "-")
    'Print #1, "Date"; Tab(13); "Particulars"; Tab(51); "    Debit"; Tab(67); "    Credit"
    Print #1, "Date"; Tab(13); "Particulars"; Tab(51); "        Credit"; Tab(67); "        Debit"
    Print #1, String(PageWidth, "-")
    If PageCount > 1 Then
        Print #1, Tab(45); "b/f"; Tab(51); Format(Format(CreditBalance, "0.00"), "@@@@@@@@@@@@@@");
        Print #1, Spc(2); Format(Format(DebitBalance, "0.00"), "@@@@@@@@@@@@@@")
        Print #1,
    End If
    
    LineCount = 0
    NewPage = True

'HeaderLength = 7
End Sub
Private Sub DaybookFooter()
    Print #1,
    Print #1, String(PageWidth, "-")
    Print #1, Tab(45); "c/d"; Tab(51); Format(Format(CreditBalance, "0.00"), "@@@@@@@@@@@@@@");
    Print #1, Spc(2); Format(Format(DebitBalance, "0.00"), "@@@@@@@@@@@@@@")
    Print #1, String(PageWidth, "-")
    'Print #1,
    Print #1,
    'Print #1, "Page Length : "; LineCount + HeaderLength + FooterLength
    Print #1, Chr(12)
    'FooterLength = 5
    'If LineCount + HeaderLength + FooterLength > 67 Then MsgBox PageCount
    'If LineCount + HeaderLength + FooterLength < 65 Then MsgBox PageCount
End Sub
Private Sub DaybookDetail()
    Dim j As Integer

    If (DaybookDate <> PrintedDate) Or NewPage Then
    Print #1, Format$(DaybookDetailRS!acn_date, "dd/mm/yyyy"); Spc(2);
    PrintedDate = Format$(DaybookDetailRS!acn_date, "dd/mm/yyyy")
    NewPage = False
    Else
    Print #1, Tab(13);
    End If
    Print #1, Mid(DaybookDetailRS!ACTITLE, 1, 36)
    LineCount = LineCount + 1
    
    j = PrintNarration(DaybookDetailRS!particular, 1)
    'Print #1, Tab(13); Spc(3); Mid(DaybookDetailRS!particular, 1, 30);
    Print #1, Tab(51);
    Print #1, ZeroSup(DaybookDetailRS!credit);
    Print #1, Spc(2);
    Print #1, ZeroSup(DaybookDetailRS!debit);
    'LineCount = LineCount + 1
    
    DebitBalance = DebitBalance + DaybookDetailRS!debit
    CreditBalance = CreditBalance + DaybookDetailRS!credit
    CashInHand = CreditBalance - DebitBalance
    
End Sub
Private Sub DaybookSummary()
        Print #1, Tab(51); String(30, "-")
        Print #1, Tab(51); Format(Format(CreditBalance, "0.00"), "@@@@@@@@@@@@@@");
        Print #1, Spc(2); Format(Format(DebitBalance, "0.00"), "@@@@@@@@@@@@@@")
        Print #1, Tab(51); String(30, "-")

        CreditBalance = IIf(CashInHand > 0, CashInHand, 0)
        DebitBalance = IIf(CashInHand < 0, -CashInHand, 0)

        Print #1,
        Print #1, Tab(36); "Cash In Hand";
        If CashInHand >= 0 Then
            Print #1, Tab(51); Format(Format(CashInHand, "0.00"), "@@@@@@@@@@@@@@")
        Else
            Print #1, Tab(67); Format(Format(-CashInHand, "0.00"), "@@@@@@@@@@@@@@")
        End If
        Print #1,

        Print #1,
        Print #1, String(PageWidth, "-")
        Print #1,
        Print #1,
        Print #1, Chr(12)
        'summaryLength = 10

End Sub

Private Sub PrepareDaybook()
    Dim i As Long
    Dim DaybookStart As Boolean
    PageLength = 66
    PageWidth = 80
    HeaderLength = 8
    FooterLength = 5
    DetailLength = PageLength - (HeaderLength + FooterLength + 1)
    PageCount = 0
    LineCount = DetailLength + 1
    cmdPrint.Enabled = False
    CashInHand = 0
    DebitBalance = 0
    CreditBalance = 0
    NewDate = False
    DaybookDate = FromDate - 1
    i = 0
    DaybookStart = True

    
    Set DaybookDetailRS = New Recordset
    DaybookDetailRS.Open "select BALANCETYP,BALANCE FROM " & MasterTable & " where actitle='CASH'" _
     , db, adOpenStatic, adLockReadOnly, adCmdText
    If DaybookDetailRS.EOF = False And DaybookDetailRS.BOF = False Then
        If DaybookDetailRS!balancetyp = "D" Then
            CashInHand = DaybookDetailRS!Balance
        Else
            CashInHand = -(DaybookDetailRS!Balance)
        End If
        
    Else
        CashInHand = 0
    End If
    DaybookDetailRS.Close
    Set DaybookDetailRS = New Recordset
    DaybookDetailRS.Open "select sum(credit-debit) FROM " & TransactionTable & " where " _
     & "acn_date between {" & Format(StartDate, "mm/dd/yyyy") & "} and {" & Format(FromDate - 1, "mm/dd/yyyy") & "} ", _
     db, adOpenStatic, adLockReadOnly, adCmdText
     
     With DaybookDetailRS
        If Not (.BOF = True And .EOF = True) Then
            CashInHand = CashInHand + .Fields(0).Value
        End If
        .Close
    End With
    
    Set DaybookDetailRS = New Recordset
    
    DaybookDetailRS.Open "select acn_date,entry_id,actitle,particular,debit," _
     & "credit FROM " & TransactionTable & " e," & MasterTable & " m where e.acnumber=m.acnumber and " _
     & "acn_date between {" & Format(FromDate, "mm/dd/yyyy") & "} and {" & Format(ToDate, "mm/dd/yyyy") & "} order by " _
     & "acn_date,entry_id ", db, adOpenStatic, adLockReadOnly, adCmdText
    
    NumberOfRecords = DaybookDetailRS.RecordCount
    
    
    
    If NumberOfRecords = 0 Then
        MsgBox "No Records"
        Exit Sub
    End If
    
    DaybookDetailRS.MoveFirst
    DaybookDate = DaybookDetailRS!acn_date
    DebitBalance = IIf(CashInHand < 0, -CashInHand, 0)
    CreditBalance = IIf(CashInHand > 0, CashInHand, 0)
    
    Open "c:\vbprog\vba\rpt\daybook.txt" For Output As #1
    
    Do While i < NumberOfRecords
        
        If LineCount > DetailLength Then
            If Not ExcessNarration Then
                If PageCount > 0 Then
                    DaybookFooter
                End If
            End If
            If Not ExcessNarration Then
                PageCount = PageCount + 1
                DaybookHeader
            End If
        End If
        
            If DaybookStart Then
                Print #1,
                Print #1, Tab(36); "Cash In Hand";
                If CashInHand >= 0 Then
                    Print #1, Tab(51); Format(Format(CashInHand, "0.00"), "@@@@@@@@@@@@@@")
                Else
                    Print #1, Tab(67); Format(Format(-CashInHand, "0.00"), "@@@@@@@@@@@@@@")
                End If
                Print #1,
                LineCount = LineCount + 2
                DaybookStart = False
            End If
            
            If Format(DaybookDetailRS!acn_date, "dd/mm/yyyy") <> Format(DaybookDate, "dd/mm/yyyy") Then
                NewDate = True
                DaybookDate = DaybookDetailRS!acn_date
            Else
                NewDate = False
            End If
                'CashInHand = CreditBalance - DebitBalance

        If NewDate Then
                    If DetailLength - LineCount > 6 Then
                        Print #1, Tab(51); String(30, "-")
                        Print #1, Tab(51); Format(Format(CreditBalance, "0.00"), "@@@@@@@@@@@@@@");
                        Print #1, Spc(2); Format(Format(DebitBalance, "0.00"), "@@@@@@@@@@@@@@")
                        Print #1, Tab(51); String(30, "-")
        
                        CreditBalance = IIf(CashInHand > 0, CashInHand, 0)
                        DebitBalance = IIf(CashInHand < 0, -CashInHand, 0)

                        Print #1,
                        Print #1, Tab(36); "Cash In Hand";
                        If CashInHand >= 0 Then
                            Print #1, Tab(51); Format(Format(CashInHand, "0.00"), "@@@@@@@@@@@@@@")
                        Else
                            Print #1, Tab(67); Format(Format(-CashInHand, "0.00"), "@@@@@@@@@@@@@@")
                        End If
                        Print #1,
                        LineCount = LineCount + 6
                    Else
                        Print #1,
                        LineCount = LineCount + 6
                    End If
            End If

        If LineCount < DetailLength Then
            With DaybookDetailRS
                If Not .EOF Then
                    DaybookDetail
                    .MoveNext
                    i = i + 1
                End If
            End With
        Else
            LineCount = DetailLength + 1
        End If
        
    Loop
    If DetailLength - LineCount > 10 Then
        DaybookSummary
    Else
        DaybookFooter
        PageCount = PageCount + 1
        DaybookHeader
        DaybookSummary
    End If
    
        
    'Print #1, Chr(12)
    Close #1
    
End Sub

Private Sub cmdprint_Click()
    cmdPrint.Enabled = False
    PrintText (RichTextBox1.Text)
End Sub

Private Function PrintNarration(Gstring As String, Gfile As Integer) As Integer
Dim Narration1 As String
Dim Ncount As Integer
Dim FileNumber As Integer
Dim SpacePosition As Integer

FileNumber = Gfile
Ncount = 0
Gstring = Trim(Gstring)
Narration1 = ""

Do While True
    SpacePosition = InStr(Gstring, " ")
    If SpacePosition > 30 Then
        Narration1 = Mid(Gstring, 1, 30)
        Gstring = Mid(Gstring, 31)
    ElseIf Len(Gstring) <= 30 Then
        Narration1 = Mid(Gstring, 1, 30)
        Gstring = Mid(Gstring, 31)
    ElseIf Len(Gstring) > 30 And SpacePosition = 0 Then
        Narration1 = Mid(Gstring, 1, 30)
        Gstring = Mid(Gstring, 31)
    End If
Do While True
    SpacePosition = InStr(Gstring, " ")
    If SpacePosition = 0 Then Exit Do
    If Len(Narration1) + SpacePosition <= 31 Then
        Narration1 = Narration1 + Mid(Gstring, 1, SpacePosition)
        Gstring = Mid(Gstring, SpacePosition + 1)
    Else
        Exit Do
    End If
Loop
    'Debug.Print Trim(Narration1)
    Print #FileNumber, Tab(13); Spc(3); Trim(Narration1);
    Narration1 = ""
    Ncount = Ncount + 1
    LineCount = LineCount + 1
    If Len(Gstring) = 0 Then
        ExcessNarration = False
        Exit Do
    End If
    If LineCount > DetailLength Then
        Print #FileNumber,
        DaybookFooter
        Ncount = 0
        PageCount = PageCount + 1
        ExcessNarration = True
        DaybookHeader
    Else
        ExcessNarration = False
    End If
Loop
PrintNarration = Ncount
End Function
