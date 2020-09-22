VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Ledger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ledger"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   10875
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FrameLedger 
      Height          =   6735
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   10575
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   300
         Left            =   4050
         TabIndex        =   2
         Top             =   6240
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   5730
         TabIndex        =   1
         Top             =   6240
         Width           =   1095
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   5835
         Left            =   180
         TabIndex        =   3
         Top             =   180
         Width           =   10245
         _ExtentX        =   18071
         _ExtentY        =   10292
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"frmLedger.frx":0000
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
End
Attribute VB_Name = "Ledger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'***********************************************
'This Software is developed by craceinfotech.
'Web site : http://www.craceinfotech.com
'email    : craceinfotech.yahoo.com
'date     : 18.03.2006
'***********************************************

Dim LineCount As Integer
Dim HeaderLength As Integer
Dim FooterLength As Integer
Dim DetailLength As Integer
Dim PageLength As Integer
Dim PageWidth As Integer
Dim PageCount As Integer
Dim HeadBalance As Currency
Dim LedgerDetailRS As Recordset

Dim NumberOfRecords As Long
Dim ExcessNarration As Boolean
Dim DebitBalance As Currency
Dim CreditBalance As Currency
Dim NewDate As Boolean
Dim LedgerDate As Date

Dim NewPage As Boolean
Dim PrintedDate As Date

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbHourglass
    PrepareLedger
    RichTextBox1.LoadFile "c:\vbprog\vba\rpt\Ledger.txt", rtfText
    cmdPrint.Enabled = True
    RichTextBox1.SetFocus
    Screen.MousePointer = vbDefault
End Sub

Private Sub LedgerHeader()
    'IMPORTANT:10+2+36+2+14+2+14=80
    Print #1,
    Print #1, PADC(CompanyName, PageWidth); Tab(PageWidth - (7 + 4)); "Page : "; Format(PageCount, "@@@@")
    Print #1, PADC("STATEMENT OF ACCOUNTS", PageWidth)
    Print #1,
    Print #1, PADC(SelectedHead, PageWidth)
    Print #1, PADC("(From " & Format(FromDate, "dd/mm/yyyy") & " to " & Format(ToDate, "dd/mm/yyyy") & ")", PageWidth)
    Print #1, String(PageWidth, "-")
    Print #1, "Date"; Tab(13); "Particulars"; Tab(51); "    Debit"; Tab(67); "    Credit"
    Print #1, String(PageWidth, "-")
    
    If PageCount = 1 Then
        Print #1, Tab(23); "Opening Balance";
        If HeadBalance >= 0 Then
            Print #1, Tab(67); Format(Format(HeadBalance, "0.00"), "@@@@@@@@@@@@@@")
        Else
            Print #1, Tab(51); Format(Format(-HeadBalance, "0.00"), "@@@@@@@@@@@@@@")
        End If

        Print #1,
    Else
        Print #1, Tab(45); "b/f"; Tab(51); Format(Format(DebitBalance, "0.00"), "@@@@@@@@@@@@@@");
        Print #1, Spc(2); Format(Format(CreditBalance, "0.00"), "@@@@@@@@@@@@@@")
        Print #1,
    End If
    
    LineCount = 0
    NewPage = True

'HeaderLength = 11
End Sub
Private Sub LedgerFooter()
    Print #1,
    Print #1, String(PageWidth, "-")
    Print #1, Tab(45); "c/d"; Tab(51); Format(Format(DebitBalance, "0.00"), "@@@@@@@@@@@@@@");
    Print #1, Spc(2); Format(Format(CreditBalance, "0.00"), "@@@@@@@@@@@@@@")
    Print #1,
    'Print #1, "Page Length : "; LineCount + HeaderLength + FooterLength
    Print #1, Chr(12)
    'FooterLength = 5
    'If LineCount + HeaderLength + FooterLength > 67 Then MsgBox PageCount
    'If LineCount + HeaderLength + FooterLength < 65 Then MsgBox PageCount
End Sub
Private Sub LedgerDetail()
    Dim j As Integer

    If (LedgerDate <> PrintedDate) Or NewPage Then
    Print #1, Format$(LedgerDetailRS!acn_date, "dd/mm/yyyy"); Spc(2);
    PrintedDate = Format$(LedgerDetailRS!acn_date, "dd/mm/yyyy")
    NewPage = False
    Else
    Print #1, Tab(13);
    End If
   
    j = PrintNarration(LedgerDetailRS!particular, 1)
    'Print #1, Tab(13);  Mid(LedgerDetailRS!particular, 1, 36);
    Print #1, Tab(51);
    Print #1, ZeroSup(LedgerDetailRS!debit);
    Print #1, Spc(2);
    Print #1, ZeroSup(LedgerDetailRS!credit)
    Print #1,
    LineCount = LineCount + 1
    
    DebitBalance = DebitBalance + LedgerDetailRS!debit
    CreditBalance = CreditBalance + LedgerDetailRS!credit
    HeadBalance = CreditBalance - DebitBalance
End Sub
Private Sub LedgerSummary()
        If NumberOfRecords > 0 Then
            Print #1, Tab(51); String(30, "-")
            Print #1, Tab(51); Format(Format(DebitBalance, "0.00"), "@@@@@@@@@@@@@@");
            Print #1, Spc(2); Format(Format(CreditBalance, "0.00"), "@@@@@@@@@@@@@@")
            Print #1, Tab(51); String(30, "-")
        End If
        
        Print #1,
        Print #1, Tab(23); "Closing Balance";
        If HeadBalance >= 0 Then
            Print #1, Tab(67); Format(Format(HeadBalance, "0.00"), "@@@@@@@@@@@@@@")
        Else
            Print #1, Tab(51); Format(Format(-HeadBalance, "0.00"), "@@@@@@@@@@@@@@")
        End If
        Print #1, String(PageWidth, "-")
        Print #1,
        Print #1, Chr(12)
        'summaryLength = 10

End Sub

Private Sub PrepareLedger()
    Dim i As Long
    PageLength = 66
    PageWidth = 80
    HeaderLength = 11
    FooterLength = 5
    DetailLength = PageLength - (HeaderLength + FooterLength + 1)
    PageCount = 0
    LineCount = DetailLength + 1
    cmdPrint.Enabled = False
    HeadBalance = 0
    DebitBalance = 0
    CreditBalance = 0
    NewDate = False
    LedgerDate = FromDate - 1
    i = 0
    
    Set LedgerDetailRS = New Recordset
    LedgerDetailRS.Open "select BALANCETYP,BALANCE FROM " & MasterTable & " where acnumber = " & SelectedRecord _
     , db, adOpenStatic, adLockReadOnly, adCmdText
    If LedgerDetailRS.EOF = False And LedgerDetailRS.BOF = False Then
        If LedgerDetailRS!balancetyp = "C" Then
            HeadBalance = LedgerDetailRS!Balance
        Else
            HeadBalance = -(LedgerDetailRS!Balance)
        End If
        
    Else
        HeadBalance = 0
    End If
    LedgerDetailRS.Close
    Set LedgerDetailRS = New Recordset
    LedgerDetailRS.Open "select sum(credit-debit) FROM " & TransactionTable & " where " _
     & "acn_date between {" & Format(StartDate, "mm/dd/yyyy") & "} and {" & Format(FromDate - 1, "mm/dd/yyyy") & "} " _
     & "and acnumber = " & SelectedRecord, db, adOpenStatic, adLockReadOnly, adCmdText
     
     With LedgerDetailRS
        If Not (.BOF = True And .EOF = True) Then
            HeadBalance = HeadBalance + .Fields(0).Value
        End If
        .Close
    End With
    
    Set LedgerDetailRS = New Recordset
    
    LedgerDetailRS.Open "select acnumber,acn_date,entry_id,particular,debit," _
     & "credit FROM " & TransactionTable & " where acnumber = " & SelectedRecord & " and " _
     & "acn_date between {" & Format(FromDate, "mm/dd/yyyy") & "} and {" & Format(ToDate, "mm/dd/yyyy") & "} order by " _
     & "acn_date,entry_id ", db, adOpenStatic, adLockReadOnly, adCmdText
    
    NumberOfRecords = LedgerDetailRS.RecordCount
    
    
        If NumberOfRecords > 0 Then
            LedgerDetailRS.MoveFirst
            LedgerDate = LedgerDetailRS!acn_date
        End If
        
        DebitBalance = IIf(HeadBalance < 0, -HeadBalance, 0)
        CreditBalance = IIf(HeadBalance > 0, HeadBalance, 0)
    
    Open "c:\vbprog\vba\rpt\Ledger.txt" For Output As #1
    
    Do While i < NumberOfRecords
        
        If LineCount > DetailLength Then
            If Not ExcessNarration Then
                If PageCount > 0 Then
                    LedgerFooter
                End If
            End If
            If Not ExcessNarration Then
                PageCount = PageCount + 1
                LedgerHeader
            End If
        End If
        
            
            If Format(LedgerDetailRS!acn_date, "dd/mm/yyyy") <> Format(LedgerDate, "dd/mm/yyyy") Then
                NewDate = True
                LedgerDate = LedgerDetailRS!acn_date
               ' HeadBalance = CreditBalance - DebitBalance
            Else
                NewDate = False
            End If

        If LineCount < DetailLength Then
            With LedgerDetailRS
                If Not .EOF Then
                    LedgerDetail
                    .MoveNext
                    i = i + 1
                End If
            End With
        Else
            LineCount = DetailLength + 1
        End If
    
    Loop
    
    If NumberOfRecords > 0 Then
        If DetailLength - LineCount > 10 Then
            LedgerSummary
        Else
            LedgerFooter
            LedgerHeader
            LedgerSummary
        End If
    Else
        PageCount = 1
        LedgerHeader
        LedgerSummary
    End If
        
    'Print #1, Chr(12)
    Close #1
    
End Sub

Private Sub cmdprint_Click()
    cmdPrint.Enabled = False
    PrintText (RichTextBox1.Text)
    ContinueProcess = True
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
    If SpacePosition > 36 Then
        Narration1 = Mid(Gstring, 1, 36)
        Gstring = Mid(Gstring, 36 + 1)
    ElseIf Len(Gstring) <= 36 Then
        Narration1 = Mid(Gstring, 1, 36)
        Gstring = Mid(Gstring, 36 + 1)
    ElseIf Len(Gstring) > 36 And SpacePosition = 0 Then
        Narration1 = Mid(Gstring, 1, 36)
        Gstring = Mid(Gstring, 36 + 1)
    End If
Do While True
    SpacePosition = InStr(Gstring, " ")
    If SpacePosition = 0 Then Exit Do
    If Len(Narration1) + SpacePosition <= 36 + 1 Then
        Narration1 = Narration1 + Mid(Gstring, 1, SpacePosition)
        Gstring = Mid(Gstring, SpacePosition + 1)
    Else
        Exit Do
    End If
Loop
    'Debug.Print Trim(Narration1)
    Print #FileNumber, Tab(13); Trim(Narration1);
    Narration1 = ""
    Ncount = Ncount + 1
    LineCount = LineCount + 1
    If Len(Gstring) = 0 Then
        ExcessNarration = False
        Exit Do
    End If
    If LineCount > DetailLength Then
        Print #FileNumber,
        LedgerFooter
        Ncount = 0
        PageCount = PageCount + 1
        ExcessNarration = True
        LedgerHeader
    Else
        ExcessNarration = False
    End If
Loop
PrintNarration = Ncount
End Function
