VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form AllLedger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "All Heads Ledger "
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
   Begin VB.Frame FrameLedger 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   10515
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   300
         Left            =   3990
         TabIndex        =   2
         Top             =   6240
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   5670
         TabIndex        =   1
         Top             =   6240
         Width           =   1095
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   5835
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   10245
         _ExtentX        =   18071
         _ExtentY        =   10292
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"frmAllLedger.frx":0000
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
Attribute VB_Name = "AllLedger"
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
'Dim HeadBalance As Currency
Dim TrialRS As Recordset
Dim MasterRS As Recordset
Dim TransactRS As Recordset
Dim NumberOfRecords As Long
Dim NumberOfRecords1 As Long
Dim DebitBalance As Currency
Dim CreditBalance As Currency
Dim NewPage As Boolean

'FROM LEDGER FORM START
Dim HeadBalance As Currency
Dim LedgerDetailRS As Recordset
Dim ExcessNarration As Boolean
Dim NewDate As Boolean
Dim LedgerDate As Date
'Dim NewPage As Boolean
Dim PrintedDate As Date
'FROM LEDGER FORM END



Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    PrepareAllLedger
    RichTextBox1.LoadFile "c:\vbprog\vba\rpt\AllLedger.txt", rtfText
    cmdPrint.Enabled = True
    RichTextBox1.SetFocus
    Screen.MousePointer = vbDefault
End Sub

Private Sub TrialHeader()
    'IMPORTANT:48+2+14+2+14=80
    Print #1,
    Print #1, PADC(CompanyName, PageWidth); Tab(PageWidth - (7 + 4)); "Page : "; Format(PageCount, "@@@@")
    Print #1, PADC("TRIAL BALANCE", PageWidth)
    'Print #1,
    Print #1, PADC("As on " & Format(ToDate, "dd/mm/yyyy"), PageWidth)
    Print #1, String(PageWidth, "-")
    Print #1, "Particulars"; Tab(51); Spc(9); "Debit"; Tab(67); Spc(8); "Credit"
    Print #1, String(PageWidth, "-")
    
    If PageCount > 1 Then
        Print #1, Tab(45); "b/f"; Tab(51); Format(Format(DebitBalance, "0.00"), "@@@@@@@@@@@@@@");
        Print #1, Spc(2); Format(Format(CreditBalance, "0.00"), "@@@@@@@@@@@@@@")
        Print #1,
    End If
    
    LineCount = 0
    NewPage = True

'HeaderLength = 9
End Sub
Private Sub TrialFooter()
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
Private Sub TrialDetail()
Print #1, Mid(TrialRS!ACTITLE, 1, 48);
If TrialRS!head_bal < 0 Then
Print #1, Tab(51); ZeroSup(-1 * TrialRS!head_bal)
DebitBalance = DebitBalance + (-1 * TrialRS!head_bal)
Else
Print #1, Tab(67); ZeroSup(TrialRS!head_bal)
CreditBalance = CreditBalance + TrialRS!head_bal
End If
LineCount = LineCount + 1
End Sub
Private Sub TrialSummary()
        If NumberOfRecords > 0 Then
            Print #1, Tab(51); String(30, "-")
            Print #1, Tab(51); Format(Format(DebitBalance, "0.00"), "@@@@@@@@@@@@@@");
            Print #1, Spc(2); Format(Format(CreditBalance, "0.00"), "@@@@@@@@@@@@@@")
            Print #1, Tab(51); String(30, "-")
        End If
        
        Print #1,
        If DebitBalance <> CreditBalance Then
        
        Print #1, Tab(23); "Difference";
        If CreditBalance - DebitBalance > 0 Then
            Print #1, Tab(67); Format(Format(CreditBalance - DebitBalance, "0.00"), "@@@@@@@@@@@@@@")
        Else
            Print #1, Tab(51); Format(Format(DebitBalance - CreditBalance, "0.00"), "@@@@@@@@@@@@@@")
        End If
        End If
        Print #1,
        Print #1, String(PageWidth, "-")
        Print #1, Chr(12)
        'summaryLength = 10

End Sub

Private Sub PrepareAllLedger()
    Dim i As Long
    Dim j As Long
    Dim CashTrans As Currency
    PageLength = 66
    PageWidth = 80
    HeaderLength = 9
    FooterLength = 5
    DetailLength = PageLength - (HeaderLength + FooterLength + 1)
    PageCount = 0
    LineCount = DetailLength + 1
    cmdPrint.Enabled = False
    CashTrans = 0
    DebitBalance = 0
    CreditBalance = 0
    i = 0
    j = 0
        
    Set MasterRS = New Recordset
    MasterRS.Open "select acnumber,actitle,balancetyp,balance,fnlrptcode,fnlrptposi FROM " & MasterTable & " order by actitle", _
        db, adOpenStatic, adLockReadOnly, adCmdText
    
    Set TrialRS = New Recordset
    TrialRS.Fields.Append "ACNUMBER", adInteger, , adFldKeyColumn
    TrialRS.Fields.Append "ACTITLE", adChar, 50
    TrialRS.Fields.Append "BALANCETYP", adChar, 1
    TrialRS.Fields.Append "BALANCE", adCurrency
    TrialRS.Fields.Append "FNLRPTCODE", adInteger
    TrialRS.Fields.Append "FNLRPTPOSI", adInteger
    TrialRS.Fields.Append "TRANSACT", adCurrency
    TrialRS.Fields.Append "HEAD_BAL", adCurrency
    TrialRS.Fields.Append "NO_NEED", adInteger
    
    TrialRS.CursorLocation = adUseClient
    TrialRS.CursorType = adOpenStatic
    TrialRS.LockType = adLockOptimistic
    TrialRS.Open
    
    With MasterRS
        .MoveFirst
        Do While Not .EOF
            TrialRS.AddNew
            TrialRS!ACNUMBER = MasterRS("acnumber").Value
            TrialRS!ACTITLE = Trim(MasterRS("actitle").Value)
            TrialRS!balancetyp = MasterRS("balancetyp").Value
            TrialRS!Balance = IIf(MasterRS("balancetyp").Value = "C", MasterRS("balance").Value, -1 * MasterRS("balance").Value)
            TrialRS!fnlrptcode = MasterRS("fnlrptcode").Value
            TrialRS!fnlrptposi = MasterRS("fnlrptposi").Value
            TrialRS!TRANSACT = 0
            TrialRS!head_bal = 0
            TrialRS!no_need = 0
            TrialRS.Update
            .MoveNext
        Loop
    End With

    
    Set TransactRS = New Recordset
    TransactRS.Open "select ACNUMBER,SUM(CREDIT-DEBIT) as TRANSACT FROM " & TransactionTable & " group by ACNUMBER", _
        db, adOpenStatic, adLockReadOnly, adCmdText
    
    With TransactRS
    If .BOF = False And .EOF = False Then
        .MoveFirst
        Do While Not .EOF
            TrialRS.Find "acnumber = " & TransactRS("acnumber").Value, , adSearchForward, adBookmarkFirst
            If Not TrialRS.EOF Then
                TrialRS!TRANSACT = TransactRS("transact").Value
                TrialRS.Update
                CashTrans = CashTrans + TransactRS("transact").Value
            End If
            .MoveNext
        Loop
    End If
    End With
    
    With TrialRS
    If .BOF = False And .EOF = False Then
        .MoveFirst
        .Find "actitle='CASH'", , adSearchForward
        If Not .EOF Then
            ' to suppress cash ledger
            'TrialRS!TRANSACT = -1 * CashTrans
            '.Update
            .Delete
        End If
        'ERROR POINT
        .MoveFirst
        Do While Not .EOF
                TrialRS!head_bal = TrialRS("transact").Value + TrialRS("balance").Value
                .Update
            .MoveNext
        Loop
    End If
    End With

    TrialRS.MoveFirst

    TrialRS.Filter = "transact <> 0 or balance <> 0"
    NumberOfRecords = TrialRS.RecordCount
    'TrialRS.MoveFirst
    If TrialRS.EOF = False And TrialRS.BOF = False Then
        TrialRS.MoveFirst
    Else
        MsgBox "Not Sufficient Entries"
        Exit Sub
    End If

    j = 0
    LineCount = DetailLength + 1
    DebitBalance = 0
    CreditBalance = 0
        
        If NumberOfRecords > 0 Then
            TrialRS.MoveFirst
        End If
        

    
    Open "c:\vbprog\vba\rpt\AllLedger.txt" For Output As #1
    
    Do While j < NumberOfRecords
        SelectedRecord = TrialRS("acnumber").Value
        PrepareLedger
        TrialRS.MoveNext
        j = j + 1
        Loop
        ' OLD TRIAL CODE
'        If LineCount > DetailLength Then
'                If PageCount > 0 Then
'                    TrialFooter
'                End If
'                PageCount = PageCount + 1
'                TrialHeader
'        End If
'
'        If LineCount < DetailLength Then
'            With TrialRS
'                If Not .EOF Then
'                    TrialDetail
'                    .MoveNext
'                    i = i + 1
'                End If
'            End With
'        Else
'            LineCount = DetailLength + 1
'        End If
'
'    Loop
'
'
'        If DetailLength - LineCount > 10 Then
'            TrialSummary
'        Else
'            TrialFooter
'            TrialHeader
'            TrialSummary
'        End If
'

    Close #1
    'OLD TRIAL CODE
    
End Sub

Private Sub cmdprint_Click()
    cmdPrint.Enabled = False
    PrintText (RichTextBox1.Text)
    ContinueProcess = True
End Sub

'*****************************************LEDGER FORM COPY START


Private Sub cmdClose1_Click()
    Unload Me
End Sub

Private Sub Form1_Activate()
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
    
    NumberOfRecords1 = LedgerDetailRS.RecordCount
    
    
        If NumberOfRecords1 > 0 Then
            LedgerDetailRS.MoveFirst
            LedgerDate = LedgerDetailRS!acn_date
        End If
        
        DebitBalance = IIf(HeadBalance < 0, -HeadBalance, 0)
        CreditBalance = IIf(HeadBalance > 0, HeadBalance, 0)
    
'    Open "c:\vbprog\vba\rpt\Ledger.txt" For Output As #1
    
    Do While i < NumberOfRecords1
        
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
    
    If NumberOfRecords1 > 0 Then
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
 '   Close #1
    
End Sub

Private Sub cmdprint1_Click()
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
'*******************************LEDGER FORM COPY END
