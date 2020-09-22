VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Ledger1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ledger1"
   ClientHeight    =   7020
   ClientLeft      =   -570
   ClientTop       =   1005
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   10875
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameLedger 
      Height          =   6735
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   10575
      Begin VB.OptionButton OptIndex 
         Caption         =   "Index"
         Height          =   435
         Left            =   9060
         TabIndex        =   5
         Top             =   6120
         Width           =   735
      End
      Begin VB.OptionButton OptLedger 
         Caption         =   "Ledger"
         Height          =   315
         Left            =   7800
         TabIndex        =   4
         Top             =   6180
         Width           =   915
      End
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
         Left            =   0
         TabIndex        =   3
         Top             =   120
         Width           =   10245
         _ExtentX        =   18071
         _ExtentY        =   10292
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"frmLedger1.frx":0000
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
Attribute VB_Name = "Ledger1"
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
Dim PageNumber As Integer
Dim HeadBalance As Currency
Dim LedgerMasterRS As Recordset
Dim LedgerDetailRS As Recordset
Dim NumberOfRecords As Long
Dim ExcessNarration As Boolean
Dim DebitBalance As Currency
Dim CreditBalance As Currency
Dim NewDate As Boolean
Dim LedgerDate As Date
Dim NewPage As Boolean
Dim NewLedger As Boolean
Dim PrintedDate As Date
Dim IsOpen As Boolean
Dim IsTrans As Boolean

Dim iLineCount As Integer
Dim iHeaderLength As Integer
Dim iFooterLength As Integer
Dim iDetailLength As Integer
Dim iPageNumber As Integer
Dim SerialNumber As Integer
Dim NewIndex As Boolean

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbHourglass
    Open "c:\vbprog\vba\rpt\Ledger.txt" For Output As #1
    Open "c:\vbprog\vba\rpt\index.txt" For Output As #2
    cmdPrint.Enabled = False
    SelectMaster
    Close #1
    Close #2
    Screen.MousePointer = vbDefault
    'RichTextBox1.LoadFile "c:\vbprog\vba\rpt\Ledger.txt", rtfText
    'RichTextBox1.LoadFile "c:\vbprog\vba\rpt\index.txt", rtfText
    'cmdPrint.Enabled = True
    'RichTextBox1.SetFocus
    OptLedger.Value = True
    'Screen.MousePointer = vbDefault
End Sub

Private Sub LedgerHeader()
    'IMPORTANT:10+2+36+2+14+2+14=80
    Print #1,
    Print #1, PADC(CompanyName, PageWidth); Tab(PageWidth - (7 + 4)); "Page : "; Format(PageNumber, "@@@@")
    Print #1, PADC("STATEMENT OF ACCOUNTS", PageWidth)
    Print #1,
    Print #1, PADC(SelectedHead, PageWidth)
    Print #1, PADC("(From " & Format(FromDate, "dd/mm/yyyy") & " to " & Format(ToDate, "dd/mm/yyyy") & ")", PageWidth)
    Print #1, String(PageWidth, "-")
    Print #1, "Date"; Tab(13); "Particulars"; Tab(51); "    Debit"; Tab(67); "    Credit"
    Print #1, String(PageWidth, "-")
    
    If NewLedger = True Then
        Print #1, Tab(23); "Opening Balance";
        If HeadBalance >= 0 Then
            Print #1, Tab(67); Format(Format(HeadBalance, "0.00"), "@@@@@@@@@@@@@@")
        Else
            Print #1, Tab(51); Format(Format(-HeadBalance, "0.00"), "@@@@@@@@@@@@@@")
        End If
        NewLedger = False
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
    'If LineCount + HeaderLength + FooterLength > 67 Then MsgBox pagenumber
    'If LineCount + HeaderLength + FooterLength < 65 Then MsgBox pagenumber
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

    LineCount = DetailLength + 1
    

    DebitBalance = IIf(HeadBalance < 0, -HeadBalance, 0)
    CreditBalance = IIf(HeadBalance > 0, HeadBalance, 0)
    
    NewDate = False
    LedgerDate = FromDate - 1
    i = 0
    
    Set LedgerDetailRS = New Recordset
    
    LedgerDetailRS.Open "select acnumber,acn_date,entry_id,particular,debit," _
     & "credit FROM " & TransactionTable & " where acnumber = " & SelectedRecord _
     & " order by " _
     & "acn_date,entry_id ", db, adOpenStatic, adLockReadOnly, adCmdText
    
    NumberOfRecords = LedgerDetailRS.RecordCount
    
        If NumberOfRecords > 0 Then
            LedgerDetailRS.MoveFirst
            LedgerDate = LedgerDetailRS!acn_date
        Else
            PageNumber = PageNumber + 1
            LedgerHeader
            LedgerSummary
        End If
        
    Do While i < NumberOfRecords
        
        If LineCount > DetailLength Then
            If Not ExcessNarration Then
                If NewLedger = False Then
                    LedgerFooter
                End If
            End If
            If Not ExcessNarration Then
                PageNumber = PageNumber + 1
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
            PageNumber = PageNumber + 1
            LedgerHeader
            LedgerSummary
        End If
    End If
    
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
        PageNumber = PageNumber + 1
        ExcessNarration = True
        LedgerHeader
    Else
        ExcessNarration = False
    End If
Loop
PrintNarration = Ncount
End Function

Public Sub SelectMaster()
    
    PageNumber = 0
    PageLength = 66
    PageWidth = 80
    HeaderLength = 11
    FooterLength = 5
    DetailLength = PageLength - (HeaderLength + FooterLength + 1)
    
    NewIndex = True
    SerialNumber = 0
    iPageNumber = 1
    iHeaderLength = 10
    iFooterLength = 4
    iDetailLength = PageLength - (iHeaderLength + iFooterLength + 1)
    
    iLineCount = iDetailLength + 1
    
    
    Set LedgerMasterRS = New Recordset
    LedgerMasterRS.Open "select ACNUMBER,ACTITLE,BALANCETYP,BALANCE FROM " & MasterTable & " where acnumber<>1 order by actitle " _
     , db, adOpenStatic, adLockReadOnly, adCmdText

    If LedgerMasterRS.EOF = True And LedgerMasterRS.BOF = True Then
        MsgBox "No Master Records"
        'Exit Sub
        Unload Me
    End If

    With LedgerMasterRS
    .MoveFirst
    Do While Not .EOF
        SelectedRecord = .Fields("acnumber").Value
        SelectedHead = Trim(.Fields("actitle").Value)

        If .Fields("balancetyp").Value = "C" Then
            HeadBalance = .Fields("balance").Value
        Else
            HeadBalance = .Fields("balance").Value * -1
        End If

        If .Fields("balance").Value > 0 Then
            IsTrans = True
        Else
            Set LedgerDetailRS = New Recordset
            LedgerDetailRS.Open "select count(*) FROM " & TransactionTable & " where " _
            & " acnumber = " & SelectedRecord, db, adOpenStatic, adLockReadOnly, adCmdText
        
            With LedgerDetailRS
                If .Fields(0).Value = 0 Then
                    .Close
                    'LedgerMasterRS.MoveNext
                    'Loop
                    IsTrans = False
                Else
                    IsTrans = True
                    .Close
                End If
                
            End With
        
        End If
        
        If IsTrans = True Then
            NewLedger = True
            PrepareIndex
            PrepareLedger
            IsTrans = False
        End If
        
        'If Not .EOF Then
            .MoveNext
        'End If
    Loop
    End With
    
    LedgerMasterRS.Close
    
        If iDetailLength - iLineCount > 10 Then
            IndexSummary
        Else
            IndexFooter
            iPageNumber = iPageNumber + 1
            IndexHeader
            IndexSummary
        End If

End Sub
'**********
Private Sub PrepareIndex()
    Dim j As Long
    
    j = 0
   
    'iLineCount = iDetailLength + 1
    
    If iLineCount > iDetailLength Then
        If NewIndex = False Then
            IndexFooter
            iPageNumber = iPageNumber + 1
        End If
    
        'iPageNumber = iPageNumber + 1
        IndexHeader
        iLineCount = 0
    End If
        
    IndexDetail
        

    
End Sub

Private Sub IndexHeader()
    'IMPORTANT:10+2+36+2+14+2+14=80
    Print #2,
    Print #2, PADC(CompanyName, PageWidth); Tab(PageWidth - (7 + 4)); "Page : "; Format(iPageNumber, "@@@@")
    Print #2, PADC("INDEX", PageWidth)
    Print #2,
    
    Print #2, String(PageWidth, "-")
    Print #2, "S.No"; Tab(13); "Account Head"; Tab(61); "Page No"
    '; Tab(67); "    Credit"
    Print #2, String(PageWidth, "-")
    Print #2,
    
    
'    iLineCount = 0
    NewIndex = False

'HeaderLength = 8
End Sub
Private Sub IndexFooter()
    Print #2,
    Print #2, String(PageWidth, "-")
    Print #2,
    Print #2, Chr(12)
    'FooterLength = 4
End Sub
Private Sub IndexDetail()
        
    Print #2, Format(SerialNumber + 1, "@@@@") + "."; Spc(2);
    Print #2, Tab(13); Mid(SelectedHead, 1, 40);
    Print #2, Tab(61);
    Print #2, Format(PageNumber + 1, "@@@@")
    Print #2,
    
    iLineCount = iLineCount + 2
    SerialNumber = SerialNumber + 1

End Sub
Private Sub IndexSummary()
        Print #2,
        Print #2,
        Print #2, String(PageWidth, "=")
        Print #2,
        Print #2, Chr(12)
        'summaryLength = 5
End Sub


Private Sub OptLedger_Click()
    cmdPrint.Enabled = False
    RichTextBox1.LoadFile "c:\vbprog\vba\rpt\Ledger.txt", rtfText
    'RichTextBox1.LoadFile "c:\vbprog\vba\rpt\index.txt", rtfText
    cmdPrint.Enabled = True
    RichTextBox1.SetFocus
End Sub

Private Sub OptIndex_Click()
    cmdPrint.Enabled = False
    'RichTextBox1.LoadFile "c:\vbprog\vba\rpt\Ledger.txt", rtfText
    RichTextBox1.LoadFile "c:\vbprog\vba\rpt\index.txt", rtfText
    cmdPrint.Enabled = True
    RichTextBox1.SetFocus
End Sub
