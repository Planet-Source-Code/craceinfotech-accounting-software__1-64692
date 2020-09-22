VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Daybook1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daybook"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   9510
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   300
      Left            =   4995
      TabIndex        =   2
      Top             =   5820
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   300
      Left            =   3315
      TabIndex        =   1
      Top             =   5820
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   5295
      Left            =   350
      TabIndex        =   0
      Top             =   180
      Width           =   8810
      _ExtentX        =   15558
      _ExtentY        =   9340
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmDaybook1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Daybook1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LineCount As Integer
Dim HeaderLength As Integer
Dim FooterLength As Integer
Dim PageLength As Integer
Dim PageWidth As Integer
Dim PageCount As Integer
Dim CashInHand As Currency
Dim DaybookDetailRS As Recordset
Dim DaybookOpeningRS As Recordset
Dim NumberOfRecords As Long
Dim ExcessNarration As Boolean
Dim DebitBalance As Currency
Dim CreditBalance As Currency
Dim NewDate As Boolean
Dim DaybookDate As Date


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
'Stop
'testnarr
'End
Screen.MousePointer = vbHourglass
PrepareDaybook
If NumberOfRecords > 0 Then
    RichTextBox1.LoadFile "c:\back\new_acn\vba\daybook.txt", rtfText
Else
    Unload Me
End If
Screen.MousePointer = vbDefault

End Sub

Private Sub DaybookHeader()
    'IMPORTANT:10+2+36+2+14+2+14=80
    Print #1,
    Print #1, PADC("LOVELY OFFSET PRINTERS PRIVATE LTD - SIVAKASI.", PageWidth); Tab(PageWidth - (7 + 4)); "Page : "; Format(PageCount, "@@@@")
    Print #1, PADC("DAY BOOK", PageWidth)
    Print #1, String(PageWidth, "-")
    Print #1, "Date"; Tab(13); "Particulars"; Tab(51); "    Debit"; Tab(67); "    Credit"
    Print #1, String(PageWidth, "-")
    If PageCount > 1 Then
        Print #1, Tab(45); "b/f"; Tab(51); Format(Format(DebitBalance, "0.00"), "@@@@@@@@@@@@@@");
        Print #1, Spc(2); Format(Format(CreditBalance, "0.00"), "@@@@@@@@@@@@@@")
        Print #1,
    End If


'HeaderLength = 7
End Sub
Private Sub DaybookFooter()
    Print #1,
    Print #1, Tab(45); "c/d"; Tab(51); Format(Format(DebitBalance, "0.00"), "@@@@@@@@@@@@@@");
    Print #1, Spc(2); Format(Format(CreditBalance, "0.00"), "@@@@@@@@@@@@@@")
    Print #1, String(PageWidth, "-")
    Print #1,
    Print #1,
    Print #1, Chr(12)
    'FooterLength = 5
    'If LineCount + FooterLength > 67 Then MsgBox PageCount
    'If LineCount + FooterLength < 65 Then MsgBox PageCount
End Sub
Private Sub DaybookDetail()
    Dim j As Integer
    If Format(DaybookDetailRS!acn_date, "dd/mm/yyyy") <> Format(DaybookDate, "dd/mm/yyyy") Then
        NewDate = True
        DaybookDate = DaybookDetailRS!acn_date
        CashInHand = CreditBalance - DebitBalance
        
                
    Else
        NewDate = False
    End If
    
    If NewDate Then
        If PageCount = 1 Then
        Print #1,
        Print #1, Tab(36); "Cash In Hand";
        If CashInHand >= 0 Then
            Print #1, Tab(67); Format(Format(CashInHand, "0.00"), "@@@@@@@@@@@@@@")
        Else
            Print #1, Tab(51); Format(Format(-CashInHand, "0.00"), "@@@@@@@@@@@@@@")
        End If
        Print #1,
        LineCount = LineCount + 3
        
        ElseIf LineCount < (PageLength - (FooterLength + 1 + 5)) Then
        
        'Print #1,
        Print #1, Tab(51); String(30, "-")
        Print #1, Tab(51); Format(Format(DebitBalance, "0.00"), "@@@@@@@@@@@@@@");
        Print #1, Spc(2); Format(Format(CreditBalance, "0.00"), "@@@@@@@@@@@@@@")
        Print #1, Tab(51); String(30, "-")
        
        CreditBalance = IIf(CashInHand > 0, CashInHand, 0)
        DebitBalance = IIf(CashInHand < 0, -CashInHand, 0)

        Print #1,
        Print #1, Tab(36); "Cash In Hand";
        If CashInHand >= 0 Then
            Print #1, Tab(67); Format(Format(CashInHand, "0.00"), "@@@@@@@@@@@@@@")
        Else
            Print #1, Tab(51); Format(Format(-CashInHand, "0.00"), "@@@@@@@@@@@@@@")
        End If
        Print #1,
        LineCount = LineCount + 5
        Else
        Print #1,
        LineCount = LineCount + 5
        Exit Sub
        End If
    End If
    
        
        
    Print #1, Format$(DaybookDetailRS!acn_date, "dd/mm/yyyy");
    Print #1, Spc(2);
    Print #1, Mid(DaybookDetailRS!ACTITLE, 1, 36)
    LineCount = LineCount + 1
    
    j = PrintNarration(DaybookDetailRS!particular, 1)
    Print #1, Tab(51);
    Print #1, ZeroSup(DaybookDetailRS!debit);
    Print #1, Spc(2);
    Print #1, ZeroSup(DaybookDetailRS!credit)
    DebitBalance = DebitBalance + DaybookDetailRS!debit
    CreditBalance = CreditBalance + DaybookDetailRS!credit
    'LineCount = LineCount + 1
    
End Sub
Private Sub DaybookSummary()

End Sub

Private Sub PrepareDaybook()
    Dim i As Long
    PageLength = 66
    PageWidth = 80
    HeaderLength = 8
    FooterLength = 5
    PageCount = 0
    'LineCount = 1
    cmdPrint.Enabled = False
    CashInHand = 0
    DebitBalance = 0
    CreditBalance = 0
    NewDate = False
    DaybookDate = FromDate - 1
    
    Set DaybookDetailRS = New Recordset
    DaybookDetailRS.Open "select sum(credit-debit) from entries where " _
     & "acn_date between {" & Format(StartDate, "mm/dd/yyyy") & "} and {" & Format(FromDate - 1, "mm/dd/yyyy") & "} ", _
     db, adOpenStatic, adLockReadOnly, adCmdText
     
     With DaybookDetailRS
        If Not (.BOF = True And .EOF = True) Then
            CashInHand = .Fields(0).Value
        Else
            CashInHand = 0
        End If
        .Close
    End With
    
    Set DaybookDetailRS = New Recordset
    
    DaybookDetailRS.Open "select acn_date,entry_id,actitle,particular,debit," _
     & "credit from entries e,master m where e.acnumber=m.acnumber and " _
     & "acn_date between {" & Format(FromDate, "mm/dd/yyyy") & "} and {" & Format(ToDate, "mm/dd/yyyy") & "} order by " _
     & "acn_date,entry_id ", db, adOpenStatic, adLockReadOnly, adCmdText
    
    NumberOfRecords = DaybookDetailRS.RecordCount
    
    If NumberOfRecords = 0 Then
        MsgBox "No Records"
        Exit Sub
    End If
    
    DaybookDetailRS.MoveFirst
    Open "c:\back\new_acn\vba\daybook.txt" For Output As #1
    
    Do While i <= NumberOfRecords
        
        If Not ExcessNarration Then
            PageCount = PageCount + 1
            DaybookHeader
            LineCount = HeaderLength
        End If
        

        Do While LineCount < (PageLength - (FooterLength + 1))
            i = i + 1
            With DaybookDetailRS
                If Not .EOF Then
                    DaybookDetail
                    .MoveNext
                    
                Else
                'LineCount = (PageLength - (FooterLength + 1))
                'i = i + 1
                Exit Do
                End If
            End With
            'LineCount = LineCount + 1
            'i = i + 1
        Loop
        If Not ExcessNarration Then
            DaybookFooter
        End If
        
        
    Loop
    
    'Print #1, Chr(12)
    Close #1
    cmdPrint.Enabled = True
End Sub

Private Sub cmdprint_Click()
    cmdPrint.Enabled = False
    'sWrittenData = RichTextBox1.Text
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
    If LineCount >= (PageLength - (FooterLength + 1)) Then
        DaybookFooter
        Ncount = 0
        PageCount = PageCount + 1
        ExcessNarration = True
        DaybookHeader
        LineCount = HeaderLength
    Else
        ExcessNarration = False
    End If
Loop
PrintNarration = Ncount
End Function
