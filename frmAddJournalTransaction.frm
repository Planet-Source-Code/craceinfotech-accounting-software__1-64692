VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form AddJournalTransaction 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Journal Transaction"
   ClientHeight    =   5295
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   8985
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCash 
      CausesValidation=   0   'False
      DataField       =   "entry_id"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   420
      Width           =   2235
   End
   Begin VB.CommandButton cmdIncrease 
      Caption         =   "&+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7635
      TabIndex        =   1
      Top             =   435
      Width           =   495
   End
   Begin VB.CommandButton cmdDecrease 
      Caption         =   "&-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   2
      Top             =   420
      Width           =   495
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H80000008&
      Height          =   1875
      Left            =   75
      TabIndex        =   16
      Top             =   960
      Width           =   8820
      Begin VB.ComboBox cboHead 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   180
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   840
         Width           =   4935
      End
      Begin VB.TextBox txtCredit 
         DataField       =   "credit"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5246
         TabIndex        =   4
         Text            =   "1234567890"
         Top             =   840
         Width           =   1635
      End
      Begin VB.TextBox txtDebit 
         DataField       =   "debit"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   7012
         TabIndex        =   5
         Text            =   "1234567890"
         Top             =   840
         Width           =   1635
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "CREDIT"
         Height          =   195
         Index           =   3
         Left            =   5763
         TabIndex        =   19
         Top             =   240
         Width           =   600
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "DEBIT"
         Height          =   195
         Index           =   6
         Left            =   7589
         TabIndex        =   18
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "PARTICULARS"
         Height          =   195
         Index           =   7
         Left            =   2085
         TabIndex        =   17
         Top             =   240
         Width           =   1125
      End
      Begin VB.Line Line1 
         BorderStyle     =   6  'Inside Solid
         X1              =   5
         X2              =   8815
         Y1              =   540
         Y2              =   540
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   390
      Left            =   6742
      TabIndex        =   8
      Top             =   4020
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   6742
      TabIndex        =   7
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox txtParticular 
      DataField       =   "particular"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   262
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "frmAddJournalTransaction.frx":0000
      Top             =   3120
      Width           =   4935
   End
   Begin VB.TextBox txtEntryNo 
      CausesValidation=   0   'False
      DataField       =   "entry_id"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3189
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   420
      Width           =   1275
   End
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   0
      ScaleHeight     =   390
      ScaleWidth      =   8985
      TabIndex        =   14
      Top             =   4905
      Width           =   8985
   End
   Begin MSMask.MaskEdBox mskAcnDate 
      Height          =   390
      Left            =   5433
      TabIndex        =   0
      Top             =   420
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   688
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "No  "
      Height          =   195
      Index           =   0
      Left            =   2622
      TabIndex        =   15
      Top             =   518
      Width           =   300
   End
   Begin VB.Label lblLabels 
      Caption         =   "CREDIT"
      Height          =   255
      Index           =   5
      Left            =   6360
      TabIndex        =   13
      Top             =   1500
      Width           =   1635
   End
   Begin VB.Label lblLabels 
      Caption         =   "DEBIT"
      Height          =   255
      Index           =   4
      Left            =   4320
      TabIndex        =   12
      Top             =   1500
      Width           =   1635
   End
   Begin VB.Label lblLabels 
      Caption         =   "PARTICULARS"
      Height          =   255
      Index           =   2
      Left            =   540
      TabIndex        =   11
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "DATE"
      Height          =   195
      Index           =   1
      Left            =   4731
      TabIndex        =   10
      Top             =   518
      Width           =   435
   End
End
Attribute VB_Name = "AddJournalTransaction"
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

Dim AddTransactionRS As Recordset
Attribute AddTransactionRS.VB_VarHelpID = -1
Dim adoMasterRS As Recordset
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean


Private Sub cboHead_Validate(Cancel As Boolean)
If cboHead.ListIndex = -1 Then
    'MsgBox Me.ActiveControl
    MsgBox "First you should select a Head"
    'Cancel = True
End If

End Sub

Private Sub Form_Load()
    
  Set adoMasterRS = New Recordset
  adoMasterRS.Open "select actitle,acnumber FROM " & MasterTable & " order by actitle", _
    db, adOpenStatic, adLockReadOnly, adCmdText
  With adoMasterRS
    .MoveFirst
    Do While Not .EOF
    cboHead.AddItem !ACTITLE
    cboHead.ItemData(cboHead.NewIndex) = !ACNUMBER
    .MoveNext
    Loop
    .Close
  End With
    
    Set adoMasterRS = Nothing
    mskAcnDate.Text = Format(CurrentDate, "dd/mm/yyyy")
End Sub


Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub


Private Sub cmdCancel_Click()
    txtCash = Format(Format(CashInHand, "0.00"), "@@@@@@@@@@@@")
    SelectedRecord = Val(txtEntryNo)
    'ViewTransaction.ViewTransactionRS.Requery
    ViewJournalTransaction.RefreshTransaction
    ViewJournalTransaction.FillGridTrans
    Me.Hide
    ViewJournalTransaction.Show
End Sub


Private Sub ClearFields()
    txtEntryNo = ""
    'txtAcnDate = Format(CurrentDate, "dd/mm/yyyy")
    txtCash = Format(Format(CashInHand, "0.00"), "@@@@@@@@@@@@")
    mskAcnDate.Text = Format(CurrentDate, "dd/mm/yyyy")
    cboHead.ListIndex = -1
    txtDebit = "0.00"
    txtDebit = Format(Format(Val(txtDebit), "0.00"), "@@@@@@@@@@")
    txtCredit = "0.00"
    txtCredit = Format(Format(Val(txtCredit), "0.00"), "@@@@@@@@@@")
    txtParticular = ""
End Sub


Public Sub LoadRecord()
With AddTransactionRS
    !entry_id = Val(txtEntryNo)
    '!acn_date = Format(txtAcnDate, "dd/mm/yyyy")
    !acn_date = mskAcnDate.Text
    
    !ACNUMBER = cboHead.ItemData(cboHead.ListIndex)
    
    !debit = Val(txtDebit)
    !credit = Val(txtCredit)

    !particular = Trim(txtParticular)
    !cj = "J"
    CashInHand = CashInHand + Val(txtCredit) - Val(txtDebit)

End With
End Sub

Private Sub cmdSave_Click()
 ' On Error GoTo AddErr
'Dim strhead As String
  
  Set AddTransactionRS = New Recordset
    
  AddTransactionRS.Open "select MAX(entry_id)" _
    & " FROM " & TransactionTable, db, adOpenStatic, adLockReadOnly, adCmdText

If AddTransactionRS.BOF = False And AddTransactionRS.EOF = False Then
  txtEntryNo = AddTransactionRS.Fields(0).Value + 1
Else
  txtEntryNo = 1
End If
  AddTransactionRS.Close

  
If cboHead.ListIndex < 0 Then
    'IncompleteEntry = True
    MsgBox "Select a Head"
    cboHead.SetFocus
    Exit Sub
ElseIf (Val(txtDebit) = 0 And Val(txtCredit) = 0) Then
    MsgBox "Both Amounts should not be empty"
    txtDebit.SetFocus
    Exit Sub
ElseIf Len(Trim(txtParticular)) = 0 Then
    MsgBox "Narration should not be empty"
    txtParticular.SetFocus
    Exit Sub
End If


  Set AddTransactionRS = New Recordset
  AddTransactionRS.Open "select entry_id,acn_date,acnumber,particular,debit,credit,cj FROM " & TransactionTable & " where entry_id= -1", db, adOpenStatic, adLockOptimistic, adCmdText
  AddTransactionRS.AddNew
  LoadRecord
  AddTransactionRS.Update
  AddTransactionRS.Close
  'txtCash = Format(Format(CashInHand, "0.00"), "@@@@@@@@@@@@")
'    SelectedRecord = Val(txtEntryNo)
'    ViewTransaction.ViewTransactionRS.Requery
'    ViewTransaction.FillGridTrans
'
'      Me.Hide
'      ViewTransaction.Show
    
    ClearFields
    cboHead.SetFocus

  Exit Sub
'AddErr:
'  MsgBox Err.Description

End Sub



Private Sub Form_Activate()
'cmdSave.Enabled = False
'IncompleteEntry = False
  ClearFields
cboHead.SetFocus
End Sub



Private Sub txtCredit_GotFocus()
txtCredit = Format(Val(txtCredit), "0.00")
txtCredit.SelStart = 0
txtCredit.SelLength = Len(txtCredit)

End Sub

Private Sub txtCredit_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyBack) Then Exit Sub
    If InStr("1234567890.", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtCredit_LostFocus()
txtCredit = Format(Format(Val(txtCredit), "0.00"), "@@@@@@@@@@")
End Sub

Private Sub txtCredit_Validate(Cancel As Boolean)
If Val(txtCredit) > 0 Then
    If Val(txtDebit) > 0 Then
        txtDebit = "0.00"
        txtDebit = Format(Format(Val(txtDebit), "0.00"), "@@@@@@@@@@")
    End If
'Else
'    If Val(txtDebit) = 0 Then
'        MsgBox "Both Amounts should not be zero"
'    End If
End If



End Sub

Private Sub txtDebit_GotFocus()
    txtDebit = Format(Val(txtDebit), "0.00")
    txtDebit.SelStart = 0
    txtDebit.SelLength = Len(txtDebit)
End Sub

Private Sub txtDebit_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyBack) Then Exit Sub
    If InStr("1234567890.", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtDebit_LostFocus()
txtDebit = Format(Format(Val(txtDebit), "0.00"), "@@@@@@@@@@")
End Sub

Private Sub txtDebit_Validate(Cancel As Boolean)
If Val(txtDebit) > 0 Then
    If Val(txtCredit) > 0 Then
        txtCredit = "0.00"
        txtCredit = Format(Format(Val(txtCredit), "0.00"), "@@@@@@@@@@")
    End If
Else
    If Val(txtCredit) = 0 Then
        MsgBox "Both Amounts can not be empty"
    End If
End If



End Sub

Private Sub txtParticular_GotFocus()
    txtParticular.SelStart = 0
    txtParticular.SelLength = Len(Trim(txtParticular))
End Sub

Private Sub txtParticular_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtParticular_LostFocus()
    txtParticular.Text = Trim(txtParticular)
End Sub

Private Sub txtParticular_Validate(Cancel As Boolean)
    If Len(Trim(txtParticular)) = 0 Then
        MsgBox "Narration can not be empty"
        'Cancel = True
    End If
End Sub
'**************



'Private Sub CancelButton_Click()
'    mskAcnDate.Text = Format(CurrentDate, "dd/mm/yyyy")
'    Unload Me
'End Sub

Private Sub cmdIncrease_Click()
If mskAcnDate.Text < EndDate Then
mskAcnDate.Text = CurrentDate + 1
CurrentDate = mskAcnDate.Text
If ToDate < CurrentDate Then
    ToDate = CurrentDate
End If
ShowMainCaption
'frmMain.Caption = Trim(ViewCompanyRS("cotitle").Value) + " - (" + ViewCompanyRS("coyear").Value + ")                   " + Format(CurrentDate, "dd/mm/yyyy")
Else
MsgBox "Date out of Range"
End If
txtCash = Format(Format(CashInHand, "0.00"), "@@@@@@@@@@@@")
End Sub

Private Sub cmdDecrease_Click()
If mskAcnDate.Text > StartDate + 1 Then
mskAcnDate.Text = CurrentDate - 1
CurrentDate = mskAcnDate.Text
If ToDate < CurrentDate Then
    ToDate = CurrentDate
End If
ShowMainCaption
'frmMain.Caption = Trim(ViewCompanyRS("cotitle").Value) + " - (" + ViewCompanyRS("coyear").Value + ")                   " + Format(CurrentDate, "dd/mm/yyyy")
Else
MsgBox "Date out of Range"
End If
txtCash = Format(Format(CashInHand, "0.00"), "@@@@@@@@@@@@")
End Sub

Private Sub mskAcnDate_GotFocus()
    mskAcnDate.SelStart = 0
    mskAcnDate.SelLength = Len(mskAcnDate.Text)
End Sub

Private Sub mskAcnDate_Validate(Cancel As Boolean)
    If Not IsDate(mskAcnDate.Text) Then
        MsgBox "Enter a Date"
        Cancel = True
        mskAcnDate.SetFocus
    Else
        If mskAcnDate.Text >= StartDate And mskAcnDate.Text <= EndDate Then
            CurrentDate = mskAcnDate.Text
            If ToDate < CurrentDate Then
                ToDate = CurrentDate
            End If
            ShowMainCaption
        Else
        MsgBox "Date out of Range"
        Cancel = True
        End If
    End If
    txtCash = Format(Format(CashInHand, "0.00"), "@@@@@@@@@@@@")
End Sub


