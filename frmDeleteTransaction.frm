VERSION 5.00
Begin VB.Form DeleteTransaction 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete Transaction"
   ClientHeight    =   4800
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   8985
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      ForeColor       =   &H80000008&
      Height          =   1875
      Left            =   60
      TabIndex        =   14
      Top             =   720
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
         Locked          =   -1  'True
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         TabStop         =   0   'False
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
         Left            =   7012
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
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
         Left            =   5264
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "1234567890"
         Top             =   840
         Width           =   1635
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "CREDIT"
         Height          =   195
         Index           =   3
         Left            =   7529
         TabIndex        =   17
         Top             =   240
         Width           =   600
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "DEBIT"
         Height          =   195
         Index           =   6
         Left            =   5841
         TabIndex        =   16
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "PARTICULARS"
         Height          =   195
         Index           =   7
         Left            =   2085
         TabIndex        =   15
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
      Left            =   6720
      TabIndex        =   0
      Top             =   3780
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   6720
      TabIndex        =   1
      Top             =   3120
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
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "frmDeleteTransaction.frx":0000
      Top             =   2880
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
      Left            =   4320
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "12345678"
      Top             =   180
      Width           =   1395
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
      TabIndex        =   12
      Top             =   4410
      Width           =   8985
   End
   Begin VB.TextBox txtAcnDate 
      CausesValidation=   0   'False
      DataField       =   "acn_date"
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
      Left            =   7080
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   180
      Width           =   1635
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "No  "
      Height          =   195
      Index           =   0
      Left            =   3720
      TabIndex        =   13
      Top             =   285
      Width           =   300
   End
   Begin VB.Label lblLabels 
      Caption         =   "CREDIT"
      Height          =   255
      Index           =   5
      Left            =   6345
      TabIndex        =   11
      Top             =   1260
      Width           =   1635
   End
   Begin VB.Label lblLabels 
      Caption         =   "DEBIT"
      Height          =   255
      Index           =   4
      Left            =   4305
      TabIndex        =   10
      Top             =   1260
      Width           =   1635
   End
   Begin VB.Label lblLabels 
      Caption         =   "PARTICULARS"
      Height          =   255
      Index           =   2
      Left            =   525
      TabIndex        =   9
      Top             =   1200
      Width           =   3375
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "DATE"
      Height          =   195
      Index           =   1
      Left            =   6120
      TabIndex        =   8
      Top             =   285
      Width           =   435
   End
End
Attribute VB_Name = "DeleteTransaction"
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

Dim DeleteTransactionRS As Recordset
Attribute DeleteTransactionRS.VB_VarHelpID = -1
Dim adoMasterRS As Recordset
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean

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
End Sub


Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub


Private Sub cmdCancel_Click()
    Me.Hide
    ViewCashTransaction.Show
End Sub


Private Sub ClearFields()
    txtEntryNo = ""
    txtAcnDate = Format(CurrentDate, "dd/mm/yyyy")
    cboHead.ListIndex = -1
    txtDebit = "0.00"
    txtDebit = Format(Format(Val(txtDebit), "0.00"), "@@@@@@@@@@")
    txtCredit = "0.00"
    txtCredit = Format(Format(Val(txtCredit), "0.00"), "@@@@@@@@@@")
    txtParticular = ""
End Sub

Private Sub cmdSave_Click()
 ' On Error GoTo AddErr

  
  DeleteTransactionRS.Delete
  DeleteTransactionRS.Close
  SelectedRecord = SelectedRecord + 1
    
  ViewCashTransaction.ViewTransactionRS.Requery
  ViewCashTransaction.FillGridTrans
  
    Me.Hide
    ViewCashTransaction.Show
  Exit Sub
'AddErr:
'  MsgBox Err.Description

End Sub

Private Sub Form_Activate()
  Set DeleteTransactionRS = New Recordset
  DeleteTransactionRS.Open "select entry_id,acn_date,acnumber,particular,debit,credit FROM " & TransactionTable & " where entry_id= " & SelectedRecord, db, adOpenStatic, adLockOptimistic, adCmdText

  ClearFields
  GetFields

End Sub

Private Sub GetFields()
    txtEntryNo = Format(DeleteTransactionRS("entry_id").Value, "@@@@@@@@")
    txtAcnDate = Format(DeleteTransactionRS("acn_date").Value, "dd/mm/yyyy")
    
    Dim i As Integer
    For i = 0 To cboHead.ListCount - 1
        If cboHead.ItemData(i) = DeleteTransactionRS!ACNUMBER Then Exit For
    Next i
    cboHead.ListIndex = i
    
    txtDebit = Format(Format(DeleteTransactionRS("debit").Value, "0.00"), "@@@@@@@@@@")
    txtCredit = Format(Format(DeleteTransactionRS("credit").Value, "0.00"), "@@@@@@@@@@")
    txtParticular = Trim(DeleteTransactionRS("particular").Value)

End Sub




