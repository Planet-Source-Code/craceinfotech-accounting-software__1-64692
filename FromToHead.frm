VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FromToHead 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Head and Period"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   300
      Left            =   2010
      TabIndex        =   3
      Top             =   2700
      Width           =   1095
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   300
      Left            =   3465
      TabIndex        =   4
      Top             =   2700
      Width           =   1095
   End
   Begin VB.Frame FrameSelect 
      Height          =   2310
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   6090
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
         Left            =   1260
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1560
         Width           =   4575
      End
      Begin MSMask.MaskEdBox mskFromDt 
         Height          =   390
         Left            =   1260
         TabIndex        =   0
         Top             =   360
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
      Begin MSMask.MaskEdBox mskToDt 
         Height          =   390
         Left            =   1260
         TabIndex        =   1
         Top             =   960
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "From Date : "
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   465
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "To Date   :   "
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   1065
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Head       :"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   1665
         Width           =   750
      End
   End
End
Attribute VB_Name = "FromToHead"
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

Dim LedgerStart As Boolean

Private Sub CancelButton_Click()
    mskFromDt.Text = Format(FromDate, "dd/mm/yyyy")
    mskToDt.Text = Format(ToDate, "dd/mm/yyyy")
    ContinueProcess = False
    Unload Me
End Sub

Private Sub Form_Activate()
    If LedgerStart Then
        mskFromDt.SetFocus
    Else
        cboHead.SetFocus
    End If
    
End Sub

Private Sub Form_Load()
  LedgerStart = True
  Set ViewMasterRS = New Recordset
  ViewMasterRS.Open "select acnumber,actitle" _
    & " FROM " & MasterTable & " order by actitle", db, adOpenStatic, adLockReadOnly, adCmdText
    
    'ViewMasterRS.Open "select acnumber,actitle,balancetyp,balance," _
    '& " FROM " &  MASTERTABLE  &  " order by actitle", db, adOpenStatic, adLockReadOnly, adCmdText
    
    'cboload = False
  With ViewMasterRS
If .EOF = False And .BOF = False Then
    .MoveFirst
    Do While Not .EOF
        With cboHead
            .AddItem ViewMasterRS("actitle").Value
            .ItemData(.NewIndex) = ViewMasterRS("acnumber").Value
        End With
        .MoveNext
    Loop
    .MoveFirst
    'cboload = True
End If
    .Close
End With
    cboHead.ListIndex = 0
    SelectedRecord = cboHead.ItemData(cboHead.ListIndex)
    Set ViewMasterRS = Nothing
    
    mskFromDt.Text = Format(FromDate, "dd/mm/yyyy")
    mskToDt.Text = Format(ToDate, "dd/mm/yyyy")
End Sub

Private Sub mskFromDt_GotFocus()
    mskFromDt.SelStart = 0
    mskFromDt.SelLength = Len(mskFromDt.Text)
End Sub


Private Sub mskFromDt_Validate(Cancel As Boolean)
    If Not IsDate(mskFromDt.Text) Then
        MsgBox "Enter a Date"
        Cancel = True
    End If

End Sub

Private Sub mskToDt_GotFocus()
    mskToDt.SelStart = 0
    mskToDt.SelLength = Len(mskToDt.Text)
End Sub

Private Sub mskToDt_Validate(Cancel As Boolean)
    If Not IsDate(mskToDt.Text) Then
        MsgBox "Enter a Date"
        Cancel = True
    End If
End Sub

Private Sub OKButton_Click()
    If Not IsDate(mskFromDt.Text) Then
        MsgBox "Enter a valid From date"
        Exit Sub
    ElseIf Not IsDate(mskToDt.Text) Then
        MsgBox "Enter a valid To date"
        Exit Sub
    End If
    
    If mskFromDt.Text >= StartDate And mskFromDt.Text <= EndDate And mskFromDt.Text <= SystemDate Then
        FromDate = mskFromDt.Text
    Else
        MsgBox "Enter a Valid From Date"
        mskFromDt.SetFocus
        Exit Sub
    End If
    
    If mskToDt.Text >= StartDate And mskToDt.Text <= EndDate And mskToDt.Text <= SystemDate And mskToDt >= FromDate Then
        ToDate = mskToDt.Text
    Else
        MsgBox "Enter a Valid To Date"
        mskToDt.SetFocus
        Exit Sub
    End If
    If cboHead.ListIndex <> -1 Then
        SelectedRecord = cboHead.ItemData(cboHead.ListIndex)
        SelectedHead = Trim(cboHead.Text)
    Else
        MsgBox "Select a Head"
        cboHead.SetFocus
        Exit Sub
    End If
    
    ContinueProcess = True
    Me.Hide
    
    Ledger.Show 1
    If ContinueProcess Then
        LedgerStart = False
        Me.Show 1
    End If

End Sub

