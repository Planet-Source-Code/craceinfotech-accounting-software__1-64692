VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form AsOnTrading 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Trading Account"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   300
      Left            =   870
      TabIndex        =   0
      Top             =   2460
      Width           =   1095
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   300
      Left            =   2295
      TabIndex        =   1
      Top             =   2460
      Width           =   1095
   End
   Begin VB.Frame FrameSelect 
      Height          =   1950
      Left            =   300
      TabIndex        =   2
      Top             =   180
      Width           =   3570
      Begin VB.TextBox txtSTOCK 
         DataField       =   "balance"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
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
         Left            =   1440
         MaxLength       =   11
         TabIndex        =   5
         Top             =   1140
         Width           =   1935
      End
      Begin MSMask.MaskEdBox mskToDt 
         Height          =   390
         Left            =   1440
         TabIndex        =   3
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Closing Stock  :  "
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   1245
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "As On   :   "
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   525
         Width           =   750
      End
   End
End
Attribute VB_Name = "AsOnTrading"
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

Private Sub CancelButton_Click()
    mskToDt.Text = Format(ToDate, "dd/mm/yyyy")
    ContinueProcess = False
    Unload Me
End Sub

Private Sub Form_Activate()
       txtSTOCK = Format(StockInHand, "0.00")
       txtSTOCK = Format(Format(Val(txtSTOCK), "0.00"), "@@@@@@@@@@@")
       mskToDt.SetFocus
End Sub

Private Sub Form_Load()
    mskToDt.Text = Format(ToDate, "dd/mm/yyyy")
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
    If Not IsDate(mskToDt.Text) Then
        MsgBox "Enter a valid To date"
        Exit Sub
    End If
        
    If mskToDt.Text >= StartDate And mskToDt.Text <= EndDate And mskToDt.Text <= SystemDate Then
        ToDate = mskToDt.Text
        FromDate = StartDate + 1
    Else
        MsgBox "Enter a Valid To Date"
        mskToDt.SetFocus
        Exit Sub
    End If
    StockInHand = Format(Val(txtSTOCK), "0.00")
    ContinueProcess = True
    Me.Hide
    
    Trading.Show 1
    If ContinueProcess Then
        Me.Show 1
    End If

End Sub

Private Sub txtstock_GotFocus()
txtSTOCK = Format(Val(txtSTOCK), "0.00")
txtSTOCK.SelStart = 0
txtSTOCK.SelLength = Len(txtSTOCK)
End Sub

Private Sub txtstock_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyBack) Then Exit Sub
    If InStr("1234567890.", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtstock_LostFocus()
txtSTOCK = Format(Format(Val(txtSTOCK), "0.00"), "@@@@@@@@@@@")
End Sub

Private Sub txtstock_Validate(Cancel As Boolean)

If IsNumeric(txtSTOCK) = False Then
    MsgBox "Enter Zero or a valid Positive number"
    Cancel = True
End If

End Sub

