VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FromTo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select the Period"
   ClientHeight    =   2220
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   300
      Left            =   2362
      TabIndex        =   1
      Top             =   1740
      Width           =   1095
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   300
      Left            =   907
      TabIndex        =   0
      Top             =   1740
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   195
      TabIndex        =   2
      Top             =   0
      Width           =   3975
      Begin MSMask.MaskEdBox mskFromDt 
         Height          =   390
         Left            =   1620
         TabIndex        =   5
         Top             =   300
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
         Left            =   1605
         TabIndex        =   6
         Top             =   900
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
         Caption         =   "To Date: "
         Height          =   195
         Left            =   465
         TabIndex        =   4
         Top             =   1005
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "From Date: "
         Height          =   195
         Left            =   450
         TabIndex        =   3
         Top             =   405
         Width           =   825
      End
   End
End
Attribute VB_Name = "FromTo"
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
    mskFromDt.Text = Format(FromDate, "dd/mm/yyyy")
    mskToDt.Text = Format(ToDate, "dd/mm/yyyy")
    ContinueProcess = False
    Unload Me
End Sub

Private Sub Form_Activate()
    mskFromDt.SetFocus
End Sub

Private Sub Form_Load()
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
    
    If (mskFromDt.Text >= StartDate) And (mskFromDt.Text <= EndDate) And (mskFromDt.Text <= SystemDate) Then
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
    
    ContinueProcess = True
    Unload Me
End Sub
