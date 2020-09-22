VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form AsOnTrial 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Trial Balance"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   300
      Left            =   870
      TabIndex        =   0
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   300
      Left            =   2295
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Frame FrameSelect 
      Height          =   1125
      Left            =   300
      TabIndex        =   2
      Top             =   180
      Width           =   3570
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
Attribute VB_Name = "AsOnTrial"
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
    Else
        MsgBox "Enter a Valid To Date"
        mskToDt.SetFocus
        Exit Sub
    End If
    
    ContinueProcess = True
    Me.Hide
    
    Trial.Show 1
    If ContinueProcess Then
        Me.Show 1
    End If

End Sub

