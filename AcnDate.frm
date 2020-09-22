VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form AcnDate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select   Account  Date"
   ClientHeight    =   1695
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "-"
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
      Left            =   4680
      TabIndex        =   5
      Top             =   308
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
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
      Left            =   3895
      TabIndex        =   4
      Top             =   308
      Width           =   495
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   390
      Left            =   1640
      TabIndex        =   0
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
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2835
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1395
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Account Date: "
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   398
      Width           =   1080
   End
End
Attribute VB_Name = "AcnDate"
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
    MaskEdBox1.Text = Format(CurrentDate, "dd/mm/yyyy")
    Unload Me
End Sub

Private Sub Command1_Click()
If MaskEdBox1.Text < EndDate Then
MaskEdBox1.Text = CurrentDate + 1
CurrentDate = MaskEdBox1.Text
If ToDate < CurrentDate Then
    ToDate = CurrentDate
End If
'ShowMainCaption
'frmMain.Caption = Trim(ViewCompanyRS("cotitle").Value) + " - (" + ViewCompanyRS("coyear").Value + ")                   " + Format(CurrentDate, "dd/mm/yyyy")
Else
MsgBox "Date out of Range"
End If
End Sub

Private Sub Command2_Click()
If MaskEdBox1.Text > StartDate + 1 Then
MaskEdBox1.Text = CurrentDate - 1
CurrentDate = MaskEdBox1.Text
If ToDate < CurrentDate Then
    ToDate = CurrentDate
End If
'ShowMainCaption
'frmMain.Caption = Trim(ViewCompanyRS("cotitle").Value) + " - (" + ViewCompanyRS("coyear").Value + ")                   " + Format(CurrentDate, "dd/mm/yyyy")
Else
MsgBox "Date out of Range"
End If

End Sub

Private Sub Form_Activate()
    'OKButton.SetFocus
    MaskEdBox1.SetFocus
End Sub

Private Sub Form_Load()
    MaskEdBox1.Text = Format(CurrentDate, "dd/mm/yyyy")
    
End Sub

Private Sub MaskEdBox1_GotFocus()
MaskEdBox1.SelStart = 0
MaskEdBox1.SelLength = Len(MaskEdBox1.Text)
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
    If Not IsDate(MaskEdBox1.Text) Then
        MsgBox "Enter a Date"
        Cancel = True
        MaskEdBox1.SetFocus
    End If
    
End Sub

Private Sub OKButton_Click()
    If IsDate(MaskEdBox1.Text) And MaskEdBox1.Text >= StartDate And MaskEdBox1.Text <= EndDate Then
        CurrentDate = MaskEdBox1.Text
        If ToDate < CurrentDate Then
            ToDate = CurrentDate
        End If
    Else
        MsgBox "Enter a Valid Date"
        MaskEdBox1.SetFocus
        Exit Sub
    End If
    CalculateCashInHand
 '   ShowMainCaption
    'frmMain.Caption = Trim(ViewCompanyRS("cotitle").Value) + " - (" + ViewCompanyRS("coyear").Value + ")                   " + Format(CurrentDate, "dd/mm/yyyy")
    Unload Me
End Sub
