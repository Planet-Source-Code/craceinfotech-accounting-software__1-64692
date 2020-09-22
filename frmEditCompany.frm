VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form EditCompany 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Company"
   ClientHeight    =   3435
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   11130
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   11130
      TabIndex        =   8
      Top             =   2835
      Width           =   11130
      Begin VB.CommandButton cmdSave 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   300
         Left            =   4448
         TabIndex        =   5
         ToolTipText     =   "To Select a Company or Close this form"
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         CausesValidation=   0   'False
         Height          =   300
         Left            =   5588
         TabIndex        =   6
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   11130
      TabIndex        =   7
      Top             =   3135
      Width           =   11130
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   180
      TabIndex        =   9
      Top             =   180
      Width           =   10575
      Begin MSMask.MaskEdBox mskStart 
         Height          =   390
         Left            =   1500
         TabIndex        =   3
         Top             =   1740
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   688
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
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
      Begin VB.TextBox txtCompany 
         DataField       =   "balance"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1500
         MaxLength       =   60
         TabIndex        =   1
         Top             =   382
         Width           =   8595
      End
      Begin VB.TextBox txtCode 
         DataField       =   "acnumber"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   1057
         Width           =   1755
      End
      Begin MSMask.MaskEdBox mskEnd 
         Height          =   390
         Left            =   8340
         TabIndex        =   4
         Top             =   1740
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   688
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
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
      Begin MSMask.MaskEdBox mskYear 
         Height          =   390
         Left            =   8340
         TabIndex        =   2
         Top             =   1057
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   688
         _Version        =   393216
         MaxLength       =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "####-####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   14
         Top             =   1155
         Width           =   375
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Start Date"
         Height          =   195
         Index           =   4
         Left            =   300
         TabIndex        =   13
         Top             =   1838
         Width           =   720
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "End Date"
         Height          =   195
         Index           =   5
         Left            =   7320
         TabIndex        =   12
         Top             =   1838
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Company"
         Height          =   195
         Left            =   300
         TabIndex        =   11
         Top             =   480
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Year"
         Height          =   195
         Left            =   7320
         TabIndex        =   10
         Top             =   1155
         Width           =   330
      End
   End
End
Attribute VB_Name = "EditCompany"
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

Dim adoCompanyRS As Recordset
Dim adoSecondaryRS As Recordset
Dim IncompleteEntry As Boolean
Dim oldtitle  As String
Dim oldyear As String


Private Sub cmdSave_Click()
 ' On Error GoTo AddErr
Dim strCompany As String
Dim stryear As String
  Set adoSecondaryRS = New Recordset
  
If txtCompany <> oldtitle Or mskYear.Text <> oldyear Then

    strCompany = Replace(txtCompany, "'", "''")
    stryear = mskYear.Text
  
    adoSecondaryRS.Open "select count(*)" _
     & " from company where cotitle= '" & strCompany & "' and coyear= '" & stryear & "'", db, adOpenStatic, adLockReadOnly, adCmdText

    If adoSecondaryRS.Fields(0).Value > 0 Then
        MsgBox "Company already exists"
        txtCompany.SetFocus
        Exit Sub
    End If

    adoSecondaryRS.Close
  
End If

  
If Len(txtCompany) = 0 Then
    MsgBox "Name should not be empty"
    txtCompany.SetFocus
    Exit Sub
ElseIf Len(mskYear.ClipText) = 0 Then
    'IncompleteEntry = True
    MsgBox "Year should  not be empty"
    mskYear.SetFocus
    Exit Sub
ElseIf Len(mskStart.ClipText) = 0 Then
    'IncompleteEntry = True
    MsgBox "Start Date should  not be empty"
    mskStart.SetFocus
    Exit Sub
ElseIf Len(mskEnd.ClipText) = 0 Then
    'IncompleteEntry = True
    MsgBox "End Date should  not be empty"
    mskEnd.SetFocus
    Exit Sub
End If


  
    

  
    LoadRecord
    adoCompanyRS.Update
    ViewCompanyRS.Requery
    
    If txtCompany <> oldtitle Or mskYear.Text <> oldyear Then
          With ViewCompanyRS
            If .EOF = False And .BOF = False Then
            .MoveFirst
            Do While Not .EOF
                With ViewCompany.cboHead
                    .AddItem Trim(ViewCompanyRS("cotitle").Value) + " - (" + ViewCompanyRS("coyear").Value + ")"
                    .ItemData(.NewIndex) = ViewCompanyRS("conumber").Value
                End With
                .MoveNext
            Loop
            .MoveFirst
            'cboload = True
            End If
        End With
    End If

    ViewCompany.cboHead.Text = Trim(adoCompanyRS!cotitle) + " - (" + adoCompanyRS!COYear + ")"
    adoCompanyRS.Close

    Me.Hide
    ViewCompany.Show
  Exit Sub
'AddErr:
'  MsgBox Err.Description

End Sub

Private Sub Form_Activate()

  Set adoCompanyRS = New Recordset

  adoCompanyRS.Open "select conumber,cotitle,coyear,costart,coend" _
    & " from  company where conumber= " & SelectedCompany, db, adOpenStatic, adLockOptimistic, adCmdText
  
  'ClearFields
  GetFields
  IncompleteEntry = False
  txtCompany.SetFocus

End Sub


Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
    ViewCompany.Show
End Sub

Private Sub ClearFields()
    txtCode = ""
    txtCompany = ""
    mskYear = "2001-2002"
    mskStart = Format(#4/1/2001#, "mm/dd/yyyy")
    mskEnd = Format(#3/31/2002#, "mm/dd/yyyy")
    End Sub

Public Sub LoadRecord()
With adoCompanyRS
    !CONUMBER = Val(txtCode)
    !cotitle = Trim(txtCompany)
    !COYear = mskYear.Text
    !costart = Format(mskStart, "dd/mm/yyyy")
    !COend = Format(mskEnd, "dd/mm/yyyy")
End With
End Sub


Private Sub mskEnd_GotFocus()
    mskEnd.SelStart = 0
    mskEnd.SelLength = Len(mskEnd)
End Sub

Private Sub mskEnd_Validate(Cancel As Boolean)
If Not IsDate(mskEnd.Text) Then
    MsgBox "Enter a valid date"
    Cancel = True
    Exit Sub
End If
    
If Right(mskEnd.Text, 4) <> Right(mskYear.Text, 4) Then
    MsgBox "Enter a valid date"
    Cancel = True
    Exit Sub
End If

End Sub

Private Sub mskStart_GotFocus()
    mskStart.SelStart = 0
    mskStart.SelLength = Len(mskStart)
End Sub

Private Sub mskStart_Validate(Cancel As Boolean)

If Not IsDate(mskStart.Text) Then
    MsgBox "Enter a valid date"
    Cancel = True
    Exit Sub
End If

If Right(mskStart.Text, 4) <> Left(mskYear.Text, 4) Then
    MsgBox "Enter a valid date"
    Cancel = True
    Exit Sub
End If

End Sub

Private Sub mskYear_GotFocus()
    mskYear.SelStart = 0
    mskYear.SelLength = Len(mskYear)
End Sub

Private Sub mskYear_Validate(Cancel As Boolean)
If Len(mskYear.ClipText) <> 8 Then
    MsgBox "Year should be entered like 2001-2002"
    Cancel = True
    Exit Sub
End If

If Not (Val(Mid(mskYear.Text, 1, 2)) >= 19 And Val(Mid(mskYear.Text, 1, 2)) <= 20) Then
    MsgBox "Year should start with 19 or 20"
    Cancel = True
    Exit Sub
End If

If Val(Mid(mskYear.Text, 6, 4)) <> Val(Mid(mskYear.Text, 1, 4)) + 1 Then
    MsgBox "Ending Year should be one year more than Starting Year"
    Cancel = True
    Exit Sub
End If
'mskStart = "01/04/" + Mid(mskYear.Text, 1, 4)
'mskEnd = "31/03/" + Mid(mskYear.Text, 6, 4)

End Sub

Private Sub txtcompany_GotFocus()
    txtCompany.SelStart = 0
    txtCompany.SelLength = Len(txtCompany)
End Sub

Private Sub txtCompany_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtCompany_Validate(Cancel As Boolean)
    If Len(Trim(txtCompany)) = 0 Then
        MsgBox "Name can not be empty"
        Cancel = True
    End If
End Sub
Private Sub GetFields()

    txtCompany = Trim(adoCompanyRS!cotitle)
    oldtitle = Trim(adoCompanyRS!cotitle)
    txtCode = Format(adoCompanyRS("conumber").Value, "@@@@")
    mskYear.Text = adoCompanyRS!COYear
    oldyear = adoCompanyRS!COYear
    mskStart.Text = Format(adoCompanyRS!costart, "dd/mm/yyyy")
    mskEnd.Text = Format(adoCompanyRS!COend, "dd/mm/yyyy")
End Sub


