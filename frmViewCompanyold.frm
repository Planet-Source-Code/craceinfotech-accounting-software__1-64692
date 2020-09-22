VERSION 5.00
Begin VB.Form ViewCompanyold 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Company"
   ClientHeight    =   4695
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   11130
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
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
      TabIndex        =   10
      Top             =   4095
      Width           =   11130
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Select"
         Default         =   -1  'True
         Height          =   300
         Left            =   6195
         TabIndex        =   1
         ToolTipText     =   "To Select a Company or Close this form"
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Quit"
         Height          =   300
         Left            =   7335
         TabIndex        =   5
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Remove"
         Height          =   300
         Left            =   5016
         TabIndex        =   4
         ToolTipText     =   "Delete Company"
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Change"
         Height          =   300
         Left            =   3858
         TabIndex        =   3
         ToolTipText     =   "Change Company's Name, Year etc."
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&New"
         Height          =   300
         Left            =   2700
         TabIndex        =   2
         ToolTipText     =   "Create a New Company"
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
      TabIndex        =   8
      Top             =   4395
      Width           =   11130
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   9
         Top             =   0
         Width           =   3360
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   278
      TabIndex        =   11
      Top             =   480
      Width           =   10575
      Begin VB.TextBox txtEnd 
         DataField       =   "balance"
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
         Left            =   8340
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1740
         Width           =   1755
      End
      Begin VB.TextBox txtStart 
         DataField       =   "balance"
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
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1740
         Width           =   1755
      End
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
         Left            =   1500
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   8595
      End
      Begin VB.TextBox txtCode 
         DataField       =   "acnumber"
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
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1755
      End
      Begin VB.TextBox txtYear 
         DataField       =   "balance"
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
         Left            =   8340
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1755
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   16
         Top             =   1155
         Width           =   375
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Start Date"
         Height          =   195
         Index           =   4
         Left            =   300
         TabIndex        =   15
         Top             =   1838
         Width           =   720
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "End Date"
         Height          =   195
         Index           =   5
         Left            =   7320
         TabIndex        =   14
         Top             =   1838
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Company"
         Height          =   195
         Left            =   300
         TabIndex        =   13
         Top             =   480
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Year"
         Height          =   195
         Left            =   7320
         TabIndex        =   12
         Top             =   1178
         Width           =   330
      End
   End
   Begin VB.Line Line3 
      X1              =   23
      X2              =   11093
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line2 
      X1              =   4620
      X2              =   5820
      Y1              =   2100
      Y2              =   2580
   End
   Begin VB.Line Line1 
      X1              =   4620
      X2              =   5820
      Y1              =   2100
      Y2              =   2580
   End
End
Attribute VB_Name = "ViewCompanyold"
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

Dim oldcompany As Integer
Dim mvBookMark As Variant
Dim cboload As Boolean

Private Sub cboHead_Click()
lblStatus.Caption = cboHead.Text

With ViewCompanyRS
    .MoveFirst
    .Find "conumber= " & cboHead.ItemData(cboHead.ListIndex)
    If .EOF Then
        MsgBox "Unable to find record"
        Exit Sub
    End If
    
    GetFields
    oldcompany = SelectedCompany
    SelectedCompany = cboHead.ItemData(cboHead.ListIndex)
End With
End Sub

Private Sub cmdSelect_Click()
 
  cmdClose.Enabled = False
  cmdAdd.Enabled = False
  cmdEdit.Enabled = False
  cmdDelete.Enabled = False
  
  CompanySelected = True

CompanyName = Trim(ViewCompanyRS!cotitle)
FinancialYear = Trim(ViewCompanyRS!COYear)
'frmMain.Caption = Trim(ViewCompanyRS("cotitle").Value) + " - (" + ViewCompanyRS("coyear").Value + ")    " + Format(CurrentDate, "dd/mm/yyyy")
StartDate = Format(ViewCompanyRS!costart, "dd/mm/yyyy")
EndDate = Format(ViewCompanyRS!COend, "dd/mm/yyyy")
FromDate = StartDate
CurrentDate = Format(Now(), "dd/mm/yyyy")
If CurrentDate > EndDate Then CurrentDate = EndDate
ToDate = CurrentDate
MasterTable = "MASTER" + Right(Format(Format(SelectedCompany, "000"), "@@@"), 2)
TransactionTable = "ENTRY" + Right(Format(Format(SelectedCompany, "000"), "@@@"), 2)
StockInHand = Format(ViewCompanyRS!stock, "0000000000.00")
CalculateToDate
CalculateCashInHand
  
  Me.Hide
  'fmMain.Caption = Trim(ViewCompanyRS("cotitle").Value) + " - (" + ViewCompanyRS("coyear").Value + ")                   " + Format(CurrentDate, "dd/mm/yyyy") + " Cash In Hand Rs. " & Format(CashInHand, "0.00")
  ShowMainCaption
  frmMain.Show
End Sub

Private Sub Form_Activate()
If cboHead.ListCount = 0 Then
    ClearFields
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    If MsgBox("No Master records." & vbCrLf & "Do you want to add now?", vbYesNo + vbDefaultButton1 + vbInformation, "Add Master") = vbYes Then
        cmdAdd.Value = True
    End If
Else
    If Not CompanySelected Then
        cmdEdit.Enabled = True
        cmdDelete.Enabled = True
    End If
    ClearFields
    Dim i  As Integer
    For i = 0 To cboHead.ListCount
        If cboHead.ItemData(i) = SelectedCompany Then
            Exit For
        End If
    Next
    GetFields
    cboHead.SetFocus
End If

End Sub

Private Sub Form_Load()
  
  Set ViewCompanyRS = New Recordset
  ViewCompanyRS.Open "select conumber,cotitle,coyear,costart,coend,stock" _
    & " from  company order by cotitle, coyear", db, adOpenStatic, adLockReadOnly, adCmdText
    cboload = False
  With ViewCompanyRS
If .EOF = False And .BOF = False Then
    .MoveFirst
    Do While Not .EOF
        With cboHead
            .AddItem Trim(ViewCompanyRS("cotitle").Value) + " - (" + ViewCompanyRS("coyear").Value + ")"
            .ItemData(.NewIndex) = ViewCompanyRS("conumber").Value
        End With
        .MoveNext
    Loop
    .MoveFirst
    cboload = True
End If
End With
  ClearFields
    With ViewCompanyRS
    If .EOF = False And .BOF = False Then
        ViewCompanyRS.MoveFirst
        cboHead.ListIndex = 0
    End If
    End With

End Sub

Private Sub Form_Resize()
  On Error Resume Next
  lblStatus.Width = Me.ScaleWidth
  lblStatus.Left = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub cmdAdd_Click()
'    Me.Hide
 '   AddCompany.Show 1
 MsgBox "NO PROVISION FOR ADDING A COMPANY"
End Sub

Private Sub cmdEdit_Click()
  'On Error GoTo EditErr


With ViewCompanyRS
    mvBookMark = .Bookmark
    SelectedCompany = !CONUMBER
End With

    Me.Hide
    EditCompany.Show 1
  
End Sub
Private Sub cmdClose_Click()
  Unload Me
  End
End Sub

Private Sub ClearFields()
        txtCode = ""
        txtYear = ""
        txtStart = ""
        txtEnd = ""
End Sub

Private Sub GetFields()

    txtCode = Format(ViewCompanyRS("conumber").Value, "@@@@@@@@@@")
    txtYear = ViewCompanyRS!COYear
    txtStart = Format(ViewCompanyRS!costart, "dd/mm/yyyy")
    txtEnd = Format(ViewCompanyRS!COend, "dd/mm/yyyy")

End Sub

'Public Sub CalculateToDate()
'    Dim Temprs As Recordset
'    Set Temprs = New Recordset
'
'    Temprs.Open "select MAX(ACN_DATE)" _
'    & " from  " & TransactionTable, db, adOpenStatic, adLockReadOnly, adCmdText
'    If Temprs.RecordCount > 0 Then
'        If IsDate(Temprs.Fields(0).Value) Then
'            CurrentDate = Format(Temprs.Fields(0).Value, "dd/mm/yyyy")
'        ToDate = CurrentDate
'        End If
'    End If
'    Temprs.Close
'End Sub


