VERSION 5.00
Begin VB.Form ViewCompany 
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
      Begin VB.CommandButton cmdRestore 
         Caption         =   "&Restore"
         Height          =   300
         Left            =   7354
         TabIndex        =   20
         ToolTipText     =   "To Restore  a Company's Data from backup"
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdBackup 
         Caption         =   "&Back Up"
         Height          =   300
         Left            =   6186
         TabIndex        =   19
         ToolTipText     =   "To Back up  a Company's Data"
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Select"
         Height          =   300
         Left            =   5018
         TabIndex        =   1
         ToolTipText     =   "To Select a Company or Close this form"
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Quit"
         Height          =   300
         Left            =   8522
         TabIndex        =   5
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Remove"
         Height          =   300
         Left            =   3850
         TabIndex        =   4
         ToolTipText     =   "Delete Company"
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Change"
         Height          =   300
         Left            =   2682
         TabIndex        =   3
         ToolTipText     =   "Change Company's Name, Year etc."
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&New"
         Default         =   -1  'True
         Height          =   300
         Left            =   1514
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
Attribute VB_Name = "ViewCompany"
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

Private Sub cmdBackup_Click()
    Dim retval As Long
    Dim fLen As Integer, filepath As String
    CompanyName = Trim(ViewCompanyRS!cotitle)
    CompanyYear = ViewCompanyRS("coyear").Value
    
    MasterTable = "C:\VBPROG\VBA\DAT\MASTER" + Right(Format(Format(SelectedCompany, "000"), "@@@"), 2) + ".DBF"
    TransactionTable = "C:\VBPROG\VBA\DAT\ENTRY" + Right(Format(Format(SelectedCompany, "000"), "@@@"), 2) + ".DBF"
    ZipFile = "BACKUP" + Right(Format(Format(SelectedCompany, "000"), "@@@"), 2) + ".ZIP"

    Open "c:\vbprog\vba\dat\filelist.txt" For Output As #1
    
    Print #1, MasterTable
    Print #1, TransactionTable
    'Print #1, Chr(12)
    Close #1
     
    retval = ExecCmd("C:\Program Files\WinZip\WINZIP32.EXE -a a:\" & ZipFile & "  @c:\vbprog\vba\dat\filelist.txt")
    'retval = ExecCmd("C:\Program Files\WinZip\WINZIP32.EXE -a d:\natarajan\test\backups\" & ZipFile & "  @c:\vbprog\vba\dat\filelist.txt")
    'retval = ExecCmd("C:\Program Files\WinZip\WINZIP32.EXE -e d:\natarajan\test\backups\backup.zip d:\natarajan\test\backups\test")
    
    'filepath = "d:\natarajan\test\backups\" & ZipFile
    filepath = "a:\" & ZipFile
    On Error Resume Next
    fLen = Len(Dir$(filepath))
    If Err Or fLen = 0 Then
        MsgBox "Backup Failure"
        Exit Sub
    Else
        MsgBox "Backup Over"
    End If
    
    cboHead.SetFocus
End Sub

Private Sub cmdDelete_Click()
On Error GoTo DeleteErr
    Dim intIndex As Integer
    If cboHead.ListCount < 2 Then
        MsgBox "Single Company Exists." & vbCrLf & "Can not be Deleted"
        Exit Sub
    End If
    If MsgBox("Delete this Company?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
        Dim Temprs As Recordset
        Set Temprs = New Recordset
        Temprs.Open "select conumber,covalid from company a where a.conumber=" & ViewCompanyRS!CONUMBER, db, adOpenStatic, adLockOptimistic, adCmdText
        Temprs("covalid").Value = "N"
        Temprs.Update
        Temprs.Close
        Set Temprs = Nothing
        ViewCompanyRS.Requery
        intIndex = cboHead.ListIndex
        cboHead.RemoveItem intIndex
        If intIndex = cboHead.ListCount Then
            cboHead.ListIndex = intIndex - 1
        Else
            cboHead.ListIndex = intIndex
        End If
    End If
    Exit Sub
DeleteErr:
    MsgBox "Some Errors occured. Please contact Mr.R.Natarajan"
 End Sub


Private Sub cmdRestore_Click()
    Me.Hide
    RestoreCompany.Show 1
    Unload Me
    End
End Sub
    

Private Sub cmdSelect_Click()
 
  cmdClose.Enabled = False
  cmdAdd.Enabled = False
  cmdEdit.Enabled = False
  cmdDelete.Enabled = False
  cmdBackup.Enabled = False
  cmdRestore.Enabled = False
  
  CompanySelected = True

CompanyName = Trim(ViewCompanyRS!cotitle)
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
  ViewCompanyRS.Open "select conumber,cotitle,coyear,costart,coend,covalid,STOCK" _
    & " from  company where covalid='Y' order by cotitle, coyear", db, adOpenStatic, adLockReadOnly, adCmdText
    cboload = False
  With ViewCompanyRS
If .EOF = False And .BOF = False Then
    .MoveFirst
    cboHead.Clear
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
    Me.Hide
    AddCompany.Show 1
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

