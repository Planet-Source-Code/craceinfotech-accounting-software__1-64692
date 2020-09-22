VERSION 5.00
Begin VB.Form RestoreCompany 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Restore Company"
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
      TabIndex        =   6
      Top             =   4095
      Width           =   11130
      Begin VB.CommandButton cmdRestore 
         Caption         =   "&Restore"
         Height          =   300
         Left            =   4434
         TabIndex        =   15
         ToolTipText     =   "To Restore  a Company's Data from backup"
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Quit"
         Height          =   300
         Left            =   5602
         TabIndex        =   1
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
      TabIndex        =   4
      Top             =   4395
      Width           =   11130
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   5
         Top             =   0
         Width           =   3360
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   278
      TabIndex        =   7
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   2
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
         TabIndex        =   3
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
         TabIndex        =   12
         Top             =   1155
         Width           =   375
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Start Date"
         Height          =   195
         Index           =   4
         Left            =   300
         TabIndex        =   11
         Top             =   1838
         Width           =   720
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "End Date"
         Height          =   195
         Index           =   5
         Left            =   7320
         TabIndex        =   10
         Top             =   1838
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Company"
         Height          =   195
         Left            =   300
         TabIndex        =   9
         Top             =   480
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Year"
         Height          =   195
         Left            =   7320
         TabIndex        =   8
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
Attribute VB_Name = "RestoreCompany"
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
Dim RestorecompanyRS As Recordset

Private Sub cboHead_Click()
lblStatus.Caption = cboHead.Text

With RestorecompanyRS
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

Private Sub RestoreComp()
On Error GoTo DeleteErr
    Dim intIndex As Integer
    'If MsgBox("Restore this Company?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
        Dim Temprs As Recordset
        Set Temprs = New Recordset
        Temprs.Open "select conumber,covalid from company a where a.conumber=" & RestorecompanyRS!CONUMBER, db, adOpenStatic, adLockOptimistic, adCmdText
        Temprs("covalid").Value = "Y"
        Temprs.Update
        Temprs.Close
        Set Temprs = Nothing
        RestorecompanyRS.Requery
        intIndex = cboHead.ListIndex
        cboHead.RemoveItem intIndex
        If intIndex = cboHead.ListCount Then
            cboHead.ListIndex = intIndex - 1
        Else
            cboHead.ListIndex = intIndex
        End If
    'End If
    Exit Sub
DeleteErr:
    MsgBox "Some Errors occured. Please contact Mr.R.Natarajan"
 End Sub

Private Sub cmdRestore_Click()
Dim retval As Long
Dim fLen As Integer, filepath As String

If cboHead.ListCount = 0 Then
    ClearFields
    MsgBox "No Companys to restore."
    Exit Sub
End If

    CompanyName = Trim(RestorecompanyRS!cotitle)
    CompanyYear = RestorecompanyRS("coyear").Value

    MasterTable = "C:\VBPROG\VBA\DAT\MASTER" + Right(Format(Format(SelectedCompany, "000"), "@@@"), 2) + ".DBF"
    TransactionTable = "C:\VBPROG\VBA\DAT\ENTRY" + Right(Format(Format(SelectedCompany, "000"), "@@@"), 2) + ".DBF"
    ZipFile = "BACKUP" + Right(Format(Format(SelectedCompany, "000"), "@@@"), 2) + ".ZIP"

    'filepath = "d:\natarajan\test\backups\" & ZipFile
    filepath = "a:\" & ZipFile
    On Error Resume Next
    fLen = Len(Dir$(filepath))
    If Err Or fLen = 0 Then
        MsgBox "Backup File Missing" & vbCrLf & "Unable to Restore "
        Exit Sub
    End If

    'retval = ExecCmd("C:\Program Files\WinZip\WINZIP32.EXE -e d:\natarajan\test\backups\" & ZipFile & " C:\VBPROG\VBA\DAT")
    retval = ExecCmd("C:\Program Files\WinZip\WINZIP32.EXE -e a:\" & ZipFile & " C:\VBPROG\VBA\DAT")
    'retval = ExecCmd("C:\Program Files\WinZip\WINZIP32.EXE -e d:\natarajan\test\backups\backup.zip d:\natarajan\test\backups\test")

    fLen = Len(Dir$(MasterTable))
    If Err Or fLen = 0 Then
        MsgBox "Restore Failure"
        Exit Sub
    End If

    fLen = Len(Dir$(TransactionTable))
    If Err Or fLen = 0 Then
        MsgBox "Restore Failure"
        Exit Sub
    Else
        RestoreComp
        MsgBox "Restore Over "
If cboHead.ListCount = 0 Then
    ClearFields
    MsgBox "No More Companys to restore."
    cmdClose.Value = True
End If
    End If


    cboHead.SetFocus
End Sub


Private Sub Form_Activate()
If cboHead.ListCount = 0 Then
    ClearFields
    MsgBox "No Companys to restore."
        cmdClose.Value = True
    
Else
    If Not CompanySelected Then
'        cmdEdit.Enabled = True
'        cmdDelete.Enabled = True
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
  
  Set RestorecompanyRS = New Recordset
  RestorecompanyRS.Open "select conumber,cotitle,coyear,costart,coend,covalid" _
    & " from  company where covalid='N' order by cotitle, coyear", db, adOpenStatic, adLockReadOnly, adCmdText
    cboload = False
  With RestorecompanyRS
If .EOF = False And .BOF = False Then
    .MoveFirst
    cboHead.Clear
    Do While Not .EOF
        With cboHead
            .AddItem Trim(RestorecompanyRS("cotitle").Value) + " - (" + RestorecompanyRS("coyear").Value + ")"
            .ItemData(.NewIndex) = RestorecompanyRS("conumber").Value
        End With
        .MoveNext
    Loop
    .MoveFirst
    cboload = True
Else
    MsgBox "No Companys to Restore"
    cmdClose.Value = True
End If
End With
  ClearFields
    With RestorecompanyRS
    If .EOF = False And .BOF = False Then
        RestorecompanyRS.MoveFirst
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

    txtCode = Format(RestorecompanyRS("conumber").Value, "@@@@@@@@@@")
    txtYear = RestorecompanyRS!COYear
    txtStart = Format(RestorecompanyRS!costart, "dd/mm/yyyy")
    txtEnd = Format(RestorecompanyRS!COend, "dd/mm/yyyy")

End Sub

