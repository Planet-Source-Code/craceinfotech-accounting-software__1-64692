VERSION 5.00
Begin VB.Form ViewMaster 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View Master"
   ClientHeight    =   4695
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   8100
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   8100
      TabIndex        =   17
      Top             =   4095
      Width           =   8100
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   300
         Left            =   5243
         TabIndex        =   10
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   4083
         TabIndex        =   9
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   300
         Left            =   2923
         TabIndex        =   8
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Default         =   -1  'True
         Height          =   300
         Left            =   1763
         TabIndex        =   7
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
      ScaleWidth      =   8100
      TabIndex        =   15
      Top             =   4395
      Width           =   8100
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "frmViewMaster.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "frmViewMaster.frx":0192
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "frmViewMaster.frx":0324
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "frmViewMaster.frx":04B6
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   16
         Top             =   0
         Width           =   3360
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3435
      Left            =   713
      TabIndex        =   18
      Top             =   300
      Width           =   6675
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
         Left            =   1740
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   375
         Width           =   4575
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
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   960
         Width           =   1755
      End
      Begin VB.TextBox txtRptOrder 
         DataField       =   "fnlrptposi"
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
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2700
         Width           =   1755
      End
      Begin VB.ComboBox cboReport 
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
         Left            =   1740
         Locked          =   -1  'True
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2100
         Width           =   4575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "DB"
         Height          =   495
         Left            =   5040
         TabIndex        =   4
         Top             =   1455
         Width           =   675
      End
      Begin VB.OptionButton Option1 
         Caption         =   "CR"
         Height          =   495
         Left            =   4020
         TabIndex        =   3
         Top             =   1455
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.TextBox txtOPBAL 
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
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1500
         Width           =   1755
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   23
         Top             =   1005
         Width           =   375
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Report Type"
         Height          =   195
         Index           =   4
         Left            =   300
         TabIndex        =   22
         Top             =   2205
         Width           =   885
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Report Position"
         Height          =   195
         Index           =   5
         Left            =   300
         TabIndex        =   21
         Top             =   2745
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Select Head"
         Height          =   195
         Left            =   300
         TabIndex        =   20
         Top             =   480
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Opening Balance :"
         Height          =   195
         Left            =   300
         TabIndex        =   19
         Top             =   1605
         Width           =   1320
      End
   End
End
Attribute VB_Name = "ViewMaster"
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


Dim mvBookMark As Variant
Dim cboload As Boolean

Private Sub cboHead_Click()

With ViewMasterRS
    .MoveFirst
    .Find "acnumber= " & cboHead.ItemData(cboHead.ListIndex)
    If .EOF Then
        MsgBox "Unable to find record"
        Exit Sub
    End If
    
    GetFields
End With
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
    cmdEdit.Enabled = True
    cmdDelete.Enabled = True
End If

End Sub

Private Sub Form_Load()
Dim adoReportRS As Recordset
Set adoReportRS = New Recordset

adoReportRS.Open "select rep_type,typeid from type order by rep_type", db, adOpenStatic, adLockReadOnly, adCmdText

With adoReportRS
    .MoveFirst
    Do While Not .EOF
    With cboReport
        .AddItem adoReportRS!rep_type
        .ItemData(.NewIndex) = adoReportRS!typeid
        adoReportRS.MoveNext
    End With
    Loop
End With

adoReportRS.Close
Set adoReportRS = Nothing
  
  Set ViewMasterRS = New Recordset
  ViewMasterRS.Open "select acnumber,actitle,balancetyp,balance," _
    & "fnlrptcode,fnlrptposi from " & MasterTable & " order by actitle", db, adOpenStatic, adLockReadOnly, adCmdText
    cboload = False
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
    cboload = True
End If
End With
  ClearFields
    With ViewMasterRS
    If .EOF = False And .BOF = False Then
        ViewMasterRS.MoveFirst
        cboHead.Text = ViewMasterRS("actitle").Value
        GetFields
    End If
    End With

End Sub

Private Sub Form_Resize()
  On Error Resume Next
  lblStatus.Width = Me.Width - 1500
  cmdNext.Left = lblStatus.Width + 700
  cmdLast.Left = cmdNext.Left + 340
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub cmdAdd_Click()
    Me.Hide
    AddMaster.Show 1
End Sub

Private Sub cmdDelete_Click()
SelectedRecord = ViewMasterRS!ACNUMBER
SelectedHead = ViewMasterRS!ACTITLE

With ViewMasterRS
    mvBookMark = .Bookmark
    .MoveNext
    If .EOF Then .MoveLast
    NextRecord = !ACNUMBER
    .Bookmark = mvBookMark
End With

    Me.Hide
    DeleteMaster.Show 1
End Sub
Private Sub cmdEdit_Click()
  'On Error GoTo EditErr


With ViewMasterRS
    mvBookMark = .Bookmark
    SelectedRecord = !ACNUMBER
    SelectedHead = Trim(!ACTITLE)
End With

    Me.Hide
    EditMaster.Show 1
  
End Sub
Private Sub cmdClose_Click()
    Dim f As Form
    For Each f In Forms
        If (f.Name <> Me.Name) And (f.Name <> frmMain.Name) Then
            Unload f
        End If
    Next f

  Unload Me
  'frmMain.Show
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError

  ViewMasterRS.MoveFirst
  cboHead.Text = ViewMasterRS("actitle").Value
  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  ViewMasterRS.MoveLast
  cboHead.Text = ViewMasterRS("actitle").Value
  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  On Error GoTo GoNextError

  If Not ViewMasterRS.EOF Then ViewMasterRS.MoveNext
  If ViewMasterRS.EOF And ViewMasterRS.RecordCount > 0 Then
    Beep
     'moved off the end so go back
    ViewMasterRS.MoveLast
  End If
  'show the current record
  cboHead.Text = ViewMasterRS("actitle").Value
  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError

  If Not ViewMasterRS.BOF Then ViewMasterRS.MovePrevious
  If ViewMasterRS.BOF And ViewMasterRS.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    ViewMasterRS.MoveFirst
  End If
  'show the current record
  'mbDataChanged = False
  cboHead.Text = ViewMasterRS("actitle").Value
  Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub ClearFields()
    txtCode = ""
   ' txtHead = ""
   ' txtCRDB = ""
   Option1.Value = True
    txtOPBAL = ""
   ' txtRptType = ""
   cboReport.ListIndex = -1
    txtRptOrder = ""
End Sub

Private Sub GetFields()
    txtCode = Format(ViewMasterRS("acnumber").Value, "@@@@@@@@@@@")
    
    If ViewMasterRS("balancetyp").Value = "C" Then
        Option1.Value = True
    ElseIf ViewMasterRS("balancetyp").Value = "D" Then
        Option2.Value = True
    End If
     
    txtOPBAL = Format(Format(ViewMasterRS("balance").Value, "0.00"), "@@@@@@@@@@@")
   
Dim i As Integer

For i = 0 To cboReport.ListCount - 1
If cboReport.ItemData(i) = ViewMasterRS!fnlrptcode Then Exit For
Next i
cboReport.ListIndex = i
    
    txtRptOrder = Format(ViewMasterRS("fnlrptposi").Value, "@@@@@@@@@@@")
        'txtRptOrder = ViewMasterRS("fnlrptcode").Value
    lblStatus.Caption = "Record: " & CStr(ViewMasterRS.AbsolutePosition) & " of " & CStr(ViewMasterRS.RecordCount)
End Sub

