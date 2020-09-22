VERSION 5.00
Begin VB.Form EditMaster 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Master"
   ClientHeight    =   4410
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   7065
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   7065
      TabIndex        =   9
      Top             =   3810
      Width           =   7065
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         CausesValidation=   0   'False
         Height          =   300
         Left            =   3765
         TabIndex        =   8
         Top             =   60
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   300
         Left            =   2220
         TabIndex        =   7
         Top             =   60
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3435
      Left            =   180
      TabIndex        =   10
      Top             =   180
      Width           =   6675
      Begin VB.TextBox txtHead 
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
         TabIndex        =   0
         Top             =   382
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
         Left            =   1740
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
         TabIndex        =   15
         Top             =   1005
         Width           =   375
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Report Type"
         Height          =   195
         Index           =   4
         Left            =   300
         TabIndex        =   14
         Top             =   2205
         Width           =   885
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Report Position"
         Height          =   195
         Index           =   5
         Left            =   300
         TabIndex        =   13
         Top             =   2745
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Head"
         Height          =   195
         Left            =   300
         TabIndex        =   12
         Top             =   480
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Opening Balance :"
         Height          =   195
         Left            =   300
         TabIndex        =   11
         Top             =   1605
         Width           =   1320
      End
   End
End
Attribute VB_Name = "EditMaster"
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

Dim adoMasterRS As Recordset
Attribute adoMasterRS.VB_VarHelpID = -1
Dim adoSecondaryRS As Recordset
Dim IncompleteEntry As Boolean

Private Sub cboReport_Validate(Cancel As Boolean)
    If cboReport.ListIndex = -1 Then
        MsgBox "Select Report Type"
        Cancel = True
    End If
   
End Sub

Private Sub cmdSave_Click()
 ' On Error GoTo AddErr
Dim strhead As String
  Set adoSecondaryRS = New Recordset
  
If txtHead <> SelectedHead Then
strhead = Replace(txtHead, "'", "''")
  
  adoSecondaryRS.Open "select count(*)" _
    & " FROM " & MasterTable & " where actitle= '" & strhead & "'", db, adOpenStatic, adLockReadOnly, adCmdText

If adoSecondaryRS.Fields(0).Value > 0 Then
    MsgBox "Head already exists"
    txtHead.SetFocus
    Exit Sub
End If
    adoSecondaryRS.Close
End If

'If Val(txtRptOrder) = 0 Then
'  adoSecondaryRS.Open "select MAX(fnlrptposi)" _
'    & " FROM " &  MASTERTABLE  &  " ", db, adOpenStatic, adLockReadOnly, adCmdText
'
'    If adoSecondaryRS.BOF = False And adoSecondaryRS.EOF = False Then
'        txtRptOrder = adoSecondaryRS.Fields(0).Value + 1
'    Else
'        txtRptOrder = 1
'    End If
'
'  adoSecondaryRS.Close
'End If
'If Val(txtRptOrder) <> adoMasterRS!fnlrptposi Or cboReport.ItemData(cboReport.ListIndex) <> adoMasterRS!fnlrptcode Then
'  adoSecondaryRS.Open "select count(*)" _
'    & " FROM " &  MASTERTABLE  &  " where fnlrptposi=" & Val(txtRptOrder) & "and fnlrptcode=" & cboReport.ItemData(cboReport.ListIndex), db, adOpenStatic, adLockReadOnly, adCmdText
'  If adoSecondaryRS.Fields(0).Value > 0 Then
'    MsgBox "Already a Head exists in this position"
'    adoSecondaryRS.Close
'    adoSecondaryRS.Open "select max(fnlrptposi)" _
'    & " FROM " &  MASTERTABLE  &  " where  fnlrptcode=" & cboReport.ItemData(cboReport.ListIndex), db, adOpenStatic, adLockReadOnly, adCmdText
'    txtRptOrder = adoSecondaryRS.Fields(0).Value + 1
'    txtRptOrder.SetFocus
'        Exit Sub
'  End If
'
'End If
  
If Len(txtHead) = 0 Then
    'IncompleteEntry = True
    MsgBox "Head should not be empty"
    txtHead.SetFocus
    Exit Sub
ElseIf cboReport.ListIndex < 0 Then
    'IncompleteEntry = True
    MsgBox "Select a Report Type"
    cboReport.SetFocus
    Exit Sub
'Else
'    IncompleteEntry = False
End If

  
    
    LoadRecord
    adoMasterRS.Update
    ViewMasterRS.Requery
    If adoMasterRS!ACTITLE <> SelectedHead Then
        ViewMaster.cboHead.RemoveItem ViewMaster.cboHead.ListIndex
        ViewMaster.cboHead.AddItem adoMasterRS!ACTITLE
        ViewMaster.cboHead.ItemData(ViewMaster.cboHead.NewIndex) = adoMasterRS!ACNUMBER
    End If
    ViewMaster.cboHead.Text = adoMasterRS("actitle").Value
    adoMasterRS.Close

    Me.Hide
    ViewMaster.Show
  Exit Sub
'AddErr:
'  MsgBox Err.Description

End Sub



Private Sub Form_Activate()
  Set adoMasterRS = New Recordset
   
  adoMasterRS.Open "select acnumber,actitle,balancetyp,balance," _
    & "fnlrptcode,fnlrptposi FROM " & MasterTable & " where acnumber= " & SelectedRecord, db, adOpenStatic, adLockOptimistic, adCmdText
  ClearFields
  GetFields
  IncompleteEntry = False
  txtHead.SetFocus
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

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
    ViewMaster.Show
End Sub

Private Sub ClearFields()
    txtCode = ""
    txtHead = ""
    Option1.Value = True
    txtOPBAL = "0.00"
    txtOPBAL = Format(Format(Val(txtOPBAL), "0.00"), "@@@@@@@@@@@")
    cboReport.ListIndex = -1
    txtRptOrder = ""
End Sub

Public Sub LoadRecord()
With adoMasterRS
    '!ACNUMBER = Val(txtCode)
    !ACTITLE = txtHead
        
    If Option1.Value = True Then
        !balancetyp = "C"
    ElseIf Option2.Value = True Then
        !balancetyp = "D"
    End If
    
    !Balance = Val(txtOPBAL)
    
    !fnlrptcode = cboReport.ItemData(cboReport.ListIndex)
    !fnlrptposi = Val(txtRptOrder)
    
End With
End Sub


Private Sub txtHead_GotFocus()
    txtHead.SelStart = 0
    txtHead.SelLength = Len(txtHead)
End Sub

Private Sub txtHead_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtHead_Validate(Cancel As Boolean)
    If Len(Trim(txtHead)) = 0 Then
        MsgBox "Name can not be empty"
        Cancel = True
    End If
End Sub

Private Sub txtOPBAL_GotFocus()
txtOPBAL = Format(Val(txtOPBAL), "0.00")
txtOPBAL.SelStart = 0
txtOPBAL.SelLength = Len(txtOPBAL)
End Sub

Private Sub txtOPBAL_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyBack) Then Exit Sub
    If InStr("1234567890.", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtOPBAL_LostFocus()
txtOPBAL = Format(Format(Val(txtOPBAL), "0.00"), "@@@@@@@@@@@")
End Sub

Private Sub txtOPBAL_Validate(Cancel As Boolean)

If IsNumeric(txtOPBAL) = False Then
    MsgBox "Enter Zero or a valid Positive number"
    Cancel = True
End If

End Sub

Private Sub GetFields()
    txtHead = Trim(adoMasterRS("actitle").Value)
    
    txtCode = Format(adoMasterRS("acnumber").Value, "@@@@@@@@@@@")
    
    If adoMasterRS("balancetyp").Value = "C" Then
        Option1.Value = True
    ElseIf adoMasterRS("balancetyp").Value = "D" Then
        Option2.Value = True
    End If
     
    txtOPBAL = Format(Format(adoMasterRS("balance").Value, "0.00"), "@@@@@@@@@@@")
   
    Dim i As Integer
    For i = 0 To cboReport.ListCount - 1
        If cboReport.ItemData(i) = adoMasterRS!fnlrptcode Then Exit For
    Next i
    cboReport.ListIndex = i
   
    txtRptOrder = adoMasterRS("fnlrptposi").Value

End Sub


Private Sub txtRptOrder_GotFocus()
txtRptOrder = Format(Val(txtRptOrder), "0")
txtRptOrder.SelStart = 0
txtRptOrder.SelLength = Len(txtRptOrder)

End Sub


Private Sub txtRptOrder_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyBack) Then Exit Sub
    If InStr("1234567890", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtRptOrder_LostFocus()
txtRptOrder = Format(Format(Val(txtRptOrder), "0"), "@@@@@@@@@@@")
End Sub

