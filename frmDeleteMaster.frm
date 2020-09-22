VERSION 5.00
Begin VB.Form DeleteMaster 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete Master"
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
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
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
         Width           =   675
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
         TabStop         =   0   'False
         Top             =   2700
         Width           =   495
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
         TabStop         =   0   'False
         Top             =   2100
         Width           =   4575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "DB"
         Height          =   495
         Left            =   5040
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1455
         Width           =   675
      End
      Begin VB.OptionButton Option1 
         Caption         =   "CR"
         Height          =   495
         Left            =   4020
         TabIndex        =   3
         TabStop         =   0   'False
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
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
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
Attribute VB_Name = "DeleteMaster"
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

Private Sub cmdSave_Click()
 ' On Error GoTo AddErr
    
    With ViewMaster
        If .cboHead.Text <> SelectedHead Then
        .cboHead.Text = SelectedHead
    End If
    End With
    

    ViewMaster.cboHead.RemoveItem (ViewMaster.cboHead.ListIndex)

    adoMasterRS.Delete
    adoMasterRS.Close
    
    
    
    With ViewMasterRS
        .Requery
        If .BOF = False And .EOF = False Then
        .Find "acnumber=" & NextRecord
        If .EOF Then .MoveLast
        ViewMaster.cboHead.Text = !ACTITLE
        End If
    End With
    
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


Private Sub GetFields()
    txtHead = ViewMasterRS("actitle").Value
    
    txtCode = Format(ViewMasterRS("acnumber").Value, "@@@@")
    
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
   
    txtRptOrder = ViewMasterRS("fnlrptposi").Value

End Sub


