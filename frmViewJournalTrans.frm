VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form ViewJournalTransaction 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   11520
      TabIndex        =   1
      Top             =   6090
      Width           =   11520
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Default         =   -1  'True
         Height          =   300
         Left            =   3473
         TabIndex        =   5
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   300
         Left            =   4633
         TabIndex        =   4
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   5793
         TabIndex        =   3
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   300
         Left            =   6953
         TabIndex        =   2
         Top             =   120
         Width           =   1095
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdTransact 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   11235
      _ExtentX        =   19817
      _ExtentY        =   10186
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      FocusRect       =   2
      HighLight       =   2
      AllowUserResizing=   1
      FormatString    =   $"frmViewJournalTrans.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "ViewJournalTransaction"
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

Public ViewTransactionRS As Recordset
Attribute ViewTransactionRS.VB_VarHelpID = -1

Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Private Sub Form_Activate()
    Static ShownFirstTime As Boolean
    Me.Caption = "Transactions for " & Format(CurrentDate, "dd/mm/yyyy")
    If grdTransact.Rows = 1 Then
    
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    If Not ShownFirstTime Then
        If MsgBox("No Transactions for this date!" & vbCrLf & "Do you want to add now?", vbYesNo + vbDefaultButton1 + vbInformation, "Add Master") = vbYes Then
            cmdAdd.Value = True
        End If
    End If
Else
    cmdEdit.Enabled = True
    cmdDelete.Enabled = True
    grdTransact.SetFocus
    grdTransact.TopRow = grdTransact.Row
End If
    ShownFirstTime = True
    End Sub

Private Sub Form_Load()

'    CurrentDate = #4/25/2001#
    
    'Set ViewTransactionRS = New Recordset
    'ViewTransactionRS.Open "select entry_id,actitle,debit,credit,particular FROM " & TransactionTable & " entries," & MasterTable & " master where entries.acnumber=master.acnumber and acn_date={" & Format(CurrentDate, "mm/dd/yyyy") & "} order by entry_id", db, adOpenStatic, adLockReadOnly
    RefreshTransaction
    FillGridTrans
    mbDataChanged = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub


Private Sub grdtransact_Click()
'grdTransact.Col = 0
'SelectedRecord = grdTransact.Text
'frmentry.Show 1
End Sub

Private Sub cmdAdd_Click()
    grdTransact.Col = 0
    If grdTransact.Rows > 1 Then
        SelectedRecord = grdTransact.Text
    Else
        SelectedRecord = 0
    End If
    Me.Hide
    AddJournalTransaction.Show 1
End Sub

Private Sub cmdDelete_Click()
    grdTransact.Col = 0
    SelectedRecord = grdTransact.Text
    Me.Hide
    DeleteTransaction.Show 1


'SelectedRecord = ViewMasterRS!acnumber
'SelectedHead = ViewMasterRS!ACTITLE
'
'With ViewMasterRS
'    mvBookMark = .Bookmark
'    .MoveNext
'    If .EOF Then .MoveLast
'    NextRecord = !acnumber
'    .Bookmark = mvBookMark
'End With
'
'    Me.Hide
'    DeleteMaster.Show 1
End Sub
Private Sub cmdEdit_Click()
  'On Error GoTo EditErr
    grdTransact.Col = 0
    SelectedRecord = grdTransact.Text
    Me.Hide
    EditTransaction.Show 1

'
'With ViewMasterRS
'    mvBookMark = .Bookmark
'    SelectedRecord = !ACNUMBER
'    SelectedHead = Trim(!ACTITLE)
'End With
'
'    Me.Hide
'    EditMaster.Show 1
  
End Sub
Private Sub cmdClose_Click()
'Dim f As Form
'For Each f In Forms
'    If (f.Name <> Me.Name) Or (f.Name = frmMain.Name) Then Unload f
'Next f
  Unload Me
  'frmMain.Show
End Sub


Public Sub FillGridTrans()
    Dim recordcounter As Integer
    Dim i As Integer
    grdTransact.Cols = ViewTransactionRS.Fields.Count
    grdTransact.Rows = ViewTransactionRS.RecordCount + 1
    recordcounter = ViewTransactionRS.RecordCount
  
  With ViewTransactionRS
    If .EOF = False And .BOF = False Then
    .MoveFirst
    grdTransact.Row = 1
    For i = 1 To recordcounter
        grdTransact.Row = i
        
        grdTransact.Col = 0
        grdTransact.Text = ViewTransactionRS.Fields(0).Value
   
        grdTransact.Col = 1
        grdTransact.Text = ViewTransactionRS.Fields(1).Value
    
        grdTransact.Col = 2
        grdTransact.Text = Format(ViewTransactionRS.Fields(2).Value, "0.00")
    
        grdTransact.Col = 3
        grdTransact.Text = Format(ViewTransactionRS.Fields(3).Value, "0.00")
       
        grdTransact.Col = 4
        grdTransact.Text = ViewTransactionRS.Fields(4).Value
      
        .MoveNext
  Next i
  End If
  End With
  
  If SelectedRecord > 0 Then
  With grdTransact
    
    .Col = 0
    For i = 1 To .Rows - 1
        .Row = i
        If Val(.Text) = SelectedRecord Then Exit For
    Next i
  End With
  End If
  'grdTransact.Row = 1
  grdTransact.Col = 0

End Sub

Public Sub RefreshTransaction()
    Set ViewTransactionRS = New Recordset
    ViewTransactionRS.Open "select entry_id,actitle,credit,debit,particular FROM " & TransactionTable & " entries," & MasterTable & " master where entries.acnumber=master.acnumber and acn_date={" & Format(CurrentDate, "mm/dd/yyyy") & "} and cj= '" & "J" & "' order by entry_id", db, adOpenStatic, adLockReadOnly
End Sub
