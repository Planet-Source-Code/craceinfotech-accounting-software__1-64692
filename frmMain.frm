VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "www.craceinfotech.com"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuCompany 
      Caption         =   "&Company"
   End
   Begin VB.Menu mnuMaster 
      Caption         =   "&Masters"
   End
   Begin VB.Menu mnuEntry 
      Caption         =   "&Entries"
      Begin VB.Menu mnuEntryDate 
         Caption         =   "Account &Date"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuEntryTrans 
         Caption         =   "&Transactions"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "&Reports"
      Begin VB.Menu mnuReportsCashDaybook 
         Caption         =   "&Daybook"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuReportsLedger 
         Caption         =   "&Ledger (Particular)"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuReportsTrial 
         Caption         =   "&Trial Balance"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuReportsTrading 
         Caption         =   "Trading &Account"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuReportsProfit 
         Caption         =   "&Profit && Loss Account "
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuReportsBalance 
         Caption         =   "&Balance Sheet"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuReportsAllLedger 
         Caption         =   "&Ledger (All)"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "frmMain"
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

Private Sub mnuCompany_Click()
    ViewCompany.Show 1
End Sub

Private Sub mnuEntryDate_Click()
    AcnDate.Show 1
End Sub

Private Sub mnuEntryJournal_Click()
    ViewJournalTransaction.Show 1
    'ViewTransaction1.Show 1
End Sub

Private Sub mnuEntryTrans_Click()
    ViewCashTransaction.Show 1
End Sub

Private Sub mnuExit_Click()
    Dim f As Form
    For Each f In Forms
        If f.Name <> Me.Name Then Unload f
    Next f
    
    Unload Me
    
    
    
    'End
End Sub

Private Sub mnuMaster_Click()
    'ViewMaster.Left = (Me.ScaleWidth - ViewMaster.Width) / 2
    'ViewMaster.Top = (Me.ScaleHeight - ViewMaster.Height) / 2
    ViewMaster.Show 1
End Sub

Private Sub mnuReportsAllLedger_Click()
    Ledger1.Show 1
End Sub

Private Sub mnuReportsBalance_Click()
    AsOnBalance.Show 1
End Sub

Private Sub mnuReportsCashDaybook_Click()
    FromTo.Show 1
    If ContinueProcess Then
        CashDaybook.Show 1
    End If
End Sub
Private Sub mnuReportsJournalDaybook_Click()
    FromTo.Show 1
    If ContinueProcess Then
        JournalDaybook.Show 1
    End If
End Sub
Private Sub mnuReportsLedger_Click()
    FromToHead.Show 1
End Sub

Private Sub mnuReportsProfit_Click()
    AsOnProfit.Show 1
End Sub

Private Sub mnuReportsTrading_Click()
    AsOnTrading.Show 1
End Sub

Private Sub mnuReportsTrial_Click()
    AsOnTrial.Show 1
End Sub
