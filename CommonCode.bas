Attribute VB_Name = "CommonCode"
Option Explicit
'***********************************************
'This Software is developed by craceinfotech.
'Web site : http://www.craceinfotech.com
'email    : craceinfotech.yahoo.com
'date     : 18.03.2006
'***********************************************



Public db As Connection
Public ViewCompanyRS As Recordset
Public ViewMasterRS As Recordset
Public SelectedRecord As Long
Public SelectedHead As String
Public SelectedCompany  As Integer
Public NextRecord As Long
Public CurrentDate As Date
Public StartDate As Date
Public EndDate As Date
Public FromDate As Date
Public ToDate As Date
Public SystemDate As Date
Public ContinueProcess As Boolean
Public CompanyName As String
Public CompanyYear As String
Public FinancialYear As String
Public StockInHand As Currency
Public CompanySelected As Boolean
Public MasterTable As String
Public TransactionTable As String
Public CashInHand As Currency
Public EnabledDate As Date
Public LoginSucceeded As Boolean
Public ZipFile As String

  Public Type DOCINFO
      pDocName As String
      pOutputFile As String
      pDatatype As String
  End Type

  Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal _
     hPrinter As Long) As Long
  Public Declare Function EndDocPrinter Lib "winspool.drv" (ByVal _
     hPrinter As Long) As Long
  Public Declare Function EndPagePrinter Lib "winspool.drv" (ByVal _
     hPrinter As Long) As Long
  Public Declare Function OpenPrinter Lib "winspool.drv" Alias _
     "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, _
      ByVal pDefault As Long) As Long
  Public Declare Function StartDocPrinter Lib "winspool.drv" Alias _
     "StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, _
     pDocInfo As DOCINFO) As Long
  Public Declare Function StartPagePrinter Lib "winspool.drv" (ByVal _
     hPrinter As Long) As Long
  Public Declare Function WritePrinter Lib "winspool.drv" (ByVal _
     hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, _
     pcWritten As Long) As Long
Declare Function AbortDoc Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function AbortPrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public lhPrinter As Long

Sub Main()

ChDrive "c"
ChDir "c:\VBPROG\VBA\DAT"
Dim i As Boolean
'Stop
    Set db = New Connection
    db.CursorLocation = adUseClient
    db.Open "PROVIDER=MSDASQL;dsn=VBA;uid=;pwd=;"
    
    
    
'CompanyName = Trim(ViewCompanyRS!cotitle)
'frmMain.Caption = CompanyName
'StartDate = ViewCompanyRS!costart
'EndDate = ViewCompanyRS!coend
'FromDate = StartDate
'CurrentDate = Format(Now(), "dd/mm/yyyy")
'If CurrentDate > EndDate Then CurrentDate = EndDate
'ToDate = CurrentDate
EnabledDate = Format("28/10/2007", "dd/mm/yyyy")
SystemDate = Format(Now(), "dd/mm/yyyy")
If SystemDate > EnabledDate Then
    MsgBox "Trial period is over"
    End
End If


LoginSucceeded = False
LoginSucceeded = True 'Password protection disabled.
'frmLogin.Show 1 Password protection disabled.
If LoginSucceeded Then
  'ViewCompany.Show 1
  ViewCompany.cmdSelect.Value = True
Else
    MsgBox "Incorrect Username or Password"
    End
End If

End Sub



Public Function PADC(Gstring As String, Glength As Integer)
    PADC = Space((Glength - Len(Gstring)) \ 2) & Gstring
End Function
Public Sub PrintText(sWrittenData As String)
Dim lReturn As Long
Dim lDoc As Long
Dim MyDocInfo As DOCINFO
Dim lpcWritten As Long
'Dim sWrittenData As String
'cmdPrint.Enabled = False
lReturn = OpenPrinter(Printer.DeviceName, lhPrinter, 0)
      MyDocInfo.pDocName = "Accounts Report"
      MyDocInfo.pOutputFile = vbNullString
      MyDocInfo.pDatatype = vbNullString
      lDoc = StartDocPrinter(lhPrinter, 1, MyDocInfo)
      Call StartPagePrinter(lhPrinter)
lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
                Len(sWrittenData), lpcWritten)
lReturn = EndPagePrinter(lhPrinter)
lReturn = EndDocPrinter(lhPrinter)
lReturn = ClosePrinter(lhPrinter)
sWrittenData = ""
End Sub



Public Function ZeroSup(Gnumber As Currency)
    ZeroSup = IIf(Gnumber > 0, Format(Format(Gnumber, "0.00"), "@@@@@@@@@@@@@@"), Space(14))
End Function

Public Sub CalculateCashInHand()
    Dim Temprs As Recordset
    Set Temprs = New Recordset
    Temprs.Open "select BALANCETYP,BALANCE FROM " & MasterTable & " where actitle='CASH'" _
     , db, adOpenStatic, adLockReadOnly, adCmdText
    If Temprs.EOF = False And Temprs.BOF = False Then
        If Temprs!balancetyp = "D" Then
            CashInHand = Temprs!Balance
        Else
            CashInHand = -(Temprs!Balance)
        End If
        
    Else
        CashInHand = 0
    End If
    Temprs.Close

    Set Temprs = New Recordset
    
    Temprs.Open "select sum(credit-debit) FROM " & TransactionTable & " where " _
     & "acn_date between {" & Format(StartDate, "mm/dd/yyyy") & "} and {" & Format(CurrentDate, "mm/dd/yyyy") & "} ", _
     db, adOpenStatic, adLockReadOnly, adCmdText
     
     With Temprs
        If Not (.BOF = True And .EOF = True) Then
            CashInHand = CashInHand + .Fields(0).Value
        End If
        .Close
    End With
    
    Set Temprs = Nothing

End Sub
Public Sub ShowMainCaption()
    CalculateCashInHand
    frmMain.Caption = "www.craceinfotech.com    " + Trim(ViewCompanyRS("cotitle").Value) + " - (" + ViewCompanyRS("coyear").Value + ") "
    'Account Date: " + Format(CurrentDate, "dd/mm/yyyy") + "     Cash In Hand Rs.   " & Format(CashInHand, "0.00")
End Sub
Public Sub CalculateToDate()
    Dim Temprs As Recordset
    Set Temprs = New Recordset
        
    Temprs.Open "select MAX(ACN_DATE)" _
    & " from  " & TransactionTable, db, adOpenStatic, adLockReadOnly, adCmdText
    If Temprs.RecordCount > 0 Then
        If IsDate(Temprs.Fields(0).Value) Then
            CurrentDate = Format(Temprs.Fields(0).Value, "dd/mm/yyyy")
        ToDate = CurrentDate
        End If
    End If
    Temprs.Close
End Sub



Public Sub UpdateStock()
 Dim Temprs As Recordset
    Set Temprs = New Recordset
    Temprs.Open "select conumber,stock from company where conumber=" & SelectedCompany _
     , db, adOpenStatic, adLockOptimistic, adCmdText
    
    If Temprs.EOF = False And Temprs.BOF = False Then
        Temprs!stock = StockInHand
        Temprs.Update
    End If
    Temprs.Close
End Sub
