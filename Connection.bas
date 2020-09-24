Attribute VB_Name = "Connection"
Public conn As ADODB.Connection
Dim str1 As String
Public FormName As String

Public Sub Main()
On Error GoTo merr
Set conn = New ADODB.Connection
Call Office97
Exit Sub
merr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub Office97()
On Err GoTo Off2000   ' If an error occurs while connecting with Office97 then try to Connect with Office2000
str1 = "provider=microsoft.jet.oledb.3.51;data source=" 'use for connecting with database file
str1 = str1 & App.Path & "\Project97.mdb" 'exe and mdb file should  be in the same folder
conn.Open str1
Call ConnectTables   ' All Recordsets are declared in this function (declared in Connect_Tables module)
Call ExpiryCheck  ' Checking and Updating of Software Expiry (declared in Check_Expiry module)
'Load frmControl
'frmControl.Show
Load frmSplash1
frmSplash1.Show
Exit Sub
Off2000:
Call Office2000
End Sub

Private Sub Office2000()
On Err GoTo Err2000
str1 = "provider=microsoft.jet.oledb.4.0;data source=" 'use for connecting with database file
str1 = str1 & App.Path & "\Project97.mdb" 'exe and mdb file should  be in the same folder
conn.Open str1
Call ConnectTables
Call ExpiryCheck
'Load frmControl
'frmControl.Show
Load frmSplash1
frmSplash1.Show
Exit Sub
Err2000:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Public Sub UnLoadAll()
'This will remove all the forms from the memory
Unload frmAbout
Unload frmAttn
Unload frmControl
Unload frmIncomeExpense
Unload frmLeaveSlip
Unload frmMain
Unload frmMaster
Unload frmMastEmp
Unload frmMastWD
Unload frmPay
Unload frmPayDeduct
Unload frmProfitLossMain
Unload frmReports
Unload frmSplash1
End Sub
