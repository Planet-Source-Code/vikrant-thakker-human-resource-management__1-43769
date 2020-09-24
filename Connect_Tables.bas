Attribute VB_Name = "Connect_Tables"
Public rsEmp As ADODB.Recordset 'connecting with table mastemployee
Public rsRetired As ADODB.Recordset
Public rsCaste As ADODB.Recordset 'connecting with table mastcaste
Public rsClass As ADODB.Recordset
Public rsDesig As ADODB.Recordset
Public rsType As ADODB.Recordset
Public rsSection As ADODB.Recordset
Public rsAttd As ADODB.Recordset
Public rsPS As ADODB.Recordset
Public rsWD As ADODB.Recordset
Public rsPD As ADODB.Recordset
Public rsLS As ADODB.Recordset  ' rsLS is for table LeaveSlip
Public rsTmpLS As ADODB.Recordset ' for table TempLeave
Public rsExpiry As ADODB.Recordset
Public rsIncome As ADODB.Recordset  'Connects to ProfitLoss table
Public rsPrintAttd As ADODB.Recordset
Public rsMntProfitLoss As ADODB.Recordset
Public rsICard As ADODB.Recordset
Public rsPrintPaySlip As ADODB.Recordset

Public Sub ConnectTables()
Set rsClass = New ADODB.Recordset
rsClass.Open "select * from MastClass", conn, adOpenStatic, adLockOptimistic

Set rsType = New ADODB.Recordset
rsType.Open "select * from MastType", conn, adOpenStatic, adLockOptimistic

Set rsEmp = New ADODB.Recordset
rsEmp.Open "select * from MastEmployee", conn, adOpenStatic, adLockOptimistic

Set rsICard = New ADODB.Recordset
rsICard.Open "select * from ICard", conn, adOpenStatic, adLockOptimistic

Set rsMntProfitLoss = New ADODB.Recordset
rsMntProfitLoss.Open "select * from MonthlyProfitLoss", conn, adOpenStatic, adLockOptimistic

Set rsPrintPaySlip = New ADODB.Recordset
rsPrintPaySlip.Open "select * from PaySlip", conn, adOpenStatic, adLockOptimistic

Set rsRetired = New ADODB.Recordset
rsRetired.Open "select * from MastRetiredEmp", conn, adOpenStatic, adLockOptimistic

Set rsWD = New ADODB.Recordset
rsWD.Open "select * from MastWorkingDays", conn, adOpenStatic, adLockOptimistic

Set rsCaste = New ADODB.Recordset
rsCaste.Open "select * from MastCaste", conn, adOpenStatic, adLockOptimistic

Set rsTmpLS = New ADODB.Recordset
rsTmpLS.Open "select * from TempLeave", conn, adOpenStatic, adLockOptimistic

Set rsPS = New ADODB.Recordset
rsPS.Open "select * from PaySlip", conn, adOpenStatic, adLockOptimistic

Set rsAttd = New ADODB.Recordset
rsAttd.Open "select * from MastAttendance", conn, adOpenStatic, adLockOptimistic

Set rsPrintAttd = New ADODB.Recordset
rsPrintAttd.Open "select * from PrintAttn", conn, adOpenStatic, adLockOptimistic


Set rsSection = New ADODB.Recordset
rsSection.Open "select * from MastSection", conn, adOpenStatic, adLockOptimistic

Set rsPD = New ADODB.Recordset
rsPD.Open "select * from PayDeduction", conn, adOpenStatic, adLockOptimistic

Set rsDesig = New ADODB.Recordset
rsDesig.Open "select * from MastDesignation", conn, adOpenStatic, adLockOptimistic

Set rsEmp = New ADODB.Recordset
rsEmp.Open "select * from MastEmployee", conn, adOpenStatic, adLockOptimistic

Set rsLS = New ADODB.Recordset
rsLS.Open "select * from LeaveMast", conn, adOpenStatic, adLockOptimistic

Set rsPL = New ADODB.Recordset
rsPL.Open "select * from ProfitLoss", conn, adOpenStatic, adLockOptimistic


Set rsExpiry = New ADODB.Recordset
rsExpiry.Open "select * from MastExpiry", conn, adOpenStatic, adLockOptimistic

End Sub
