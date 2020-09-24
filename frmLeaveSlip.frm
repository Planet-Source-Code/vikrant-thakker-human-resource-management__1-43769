VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmLeaveSlip 
   BackColor       =   &H00000000&
   Caption         =   "Leave Slip"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0041E9D8&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   8700
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   7500
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H0041E9D8&
      Caption         =   "E&xit"
      Height          =   420
      Left            =   5415
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7575
      Width           =   870
   End
   Begin VB.Frame Framebutton 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   2625
      TabIndex        =   0
      Top             =   6765
      Width           =   6495
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H0080C0FF&
         Caption         =   "Cancel"
         Height          =   420
         Left            =   2790
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdModify 
         BackColor       =   &H0080C0FF&
         Caption         =   "Modify"
         Height          =   420
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   60
         Width           =   915
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H0080C0FF&
         Caption         =   "Save"
         Height          =   420
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H0080C0FF&
         Caption         =   "Next"
         Height          =   420
         Left            =   5535
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdPrev 
         BackColor       =   &H0080C0FF&
         Caption         =   "Previous"
         Height          =   420
         Left            =   4590
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   45
         Width           =   915
      End
      Begin VB.CommandButton cmdRemove 
         BackColor       =   &H0080C0FF&
         Caption         =   "Remove"
         Height          =   420
         Left            =   3690
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H0080C0FF&
         Caption         =   "Add"
         Height          =   420
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   45
         Width           =   870
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   5550
      Left            =   1185
      TabIndex        =   9
      Top             =   720
      Width           =   9405
      Begin VB.CommandButton cmdHelp 
         BackColor       =   &H0041E9D8&
         Caption         =   "&List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3510
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   240
         Width           =   780
      End
      Begin VB.TextBox txtTo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   1620
         MaxLength       =   20
         TabIndex        =   21
         Top             =   2745
         Width           =   1860
      End
      Begin VB.TextBox txtFrom 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   1620
         MaxLength       =   20
         TabIndex        =   20
         Top             =   2115
         Width           =   1860
      End
      Begin VB.ComboBox cmbLType 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1620
         TabIndex        =   19
         Top             =   1530
         Width           =   1860
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1620
         MaxLength       =   5
         TabIndex        =   13
         Top             =   225
         Width           =   1860
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   1620
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   12
         Top             =   855
         Width           =   1860
      End
      Begin VB.TextBox txtSalary 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   7260
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   11
         Top             =   855
         Width           =   1860
      End
      Begin VB.TextBox txtDesig 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   7260
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   10
         Top             =   225
         Width           =   1860
      End
      Begin RichTextLib.RichTextBox txtReason 
         Height          =   1230
         Left            =   1620
         TabIndex        =   25
         Top             =   3375
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   2170
         _Version        =   393217
         BackColor       =   16777215
         BorderStyle     =   0
         TextRTF         =   $"frmLeaveSlip.frx":0000
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "DD/MM/YYYY"
         ForeColor       =   &H00EBCCB4&
         Height          =   375
         Left            =   3600
         TabIndex        =   32
         Top             =   2835
         Width           =   1230
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "DD/MM/YYYY"
         ForeColor       =   &H00EBCCB4&
         Height          =   375
         Left            =   3600
         TabIndex        =   31
         Top             =   2220
         Width           =   1230
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL LEAVES"
         ForeColor       =   &H00EBCCB4&
         Height          =   330
         Left            =   -540
         TabIndex        =   29
         Top             =   4980
         Width           =   1905
      End
      Begin VB.Label lblTotDays 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   1620
         TabIndex        =   28
         Top             =   4905
         Width           =   1320
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TO (DATE)"
         ForeColor       =   &H00EBCCB4&
         Height          =   330
         Left            =   75
         TabIndex        =   24
         Top             =   2880
         Width           =   1245
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FROM"
         ForeColor       =   &H00EBCCB4&
         Height          =   330
         Left            =   75
         TabIndex        =   23
         Top             =   2250
         Width           =   1245
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "REASON"
         ForeColor       =   &H00EBCCB4&
         Height          =   330
         Left            =   75
         TabIndex        =   22
         Top             =   3510
         Width           =   1245
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "LEAVE TYPE"
         ForeColor       =   &H00EBCCB4&
         Height          =   330
         Left            =   75
         TabIndex        =   18
         Top             =   1620
         Width           =   1245
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "EMP. CODE"
         ForeColor       =   &H00EBCCB4&
         Height          =   330
         Left            =   75
         TabIndex        =   17
         Top             =   360
         Width           =   1245
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "NAME"
         ForeColor       =   &H00EBCCB4&
         Height          =   330
         Left            =   75
         TabIndex        =   16
         Top             =   990
         Width           =   1245
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "SALARY"
         ForeColor       =   &H00EBCCB4&
         Height          =   330
         Left            =   5715
         TabIndex        =   15
         Top             =   990
         Width           =   1245
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DESIGNATION"
         ForeColor       =   &H00EBCCB4&
         Height          =   330
         Left            =   5715
         TabIndex        =   14
         Top             =   360
         Width           =   1245
      End
   End
   Begin MSDataGridLib.DataGrid Dg 
      Height          =   7260
      Left            =   1830
      TabIndex        =   27
      Top             =   135
      Visible         =   0   'False
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   12806
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   0
      ForeColor       =   12648447
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Employee List"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "LEAVE SLIP"
      DataField       =   " "
      DataSource      =   " "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   5025
      TabIndex        =   33
      Top             =   180
      Width           =   1665
   End
End
Attribute VB_Name = "frmLeaveSlip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 'Option Explicit
Dim rsEmpList As Recordset
Dim SetDates, FirstMonth, LastMonth, FirstMonthDays, LastMonthDays As Integer
Dim Modify, Add, RecordFound, WDFound As Boolean
Dim WorkingDays As Integer
Dim StartDate As Date

'This is to search for the entered employee code and
'entering its information in the required text boxes.
Private Sub Search()
If rsEmp.RecordCount > 0 Then rsEmp.MoveFirst
    For i = 0 To rsEmp.RecordCount - 1 Step 1
        If (txtCode = rsEmp!Code) Then ' If the employee code is found then
        'Enter the data of employee in respected text boxes
            If Not IsNull(rsEmp!Name) Then txtName.Text = rsEmp!Name
            If Not IsNull(rsEmp!Desig) Then txtDesig.Text = rsEmp!Desig
            If Not IsNull(rsEmp!Salary) Then txtSalary.Text = rsEmp!Salary
            SendKeys "{TAB}"
   If Add = False Then 'If add Button is not pressed
        cmdModify.Enabled = True
        cmdRemove.Enabled = True
   End If
            Exit Sub
        End If
    rsEmp.MoveNext
    If rsEmp.EOF = True Then
        MsgBox "Invalid Employee Code !", vbOKOnly, "OASYS"
        txtCode.Text = ""
        txtCode.SetFocus
        Exit Sub
    End If
    Next
End Sub
'This function is to Check whether the entered Empoyee Code is Valid (Exists or not)
Private Sub ValidateCode()
If rsEmp.RecordCount > 0 Then rsEmp.MoveFirst
    For i = 0 To rsEmp.RecordCount - 1 Step 1
        If (txtCode = rsEmp!Code) Then
        ValidCode = True 'ValidCode = True if Employee Code does exist
            Exit Sub
        End If
    rsEmp.MoveNext
    If rsEmp.EOF = True Then
    ValidCode = False
        MsgBox "Invalid Employee Code !", vbOKOnly, "OASYS"
        Exit Sub
    End If
    Next

End Sub

Private Sub cmbLType_GotFocus()
Call txt_GotFocus
End Sub

Private Sub cmbLType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub cmbLType_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub cmdAdd_Click()
On Error GoTo aerr
Modify = False
Add = True

Call Modi
txtCode.Enabled = True
txtName.Enabled = True
txtDesig.Enabled = True
txtSalary.Enabled = True
cmbLType.Enabled = True
txtFrom.Enabled = True
txtTo.Enabled = True
txtReason.Enabled = True
cmdHelp.Enabled = True

txtCode.Text = ""
txtName.Text = ""
txtDesig.Text = ""
txtSalary.Text = ""
txtFrom.Text = ""
txtTo.Text = ""
txtReason.Text = ""

rsLS.AddNew

cmdCancel.Enabled = True
cmdSave.Enabled = True
cmdModify.Enabled = False
cmdRemove.Enabled = False
cmdPrev.Enabled = False
cmdNext.Enabled = False
cmdAdd.Enabled = False
cmdClose.Enabled = False

txtCode.SetFocus
cmdSave.Enabled = False
Exit Sub
aerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
Unload frmLeaveSlip
frmMain.Show
End Sub

'Cancel Adding or Modifying a record
Private Sub cmdCancel_Click()
On Error GoTo cerr
Modify = False
Add = False
Call Modi
rsLS.CancelUpdate
txtCode.Enabled = False
txtName.Enabled = False
txtDesig.Enabled = False
txtSalary.Enabled = False
cmbLType.Enabled = False
txtFrom.Enabled = False
txtTo.Enabled = False
txtReason.Enabled = False
cmdHelp.Enabled = False

cmdAdd.Enabled = True
cmdModify.Enabled = True
cmdPrev.Enabled = True
cmdNext.Enabled = True
cmdRemove.Enabled = True
cmdCancel.Enabled = False
cmdSave.Enabled = False
cmdClose.Enabled = True

Call cmdPrev_Click
Exit Sub
cerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

'Close Leave Slip Form and goto Main Form
Private Sub cmdClose_Click()
On Error GoTo eerr
frmMain.Show
Unload Me
Exit Sub
eerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

'Show the data grid and hide Record Entry Form
Private Sub cmdHelp_Click()
Framebutton.Visible = False
frame1.Visible = False
cmdClose.Visible = False
dg.Refresh
dg.Visible = True
cmdOK.Visible = True
End Sub

'To Modify the existing Record
Private Sub cmdModify_Click()
On Error GoTo merr
Modify = True
Add = False
Call Modi
txtCode.Enabled = True
txtName.Enabled = True
txtDesig.Enabled = True
txtSalary.Enabled = True
cmbLType.Enabled = True
txtFrom.Enabled = True
txtTo.Enabled = True
txtReason.Enabled = True
cmdHelp.Enabled = True

cmdCancel.Enabled = True
cmdSave.Enabled = True
cmdPrev.Enabled = False
cmdNext.Enabled = False
cmdAdd.Enabled = False
cmdRemove.Enabled = False
cmdModify.Enabled = False
cmdClose.Enabled = False

txtCode.SetFocus
Exit Sub
merr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

'Show the Next Record
Private Sub cmdNext_Click()
On Error GoTo nerr
Modify = False
Add = False
Call Modi

If rsLS.RecordCount = 0 Then
Call ClearAll
    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    cmdCancel.Enabled = False
    cmdSave.Enabled = False
    Exit Sub
Else
    cmdModify.Enabled = True
End If
If rsLS.EOF = False Then rsLS.MoveNext
If rsLS.EOF = True Then rsLS.MoveLast
    showall
Exit Sub
nerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

'To Hide the Data Grid and show Data Entry Form
Private Sub cmdOK_Click()
frame1.Visible = True
Framebutton.Visible = True
cmdClose.Visible = True
cmdOK.Visible = False
dg.Visible = False
End Sub

'Show Previous Record
Private Sub cmdPrev_Click()
On Error GoTo perr
Modify = False
Add = False
Call Modi

If rsLS.RecordCount = 0 Then
Call ClearAll
    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    cmdCancel.Enabled = False
    cmdSave.Enabled = False
    Exit Sub
Else
    cmdModify.Enabled = True
End If
If rsLS.BOF = False Then rsLS.MovePrevious
If rsLS.BOF = True Then rsLS.MoveFirst
showall
    cmdRemove.Enabled = True
Exit Sub
perr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

'Delete an existing Record
Private Sub cmdRemove_Click()
On Error GoTo rerr
Modify = False
Add = False
Call Modi
 If rsLS.RecordCount = 0 Then
 Call ClearAll
    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    cmdCancel.Enabled = False
    cmdSave.Enabled = False
    Exit Sub
End If
       rsLS.Delete
         Call cmdNext_Click
Exit Sub
rerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

'Save the New or Modified Record
Private Sub cmdSave_Click()
On Error GoTo serr

'The entry of Working Days for the month of Leave Date should exist
If WDFound = False Then
    MsgBox "Working Days of the Month in the given date cannot be found !", vbCritical, "OASYS"
    Exit Sub
End If
If Trim(txtCode.Text) = "" Or Trim(txtName.Text) = "" Or Trim(txtFrom.Text) = "" Or Trim(txtTo.Text) = "" Or (cmbLType.Text) = "" Then
    Exit Sub
End If

Add = False

If RecordFound = True Then Exit Sub
If txtCode.Text <> "" Then rsLS!Code = txtCode.Text
If txtName.Text <> "" Then rsLS!Name = txtName.Text
If txtDesig.Text <> "" Then rsLS!Desig = txtDesig.Text
If txtSalary.Text <> "" Then rsLS!Salary = txtSalary.Text
If cmbLType.Text <> "" Then rsLS!LType = cmbLType.Text
If txtFrom.Text <> "" Then rsLS!From = txtFrom.Text
If txtTo.Text <> "" Then rsLS!To = txtTo.Text
If txtReason.Text <> "" Then rsLS!reason = txtReason.Text
If lblTotDays.Caption <> "" Then rsLS!totDays = lblTotDays.Caption

If Year(Date) < Year(rsExpiry!CurrentYear) Then
    MsgBox "You cannot work in previous working year ! Contact Developer", vbCritical, "OASYS"
Exit Sub
ElseIf Year(Date) > Year(rsExpiry!CurrentYear) Then
    MsgBox "You cannot work in next working year ! Contact Developer", vbCritical, "OASYS"
Exit Sub
End If

rsLS.Update

cmdSave.Enabled = False
cmdAdd.Enabled = True
cmdModify.Enabled = True
cmdPrev.Enabled = True
cmdNext.Enabled = True
cmdCancel.Enabled = False
cmdRemove.Enabled = True
cmdClose.Enabled = True
cmdSave.Enabled = False

Call DisableAll  'Disable all the textboxes

If cmdAdd.Enabled = True Then cmdAdd.SetFocus
Exit Sub
serr:
MsgBox Err.Description, vbOKOnly, "OASYS"

Call EnableAll
Call ClearAll
txtCode.SetFocus
End Sub

'This code is to find if the same employee has any previous leave record in the same month
' If there is any leave in the same month, then add the new Leave to the previous Leave
Private Sub FindRecord()
If rsTmpLS.RecordCount > 0 Then rsTmpLS.MoveFirst
For i = 0 To rsTmpLS.RecordCount - 1 Step 1
    If rsTmpLS!Code = txtCode.Text And rsTmpLS!Year = Year(StartDate) And rsTmpLS!Month = Month(StartDate) Then
        RecordFound = True
        rsTmpLS!totDays = Val(lblTotDays.Caption) + Val(rsTmpLS!totDays)
        rsTmpLS.Update
    End If
If rsTmpLS.EOF = False Then rsTmpLS.MoveNext
If rsTmpLS.EOF = True Then
RecordFound = False
Exit Sub
End If
Next
End Sub

'Suppose an Employee takes an Leave
'From : 28/01/03
'To   : 3/04/03
'Now in this case we have to calculate and make an
'database entry of total 4 months. Jan, Feb,March,April

'In this case,
'FirstMonth = 01  (January)
'LastMonth = 04 (April)

'So first we will calculate and enter the no. of days leave taken in the FirstMonth
'Total Leaves in FirstMonth = Total Working Days - 28    (Days left after 28 : 28,29,30,31)
'                           = 31-28 = 4

'Similarly, Leaves in LastMonth = 3  (Days From 1st date... 1,2,3)

'Now as we have calculated the Total Leaves in the First and the Last Month..
'We are just left to calculate the leaves of the months between FirstMonth and LastMonth...

'As we know that, employee remained fully absent in the months between First and Last Month
'Therefore, we get get the Leave taken, by simply finding the Total No. Of Working Days of that month...
'WorkingDays are already entered manually for each month and year, through Working Day Master Form

'Thus, finally we can calculate Total Leave Days between 28/01/03 to 3/04/03 as
'TotalLeave for FirstMonth = 4    (For Jan.)
'TotalLeave for LastMonth = 3     (For April)
'TotalLeave for SecondMonth = 28  (for February)
'TotalLeave for ThirdMonth = 31   (for March)

'Therefore , TotalLeaves = 4 + 3 + 28 + 31 = 66

'This is to Findout No. of Working Days and calculate
'the days attended in the FirstMonth
Private Sub FindWorkingDays()
If rsWD.RecordCount > 0 Then rsWD.MoveFirst
For i = 0 To rsWD.RecordCount - 1 Step 1
    If rsWD!Year = Year(txtFrom.Text) And rsWD!Month = Month(txtFrom.Text) Then
        WorkingDays = rsWD!WD
        WDFound = True
        Exit Sub
    End If
    If rsWD.EOF = False Then rsWD.MoveNext
    If rsWD.EOF = True Then
          MsgBox "Working Days not found for the entered Date!"
          WDFound = False
    End If
Next
End Sub

'This is to findout No. of WorkingDays and calculate
'the days attended in the LastMonth
Private Sub FindWorkingDays1()
If rsWD.RecordCount > 0 Then rsWD.MoveFirst
For i = 0 To rsWD.RecordCount - 1 Step 1
    If rsWD!Year = Year(txtTo.Text) And rsWD!Month = Month(txtTo.Text) Then
        WorkingDays = rsWD!WD
        rsTmpLS!DaysAttn = WorkingDays - totDays
        WDFound = True
        Exit Sub
    End If
    If rsWD.EOF = False Then rsWD.MoveNext
    If rsWD.EOF = True Then
          MsgBox "Working Days not found for the entered Date!"
          WDFound = False
    End If
Next
End Sub
Private Sub FindWorkingStartDays()
'Find the No. of Working Days in the Months between FirstMonth and LastMonth
If rsWD.RecordCount > 0 Then rsWD.MoveFirst
For i = 0 To rsWD.RecordCount - 1 Step 1
    If rsWD!Year = Year(StartDate) And rsWD!Month = Month(StartDate) Then
        WorkingDays = rsWD!WD
        WDFound = True
        Exit Sub
    End If
    If rsWD.EOF = False Then rsWD.MoveNext
    If rsWD.EOF = True Then
          MsgBox "Working Days not found for the entered Date!"
          WDFound = False
    End If
Next
End Sub

'To calculate and save the Total Leave Days for each month
Private Sub calc()
FromDate = txtFrom.Text  ' Stores the Date entered in txtFrom in FromDate variable
ToDate = txtTo.Text      ' Stores the Date entered in txtTo in ToDate Variable

'Suppose we have entered the following dates
' FROM DATE : 18/02/03
' TO DATE   : 02/04/03

FirstMonth = Month(FromDate)  'Here FirstMonth = 02
LastMonth = Month(ToDate)     'Here LastMonth = 04

'If 'FROM' and 'TO' dates have same month and year then
' we only need to calculate the difference of days
If Year(FromDate) = Year(ToDate) Then
    If FirstMonth = LastMonth Then
    totDays = DateDiff("D", FromDate, ToDate) + 1 'Calculate the difference of days
    rsTmpLS.AddNew
        rsTmpLS!Code = txtCode.Text
        rsTmpLS!Year = Year(FromDate)
        rsTmpLS!Month = FirstMonth
        rsTmpLS!totDays = totDays
        Call FindWorkingDays 'Find the Total Working days in FirstMonth

'Days Attended = WorkingDays - no. of leaves(totDays)
        rsTmpLS!DaysAttn = WorkingDays - totDays

'To calculate the Excess Leaves for FirstMonth
If totDays > 2 Then
    rsTmpLS!ExcessLeave = totDays - 2
Else
    rsTmpLS!ExcessLeave = 0
End If
        If WDFound = False Then  ' If WorkingDays for the entered month of the date is not found then cancel update
            rsTmpLS.CancelUpdate
            Exit Sub
        ElseIf WDFound = True Then
            rsTmpLS!TotWorkingDays = WorkingDays
            rsTmpLS.Update
        End If

    End If
End If

' If 'FROM' and 'TO' dates have different Months then
    If Year(FromDate) <> Year(ToDate) Or FirstMonth <> LastMonth Then
        GetDays (FromDate)
        FirstMonthDays = SetDates  'Max. days in a Starting Month
        
'Calculate and Add the TotalDays of Leave in the FirstMonth
        rsTmpLS.AddNew
        rsTmpLS!Code = txtCode.Text
        rsTmpLS!Year = Year(FromDate)
        rsTmpLS!Month = FirstMonth
        rsTmpLS!totDays = FirstMonthDays - Day(FromDate) + 1 'Here our FirstMonth=02. Now, FirstMonthDays=28. ThereFore, TotDays of Leave in FirstMonth is 28-18=10
        totDays = FirstMonthDays - Day(FromDate) + 1
        Call FindWorkingDays
        rsTmpLS!DaysAttn = WorkingDays - totDays
If totDays > 2 Then
    rsTmpLS!ExcessLeave = totDays - 2
Else
    rsTmpLS!ExcessLeave = 0
End If
        If WDFound = False Then  ' If WorkingDays for the entered month of the date is not found then cancel update
            rsTmpLS.CancelUpdate
            Exit Sub
        ElseIf WDFound = True Then
            rsTmpLS!TotWorkingDays = WorkingDays
            rsTmpLS.Update
        End If
        
'Calculate and Add the TotalDays of Leave in the LastMonth
        rsTmpLS.AddNew
        totDays = Day(txtTo.Text)
        rsTmpLS!Code = txtCode.Text
        rsTmpLS!Year = Year(ToDate)
        rsTmpLS!Month = Month(txtTo.Text)  'LastMonth
        'rsTmpLS!TotDays = Day(todate)
        rsTmpLS!totDays = totDays
If totDays > 2 Then
    rsTmpLS!ExcessLeave = totDays - 2
Else
    rsTmpLS!ExcessLeave = 0
End If
        Call FindWorkingDays1

        If WDFound = False Then  ' If WorkingDays for the entered month of the date is not found then cancel update
            rsTmpLS.CancelUpdate
            Exit Sub
        ElseIf WDFound = True Then
            rsTmpLS!TotWorkingDays = WorkingDays
            rsTmpLS!DaysAttn = WorkingDays - totDays
            rsTmpLS.Update
        End If
        
'Calculate and Add the TotalDays of Leave of all the Months
'between FirstMonth and LastMonth

'For this first we need to know the Total No. of Months
'between FirstMonth and LastMonth

TotMnt = DateDiff("M", FromDate, ToDate) + 1

AddDays = FirstMonthDays - Day(FromDate) + 1
StartDate = DateAdd("D", AddDays, FromDate)

SubDate = -Day(ToDate)
Enddate = DateAdd("D", SubDate, ToDate)

For i = 1 To TotMnt Step 1

GetDays (StartDate)
MaxDays = SetDates

If StartDate < Enddate Then
    rsTmpLS.AddNew
    rsTmpLS!Code = txtCode.Text
    rsTmpLS!Year = Year(StartDate)
    rsTmpLS!Month = Month(StartDate)

Call FindWorkingStartDays
        rsTmpLS!totDays = WorkingDays
        totDays = rsTmpLS!totDays
    '    rsTmpLS!DaysAttn = WorkingDays - MaxDays
     rsTmpLS!DaysAttn = 0

If totDays > 2 Then
    rsTmpLS!ExcessLeave = totDays - 2
Else
    rsTmpLS!ExcessLeave = 0
End If

If WDFound = False Then  ' If WorkingDays for the entered month of the date is not found then cancel update
    rsTmpLS.CancelUpdate
    Exit Sub
ElseIf WDFound = True Then
    rsTmpLS!TotWorkingDays = WorkingDays
    rsTmpLS.Update
End If
'rsTmpLS.Update

AddDays = MaxDays + 1
    StartDate = DateAdd("D", AddDays, StartDate)
Else
    Exit Sub
End If

Next
End If
End Sub

'This is to get the maximum no. of days in any month...
'eg. January = 31, February = 28, March = 31, April = 30...
Public Function GetDays(pDate As Date) As String

 Select Case Month(pDate)
  Case 1, 3, 5, 7, 8, 10, 12
   SetDates = "31"
  Case 4, 6, 9, 11
   SetDates = "30"
   '
  Case 2
   If (Year(pDate) Mod 4) = 0 Then
    SetDates = "29"
   Else
    SetDates = "28"
   End If
 End Select
 
 Select Case i
  Case 1, 3, 5, 7, 8, 10, 12
   SetDates = "31"
  Case 4, 6, 9, 11
   SetDates = "30"
   '
  Case 2
   If (Year(pDate) Mod 4) = 0 Then
    SetDates = "29"
   Else
    SetDates = "28"
   End If
 End Select
End Function


Private Sub Form_Activate()
cmdAdd.SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo ferr
Set rsEmpList = New ADODB.Recordset
rsEmpList.Open "select Code,Name,Desig,Salary from MastEmployee", conn, adOpenStatic, adLockOptimistic
Set dg.DataSource = rsEmpList  'Load the employee data in the Data Grid
dg.Refresh

cmbLType.AddItem ("CASUAL")
cmbLType.AddItem ("MEDICAL")

    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    cmdCancel.Enabled = False
    
If rsLS.RecordCount > 0 Then
    cmdAdd.Enabled = True
    cmdPrev.Enabled = True
    cmdNext.Enabled = True
End If
    cmdSave.Enabled = False
    
Exit Sub
ferr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub


Private Sub txtCode_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Add = True Then Call Search
    'SendKeys "{TAB}"
    'KeyAscii = 0
End If
End Sub

Private Sub txtCode_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtDesig_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtDesig_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtDesig_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtFrom_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtFrom_KeyPress(KeyAscii As Integer)
On Error GoTo derr
Select Case KeyAscii
    
    Case vbKeyBack, 48 To 57, vbKeyReturn, 47, vbKeyEscape
    Case Else
         MsgBox "No Special Characters are allowed! Please enter Numbers and / only! ", vbOKOnly, "OASYS"
         KeyAscii = 0
         txtFrom.SetFocus
        
         Exit Sub
    End Select

If KeyAscii = 13 Then
   datevali (txtFrom.Text)
End If
Exit Sub
derr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub txtFrom_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtname_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtname_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub showall()
On Error GoTo showerr
txtCode.Text = ""
txtName.Text = ""
txtDesig.Text = ""
txtSalary.Text = ""
'cmbLType.Text = ""
txtFrom.Text = ""
txtTo.Text = ""
txtReason.Text = ""
lblTotDays.Caption = ""

If rsLS.RecordCount = 0 Then
    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    Exit Sub
ElseIf rsLS.RecordCount > 0 Then
    cmdModify.Enabled = True
    cmdPrev.Enabled = True
    cmdNext.Enabled = True
    cmdRemove.Enabled = True
End If
If Not IsNull(rsLS!Name) Then txtName.Text = rsLS!Name
If Not IsNull(rsLS!Code) Then txtCode.Text = rsLS!Code
If Not IsNull(rsLS!Desig) Then txtDesig.Text = rsLS!Desig
If Not IsNull(rsLS!Salary) Then txtSalary.Text = rsLS!Salary
If Not IsNull(rsLS!LType) Then cmbLType.Text = rsLS!LType
If Not IsNull(rsLS!From) Then txtFrom.Text = rsLS!From
If Not IsNull(rsLS!To) Then txtTo.Text = rsLS!To
If Not IsNull(rsLS!reason) Then txtReason.Text = rsLS!reason
If Not IsNull(rsLS!totDays) Then lblTotDays.Caption = rsLS!totDays
Exit Sub
showerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub


Private Sub txt_GotFocus()
On Error GoTo focerr
    If Trim(txtName.Text) = "" Or Trim(txtCode.Text) = "" Or Trim(txtDesig.Text) = "" Or Trim(txtSalary.Text) = "" Or (cmbLType.Text) = "" Or Trim(txtFrom.Text) = "" Or Trim(txtTo.Text) = "" Or Trim(txtReason.Text) = "" Then
        cmdSave.Enabled = False
        
    Else
        cmdSave.Enabled = True
    End If
Exit Sub
focerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub Modi()
On Error GoTo merr
If Modify = False Then   ' If cmdmodify is clicked then
txtCode.Enabled = False
txtName.Enabled = False
txtDesig.Enabled = False
txtSalary.Enabled = False
cmbLType.Enabled = False
txtFrom.Enabled = False
txtTo.Enabled = False
txtReason.Enabled = False
cmdHelp.Enabled = False

ElseIf Modify = True Then  ' If cmdmodify is not clicked then
txtCode.Enabled = True
txtName.Enabled = True
txtDesig.Enabled = True
txtSalary.Enabled = True
cmbLType.Enabled = True
txtFrom.Enabled = True
txtTo.Enabled = True
txtReason.Enabled = True
cmdHelp.Enabled = True
End If
Exit Sub
merr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub dg_DblClick()
On Error GoTo dgerr

txtCode.Text = dg.Columns(0).Text
txtName.Text = dg.Columns(1).Text
txtDesig.Text = dg.Columns(2).Text
txtSalary.Text = dg.Columns(3).Text
Call cmdOK_Click
cmbLType.SetFocus
Exit Sub
dgerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub dg_KeyPress(KeyAscii As Integer)
On Error GoTo dgkerr
If KeyAscii = 13 Then
txtCode.Text = dg.Columns(0).Text
txtName.Text = dg.Columns(1).Text
txtDesig.Text = dg.Columns(2).Text
txtSalary.Text = dg.Columns(3).Text
Call cmdOK_Click
cmbLType.SetFocus
End If
Call cmdOK_Click
txtFrom.SetFocus

Exit Sub
dgkerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub txtReason_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtReason_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtReason_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtSalary_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtSalary_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtSalary_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtTo_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtTo_KeyPress(KeyAscii As Integer)
On Error GoTo derr
Select Case KeyAscii
    
    Case vbKeyBack, 48 To 57, vbKeyReturn, 47, vbKeyEscape
    Case Else
         MsgBox "No Special Characters are allowed! Please enter Numbers and / only! ", vbOKOnly, "OASYS"
         KeyAscii = 0
         txtTo.SetFocus
         Exit Sub
    End Select

If KeyAscii = 13 Then
   datevali (txtTo.Text)
   If Trim(txtFrom.Text) <> "" And Trim(txtTo.Text) <> "" Then
        totDays = DateDiff("D", txtFrom.Text, txtTo.Text) + 1
        If totDays < 0 Then
            MsgBox "From date cannot be smaller then To date !", vbCritical, "OASYS"
            Exit Sub
        End If
        lblTotDays.Caption = totDays
        Call calc
    End If
End If

Exit Sub
derr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub txtTo_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub EnableAll()
txtCode.Enabled = True
txtName.Enabled = True
txtDesig.Enabled = True
txtSalary.Enabled = True
cmbLType.Enabled = True
txtFrom.Enabled = True
txtTo.Enabled = True
txtReason.Enabled = True
cmdHelp.Enabled = True
End Sub

Private Sub DisableAll()
txtCode.Enabled = False
txtName.Enabled = False
txtDesig.Enabled = False
txtSalary.Enabled = False
cmbLType.Enabled = False
txtFrom.Enabled = False
txtTo.Enabled = False
txtReason.Enabled = False
cmdHelp.Enabled = False
End Sub
Private Sub ClearAll()
txtCode.Text = ""
txtName.Text = ""
txtDesig.Text = ""
txtSalary.Text = ""
cmbLType.Text = ""
txtFrom.Text = ""
txtTo.Text = ""
txtReason.Text = ""
End Sub


'For Date Validation in Textboxes
Private Function datevali(dtt)
On Error GoTo dvalerr
d1 = 0
dd = 0
m1 = 0
mm = 0
Y1 = 0
yy = 0

d1 = InStr(1, dtt, "/")
If d1 > 0 Then
    dd = Mid(dtt, 1, d1 - 1)
Else
    MsgBox "Please enter / after date"
    Me.SetFocus
End If
    dlen = Len(dtt)
    dlen = dlen - d1
    mmid = Mid(dtt, d1 + 1, dlen)
    m1 = InStr(1, mmid, "/")
If m1 > 0 Then
    mm = Mid(mmid, 1, m1 - 1)
Else
    MsgBox "Please enter / after month"
    Me.SetFocus
End If
    dlen = Len(mmid)
    dlen = dlen - m1
    yy = Mid(mmid, m1 + 1, dlen)
If Len(dd) > 2 Then
    MsgBox "Please enter date of two digit"
    Me.SetFocus
ElseIf Val(dd) > 31 Or Val(dd) < 1 Then
    MsgBox "Plz enter date of less then 31"
    Me.SetFocus
ElseIf Val(mm) > 12 Or Val(mm) < 1 Then
    MsgBox "Plz enter month of less/equal then 12"
    Me.SetFocus
ElseIf Len(mm) > 2 Then
    MsgBox "Plz enter month of two digit"
    Me.SetFocus
ElseIf Len(yy) <> 4 Then
    MsgBox "Plz enter year of 4 digit"
    Me.SetFocus
ElseIf (Val(yy) < 1) Then
    MsgBox "Please enteryear Between financial year"
    Me.SetFocus
Else
    SendKeys "{TAB}"
    KeyAscii = 0
End If

Exit Function
dvalerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Function
