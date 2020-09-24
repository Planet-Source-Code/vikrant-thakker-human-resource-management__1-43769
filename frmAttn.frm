VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmAttn 
   BackColor       =   &H00000000&
   Caption         =   "Attendance Slip"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CR 
      Left            =   420
      Top             =   7440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H0041E9D8&
      Caption         =   "&PRINT"
      Height          =   540
      Left            =   4380
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7500
      Width           =   1140
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H0041E9D8&
      Caption         =   "&CLEAR"
      Height          =   540
      Left            =   5700
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7500
      Width           =   1140
   End
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
      Left            =   9660
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7500
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H0041E9D8&
      Caption         =   "E&XIT"
      Height          =   540
      Left            =   7020
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7500
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   6225
      Left            =   2625
      TabIndex        =   16
      Top             =   750
      Width           =   7095
      Begin VB.ComboBox txtDummyMnt 
         Height          =   315
         Left            =   6180
         TabIndex        =   31
         Text            =   "Dummy"
         Top             =   3300
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cmbYear 
         Height          =   315
         Left            =   5880
         TabIndex        =   1
         Top             =   2700
         Width           =   1035
      End
      Begin VB.TextBox txtLeave 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   " "
         DataSource      =   " "
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   3375
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   " "
         Top             =   4980
         Width           =   2775
      End
      Begin VB.ComboBox cmbDummyYear 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmAttn.frx":0000
         Left            =   5880
         List            =   "frmAttn.frx":0002
         TabIndex        =   10
         Text            =   "Dummy"
         Top             =   2280
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.ComboBox txtMnt 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3375
         TabIndex        =   2
         Top             =   3285
         Width           =   2760
      End
      Begin VB.TextBox txtAttd 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   " "
         DataSource      =   " "
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   " "
         Top             =   4425
         Width           =   2775
      End
      Begin VB.TextBox txtWD 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   " "
         DataSource      =   " "
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   3360
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   3870
         Width           =   2775
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   " "
         DataSource      =   " "
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   3375
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1005
         Width           =   2775
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   " "
         DataSource      =   " "
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   3360
         MaxLength       =   5
         MousePointer    =   10  'Up Arrow
         TabIndex        =   0
         ToolTipText     =   "Employee Code"
         Top             =   450
         Width           =   975
      End
      Begin VB.TextBox txtDOJ 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   " "
         DataSource      =   " "
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2115
         Width           =   1215
      End
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
         Height          =   405
         Left            =   4410
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   405
         Width           =   780
      End
      Begin VB.TextBox txtDesig 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   " "
         DataSource      =   " "
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   3375
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1575
         Width           =   2775
      End
      Begin VB.TextBox txtType 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   " "
         DataSource      =   " "
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   3375
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   " "
         Top             =   2700
         Width           =   1215
      End
      Begin VB.Label lblExcessLeave 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   3420
         TabIndex        =   30
         Top             =   5565
         Width           =   1320
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "EXCESS LEAVE DAYS"
         ForeColor       =   &H00EBCCB4&
         Height          =   330
         Left            =   1320
         TabIndex        =   29
         Top             =   5580
         Width           =   2160
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL DAYS OF LEAVE"
         ForeColor       =   &H00EBCCB4&
         Height          =   195
         Left            =   1140
         TabIndex        =   28
         Top             =   5040
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "YEAR"
         ForeColor       =   &H00EBCCB4&
         Height          =   285
         Left            =   5220
         TabIndex        =   26
         Top             =   2745
         Width           =   780
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "NAME"
         ForeColor       =   &H00EBCCB4&
         Height          =   195
         Left            =   2610
         TabIndex        =   25
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "DESIGNATION"
         ForeColor       =   &H00EBCCB4&
         Height          =   195
         Left            =   1980
         TabIndex        =   24
         Top             =   1650
         Width           =   1125
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "ENTER THE MONTH"
         ForeColor       =   &H00EBCCB4&
         Height          =   195
         Left            =   1425
         TabIndex        =   23
         Top             =   3390
         Width           =   1575
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL WORKING DAYS"
         ForeColor       =   &H00EBCCB4&
         Height          =   195
         Left            =   1170
         TabIndex        =   22
         Top             =   3915
         Width           =   1845
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "NO OF DAYS ATTENDED"
         ForeColor       =   &H00EBCCB4&
         Height          =   195
         Left            =   1125
         TabIndex        =   21
         Top             =   4485
         Width           =   1905
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "DATE OF JOIN"
         ForeColor       =   &H00EBCCB4&
         Height          =   195
         Left            =   1980
         TabIndex        =   20
         Top             =   2145
         Width           =   1095
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "TYPE"
         ForeColor       =   &H00EBCCB4&
         Height          =   195
         Left            =   2610
         TabIndex        =   19
         Top             =   2745
         Width           =   420
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "EMP CODE"
         ForeColor       =   &H00EBCCB4&
         Height          =   195
         Left            =   2235
         TabIndex        =   18
         Top             =   540
         Width           =   855
      End
   End
   Begin MSDataGridLib.DataGrid Dg 
      Height          =   6480
      Left            =   1260
      TabIndex        =   15
      Top             =   720
      Visible         =   0   'False
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   11430
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "ATTENDANCE SLIP"
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
      Left            =   4830
      TabIndex        =   13
      Top             =   225
      Width           =   2685
   End
End
Attribute VB_Name = "frmAttn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsEmpList, rsValidList As Recordset
Dim CodeFound As Boolean

' BOOLEANS DECLARED AND THE FUNCTIONS

' Modify = True if Modify button is clicked
' Add = True if Add button is clicked
' ValidCode = True if the employee code entered is valid (does exist)
' CodeFound = True if the Employee Code entered is found
' USER DEFINED FUNCTIONS AND THEIR PURPOSE

' EnableAll  : enable all the text boxes
' DisableAll : disable all the text boxes
' ClearAll   : Clear (Blank) all the text boxes
' Search     : When we enter the Employee Code and press enter key,
'              it searces for the Emp.Name, Desig, Type and DOJ
'              of the employee with that Code, and enters the
'              data in these fields.
' Modi       : If modify = True then Enables all the text boxes
'              If modify = False then Disable all the text boxes


'This is to search for the entered employee code and
'entering its information in the text boxes.
'eg. Once we enter the Emp.Code and press enter key,
'some of the textboxes should be filled automatically.
Private Sub Search()
If rsEmp.RecordCount > 0 Then rsEmp.MoveFirst
    For i = 0 To rsEmp.RecordCount - 1 Step 1
        If (txtCode = rsEmp!Code) Then ' If the employee code is found then
'Enter the data of employee in respected text boxes
            If Not IsNull(rsEmp!Name) Then txtName.Text = rsEmp!Name
            If Not IsNull(rsEmp!Desig) Then txtDesig.Text = rsEmp!Desig
            If Not IsNull(rsEmp!Type) Then txtType.Text = rsEmp!Type
            If Not IsNull(rsEmp!DOJ) Then txtDOJ.Text = rsEmp!DOJ
            CodeFound = True   ' Employee code is found
            Exit Sub
        End If
    rsEmp.MoveNext
    If rsEmp.EOF = True Then
        CodeFound = False ' Employee Code is not Found
        Exit Sub
    End If
    Next
End Sub

Private Sub cmbYear_GotFocus()
Call AddYears
End Sub

'As we press Enter Key after selecting an year,
'The months for that particular year, should be
'added in the cmbMonth
Private Sub cmbYear_KeyPress(KeyAscii As Integer)
txtMnt.Clear
If KeyAscii = 13 Then
Call AddMonths
Call ClearMonths
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

'For clearing all the TextBoxes and making it ready for NEW Entry
Private Sub cmdClear_Click()
Call ClearAll
txtCode.SetFocus
End Sub

'Closes the Attendance Form, and loads Main Form
Private Sub cmdClose_Click()
    On Error GoTo eerr
    frmMain.Show
    Unload Me
    Exit Sub
eerr:
    MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

' This is to show the Data Grid and hide Data Entry Screen
Private Sub cmdHelp_Click()
    frame1.Visible = False
    cmdClose.Visible = False
    cmdClear.Visible = False
    cmdPrint.Visible = False
    dg.Visible = True
    cmdOK.Visible = True
End Sub

' Hide the Data grid and show the dataentry screen
Private Sub cmdOK_Click()
    frame1.Visible = True
    cmdClose.Visible = True
    cmdClear.Visible = True
    cmdPrint.Visible = True
    cmdOK.Visible = False
    dg.Visible = False
    cmbYear.SetFocus
End Sub

Private Sub cmdPrint_Click()
On Error GoTo perr
If rsPrintAttd.RecordCount > 0 Then rsPrintAttd.MoveFirst
    For i = 0 To rsPrintAttd.RecordCount - 1 Step 1
    If rsPrintAttd.BOF = False Or rsPrintAttd.EOF = False Then
        rsPrintAttd.Delete
        If rsPrintAttd.EOF = False Then rsPrintAttd.MoveNext
    End If
    Next
Call AddPrintRecord
Call ShowReport
Exit Sub
perr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub
Private Sub AddPrintRecord()
rsPrintAttd.AddNew
If Trim(txtCode.Text) <> "" Then rsPrintAttd!Code = txtCode.Text
If Trim(txtName.Text) <> "" Then rsPrintAttd!Name = txtName.Text
If Trim(txtDesig.Text) <> "" Then rsPrintAttd!Desig = txtDesig.Text
If Trim(txtDOJ.Text) <> "" Then rsPrintAttd!DOJ = txtDOJ.Text
If Trim(txtType.Text) <> "" Then rsPrintAttd!Type = txtType.Text
If Trim(txtMnt.Text) <> "" Then rsPrintAttd!Month = txtMnt.Text
If Trim(cmbYear.Text) <> "" Then rsPrintAttd!Year = cmbYear.Text
If Trim(txtWD.Text) <> "" Then rsPrintAttd!WD = txtWD.Text
If Trim(txtAttd.Text) <> "" Then rsPrintAttd!Attd = txtAttd.Text
If Trim(txtLeave.Text) <> "" Then rsPrintAttd!Leave = txtLeave.Text
If lblExcessLeave.Caption <> "" Then rsPrintAttd!ExcessLeave = lblExcessLeave.Caption
rsPrintAttd.Update
End Sub

Private Sub ShowReport()
On Error GoTo Attnerr
Cr.Reset
Cr.ReportTitle = "Monthly Attendance Slip"
Cr.WindowState = crptMaximized
Cr.WindowTitle = "Monthly Attendance Slip"
Cr.WindowShowGroupTree = True
Cr.DataFiles(0) = App.Path & "\Project97.mdb"
Cr.ReportFileName = App.Path & "\Reports\Attendance\Monthly Attendance Slip.rpt"

Cr.Destination = crptToWindow
Cr.WindowShowRefreshBtn = True
Cr.Action = 1
Exit Sub
Attnerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub Form_Activate()
    txtCode.SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo ferr
    Set rsEmpList = New ADODB.Recordset
    rsEmpList.Open "select code,name,desig,doj,type from MastEmployee", conn, adOpenStatic, adLockOptimistic
    cmbYear.Text = Year(Date)  'The year box by default should show the current year

'Load the data in the Data Grid
    Set dg.DataSource = rsEmpList
    dg.Refresh
             
     Exit Sub
ferr:
    MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then

Call Search  'Enter the data automatically in other fields, relative to the entered Employee Code

    If CodeFound = True Then
    Call AddYears
    Call ClearYears
        SendKeys "{TAB}"
        KeyAscii = 0
    ElseIf CodeFound = False Then
        Call ClearAll   'Clear all the fields
        txtCode.SetFocus
        MsgBox "Invalid Code !", vbCritical, "OASYS"
    Exit Sub
End If
End If
End Sub
Private Sub AddYears()
Call ValidList
cmbDummyYear.Clear   ' First remove all the data from the Years combo box
If rsValidList.RecordCount > 0 Then rsValidList.MoveFirst
    For i = 0 To rsValidList.RecordCount - 1 Step 1
        cmbDummyYear.AddItem (rsValidList!Year)
    If rsValidList.EOF = False Then rsValidList.MoveNext
    Next
End Sub
Private Sub ClearYears()
cmbYear.Clear
If cmbDummyYear.ListCount > 0 Then
'MsgBox cmbdummyYear.ListCount
For i = 0 To cmbDummyYear.ListCount Step 1
restart:
If i = cmbDummyYear.ListCount Then Exit Sub

YR = cmbDummyYear.List(i)

    For no = 0 To cmbYear.ListCount Step 1
        If cmbYear.List(no) = YR Then
            i = i + 1  'if year already exists in the list
            no = 0
            GoTo restart
        End If
        If no = cmbYear.ListCount Then
            cmbYear.AddItem (YR) ' if year does not exist then add it
        End If
    Next no
Next i
End If
End Sub
Private Sub AddMonths()
Call ValidList
txtDummyMnt.Clear  'First clear the Months Combo box
If Trim(cmbYear.Text) <> "" Then
    If rsValidList.RecordCount > 0 Then rsValidList.MoveFirst
    For i = 0 To rsValidList.RecordCount - 1 Step 1
        If cmbYear.Text = rsValidList!Year Then
            txtDummyMnt.AddItem (rsValidList!Month)  ' Add Months in the combo box
        End If
        If rsValidList.EOF = False Then rsValidList.MoveNext
    Next
End If
End Sub

Private Sub ClearMonths()
txtMnt.Clear
If txtDummyMnt.ListCount > 0 Then
'MsgBox cmbdummyYear.ListCount
For i = 0 To txtDummyMnt.ListCount Step 1
restart:
If i = txtDummyMnt.ListCount Then Exit Sub

MNT = txtDummyMnt.List(i)

    For no = 0 To txtMnt.ListCount Step 1
        If txtMnt.List(no) = MNT Then
            i = i + 1  'if year already exists in the list
            no = 0
            GoTo restart
        End If
        If no = txtMnt.ListCount Then
            txtMnt.AddItem (MNT) ' if year does not exist then add it
        End If
    Next no
Next i
End If
End Sub
Private Sub ValidList()
Set rsValidList = New ADODB.Recordset
rsValidList.Open "select * from TempLeave where Code=" & "'" & txtCode.Text & "'", conn, adOpenStatic, adLockOptimistic
End Sub

'Add the months in cmbMonth, based on the Year selected in the cmbYear
Private Sub txtMnt_GotFocus()
Call AddMonths
Call ClearMonths
End Sub

'As we select a month, corresponding Working Days of that
'Month and selected year should be entered automatically.
Private Sub txtMnt_LostFocus()
If Trim(txtMnt.Text) <> "" And Trim(cmbYear.Text) <> "" Then
    If rsWD.RecordCount > 0 Then rsWD.MoveFirst
    For i = 0 To rsWD.RecordCount - 1 Step 1
        If ((cmbYear = rsWD!Year) And (txtMnt.Text = rsWD!Month)) Then
            txtWD.Text = rsWD!WD 'Enter the corresponding Working Days
            Exit Sub
        End If
    If rsWD.EOF = False Then rsWD.MoveNext
    If rsWD.EOF = True Then MsgBox "WorkingDays Entry missing for selected Month !", vbCritical, "OASYS"
    Next
End If
End Sub

'This is to Calculate and display the Total Leave Days,
'Days Attended and Excess Leave taken in a selected Month and Year
'Suppose an Employee "xyz" has taken a Leave for two days
'in a month of November
'1st time From     : 05/11/03 to 07/11/03  Here TotDays = 3
'and 2nd time From : 22/11/03 to 25/11/03  Here TotDays = 4

'So, we get Total Leave Days of November = 3+4 = 7
Private Sub TotalLeaveDays()
Dim TotalLeave, EntriesFound As Integer
MNT = txtMnt.ListIndex + 1  ' By this we will get the No. of the Month. eg. if January=1, February=2, March=3.....

TotalLeave = 0
EntriesFound = 0
Call ValidList

If rsValidList.RecordCount > 0 Then rsValidList.MoveFirst
For i = 0 To rsValidList.RecordCount - 1 Step 1
    If rsValidList!Code = txtCode.Text And rsValidList!Year = cmbYear.Text And rsValidList!Month = txtMnt.Text Then
        TotalLeave = TotalLeave + rsValidList!totDays
        EntriesFound = EntriesFound + 1
    End If
If rsValidList.EOF = False Then rsValidList.MoveNext
Next
If EntriesFound = 0 Then
    MsgBox "Attendance Entry Not Found for the corresponding Month and Year!", vbOKOnly, "OASYS"
    Exit Sub
Else
    txtAttd.Text = Val(txtWD.Text) - TotalLeave 'Days Attended = Working Days - Total Leave Days
    txtLeave.Text = TotalLeave
    Call ExcessLeave 'Excess Leave = Total Leaves - 2
    cmdPrint.SetFocus
End If
End Sub

'If Leaves Taken are more then 2 then Excess Leave = Leaves Taken -2
Private Sub ExcessLeave()
If Val(txtLeave.Text) > 2 Then
    lblExcessLeave.Caption = Val(txtLeave.Text) - 2
Else
    lblExcessLeave.Caption = 0
End If
End Sub

'Enter the values in textboxes based on the record selected in the data grid
Private Sub dg_DblClick()
On Error GoTo dgerr
txtCode.Text = dg.Columns(0).Text
txtName.Text = dg.Columns(1).Text
txtDesig.Text = dg.Columns(2).Text
txtDOJ.Text = dg.Columns(3).Text
txtType.Text = dg.Columns(4).Text
Call cmdOK_Click
cmbYear.SetFocus
Exit Sub
dgerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

'Enter the values in textboxes based on the record selected in the data grid
Private Sub dg_KeyPress(KeyAscii As Integer)
On Error GoTo dgkerr
If KeyAscii = 13 Then
    txtCode.Text = dg.Columns(0).Text
    txtName.Text = dg.Columns(1).Text
    txtDesig.Text = dg.Columns(2).Text
    txtDOJ.Text = dg.Columns(3).Text
    txtType.Text = dg.Columns(4).Text
    Call cmdOK_Click
    cmbYear.SetFocus
    cmbYear.Text = Year(Date)
End If
Exit Sub
dgkerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub txtMnt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

If Trim(txtMnt.Text) = "" Or Trim(cmbYear.Text) = "" Then
    MsgBox ("Year and Month Fields cannot be empty !")
    cmbYear.SetFocus
    Exit Sub
End If
Call GetWorkingDays
Call TotalLeaveDays

End If
End Sub

'This is to find and display the Working Days in the Month of a selected Year
Private Sub GetWorkingDays()
If Trim(txtMnt.Text) <> "" And Trim(cmbYear.Text) <> "" Then
    If rsWD.RecordCount > 0 Then rsWD.MoveFirst
    For i = 0 To rsWD.RecordCount - 1 Step 1
        If ((cmbYear = rsWD!Year) And (txtMnt.Text = rsWD!Month)) Then
            txtWD.Text = rsWD!WD  'Display the Working days in the textbox
            Exit Sub
        End If
    If rsWD.EOF = False Then rsWD.MoveNext
    If rsWD.EOF = True Then
        MsgBox "Invalid Year !"
        cmbYear.SetFocus
        Exit Sub
    End If
    Next
End If
End Sub
Private Sub txtWD_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtDOJ_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub cmbWD_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtAttd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub ClearAll()  ' Clear all the text boxes
txtCode.Text = ""
txtName.Text = ""
txtDesig.Text = ""
txtType.Text = ""
txtDOJ.Text = ""
txtMnt.Text = ""
txtWD.Text = ""
txtAttd.Text = ""
cmbYear.Text = ""
End Sub
