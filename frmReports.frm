VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmReports 
   Caption         =   "Reports ...."
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "REPORTS"
      ForeColor       =   &H00000000&
      Height          =   8655
      Left            =   0
      TabIndex        =   10
      ToolTipText     =   "Select the type of Report that you want to view"
      Top             =   0
      Width           =   3555
      Begin VB.OptionButton OptProfitLoss 
         BackColor       =   &H00E6AC7D&
         Caption         =   "PROFIT AND LOSS ANALYSIS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   300
         TabIndex        =   23
         Top             =   6180
         Width           =   2925
      End
      Begin VB.OptionButton optEmpAttn 
         BackColor       =   &H00E6AC7D&
         Caption         =   "Employee Attendance Records"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   300
         TabIndex        =   22
         Top             =   5580
         Width           =   2925
      End
      Begin VB.CommandButton cmdShow1 
         BackColor       =   &H0041E9D8&
         Caption         =   "&Show"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   " View the Report"
         Top             =   7320
         Width           =   1365
      End
      Begin VB.OptionButton optRetiredEmp 
         BackColor       =   &H00E6AC7D&
         Caption         =   "Retired Employee Records"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   300
         TabIndex        =   20
         Top             =   2520
         Width           =   2925
      End
      Begin VB.OptionButton optMonthlyAttn 
         BackColor       =   &H00E6AC7D&
         Caption         =   "Monthly Attendance Report"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   330
         TabIndex        =   19
         Top             =   4980
         Width           =   2925
      End
      Begin VB.OptionButton optEmpLeave 
         BackColor       =   &H00E6AC7D&
         Caption         =   "Employee Leave Records"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   300
         TabIndex        =   18
         Top             =   4380
         Width           =   2925
      End
      Begin VB.OptionButton optEmpPerDet 
         BackColor       =   &H00E6AC7D&
         Caption         =   "Employee Personal Details"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   315
         TabIndex        =   17
         Top             =   660
         Width           =   2925
      End
      Begin VB.OptionButton optEmpOrg 
         BackColor       =   &H00E6AC7D&
         Caption         =   "Employee Organizational Info"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   330
         TabIndex        =   16
         Top             =   1905
         Width           =   2925
      End
      Begin VB.OptionButton optContact 
         BackColor       =   &H00E6AC7D&
         Caption         =   "Employee Contact Info"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   330
         TabIndex        =   15
         Top             =   1290
         Width           =   2925
      End
      Begin VB.OptionButton optDatewiseRet 
         BackColor       =   &H00E6AC7D&
         Caption         =   "Datewise Retirement Report"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   330
         TabIndex        =   14
         Top             =   3780
         Width           =   2925
      End
      Begin VB.OptionButton optEmpRet 
         BackColor       =   &H00E6AC7D&
         Caption         =   "Employee Retirement Report"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   330
         TabIndex        =   13
         Top             =   3150
         Width           =   2925
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H0041E9D8&
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   1860
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   7320
         Width           =   1365
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   8655
      Left            =   3195
      TabIndex        =   6
      Top             =   0
      Width           =   8715
      Begin VB.OptionButton optDatewisePaySlip 
         BackColor       =   &H00E6AC7D&
         Caption         =   "Datewise Payslips"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   600
         TabIndex        =   25
         Top             =   870
         Visible         =   0   'False
         Width           =   2925
      End
      Begin VB.OptionButton optPaySlip 
         BackColor       =   &H00E6AC7D&
         Caption         =   "Employeewise Payslips"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   645
         TabIndex        =   24
         Top             =   360
         Visible         =   0   'False
         Width           =   2925
      End
      Begin Crystal.CrystalReport CR1 
         Left            =   780
         Top             =   5760
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin Crystal.CrystalReport Cr 
         Left            =   6660
         Top             =   1860
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.TextBox txtRep 
         Height          =   315
         Left            =   6600
         TabIndex        =   11
         Text            =   "MemReg"
         Top             =   1140
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Frame frmview 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Quick View"
         Height          =   1995
         Left            =   1320
         TabIndex        =   7
         Top             =   4140
         Visible         =   0   'False
         Width           =   5115
         Begin VB.CommandButton cmdLastMonth 
            BackColor       =   &H0041E9D8&
            Caption         =   "Last Month"
            Height          =   375
            Left            =   600
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "View Last Month's Report"
            Top             =   1185
            Width           =   1095
         End
         Begin VB.CommandButton cmdLastWeek 
            BackColor       =   &H0041E9D8&
            Caption         =   "Last Week"
            Height          =   375
            Left            =   600
            Style           =   1  'Graphical
            TabIndex        =   0
            ToolTipText     =   "View Last week's Report"
            Top             =   600
            Width           =   1095
         End
         Begin VB.CommandButton cmd3rd 
            BackColor       =   &H0041E9D8&
            Caption         =   "3rd Quarter"
            Height          =   375
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "View reports for July to Sept."
            Top             =   600
            Width           =   1095
         End
         Begin VB.CommandButton cmd2nd 
            BackColor       =   &H0041E9D8&
            Caption         =   "2nd Quarter"
            Height          =   375
            Left            =   1965
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "View reports for April to June"
            Top             =   1185
            Width           =   1050
         End
         Begin VB.CommandButton cmd1st 
            BackColor       =   &H0041E9D8&
            Caption         =   "1st Quarter"
            Height          =   375
            Left            =   1980
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "View the reports for the 1st quarter (Months January to March)"
            Top             =   600
            Width           =   1050
         End
         Begin VB.CommandButton cmd4th 
            BackColor       =   &H0041E9D8&
            Caption         =   "4th Quarter"
            Height          =   375
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "View reports for Oct. to Dec."
            Top             =   1200
            Width           =   1110
         End
      End
      Begin VB.Label lblTDate 
         Caption         =   "Label8"
         Height          =   375
         Left            =   7140
         TabIndex        =   9
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblfDate 
         Caption         =   "Label8"
         Height          =   375
         Left            =   5700
         TabIndex        =   8
         Top             =   480
         Visible         =   0   'False
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim YR As Integer
Private Sub cmdExit_Click()
frmControl.Show
frmReports.Hide
End Sub

Private Sub cmdShow1_Click()
If optEmpPerDet.Value = True Then
    Call EmpPerDet
ElseIf optContact.Value = True Then
    Call EmpContact
ElseIf optEmpOrg.Value = True Then
    Call EmpOrg
ElseIf optRetiredEmp.Value = True Then
    Call RetiredEmp
ElseIf optEmpRet.Value = True Then
    Call EmpRet
ElseIf optDatewiseRet.Value = True Then
    Call DateWiseRet
ElseIf optEmpLeave.Value = True Then
    Call EmpLeave
ElseIf optMonthlyAttn.Value = True Then
    Call MonthlyAttn
ElseIf optEmpAttn.Value = True Then
    Call EmpAttn
End If
Cr.Action = 1
Call AllFalse
End Sub


Private Sub optContact_Click()
frmview.Visible = False
End Sub

Private Sub EmpContact()
On Error GoTo EmpContacterr
Cr.Reset
Cr.ReportTitle = "Employee Contact Info"
Cr.WindowState = crptMaximized
Cr.WindowTitle = "Employee Contact Info"
Cr.WindowShowGroupTree = True
Cr.DataFiles(0) = App.Path & "\Project97.mdb"
Cr.ReportFileName = App.Path & "\Reports\Employee Details\rptEmpContact.rpt"

Cr.Destination = crptToWindow
Cr.WindowShowRefreshBtn = True

Exit Sub
EmpContacterr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub optContact_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub optDatewiseRet_Click()
frmview.Visible = False
End Sub

Private Sub OptEmpAttn_Click()
frmview.Visible = False
End Sub

Private Sub EmpAttn()
On Error GoTo Attnerr
Cr.Reset
Cr.ReportTitle = "Employee Attendance Report"
Cr.WindowState = crptMaximized
Cr.WindowTitle = "Employee Attendance Report"
Cr.WindowShowGroupTree = True
Cr.DataFiles(0) = App.Path & "\Project97.mdb"
Cr.ReportFileName = App.Path & "\Reports\Attendance\Employee Attendance Records.rpt"

Cr.Destination = crptToWindow
Cr.WindowShowRefreshBtn = True

Exit Sub
Attnerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub DateWiseRet()
On Error GoTo errDateWiseRet
Cr.Reset
Cr.ReportTitle = "Datewise Emp. Retirement Report"
Cr.WindowState = crptMaximized
Cr.WindowTitle = "Datewise Emp. Retirement Report"
Cr.WindowShowGroupTree = True
Cr.DataFiles(0) = App.Path & "\Project97.mdb"
Cr.ReportFileName = App.Path & "\Reports\Retirement\rptDatewiseRet.rpt"
Cr.Destination = crptToWindow
Cr.WindowShowRefreshBtn = True
Exit Sub
errDateWiseRet:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub optDateWiseRet_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub optEmpOrg_Click()
frmview.Visible = False
End Sub
Private Sub EmpOrg()
On Error GoTo EmpOrgerr
Cr.Reset
Cr.ReportTitle = "Employee Organizational Info"
Cr.WindowState = crptMaximized
Cr.WindowTitle = "Employee Organizational Info"
Cr.WindowShowGroupTree = True
Cr.DataFiles(0) = App.Path & "\Project97.mdb"
Cr.ReportFileName = App.Path & "\Reports\Employee Details\rptEmpOrg.rpt"
Cr.Destination = crptToWindow
Cr.WindowShowRefreshBtn = True
Exit Sub
EmpOrgerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub
Private Sub optEmpOrg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub optMonthlyAttn_Click()
frmview.Visible = False
End Sub

Private Sub MonthlyAttn()
On Error GoTo Attnerr
Cr.Reset
Cr.ReportTitle = "Monthly Attendance Report"
Cr.WindowState = crptMaximized
Cr.WindowTitle = "Monthly Attendance Report"
Cr.WindowShowGroupTree = True
Cr.DataFiles(0) = App.Path & "\Project97.mdb"
Cr.ReportFileName = App.Path & "\Reports\Attendance\Monthly Attendance Records.rpt"

Cr.Destination = crptToWindow
Cr.WindowShowRefreshBtn = True
Exit Sub
Attnerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub optMonthlyAttn_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub optEmpOrgactInfo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub optEmpLeave_Click()
frmview.Visible = True
End Sub
Private Sub EmpLeave()
On Error GoTo errEmpLeave
Cr.Reset
Cr.ReportTitle = "Employee Leave Record"
Cr.WindowState = crptMaximized
Cr.WindowTitle = "Employee Leave Record"
Cr.WindowShowGroupTree = True
Cr.DataFiles(0) = App.Path & "\Project97.mdb"
Cr.ReportFileName = App.Path & "\Reports\Leave\rptEmpLeave.rpt"

Cr.Destination = crptToWindow
Cr.WindowShowRefreshBtn = True

Exit Sub
errEmpLeave:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub
Private Sub optEmpLeave_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub optEmpRet_Click()
frmview.Visible = False
End Sub
Private Sub EmpRet()
'on Error GoTo errEmpRet
Cr.Reset
Cr.ReportTitle = "Employee Retirement Report"
Cr.WindowState = crptMaximized
Cr.WindowTitle = "Employee Retirement Report"
Cr.WindowShowGroupTree = True
Cr.DataFiles(0) = App.Path & "\Project97.mdb"
Cr.ReportFileName = App.Path & "\Reports\Retirement\rptEmpRet.rpt"
Cr.Destination = crptToWindow
Cr.WindowShowRefreshBtn = True

CR1.Reset
CR1.ReportTitle = "Employee Retirement Report"
CR1.WindowState = crptMaximized
CR1.WindowTitle = "Employee Retirement Report"
CR1.WindowShowGroupTree = True
CR1.DataFiles(0) = App.Path & "\Project97.mdb"
CR1.ReportFileName = App.Path & "\Reports\Retirement\rptRetiredEmployees.rpt"
CR1.Destination = crptToWindow
CR1.WindowShowRefreshBtn = True

Exit Sub
errEmpRet:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub
Private Sub optEmpRet_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub optEmpAttn_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub optEmpPerDet_Click()
frmview.Visible = False
End Sub
Private Sub EmpPerDet()
On Error GoTo EmpPerDeterr
Cr.Reset
Cr.ReportTitle = "Employee Personal Details"
Cr.WindowState = crptMaximized
Cr.WindowTitle = "Employee Personal Details"
Cr.WindowShowGroupTree = True

Cr.DataFiles(0) = App.Path & "\Project97.mdb"
Cr.ReportFileName = App.Path & "\Reports\Employee Details\rptEmpPersonalDetails.rpt"
Cr.Destination = crptToWindow
Cr.WindowShowRefreshBtn = True
Cr.WindowShowPrintBtn = True
Exit Sub
EmpPerDeterr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub
Private Sub optEmpPerDet_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub OptProfitLoss_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub
Private Sub optProfitLoss_Click()
frmview.Visible = False
YR = InputBox("Enter the Year : ", "PROFIT AND LOSS ANALYSIS", Year(Date))
    Call ProfitLoss
    Cr.SelectionFormula = "{MonthlyProfitLoss.Year}" & "=" & YR
    Cr.Action = 1
    Call AllFalse
End Sub
Private Sub ProfitLoss()
On Error GoTo EmpOrgerr
Cr.Reset
Cr.ReportTitle = "PROFIT AND LOSS ANALYSIS"
Cr.WindowState = crptMaximized
Cr.WindowTitle = "PROFIT AND LOSS ANALYSIS"
Cr.WindowShowGroupTree = True
Cr.DataFiles(0) = App.Path & "\Project97.mdb"
Cr.ReportFileName = App.Path & "\Reports\ProfitLoss\MonthlyProfitLoss.rpt"
Cr.Destination = crptToWindow
Cr.WindowShowRefreshBtn = True
Exit Sub
EmpOrgerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub optRetiredEmp_Click()
frmview.Visible = False
End Sub
Private Sub RetiredEmp()
On Error GoTo EmpPerDeterr
Cr.Reset
Cr.ReportTitle = "Retired Employee Records"
Cr.WindowState = crptMaximized
Cr.WindowTitle = "Retired Employee Records"
Cr.WindowShowGroupTree = True

Cr.DataFiles(0) = App.Path & "\Project97.mdb"
Cr.ReportFileName = App.Path & "\Reports\Retirement\rptRetiredEmployees.rpt"
Cr.Destination = crptToWindow
Cr.WindowShowRefreshBtn = True
Exit Sub
EmpPerDeterr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub optSpec_Click()
framedate.Visible = True
txtFrom.Visible = True
txtTo.Visible = False
End Sub

Private Sub optToday_Click()
framedate.Visible = False
End Sub
Private Sub cmdLastWeek_Click()
If optEmpLeave.Value = True Then
Call EmpLeave
Cr.SelectionFormula = "{LeaveMast.From}=LastFullWeek"
Cr.ReportTitle = "LAST WEEK : Employee Leave Records"
Cr.WindowTitle = "LAST WEEK :Employee Leave Record"
Cr.Action = 1
End If
End Sub
Private Sub cmdLastMonth_Click()
If optEmpLeave.Value = True Then
Call EmpLeave
Cr.SelectionFormula = "{LeaveMast.From}=LastFullmonth"
Cr.ReportTitle = "LAST MONTH : Employee Leave Records"
Cr.WindowTitle = "LAST MONTH :Employee Leave Record"
Cr.Action = 1
End If
End Sub
Private Sub cmd1st_Click()
If optEmpLeave.Value = True Then
Call EmpLeave
Cr.SelectionFormula = "{LeaveMast.From}=Calendar1stQtr"
Cr.WindowTitle = "1ST QUARTER : Employee Leave Record"
Cr.ReportTitle = "1ST QUARTER : Employee Leave Records"
Cr.Action = 1
End If
End Sub
Private Sub cmd2nd_Click()
If optEmpLeave.Value = True Then
Call EmpLeave
Cr.SelectionFormula = "{LeaveMast.From}=Calendar2ndQtr"
Cr.ReportTitle = "2ND QUARTER : Employee Leave Records"
Cr.WindowTitle = "2ND QUARTER : Employee Leave Record"
Cr.Action = 1
End If
End Sub
Private Sub cmd3rd_Click()
If optEmpLeave.Value = True Then
Call EmpLeave
Cr.SelectionFormula = "{LeaveMast.From}=Calendar3rdQtr"
Cr.ReportTitle = "3RD QUARTER : Employee Leave Records"
Cr.WindowTitle = "3RD QUARTER : Employee Leave Record"
Cr.Action = 1
End If
End Sub
Private Sub cmd4th_Click()
If optEmpLeave.Value = True Then
Call EmpLeave
Cr.SelectionFormula = "{LeaveMast.From}=Calendar4thQtr"
Cr.ReportTitle = "4TH QUARTER : Employee Leave Records"
Cr.WindowTitle = "4TH QUARTER :Employee Leave Record"
Cr.Action = 1
End If
End Sub

Private Sub AllFalse()
optEmpPerDet.Value = False
optEmpOrg.Value = False
optContact.Value = False
optDatewiseRet.Value = False
optEmpAttn.Value = False
optEmpLeave.Value = False
optEmpRet.Value = False
OptProfitLoss.Value = False
optRetiredEmp.Value = False
End Sub

