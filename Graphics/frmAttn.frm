VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAttn 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Attendance Form"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   60
      TabIndex        =   32
      Top             =   1560
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Height          =   5685
      Left            =   3105
      TabIndex        =   21
      Top             =   450
      Width           =   7035
      Begin VB.ComboBox cmbYear 
         Height          =   315
         ItemData        =   "frmAttn.frx":0000
         Left            =   6060
         List            =   "frmAttn.frx":0022
         TabIndex        =   5
         Top             =   2700
         Width           =   825
      End
      Begin VB.ComboBox txtMnt 
         Height          =   315
         ItemData        =   "frmAttn.frx":0062
         Left            =   3375
         List            =   "frmAttn.frx":0064
         TabIndex        =   6
         Top             =   3285
         Width           =   2760
      End
      Begin VB.TextBox txtAttd 
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         DataField       =   " "
         DataSource      =   " "
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   3360
         TabIndex        =   8
         Text            =   " "
         Top             =   4425
         Width           =   2775
      End
      Begin VB.TextBox txtWD 
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         DataField       =   " "
         DataSource      =   " "
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   3360
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   3870
         Width           =   2775
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         DataField       =   " "
         DataSource      =   " "
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   3375
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1005
         Width           =   2775
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         DataField       =   " "
         DataSource      =   " "
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   3360
         MousePointer    =   10  'Up Arrow
         TabIndex        =   0
         ToolTipText     =   "Employee Code"
         Top             =   450
         Width           =   975
      End
      Begin VB.TextBox txtDOJ 
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         DataField       =   " "
         DataSource      =   " "
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   3
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
         TabIndex        =   22
         Top             =   405
         Width           =   780
      End
      Begin VB.TextBox txtDesig 
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         DataField       =   " "
         DataSource      =   " "
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   3375
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1575
         Width           =   2775
      End
      Begin VB.TextBox txtType 
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         DataField       =   " "
         DataSource      =   " "
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   3375
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   " "
         Top             =   2700
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "YEAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5220
         TabIndex        =   31
         Top             =   2745
         Width           =   780
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "NAME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   240
         Left            =   2580
         TabIndex        =   30
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "DESIGNATION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   195
         Left            =   1830
         TabIndex        =   29
         Top             =   1650
         Width           =   1305
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "ENTER THE MONTH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   195
         Left            =   1305
         TabIndex        =   28
         Top             =   3390
         Width           =   1815
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL WORKING DAYS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   195
         Left            =   990
         TabIndex        =   27
         Top             =   3915
         Width           =   2130
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "NO OF DAYS ATTENDED"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   195
         Left            =   945
         TabIndex        =   26
         Top             =   4485
         Width           =   2205
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "DATE OF JOIN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   195
         Left            =   1800
         TabIndex        =   25
         Top             =   2145
         Width           =   1290
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "TYPE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   195
         Left            =   2610
         TabIndex        =   24
         Top             =   2745
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "EMP CODE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   195
         Left            =   2160
         TabIndex        =   23
         Top             =   540
         Width           =   1005
      End
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
      Left            =   6030
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7290
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Framebutton 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   3375
      TabIndex        =   17
      Top             =   6300
      Width           =   6495
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H0080C0FF&
         Caption         =   "Add"
         Height          =   420
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdRemove 
         BackColor       =   &H0080C0FF&
         Caption         =   "Remove"
         Height          =   420
         Left            =   3690
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdPrev 
         BackColor       =   &H0080C0FF&
         Caption         =   "Previous"
         Height          =   420
         Left            =   4590
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   45
         Width           =   915
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H0080C0FF&
         Caption         =   "Next"
         Height          =   420
         Left            =   5535
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H0080C0FF&
         Caption         =   "Save"
         Height          =   420
         Left            =   1890
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdModify 
         BackColor       =   &H0080C0FF&
         Caption         =   "Modify"
         Height          =   420
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   45
         Width           =   915
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H0080C0FF&
         Caption         =   "Cancel"
         Height          =   420
         Left            =   2790
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   45
         Width           =   870
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H0080C0FF&
      Caption         =   "E&xit"
      Height          =   420
      Left            =   6165
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7110
      Width           =   870
   End
   Begin MSAdodcLib.Adodc AD 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\WIN98\Desktop\Project\Project.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\WIN98\Desktop\Project\Project.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select Code,Name,Desig,Type, DOJ from MastEmployee"
      Caption         =   "Employee List"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid Dg 
      Height          =   7080
      Left            =   1800
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   12488
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   16744576
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
      Caption         =   "ATTENDANCE FORM"
      DataField       =   " "
      DataSource      =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   4995
      TabIndex        =   16
      Top             =   405
      Width           =   2475
   End
End
Attribute VB_Name = "frmAttn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsEmpList As Recordset
Dim Modify, Add, ValidCode As Boolean

' BOOLEANS DECLARED AND THE FUNCTIONS

' Modify = True if Modify button is clicked
' Add = True if Add button is clicked
' ValidCode = True if the employee code entered is valid (does exist)

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


Private Sub Search()
If rsEmp.RecordCount > 0 Then rsEmp.MoveFirst
    For i = 0 To rsEmp.RecordCount - 1 Step 1
        If (txtCode = rsEmp!Code) Then ' If the employee code is found then
        'Enter the data of employee in respected text boxes
            If Not IsNull(rsEmp!Name) Then txtName.Text = rsEmp!Name
            If Not IsNull(rsEmp!Desig) Then txtDesig.Text = rsEmp!Desig
            If Not IsNull(rsEmp!Type) Then txtType.Text = rsEmp!Type
            If Not IsNull(rsEmp!DOJ) Then txtDOJ.Text = rsEmp!DOJ
   If Add = False Then 'If add Button is not pressed
        cmdModify.Enabled = True
        cmdRemove.Enabled = True
   End If
            Exit Sub
        End If
    rsEmp.MoveNext
    If rsEmp.EOF = True Then
        MsgBox "Invalid Employee Code !", vbOKOnly, "Office Automation"
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
        MsgBox "Invalid Employee Code !", vbOKOnly, "Office Automation"
        Exit Sub
    End If
    Next

End Sub

Private Sub cmbYear_KeyPress(KeyAscii As Integer)
txtMnt.Clear
If KeyAscii = 13 Then
If Trim(cmbYear.Text) <> "" Then
    If rsWD.RecordCount > 0 Then rsWD.MoveFirst
    For i = 0 To rsWD.RecordCount - 1 Step 1
        If cmbYear.Text = rsWD!Year Then
            txtMnt.AddItem (rsWD!Month)
        End If
        If rsWD.EOF = False Then rsWD.MoveNext
    Next
End If
    SendKeys "{TAB}"
    KeyAscii = 0
End If

End Sub

Private Sub cmdAdd_Click()
On Error GoTo aerr
Modify = False
Add = True   ' As Add button is clicked, Boolean Add = True

Call Modi
Call EnableAll  ' Enable all the text boxes
Call ClearAll   ' Blank all the text boxes

rsAttd.AddNew

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
MsgBox Err.Description, vbOKOnly, "Office Automation"
End Sub

Private Sub cmdCancel_Click()
On Error GoTo cerr
Modify = False
Add = False
Call Modi
rsAttd.CancelUpdate
Call DisableAll   ' Disable all the text boxes

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
MsgBox Err.Description, vbOKOnly, "Office Automation"
End Sub

Private Sub cmdClose_Click()
On Error GoTo eerr
frmMain.Show
Unload Me
Exit Sub
eerr:
MsgBox Err.Description, vbOKOnly, "Office Automation"
End Sub

Private Sub cmdHelp_Click()  ' This is to show the Data Grid
Framebutton.Visible = False
Frame1.Visible = False
cmdClose.Visible = False
Dg.Refresh
Dg.Visible = True
cmdOK.Visible = True
End Sub

Private Sub cmdModify_Click()
On Error GoTo merr
Modify = True  ' Modify button is clicked
Add = False
Call Modi
Call EnableAll  ' Enable all the text boxes

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
MsgBox Err.Description, vbOKOnly, "Office Automation"
End Sub

Private Sub cmdNext_Click()
On Error GoTo nerr
Modify = False
Add = False
Call Modi

If rsAttd.RecordCount = 0 Then  'If there are no records (data) in the table
    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    cmdCancel.Enabled = False
    Exit Sub
Else
    cmdModify.Enabled = True
End If
rsAttd.MoveNext
If rsAttd.EOF = True Then rsAttd.MoveLast
    showall
Exit Sub
nerr:
MsgBox Err.Description, vbOKOnly, "Office Automation"
End Sub

Private Sub cmdOK_Click()   ' Hide the Data grid
Frame1.Visible = True
Framebutton.Visible = True
cmdClose.Visible = True
cmdOK.Visible = False
Dg.Visible = False
End Sub

Private Sub cmdPrev_Click()
On Error GoTo perr
Modify = False
Add = False
Call Modi

If rsAttd.RecordCount = 0 Then
    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    cmdCancel.Enabled = False
    Exit Sub
End If
If rsAttd.BOF = False And rsAttd.EOF = False Then rsAttd.MovePrevious
If rsAttd.BOF = True Then rsAttd.MoveFirst
showall
    cmdRemove.Enabled = True
Exit Sub
perr:
MsgBox Err.Description, vbOKOnly, "Office Automation"
End Sub

Private Sub cmdRemove_Click()
On Error GoTo rerr
Modify = False
Add = False
Call Modi
 If rsAttd.RecordCount = 0 Then
    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    cmdCancel.Enabled = False
    cmdSave.Enabled = False
    Exit Sub
End If
       rsAttd.Delete
       If rsAttd.EOF = False Then rsAttd.MoveNext
       showall
Exit Sub
rerr:
MsgBox Err.Description, vbOKOnly, "Office Automation"
End Sub

Private Sub cmdSave_Click()
On Error GoTo serr
Add = False

If Trim(txtCode.Text) = "" Or Trim(txtName.Text) = "" Or Trim(txtWD.Text) = "" Or Trim(txtMnt.Text) = "" Or (txtAttd.Text) = "" Then
    Exit Sub
End If

Call ValidateCode
If ValidCode = False Then Exit Sub

If (Val(txtWD.Text) < Val(txtAttd.Text)) Then
    MsgBox "No. of working Days cannot be less then No. of days attended !", vbCritical, "Office Automation"
    txtAttd.Text = ""
    txtAttd.SetFocus
    Exit Sub
End If

If (Val(txtAttd.Text) < 0) Then
    MsgBox "No. of Days Attended cannot be less then 0 !", vbCritical, "Office Automation"
    txtAttd.Text = ""
    txtAttd.SetFocus
    Exit Sub
End If

If txtCode.Text <> "" Then rsAttd!Code = txtCode.Text
If txtName.Text <> "" Then rsAttd!Name = txtName.Text
If txtDOJ.Text <> "" Then rsAttd!DOJ = txtDOJ.Text
If txtMnt.Text <> "" Then rsAttd!Month = txtMnt.Text
If txtWD.Text <> "" Then rsAttd!WD = txtWD.Text
If txtAttd.Text <> "" Then rsAttd!Attd = txtAttd.Text
If txtType.Text <> "" Then rsAttd!Type = txtType.Text
If txtDesig.Text <> "" Then rsAttd!Desig = txtDesig.Text
If cmbYear.Text <> "" Then rsAttd!Year = cmbYear.Text
If Year(Date) < Year(rsExpiry!CurrentYear) Then
    MsgBox "You cannot work in previous working year ! Contact Developer", vbCritical, "Office Automation"
Exit Sub
ElseIf Year(Date) > Year(rsExpiry!CurrentYear) Then
    MsgBox "You cannot work in next working year ! Contact Developer", vbCritical, "Office Automation"
Exit Sub
End If


rsAttd.Update

cmdSave.Enabled = False
cmdAdd.Enabled = True
cmdModify.Enabled = True
cmdPrev.Enabled = True
cmdNext.Enabled = True
cmdCancel.Enabled = False
cmdRemove.Enabled = True
cmdClose.Enabled = True
cmdSave.Enabled = False

Call DisableAll
cmdAdd.SetFocus

If cmdAdd.Enabled = True Then cmdAdd.SetFocus
Exit Sub
serr:
MsgBox Err.Description, vbOKOnly, "Office Automation"

Call EnableAll
Call ClearAll

txtCode.SetFocus
End Sub

Private Sub Command1_Click()
        Call TotalLeaveDays
End Sub

Private Sub Form_Activate()
cmdAdd.SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo ferr
'Set rsEmpList = New ADODB.Recordset
'rsEmpList.Open "select * from MastEmployee", conn, adOpenStatic, adLockOptimistic
cmbYear.Text = Year(Date)

Set Dg.DataSource = AD
Dg.Refresh

    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    cmdCancel.Enabled = False
    
If rsAttd.RecordCount > 0 Then
    cmdAdd.Enabled = True
    cmdPrev.Enabled = True
    cmdNext.Enabled = True
End If
    cmdSave.Enabled = False
    
Exit Sub
ferr:
MsgBox Err.Description, vbOKOnly, "Office Automation"
End Sub

Private Sub txtCode_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Add = True Then Call Search
SendKeys "{TAB}"
KeyAscii = 0
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

Private Sub txtMnt_LostFocus()
If Trim(txtMnt.Text) = "" Or Trim(cmbYear.Text) = "" Then
    MsgBox ("Year and Month Fields cannot be empty !")
    cmbYear.SetFocus
    Exit Sub
End If
If Trim(txtMnt.Text) <> "" And Trim(cmbYear.Text) <> "" Then
    If rsWD.RecordCount > 0 Then rsWD.MoveFirst
    For i = 0 To rsWD.RecordCount - 1 Step 1
        If ((cmbYear = rsWD!Year) And (txtMnt.Text = rsWD!Month)) Then
            txtWD.Text = rsWD!WD
            txtWD.Locked = True
            txtAttd.SetFocus
            Exit Sub
        End If
    If rsWD.EOF = False Then rsWD.MoveNext
    If rsWD.EOF = True Then MsgBox "Invalid Year !"
    Next
'Call TotalLeaveDays
End If

End Sub

Private Sub TotalLeaveDays()
Dim TotalLeave As Integer
mnt = txtMnt.ListIndex + 1  ' By this we will get the No. of the Month. eg. if January=1, February=2, March=3.....

TotalLeave = 0
If rsTmpLS.RecordCount > 0 Then rsTmpLS.MoveFirst
For i = 0 To rsTmpLS.RecordCount - 1 Step 1
    If rsTmpLS!Code = txtCode.Text And rsTmpLS!Year = cmbYear.Text And rsTmpLS!Month = txtMnt.Text Then
        TotalLeave = TotalLeave + rsTmpLS!TotDays
    End If
If rsTmpLS.EOF = False Then rsTmpLS.MoveNext
Next
txtAttd.Text = Val(txtWD.Text) - TotalLeave
'txtTotDays.Text = TotalLeave
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
Call ClearAll   ' Clear (Blank) all the text boxes

If rsAttd.RecordCount = 0 Then
    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    Exit Sub
ElseIf rsAttd.RecordCount > 0 Then
    cmdModify.Enabled = True
    cmdPrev.Enabled = True
    cmdNext.Enabled = True
    cmdRemove.Enabled = True
End If
If Not IsNull(rsAttd!Name) Then txtName.Text = rsAttd!Name
If Not IsNull(rsAttd!Code) Then txtCode.Text = rsAttd!Code
If Not IsNull(rsAttd!Desig) Then txtDesig.Text = rsAttd!Desig
If Not IsNull(rsAttd!Month) Then txtMnt.Text = rsAttd!Month
If Not IsNull(rsAttd!Type) Then txtType.Text = rsAttd!Type
If Not IsNull(rsAttd!WD) Then txtWD.Text = rsAttd!WD
If Not IsNull(rsAttd!Attd) Then txtAttd.Text = rsAttd!Attd
If Not IsNull(rsAttd!DOJ) Then txtDOJ.Text = rsAttd!DOJ
If Not IsNull(rsAttd!Year) Then cmbYear.Text = rsAttd!Year
Exit Sub
showerr:
MsgBox Err.Description, vbOKOnly, "Office Automation"
End Sub

Private Sub txt_GotFocus()
On Error GoTo focerr
    If Trim(txtName.Text) = "" Or Trim(txtCode.Text) = "" Or Trim(txtMnt.Text) = "" Or Trim(txtWD.Text) = "" Or (txtType.Text) = "" Or Trim(txtAttd.Text) = "" Then
        cmdSave.Enabled = False
        
    Else
        cmdSave.Enabled = True
    End If
Exit Sub
focerr:
MsgBox Err.Description, vbOKOnly, "Office Automation"
End Sub

Private Sub Modi()
On Error GoTo merr
If Modify = False Then   ' If cmdmodify is clicked then
    Call DisableAll
ElseIf Modify = True Then  ' If cmdmodify is not clicked then
    Call EnableAll
End If
Exit Sub
merr:
MsgBox Err.Description, vbOKOnly, "Office Automation"
End Sub

Private Sub dg_DblClick()
On Error GoTo dgerr

txtCode.Text = Dg.Columns(0).Text
txtName.Text = Dg.Columns(1).Text
txtDesig.Text = Dg.Columns(2).Text
txtType.Text = Dg.Columns(3).Text
txtDOJ.Text = Dg.Columns(4).Text
Call cmdOK_Click
cmbYear.SetFocus
Exit Sub
dgerr:
MsgBox Err.Description, vbOKOnly, "Office Automation"
End Sub

Private Sub dg_KeyPress(KeyAscii As Integer)
On Error GoTo dgkerr
If KeyAscii = 13 Then
txtCode.Text = Dg.Columns(0).Text
txtName.Text = Dg.Columns(1).Text
txtDesig.Text = Dg.Columns(2).Text
txtType.Text = Dg.Columns(3).Text
txtDOJ.Text = Dg.Columns(4).Text
Call cmdOK_Click
cmbYear.SetFocus
End If
Call cmdOK_Click
cmbYear.SetFocus

Exit Sub
dgkerr:
MsgBox Err.Description, vbOKOnly, "Office Automation"
End Sub

Private Sub txtMnt_GotFocus()
Call txt_GotFocus
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

   ' SendKeys "{TAB}"
   ' KeyAscii = 0
End If
End Sub
Private Sub GetWorkingDays()
If Trim(txtMnt.Text) <> "" And Trim(cmbYear.Text) <> "" Then
    If rsWD.RecordCount > 0 Then rsWD.MoveFirst
    For i = 0 To rsWD.RecordCount - 1 Step 1
        If ((cmbYear = rsWD!Year) And (txtMnt.Text = rsWD!Month)) Then
            txtWD.Text = rsWD!WD
            txtWD.Locked = True
            txtAttd.SetFocus
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
Private Sub txtMnt_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtWD_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtWD_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtWD_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtDOJ_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtDOJ_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtDOJ_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub cmbWD_GotFocus()
Call txt_GotFocus
End Sub

Private Sub cmbWD_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub cmbWD_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtAttd_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtAttd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtAttd_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub


Private Sub txtType_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtType_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub
Private Sub DisableAll()  ' Disable all the text boxes
txtCode.Enabled = False
txtName.Enabled = False
txtDesig.Enabled = False
txtType.Enabled = False
txtWD.Enabled = False
txtMnt.Enabled = False
txtAttd.Enabled = False
txtDOJ.Enabled = False
cmdHelp.Enabled = False
cmbYear.Enabled = False
End Sub
Private Sub EnableAll()  ' Enable all the text boxes
txtCode.Enabled = True
txtName.Enabled = True
txtDesig.Enabled = True
txtWD.Enabled = True
txtType.Enabled = True
txtMnt.Enabled = True
txtAttd.Enabled = True
txtDOJ.Enabled = True
cmdHelp.Enabled = True
cmbYear.Enabled = True
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
