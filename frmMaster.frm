VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMaster 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Master"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdList 
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
      Height          =   390
      Left            =   5580
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2640
      Width           =   885
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
      Height          =   390
      Left            =   5580
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4320
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Frame frame1 
      BackColor       =   &H00000000&
      Height          =   2355
      Left            =   360
      TabIndex        =   10
      Top             =   120
      Width           =   6075
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1905
         MaxLength       =   5
         TabIndex        =   0
         Top             =   720
         Width           =   1860
      End
      Begin VB.TextBox txtDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1905
         MaxLength       =   25
         TabIndex        =   1
         Top             =   1350
         Width           =   1860
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00EBCCB4&
         Height          =   330
         Left            =   420
         TabIndex        =   13
         Top             =   735
         Width           =   1365
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00EBCCB4&
         Height          =   330
         Left            =   420
         TabIndex        =   12
         Top             =   1365
         Width           =   1410
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H0041E9D8&
      Caption         =   "E&xit"
      Height          =   420
      Left            =   2925
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4320
      Width           =   870
   End
   Begin VB.Frame Framebutton 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   135
      TabIndex        =   2
      Top             =   3510
      Width           =   6495
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H0080C0FF&
         Caption         =   "Add"
         Height          =   420
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H0080C0FF&
         Caption         =   "Cancel"
         Height          =   420
         Left            =   2790
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdModify 
         BackColor       =   &H0080C0FF&
         Caption         =   "Modify"
         Height          =   420
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   60
         Width           =   915
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H0080C0FF&
         Caption         =   "Save"
         Height          =   420
         Left            =   1890
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H0080C0FF&
         Caption         =   "Next"
         Height          =   420
         Left            =   5535
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdPrev 
         BackColor       =   &H0080C0FF&
         Caption         =   "Previous"
         Height          =   420
         Left            =   4590
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   60
         Width           =   915
      End
      Begin VB.CommandButton cmdRemove 
         BackColor       =   &H0080C0FF&
         Caption         =   "Remove"
         Height          =   420
         Left            =   3690
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   45
         Width           =   870
      End
   End
   Begin MSDataGridLib.DataGrid dg 
      Height          =   3015
      Left            =   1080
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   5318
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   15499943
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Caption         =   "Code List"
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
End
Attribute VB_Name = "frmMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rsMaster As ADODB.Recordset
Dim Modify, Add, Search, RecordFound As Boolean

' BOOLEANS DECKARED AND THE PURPOSE
' Modify : To check if modify button is clicked
' Add    : To check if Add button is clicked


'When we clicked any of the button in the Main Form (frmMain),
'A particular string value, gets stored in the 'FormName' that we
'have declared as a public String
'Thus from the string of FormName, we can get to know, which
'button in the Main Form was clicked.
'eg. Type, Caste, Desig... etc
'Now depending on the String of FormName, we decide, with which
'table should we connect recordset rsMaster
'MakeConnection function deals with all this...
Private Sub MakeConnection()
On Err GoTo errConn
    If FormName = "Class" Then
        Set rsMaster = New ADODB.Recordset
        rsMaster.Open "select * from MastClass", conn, adOpenStatic, adLockOptimistic
    
    ElseIf FormName = "Type" Then
        Set rsMaster = New ADODB.Recordset
        rsMaster.Open "select * from MastType", conn, adOpenStatic, adLockOptimistic
    
    ElseIf FormName = "Caste" Then
        Set rsMaster = New ADODB.Recordset
        rsMaster.Open "select * from MastCaste", conn, adOpenStatic, adLockOptimistic
    
    ElseIf FormName = "Section" Then
        Set rsMaster = New ADODB.Recordset
        rsMaster.Open "select * from MastSection", conn, adOpenStatic, adLockOptimistic
    
    ElseIf FormName = "Desig" Then
        Set rsMaster = New ADODB.Recordset
        rsMaster.Open "select * from MastDesignation", conn, adOpenStatic, adLockOptimistic
    ElseIf FormName = "Income" Then
        Set rsMaster = New ADODB.Recordset
        rsMaster.Open "select * from MastIncome", conn, adOpenStatic, adLockOptimistic
        Label1.Caption = "Income Code"
        Label2.Caption = "Desc."
    ElseIf FormName = "Expense" Then
        Set rsMaster = New ADODB.Recordset
        rsMaster.Open "select * from MastExpense", conn, adOpenStatic, adLockOptimistic
        Label1.Caption = "Expense Code"
        Label2.Caption = "Desc."
    End If
Exit Sub
errConn:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub cmdAdd_Click()
On Error GoTo aerr
    Modify = False    'Modify = True only when Modify Button is clicked
    Call Modi   ' This function shows what to do if Modify=True or Modify=False
    txtCode.Enabled = True
    txtDesc.Enabled = True
    txtCode.Text = ""
    txtDesc.Text = ""
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
    frmMaster.Hide
    frmMain.Show
End Sub

Private Sub cmdCancel_Click()
On Error GoTo cerr
    Modify = False
    Call Modi
    txtCode.Enabled = False
    txtDesc.Enabled = False
    cmdAdd.Enabled = True
    cmdModify.Enabled = True
    cmdPrev.Enabled = True
    cmdNext.Enabled = True
    cmdRemove.Enabled = True
    cmdCancel.Enabled = False
    cmdSave.Enabled = False
    cmdClose.Enabled = True
    
    If rsMaster.RecordCount = 0 Then
    txtCode.Text = ""
    txtDesc.Text = ""
    cmdAdd.Enabled = True
    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    cmdCancel.Enabled = False
    Exit Sub
    End If
    Call showall
    Exit Sub
cerr:
    MsgBox Err.Description, vbOKOnly, "OASYS"
    frmMaster.Hide
    frmMain.Show
End Sub

Private Sub cmdClose_Click()
On Error GoTo eerr
If FormName = "Income" Or FormName = "Expense" Then
    frmProfitLossMain.Show
Else
    frmMain.Show
End If
    Unload Me
Exit Sub
eerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
    frmMaster.Hide
    frmMain.Show
End Sub

Private Sub cmdList_Click()
rsMaster.Requery
frame1.Visible = False
dg.Visible = True
Framebutton.Visible = False
cmdOK.Visible = True
cmdClose.Visible = False
cmdList.Visible = False
End Sub

Private Sub cmdModify_Click()
On Error GoTo merr
    Modify = True
    Call Modi
    txtCode.Enabled = True
    txtDesc.Enabled = True
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
    frmMaster.Hide
    frmMain.Show
End Sub

Private Sub cmdNext_Click()
On Error GoTo nerr
    Modify = False
    Call Modi
    
    If rsMaster.RecordCount = 0 Then
        txtCode.Text = ""
        txtDesc.Text = ""
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
        If rsMaster.EOF = False Then rsMaster.MoveNext
        If rsMaster.EOF = True Then rsMaster.MoveLast
        showall
        cmdRemove.Enabled = True

Exit Sub
nerr:
    MsgBox Err.Description, vbOKOnly, "OASYS"
    frmMaster.Hide
    frmMain.Show
End Sub

Private Sub cmdOK_Click()
frame1.Visible = True
dg.Visible = False
Framebutton.Visible = True
cmdOK.Visible = False
cmdClose.Visible = True
cmdList.Visible = True
End Sub

Private Sub cmdPrev_Click()
On Error GoTo perr
    Modify = False
    Call Modi

'if there are no records in the table then disable the below buttons
    If rsMaster.RecordCount = 0 Then
        txtCode.Text = ""
        txtDesc.Text = ""
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
    
    If rsMaster.BOF = False Then rsMaster.MovePrevious
    If rsMaster.BOF = True Then rsMaster.MoveFirst
    showall
    cmdRemove.Enabled = True
    Exit Sub
perr:
    MsgBox Err.Description, vbOKOnly, "OASYS"
    frmMaster.Hide
    frmMain.Show
End Sub

Private Sub cmdRemove_Click()
On Error GoTo rerr
    Modify = False
    Call Modi
If rsMaster.RecordCount = 0 Then
    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    cmdCancel.Enabled = False
    cmdSave.Enabled = False
    Exit Sub
End If
      
       rsMaster.Delete
        Call cmdNext_Click
Exit Sub
rerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
    frmMaster.Hide
    frmMain.Show
End Sub
Private Sub FindRecord()
On Error GoTo ErrFind
    If rsMaster.RecordCount > 0 Then rsMaster.MoveFirst
    For i = 0 To rsMaster.RecordCount - 1 Step 1
        If rsMaster!Code = txtCode.Text Then
            MsgBox "Code Already Exists!", vbCritical, "OASYS"
            txtCode.Text = ""
            txtDesc.Text = ""
            txtCode.SetFocus
            RecordFound = True
            Exit Sub
        End If
        If rsMaster.EOF = False Then rsMaster.MoveNext
        If rsMaster.EOF = True Then
        RecordFound = False
        End If
    Next
Exit Sub
ErrFind:
    MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub
Private Sub cmdSave_Click()
On Error GoTo serr

If Trim(txtCode.Text) = "" Or Trim(txtDesc.Text) = "" Then
    Exit Sub
End If
If Modify = False Then
Call FindRecord
    If RecordFound = True Then  'If Code already exists then Exit Sub
        Exit Sub
    ElseIf RecordFound = False Then
        rsMaster.AddNew  'If Code does not exist then Add this record
    End If
End If
If txtCode.Text <> "" Then rsMaster!Code = txtCode.Text
If txtDesc.Text <> "" Then rsMaster!Desc = txtDesc.Text

If Year(Date) < Year(rsExpiry!CurrentYear) Then
    MsgBox "You cannot work in previous working year ! Contact Developer", vbCritical, "OASYS"
Exit Sub
ElseIf Year(Date) > Year(rsExpiry!CurrentYear) Then
    MsgBox "You cannot work in next working year ! Contact Developer", vbCritical, "OASYS"
Exit Sub
End If

rsMaster.Update


cmdAdd.Enabled = True
cmdModify.Enabled = True
cmdPrev.Enabled = True
cmdNext.Enabled = True
cmdCancel.Enabled = False
cmdRemove.Enabled = True
cmdClose.Enabled = True
cmdSave.Enabled = False

txtCode.Enabled = False
txtDesc.Enabled = False

If cmdAdd.Enabled = True Then cmdAdd.SetFocus
Exit Sub
serr:
MsgBox Err.Description, vbOKOnly, "OASYS"
    frmMaster.Hide
    frmMain.Show
End Sub

Private Sub Form_Activate()
cmdAdd.SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo ferr
    Call MakeConnection
    Set dg.DataSource = rsMaster
    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    cmdCancel.Enabled = False
If rsMaster.RecordCount > 0 Then
    cmdAdd.Enabled = True
    cmdPrev.Enabled = True
    cmdNext.Enabled = True
End If
    cmdSave.Enabled = False
Exit Sub
ferr:
MsgBox Err.Description, vbOKOnly, "OASYS"
    frmMaster.Hide
    frmMain.Show
End Sub

Private Sub txtCode_GotFocus()
Call txt_GotFocus
txtCode.BackColor = &HC00000
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtCode_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtCode_LostFocus()
txtCode.BackColor = vbBlack
End Sub

Private Sub txtDesc_GotFocus()
Call txt_GotFocus
txtDesc.BackColor = &HC00000
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtDesc_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub showall()
'On Error GoTo showerr
txtCode.Text = ""
txtDesc.Text = ""

If rsMaster.RecordCount = 0 Then
    txtCode.Text = ""
    txtDesc.Text = ""
    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    Exit Sub
ElseIf rsMaster.RecordCount > 0 Then
    cmdModify.Enabled = True
    cmdPrev.Enabled = True
    cmdNext.Enabled = True
    cmdRemove.Enabled = True
End If
If Not IsNull(rsMaster!Desc) Then txtDesc.Text = rsMaster!Desc
If Not IsNull(rsMaster!Code) Then txtCode.Text = rsMaster!Code
Exit Sub
showerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub


Private Sub txt_GotFocus()
On Error GoTo focerr
    If Trim(txtDesc.Text) = "" Or Trim(txtCode.Text) = "" Then
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
txtDesc.Enabled = False

ElseIf Modify = True Then  ' If cmdmodify is not clicked then
txtCode.Enabled = True
txtDesc.Enabled = True
End If
Exit Sub
merr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub txtDesc_LostFocus()
txtDesc.BackColor = vbBlack
End Sub
