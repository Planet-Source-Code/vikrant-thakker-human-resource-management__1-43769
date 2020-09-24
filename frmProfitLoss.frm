VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmProfitLoss 
   BackColor       =   &H00000000&
   Caption         =   "Profit-Loss Entry"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Framebutton 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   0
      TabIndex        =   9
      Top             =   3570
      Width           =   6495
      Begin VB.CommandButton cmdRemove 
         BackColor       =   &H0080C0FF&
         Caption         =   "Remove"
         Height          =   420
         Left            =   3690
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdPrev 
         BackColor       =   &H0080C0FF&
         Caption         =   "Previous"
         Height          =   420
         Left            =   4590
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   60
         Width           =   915
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H0080C0FF&
         Caption         =   "Next"
         Height          =   420
         Left            =   5535
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H0080C0FF&
         Caption         =   "Save"
         Height          =   420
         Left            =   1890
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdModify 
         BackColor       =   &H0080C0FF&
         Caption         =   "Modify"
         Height          =   420
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   60
         Width           =   915
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H0080C0FF&
         Caption         =   "Cancel"
         Height          =   420
         Left            =   2790
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H0080C0FF&
         Caption         =   "Add"
         Height          =   420
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   45
         Width           =   870
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H0041E9D8&
      Caption         =   "E&xit"
      Height          =   420
      Left            =   2790
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4380
      Width           =   870
   End
   Begin VB.Frame frame1 
      BackColor       =   &H00000000&
      Height          =   2895
      Left            =   420
      TabIndex        =   5
      Top             =   120
      Width           =   6075
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1860
         TabIndex        =   21
         Text            =   "Combo1"
         Top             =   2280
         Width           =   1890
      End
      Begin VB.TextBox txtAmt 
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
         MaxLength       =   200
         TabIndex        =   19
         Top             =   1800
         Width           =   1860
      End
      Begin VB.CheckBox chkSameDate 
         BackColor       =   &H00000000&
         Caption         =   "Same Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0FF&
         Height          =   375
         Left            =   4200
         TabIndex        =   18
         Top             =   2280
         Width           =   1755
      End
      Begin VB.OptionButton optExpense 
         BackColor       =   &H00000000&
         Caption         =   "&Expense"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   1260
         Width           =   1395
      End
      Begin VB.OptionButton optIncome 
         BackColor       =   &H00000000&
         Caption         =   "&Income"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   1260
         Width           =   1395
      End
      Begin VB.TextBox txtDate 
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
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
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
         TabIndex        =   20
         Top             =   1815
         Width           =   1410
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Source"
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
         TabIndex        =   7
         Top             =   2325
         Width           =   1410
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         TabIndex        =   6
         Top             =   735
         Width           =   1365
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
      Height          =   390
      Left            =   5445
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4380
      Visible         =   0   'False
      Width           =   885
   End
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
      Left            =   5445
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   885
   End
   Begin MSDataGridLib.DataGrid dg 
      Height          =   3015
      Left            =   945
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   180
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
Attribute VB_Name = "frmProfitLoss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rsPL As ADODB.Recordset
Dim Modify, Add, Search, RecordFound As Boolean

' BOOLEANS DECKARED AND THE PURPOSE
' Modify : To check if modify button is clicked
' Add    : To check if Add button is clicked

Private Sub cmdAdd_Click()
On Error GoTo aerr
    Modify = False    'Modify = True only when Modify Button is clicked
    Call Modi   ' This function shows what to do if Modify=True or Modify=False
    txtDate.Enabled = True
    txtSource.Enabled = True
    
    If chkSameDate.Value = 1 Then
        txtDate.Text = ""
    End If
    txtSource.Text = ""
    optIncome.Value = False
    optExpense.Value = False
        
    cmdCancel.Enabled = True
    cmdSave.Enabled = True
    cmdModify.Enabled = False
    cmdRemove.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdAdd.Enabled = False
    cmdClose.Enabled = False
    txtDate.SetFocus
    cmdSave.Enabled = False
    Exit Sub
aerr:
    MsgBox Err.Description, vbOKOnly, "Office Automation"
    frmMaster.Hide
    frmMain.Show
End Sub

Private Sub cmdCancel_Click()
On Error GoTo cerr
    Modify = False
    Call Modi
    txtDate.Enabled = False
    txtSource.Enabled = False
    optIncome.Enabled = False
    optExpense.Enabled = False
    chkSameDate.Enabled = False
    
    cmdAdd.Enabled = True
    cmdModify.Enabled = True
    cmdPrev.Enabled = True
    cmdNext.Enabled = True
    cmdRemove.Enabled = True
    cmdCancel.Enabled = False
    cmdSave.Enabled = False
    cmdClose.Enabled = True
    
    If rsPL.RecordCount = 0 Then
    txtDate.Text = ""
    txtSource.Text = ""
    optIncome.Value = False
    optExpense.Value = False
    
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
    MsgBox Err.Description, vbOKOnly, "Office Automation"
    frmMaster.Hide
    frmMain.Show
End Sub

Private Sub cmdClose_Click()
On Error GoTo eerr
    frmMain.Show
    Unload Me
Exit Sub
eerr:
MsgBox Err.Description, vbOKOnly, "Office Automation"
    frmMaster.Hide
    frmMain.Show
End Sub

Private Sub cmdList_Click()
rsPL.Requery
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
    txtDate.Enabled = True
    txtSource.Enabled = True
    cmdCancel.Enabled = True
    cmdSave.Enabled = True
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdAdd.Enabled = False
    cmdRemove.Enabled = False
    cmdModify.Enabled = False
    cmdClose.Enabled = False
    txtDate.SetFocus
Exit Sub
merr:
MsgBox Err.Description, vbOKOnly, "Office Automation"
    frmMaster.Hide
    frmMain.Show
End Sub

Private Sub cmdNext_Click()
On Error GoTo nerr
    Modify = False
    Call Modi
    
    If rsPL.RecordCount = 0 Then
        txtDate.Text = ""
        txtSource.Text = ""
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
        If rsPL.EOF = False Then rsPL.MoveNext
        If rsPL.EOF = True Then rsPL.MoveLast
        showall
        cmdRemove.Enabled = True

Exit Sub
nerr:
    MsgBox Err.Description, vbOKOnly, "Office Automation"
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
    If rsPL.RecordCount = 0 Then
        txtDate.Text = ""
        txtSource.Text = ""
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
    
    If rsPL.BOF = False Then rsPL.MovePrevious
    If rsPL.BOF = True Then rsPL.MoveFirst
    showall
    cmdRemove.Enabled = True
    Exit Sub
perr:
    MsgBox Err.Description, vbOKOnly, "Office Automation"
    frmMaster.Hide
    frmMain.Show
End Sub

Private Sub cmdRemove_Click()
On Error GoTo rerr
    Modify = False
    Call Modi
If rsPL.RecordCount = 0 Then
    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    cmdCancel.Enabled = False
    cmdSave.Enabled = False
    Exit Sub
End If
      
       rsPL.Delete
        Call cmdNext_Click
Exit Sub
rerr:
MsgBox Err.Description, vbOKOnly, "Office Automation"
    frmMaster.Hide
    frmMain.Show
End Sub
Private Sub FindRecord()
On Error GoTo ErrFind
    If rsPL.RecordCount > 0 Then rsPL.MoveFirst
    For i = 0 To rsPL.RecordCount - 1 Step 1
        If rsPL!Date = txtDate.Text Then
            MsgBox "Date Already Exists!", vbCritical, "Office Automation"
            txtDate.Text = ""
            txtSource.Text = ""
            txtDate.SetFocus
            RecordFound = True
            Exit Sub
        End If
        If rsPL.EOF = False Then rsPL.MoveNext
        If rsPL.EOF = True Then
        RecordFound = False
        End If
    Next
Exit Sub
ErrFind:
    MsgBox Err.Description, vbOKOnly, "Office Automation"
End Sub
Private Sub cmdSave_Click()
On Error GoTo serr

If Trim(txtDate.Text) = "" Or Trim(txtSource.Text) = "" Then
    Exit Sub
End If
If Modify = False Then
Call FindRecord
    If RecordFound = True Then  'If Date already exists then Exit Sub
        Exit Sub
    ElseIf RecordFound = False Then
        rsPL.AddNew  'If Date does not exist then Add this record
    End If
End If
If txtDate.Text <> "" Then rsPL!Date = txtDate.Text
If txtSource.Text <> "" Then rsPL!Source = txtSource.Text

If Year(Date) < Year(rsExpiry!CurrentYear) Then
    MsgBox "You cannot work in previous working year ! Contact Developer", vbCritical, "Office Automation"
Exit Sub
ElseIf Year(Date) > Year(rsExpiry!CurrentYear) Then
    MsgBox "You cannot work in next working year ! Contact Developer", vbCritical, "Office Automation"
Exit Sub
End If

rsPL.Update


cmdAdd.Enabled = True
cmdModify.Enabled = True
cmdPrev.Enabled = True
cmdNext.Enabled = True
cmdCancel.Enabled = False
cmdRemove.Enabled = True
cmdClose.Enabled = True
cmdSave.Enabled = False

txtDate.Enabled = False
txtSource.Enabled = False

If cmdAdd.Enabled = True Then cmdAdd.SetFocus
Exit Sub
serr:
MsgBox Err.Description, vbOKOnly, "Office Automation"
    frmMaster.Hide
    frmMain.Show
End Sub

Private Sub Form_Activate()
cmdAdd.SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo ferr
    Call MakeConnection
    Set dg.Data = rsPL
    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    cmdCancel.Enabled = False
If rsPL.RecordCount > 0 Then
    cmdAdd.Enabled = True
    cmdPrev.Enabled = True
    cmdNext.Enabled = True
End If
    cmdSave.Enabled = False
Exit Sub
ferr:
MsgBox Err.Description, vbOKOnly, "Office Automation"
    frmMaster.Hide
    frmMain.Show
End Sub

Private Sub optIncome_Click()

End Sub

Private Sub txtDate_GotFocus()
Call txt_GotFocus
txtDate.BackColor = &HC00000
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtDate_KeyUp(KeyDate As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtDate_LostFocus()
txtDate.BackColor = vbBlack
End Sub

Private Sub txtDesc_GotFocus()
Call txt_GotFocus
txtSource.BackColor = &HC00000
End Sub

Private Sub txtSource_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtSource_KeyUp(KeyDate As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub showall()
'On Error GoTo showerr
txtDate.Text = ""
txtSource.Text = ""

If rsPL.RecordCount = 0 Then
    txtDate.Text = ""
    txtSource.Text = ""
    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    Exit Sub
ElseIf rsPL.RecordCount > 0 Then
    cmdModify.Enabled = True
    cmdPrev.Enabled = True
    cmdNext.Enabled = True
    cmdRemove.Enabled = True
End If
If Not IsNull(rsPL!Source) Then txtSource.Text = rsPL!Source
If Not IsNull(rsPL!Date) Then txtDate.Text = rsPL!Date
Exit Sub
showerr:
MsgBox Err.Description, vbOKOnly, "Office Automation"
End Sub


Private Sub txt_GotFocus()
On Error GoTo focerr
    If Trim(txtSource.Text) = "" Or Trim(txtDate.Text) = "" Then
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
txtDate.Enabled = False
txtSource.Enabled = False

ElseIf Modify = True Then  ' If cmdmodify is not clicked then
txtDate.Enabled = True
txtSource.Enabled = True
End If
Exit Sub
merr:
MsgBox Err.Description, vbOKOnly, "Office Automation"
End Sub

Private Sub txtSource_LostFocus()
txtSource.BackColor = vbBlack
End Sub

