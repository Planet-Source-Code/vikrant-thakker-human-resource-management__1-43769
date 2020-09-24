VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmIncomeExpense 
   BackColor       =   &H00000000&
   Caption         =   "Income Data Entry Form"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   6735
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
      Left            =   4260
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1500
      Width           =   765
   End
   Begin VB.Frame Framebutton 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   120
      TabIndex        =   8
      Top             =   3630
      Width           =   6495
      Begin VB.CommandButton cmdRemove 
         BackColor       =   &H0080C0FF&
         Caption         =   "Remove"
         Height          =   420
         Left            =   3690
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdPrev 
         BackColor       =   &H0080C0FF&
         Caption         =   "Previous"
         Height          =   420
         Left            =   4590
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   60
         Width           =   915
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H0080C0FF&
         Caption         =   "Next"
         Height          =   420
         Left            =   5535
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H0080C0FF&
         Caption         =   "Save"
         Height          =   420
         Left            =   1890
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdModify 
         BackColor       =   &H0080C0FF&
         Caption         =   "Modify"
         Height          =   420
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   60
         Width           =   915
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H0080C0FF&
         Caption         =   "Cancel"
         Height          =   420
         Left            =   2790
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H0080C0FF&
         Caption         =   "Add"
         Height          =   420
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   45
         Width           =   870
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H0041E9D8&
      Caption         =   "E&xit"
      Height          =   420
      Left            =   2910
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4440
      Width           =   870
   End
   Begin VB.Frame frame1 
      BackColor       =   &H00000000&
      Height          =   3135
      Left            =   300
      TabIndex        =   4
      Top             =   180
      Width           =   6075
      Begin VB.ComboBox cmbType 
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
         Height          =   360
         Left            =   1860
         TabIndex        =   1
         Top             =   1320
         Width           =   1935
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
         MaxLength       =   7
         TabIndex        =   2
         Top             =   1980
         Width           =   1860
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
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   0
         Top             =   720
         Width           =   1860
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "DD/MM/YYYY"
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
         Left            =   3960
         TabIndex        =   19
         Top             =   720
         Width           =   1365
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
         TabIndex        =   16
         Top             =   1995
         Width           =   1410
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Income Type"
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
         Top             =   1365
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
         TabIndex        =   5
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
      Left            =   5565
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4440
      Visible         =   0   'False
      Width           =   885
   End
   Begin MSDataGridLib.DataGrid dg 
      Height          =   3615
      Left            =   1200
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   6376
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
Attribute VB_Name = "frmIncomeExpense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Flag, MNT As String
Public rsEntry As ADODB.Recordset
Public rsList As ADODB.Recordset
Private Sub MakeConnection()
On Err GoTo errConn
    If FormName = "IncomeEntry" Then
        Set rsEntry = New ADODB.Recordset
        rsEntry.Open "select * from ProfitLoss where Inc_Exp='Income'", conn, adOpenStatic, adLockOptimistic
        
        Set rsList = New ADODB.Recordset
        rsList.Open "select * from MastIncome", conn, adOpenStatic, adLockOptimistic
        
        Label2.Caption = "Income Type"
    ElseIf FormName = "ExpenseEntry" Then
        Set rsEntry = New ADODB.Recordset
        rsEntry.Open "select * from ProfitLoss where Inc_Exp='Expense'", conn, adOpenStatic, adLockOptimistic
    
        Set rsList = New ADODB.Recordset
        rsList.Open "select * from MastExpense", conn, adOpenStatic, adLockOptimistic
        
        Label2.Caption = "Expense Type"
    End If
Exit Sub
errConn:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub cmdAdd_Click()
On Error GoTo aerr
    Modify = False    'Modify = True only when Modify Button is clicked
    Call Modi   ' This function shows what to do if Modify=True or Modify=False
    txtDate.Enabled = True
    txtAmt.Enabled = True
    cmbType.Enabled = True
    
    txtDate.Text = ""
    txtAmt.Text = ""
    cmbType.Text = ""
    
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
    MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub cmdCancel_Click()
On Error GoTo cerr
    Modify = False
    Call Modi
    txtDate.Enabled = False
    txtAmt.Enabled = False
    cmbType.Enabled = False
    
    cmdAdd.Enabled = True
    cmdModify.Enabled = True
    cmdPrev.Enabled = True
    cmdNext.Enabled = True
    cmdRemove.Enabled = True
    cmdCancel.Enabled = False
    cmdSave.Enabled = False
    cmdClose.Enabled = True
    
    If rsEntry.RecordCount = 0 Then
    txtDate.Text = ""
    txtAmt.Text = ""
    cmbType.Text = ""
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
End Sub

Private Sub cmdClose_Click()
On Error GoTo eerr
    frmProfitLossMain.Show
    Unload Me
Exit Sub
eerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub cmdList_Click()
rsEntry.Requery
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
    txtAmt.Enabled = True
    cmbType.Enabled = True
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
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub cmdNext_Click()
On Error GoTo nerr
    Modify = False
    Call Modi
    
    If rsEntry.RecordCount = 0 Then
        txtDate.Text = ""
        txtAmt.Text = ""
        cmbType.Text = ""
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
        If rsEntry.EOF = False Then rsEntry.MoveNext
        If rsEntry.EOF = True Then rsEntry.MoveLast
        showall
        cmdRemove.Enabled = True

Exit Sub
nerr:
    MsgBox Err.Description, vbOKOnly, "OASYS"
    
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
    If rsEntry.RecordCount = 0 Then
        txtDate.Text = ""
        txtAmt.Text = ""
        cmbType.Text = ""
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
    
    If rsEntry.BOF = False Then rsEntry.MovePrevious
    If rsEntry.BOF = True Then rsEntry.MoveFirst
    showall
    cmdRemove.Enabled = True
    Exit Sub
perr:
    MsgBox Err.Description, vbOKOnly, "OASYS"
    
End Sub

Private Sub cmdRemove_Click()
On Error GoTo rerr
    Modify = False
    Call Modi
If rsEntry.RecordCount = 0 Then
    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    cmdCancel.Enabled = False
    cmdSave.Enabled = False
    Exit Sub
End If
       rsEntry.Delete
        Call cmdNext_Click
Exit Sub
rerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
    
End Sub

Private Sub cmdSave_Click()
On Error GoTo serr

If Trim(txtDate.Text) = "" Or Trim(txtAmt.Text) = "" Or Trim(cmbType.Text) = "" Then
    Exit Sub
End If
If Modify = False Then
        rsEntry.AddNew  'If Date does not exist then Add this record
End If
If txtDate.Text <> "" Then rsEntry!Date = txtDate.Text
If txtAmt.Text <> "" Then rsEntry!Amt = txtAmt.Text
If cmbType.Text <> "" Then rsEntry!Type = cmbType.Text
If FormName = "IncomeEntry" Then
    rsEntry!Inc_Exp = "Income"
ElseIf FormName = "ExpenseEntry" Then
    rsEntry!Inc_Exp = "Expense"
End If

If Year(Date) < Year(rsExpiry!CurrentYear) Then
    MsgBox "You cannot work in previous working year ! Contact Developer", vbCritical, "OASYS"
Exit Sub
ElseIf Year(Date) > Year(rsExpiry!CurrentYear) Then
    MsgBox "You cannot work in next working year ! Contact Developer", vbCritical, "OASYS"
Exit Sub
End If

rsEntry.Update
Call MntProfitLoss

cmdAdd.Enabled = True
cmdModify.Enabled = True
cmdPrev.Enabled = True
cmdNext.Enabled = True
cmdCancel.Enabled = False
cmdRemove.Enabled = True
cmdClose.Enabled = True
cmdSave.Enabled = False

txtDate.Enabled = False
txtAmt.Enabled = False
cmbType.Enabled = False
If cmdAdd.Enabled = True Then cmdAdd.SetFocus
Exit Sub
serr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub
Private Sub GetMonth()
If Month(txtDate.Text) = 1 Then
    MNT = "JANUARY"
ElseIf Month(txtDate.Text) = 2 Then
    MNT = "FEBRUARY"
ElseIf Month(txtDate.Text) = 3 Then
    MNT = "MARCH"
ElseIf Month(txtDate.Text) = 4 Then
    MNT = "APRIL"
ElseIf Month(txtDate.Text) = 5 Then
    MNT = "MAY"
ElseIf Month(txtDate.Text) = 6 Then
    MNT = "JUNE"
ElseIf Month(txtDate.Text) = 7 Then
    MNT = "JULY"
ElseIf Month(txtDate.Text) = 8 Then
    MNT = "AUGUST"
ElseIf Month(txtDate.Text) = 9 Then
    MNT = "SEPTEMBER"
ElseIf Month(txtDate.Text) = 10 Then
    MNT = "OCTOBER"
ElseIf Month(txtDate.Text) = 11 Then
    MNT = "NOVEMBER"
ElseIf Month(txtDate.Text) = 12 Then
    MNT = "DECEMBER"
End If
End Sub
Private Sub MntProfitLoss()
Call GetMonth
If rsMntProfitLoss.RecordCount > 0 Then rsMntProfitLoss.MoveFirst

For i = 0 To rsMntProfitLoss.RecordCount - 1 Step 1
    If rsMntProfitLoss!Year = Year(txtDate.Text) And rsMntProfitLoss!Month = MNT Then
        If FormName = "IncomeEntry" Then
            rsMntProfitLoss!Income = rsMntProfitLoss!Income + Val(txtAmt.Text)
            rsMntProfitLoss!Expense = rsMntProfitLoss!Expense + 0
        ElseIf FormName = "ExpenseEntry" Then
            rsMntProfitLoss!Expense = rsMntProfitLoss!Expense + Val(txtAmt.Text)
            rsMntProfitLoss!Income = rsMntProfitLoss!Income + 0
        End If
            
            rsMntProfitLoss!NetProfit = rsMntProfitLoss!Income - rsMntProfitLoss!Expense
                    
        rsMntProfitLoss.Update
        Exit Sub
    End If
If rsMntProfitLoss.EOF = False Then rsMntProfitLoss.MoveNext
Next
        rsMntProfitLoss.AddNew
        rsMntProfitLoss!Year = Year(txtDate.Text)
        rsMntProfitLoss!Month = MNT
        If FormName = "IncomeEntry" Then
            rsMntProfitLoss!Income = rsMntProfitLoss!Income + Val(txtAmt.Text)
            rsMntProfitLoss!Expense = 0
        ElseIf FormName = "ExpenseEntry" Then
            rsMntProfitLoss!Expense = rsMntProfitLoss!Expense + Val(txtAmt.Text)
            rsMntProfitLoss!Income = 0
        End If
            
            rsMntProfitLoss!NetProfit = rsMntProfitLoss!Income - rsMntProfitLoss!Expense
        rsMntProfitLoss.Update
End Sub
Private Sub Form_Activate()
cmdAdd.SetFocus
End Sub

Private Sub Form_Load()
'On Error GoTo ferr
Call MakeConnection
Set dg.DataSource = rsList
    
If rsList.RecordCount > 0 Then
rsList.MoveFirst
    For i = 0 To rsList.RecordCount - 1 Step 1
cmbType.AddItem (rsList!Code)
rsList.MoveNext
    Next
End If
   
    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    cmdCancel.Enabled = False
If rsEntry.RecordCount > 0 Then
    cmdAdd.Enabled = True
    cmdPrev.Enabled = True
    cmdNext.Enabled = True
End If
    cmdSave.Enabled = False
Exit Sub
ferr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub txtDate_GotFocus()
Call txt_GotFocus
txtDate.BackColor = &HC00000
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    datevali (txtDate.Text)
    'SendKeys "{TAB}"
    'KeyAscii = 0
End If
End Sub

Private Sub txtDate_KeyUp(KeyDate As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtDate_LostFocus()
txtDate.BackColor = vbBlack
End Sub

Private Sub txtAmt_GotFocus()
Call txt_GotFocus
txtAmt.BackColor = &HC00000
End Sub

Private Sub txtAmt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtAmt_KeyUp(KeyDate As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtAmt_LostFocus()
txtAmt.BackColor = vbBlack
End Sub

Private Sub cmbType_GotFocus()
Call txt_GotFocus
cmbType.BackColor = &HC00000
End Sub

Private Sub cmbType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub cmbType_KeyUp(KeyDate As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub cmbType_LostFocus()
cmbType.BackColor = vbBlack
End Sub

Private Sub showall()
'On Error GoTo showerr
txtDate.Text = ""
txtAmt.Text = ""
cmbType.Text = ""
If rsEntry.RecordCount = 0 Then
    txtDate.Text = ""
    txtAmt.Text = ""
    cmbType.Text = ""
    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    Exit Sub
ElseIf rsEntry.RecordCount > 0 Then
    cmdModify.Enabled = True
    cmdPrev.Enabled = True
    cmdNext.Enabled = True
    cmdRemove.Enabled = True
End If
If Not IsNull(rsEntry!Amt) Then txtAmt.Text = rsEntry!Amt
If Not IsNull(rsEntry!Date) Then txtDate.Text = rsEntry!Date
If Not IsNull(rsEntry!Type) Then cmbType.Text = rsEntry!Type
Exit Sub
showerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub


Private Sub txt_GotFocus()
On Error GoTo focerr
    If Trim(txtAmt.Text) = "" Or Trim(txtDate.Text) = "" Or Trim(cmbType.Text) = "" Then
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
txtDate.Enabled = False
txtAmt.Enabled = False
cmbType.Enabled = False
ElseIf Modify = True Then  ' If cmdmodify is not clicked then
txtDate.Enabled = True
txtAmt.Enabled = True
cmbType.Enabled = True
End If
Exit Sub
merr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub




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
    MsgBox "Please enter year Between financial year"
    Me.SetFocus
Else
    SendKeys "{TAB}"
    KeyAscii = 0
End If

Exit Function
dvalerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Function
