VERSION 5.00
Begin VB.Form frmMastWD 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Working Days Master"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbMnt 
      Height          =   315
      ItemData        =   "frmMastWD.frx":0000
      Left            =   2880
      List            =   "frmMastWD.frx":0002
      TabIndex        =   1
      Top             =   1260
      Width           =   1860
   End
   Begin VB.TextBox txtWD 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   2880
      MaxLength       =   5
      TabIndex        =   2
      Top             =   1890
      Width           =   1860
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H0041E9D8&
      Caption         =   "E&xit"
      Height          =   420
      Left            =   2925
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3600
      Width           =   870
   End
   Begin VB.Frame Framebutton 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   135
      TabIndex        =   5
      Top             =   2790
      Width           =   6495
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
      Begin VB.CommandButton cmdModify 
         BackColor       =   &H0080C0FF&
         Caption         =   "Modify"
         Height          =   420
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   45
         Width           =   915
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
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H0080C0FF&
         Caption         =   "Next"
         Height          =   420
         Left            =   5535
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdPrev 
         BackColor       =   &H0080C0FF&
         Caption         =   "Previous"
         Height          =   420
         Left            =   4590
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   45
         Width           =   915
      End
      Begin VB.CommandButton cmdRemove 
         BackColor       =   &H0080C0FF&
         Caption         =   "Remove"
         Height          =   420
         Left            =   3660
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H0080C0FF&
         Caption         =   "Add"
         Height          =   420
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   45
         Width           =   870
      End
   End
   Begin VB.TextBox txtYear 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   2880
      MaxLength       =   5
      TabIndex        =   0
      Top             =   630
      Width           =   1860
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Working Days"
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
      Left            =   1215
      TabIndex        =   14
      Top             =   2025
      Width           =   1365
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
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
      Left            =   1935
      TabIndex        =   4
      Top             =   1395
      Width           =   690
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
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
      Left            =   2055
      TabIndex        =   3
      Top             =   765
      Width           =   735
   End
End
Attribute VB_Name = "frmMastWD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim Modify, RecordFound As Boolean

Private Sub cmdAdd_Click()
On Error GoTo aerr
Modify = False
Call Modi
txtYear.Enabled = True
cmbMnt.Enabled = True
txtWD.Enabled = True
       
txtYear.Text = ""
cmbMnt.Text = ""
txtWD.Text = ""

'rsWD.AddNew

cmdCancel.Enabled = True
cmdSave.Enabled = True
cmdModify.Enabled = False
cmdRemove.Enabled = False
cmdPrev.Enabled = False
cmdNext.Enabled = False
cmdAdd.Enabled = False
cmdClose.Enabled = False

txtYear.SetFocus
cmdSave.Enabled = False
Exit Sub
aerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub cmdCancel_Click()
On Error GoTo cerr
Modify = False
Call Modi
'rsWD.CancelUpdate
txtYear.Enabled = False
cmbMnt.Enabled = False
txtWD.Enabled = False

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

Private Sub cmdClose_Click()
On Error GoTo eerr
frmMain.Show
Unload Me
Exit Sub
eerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub cmdModify_Click()
On Error GoTo merr
Modify = True
Call Modi
txtYear.Enabled = True
cmbMnt.Enabled = True
txtWD.Enabled = True

cmdCancel.Enabled = True
cmdSave.Enabled = True
cmdPrev.Enabled = False
cmdNext.Enabled = False
cmdAdd.Enabled = False
cmdRemove.Enabled = False
cmdModify.Enabled = False
cmdClose.Enabled = False

txtYear.SetFocus
Exit Sub
merr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub cmdNext_Click()
On Error GoTo nerr
Modify = False
Call Modi

If rsWD.RecordCount = 0 Then
txtYear.Text = ""
cmbMnt.Text = ""
txtWD.Text = ""
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
If rsWD.EOF = False Then rsWD.MoveNext
If rsWD.EOF = True Then rsWD.MoveLast
    showall
Exit Sub
nerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub cmdPrev_Click()
On Error GoTo perr
Modify = False
Call Modi

If rsWD.RecordCount = 0 Then
txtYear.Text = ""
cmbMnt.Text = ""
txtWD.Text = ""
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

If rsWD.BOF = False Then rsWD.MovePrevious
If rsWD.BOF = True Then rsWD.MoveFirst
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
 If rsWD.RecordCount = 0 Then
    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    cmdCancel.Enabled = False
    cmdSave.Enabled = False
    Exit Sub
End If
       rsWD.Delete
       Call cmdNext_Click
Exit Sub
rerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub FindRecord()
If rsWD.RecordCount > 0 Then rsWD.MoveFirst
For i = 0 To rsWD.RecordCount - 1 Step 1
    If rsWD!Year = txtYear.Text And rsWD!Month = cmbMnt.ListIndex + 1 Then
    MsgBox cmbMnt.ListIndex + 1
        MsgBox "Year and Month Already Exists!", vbCritical, "OASYS"
        txtYear.Text = ""
        cmbMnt.Text = ""
        txtYear.SetFocus
        RecordFound = True
        Exit Sub
    End If
    If rsWD.EOF = False Then rsWD.MoveNext
If rsWD.EOF = True Then
    RecordFound = False
End If
Next
End Sub
Private Sub cmdSave_Click()
On Error GoTo serr

If Trim(txtYear.Text) = "" Or Trim(cmbMnt.Text) = "" Or Trim(txtWD.Text) = "" Then
    Exit Sub
End If

If Modify = False Then
Call FindRecord
    If RecordFound = True Then  'If Code already exists then Exit Sub
        Exit Sub
    ElseIf RecordFound = False Then
        rsWD.AddNew  'If Code does not exist then Add this record
    End If
End If

If txtYear.Text <> "" Then rsWD!Year = txtYear.Text
If cmbMnt.Text <> "" Then rsWD!Month = cmbMnt.ListIndex + 1
If txtWD.Text <> "" Then rsWD!WD = txtWD.Text

If Year(Date) < Year(rsExpiry!CurrentYear) Then
    MsgBox "You cannot work in previous working year ! Contact Developer", vbCritical, "OASYS"
Exit Sub
ElseIf Year(Date) > Year(rsExpiry!CurrentYear) Then
    MsgBox "You cannot work in next working year ! Contact Developer", vbCritical, "OASYS"
Exit Sub
End If

rsWD.Update

cmdSave.Enabled = False
cmdAdd.Enabled = True
cmdModify.Enabled = True
cmdPrev.Enabled = True
cmdNext.Enabled = True
cmdCancel.Enabled = False
cmdRemove.Enabled = True
cmdClose.Enabled = True
cmdSave.Enabled = False

txtYear.Enabled = False
cmbMnt.Enabled = False
txtWD.Enabled = False
'txtYear.Text = ""
'cmbMnt.Text = ""
'txtwd.Text = ""
'txtYear.SetFocus

If cmdAdd.Enabled = True Then cmdAdd.SetFocus
Exit Sub
serr:
MsgBox Err.Description, vbOKOnly, "OASYS"

txtYear.Enabled = True
cmbMnt.Enabled = True

txtYear.Text = ""
cmbMnt.Text = ""
txtWD.Text = ""
txtYear.SetFocus
End Sub
Private Sub AddMonths()
'Add the names of Months in the Month combobox
cmbMnt.AddItem ("January")
cmbMnt.AddItem ("February")
cmbMnt.AddItem ("March")
cmbMnt.AddItem ("April")
cmbMnt.AddItem ("May")
cmbMnt.AddItem ("June")
cmbMnt.AddItem ("July")
cmbMnt.AddItem ("August")
cmbMnt.AddItem ("September")
cmbMnt.AddItem ("October")
cmbMnt.AddItem ("November")
cmbMnt.AddItem ("December")
End Sub
Private Sub Form_Activate()
cmdAdd.SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo ferr

Call AddMonths  'This function Adds names of the months in the combobox
    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    cmdCancel.Enabled = False
    
 ' End If
    
If rsWD.RecordCount > 0 Then
    cmdAdd.Enabled = True
    cmdPrev.Enabled = True
    cmdNext.Enabled = True
End If
    cmdSave.Enabled = False
    
Exit Sub
ferr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub txtYear_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtYear_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtYear_KeyUp(KeyYear As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub cmbMnt_GotFocus()
Call txt_GotFocus
End Sub

Private Sub cmbMnt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub cmbMnt_KeyUp(KeyYear As Integer, Shift As Integer)
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

Private Sub txtWD_KeyUp(KeyYear As Integer, Shift As Integer)
Call txt_GotFocus
End Sub
Private Sub showall()
On Error GoTo showerr
txtYear.Text = ""
cmbMnt.Text = ""
txtWD.Text = ""

If rsWD.RecordCount = 0 Then
    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    Exit Sub
ElseIf rsWD.RecordCount > 0 Then
    cmdModify.Enabled = True
    cmdPrev.Enabled = True
    cmdNext.Enabled = True
    cmdRemove.Enabled = True
End If
If Not IsNull(rsWD!WD) Then txtWD.Text = rsWD!WD
If Not IsNull(rsWD!Year) Then txtYear.Text = rsWD!Year
If Not IsNull(rsWD!Month) Then cmbMnt.Text = cmbMnt.List(rsWD!Month - 1)
Exit Sub
showerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub


Private Sub txt_GotFocus()
On Error GoTo focerr
    If Trim(cmbMnt.Text) = "" Or Trim(txtYear.Text) = "" Or Trim(txtWD.Text) = "" Then
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
txtYear.Enabled = False
cmbMnt.Enabled = False
txtWD.Enabled = False

ElseIf Modify = True Then  ' If cmdmodify is not clicked then
txtYear.Enabled = True
cmbMnt.Enabled = True
txtWD.Enabled = True
End If
Exit Sub
merr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub
