VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmMastEmp 
   BackColor       =   &H00000000&
   Caption         =   "Employee Master"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPrintICard 
      BackColor       =   &H0041E9D8&
      Caption         =   "GENERATE IDENTITY &CARD"
      Height          =   780
      Left            =   10140
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   7140
      Width           =   1650
   End
   Begin VB.Frame Framebutton 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   2700
      TabIndex        =   40
      Top             =   7110
      Width           =   6495
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H0080C0FF&
         Caption         =   "Add"
         Height          =   420
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdRemove 
         BackColor       =   &H0080C0FF&
         Caption         =   "Remove"
         Height          =   420
         Left            =   3690
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdPrev 
         BackColor       =   &H0080C0FF&
         Caption         =   "Previous"
         Height          =   420
         Left            =   4590
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   45
         Width           =   915
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H0080C0FF&
         Caption         =   "Next"
         Height          =   420
         Left            =   5535
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H0080C0FF&
         Caption         =   "Save"
         Height          =   420
         Left            =   1890
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdModify 
         BackColor       =   &H0080C0FF&
         Caption         =   "Modify"
         Height          =   420
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   45
         Width           =   915
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H0080C0FF&
         Caption         =   "Cancel"
         Height          =   420
         Left            =   2790
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   45
         Width           =   870
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H0041E9D8&
      Caption         =   "E&xit"
      Height          =   420
      Left            =   5490
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   7830
      Width           =   870
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   7080
      Left            =   60
      TabIndex        =   18
      Top             =   -60
      Width           =   11760
      Begin Crystal.CrystalReport CR 
         Left            =   300
         Top             =   3900
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H0041E9D8&
         Caption         =   "&Search"
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
         Left            =   3645
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   1215
         Width           =   780
      End
      Begin VB.ComboBox cmbDesig 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2610
         TabIndex        =   2
         Top             =   2385
         Width           =   2805
      End
      Begin VB.ComboBox cmbCaste 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2610
         TabIndex        =   6
         Top             =   5400
         Width           =   1320
      End
      Begin VB.CommandButton cmbBrowse 
         Caption         =   "&Browse"
         Height          =   375
         Left            =   10800
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   5895
         Width           =   735
      End
      Begin MSComDlg.CommonDialog cd 
         Left            =   10920
         Top             =   5280
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txtPath 
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         DataField       =   " "
         DataSource      =   " "
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7680
         TabIndex        =   17
         Text            =   " "
         Top             =   5940
         Width           =   3045
      End
      Begin VB.ComboBox cmbSection 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7695
         TabIndex        =   12
         Text            =   "Combo3"
         Top             =   2925
         Width           =   1950
      End
      Begin VB.ComboBox cmbType 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7695
         TabIndex        =   11
         Text            =   "Combo2"
         Top             =   2385
         Width           =   1140
      End
      Begin VB.ComboBox cmbSex 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7695
         TabIndex        =   10
         Text            =   "Male"
         Top             =   1800
         Width           =   1185
      End
      Begin VB.ComboBox cmbClass 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2610
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   5940
         Width           =   2850
      End
      Begin VB.CommandButton Command1 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   9750
         TabIndex        =   20
         Top             =   7425
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Help"
         Height          =   420
         Left            =   1350
         TabIndex        =   19
         Top             =   7425
         Width           =   1155
      End
      Begin VB.Timer Timer1 
         Interval        =   1500
         Left            =   3270
         Top             =   225
      End
      Begin VB.TextBox txtPhone 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   " "
         DataSource      =   " "
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   2595
         TabIndex        =   5
         Text            =   " "
         Top             =   4830
         Width           =   2775
      End
      Begin VB.TextBox txtFName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   " "
         DataSource      =   " "
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   2595
         TabIndex        =   3
         Text            =   " "
         Top             =   2940
         Width           =   2775
      End
      Begin VB.TextBox txtAdd 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   " "
         DataSource      =   " "
         ForeColor       =   &H00000000&
         Height          =   1095
         Left            =   2595
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   3510
         Width           =   2775
      End
      Begin VB.TextBox txtSalary 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   " "
         DataSource      =   " "
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2565
         TabIndex        =   8
         Text            =   " "
         Top             =   6510
         Width           =   5295
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   " "
         DataSource      =   " "
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   2595
         TabIndex        =   1
         Text            =   " "
         Top             =   1815
         Width           =   2775
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   " "
         DataSource      =   " "
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   2595
         MousePointer    =   10  'Up Arrow
         TabIndex        =   0
         ToolTipText     =   "Employee Code"
         Top             =   1260
         Width           =   975
      End
      Begin VB.TextBox txtQual 
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         DataField       =   " "
         DataSource      =   " "
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7680
         TabIndex        =   16
         Text            =   " "
         Top             =   5385
         Width           =   1695
      End
      Begin VB.TextBox txtDOB 
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         DataField       =   " "
         DataSource      =   " "
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   7680
         MaxLength       =   10
         TabIndex        =   13
         Text            =   " "
         Top             =   3540
         Width           =   1320
      End
      Begin VB.TextBox txtDOJ 
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         DataField       =   " "
         DataSource      =   " "
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   7680
         MaxLength       =   10
         TabIndex        =   14
         Text            =   " "
         Top             =   4260
         Width           =   1335
      End
      Begin VB.TextBox txtDOR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         DataField       =   " "
         DataSource      =   " "
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7680
         MaxLength       =   10
         TabIndex        =   15
         Text            =   " "
         Top             =   4800
         Width           =   1350
      End
      Begin VB.TextBox txtBasicPay 
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         DataField       =   " "
         DataSource      =   " "
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7680
         TabIndex        =   9
         Text            =   " "
         Top             =   1215
         Width           =   1215
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "DD/MM/YYYY"
         ForeColor       =   &H00EBCCB4&
         Height          =   375
         Left            =   9120
         TabIndex        =   54
         Top             =   4860
         Width           =   1230
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "DD/MM/YYYY"
         ForeColor       =   &H00EBCCB4&
         Height          =   375
         Left            =   9120
         TabIndex        =   52
         Top             =   3600
         Width           =   1230
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "DD/MM/YYYY"
         ForeColor       =   &H00EBCCB4&
         Height          =   375
         Left            =   9120
         TabIndex        =   51
         Top             =   4335
         Width           =   1230
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "PIC PATH"
         ForeColor       =   &H00EBCCB4&
         Height          =   195
         Left            =   6525
         TabIndex        =   48
         Top             =   5985
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "EMP CODE"
         ForeColor       =   &H00EBCCB4&
         Height          =   195
         Left            =   1470
         TabIndex        =   38
         Top             =   1350
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "NAME"
         ForeColor       =   &H00EBCCB4&
         Height          =   195
         Left            =   1845
         TabIndex        =   37
         Top             =   1890
         Width           =   480
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "SECTION"
         ForeColor       =   &H00EBCCB4&
         Height          =   195
         Left            =   6555
         TabIndex        =   36
         Top             =   3030
         Width           =   720
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "DESIGNATION"
         ForeColor       =   &H00EBCCB4&
         Height          =   195
         Left            =   1155
         TabIndex        =   35
         Top             =   2460
         Width           =   1125
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "FATHER'S NAME"
         ForeColor       =   &H00EBCCB4&
         Height          =   195
         Left            =   960
         TabIndex        =   34
         Top             =   3030
         Width           =   1290
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "CASTE"
         ForeColor       =   &H00EBCCB4&
         Height          =   195
         Left            =   1725
         TabIndex        =   33
         Top             =   5460
         Width           =   525
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "CLASS"
         ForeColor       =   &H00EBCCB4&
         Height          =   195
         Left            =   1755
         TabIndex        =   32
         Top             =   6000
         Width           =   510
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   " ADDRESS"
         ForeColor       =   &H00EBCCB4&
         Height          =   195
         Left            =   1395
         TabIndex        =   31
         Top             =   3555
         Width           =   825
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "PHONE"
         ForeColor       =   &H00EBCCB4&
         Height          =   195
         Left            =   1665
         TabIndex        =   30
         Top             =   4890
         Width           =   570
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "SALARY"
         ForeColor       =   &H00EBCCB4&
         Height          =   195
         Left            =   1650
         TabIndex        =   29
         Top             =   6615
         Width           =   630
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "EMPLOYEE MASTER"
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
         Left            =   4740
         TabIndex        =   28
         Top             =   225
         Width           =   2445
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "SEX"
         ForeColor       =   &H00EBCCB4&
         Height          =   195
         Left            =   6930
         TabIndex        =   27
         Top             =   1830
         Width           =   315
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "TYPE"
         ForeColor       =   &H00EBCCB4&
         Height          =   195
         Left            =   6840
         TabIndex        =   26
         Top             =   2430
         Width           =   420
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "DATE OF BIRTH"
         ForeColor       =   &H00EBCCB4&
         Height          =   195
         Left            =   6015
         TabIndex        =   25
         Top             =   3630
         Width           =   1230
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   " DATE OF JOIN"
         ForeColor       =   &H00EBCCB4&
         Height          =   195
         Left            =   6135
         TabIndex        =   24
         Top             =   4320
         Width           =   1140
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "RETIRING DATE"
         ForeColor       =   &H00EBCCB4&
         Height          =   195
         Left            =   6000
         TabIndex        =   23
         Top             =   4860
         Width           =   1260
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "QUALIFICATION"
         ForeColor       =   &H00EBCCB4&
         Height          =   195
         Left            =   6045
         TabIndex        =   22
         Top             =   5415
         Width           =   1215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "BASIC PAY"
         ForeColor       =   &H00EBCCB4&
         Height          =   195
         Left            =   6480
         TabIndex        =   21
         Top             =   1275
         Width           =   825
      End
      Begin VB.Image Img 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1845
         Left            =   9690
         Stretch         =   -1  'True
         Top             =   270
         Width           =   1860
      End
   End
End
Attribute VB_Name = "frmMastEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 'Option Explicit
Dim Modify, RecordFound As Boolean
'This will open the 'Open File' Dialogbox, thus making it easier
'for user to select the path of the picture file.
Private Sub cmbBrowse_Click()
On Error GoTo errbrowse
cd.Action = 1   'Returns or sets the type of dialog box to be displayed.. ie. 1 = Open Dialog box, 2 = Save As dialogbox, 3 = Color Dialogbox....
txtPath.Text = cd.FileName  'This will enter the Path of the filename that we have selected in the dialog box
If txtPath.Text <> "" Then
    Img.Picture = LoadPicture(txtPath.Text) ' Loads the picture from the entered path into the Image Box
ElseIf txtPath.Text = "" Then
    Img.Picture = LoadPicture(none)
End If
Exit Sub
errbrowse:
MsgBox "Invalid File!", vbCritical, "OASYS"
End Sub

Private Sub cmbCaste_GotFocus()
Call txt_GotFocus
End Sub

Private Sub cmbCaste_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub cmbCaste_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub cmbClass_GotFocus()
Call txt_GotFocus
End Sub

Private Sub cmbClass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub cmbClass_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub cmbSection_GotFocus()
Call txt_GotFocus
End Sub

Private Sub cmbSection_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub cmbSection_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub
Private Sub cmbSex_GotFocus()
Call txt_GotFocus
End Sub

Private Sub cmbSex_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub cmbSex_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub cmbType_GotFocus()
Call txt_GotFocus
End Sub

Private Sub cmbType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub cmbType_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub
'======================================
Private Sub cmdPrintICard_Click()
On Error GoTo perr
If rsICard.RecordCount > 0 Then rsICard.MoveFirst
    For i = 0 To rsICard.RecordCount - 1 Step 1
    If rsICard.BOF = False Or rsICard.EOF = False Then
        rsICard.Delete
        If rsICard.EOF = False Then rsICard.MoveNext
    End If
    Next
Call AddPrintRecord
Call GenerateICard
    Ans = MsgBox("Do you want to create an I-Card ?", vbYesNo, "OASYS")
    If Ans = vbYes Then
        Cr.Action = 1
    Else
        Exit Sub
    End If
Exit Sub
perr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub
Private Sub AddPrintRecord()
rsICard.AddNew
If Trim(txtCode.Text) <> "" Then rsICard!Code = txtCode.Text
If Trim(txtName.Text) <> "" Then rsICard!Name = txtName.Text
If Trim(cmbDesig.Text) <> "" Then rsICard!Desig = cmbDesig.Text
If Trim(txtDOJ.Text) <> "" Then rsICard!DOJ = txtDOJ.Text
If Trim(txtDOR.Text) <> "" Then rsICard!DOR = txtDOR.Text
If Trim(txtAdd.Text) <> "" Then rsICard!address = txtAdd.Text
rsICard.Update
End Sub

Private Sub RemoveOldData()
On Error GoTo perr
If rsICard.RecordCount > 0 Then rsICard.MoveFirst
    For i = 0 To rsICard.RecordCount - 1 Step 1
    If rsICard.BOF = False Or rsICard.EOF = False Then
        rsICard.Delete
        If rsICard.EOF = False Then rsICard.MoveNext
    End If
    Next
Exit Sub
perr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub AddPrintData()
rsICard.AddNew
If Trim(txtCode.Text) <> "" Then rsICard!Code = txtCode.Text
If Trim(txtName.Text) <> "" Then rsICard!Name = txtName.Text
If Trim(cmbDesig.Text) <> "" Then rsICard!Desig = cmbDesig.Text
If Trim(txtDOJ.Text) <> "" Then rsICard!DOJ = txtDOJ.Text
If Trim(txtDOR.Text) <> "" Then rsICard!DOR = txtDOR.Text
If Trim(txtAdd.Text) <> "" Then rsICard!address = txtAdd.Text
rsICard.Update
End Sub

Private Sub GenerateICard()
On Error GoTo errICard
Cr.Reset
Cr.ReportTitle = "EMPLOYEE IDENTITY CARD"
Cr.WindowState = crptMaximized
Cr.WindowTitle = "EMPLOYEE IDENTITY CARD"
Cr.WindowShowGroupTree = True
Cr.DataFiles(0) = App.Path & "\Project97.mdb"
Cr.ReportFileName = App.Path & "\Reports\Employee Details\rptICard.rpt"

Cr.Destination = crptToWindow
Cr.WindowShowRefreshBtn = True
Cr.WindowShowGroupTree = False

'CR.Action = 1
Exit Sub
errICard:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub
'This is to search for the entered employee code and
'entering its information in the text boxes.
'eg. Once we enter the Emp.Code and press enter key,
'all the textboxes should be filled automatically.
Private Sub cmdSearch_Click()
On Error GoTo serr
'Open the Inputbox and ask for the Employee Code from the user
Ans = InputBox("Enter the Employee Code to Search")
If rsEmp.RecordCount = 0 Then
    MsgBox "No Employee Records in the database !", vbOKOnly, "OASYS"
End If
If rsEmp.RecordCount > 0 Then rsEmp.MoveFirst
    For i = 0 To rsEmp.RecordCount - 1 Step 1
        If (Ans = rsEmp!Code) Then
        Call ClearAll
            If Not IsNull(rsEmp!Name) Then txtName.Text = rsEmp!Name
            If Not IsNull(rsEmp!Code) Then txtCode.Text = rsEmp!Code
            If Not IsNull(rsEmp!address) Then txtAdd.Text = rsEmp!address
            If Not IsNull(rsEmp!Desig) Then cmbDesig.Text = rsEmp!Desig
            If Not IsNull(rsEmp!FName) Then txtFName.Text = rsEmp!FName
            If Not IsNull(rsEmp!Phone) Then txtPhone.Text = rsEmp!Phone
            If Not IsNull(rsEmp!Caste) Then cmbCaste.Text = rsEmp!Caste
            If Not IsNull(rsEmp!Class) Then cmbClass.Text = rsEmp!Class
            If Not IsNull(rsEmp!Salary) Then txtSalary.Text = rsEmp!Salary
            If Not IsNull(rsEmp!BasicPay) Then txtBasicPay.Text = rsEmp!BasicPay
            If Not IsNull(rsEmp!Sex) Then cmbSex.Text = rsEmp!Sex
            If Not IsNull(rsEmp!Type) Then cmbType.Text = rsEmp!Type
            If Not IsNull(rsEmp!Sect) Then cmbSection.Text = rsEmp!Sect
            If Not IsNull(rsEmp!DOB) Then txtDOB.Text = rsEmp!DOB
            If Not IsNull(rsEmp!DOJ) Then txtDOJ.Text = rsEmp!DOJ
            If Not IsNull(rsEmp!DOR) Then txtDOR.Text = rsEmp!DOR
            If Not IsNull(rsEmp!Path) Then txtPath.Text = rsEmp!Path
            If Not IsNull(rsEmp!Qual) Then txtQual.Text = rsEmp!Qual

Call DisableAll  'Disable all the fields on form
cmdModify.Enabled = True
cmdRemove.Enabled = True
   
        If Not IsNull(rsEmp!Path) Then    ' For showing employee photo in the Image Box
            On Error GoTo FileNotFound
            Img.Picture = LoadPicture(rsEmp!Path) 'Load the employee picture into the Image Box
        ElseIf IsNull(rsEmp!Path) Then
            Img.Picture = LoadPicture(none) 'If there is no path, then keep the image box blank
        End If
            
            Exit Sub
        End If
    rsEmp.MoveNext
    If rsEmp.EOF = True Then
        MsgBox "Invalid Employee Code !", vbOKOnly, "OASYS"
        Exit Sub
    End If
    Next
Exit Sub
FileNotFound:
        MsgBox "Picture file not found ! Please set the correct Path !", vbOKOnly, "OASYS"
        Img.Picture = LoadPicture(none)
Exit Sub
serr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub Form_Load()
On Error GoTo ferr
    cmdSave.Enabled = False
    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    cmdCancel.Enabled = False
    
If rsEmp.RecordCount > 0 Then
    cmdAdd.Enabled = True
    cmdPrev.Enabled = True
    cmdNext.Enabled = True
End If

cmbSex.AddItem ("MALE")
cmbSex.AddItem ("FEMALE")
    
'Add Classes from the database into the Class combobox
If rsClass.RecordCount > 0 Then
rsClass.MoveFirst
    For i = 0 To rsClass.RecordCount - 1 Step 1
cmbClass.AddItem (rsClass!Code)
rsClass.MoveNext
    Next
End If
cmbClass.Text = cmbClass.List(0)

'Add Types from the database into the Type combobox
If rsType.RecordCount > 0 Then
    rsType.MoveFirst
    For i = 0 To rsType.RecordCount - 1 Step 1
    cmbType.AddItem (rsType!Code)
    rsType.MoveNext
    Next
    End If
cmbType.Text = cmbType.List(0)

'Add Sections from the database into the Section combobox
If rsSection.RecordCount > 0 Then
    rsSection.MoveFirst
    For i = 0 To rsSection.RecordCount - 1 Step 1
    cmbSection.AddItem (rsSection!Code)
    rsSection.MoveNext
    Next
    End If
cmbSection.Text = cmbSection.List(0)

'Add Castes from the database into the Caste combobox
If rsCaste.RecordCount > 0 Then
    rsCaste.MoveFirst
    For i = 0 To rsCaste.RecordCount - 1 Step 1
    cmbCaste.AddItem (rsCaste!Code)
    rsCaste.MoveNext
    Next
    End If
cmbCaste.Text = cmbCaste.List(0)

'Add Designations from the database into the Designation combobox
If rsDesig.RecordCount > 0 Then
    rsDesig.MoveFirst
    For i = 0 To rsDesig.RecordCount - 1 Step 1
    cmbDesig.AddItem (rsDesig!Code)
    rsDesig.MoveNext
    Next
    End If
cmbDesig.Text = cmbDesig.List(0)

Exit Sub
ferr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub cmdAdd_Click()
On Error GoTo aerr
Modify = False
Call Modi
Call EnableAll
Call ClearAll

'rsEmp.AddNew
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
End Sub

Private Sub cmdCancel_Click()
On Error GoTo cerr
Modify = False
Call Modi
'rsEmp.CancelUpdate
Call DisableAll

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
Call EnableAll

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

Private Sub cmdNext_Click()
On Error GoTo nerr
Modify = False
Call Modi

If rsEmp.RecordCount = 0 Then
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
If rsEmp.EOF = False Then rsEmp.MoveNext
If rsEmp.EOF = True Then rsEmp.MoveLast
showall

Exit Sub
nerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub cmdPrev_Click()
On Error GoTo perr
Modify = False
Call Modi

If rsEmp.RecordCount = 0 Then
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
If rsEmp.BOF = False Then rsEmp.MovePrevious
If rsEmp.BOF = True Then rsEmp.MoveFirst
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
 If rsEmp.RecordCount = 0 Then
    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    cmdCancel.Enabled = False
    cmdSave.Enabled = False
    Exit Sub
End If
       rsEmp.Delete
         Call cmdNext_Click
Exit Sub
rerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

'This checks if the entered Code already exists.
'As the Emp.Code should be unique, it will not allow
'any two employees with same Code.
Private Sub FindRecord()
If rsEmp.RecordCount > 0 Then rsEmp.MoveFirst
For i = 0 To rsEmp.RecordCount - 1 Step 1
    If rsEmp!Code = txtCode.Text Then
        MsgBox "Code Already Exists!", vbCritical, "OASYS"
        txtCode.Text = ""
        txtCode.SetFocus
        RecordFound = True
        Exit Sub
    End If
    If rsEmp.EOF = False Then rsEmp.MoveNext
If rsEmp.EOF = True Then
    RecordFound = False
End If
Next
End Sub

Private Sub cmdSave_Click()
On Error GoTo serr
If Trim(txtCode.Text) = "" Or Trim(txtName.Text) = "" Then
    Exit Sub
End If

If Modify = False Then
Call FindRecord
    If RecordFound = True Then  'If Code already exists then Exit Sub
        Exit Sub
    ElseIf RecordFound = False Then
        rsEmp.AddNew  'If Code does not exist then Add this record
    End If
End If

If Not txtCode.Text <> "" Then rsEmp!Code = txtCode.Text
If txtCode.Text <> "" Then rsEmp!Code = txtCode.Text
If txtName.Text <> "" Then rsEmp!Name = txtName.Text
If txtAdd.Text <> "" Then rsEmp!address = txtAdd.Text
If cmbDesig.Text <> "" Then rsEmp!Desig = cmbDesig.Text
If txtFName.Text <> "" Then rsEmp!FName = txtFName.Text
If txtPhone.Text <> "" Then rsEmp!Phone = txtPhone.Text
If cmbCaste.Text <> "" Then rsEmp!Caste = cmbCaste.Text
If cmbClass.Text <> "" Then rsEmp!Class = cmbClass.Text
If txtSalary.Text <> "" Then rsEmp!Salary = txtSalary.Text
If txtBasicPay.Text <> "" Then rsEmp!BasicPay = txtBasicPay.Text
If cmbSex.Text <> "" Then rsEmp!Sex = cmbSex.Text
If cmbType.Text <> "" Then rsEmp!Type = cmbType.Text
If cmbSection.Text <> "" Then rsEmp!Sect = cmbSection.Text
If txtDOB.Text <> "" Then rsEmp!DOB = txtDOB.Text
If txtDOJ.Text <> "" Then rsEmp!DOJ = txtDOJ.Text
If txtDOR.Text <> "" Then rsEmp!DOR = txtDOR.Text
If txtPath.Text <> "" Then rsEmp!Path = txtPath.Text
If txtQual.Text <> "" Then rsEmp!Qual = txtQual.Text

If Year(Date) < Year(rsExpiry!CurrentYear) Then
    MsgBox "You cannot work in previous working year ! Contact Developer", vbCritical, "OASYS"
Exit Sub
ElseIf Year(Date) > Year(rsExpiry!CurrentYear) Then
    MsgBox "You cannot work in next working year ! Contact Developer", vbCritical, "OASYS"
Exit Sub
End If

rsEmp.Update

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

'txtCode.Text = ""
'txtDesc.Text = ""
'txtCode.SetFocus

If cmdAdd.Enabled = True Then cmdAdd.SetFocus
Exit Sub
serr:
MsgBox Err.Description, vbOKOnly, "OASYS"
Call EnableAll
Call ClearAll
txtCode.SetFocus
End Sub

Private Sub Form_Activate()
cmdAdd.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub txtAdd_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtAdd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtAdd_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtBasicPay_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtBasicPay_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtBasicPay_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub
Private Sub txtCode_GotFocus()
Call txt_GotFocus
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

Private Sub txtDOB_KeyPress(KeyAscii As Integer)
On Error GoTo derr
Select Case KeyAscii
    
    Case vbKeyBack, 48 To 57, vbKeyReturn, 47, vbKeyEscape
    Case Else
         MsgBox "No Special Characters are allowed! Please enter Numbers and / only! ", vbOKOnly, "OASYS"
         KeyAscii = 0
         txtDOB.SetFocus
         Exit Sub
    End Select

If KeyAscii = 13 Then
   datevali (txtDOB.Text)
End If
Exit Sub
derr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub txtDOJ_KeyPress(KeyAscii As Integer)
On Error GoTo derr
Select Case KeyAscii
    
    Case vbKeyBack, 48 To 57, vbKeyReturn, 47, vbKeyEscape
    Case Else
         MsgBox "No Special Characters are allowed! Please enter Numbers and / only! ", vbOKOnly, "OASYS"
         KeyAscii = 0
         txtDOJ.SetFocus
        
         Exit Sub
    End Select

If KeyAscii = 13 Then
   datevali (txtDOJ.Text)
End If
Exit Sub
derr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub txtDOR_KeyPress(KeyAscii As Integer)
On Error GoTo derr
Select Case KeyAscii
    
    Case vbKeyBack, 48 To 57, vbKeyReturn, 47, vbKeyEscape
    Case Else
         MsgBox "No Special Characters are allowed! Please enter Numbers and / only! ", vbOKOnly, "OASYS"
         KeyAscii = 0
         txtDOR.SetFocus
        
         Exit Sub
    End Select

If KeyAscii = 13 Then
   datevali (txtDOR.Text)
End If
Exit Sub
derr:
MsgBox Err.Description, vbOKOnly, "OASYS"

End Sub

Private Sub txtFname_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtFname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtFname_KeyUp(KeyCode As Integer, Shift As Integer)
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

Private Sub cmbDesig_GotFocus()
Call txt_GotFocus
End Sub

Private Sub cmbDesig_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub cmbDesig_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub
Private Sub showall()
On Error GoTo showerr
Call ClearAll
If rsEmp.RecordCount = 0 Then
    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    Exit Sub
ElseIf rsEmp.RecordCount > 0 Then
    cmdModify.Enabled = True
    cmdPrev.Enabled = True
    cmdNext.Enabled = True
    cmdRemove.Enabled = True
End If

If Not IsNull(rsEmp!Name) Then txtName.Text = rsEmp!Name
If Not IsNull(rsEmp!Code) Then txtCode.Text = rsEmp!Code
If Not IsNull(rsEmp!address) Then txtAdd.Text = rsEmp!address
If Not IsNull(rsEmp!Desig) Then cmbDesig.Text = rsEmp!Desig
If Not IsNull(rsEmp!FName) Then txtFName.Text = rsEmp!FName
If Not IsNull(rsEmp!Phone) Then txtPhone.Text = rsEmp!Phone
If Not IsNull(rsEmp!Caste) Then cmbCaste.Text = rsEmp!Caste
If Not IsNull(rsEmp!Class) Then cmbClass.Text = rsEmp!Class
If Not IsNull(rsEmp!Salary) Then txtSalary.Text = rsEmp!Salary
If Not IsNull(rsEmp!BasicPay) Then txtBasicPay.Text = rsEmp!BasicPay
If Not IsNull(rsEmp!Sex) Then cmbSex.Text = rsEmp!Sex
If Not IsNull(rsEmp!Type) Then cmbType.Text = rsEmp!Type
If Not IsNull(rsEmp!Sect) Then cmbSection.Text = rsEmp!Sect
If Not IsNull(rsEmp!DOB) Then txtDOB.Text = rsEmp!DOB
If Not IsNull(rsEmp!DOJ) Then txtDOJ.Text = rsEmp!DOJ
If Not IsNull(rsEmp!DOR) Then txtDOR.Text = rsEmp!DOR
If Not IsNull(rsEmp!Path) Then txtPath.Text = rsEmp!Path
If Not IsNull(rsEmp!Qual) Then txtQual.Text = rsEmp!Qual

If Not IsNull(rsEmp!Path) Then    ' For showing employee photo in the Image Box
On Error GoTo FileNotFound
    Img.Picture = LoadPicture(rsEmp!Path)
ElseIf IsNull(rsEmp!Path) Then
    Img.Picture = LoadPicture(none)
End If

Exit Sub
FileNotFound:
    MsgBox "Picture file not found ! Please set the correct Path !", vbOKOnly, "OASYS"
    Img.Picture = LoadPicture(none)

Exit Sub
showerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub


Private Sub txt_GotFocus()
On Error GoTo focerr
    If Trim(txtName.Text) = "" Or Trim(txtCode.Text) = "" Then
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
    Call DisableAll
ElseIf Modify = True Then  ' If cmdmodify is not clicked then
    Call EnableAll
End If
Exit Sub
merr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub DisableAll()
txtCode.Enabled = False
txtName.Enabled = False
txtAdd.Enabled = False
cmbDesig.Enabled = False
txtFName.Enabled = False
txtPhone.Enabled = False
cmbCaste.Enabled = False
cmbClass.Enabled = False
txtSalary.Enabled = False
txtBasicPay.Enabled = False
cmbSex.Enabled = False
txtPath.Enabled = False
cmbType.Enabled = False
cmbSection.Enabled = False
txtDOB.Enabled = False
txtDOJ.Enabled = False
txtDOR.Enabled = False
txtQual.Enabled = False
End Sub

Private Sub EnableAll()
txtCode.Enabled = True
txtName.Enabled = True
cmbDesig.Enabled = True
txtFName.Enabled = True
txtPhone.Enabled = True
cmbCaste.Enabled = True
cmbClass.Enabled = True
txtSalary.Enabled = True
txtBasicPay.Enabled = True
cmbSex.Enabled = True
cmbType.Enabled = True
cmbSection.Enabled = True
txtDOB.Enabled = True
txtDOJ.Enabled = True
txtDOR.Enabled = True
txtQual.Enabled = True
txtPath.Enabled = True
txtAdd.Enabled = True
End Sub

Private Sub ClearAll()
txtCode.Text = ""
txtName.Text = ""
txtAdd.Text = ""
cmbDesig.Text = ""
txtFName.Text = ""
txtPhone.Text = ""
cmbCaste.Text = ""
cmbClass.Text = ""
txtSalary.Text = ""
txtBasicPay.Text = ""
cmbSex.Text = ""
cmbType.Text = ""
cmbSection.Text = ""
txtAdd.Text = ""
txtDOB.Text = ""
txtDOJ.Text = ""
txtDOR.Text = ""
txtQual.Text = ""
txtPath.Text = ""
Img.Picture = LoadPicture(none)
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

Private Sub txtPath_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtPath_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtPath_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtPath_LostFocus()
If txtPath.Text <> "" Then
    Img.Picture = LoadPicture(txtPath.Text) ' Loads the picture from the entered path into the Image Box
ElseIf txtPath.Text = "" Then
    Img.Picture = LoadPicture(none)
End If
End Sub

Private Sub txtPhone_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtPhone_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtPhone_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtQual_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtQual_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtQual_KeyUp(KeyCode As Integer, Shift As Integer)
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
