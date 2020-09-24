VERSION 5.00
Begin VB.Form frmMain1 
   BackColor       =   &H00000000&
   Caption         =   "Main"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox CR 
      Height          =   480
      Left            =   5220
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   13
      Top             =   3780
      Width           =   1200
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Pay Slip"
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
      Height          =   240
      Left            =   8100
      TabIndex        =   12
      Top             =   5175
      Width           =   915
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " &Deduction/Cash"
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
      Height          =   240
      Left            =   7770
      TabIndex        =   11
      Top             =   2880
      Width           =   1755
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   5535
      TabIndex        =   10
      Top             =   7515
      Width           =   465
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&About"
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
      Height          =   240
      Left            =   8235
      TabIndex        =   5
      Top             =   495
      Width           =   645
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Leave Slip"
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
      Height          =   240
      Left            =   7890
      TabIndex        =   7
      Top             =   4005
      Width           =   1155
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Attendance"
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
      Height          =   240
      Left            =   7875
      TabIndex        =   8
      Top             =   1710
      Width           =   1215
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee"
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
      Height          =   240
      Left            =   2385
      TabIndex        =   9
      Top             =   6300
      Width           =   1095
   End
   Begin VB.Image ImgEmpB 
      Height          =   975
      Left            =   1170
      Picture         =   "frmMain1.frx":0000
      Top             =   6030
      Width           =   3570
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Reports"
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
      Height          =   240
      Left            =   8085
      TabIndex        =   6
      Top             =   6345
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Working Days"
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
      Height          =   240
      Left            =   2145
      TabIndex        =   4
      Top             =   5175
      Width           =   1485
   End
   Begin VB.Image ImgWDB 
      Height          =   975
      Left            =   1170
      Picture         =   "frmMain1.frx":0586
      Top             =   4860
      Width           =   3570
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Section"
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
      Height          =   240
      Left            =   2520
      TabIndex        =   3
      Top             =   4005
      Width           =   825
   End
   Begin VB.Image ImgSectionB 
      Height          =   975
      Left            =   1260
      Picture         =   "frmMain1.frx":0B0C
      Top             =   3690
      Width           =   3570
   End
   Begin VB.Label lblDesig 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Designation"
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
      Height          =   240
      Left            =   2295
      TabIndex        =   2
      Top             =   2880
      Width           =   1275
   End
   Begin VB.Image ImgDesigB 
      Height          =   975
      Left            =   1260
      Picture         =   "frmMain1.frx":1092
      Top             =   2565
      Width           =   3570
   End
   Begin VB.Label lblCaste 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caste"
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
      Height          =   240
      Left            =   2625
      TabIndex        =   1
      Top             =   1710
      Width           =   645
   End
   Begin VB.Image ImgCasteB 
      Height          =   975
      Left            =   1260
      Picture         =   "frmMain1.frx":1618
      Top             =   1395
      Width           =   3570
   End
   Begin VB.Image ImgCasteR 
      Height          =   975
      Left            =   1260
      Picture         =   "frmMain1.frx":1B9E
      Top             =   1395
      Visible         =   0   'False
      Width           =   3570
   End
   Begin VB.Label lblType 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
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
      Height          =   240
      Left            =   2685
      TabIndex        =   0
      Top             =   495
      Width           =   585
   End
   Begin VB.Image ImgTypeB 
      Height          =   975
      Left            =   1215
      Picture         =   "frmMain1.frx":2217
      Top             =   180
      Width           =   3570
   End
   Begin VB.Image ImgTypeR 
      Height          =   975
      Left            =   1260
      Picture         =   "frmMain1.frx":279D
      Top             =   180
      Visible         =   0   'False
      Width           =   3570
   End
   Begin VB.Image ImgDesigR 
      Height          =   975
      Left            =   1260
      Picture         =   "frmMain1.frx":2E16
      Top             =   2565
      Visible         =   0   'False
      Width           =   3570
   End
   Begin VB.Image ImgSectionR 
      Height          =   975
      Left            =   1260
      Picture         =   "frmMain1.frx":348F
      Top             =   3690
      Visible         =   0   'False
      Width           =   3570
   End
   Begin VB.Image ImgWDR 
      Height          =   975
      Left            =   1170
      Picture         =   "frmMain1.frx":3B08
      Top             =   4860
      Visible         =   0   'False
      Width           =   3570
   End
   Begin VB.Image ImgEmpR 
      Height          =   975
      Left            =   1215
      Picture         =   "frmMain1.frx":4181
      Top             =   6030
      Visible         =   0   'False
      Width           =   3570
   End
   Begin VB.Image ImgAttnB 
      Height          =   975
      Left            =   6795
      Picture         =   "frmMain1.frx":47FA
      Top             =   1395
      Width           =   3570
   End
   Begin VB.Image ImgAttnR 
      Height          =   975
      Left            =   6795
      Picture         =   "frmMain1.frx":4D80
      Top             =   1395
      Visible         =   0   'False
      Width           =   3570
   End
   Begin VB.Image ImgLSB 
      Height          =   975
      Left            =   6795
      Picture         =   "frmMain1.frx":53F9
      Top             =   3690
      Width           =   3570
   End
   Begin VB.Image ImgLSR 
      Height          =   975
      Left            =   6750
      Picture         =   "frmMain1.frx":597F
      Top             =   3690
      Visible         =   0   'False
      Width           =   3570
   End
   Begin VB.Image ImgRepB 
      Height          =   975
      Left            =   6795
      Picture         =   "frmMain1.frx":5FF8
      Top             =   6030
      Width           =   3570
   End
   Begin VB.Image ImgRepR 
      Height          =   975
      Left            =   6795
      Picture         =   "frmMain1.frx":657E
      Top             =   6030
      Visible         =   0   'False
      Width           =   3570
   End
   Begin VB.Image ImgAbtB 
      Height          =   975
      Left            =   6840
      Picture         =   "frmMain1.frx":6BF7
      Top             =   180
      Width           =   3570
   End
   Begin VB.Image ImgAbtR 
      Height          =   975
      Left            =   6840
      Picture         =   "frmMain1.frx":717D
      Top             =   180
      Visible         =   0   'False
      Width           =   3570
   End
   Begin VB.Image ImgExitR 
      Height          =   975
      Left            =   4050
      Picture         =   "frmMain1.frx":77F6
      Top             =   7200
      Width           =   3570
   End
   Begin VB.Image ImgExitB 
      Height          =   975
      Left            =   4050
      Picture         =   "frmMain1.frx":7E6F
      Top             =   7200
      Width           =   3570
   End
   Begin VB.Image ImgPSB 
      Height          =   975
      Left            =   6840
      Picture         =   "frmMain1.frx":83F5
      Top             =   4860
      Width           =   3570
   End
   Begin VB.Image ImgPSR 
      Height          =   975
      Left            =   6840
      Picture         =   "frmMain1.frx":897B
      Top             =   4860
      Visible         =   0   'False
      Width           =   3570
   End
   Begin VB.Image imgPDCB 
      Height          =   975
      Left            =   6840
      Picture         =   "frmMain1.frx":8FF4
      Top             =   2565
      Width           =   3570
   End
   Begin VB.Image imgPDCR 
      Height          =   975
      Left            =   6840
      Picture         =   "frmMain1.frx":957A
      Top             =   2565
      Visible         =   0   'False
      Width           =   3570
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&FILE"
      Begin VB.Menu mnuMax 
         Caption         =   "Ma&ximimize"
      End
      Begin VB.Menu mnuMin 
         Caption         =   "Mi&nimize"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuMaster 
      Caption         =   "&MASTERS"
      Begin VB.Menu mnuCaste 
         Caption         =   "&Caste"
      End
      Begin VB.Menu mnuClass 
         Caption         =   "C&lass"
      End
      Begin VB.Menu mnuDesig 
         Caption         =   "&Designation"
      End
      Begin VB.Menu mnuSection 
         Caption         =   "&Section"
      End
      Begin VB.Menu mnuType 
         Caption         =   "&Type"
      End
      Begin VB.Menu mnuWorkingD 
         Caption         =   "&Working Days"
      End
   End
   Begin VB.Menu mnuDataEntry 
      Caption         =   "&DATA ENTRY FORMS"
      Begin VB.Menu mnuAttn 
         Caption         =   "&Attendance"
      End
      Begin VB.Menu mnuEmp 
         Caption         =   "&Employee"
      End
      Begin VB.Menu mnuLS 
         Caption         =   "&Leave Slip"
      End
      Begin VB.Menu mnuPayDed 
         Caption         =   "&Pay Deduction/Cash"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&REPORTS"
      Begin VB.Menu mnuList 
         Caption         =   "L&IST"
         Begin VB.Menu mnuCasteList 
            Caption         =   "&CASTE LIST"
         End
         Begin VB.Menu mnuClassList 
            Caption         =   "C&LASS LIST"
         End
         Begin VB.Menu mnuDesigList 
            Caption         =   "&DESIGNATION LIST"
         End
         Begin VB.Menu mnuSectionList 
            Caption         =   "&SECTION LIST"
         End
         Begin VB.Menu mnuTypeList 
            Caption         =   "&TYPE LIST"
         End
      End
      Begin VB.Menu mnuLeaveReport 
         Caption         =   "&LEAVE REPORT"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&ABOUT"
      Begin VB.Menu mnuDeveloper 
         Caption         =   "&Developer"
      End
      Begin VB.Menu mnuDetails 
         Caption         =   "&Project Details"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&HELP"
      End
   End
End
Attribute VB_Name = "frmMain1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
ans = MsgBox("Do you want to Exit ?", vbYesNo, "Office Automation")

If ans = vbYes Then
    End
Else
    Exit Sub
End If
End Sub

Private Sub cmdMastAttn_Click()
Load frmAttn
frmAttn.Show
End Sub

Private Sub cmdMastDesig_Click()
Load frmMastDesig
frmMastDesig.Show
End Sub

Private Sub cmdMastEmp_Click()
Load frmMastEmp
frmMastEmp.Show
End Sub

Private Sub cmdMastLS_Click()
Load frmLeaveSlip
frmLeaveSlip.Show
End Sub

Private Sub cmdMastSection_Click()
Load frmMastSection
frmMastSection.Show
End Sub

Private Sub cmdMastWD_Click()
Load frmMastWD
frmMastWD.Show
End Sub

Private Sub CmdMastType_Click()
Load frmMastType
frmMastType.Show
End Sub

Private Sub cmd_about_Click()
Load frmAbout
frmAbout.Show
End Sub

Private Sub Form_Load()
Call LoadButtons 'This function is to load images in the image box, to make it look like command buttons
If rsEmp.RecordCount > 0 Then rsEmp.MoveFirst
'loop for setting yes/no
For i = 0 To rsEmp.RecordCount - 1 Step 1
    If rsEmp!DOR <= Date Then
        rsEmp!Retired = "Y"
        rsEmp.Update
    Else
        rsEmp!Retired = "N"
    End If
    If rsEmp.EOF = False Then rsEmp.MoveNext
Next
'loop for putting the data in the retired table
If rsEmp.RecordCount > 0 Then rsEmp.MoveFirst

    For i = 0 To rsEmp.RecordCount - 1 Step 1
        If rsEmp!Retired = "Y" Or rsEmp!Retired = "y" Then
            rsRetired.AddNew
                rsRetired!Code = rsEmp!Code
                rsRetired!Name = rsEmp!Name
                rsRetired!Desig = rsEmp!Desig
                rsRetired!Sect = rsEmp!Sect
                rsRetired!FName = rsEmp!FName
                rsRetired!Address = rsEmp!Address
                rsRetired!Caste = rsEmp!Caste
                rsRetired!Class = rsEmp!Class
                rsRetired!Sex = rsEmp!Sex
                rsRetired!Type = rsEmp!Type
                rsRetired!DOB = rsEmp!DOB
                rsRetired!DOJ = rsEmp!DOJ
                rsRetired!DOR = rsEmp!DOR
                rsRetired!Qual = rsEmp!Qual
                rsRetired!Salary = rsEmp!Salary
                rsRetired!Phone = rsEmp!Phone
                rsRetired!Path = rsEmp!Path
                rsRetired!Retired = rsEmp!Retired
            rsRetired.Update  'update in the retired table and delete from employee table
                        rsEmp.Delete
        End If
    If rsEmp.EOF = False Then rsEmp.MoveNext
    Next
End Sub

Private Sub ImgTypeB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgTypeB_Down
End Sub
Private Sub ImgTypeB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgTypeB_Up
End Sub
Private Sub ImgTypeR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgTypeR_Up
End Sub

Private Sub ImgCasteB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgCasteB_Down
End Sub

Private Sub ImgCasteB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgCasteB_Up
End Sub

Private Sub ImgCasteR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgCasteR_Up
End Sub

Private Sub ImgDesigB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgDesigB_Down
End Sub
Private Sub ImgDesigB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgDesigB_Up
End Sub
Private Sub ImgDesigR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgDesigR_Up
End Sub

Private Sub ImgSectionB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgSectionB.Visible = False
ImgSectionR.Visible = True
End Sub
Private Sub ImgSectionB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Load frmMastSection
frmMastSection.Show
frmMain.Hide
ImgSectionB.Visible = True
ImgSectionR.Visible = False
End Sub
Private Sub ImgSectionR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgSectionB.Visible = True
ImgSectionR.Visible = False
End Sub

Private Sub ImgWDB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgWDB.Visible = False
ImgWDR.Visible = True
End Sub
Private Sub ImgWDB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Load frmMastWD
frmMastWD.Show
frmMain.Hide
ImgWDB.Visible = True
ImgWDR.Visible = False
End Sub
Private Sub ImgWDR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgWDB.Visible = True
ImgWDR.Visible = False
End Sub

Private Sub ImgEmpB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgEmpB.Visible = False
ImgEmpR.Visible = True
End Sub
Private Sub ImgEmpB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Load frmMastEmp
frmMastEmp.Show
frmMain.Hide
ImgEmpB.Visible = True
ImgEmpR.Visible = False
End Sub
Private Sub ImgEmpR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgEmpB.Visible = True
ImgEmpR.Visible = False
End Sub

Private Sub ImgLSB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgLSB.Visible = False
ImgLSR.Visible = True
End Sub
Private Sub ImgLSB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Load frmLeaveSlip
frmLeaveSlip.Show
frmMain.Hide
ImgLSB.Visible = True
ImgLSR.Visible = False
End Sub
Private Sub ImgLSR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgLSB.Visible = True
ImgLSR.Visible = False
End Sub

Private Sub ImgRepB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgRepB.Visible = False
ImgRepR.Visible = True
End Sub
Private Sub ImgRepB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Load frmReports
frmReports.Show
frmMain.Hide
ImgRepB.Visible = True
ImgRepR.Visible = False
End Sub
Private Sub ImgRepR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgRepB.Visible = True
ImgRepR.Visible = False
End Sub

Private Sub ImgAbtB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgAbtB.Visible = False
ImgAbtR.Visible = True
End Sub
Private Sub ImgAbtB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Load frmAbout
frmAbout.Show
frmMain.Hide
ImgAbtB.Visible = True
ImgAbtR.Visible = False
End Sub
Private Sub ImgAbtR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgAbtB.Visible = True
ImgAbtR.Visible = False
End Sub

Private Sub ImgExitR_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgExitR.Visible = False
ImgExitB.Visible = True
End Sub
Private Sub ImgExitR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgExitR.Visible = True
ImgExitB.Visible = False

ans = MsgBox("Do you want to Exit ?", vbYesNo, "Office Automation")
If ans = vbYes Then
    End
Else
    Exit Sub
End If
End Sub
Private Sub ImgExitB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgExitR.Visible = True
ImgExitB.Visible = False
End Sub

Private Sub ImgPSB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgPSB.Visible = False
ImgPSR.Visible = True
End Sub
Private Sub ImgPSB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Load frmPay
frmPay.Show
frmMain.Hide
ImgPSB.Visible = True
ImgPSR.Visible = False
End Sub
Private Sub ImgPSR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgPSB.Visible = True
ImgPSR.Visible = False
End Sub

Private Sub ImgAttnB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgAttnB.Visible = False
ImgAttnR.Visible = True
End Sub
Private Sub ImgAttnB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Load frmAttn
frmAttn.Show
frmMain.Hide
ImgAttnB.Visible = True
ImgAttnR.Visible = False
End Sub
Private Sub ImgAttnR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgAttnB.Visible = True
ImgAttnR.Visible = False
End Sub

Private Sub ImgPDCB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgPDCB.Visible = False
imgPDCR.Visible = True
End Sub
Private Sub ImgPDCB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Load frmPayDeduct
frmPayDeduct.Show
frmMain.Hide
imgPDCB.Visible = True
imgPDCR.Visible = False
End Sub
Private Sub ImgpdcR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgPDCB.Visible = True
imgPDCR.Visible = False
End Sub


Private Sub mnuAttn_Click()
Load frmAttn
frmAttn.Show
End Sub

Private Sub mnuCaste_Click()
Load frmCaste
frmCaste.Show
End Sub

Private Sub mnuCasteList_Click()
On Error GoTo errCasteList
CR.Reset
CR.ReportTitle = "CASTE LIST"
CR.WindowState = crptMaximized
CR.WindowTitle = "CASTE LIST"
CR.WindowShowGroupTree = True ' side tree structure in the report
'standard format for connecting with the database
CR.DataFiles(0) = App.Path & "\Project.mdb"
CR.ReportFileName = App.Path & "\Reports\rptCasteList.rpt"


CR.Destination = crptToWindow 'the position where the information is to be displayed
CR.WindowShowRefreshBtn = True ' to specify whether or not to enable the refresh property
CR.Action = 0
Exit Sub
errCasteList:
MsgBox Err.Description, vbOKOnly, "Office Automation"
End Sub

Private Sub mnuClass_Click()
Load frmMastClass
frmMastClass.Show
End Sub

Private Sub mnuClassList_Click()
On Error GoTo errClassList
CR.Reset
CR.ReportTitle = "CLASS LIST"
CR.WindowState = crptMaximized
CR.WindowTitle = "CLASS LIST"
CR.WindowShowGroupTree = True

CR.DataFiles(0) = App.Path & "\Project.mdb"
CR.ReportFileName = App.Path & "\Reports\rptClassList.rpt"


CR.Destination = crptToWindow
CR.WindowShowRefreshBtn = True
CR.Action = 1
Exit Sub
errClassList:
MsgBox Err.Description, vbOKOnly, "Office Automation"
End Sub

Private Sub mnuDesig_Click()
Load frmMastDesig
frmMastDesig.Show
End Sub

Private Sub mnuDesigList_Click()
On Error GoTo errDesigList
CR.Reset
CR.ReportTitle = "DESIGNATION LIST"
CR.WindowState = crptMaximized
CR.WindowTitle = "DESIGNATION LIST"
CR.WindowShowGroupTree = True

CR.DataFiles(0) = App.Path & "\Project.mdb"
CR.ReportFileName = App.Path & "\Reports\rptDesignationList.rpt"

CR.Destination = crptToWindow
CR.WindowShowRefreshBtn = True
CR.Action = 1 'to make the report appear compulsory otherwise the report is not enabled
Exit Sub
errDesigList:
MsgBox Err.Description, vbOKOnly, "Office Automation"
End Sub

Private Sub mnuDeveloper_Click()
Load frmAbout
frmAbout.Show
End Sub

Private Sub mnuEmp_Click()
Load frmMastEmp
frmMastEmp.Show
End Sub

Private Sub mnuExit_Click()
ans = MsgBox("Do You want to Exit ?", vbYesNo, "Office Automation")
If ans = vbYes Then
    End
Else
    Exit Sub
End If
End Sub

Private Sub mnuLeave_Click()
Load frmLeaveSlip
frmLeaveSlip.Show
End Sub

Private Sub mnuLeaveReport_Click()
On Error GoTo errLeaveReport
CR.Reset
CR.ReportTitle = "LEAVE REPORT"
CR.WindowState = crptMaximized
CR.WindowTitle = "LEAVE REPORT"
CR.WindowShowGroupTree = True

CR.DataFiles(0) = App.Path & "\Project.mdb"
CR.ReportFileName = App.Path & "\Reports\rptLeaveReport.rpt"


CR.Destination = crptToWindow
CR.WindowShowRefreshBtn = True
CR.Action = 1
Exit Sub
errLeaveReport:
MsgBox Err.Description, vbOKOnly, "Office Automation"

End Sub

Private Sub mnuMax_Click()
frmMain.WindowState = 2
End Sub

Private Sub mnuMin_Click()
frmMain.WindowState = 1
End Sub

Private Sub mnuPayDed_Click()
Load frmPayDeduct
frmPayDeduct.Show
End Sub

Private Sub mnuRestore_Click()
frmMain.WindowState = 0
End Sub

Private Sub mnuSection_Click()
Load frmMastSection
frmMastSection.Show
End Sub

Private Sub mnuSectionList_Click()
On Error GoTo errSectionList
CR.Reset
CR.ReportTitle = "SECTION LIST"
CR.WindowState = crptMaximized
CR.WindowTitle = "SECTION LIST"
CR.WindowShowGroupTree = True

CR.DataFiles(0) = App.Path & "\Project.mdb"
CR.ReportFileName = App.Path & "\Reports\rptSectionList.rpt"

CR.Destination = crptToWindow
CR.WindowShowRefreshBtn = True
CR.Action = 1
Exit Sub
errSectionList:
MsgBox Err.Description, vbOKOnly, "Office Automation"
End Sub

Private Sub mnuType_Click()
Load frmMastType
frmMastType.Show
End Sub

Private Sub mnuTypeList_Click()
On Error GoTo errTypeList
CR.Reset
CR.ReportTitle = "TYPE LIST"
CR.WindowState = crptMaximized
CR.WindowTitle = "TYPE LIST"
CR.WindowShowGroupTree = True

CR.DataFiles(0) = App.Path & "\Project.mdb"
CR.ReportFileName = App.Path & "\Reports\rptTypeList.rpt"


CR.Destination = crptToWindow
CR.WindowShowRefreshBtn = True
CR.Action = 1
Exit Sub
errTypeList:
MsgBox Err.Description, vbOKOnly, "Office Automation"
End Sub

Private Sub mnuWorkingD_Click()
Load frmMastWD
frmMastWD.Show
End Sub

Private Sub LoadButtons()
'This Functions loads the Button images in the ImageBox to
'make it look like command buttons
'Blue.jpg and Red.jpg should be in the folder of your .exe file

ImgTypeB.Picture = LoadPicture(App.Path & "\Blue.jpg")
ImgTypeR.Picture = LoadPicture(App.Path & "\Red.jpg")

ImgCasteB.Picture = LoadPicture(App.Path & "\Blue.jpg")
ImgCasteR.Picture = LoadPicture(App.Path & "\Red.jpg")

ImgDesigB.Picture = LoadPicture(App.Path & "\Blue.jpg")
ImgDesigR.Picture = LoadPicture(App.Path & "\Red.jpg")

ImgSectionB.Picture = LoadPicture(App.Path & "\Blue.jpg")
ImgSectionR.Picture = LoadPicture(App.Path & "\Red.jpg")

ImgWDB.Picture = LoadPicture(App.Path & "\Blue.jpg")
ImgWDR.Picture = LoadPicture(App.Path & "\Red.jpg")

ImgEmpB.Picture = LoadPicture(App.Path & "\Blue.jpg")
ImgEmpR.Picture = LoadPicture(App.Path & "\Red.jpg")

ImgAbtB.Picture = LoadPicture(App.Path & "\Blue.jpg")
ImgAbtR.Picture = LoadPicture(App.Path & "\Red.jpg")

ImgAttnB.Picture = LoadPicture(App.Path & "\Blue.jpg")
ImgAttnR.Picture = LoadPicture(App.Path & "\Red.jpg")

imgPDCB.Picture = LoadPicture(App.Path & "\Blue.jpg")
imgPDCR.Picture = LoadPicture(App.Path & "\Red.jpg")

ImgLSB.Picture = LoadPicture(App.Path & "\Blue.jpg")
ImgLSR.Picture = LoadPicture(App.Path & "\Red.jpg")

ImgPSB.Picture = LoadPicture(App.Path & "\Blue.jpg")
ImgPSR.Picture = LoadPicture(App.Path & "\Red.jpg")

ImgRepB.Picture = LoadPicture(App.Path & "\Blue.jpg")
ImgRepR.Picture = LoadPicture(App.Path & "\Red.jpg")

ImgExitB.Picture = LoadPicture(App.Path & "\Blue.jpg")
ImgExitR.Picture = LoadPicture(App.Path & "\Red.jpg")

End Sub


'Functions for making Image boxes work as command button
'are declared over here
'======================================================
Private Sub ImgTypeB_Down()
ImgTypeB.Visible = False
ImgTypeR.Visible = True
End Sub
Private Sub ImgTypeB_Up()
Load frmMastType
frmMastType.Show
frmMain.Hide
ImgTypeB.Visible = True
ImgTypeR.Visible = False
End Sub
Private Sub ImgTypeR_Up()
ImgTypeB.Visible = True
ImgTypeR.Visible = False
End Sub

Private Sub ImgCasteB_Down()
ImgCasteB.Visible = False
ImgCasteR.Visible = True
End Sub
Private Sub ImgCasteB_Up()
Load frmCaste
frmCaste.Show
frmMain.Hide
ImgCasteB.Visible = True
ImgCasteR.Visible = False
End Sub
Private Sub ImgCasteR_Up()
ImgCasteB.Visible = True
ImgCasteR.Visible = False
End Sub

Private Sub ImgDesigB_Down()
ImgDesigB.Visible = False
ImgDesigR.Visible = True
End Sub
Private Sub ImgDesigB_Up()
Load frmMastDesig
frmMastDesig.Show
frmMain.Hide
ImgDesigB.Visible = True
ImgDesigR.Visible = False
End Sub
Private Sub ImgDesigR_Up()
ImgDesigB.Visible = True
ImgDesigR.Visible = False
End Sub
'On clicking on the Labels of the imagebox, it should work
'just as command buttons.
'=====================================================
Private Sub lblCaste_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgCasteB_Down
End Sub

Private Sub lblCaste_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgCasteB_Up
End Sub

Private Sub lblType_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgTypeB_Down
End Sub

Private Sub lblType_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgTypeB_Up
End Sub

Private Sub lblDesig_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgDesigB_Down
End Sub

Private Sub lblDesig_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgDesigB_Up
End Sub
