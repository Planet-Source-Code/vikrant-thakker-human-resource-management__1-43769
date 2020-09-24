VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H008080FF&
   Caption         =   "Main"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CR 
      Left            =   3720
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HUMAN RESOURCE MANAGEMENT"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009BF4C8&
      Height          =   555
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   11895
   End
   Begin VB.Label lblHelp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
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
      Left            =   1635
      TabIndex        =   13
      Top             =   7455
      Width           =   525
   End
   Begin VB.Image ImgHelpB 
      Height          =   975
      Left            =   180
      Top             =   7140
      Width           =   3570
   End
   Begin VB.Image ImgHelpR 
      Height          =   975
      Left            =   180
      Top             =   7140
      Visible         =   0   'False
      Width           =   3570
   End
   Begin VB.Label lblPS 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pay Slip"
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
      Left            =   9300
      TabIndex        =   12
      Top             =   4395
      Width           =   915
   End
   Begin VB.Label lblPD 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Deduction/Cash"
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
      Left            =   8940
      TabIndex        =   11
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
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
      Left            =   5655
      TabIndex        =   10
      Top             =   7455
      Width           =   465
   End
   Begin VB.Label lblAbt 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "About"
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
      Left            =   9435
      TabIndex        =   5
      Top             =   7455
      Width           =   645
   End
   Begin VB.Label lblLS 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Leave Slip"
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
      Left            =   9240
      TabIndex        =   7
      Top             =   3180
      Width           =   1155
   End
   Begin VB.Label lblAttn 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Attendance"
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
      Left            =   9195
      TabIndex        =   8
      Top             =   870
      Width           =   1215
   End
   Begin VB.Label lblEmp 
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
      Left            =   9270
      TabIndex        =   9
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Image ImgEmpB 
      Height          =   975
      Left            =   8010
      Top             =   5280
      Width           =   3570
   End
   Begin VB.Label lblRep 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reports"
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
      Left            =   5385
      TabIndex        =   6
      Top             =   5685
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblWD 
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
      Left            =   1245
      TabIndex        =   4
      Top             =   5775
      Width           =   1485
   End
   Begin VB.Image ImgWDB 
      Height          =   975
      Left            =   180
      Top             =   5400
      Width           =   3570
   End
   Begin VB.Label lblSection 
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
      Left            =   1500
      TabIndex        =   3
      Top             =   4545
      Width           =   825
   End
   Begin VB.Image ImgSectionB 
      Height          =   975
      Left            =   180
      Top             =   4170
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
      Left            =   1335
      TabIndex        =   2
      Top             =   3300
      Width           =   1275
   End
   Begin VB.Image ImgDesigB 
      Height          =   975
      Left            =   180
      Top             =   2925
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
      Left            =   1605
      TabIndex        =   1
      Top             =   2130
      Width           =   645
   End
   Begin VB.Image ImgCasteB 
      Height          =   975
      Left            =   180
      Top             =   1755
      Width           =   3570
   End
   Begin VB.Image ImgCasteR 
      Height          =   975
      Left            =   180
      Top             =   1755
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
      Left            =   1605
      TabIndex        =   0
      Top             =   855
      Width           =   585
   End
   Begin VB.Image ImgTypeB 
      Height          =   975
      Left            =   195
      Top             =   540
      Width           =   3570
   End
   Begin VB.Image ImgTypeR 
      Height          =   975
      Left            =   180
      Top             =   540
      Visible         =   0   'False
      Width           =   3570
   End
   Begin VB.Image ImgDesigR 
      Height          =   975
      Left            =   180
      Top             =   2925
      Visible         =   0   'False
      Width           =   3570
   End
   Begin VB.Image ImgSectionR 
      Height          =   975
      Left            =   180
      Top             =   4170
      Visible         =   0   'False
      Width           =   3570
   End
   Begin VB.Image ImgWDR 
      Height          =   975
      Left            =   150
      Top             =   5400
      Visible         =   0   'False
      Width           =   3570
   End
   Begin VB.Image ImgEmpR 
      Height          =   975
      Left            =   7995
      Top             =   5310
      Visible         =   0   'False
      Width           =   3570
   End
   Begin VB.Image ImgAttnB 
      Height          =   975
      Left            =   8055
      Top             =   555
      Width           =   3570
   End
   Begin VB.Image ImgAttnR 
      Height          =   975
      Left            =   8055
      Top             =   555
      Visible         =   0   'False
      Width           =   3570
   End
   Begin VB.Image ImgLSB 
      Height          =   975
      Left            =   8055
      Top             =   2880
      Width           =   3570
   End
   Begin VB.Image ImgLSR 
      Height          =   975
      Left            =   8070
      Top             =   2850
      Visible         =   0   'False
      Width           =   3570
   End
   Begin VB.Image ImgRepB 
      Height          =   975
      Left            =   4035
      Top             =   5310
      Visible         =   0   'False
      Width           =   3570
   End
   Begin VB.Image ImgRepR 
      Height          =   975
      Left            =   4035
      Top             =   5310
      Visible         =   0   'False
      Width           =   3570
   End
   Begin VB.Image ImgAbtB 
      Height          =   975
      Left            =   8040
      Top             =   7140
      Width           =   3570
   End
   Begin VB.Image ImgAbtR 
      Height          =   975
      Left            =   8040
      Top             =   7140
      Visible         =   0   'False
      Width           =   3570
   End
   Begin VB.Image ImgExitR 
      Height          =   975
      Left            =   4200
      Top             =   7140
      Width           =   3570
   End
   Begin VB.Image ImgExitB 
      Height          =   975
      Left            =   4170
      Top             =   7140
      Width           =   3570
   End
   Begin VB.Image ImgPSB 
      Height          =   975
      Left            =   8040
      Top             =   4080
      Width           =   3570
   End
   Begin VB.Image ImgPSR 
      Height          =   975
      Left            =   8040
      Top             =   4080
      Visible         =   0   'False
      Width           =   3570
   End
   Begin VB.Image ImgPDCB 
      Height          =   975
      Left            =   8040
      Top             =   1725
      Width           =   3570
   End
   Begin VB.Image imgPDCR 
      Height          =   975
      Left            =   8040
      Top             =   1725
      Visible         =   0   'False
      Width           =   3570
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DataMissing As Boolean
Private Sub cmdExit_Click()
Ans = MsgBox("Do you want to Exit ?", vbYesNo, "OASYS")

If Ans = vbYes Then
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
frmMain.BackColor = vbBlack
Call LoadButtons 'This function is to load images in the image box, to make it look like command buttons
Call ChangeDateFormat 'This function is to change the System Date format to "dd/mm/yyyy"
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
                rsRetired!address = rsEmp!address
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

Private Sub ImgHelpB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgHelpB_Down
End Sub

Private Sub ImgHelpB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgHelpB_Up
End Sub

Private Sub ImgHelpR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgHelpR_Up
End Sub

Private Sub ImgSectionB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgSectionB_Down
End Sub
Private Sub ImgSectionB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgSectionB_Up
End Sub
Private Sub ImgSectionR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgSectionR_Up
End Sub

Private Sub ImgWDB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgWDB_Down
End Sub
Private Sub ImgWDB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgWDB_Up
End Sub
Private Sub ImgWDR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgWDR_Up
End Sub

Private Sub ImgEmpB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DataMissing = False
rsType.Requery
rsClass.Requery
rsCaste.Requery
rsDesig.Requery
rsSection.Requery
Call CheckDatabase
If DataMissing = True Then Exit Sub
    Call ImgEmpB_Down
End Sub
Private Sub ImgEmpB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgEmpB_Up
End Sub
Private Sub ImgEmpR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgEmpR_Up
End Sub

Private Sub ImgLSB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgLSB_Down
End Sub
Private Sub ImgLSB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgLSB_Up
End Sub
Private Sub ImgLSR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgLSR_Up
End Sub

Private Sub ImgREPB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgREPB_Down
End Sub
Private Sub ImgREPB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgREPB_Up
End Sub
Private Sub ImgREPR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgREPR_Up
End Sub

Private Sub ImgAbtB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgAbtB_Down
End Sub
Private Sub ImgAbtB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgAbtB_Up
End Sub
Private Sub ImgAbtR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgAbtR_Up
End Sub

Private Sub ImgExitR_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgExitR_Down
End Sub
Private Sub ImgExitB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgExitB_Up
End Sub
Private Sub ImgExitR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgExitR_Up
End Sub

Private Sub ImgPSB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgPSB_Down
End Sub
Private Sub ImgPSB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgPSB_Up
End Sub
Private Sub ImgPSR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgPSR_Up
End Sub

Private Sub ImgAttnB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgAttnB_Down
End Sub
Private Sub ImgAttnB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgAttnB_Up
End Sub
Private Sub ImgAttnR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgAttnR_Up
End Sub

Private Sub ImgPDCB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgPDCB.Visible = False
imgPDCR.Visible = True
End Sub
Private Sub ImgPDCB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Load frmPayDeduct
frmPayDeduct.Show
frmMain.Hide
ImgPDCB.Visible = True
imgPDCR.Visible = False
End Sub
Private Sub ImgPDCR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgPDCB.Visible = True
imgPDCR.Visible = False
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

ImgPDCB.Picture = LoadPicture(App.Path & "\Blue.jpg")
imgPDCR.Picture = LoadPicture(App.Path & "\Red.jpg")

ImgLSB.Picture = LoadPicture(App.Path & "\Blue.jpg")
ImgLSR.Picture = LoadPicture(App.Path & "\Red.jpg")

ImgPSB.Picture = LoadPicture(App.Path & "\Blue.jpg")
ImgPSR.Picture = LoadPicture(App.Path & "\Red.jpg")

ImgRepB.Picture = LoadPicture(App.Path & "\Blue.jpg")
ImgRepR.Picture = LoadPicture(App.Path & "\Red.jpg")

ImgExitB.Picture = LoadPicture(App.Path & "\Blue.jpg")
ImgExitR.Picture = LoadPicture(App.Path & "\Red.jpg")

ImgHelpB.Picture = LoadPicture(App.Path & "\Blue.jpg")
ImgHelpR.Picture = LoadPicture(App.Path & "\Red.jpg")

End Sub


'Functions for making Image boxes work as command button
'are declared over here
'======================================================
Private Sub ImgTypeB_Down()
ImgTypeB.Visible = False
ImgTypeR.Visible = True
End Sub
Private Sub ImgTypeB_Up()
FormName = "Type"
Load frmMaster
frmMaster.Caption = "Type Master"
frmMaster.Show
'frmMastType.Show
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
FormName = "Caste"
Load frmMaster
frmMaster.Caption = "Caste Master"
frmMaster.Show
'frmCaste.Show
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
FormName = "Desig"
Load frmMaster
frmMaster.Caption = "Designation Master"
frmMaster.Show
'frmMastDesig.Show
frmMain.Hide
ImgDesigB.Visible = True
ImgDesigR.Visible = False
End Sub
Private Sub ImgDesigR_Up()
ImgDesigB.Visible = True
ImgDesigR.Visible = False
End Sub

Private Sub ImgHelpB_Down()
ImgHelpB.Visible = False
ImgHelpR.Visible = True
End Sub
Private Sub ImgHelpB_Up()
'FormName = "Help"
'Load frmMaster
'frmMaster.Caption = "Help Master"
'frmMaster.Show
'frmHelp.Show
'frmMain.Hide
ImgHelpB.Visible = True
ImgHelpR.Visible = False
End Sub
Private Sub ImgHelpR_Up()
ImgHelpB.Visible = True
ImgHelpR.Visible = False
End Sub

Private Sub ImgSectionB_Down()
ImgSectionB.Visible = False
ImgSectionR.Visible = True
End Sub
Private Sub ImgSectionB_Up()
FormName = "Section"
Load frmMaster
frmMaster.Caption = "Section Master"
frmMaster.Show
'frmMastSection.Show
frmMain.Hide
ImgSectionB.Visible = True
ImgSectionR.Visible = False
End Sub
Private Sub ImgSectionR_Up()
ImgSectionB.Visible = True
ImgSectionR.Visible = False
End Sub

Private Sub ImgWDB_Down()
ImgWDB.Visible = False
ImgWDR.Visible = True
End Sub
Private Sub ImgWDB_Up()
Load frmMastWD
frmMastWD.Show
frmMain.Hide
ImgWDB.Visible = True
ImgWDR.Visible = False
End Sub
Private Sub ImgWDR_Up()
ImgWDB.Visible = True
ImgWDR.Visible = False
End Sub

Private Sub ImgEmpB_Down()
ImgEmpB.Visible = False
ImgEmpR.Visible = True
End Sub
Private Sub ImgEmpB_Up()
Load frmMastEmp
frmMastEmp.Show
frmMain.Hide
ImgEmpB.Visible = True
ImgEmpR.Visible = False
End Sub
Private Sub ImgEmpR_Up()
ImgEmpB.Visible = True
ImgEmpR.Visible = False
End Sub

Private Sub ImgLSB_Down()
ImgLSB.Visible = False
ImgLSR.Visible = True
End Sub
Private Sub ImgLSB_Up()
Load frmLeaveSlip
frmLeaveSlip.Show
frmMain.Hide
ImgLSB.Visible = True
ImgLSR.Visible = False
End Sub
Private Sub ImgLSR_Up()
ImgLSB.Visible = True
ImgLSR.Visible = False
End Sub

Private Sub ImgREPB_Down()
ImgRepB.Visible = False
ImgRepR.Visible = True
End Sub
Private Sub ImgREPB_Up()
Load frmReports
frmReports.Show
frmMain.Hide
ImgRepB.Visible = True
ImgRepR.Visible = False
End Sub
Private Sub ImgREPR_Up()
ImgRepB.Visible = True
ImgRepR.Visible = False
End Sub

Private Sub ImgAbtB_Down()
ImgAbtB.Visible = False
ImgAbtR.Visible = True
End Sub
Private Sub ImgAbtB_Up()
FormName = "HRM"
Load frmAbout
frmAbout.Show
frmMain.Hide
ImgAbtB.Visible = True
ImgAbtR.Visible = False
End Sub
Private Sub ImgAbtR_Up()
ImgAbtB.Visible = True
ImgAbtR.Visible = False
End Sub

Private Sub ImgExitR_Up()
ImgExitR.Visible = True
ImgExitB.Visible = False
frmControl.Show
frmMain.Hide
End Sub
Private Sub ImgExitB_Up()
ImgExitR.Visible = True
ImgExitB.Visible = False
End Sub
Private Sub ImgExitR_Down()
ImgExitR.Visible = False
ImgExitB.Visible = True
End Sub


Private Sub ImgPSB_Down()
ImgPSB.Visible = False
ImgPSR.Visible = True
End Sub
Private Sub ImgPSB_Up()
Load frmPay
frmPay.Show
frmMain.Hide
ImgPSB.Visible = True
ImgPSR.Visible = False
End Sub
Private Sub ImgPSR_Up()
ImgPSB.Visible = True
ImgPSR.Visible = False
End Sub

Private Sub ImgPDCB_Down()
ImgPDCB.Visible = False
imgPDCR.Visible = True
End Sub
Private Sub ImgPDCB_Up()
Load frmPayDeduct
frmPayDeduct.Show
frmMain.Hide
ImgPDCB.Visible = True
imgPDCR.Visible = False
End Sub
Private Sub ImgPDCR_Up()
ImgPDCB.Visible = True
imgPDCR.Visible = False
End Sub

Private Sub ImgAttnB_Down()
ImgAttnB.Visible = False
ImgAttnR.Visible = True
End Sub
Private Sub ImgAttnB_Up()
Load frmAttn
frmAttn.Show
frmMain.Hide
ImgAttnB.Visible = True
ImgAttnR.Visible = False
End Sub
Private Sub ImgAttnR_Up()
ImgAttnB.Visible = True
ImgAttnR.Visible = False
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

Private Sub lblHelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgHelpB_Down
End Sub

Private Sub lblHelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgHelpB_Up
End Sub

Private Sub lblSection_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgSectionB_Down
End Sub

Private Sub lblSection_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgSectionB_Up
End Sub

Private Sub lblWD_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgWDB_Down
End Sub

Private Sub lblWD_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgWDB_Up
End Sub

Private Sub lblEmp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgEmpB_Down
End Sub

Private Sub lblEmp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgEmpB_Up
End Sub

Private Sub lblLS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgLSB_Down
End Sub

Private Sub lblLS_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgLSB_Up
End Sub

Private Sub lblREP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgREPB_Down
End Sub

Private Sub lblREP_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgREPB_Up
End Sub

Private Sub lblAbt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgAbtB_Down
End Sub

Private Sub lblAbt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgAbtB_Up
End Sub

Private Sub lblExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgExitR_Down
End Sub
Private Sub lblExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgExitR_Up
End Sub

Private Sub lblPS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgPSB_Down
End Sub
Private Sub lblPS_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgPSB_Up
End Sub

Private Sub lblPD_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgPDCB_Down
End Sub
Private Sub lblPD_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgPDCB_Up
End Sub

Private Sub lblAttn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgAttnB_Down
End Sub
Private Sub lblAttn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgAttnB_Up
End Sub

Private Sub CheckDatabase()
On Error GoTo CDErr
If rsClass.RecordCount = 0 Then
    MsgBox "Enter atleast 1 Class to continue !", vbOKOnly, "OASYS"
    DataMissing = True
    FormName = "Class"
    frmMaster.Caption = "Class Master"
    Load frmMaster
    frmMaster.Show
ElseIf rsCaste.RecordCount = 0 Then
    MsgBox "Enter atleast 1 Caste to continue !", vbOKOnly, "OASYS"
    DataMissing = True
    FormName = "Caste"
    frmMaster.Caption = "Caste Master"
    Load frmMaster
    frmMaster.Show
ElseIf rsDesig.RecordCount = 0 Then
    MsgBox "Enter atleast 1 Designation to continue !", vbOKOnly, "OASYS"
    DataMissing = True
    FormName = "Desig"
    frmMaster.Caption = "Designation Master"
    Load frmMaster
    frmMaster.Show
ElseIf rsSection.RecordCount = 0 Then
    MsgBox "Enter atleast 1 Section to continue !", vbOKOnly, "OASYS"
    DataMissing = True
    frmMaster.Caption = "Section Master"
    FormName = "Section"
    Load frmMaster
    frmMaster.Show
ElseIf rsType.RecordCount = 0 Then
    MsgBox "Enter atleast 1 Type to continue !", vbOKOnly, "OASYS"
    DataMissing = True
    FormName = "Type"
    frmMaster.Caption = "Type Master"
    Load frmMaster
    frmMaster.Show
End If
Exit Sub
CDErr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub
