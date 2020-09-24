VERSION 5.00
Begin VB.Form frmControl 
   BackColor       =   &H00F3A965&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MAIN FORM"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5820
   ForeColor       =   &H00C4E3CF&
   Icon            =   "frmControl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MAIN SCREEN"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   -120
      TabIndex        =   7
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label lblInv 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INVENTORY"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   6165
      TabIndex        =   6
      Top             =   5175
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label lblMan 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SYSTEM MANUAL"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   2085
      TabIndex        =   5
      Top             =   4065
      Width           =   1500
   End
   Begin VB.Label lblPL 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PROFIT AND LOSS"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   2055
      TabIndex        =   4
      Top             =   2115
      Width           =   1590
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   270
      Left            =   2640
      TabIndex        =   3
      Top             =   6075
      Width           =   435
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ABOUT"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   2595
      TabIndex        =   2
      Top             =   5085
      Width           =   600
   End
   Begin VB.Label lblRep 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "REPORTS"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   2430
      TabIndex        =   1
      Top             =   3075
      Width           =   810
   End
   Begin VB.Label lblHRM 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "H-R-M"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   2520
      TabIndex        =   0
      Top             =   1125
      Width           =   570
   End
   Begin VB.Image ImgHRMB 
      Height          =   675
      Left            =   1740
      Top             =   930
      Width           =   2190
   End
   Begin VB.Image ImgInvB 
      Height          =   675
      Left            =   5580
      Top             =   4980
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.Image ImgManB 
      Height          =   675
      Left            =   1740
      Top             =   3870
      Width           =   2190
   End
   Begin VB.Image ImgPLB 
      Height          =   675
      Left            =   1740
      Top             =   1920
      Width           =   2190
   End
   Begin VB.Image ImgAboutB 
      Height          =   675
      Left            =   1740
      Top             =   4890
      Width           =   2190
   End
   Begin VB.Image ImgRepB 
      Height          =   675
      Left            =   1740
      Top             =   2880
      Width           =   2190
   End
   Begin VB.Image ImgExitR 
      Height          =   675
      Left            =   1740
      Top             =   5880
      Width           =   2190
   End
   Begin VB.Image ImgInvR 
      Height          =   675
      Left            =   5580
      Top             =   4980
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.Image ImgHRMR 
      Height          =   675
      Left            =   1740
      Top             =   930
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.Image ImgPLR 
      Height          =   675
      Left            =   1740
      Top             =   1920
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.Image ImgManR 
      Height          =   675
      Left            =   1740
      Top             =   3870
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.Image ImgRepR 
      Height          =   675
      Left            =   1740
      Top             =   2880
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.Image ImgAboutR 
      Height          =   675
      Left            =   1740
      Top             =   4890
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.Image ImgExitB 
      Height          =   675
      Left            =   1740
      Top             =   5880
      Visible         =   0   'False
      Width           =   2190
   End
End
Attribute VB_Name = "frmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'frmControl.BackColor = vbBlack
Call LoadButtons
End Sub

Private Sub LoadButtons()
'blue.jpg
'red.jpg
ImgInvB.Picture = LoadPicture(App.Path & "\green.jpg")
ImgInvR.Picture = LoadPicture(App.Path & "\yellow.jpg")

ImgPLB.Picture = LoadPicture(App.Path & "\green.jpg")
ImgPLR.Picture = LoadPicture(App.Path & "\yellow.jpg")

ImgExitB.Picture = LoadPicture(App.Path & "\green.jpg")
ImgExitR.Picture = LoadPicture(App.Path & "\yellow.jpg")

ImgManB.Picture = LoadPicture(App.Path & "\green.jpg")
ImgManR.Picture = LoadPicture(App.Path & "\yellow.jpg")

ImgAboutB.Picture = LoadPicture(App.Path & "\green.jpg")
ImgAboutR.Picture = LoadPicture(App.Path & "\yellow.jpg")

ImgHRMB.Picture = LoadPicture(App.Path & "\green.jpg")
ImgHRMR.Picture = LoadPicture(App.Path & "\yellow.jpg")

ImgRepB.Picture = LoadPicture(App.Path & "\green.jpg")
ImgRepR.Picture = LoadPicture(App.Path & "\yellow.jpg")

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

Private Sub ImgExitR_Up()
Dim Ans As String
ImgExitR.Visible = True
ImgExitB.Visible = False
Ans = MsgBox("Do you want to Exit ?", vbYesNo, "OAYSY")
If Ans = vbYes Then
    Call UnLoadAll
    End
End If
End Sub
Private Sub ImgExitB_Up()
ImgExitR.Visible = True
ImgExitB.Visible = False
End Sub
Private Sub ImgExitR_Down()
ImgExitR.Visible = False
ImgExitB.Visible = True
End Sub

Private Sub lblExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgExitR_Down
End Sub
Private Sub lblExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgExitR_Up
End Sub


Private Sub ImgHRMB_Down()
ImgHRMB.Visible = False
ImgHRMR.Visible = True
End Sub
Private Sub ImgHRMB_Up()
Load frmMain
frmMain.Show
frmControl.Hide
ImgHRMB.Visible = True
ImgHRMR.Visible = False
End Sub
Private Sub ImgHRMR_Up()
ImgHRMB.Visible = True
ImgHRMR.Visible = False
End Sub

Private Sub ImgHRMB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgHRMB_Down
End Sub
Private Sub ImgHRMB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgHRMB_Up
End Sub
Private Sub ImgHRMR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgHRMR_Up
End Sub


Private Sub lblHRM_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgHRMB_Down
End Sub

Private Sub lblHRM_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgHRMB_Up
End Sub

'---
Private Sub ImgREPB_Down()
ImgRepB.Visible = False
ImgRepR.Visible = True
End Sub
Private Sub ImgREPB_Up()
Load frmReports
frmReports.Show
frmControl.Hide
ImgRepB.Visible = True
ImgRepR.Visible = False
End Sub
Private Sub ImgREPR_Up()
ImgRepB.Visible = True
ImgRepR.Visible = False
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


Private Sub lblREP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgREPB_Down
End Sub

Private Sub lblREP_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgREPB_Up
End Sub

'----
Private Sub ImgAboutB_Down()
ImgAboutB.Visible = False
ImgAboutR.Visible = True
End Sub
Private Sub ImgAboutB_Up()
FormName = "Control"
Load frmAbout
frmAbout.Show
frmControl.Hide
ImgAboutB.Visible = True
ImgAboutR.Visible = False
End Sub
Private Sub ImgAboutR_Up()
ImgAboutB.Visible = True
ImgAboutR.Visible = False
End Sub

Private Sub ImgAboutB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgAboutB_Down
End Sub
Private Sub ImgAboutB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgAboutB_Up
End Sub
Private Sub ImgAboutR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgAboutR_Up
End Sub


Private Sub lblAbout_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgAboutB_Down
End Sub

Private Sub lblAbout_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgAboutB_Up
End Sub

'------
Private Sub ImgINVB_Down()
ImgInvB.Visible = False
ImgInvR.Visible = True
End Sub
Private Sub ImgINVB_Up()
ImgInvB.Visible = True
ImgInvR.Visible = False
MsgBox "INVENTORY MANAGEMENT UNDER CONSTRUCTION !", vbExclamation, "OASYS"
End Sub
Private Sub ImgINVR_Up()
ImgInvB.Visible = True
ImgInvR.Visible = False
End Sub

Private Sub ImgINVB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgINVB_Down
End Sub
Private Sub ImgINVB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgINVB_Up
End Sub
Private Sub ImgINVR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgINVR_Up
End Sub

Private Sub lblINV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgINVB_Down
End Sub

Private Sub lblINV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgINVB_Up
End Sub

'----
Private Sub ImgPLB_Down()
ImgPLB.Visible = False
ImgPLR.Visible = True
End Sub
Private Sub ImgPLB_Up()
Load frmProfitLossMain
frmProfitLossMain.Show
frmControl.Hide
ImgPLB.Visible = True
ImgPLR.Visible = False
End Sub
Private Sub ImgPLR_Up()
ImgPLB.Visible = True
ImgPLR.Visible = False
End Sub

Private Sub ImgPLB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgPLB_Down
End Sub
Private Sub ImgPLB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgPLB_Up
End Sub
Private Sub ImgPLR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgPLR_Up
End Sub


Private Sub lblPL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgPLB_Down
End Sub

Private Sub lblPL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgPLB_Up
End Sub

'----
Private Sub ImgManB_Down()
ImgManB.Visible = False
ImgManR.Visible = True
End Sub
Private Sub ImgManB_Up()
ImgManB.Visible = True
ImgManR.Visible = False
MsgBox "SYSTEM MANUAL UNDER CONSTRUCTION !", vbExclamation, "OASYS"
End Sub
Private Sub ImgManR_Up()
ImgManB.Visible = True
ImgManR.Visible = False
End Sub

Private Sub ImgManB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgManB_Down
End Sub
Private Sub ImgManB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgManB_Up
End Sub
Private Sub ImgManR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgManR_Up
End Sub


Private Sub lblMan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgManB_Down
End Sub

Private Sub lblMan_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgManB_Up
End Sub

