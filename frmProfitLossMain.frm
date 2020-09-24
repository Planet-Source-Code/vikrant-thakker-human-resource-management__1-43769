VERSION 5.00
Begin VB.Form frmProfitLossMain 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PROFIT - LOSS "
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblIncS 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Income Sources"
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
      Left            =   1200
      TabIndex        =   3
      Top             =   1695
      Width           =   1695
   End
   Begin VB.Label lblExpS 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expense Sources"
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
      Height          =   315
      Left            =   900
      TabIndex        =   2
      Top             =   3045
      Width           =   2250
   End
   Begin VB.Label lblExp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expense Data Entry"
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
      Height          =   315
      Left            =   5490
      TabIndex        =   1
      Top             =   3045
      Width           =   2850
   End
   Begin VB.Label lblInc 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Income Data Entry"
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
      Height          =   315
      Left            =   5520
      TabIndex        =   0
      Top             =   1695
      Width           =   2910
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   4230
      TabIndex        =   4
      Top             =   4695
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PROFIT AND LOSS"
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
      Left            =   60
      TabIndex        =   5
      Top             =   0
      Width           =   8835
   End
   Begin VB.Image ImgExitR 
      Height          =   975
      Left            =   2730
      Top             =   4380
      Width           =   3570
   End
   Begin VB.Image imgIncSB 
      Height          =   975
      Left            =   300
      Top             =   1320
      Width           =   3570
   End
   Begin VB.Image imgExpSB 
      Height          =   975
      Left            =   300
      Top             =   2730
      Width           =   3570
   End
   Begin VB.Image imgIncB 
      Height          =   975
      Left            =   5100
      Top             =   1320
      Width           =   3570
   End
   Begin VB.Image imgExpB 
      Height          =   975
      Left            =   5100
      Top             =   2730
      Width           =   3570
   End
   Begin VB.Image imgExpSR 
      Height          =   975
      Left            =   300
      Top             =   2730
      Visible         =   0   'False
      Width           =   3570
   End
   Begin VB.Image imgIncSR 
      Height          =   975
      Left            =   300
      Top             =   1320
      Visible         =   0   'False
      Width           =   3570
   End
   Begin VB.Image ImgExpR 
      Height          =   975
      Left            =   5100
      Top             =   2730
      Visible         =   0   'False
      Width           =   3570
   End
   Begin VB.Image ImgIncR 
      Height          =   975
      Left            =   5100
      Top             =   1320
      Visible         =   0   'False
      Width           =   3570
   End
   Begin VB.Image ImgExitB 
      Height          =   975
      Left            =   2760
      Top             =   4380
      Width           =   3570
   End
End
Attribute VB_Name = "frmProfitLossMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
frmProfitLossMain.BackColor = vbBlack
Call LoadButtons
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

Private Sub LoadButtons()
imgIncSB.Picture = LoadPicture(App.Path & "\blue.jpg")
imgIncSR.Picture = LoadPicture(App.Path & "\red.jpg")

imgExpSB.Picture = LoadPicture(App.Path & "\blue.jpg")
imgExpSR.Picture = LoadPicture(App.Path & "\red.jpg")

ImgExitB.Picture = LoadPicture(App.Path & "\blue.jpg")
ImgExitR.Picture = LoadPicture(App.Path & "\red.jpg")

imgIncB.Picture = LoadPicture(App.Path & "\blue.jpg")
ImgIncR.Picture = LoadPicture(App.Path & "\red.jpg")

imgExpB.Picture = LoadPicture(App.Path & "\blue.jpg")
ImgExpR.Picture = LoadPicture(App.Path & "\red.jpg")
End Sub

Private Sub ImgIncSB_Down()
imgIncSB.Visible = False
imgIncSR.Visible = True
End Sub
Private Sub ImgincsB_Up()
FormName = "Income"
frmMaster.Caption = "Income Sources"
Load frmMaster
frmMaster.Show
frmMain.Hide
imgIncSB.Visible = True
imgIncSR.Visible = False
End Sub
Private Sub ImgIncSR_Up()
imgIncSB.Visible = True
imgIncSR.Visible = False
End Sub

Private Sub ImgExpSB_Down()
imgExpSB.Visible = False
imgExpSR.Visible = True
End Sub
Private Sub ImgExpsB_Up()
FormName = "Expense"
frmMaster.Caption = "Expense Sources"
Load frmMaster
frmMaster.Show
frmMain.Hide
imgExpSB.Visible = True
imgExpSR.Visible = False
End Sub
Private Sub ImgExpSR_Up()
imgExpSB.Visible = True
imgExpSR.Visible = False
End Sub

Private Sub ImgExitR_Up()
ImgExitR.Visible = True
ImgExitB.Visible = False
frmControl.Show
frmProfitLossMain.Hide
End Sub
Private Sub ImgExitB_Up()
ImgExitR.Visible = True
ImgExitB.Visible = False
End Sub
Private Sub ImgExitR_Down()
ImgExitR.Visible = False
ImgExitB.Visible = True
End Sub


Private Sub ImgIncB_Down()
imgIncB.Visible = False
ImgIncR.Visible = True
End Sub
Private Sub ImgincB_Up()
FormName = "IncomeEntry"
frmIncomeExpense.Caption = "Income Data Entry"
Load frmIncomeExpense
frmIncomeExpense.Show
frmMain.Hide
imgIncB.Visible = True
ImgIncR.Visible = False
End Sub
Private Sub ImgIncR_Up()
imgIncB.Visible = True
ImgIncR.Visible = False
End Sub

Private Sub ImgExpB_Down()
imgExpB.Visible = False
ImgExpR.Visible = True
End Sub
Private Sub ImgExpB_Up()
FormName = "ExpenseEntry"
frmIncomeExpense.Caption = "Expense Data Entry"
Load frmIncomeExpense
frmIncomeExpense.Show
frmMain.Hide
imgExpB.Visible = True
ImgExpR.Visible = False
End Sub
Private Sub ImgExpR_Up()
imgExpB.Visible = True
ImgExpR.Visible = False
End Sub
Private Sub ImgIncSB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgIncSB_Down
End Sub
Private Sub ImgIncSB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgincsB_Up
End Sub
Private Sub ImgIncSR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgIncSR_Up
End Sub

Private Sub ImgExpSB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgExpSB_Down
End Sub
Private Sub ImgExpSB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgExpsB_Up
End Sub
Private Sub ImgExpSR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgExpSR_Up
End Sub


Private Sub ImgIncB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgIncB_Down
End Sub
Private Sub ImgIncB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgincB_Up
End Sub
Private Sub ImgIncR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgIncR_Up
End Sub

Private Sub ImgExpB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgExpB_Down
End Sub
Private Sub ImgExpB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgExpB_Up
End Sub
Private Sub ImgExpR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgExpR_Up
End Sub


Private Sub lblIncS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgIncSB_Down
End Sub

Private Sub lblIncS_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgincsB_Up
End Sub

Private Sub lblExpS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgExpSB_Down
End Sub

Private Sub lblExpS_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgExpsB_Up
End Sub

Private Sub lblInc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgIncB_Down
End Sub

Private Sub lblInc_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgincB_Up
End Sub

Private Sub lblExp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgExpB_Down
End Sub

Private Sub lblExp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgExpB_Up
End Sub

Private Sub lblExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgExitR_Down
End Sub
Private Sub lblExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ImgExitR_Up
End Sub


