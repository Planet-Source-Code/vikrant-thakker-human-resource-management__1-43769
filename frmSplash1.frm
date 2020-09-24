VERSION 5.00
Begin VB.Form frmSplash1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OASYS - An Office Automation System"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   7560
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   4000
      Left            =   420
      Top             =   2760
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Version 1.0.0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000FD2F4&
      Height          =   255
      Index           =   7
      Left            =   5460
      TabIndex        =   12
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Contact : vikrant_thakker@yahoo.com"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EF7E72&
      Height          =   315
      Index           =   5
      Left            =   420
      TabIndex        =   11
      Top             =   3840
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Developed By Vikrant Thakker"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EF7E72&
      Height          =   255
      Index           =   4
      Left            =   420
      TabIndex        =   10
      Top             =   2040
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Developed Exclusively For  AnaSys Softwares"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EF7E72&
      Height          =   495
      Index           =   6
      Left            =   420
      TabIndex        =   9
      Top             =   2400
      Width           =   6855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   ":  Daily Profit/Loss record maintenance"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EF7E72&
      Height          =   255
      Index           =   2
      Left            =   2010
      TabIndex        =   8
      Top             =   1200
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Including :  Payroll Management System"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EF7E72&
      Height          =   255
      Index           =   1
      Left            =   1050
      TabIndex        =   7
      Top             =   840
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "OASYS - AN Office Automation System"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000FD2F4&
      Height          =   495
      Index           =   0
      Left            =   570
      TabIndex        =   6
      Top             =   300
      Width           =   6435
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Special Thanks to my friend : GEETHA THAKKAR :-)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EF7E72&
      Height          =   435
      Index           =   9
      Left            =   2280
      TabIndex        =   5
      Top             =   4920
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Undertaken as a BCA Final Semister Project"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EF7E72&
      Height          =   255
      Index           =   10
      Left            =   2880
      TabIndex        =   4
      Top             =   2880
      Width           =   4395
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "For, Indira Gandhi National Open Uni."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EF7E72&
      Height          =   255
      Index           =   11
      Left            =   3240
      TabIndex        =   3
      Top             =   3120
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "AHMEDABAD"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EF7E72&
      Height          =   255
      Index           =   12
      Left            =   4740
      TabIndex        =   2
      Top             =   3360
      Width           =   1395
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "WINNERS DON'T DO DIFFERENT THINGS. THEY DO THINGS DIFFERENTLY."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000FD2F4&
      Height          =   1215
      Index           =   1
      Left            =   2640
      TabIndex        =   1
      Top             =   5760
      Width           =   2235
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   180
      X2              =   7380
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   180
      X2              =   7380
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   2
      X1              =   60
      X2              =   7020
      Y1              =   4740
      Y2              =   4740
   End
   Begin VB.Image Image1 
      Height          =   2430
      Left            =   360
      Picture         =   "frmSplash1.frx":0000
      Top             =   4800
      Width           =   1305
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   180
      X2              =   7380
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Shape frmSplash1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   7575
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   60
      Width           =   7215
   End
   Begin VB.Image Image2 
      Height          =   1920
      Left            =   5460
      Picture         =   "frmSplash1.frx":0D56
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Jai Swaminarayan"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   13
      Left            =   5220
      TabIndex        =   0
      Top             =   6960
      Width           =   1935
   End
End
Attribute VB_Name = "frmSplash1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
Call start
End Sub

Private Sub Frame1_Click()
Call start
End Sub

Private Sub Image1_Click()
Call start
End Sub

Private Sub Timer1_Timer()
Call start
End Sub
Private Sub start()
Load frmControl
frmControl.Show
Unload Me
End Sub

