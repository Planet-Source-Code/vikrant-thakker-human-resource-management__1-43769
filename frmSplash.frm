VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4740
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7650
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7635
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   2970
         Top             =   225
      End
      Begin VB.Image Image1 
         Height          =   4515
         Left            =   30
         Top             =   120
         Width           =   7575
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
Call vicky
End Sub

Private Sub Frame1_Click()
Call vicky
End Sub

Private Sub Image1_Click()
Call vicky
End Sub

Private Sub Timer1_Timer()
Call vicky
End Sub
Private Sub vicky()
Load frmMain
frmMain.Show
Unload Me
End Sub
