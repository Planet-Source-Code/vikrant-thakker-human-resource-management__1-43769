VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPay 
   Appearance      =   0  'Flat
   BackColor       =   &H009BF4C8&
   Caption         =   "PAYSLIP"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Caption         =   "GENERATE &PAYSLIP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   7680
      Width           =   1230
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
      Height          =   540
      Left            =   9270
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   7185
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Framebutton 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   2700
      TabIndex        =   28
      Top             =   7695
      Width           =   6495
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H0080C0FF&
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdRemove 
         BackColor       =   &H0080C0FF&
         Caption         =   "Remove"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3690
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdPrev 
         BackColor       =   &H0080C0FF&
         Caption         =   "Previous"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4590
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   45
         Width           =   915
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H0080C0FF&
         Caption         =   "Next"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5535
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H0080C0FF&
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1890
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdModify 
         BackColor       =   &H0080C0FF&
         Caption         =   "Modify"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   60
         Width           =   915
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H0080C0FF&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2790
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   45
         Width           =   870
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00E0E0E0&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   7680
      Width           =   990
   End
   Begin VB.Frame frameDed_Exch 
      BackColor       =   &H009BF4C8&
      Height          =   7575
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   12135
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H009BF4C8&
         Caption         =   "Deductions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   4335
         Left            =   4680
         TabIndex        =   44
         Top             =   3240
         Width           =   7215
         Begin VB.Frame Frame4 
            BackColor       =   &H009BF4C8&
            Caption         =   "NET TOTAL"
            ForeColor       =   &H000000FF&
            Height          =   1155
            Left            =   4740
            TabIndex        =   60
            Top             =   3180
            Width           =   2475
            Begin VB.Label lblNet 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Left            =   180
               TabIndex        =   61
               Top             =   420
               Width           =   2055
            End
         End
         Begin VB.TextBox txtTotDed 
            Appearance      =   0  'Flat
            BackColor       =   &H009BF4C8&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   58
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   3720
            Width           =   1455
         End
         Begin VB.TextBox txtLeavesDed 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5460
            Locked          =   -1  'True
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   1260
            Width           =   1455
         End
         Begin VB.TextBox txtLoan 
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
            Height          =   375
            Left            =   5475
            TabIndex        =   12
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox txtQRent 
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
            Height          =   375
            Left            =   2280
            TabIndex        =   11
            Top             =   3090
            Width           =   1455
         End
         Begin VB.TextBox txtPTax 
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
            Height          =   375
            Left            =   2280
            TabIndex        =   10
            Top             =   2475
            Width           =   1455
         End
         Begin VB.TextBox txtGInsur 
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
            Height          =   375
            Left            =   2280
            TabIndex        =   9
            Top             =   1860
            Width           =   1455
         End
         Begin VB.TextBox txtProv 
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
            Height          =   375
            Left            =   2280
            TabIndex        =   8
            Top             =   1260
            Width           =   1455
         End
         Begin VB.TextBox txtLIC 
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
            Height          =   375
            Left            =   2280
            TabIndex        =   7
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label18 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   195
            Left            =   1320
            TabIndex        =   59
            Top             =   3840
            Width           =   525
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "LEAVES"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   195
            Left            =   4560
            TabIndex        =   55
            Top             =   1380
            Width           =   615
         End
         Begin VB.Label Label17 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "LOAN"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   195
            Left            =   4770
            TabIndex        =   50
            Top             =   720
            Width           =   435
         End
         Begin VB.Label Label16 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "QUARTERS RENT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   195
            Left            =   435
            TabIndex        =   49
            Top             =   3240
            Width           =   1395
         End
         Begin VB.Label Label15 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "PROFESSIONAL TAX"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   195
            Left            =   240
            TabIndex        =   48
            Top             =   2580
            Width           =   1590
         End
         Begin VB.Label Label14 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "GROUP INSURANCE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   195
            Left            =   285
            TabIndex        =   47
            Top             =   1980
            Width           =   1575
         End
         Begin VB.Label Label13 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "PROVIDANT FUND"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   195
            Left            =   435
            TabIndex        =   46
            Top             =   1380
            Width           =   1560
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "L.I.C."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   195
            Left            =   1500
            TabIndex        =   45
            Top             =   720
            Width           =   375
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H009BF4C8&
         Caption         =   "Allowances"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   4335
         Left            =   120
         TabIndex        =   39
         Top             =   3240
         Width           =   4575
         Begin VB.TextBox txtTotAll 
            Appearance      =   0  'Flat
            BackColor       =   &H009BF4C8&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   56
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   3660
            Width           =   1695
         End
         Begin VB.TextBox txtLeavesAllowed 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   3060
            Width           =   1695
         End
         Begin VB.TextBox txtConv 
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
            Height          =   375
            Left            =   1800
            TabIndex        =   6
            Top             =   2460
            Width           =   1695
         End
         Begin VB.TextBox txtMedical 
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
            Height          =   375
            Left            =   1800
            TabIndex        =   5
            Top             =   1830
            Width           =   1695
         End
         Begin VB.TextBox txtHouseRent 
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
            Height          =   375
            Left            =   1800
            TabIndex        =   4
            Top             =   1230
            Width           =   1695
         End
         Begin VB.TextBox txtDearness 
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
            Height          =   375
            Left            =   1800
            TabIndex        =   3
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   195
            Left            =   720
            TabIndex        =   57
            Top             =   3780
            Width           =   525
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "LEAVES"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   195
            Left            =   720
            TabIndex        =   53
            Top             =   3180
            Width           =   615
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "CONVEYANCE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   195
            Left            =   240
            TabIndex        =   43
            Top             =   2580
            Width           =   1095
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "MEDICAL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   195
            Left            =   615
            TabIndex        =   42
            Top             =   1950
            Width           =   705
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   " HOUSE RENT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   195
            Left            =   255
            TabIndex        =   41
            Top             =   1350
            Width           =   1110
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "DEARNESS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   195
            Left            =   480
            TabIndex        =   40
            Top             =   720
            Width           =   885
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H009BF4C8&
         Caption         =   "Employee Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   3240
         Left            =   120
         TabIndex        =   14
         Top             =   0
         Width           =   11865
         Begin Crystal.CrystalReport CR 
            Left            =   4800
            Top             =   1800
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            PrintFileLinesPerPage=   60
         End
         Begin VB.ComboBox cmbDummyMnt 
            Height          =   360
            Left            =   3120
            TabIndex        =   63
            Text            =   "DummyMnt"
            Top             =   2640
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.ComboBox cmbDummyYear 
            Height          =   360
            Left            =   3060
            TabIndex        =   62
            Text            =   "dummyYear"
            Top             =   1920
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.TextBox txtBasicPay 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "BASIC"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7590
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   1215
            Width           =   1575
         End
         Begin VB.TextBox txtCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "SNO"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1710
            TabIndex        =   0
            Top             =   405
            Width           =   1335
         End
         Begin VB.TextBox txtName 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "NAME"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1710
            Locked          =   -1  'True
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   1185
            Width           =   3495
         End
         Begin VB.TextBox txtDesig 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "DESIG"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7530
            Locked          =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   1965
            Width           =   3255
         End
         Begin VB.CommandButton cmdHelp 
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
            Height          =   405
            Left            =   3090
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   405
            Width           =   780
         End
         Begin VB.ComboBox cmbYear 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "frmPay.frx":0000
            Left            =   1665
            List            =   "frmPay.frx":0002
            TabIndex        =   1
            Top             =   1935
            Width           =   1350
         End
         Begin VB.TextBox txtSection 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "BASIC"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7605
            Locked          =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   405
            Width           =   1575
         End
         Begin VB.ComboBox txtMnt 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1665
            TabIndex        =   2
            Top             =   2640
            Width           =   1335
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "MONTH"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   750
            TabIndex        =   26
            Top             =   2670
            Width           =   600
         End
         Begin VB.Label Label28 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "BASIC PAY"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   6510
            TabIndex        =   25
            Top             =   1200
            Width           =   825
         End
         Begin VB.Label Label26 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "CODE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   975
            TabIndex        =   24
            Top             =   525
            Width           =   450
         End
         Begin VB.Label Label25 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "NAME"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   960
            TabIndex        =   23
            Top             =   1245
            Width           =   465
         End
         Begin VB.Label Label24 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "DESIGNATION"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   6135
            TabIndex        =   22
            Top             =   2025
            Width           =   1110
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "YEAR"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   420
            Left            =   930
            TabIndex        =   21
            Top             =   1980
            Width           =   555
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "SECTION"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   6660
            TabIndex        =   20
            Top             =   405
            Width           =   705
         End
      End
   End
   Begin MSDataGridLib.DataGrid Dg 
      Height          =   7140
      Left            =   1755
      TabIndex        =   37
      Top             =   0
      Visible         =   0   'False
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   12594
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   0
      ForeColor       =   12648447
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
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
      Caption         =   "Employee List"
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
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pay  Deduction /  Cash exchange for earned leave"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   405
      Left            =   1710
      TabIndex        =   38
      Top             =   405
      Width           =   7320
   End
End
Attribute VB_Name = "frmPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsEmpList, rsValidList As Recordset
Dim Modify, Add, ValidCode, DataUpdated As Boolean
'Dim TotAllowance, TotDeduction As Integer

' BOOLEANS DECLARED AND THE FUNCTIONS

' Modify = True if Modify button is clicked
' Add = True if Add button is clicked
' ValidCode = True if the employee code entered is valid (does exist)

' USER DEFINED FUNCTIONS AND THEIR PURPOSE

' EnableAll  : enable all the text boxes
' DisableAll : disable all the text boxes
' ClearAll   : Clear (Blank) all the text boxes
' Modi       : If modify = True then Enables all the text boxes
'              If modify = False then Disable all the text boxes
' TotalLeaveDays : This function calculates the Total No. Of Leaves by an employee in a particular month

Private Sub Search()
If rsPD.RecordCount = 0 Then
    MsgBox "NO ATTENDANCE RECORD OF ANY EMPLOYEE !", vbCritical, "OASYS"
    Call ClearAll
    Exit Sub
End If
If rsPD.RecordCount > 0 Then rsPD.MoveFirst
    For i = 0 To rsPD.RecordCount - 1 Step 1
        If (txtCode.Text = rsPD!Code) Then
            If Not IsNull(rsPD!Name) Then txtName.Text = rsPD!Name
            If Not IsNull(rsPD!Desig) Then txtDesig.Text = rsPD!Desig
            If Not IsNull(rsPD!Sect) Then txtSection.Text = rsPD!Sect
            If Not IsNull(rsPD!BasicPay) Then txtBasicPay.Text = rsPD!BasicPay
            SendKeys "{TAB}"
            KeyAscii = 0

   If Add = False Then 'If add Button is not pressed
        cmdModify.Enabled = True
        cmdRemove.Enabled = True
   End If
            Exit Sub
        End If
    rsPD.MoveNext
    If rsPD.EOF = True Then
        MsgBox "Please Select Employee Code from the List !", vbOKOnly, "OASYS"
        Call ClearAll
        txtCode.SetFocus
        Exit Sub
    End If
    Next
End Sub

Private Sub cmbYear_GotFocus()
Call txt_GotFocus
End Sub
Private Sub AddMonths()
Call ValidList
txtMnt.Clear  'First clear the Months Combo box
If Trim(cmbYear.Text) <> "" Then
    If rsValidList.RecordCount > 0 Then rsValidList.MoveFirst
    For i = 0 To rsValidList.RecordCount - 1 Step 1
        If cmbYear.Text = rsValidList!Year Then
            txtMnt.AddItem (rsValidList!Month)  ' Add Months in the combo box
        End If
        If rsValidList.EOF = False Then rsValidList.MoveNext
    Next
End If
End Sub

Private Sub AddYears()
Call ValidList
cmbDummyYear.Clear   ' First remove all the data from the Years combo box
If rsValidList.RecordCount > 0 Then rsValidList.MoveFirst
    For i = 0 To rsValidList.RecordCount - 1 Step 1
        cmbDummyYear.AddItem (rsValidList!Year)
    If rsValidList.EOF = False Then rsValidList.MoveNext
    Next
End Sub
Private Sub cmbYear_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call AddMonths
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub cmbYear_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub cmdAdd_Click()
'On Error GoTo aerr
Modify = False
Add = True   ' As Add button is clicked, Boolean Add = True

Call Modi
Call EnableAll  ' Enable all the text boxes
Call ClearAll   ' Blank all the text boxes

'rsPS.AddNew

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
'On Error GoTo cerr
Modify = False
Add = False
Call Modi
'rsPS.CancelUpdate
Call DisableAll   ' Disable all the text boxes

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

Private Sub cmdHelp_Click()  ' This is to show the Data Grid
Call DataGrid
Framebutton.Visible = False
frameDed_Exch.Visible = False
cmdClose.Visible = False
cmdPrint.Visible = False
dg.Refresh
dg.Visible = True
cmdOK.Visible = True
End Sub
Private Sub DataGrid()
'rsValidList.Open "select Code,Name,Year,Month,BasicPay,Sect,Desig from PayDeduction", conn, adOpenStatic, adLockOptimistic
cmbYear.Text = Year(Date)
Set dg.DataSource = rsPD  'List of employees whose entry is already there in the PayDeduction/Exchange table...
dg.Refresh
End Sub
Private Sub ValidList()
Set rsValidList = New ADODB.Recordset
rsValidList.Open "select Code,Name,Year,Month,BasicPay,Sect,Desig from PayDeduction where Code=" & "'" & txtCode.Text & "'", conn, adOpenStatic, adLockOptimistic
End Sub
Private Sub cmdModify_Click()
On Error GoTo merr
MsgBox "You cannot modify this data !", vbOKOnly, "OASYS"
Exit Sub
Modify = True  ' Modify button is clicked
Add = False
Call Modi
Call EnableAll  ' Enable all the text boxes

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
Add = False
Call Modi

If rsPD.RecordCount = 0 Then  'If there are no records (data) in the table
    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    cmdCancel.Enabled = False
    Exit Sub
Else
    cmdModify.Enabled = True
End If
If rsPD.EOF = False Then rsPD.MoveNext
If rsPD.EOF = True Then rsPD.MoveLast
    showall
Exit Sub
nerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub cmdOK_Click()   ' Hide the Data grid
frameDed_Exch.Visible = True
Framebutton.Visible = True
cmdClose.Visible = True
cmdPrint.Visible = True
cmdOK.Visible = False
dg.Visible = False
End Sub

Private Sub cmdPrev_Click()
'On Error GoTo perr
Modify = False
Add = False
Call Modi

If rsPD.RecordCount = 0 Then
    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    cmdCancel.Enabled = False
    cmdSave.Enabled = False
    Exit Sub
End If
If rsPD.BOF = False Then rsPD.MovePrevious
If rsPD.BOF = True Then rsPD.MoveFirst
showall
    cmdRemove.Enabled = True
Exit Sub
perr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub cmdRemove_Click()
On Error GoTo rerr
MsgBox "You cannot Delete this data !", vbOKOnly, "OASYS"
Exit Sub

Modify = False
Add = False
Call Modi
 If rsPD.RecordCount = 0 Then
    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    cmdCancel.Enabled = False
    cmdSave.Enabled = False
    Exit Sub
End If
       rsPD.Delete
       Call cmdNext_Click
Exit Sub
rerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub AddData()

If rsPD.RecordCount > 0 Then rsPD.MoveFirst
For i = 0 To rsPD.RecordCount - 1 Step 1
If (rsPD!Code = txtCode.Text) And (rsPD!Name = txtName.Text) And (rsPD!Year = cmbYear.Text) And (rsPD!Month = txtMnt.Text) Then

'ADD ALLOWANCE DETAILS
If txtDearness.Text <> "" Then rsPD!DA = txtDearness.Text
If txtHouseRent.Text <> "" Then rsPD!HRA = txtHouseRent.Text
If txtMedical.Text <> "" Then rsPD!MA = txtMedical.Text
If txtConv.Text <> "" Then rsPD!Conv = txtConv.Text

'ADD DEDUCTION DETAILS
If txtLIC.Text <> "" Then rsPD!LIC = txtLIC.Text
If txtProv.Text <> "" Then rsPD!PF = txtProv.Text
If txtGInsur.Text <> "" Then rsPD!GInsur = txtGInsur.Text
If txtPTax.Text <> "" Then rsPD!PTax = txtPTax.Text
If txtQRent.Text <> "" Then rsPD!QRent = txtQRent.Text
If txtLoan.Text <> "" Then rsPD!Loan = txtLoan.Text
If lblNet.Caption <> "" Then rsPD!NetPay = lblNet.Caption

If Year(Date) < Year(rsExpiry!CurrentYear) Then
    MsgBox "You cannot work in previous working year ! Contact Developer", vbCritical, "OASYS"
    DataUpdated = False
Exit Sub
ElseIf Year(Date) > Year(rsExpiry!CurrentYear) Then
    MsgBox "You cannot work in next working year ! Contact Developer", vbCritical, "OASYS"
    DataUpdated = False
Exit Sub
End If
rsPD.Update
DataUpdated = True
Exit Sub
End If
If rsPD.EOF = False Then rsPD.MoveNext
Next
End Sub

Private Sub cmdSave_Click()
On Error GoTo serr
Add = False

If Trim(txtCode.Text) = "" Or Trim(cmbYear.Text) = "" Or Trim(txtMnt.Text) = "" Then
    Exit Sub
End If

Call AddData

If DataUpdated = False Then Exit Sub  'If data is not updated successfully, then Exit Sub

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
cmdAdd.SetFocus

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

Private Sub Form_Load()
'On Error GoTo ferr
cmbYear.Text = Year(Date)

    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    cmdCancel.Enabled = False
    
If rsPD.RecordCount > 0 Then
    cmdAdd.Enabled = True
    cmdPrev.Enabled = True
    cmdNext.Enabled = True
End If
    cmdSave.Enabled = False
    
Exit Sub
ferr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub


Private Sub txtCode_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call DataGrid
Call AddYears
Call ClearYears
If Add = True Then Call Search
End If
End Sub

Private Sub txtCode_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtConv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    Call TotalAllowance
End If
End Sub

Private Sub txtDearness_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    Call TotalAllowance
End If
End Sub

Private Sub txtGInsur_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    Call TotalDeduction
End If
End Sub

Private Sub txtHouseRent_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    Call TotalAllowance
End If
End Sub

Private Sub txtLeavesAllowed_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    Call TotalAllowance
End If
End Sub

Private Sub txtLeavesDed_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    Call TotalDeduction
End If
End Sub

Private Sub txtLIC_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    Call TotalDeduction
End If
End Sub

Private Sub txtLoan_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    Call TotalDeduction
End If
End Sub

Private Sub txtMedical_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    Call TotalAllowance
End If
End Sub

Private Sub txtMnt_LostFocus()
Call LeavesTaken
'Call TotalLeaveDays
End Sub

Private Sub showall()
'On Error GoTo showerr
Call ClearAll   ' Clear (Blank) all the text boxes

If rsPD.RecordCount = 0 Then
    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    Exit Sub
ElseIf rsPD.RecordCount > 0 Then
    cmdModify.Enabled = True
    cmdPrev.Enabled = True
    cmdNext.Enabled = True
    cmdRemove.Enabled = True
End If

'SHOW EMPLOYEE DETAILS
If Not IsNull(rsPD!Name) Then txtName.Text = rsPD!Name
If Not IsNull(rsPD!Code) Then txtCode.Text = rsPD!Code
If Not IsNull(rsPD!Desig) Then txtDesig.Text = rsPD!Desig
If Not IsNull(rsPD!Year) Then cmbYear.Text = rsPD!Year
If Not IsNull(rsPD!Month) Then txtMnt.Text = rsPD!Month
If Not IsNull(rsPD!Sect) Then txtSection.Text = rsPD!Sect
If Not IsNull(rsPD!BasicPay) Then txtBasicPay.Text = rsPD!BasicPay

'SHOW ALLOWANCE DETAILS
If Not IsNull(rsPD!DA) Then txtDearness.Text = rsPD!DA
If Not IsNull(rsPD!HRA) Then txtHouseRent.Text = rsPD!HRA
If Not IsNull(rsPD!MA) Then txtMedical.Text = rsPD!MA
If Not IsNull(rsPD!Conv) Then txtConv.Text = rsPD!Conv

'SHOW DEDUCTION DETAILS
If Not IsNull(rsPD!LIC) Then txtLIC.Text = rsPD!LIC
If Not IsNull(rsPD!PF) Then txtProv.Text = rsPD!PF
If Not IsNull(rsPD!GInsur) Then txtGInsur.Text = rsPD!GInsur
If Not IsNull(rsPD!PTax) Then txtPTax.Text = rsPD!PTax
If Not IsNull(rsPD!QRent) Then txtQRent.Text = rsPD!QRent
If Not IsNull(rsPD!Loan) Then txtLoan.Text = rsPD!Loan

Call TotalAllowance
Call TotalDeduction
'FrameDed.Visible = True
'FrameExch.Visible = True
Exit Sub
showerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub txt_GotFocus()
On Error GoTo focerr
    If Trim(txtMnt.Text) = "" Or Trim(txtCode.Text) = "" Or Trim(txtMnt.Text) = "" Or Trim(cmbYear.Text) = "" Then
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

Private Sub dg_DblClick()
'On Error GoTo dgerr
txtCode.Text = dg.Columns(0).Text
txtName.Text = dg.Columns(1).Text
txtSection.Text = dg.Columns(3).Text
txtBasicPay.Text = dg.Columns(4).Text
txtDesig.Text = dg.Columns(2).Text
cmbYear.Text = dg.Columns(5).Text
txtMnt.Text = dg.Columns(6).Text
Call cmdOK_Click
cmbYear.SetFocus
Exit Sub
dgerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub dg_KeyPress(KeyAscii As Integer)
'On Error GoTo dgkerr
If KeyAscii = 13 Then
txtCode.Text = dg.Columns(0).Text
txtName.Text = dg.Columns(1).Text
txtSection.Text = dg.Columns(3).Text
txtBasicPay.Text = dg.Columns(4).Text
txtDesig.Text = dg.Columns(2).Text
cmbYear.Text = dg.Columns(5).Text
txtMnt.Text = dg.Columns(6).Text
Call cmdOK_Click
cmbYear.SetFocus
End If
Call cmdOK_Click
cmbYear.SetFocus

Exit Sub
dgkerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub txtMnt_GotFocus()
Call AddMonths
Call txt_GotFocus
End Sub
Private Sub LeavesTaken()
If Trim(txtCode.Text) <> "" And Trim(cmbYear.Text) <> "" And Trim(txtMnt.Text) <> "" Then
If rsPD.RecordCount > 0 Then rsPD.MoveFirst
For i = 0 To rsPD.RecordCount - 1 Step 1
    If (rsPD!Code = txtCode.Text) And (rsPD!Name = txtName.Text) And (rsPD!Year = cmbYear.Text) And (rsPD!Month = txtMnt.Text) Then
        txtLeavesAllowed.Text = rsPD!TotAmtAllowed
        txtLeavesDed.Text = rsPD!totamtded
        Call TotalAllowance
        Call TotalDeduction
        Exit Sub
    End If
If rsPD.EOF = False Then rsPD.MoveNext
Next
End If
End Sub
Private Sub txtMnt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

If Trim(txtMnt.Text) = "" Or Trim(cmbYear.Text) = "" Then
'    MsgBox ("Year and Month Fields cannot be empty !")
'    cmbYear.SetFocus
    Exit Sub
End If
  
    SendKeys "{TAB}"
    KeyAscii = 0
    Call LeavesTaken
'    Call TotalLeaveDays
End If
End Sub
Private Sub txtMnt_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub DisableAll()  ' Disable all the text boxes
txtCode.Enabled = False
txtName.Enabled = False
txtDesig.Enabled = False
txtSection.Enabled = False
txtBasicPay.Enabled = False
cmbYear.Enabled = False
txtMnt.Enabled = False
cmdHelp.Enabled = False
End Sub
Private Sub EnableAll()  ' Enable all the text boxes
txtCode.Enabled = True
txtName.Enabled = True
txtDesig.Enabled = True
txtSection.Enabled = True
txtBasicPay.Enabled = True
cmbYear.Enabled = True
txtMnt.Enabled = True
cmdHelp.Enabled = True
End Sub
Private Sub ClearAll()  ' Clear all the text boxes
txtCode.Text = ""
txtName.Text = ""
txtDesig.Text = ""
txtSection.Text = ""
txtBasicPay.Text = ""
cmbYear.Text = ""
txtMnt.Text = ""
txtTotAll.Text = "0"
txtTotDed.Text = "0"
lblNet.Caption = "0"
End Sub

Private Sub TotalAllowance()
TotAllowance = Val(txtDearness.Text) + Val(txtHouseRent.Text) + Val(txtMedical.Text) + Val(txtConv.Text) + Val(txtLeavesAllowed.Text)
txtTotAll.Text = Val(TotAllowance)
lblNet.Caption = Val(txtBasicPay.Text) + Val(txtTotAll.Text) - Val(txtTotDed.Text)
End Sub
Private Sub TotalDeduction()
TotDeduction = Val(txtLIC.Text) + Val(txtGInsur.Text) + Val(txtProv.Text) + Val(txtPTax.Text) + Val(txtQRent.Text) + Val(txtLoan.Text) + Val(txtLeavesDed.Text)
txtTotDed.Text = Val(TotDeduction)
lblNet.Caption = Val(txtBasicPay.Text) + Val(txtTotAll.Text) - Val(txtTotDed.Text)
End Sub
'**********************************************************

Private Sub txtProv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    Call TotalDeduction
End If
End Sub
Private Sub txtPTax_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    Call TotalDeduction
End If
End Sub

Private Sub txtQRent_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    Call TotalDeduction
End If
End Sub

Private Sub txtTotAll_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    Call TotalAllowance
End If
End Sub

Private Sub txtTotDed_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    Call TotalDeduction
End If
End Sub

Private Sub ClearYears()
cmbYear.Clear
If cmbDummyYear.ListCount > 0 Then
'MsgBox cmbdummyYear.ListCount
For i = 0 To cmbDummyYear.ListCount Step 1
restart:
If i = cmbDummyYear.ListCount Then Exit Sub

YR = cmbDummyYear.List(i)

    For no = 0 To cmbYear.ListCount Step 1
        If cmbYear.List(no) = YR Then
            i = i + 1  'if year already exists in the list
            no = 0
            GoTo restart
        End If
        If no = cmbYear.ListCount Then
            cmbYear.AddItem (YR) ' if year does not exist then add it
        End If
    Next no
Next i
End If
End Sub


'----------- FOR PRINTING PAYSLIP-------------

Private Sub cmdPrint_Click()
'On Error GoTo perr
'First delete any existing record in table PaySlip
If rsPrintPaySlip.RecordCount > 0 Then rsPrintPaySlip.MoveFirst
    For i = 0 To rsPrintPaySlip.RecordCount - 1 Step 1
    If rsPrintPaySlip.BOF = False Or rsPrintPaySlip.EOF = False Then
        rsPrintPaySlip.Delete
        If rsPrintPaySlip.EOF = False Then rsPrintPaySlip.MoveNext
    End If
    Next
Call AddPrintRecord
Call ShowReport
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
rsPrintPaySlip.AddNew
If Trim(txtCode.Text) <> "" Then rsPrintPaySlip!Code = txtCode.Text
If Trim(txtName.Text) <> "" Then rsPrintPaySlip!Name = txtName.Text
If Trim(cmbYear.Text) <> "" Then rsPrintPaySlip!Year = cmbYear.Text
If Trim(txtBasicPay.Text) <> "" Then rsPrintPaySlip!BasicPay = txtBasicPay.Text
If Trim(txtHouseRent.Text) <> "" Then rsPrintPaySlip!HRA = txtHouseRent.Text
If Trim(txtMnt.Text) <> "" Then rsPrintPaySlip!Month = txtMnt.Text
If Trim(txtDearness.Text) <> "" Then rsPrintPaySlip!DA = txtDearness.Text
If Trim(txtMedical.Text) <> "" Then rsPrintPaySlip!MA = txtMedical.Text
If Trim(txtConv.Text) <> "" Then rsPrintPaySlip!Conv = txtConv.Text
If Trim(txtLeavesAllowed.Text) <> "" Then rsPrintPaySlip!LeaveAllow = txtLeavesAllowed.Text
If Trim(txtTotAll.Text) <> "" Then rsPrintPaySlip!TotalAllowance = txtTotAll.Text
If Trim(txtProv.Text) <> "" Then rsPrintPaySlip!PF = txtProv.Text
If Trim(txtGInsur.Text) <> "" Then rsPrintPaySlip!GPF = txtGInsur.Text
If Trim(txtLIC.Text) <> "" Then rsPrintPaySlip!LIC = txtLIC.Text
If Trim(txtPTax.Text) <> "" Then rsPrintPaySlip!PTax = txtPTax.Text
If Trim(txtQRent.Text) <> "" Then rsPrintPaySlip!QRent = txtQRent.Text
If Trim(txtTotDed.Text) <> "" Then rsPrintPaySlip!TotalDeduction = txtTotDed.Text
If Trim(txtLoan.Text) <> "" Then rsPrintPaySlip!Loan = txtLoan.Text
If Trim(txtLeavesDed.Text) <> "" Then rsPrintPaySlip!LeaveDed = txtLeavesDed.Text
If lblNet.Caption <> "" Then rsPrintPaySlip!NetPay = lblNet.Caption
rsPrintPaySlip.Update
End Sub

Private Sub ShowReport()
'On Error GoTo Attnerr
Cr.Reset
Cr.ReportTitle = "MONTHLY PAY SLIP"
Cr.WindowState = crptMaximized
Cr.WindowTitle = "MONTHLY PAY SLIP"
Cr.WindowShowGroupTree = True
Cr.DataFiles(0) = App.Path & "\Project97.mdb"
Cr.ReportFileName = App.Path & "\Reports\PaySlip\rptPaySlip.rpt"

Cr.Destination = crptToWindow
Cr.WindowShowRefreshBtn = True
Cr.WindowShowGroupTree = False
'CR.Action = 1
Exit Sub
Attnerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

